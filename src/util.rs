use crate::error::Result;
use std::{ffi::OsStr, io, os::windows::prelude::OsStrExt, ptr};
use windows::{
    core::{Interface, BSTR, GUID, PCWSTR, PWSTR},
    Win32::{
        Foundation::{ERROR_SUCCESS, FILETIME, WIN32_ERROR},
        System::{
            Com::{
                CLSIDFromProgID, CLSIDFromString, CoCreateInstance, ITypeInfo, ITypeLib,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, TYPEDESC, VT_PTR, VT_SAFEARRAY,
            },
            Environment::ExpandEnvironmentStringsW,
            Ole::{OleInitialize, OleUninitialize},
            Registry::{
                RegCloseKey, RegEnumKeyExW, RegOpenKeyExW, RegQueryValueExW, HKEY, KEY_READ,
                REG_EXPAND_SZ, REG_QWORD, REG_VALUE_TYPE,
            },
        },
    },
};

thread_local!(static OLE_INITIALIZED: OleInitialized = {
    unsafe {
        OleInitialize(ptr::null_mut()).unwrap();
        OleInitialized(ptr::null_mut())
    }
});

/// RAII object that guards the fact that COM is initialized.
///
// We store a raw pointer because it's the only way at the moment to remove `Send`/`Sync` from the
// object.
struct OleInitialized(*mut ());

impl Drop for OleInitialized {
    #[inline]
    fn drop(&mut self) {
        unsafe { OleUninitialize() };
    }
}

/// Ensures that COM is initialized in this thread.
#[inline]
pub fn ole_initialized() {
    OLE_INITIALIZED.with(|_| {});
}

pub trait ToWide {
    fn to_wide(&self) -> Vec<u16>;
    fn to_wide_null(&self) -> Vec<u16>;
}

impl<T> ToWide for T
where
    T: AsRef<OsStr>,
{
    fn to_wide(&self) -> Vec<u16> {
        self.as_ref().encode_wide().collect()
    }
    fn to_wide_null(&self) -> Vec<u16> {
        self.as_ref().encode_wide().chain(Some(0)).collect()
    }
}

pub fn to_u16s<S: AsRef<OsStr>>(s: S) -> Result<Vec<u16>> {
    let mut maybe_result: Vec<u16> = s.to_wide();
    if maybe_result.iter().any(|&u| u == 0) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "strings passed to WinAPI cannot contain NULs",
        )
        .into());
    }
    maybe_result.push(0);
    Ok(maybe_result)
}

pub fn get_class_id<S: AsRef<OsStr>>(s: S) -> Result<GUID> {
    let prog_id = to_u16s(s)?;
    let prog_id = PCWSTR::from_raw(prog_id.as_ptr());

    unsafe {
        match CLSIDFromProgID(prog_id) {
            Ok(guid) => Ok(guid),
            Err(_error) => match CLSIDFromString(prog_id) {
                Ok(guid) => Ok(guid),
                Err(error) => Err(error.into()),
            },
        }
    }
}

pub fn create_instance<T: Interface>(clsid: &GUID) -> Result<T> {
    let flags = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;
    unsafe { Ok(CoCreateInstance(clsid, None, flags)?) }
}

pub fn create_com_object<S: AsRef<OsStr>, T: Interface>(s: S) -> Result<T> {
    ole_initialized();
    let class_id = get_class_id(s)?;

    create_instance(&class_id)
}

pub(crate) fn reg_open_key(hkey: HKEY, name: PCWSTR, phkey: &mut HKEY) -> WIN32_ERROR {
    unsafe { RegOpenKeyExW(hkey, name, 0, KEY_READ, phkey) }
}

pub(crate) fn reg_enum_key(hkey: HKEY, i: u32) -> PCWSTR {
    let mut buf = vec![0; 512 + 1];
    let buf_pwstr = PWSTR(buf.as_mut_ptr());
    let mut buf_size = buf.len() as u32;
    let mut ft = FILETIME::default();
    let _result = unsafe {
        RegEnumKeyExW(
            hkey,
            i,
            buf_pwstr,
            &mut buf_size,
            None,
            PWSTR::null(),
            None,
            Some(&mut ft),
        )
    };
    PCWSTR(buf_pwstr.0) //May contain a NULL buffer
}

pub fn reg_get_val(hkey: HKEY, subkey: Option<PCWSTR>) -> PCWSTR {
    let subkey_pcwstr = if let Some(subkey) = subkey {
        subkey
    } else {
        PCWSTR::null()
    };
    let mut dwtype = REG_VALUE_TYPE::default();
    let mut buf_len = 2048;
    let mut buf = Vec::with_capacity(buf_len as usize);
    loop {
        match unsafe {
            RegQueryValueExW(
                hkey,
                subkey_pcwstr,
                None,
                Some(&mut dwtype),
                Some(buf.as_mut_ptr()),
                Some(&mut buf_len),
            )
            .0
        } {
            0 => {
                // ERROR_SUCCESS
                unsafe {
                    buf.set_len(buf_len as usize);
                }
                // there shouldn't be anything higher than REG_QWORD, so just return a null string
                if dwtype.0 > REG_QWORD.0 {
                    return PCWSTR::null();
                } else {
                    let mut buf_str = String::from_utf8_lossy(&buf).to_string();
                    while buf_str.ends_with('\u{0}') {
                        buf_str.pop();
                    }
                    let buf_vec = buf_str.to_wide_null();
                    let buf_pcwstr = PCWSTR::from_raw(buf_vec.as_ptr());
                    if dwtype == REG_EXPAND_SZ {
                        let len = unsafe { ExpandEnvironmentStringsW(buf_pcwstr, None) };
                        let mut expanded_buf = vec![0; len as usize];
                        let _len = unsafe {
                            ExpandEnvironmentStringsW(buf_pcwstr, Some(&mut expanded_buf))
                        };
                        return PCWSTR::from_raw(expanded_buf.as_ptr());
                    } else {
                        return PCWSTR::from_raw(buf_pcwstr.as_ptr());
                    }
                };
            }
            234 => {
                // ERROR_MORE_DATA
                buf.reserve(buf_len as usize);
            }
            _err => return PCWSTR::null(),
        }
    }
}

pub(crate) fn reg_get_val2(hkey: HKEY, subkey: PCWSTR) -> PCWSTR {
    let mut hsubkey = HKEY::default();
    let mut val = PCWSTR::null();
    let result = unsafe { RegOpenKeyExW(hkey, subkey, 0, KEY_READ, &mut hsubkey) };
    if result == ERROR_SUCCESS {
        val = reg_get_val(hkey, None);
        unsafe { RegCloseKey(hsubkey) };
    }
    if val.is_null() {
        val = reg_get_val(hkey, Some(subkey));
    }
    val
}

pub(crate) fn ole_typedesc2val(
    typeinfo: &ITypeInfo,
    typedesc: &TYPEDESC,
    mut typedetails: Option<&mut Vec<String>>,
) -> String {
    let typestr = match typedesc.vt.0 {
        2 => "I2".into(),
        3 => "I4".into(),
        4 => "R4".into(),
        5 => "R8".into(),
        6 => "CY".into(),
        7 => "DATE".into(),
        8 => "BSTR".into(),
        11 => "BOOL".into(),
        12 => "VARIANT".into(),
        14 => "DECIMAL".into(),
        16 => "I1".into(),
        17 => "UI1".into(),
        18 => "UI2".into(),
        19 => "UI4".into(),
        20 => "I8".into(),
        21 => "UI8".into(),
        22 => "INT".into(),
        23 => "UINT".into(),
        24 => "VOID".into(),
        25 => "HRESULT".into(),
        26 => {
            let typestr: String = "PTR".into();
            if let Some(ref mut typedetails) = typedetails {
                typedetails.push(typestr);
            }
            return ole_ptrtype2val(typeinfo, typedesc, typedetails);
        }
        27 => {
            let typestr: String = "SAFEARRAY".into();
            if let Some(ref mut typedetails) = typedetails {
                typedetails.push(typestr);
            }
            return ole_ptrtype2val(typeinfo, typedesc, typedetails);
        }
        28 => "CARRAY".into(),
        29 => {
            let typestr: String = "USERDEFINED".into();
            if let Some(ref mut typedetails) = typedetails {
                typedetails.push(typestr.clone());
            }
            let str = ole_usertype2val(typeinfo, typedesc, typedetails);
            if let Some(str) = str {
                return str;
            }
            return typestr;
        }
        13 => "UNKNOWN".into(),
        9 => "DISPATCH".into(),
        10 => "ERROR".into(),
        31 => "LPWSTR".into(),
        30 => "LPSTR".into(),
        36 => "RECORD".into(),
        _ => {
            let typestr: String = "Unknown Type ".into();
            format!("{}{}", typestr, typedesc.vt.0)
        }
    };
    if let Some(typedetails) = typedetails {
        typedetails.push(typestr.clone());
    }
    typestr
}

pub(crate) fn ole_ptrtype2val(
    typeinfo: &ITypeInfo,
    typedesc: &TYPEDESC,
    typedetails: Option<&mut Vec<String>>,
) -> String {
    let mut type_ = "".into();

    if typedesc.vt == VT_PTR || typedesc.vt == VT_SAFEARRAY {
        let p = unsafe { typedesc.Anonymous.lptdesc };
        type_ = ole_typedesc2val(typeinfo, unsafe { &*p }, typedetails);
    }
    type_
}

pub(crate) fn ole_usertype2val(
    typeinfo: &ITypeInfo,
    typedesc: &TYPEDESC,
    typedetails: Option<&mut Vec<String>>,
) -> Option<String> {
    let result = unsafe { typeinfo.GetRefTypeInfo(typedesc.Anonymous.hreftype) };
    if result.is_err() {
        return None;
    }
    let reftypeinfo = result.unwrap();
    let mut bstrname = BSTR::default();
    let result = ole_docinfo_from_type(
        &reftypeinfo,
        Some(&mut bstrname),
        None,
        ptr::null_mut(),
        None,
    );
    if result.is_err() {
        return None;
    }
    let type_ = bstrname.to_string();
    if let Some(typedetails) = typedetails {
        typedetails.push(type_.clone());
    }
    Some(type_)
}

fn ole_docinfo_from_type(
    typeinfo: &ITypeInfo,
    name: Option<*mut BSTR>,
    helpstr: Option<*mut BSTR>,
    helpcontext: *mut u32,
    helpfile: Option<*mut BSTR>,
) -> Result<()> {
    let mut typelib: Option<ITypeLib> = None;
    let mut index = 0;
    unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut index)? };
    let typelib = typelib.unwrap();
    unsafe { typelib.GetDocumentation(index as i32, name, helpstr, helpcontext, helpfile)? };
    Ok(())
}

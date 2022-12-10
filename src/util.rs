use crate::error::Result;
use std::{ffi::OsStr, io, os::windows::prelude::OsStrExt, ptr};
use windows::{
    core::{Interface, BSTR, GUID, PCSTR, PCWSTR, PWSTR},
    Win32::{
        Foundation::{ERROR_SUCCESS, FILETIME, WIN32_ERROR},
        System::{
            Com::{
                CLSIDFromProgID, CLSIDFromString, CoCreateInstance, ITypeInfo, ITypeLib,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, TYPEDESC, VT_PTR, VT_SAFEARRAY,
            },
            Environment::ExpandEnvironmentStringsA,
            Ole::{OleInitialize, OleUninitialize},
            Registry::{
                RegCloseKey, RegEnumKeyExW, RegOpenKeyExW, HKEY, KEY_READ,
                REG_EXPAND_SZ, REG_VALUE_TYPE, RegQueryValueExA,
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
    let result = unsafe {
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
    if result == ERROR_SUCCESS {
        PCWSTR::from_raw(buf.as_ptr())
    } else {
        PCWSTR::null()
    }
}

pub fn reg_get_val(hkey: HKEY, subkey: Option<PCWSTR>) -> PCWSTR {
    let subkey_pcstr = if let Some(subkey) = subkey {
        let subkey_str = unsafe { subkey.to_string().unwrap() };
        let mut subkey_vec = subkey_str.into_bytes();
        subkey_vec.push(0);
        PCSTR::from_raw(subkey_vec.as_ptr())
    } else {
        PCSTR::null()
    };

    let mut dwtype = REG_VALUE_TYPE::default();
    let mut buf_len = 0;
    let result = unsafe {
        RegQueryValueExA(
            hkey,
            subkey_pcstr,
            None,
            Some(&mut dwtype),
            None,
            Some(&mut buf_len),
        )
    };

    if result == ERROR_SUCCESS {
        let mut buf = vec![0; buf_len as usize + 1];

        let result = unsafe {
            RegQueryValueExA(
                hkey,
                subkey_pcstr,
                None,
                Some(&mut dwtype),
                Some(buf.as_mut_ptr()),
                Some(&mut buf_len),
            )
        };

        if result == ERROR_SUCCESS {
            let buf_pcstr = PCSTR::from_raw(buf.as_ptr());
            if dwtype == REG_EXPAND_SZ {
                let len = unsafe { ExpandEnvironmentStringsA(buf_pcstr, None) };
                let mut expanded_buf = vec![0; len as usize + 1];
                let _len = unsafe { ExpandEnvironmentStringsA(buf_pcstr, Some(&mut expanded_buf)) };
                let expanded_buf_str = unsafe { buf_pcstr.to_string().unwrap() };
                let expanded_buf_u16vec = expanded_buf_str.to_wide_null();
                return PCWSTR::from_raw(expanded_buf_u16vec.as_ptr());
            }
            let buf_str = unsafe { buf_pcstr.to_string().unwrap() };
            let buf_vecu16 = buf_str.to_wide_null();
            let buf_pcwstr = PCWSTR::from_raw(buf_vecu16.as_ptr());
            return buf_pcwstr;
        }
        println!("In here, result is {:?}", result);
    }
    PCWSTR::null()
}

pub(crate) fn reg_get_val2(hkey: HKEY, subkey: PCWSTR) -> PCWSTR {
    let mut hsubkey = HKEY::default();
    let mut val = PCWSTR::null();
    let result = unsafe { RegOpenKeyExW(hkey, subkey, 0, KEY_READ, &mut hsubkey) };
    if result == ERROR_SUCCESS {
        val = reg_get_val(hsubkey, None);
        unsafe { RegCloseKey(hsubkey) };
    }
    if val.is_null() {
        val = reg_get_val(hkey, Some(subkey));
    }
    val
}

pub(crate) fn reg_get_val2_string(hkey: HKEY, subkey: PCWSTR) -> Option<String> {
    let result = reg_get_val2(hkey, subkey);
    match result.is_null() {
        false => {
            if let Ok(str) = unsafe { result.to_string() } {
                Some(str)
            } else {
                None
            }
        }
        true => None,
    }
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

/*pub(crate) fn check_nano_server() {
    let mut hsubkey = HKEY::default();
    let subkey =
        "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Server\\ServerLevels".to_wide_null();
    let subkey_pcwstr = PCWSTR::from_raw(subkey.as_ptr());
    let regval = "NanoServer".to_wide_null();
    let regval_pcwstr = PCWSTR::from_raw(regval.as_ptr());

    let result =
        unsafe { RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey_pcwstr, 0, KEY_READ, &mut hsubkey) };
    if result == ERROR_SUCCESS {
        let result = unsafe { RegQueryValueExW(hsubkey, regval_pcwstr, None, None, None, None) };
        if result == ERROR_SUCCESS {
            g_running_nano = TRUE;
        }
        unsafe { RegCloseKey(hsubkey) };
    }
}*/

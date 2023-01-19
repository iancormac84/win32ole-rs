use crate::{
    error::{OleError, Result},
    ToWide, G_RUNNING_NANO,
};
use std::{ffi::OsStr, ptr};
use windows::{
    core::{Interface, BSTR, GUID, PCWSTR},
    Win32::System::{
        Com::{
            CLSIDFromProgID, CLSIDFromString, CoCreateInstance, CoIncrementMTAUsage, CoInitializeEx, CoUninitialize,
            ITypeInfo, ITypeLib, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, COINIT_MULTITHREADED, CO_MTA_USAGE_COOKIE,
            TYPEDESC, VT_PTR, VT_SAFEARRAY,
        },
        Ole::{OleInitialize, OleUninitialize},
    },
};

/// Initialize a new multithreaded apartment (MTA) runtime. This will ensure
/// that an MTA is running for the process. Every new thread will implicitly
/// be in the MTA unless a different apartment type is chosen (through [`init_apartment`])
///
/// This calls `CoIncrementMTAUsage`
///
/// This function only needs to be called once per process.
pub fn init_runtime() -> windows::core::Result<CO_MTA_USAGE_COOKIE> {
    match unsafe { CoIncrementMTAUsage() } {
        // S_OK indicates the runtime was initialized
        S_OK => Ok(cookie),
        // Any other result is considered an error here.
        hr => Err(hr),
    }
}

thread_local!(static OLE_INITIALIZED: OleInitialized = {
    unsafe {
        let result = if *G_RUNNING_NANO {
            CoInitializeEx(None, COINIT_MULTITHREADED)
        } else {
            OleInitialize(ptr::null_mut())
        };
        if let Err(error) = result {
            let runtime_error = OleError::runtime(error, "failed: OLE initialization");
            panic!("{runtime_error}");
        }
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
        if *G_RUNNING_NANO {
            unsafe { CoUninitialize() };
        } else {
            unsafe { OleUninitialize() };
        }
    }
}

/// Ensures that COM is initialized in this thread.
#[inline]
pub fn ole_initialized() {
    OLE_INITIALIZED.with(|_| {});
}

pub fn get_class_id<S: AsRef<OsStr>>(s: S) -> Result<GUID> {
    let prog_id = s.to_wide_null();
    let prog_id = PCWSTR::from_raw(prog_id.as_ptr());

    unsafe {
        match CLSIDFromProgID(prog_id) {
            Ok(guid) => Ok(guid),
            Err(_error) => match CLSIDFromString(prog_id) {
                Ok(guid) => Ok(guid),
                Err(error) => Err(OleError::runtime(
                    error,
                    format!("unknown OLE server: `{}`", s.as_ref().to_str().unwrap()),
                )
                .into()),
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

pub trait TypeRef {
    fn typeinfo(&self) -> &ITypeInfo;
    fn typedesc(&self) -> &TYPEDESC;
}

pub trait ValueDescription: TypeRef {
    fn ole_typedesc2val(&self, mut typedetails: Option<&mut Vec<String>>) -> String {
        let p = unsafe { self.typedesc().Anonymous.lptdesc };
        let typestr = match unsafe { (*p).vt.0 } {
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
                return self.ole_ptrtype2val(typedetails);
            }
            27 => {
                let typestr: String = "SAFEARRAY".into();
                if let Some(ref mut typedetails) = typedetails {
                    typedetails.push(typestr);
                }
                return self.ole_ptrtype2val(typedetails);
            }
            28 => "CARRAY".into(),
            29 => {
                let typestr: String = "USERDEFINED".into();
                if let Some(ref mut typedetails) = typedetails {
                    typedetails.push(typestr.clone());
                }
                let str = self.ole_usertype2val(typedetails);
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
                format!("{}{}", typestr, self.typedesc().vt.0)
            }
        };
        if let Some(typedetails) = typedetails {
            typedetails.push(typestr.clone());
        }
        typestr
    }

    fn ole_ptrtype2val(&self, typedetails: Option<&mut Vec<String>>) -> String {
        let mut type_ = "".into();

        if self.typedesc().vt == VT_PTR || self.typedesc().vt == VT_SAFEARRAY {
            type_ = self.ole_typedesc2val(typedetails);
        }
        type_
    }

    fn ole_usertype2val(&self, typedetails: Option<&mut Vec<String>>) -> Option<String> {
        let result = unsafe {
            self.typeinfo()
                .GetRefTypeInfo(self.typedesc().Anonymous.hreftype)
        };
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
}

pub(crate) fn ole_docinfo_from_type(
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

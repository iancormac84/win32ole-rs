use std::ptr;

use windows::{
    core::{Interface, BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::{DISP_E_EXCEPTION, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH},
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{
                IDispatch, ITypeInfo, ITypeLib, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, INVOKE_FUNC,
                INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF,
            },
            Environment::ExpandEnvironmentStringsW,
            Ole::DISPID_PROPERTYPUT,
            Variant::VARIANT,
        },
    },
};
use windows_core::HSTRING;
use windows_registry::{Key, Type};

use crate::{
    error::{ComArgumentErrorType, Error, OleError, Result},
    olemethoddata::{ole_methods_from_typeinfo, OleMethodData},
    types::OleClassNames,
    util::{create_com_object, get_class_id},
    OleTypeData, OleTypeLibData,
};

/*#[inline]
pub unsafe fn ShowHTMLDialogEx<P0, P1>(
    hwndparent: P0,
    moniker: *const IMoniker,
    dialogflags: u32,
    variant_arg_in: *const VARIANT,
    options: P1,

) -> ::windows::Win32::Foundation::HWND
where
    P0: ::std::convert::Into<::windows::Win32::Foundation::HWND>,
    P1: ::std::convert::Into<::windows::core::InParam<::windows::core::PCWSTR>>,
{
    ::windows::core::link ! ( "Mshtml.dll""system" fn ShowHTMLDialogEx ( hwndparent : ::windows::Win32::Foundation:: HWND , moniker : *const :: windows::Win32::System::Com::IMoniker , dialogflags : u32 , variant_arg_in : Option<*const :: windows::Win32::System::Com::VARIANT> , options:  ) -> ::windows::Win32::Foundation:: HWND );
    ShowHTMLDialogEx(hwndparent.into(), moniker, dialogflags, variant_arg_in)
}*/

pub struct OleData {
    pub dispatch: IDispatch,
}
impl OleData {
    pub fn new<H: Into<HSTRING>>(prog_id: H) -> Result<Self> {
        Ok(OleData {
            dispatch: create_com_object(prog_id)?,
        })
    }
    pub fn get_ids_of_names<H: Into<HSTRING> + Copy>(&self, names: &[H]) -> Result<Vec<i32>> {
        let namelen = names.len();
        let mut wnames = vec![PCWSTR::null(); namelen];
        for i in 0..namelen {
            let a = &names[i].into();
            wnames[i] = PCWSTR(a.as_ptr());
        }

        let mut dispids = 0;

        unsafe {
            self.dispatch.GetIDsOfNames(
                &GUID::zeroed(),
                wnames.as_ptr(),
                wnames.len() as u32,
                GetUserDefaultLCID(),
                &mut dispids,
            )
        }?;

        let ids = unsafe { Vec::from_raw_parts(&mut dispids, wnames.len(), wnames.len()) };

        Ok(ids)
    }
    pub fn responds_to<H: Into<HSTRING>>(&self, method: H) -> bool {
        let method = method.into();
        let methods = [PCWSTR(method.as_ptr())];
        let mut dispids = 0;

        unsafe {
            self.dispatch
                .GetIDsOfNames(
                    &GUID::zeroed(),
                    methods.as_ptr(),
                    1,
                    GetUserDefaultLCID(),
                    &mut dispids,
                )
                .is_ok()
        }
    }
    fn get_type_info(&self) -> Result<ITypeInfo> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID()) };
        match typeinfo {
            Ok(typeinfo) => Ok(typeinfo),
            Err(error) => Err(OleError::interface(error, "failed to GetTypeInfo").into()),
        }
    }
    pub fn ole_type(&self) -> Result<OleTypeData> {
        let typeinfo = self.get_type_info()?;
        OleTypeData::try_from(typeinfo)
    }
    pub fn ole_typelib(&self) -> Result<OleTypeLibData> {
        let typeinfo = self.get_type_info()?;
        OleTypeLibData::try_from(&typeinfo)
    }
    fn raw_ole_methods(&self, mask: i32) -> Result<Vec<OleMethodData>> {
        let mut methods = vec![];

        let typeinfo = self.typeinfo_from_ole()?;
        methods.extend(ole_methods_from_typeinfo(typeinfo, mask)?);
        Ok(methods)
    }
    pub fn ole_methods(&self) -> Result<Vec<OleMethodData>> {
        self.raw_ole_methods(
            INVOKE_FUNC.0 | INVOKE_PROPERTYGET.0 | INVOKE_PROPERTYPUT.0 | INVOKE_PROPERTYPUTREF.0,
        )
    }
    pub fn ole_get_methods(&self) -> Result<Vec<OleMethodData>> {
        self.raw_ole_methods(INVOKE_PROPERTYGET.0)
    }
    pub fn ole_put_methods(&self) -> Result<Vec<OleMethodData>> {
        self.raw_ole_methods(INVOKE_PROPERTYPUT.0 | INVOKE_PROPERTYPUTREF.0)
    }
    pub fn ole_func_methods(&self) -> Result<Vec<OleMethodData>> {
        self.raw_ole_methods(INVOKE_FUNC.0)
    }
    fn typeinfo_from_ole(&self) -> Result<ITypeInfo> {
        let typeinfo = self.get_type_info()?;

        let mut bstrname = BSTR::default();
        unsafe { typeinfo.GetDocumentation(-1, Some(&mut bstrname), None, ptr::null_mut(), None)? };
        let type_ = bstrname;
        let mut typelib: Option<ITypeLib> = None;
        let mut i = 0;
        let result = unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut i) };
        if let Err(error) = result {
            return Err(OleError::interface(error, "failed to GetContainingTypeLib").into());
        };

        let typelib = typelib.unwrap();

        let ole_class_names = OleClassNames::from(&typelib);
        let mut ret_type_info = None;
        for (idx, class_name) in ole_class_names.enumerate() {
            if let Ok(class_name) = class_name {
                if class_name == type_ {
                    let result = unsafe { typelib.GetTypeInfo(idx as u32) };
                    if let Ok(ret_type) = result {
                        ret_type_info = Some(ret_type);
                        break;
                    }
                }
            }
        }
        Ok(ret_type_info.unwrap())
    }
    pub fn ole_query_interface<H: Into<HSTRING>>(&self, str_iid: H) -> Result<OleData> {
        let iid = get_class_id(str_iid)?;
        let mut dispatch_interface = ptr::null_mut();
        let result = unsafe { self.dispatch.query(&iid, &mut dispatch_interface) };
        let result = result.ok();
        if let Err(error) = result {
            Err(error.into())
        } else {
            let dispatch: IDispatch =
                unsafe { <IDispatch as Interface>::from_raw(dispatch_interface as *mut _) };
            Ok(OleData { dispatch })
        }
    }
    pub fn ole_method_help<S: AsRef<str>>(&self, cmdname: S) -> Result<OleMethodData> {
        let cmdname = cmdname.as_ref();
        let typeinfo = self.typeinfo_from_ole();
        let Ok(typeinfo) = typeinfo else {
            return Err(Error::Custom(format!(
                "failed to get ITypeInfo: {}",
                typeinfo.err().unwrap()
            )));
        };
        let obj = OleMethodData::from_typeinfo(typeinfo, cmdname)?;

        if let Some(obj) = obj {
            Ok(obj)
        } else {
            Err(Error::Custom(format!("not found {cmdname}",)))
        }
    }

    pub fn invoke<H: Into<HSTRING> + Copy>(
        &self,
        name: H,
        dp: &mut DISPPARAMS,
        flags: DISPATCH_FLAGS,
    ) -> Result<VARIANT> {
        let ids = self.get_ids_of_names(&[name])?;

        let mut excep = EXCEPINFO::default();
        let mut arg_err = 0;
        let mut result = VARIANT::default();

        let res = unsafe {
            self.dispatch.Invoke(
                ids[0],
                &GUID::zeroed(),
                0x0800, /*LOCALE_SYSTEM_DEFAULT*/
                flags,
                dp,
                Some(&mut result),
                Some(&mut excep),
                Some(&mut arg_err),
            )
        };

        match res {
            Ok(()) => Ok(result),
            Err(e) => Err(match e.code() {
                DISP_E_EXCEPTION => Error::Exception(excep),
                DISP_E_TYPEMISMATCH => Error::IDispatchArgument {
                    error_type: ComArgumentErrorType::TypeMismatch,
                    arg_err,
                },
                DISP_E_PARAMNOTFOUND => Error::IDispatchArgument {
                    error_type: ComArgumentErrorType::ParameterNotFound,
                    arg_err,
                },
                _ => e.into(),
            }),
        }
    }

    /// Get a property from a COM object
    ///
    pub fn get(&self, name: &str) -> Result<VARIANT> {
        let mut dp = DISPPARAMS::default();
        self.invoke(name, &mut dp, DISPATCH_PROPERTYGET)
    }

    /// Set a property on a COM object
    ///
    pub fn put(&self, name: &str, value: &mut VARIANT) -> Result<()> {
        let mut dp = DISPPARAMS {
            cArgs: 1,
            rgvarg: value,
            cNamedArgs: 1,
            ..Default::default()
        };
        let mut id = DISPID_PROPERTYPUT;
        dp.rgdispidNamedArgs = &mut id as *mut _;
        self.invoke(name, &mut dp, DISPATCH_PROPERTYPUT)?;
        Ok(())
    }

    /// Call a method on a COM object
    ///
    pub fn call(&self, name: &str, args: Vec<VARIANT>) -> Result<VARIANT> {
        let mut dp = DISPPARAMS::default();
        let args: Vec<VARIANT> = args.into_iter().rev().collect();
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_ptr() as *mut _;
        self.invoke(name, &mut dp, DISPATCH_METHOD)
    }
}

pub unsafe fn reg_get_val<N: AsRef<PCWSTR>>(key: &Key, subkey: N) -> Result<String> {
    let (ty, _) = key.raw_get_info(&subkey)?;
    let subkey = if subkey.as_ref().is_null() {
        "".to_string()
    } else {
        subkey.as_ref().to_string().unwrap()
    };
    let data = HSTRING::from(key.get_string(&subkey)?);
    if ty == Type::ExpandString {
        let len = ExpandEnvironmentStringsW(&data, None);
        let mut expanded_data = vec![0; len as usize + 1];
        unsafe { ExpandEnvironmentStringsW(&data, Some(&mut expanded_data)) };
        let expanded_data_string = String::from_utf16_lossy(&expanded_data);
        return Ok(expanded_data_string);
    }
    Ok(data.to_string())
}

/*fn clsid_from_remote<H: Into<HSTRING>, S: AsRef<str>>(host: H, com: S) -> windows::core::Result<()> {
    let host = host.into();
    let mut hlm = ptr::null_mut();
    let result = unsafe { RegConnectRegistryW(&host, HKEY::from(HKEY_LOCAL_MACHINE), hlm) };
    if result != ERROR_SUCCESS {
        return Err(windows::core::Error::from_hresult(HRESULT::from_win32(result.0)));
    };
    let mut subkey = String::from("SOFTWARE\\Classes\\");
    subkey.push_str(com.as_ref());
    subkey.push_str("\\CLSID");
    let hlm = unsafe { Key::from_raw(&hlm) };
    let result = hlm.open(subkey);
    if let Err(error) = result {
        return Err(error);
    } else {
        len = sizeof(clsid);
        err = RegQueryValueEx(hpid, "", NULL, &dwtype, (BYTE *)clsid, &len);
        if (err == ERROR_SUCCESS && dwtype == REG_SZ) {
            pbuf = ole_mb2wc(clsid, -1, cWIN32OLE_cp);
            hr = CLSIDFromString(pbuf, pclsid);
            SysFreeString(pbuf);
        }
        else {
            hr = HRESULT_FROM_WIN32(err);
        }
        RegCloseKey(hpid);
    }
    RegCloseKey(hlm);
    return hr;
}

fn ole_create_dcom(VALUE self, VALUE ole, VALUE host, VALUE others)
{
    HRESULT hr;
    CLSID   clsid;
    OLECHAR *pbuf;

    COSERVERINFO serverinfo;
    MULTI_QI multi_qi;
    DWORD clsctx = CLSCTX_REMOTE_SERVER;

    pbuf  = ole_vstr2wc(ole);
    hr = CLSIDFromProgID(pbuf, &clsid);
    if (FAILED(hr))
        hr = clsid_from_remote(host, ole, &clsid);
    if (FAILED(hr))
        hr = CLSIDFromString(pbuf, &clsid);
    SysFreeString(pbuf);
    if (FAILED(hr))
        ole_raise(hr, eWIN32OLERuntimeError,
                  "unknown OLE server: `%s'",
                  StringValuePtr(ole));
    memset(&serverinfo, 0, sizeof(COSERVERINFO));
    serverinfo.pwszName = ole_vstr2wc(host);
    memset(&multi_qi, 0, sizeof(MULTI_QI));
    multi_qi.pIID = &IID_IDispatch;
    hr = gCoCreateInstanceEx(&clsid, NULL, clsctx, &serverinfo, 1, &multi_qi);
    SysFreeString(serverinfo.pwszName);
    if (FAILED(hr))
        ole_raise(hr, eWIN32OLERuntimeError,
                  "failed to create DCOM server `%s' in `%s'",
                  StringValuePtr(ole),
                  StringValuePtr(host));

    ole_set_member(self, (IDispatch*)multi_qi.pItf);
    return self;
}*/

/*pub enum HelpTarget<'a> {
    OleType(OleTypeData),
    OleMethod(OleMethodData<'a>),
    HelpFile(PathBuf),
}

impl<'a> From<OleTypeData> for HelpTarget<'a> {
    fn from(value: OleTypeData) -> Self {
        HelpTarget::OleType(value)
    }
}

impl<'a> From<OleMethodData<'a>> for HelpTarget<'a> {
    fn from(value: OleMethodData) -> Self {
        HelpTarget::OleMethod(value)
    }
}

impl<'a> From<PathBuf> for HelpTarget<'a> {
    fn from(value: PathBuf) -> Self {
        HelpTarget::HelpFile(value)
    }
}

pub fn ole_show_help<H: Into<HelpTarget>>(target: H, helpcontext: Option<u32>) -> Result<HWND> {
    let target = target.into();
    use HelpTarget::*;
    let (helpfile, helpcontext) = match target {
        OleType(oletypedata) => {
            let helpfile = oletypedata.helpfile()?;
            if helpfile.is_empty() {
                return Err(Error::Custom(format!(
                    "no helpfile found for {}",
                    oletypedata.name
                )));
            }
            let helpcontext = oletypedata.helpcontext()?;
            (helpfile, Some(helpcontext))
        }
        OleMethod(olemethoddata) => {
            let helpfile = olemethoddata.helpfile()?;
            if helpfile.is_empty() {
                return Err(Error::Custom(format!(
                    "no helpfile found for {}",
                    olemethoddata.name()
                )));
            }
            let helpcontext = olemethoddata.helpcontext()?;
            (helpfile, Some(helpcontext))
        }
        HelpFile(helpfile) => (helpfile.to_str().unwrap().to_string(), helpcontext),
    };
    ole_show_help_(helpfile, helpcontext.unwrap_or(0) as usize)
}

fn ole_show_help_<S: AsRef<OsStr>>(helpfile: S, helpcontext: usize) -> Result<HWND> {
    let helpfile = helpfile.as_ref().to_wide_null();
    let pszfile = PCWSTR::from_raw(helpfile.as_ptr());
    let mut hwnd = unsafe {
        HtmlHelpW(
            GetDesktopWindow(),
            pszfile,
            HTML_HELP_COMMAND(0x0f),
            helpcontext,
        )
    };
    if hwnd.0 == 0 {
        hwnd = unsafe {
            HtmlHelpW(
                GetDesktopWindow(),
                pszfile,
                HTML_HELP_COMMAND(0),
                helpcontext,
            )
        };
    }
    Ok(hwnd)
}*/

#[cfg(test)]
mod tests {
    #[test]
    fn test_methods() {
        let obj = super::OleData::new("Scripting.Dictionary");
        assert!(obj.is_ok());
        let obj = obj.unwrap();

        let methods = obj.ole_methods();
        assert!(methods.is_ok());
        let methods = methods.unwrap();

        let res: Vec<&super::OleMethodData> = methods
            .iter()
            .filter_map(|m| {
                if m.invoke_kind() == "PROPERTYPUTREF" {
                    Some(m)
                } else {
                    None
                }
            })
            .collect();
        assert_eq!(res.len(), 1);
        assert_eq!(res[0].name(), "Item");
    }
}

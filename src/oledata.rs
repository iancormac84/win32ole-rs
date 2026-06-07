use std::{ffi::OsStr, os::windows::ffi::OsStrExt, ptr};

use windows::{
    core::{Interface, BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::{DISP_E_EXCEPTION, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH, ERROR_SUCCESS},
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{
                CLSIDFromProgID, CLSIDFromString, CoCreateInstanceEx, CoGetClassObject,
                CreateBindCtx, IDispatch, ITypeInfo, ITypeLib, MkParseDisplayName,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER, COSERVERINFO,
                DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
                DISPPARAMS, EXCEPINFO, INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
                INVOKE_PROPERTYPUTREF, MULTI_QI,
            },
            Environment::ExpandEnvironmentStringsW,
            Ole::{GetActiveObject, IClassFactory2, DISPID_PROPERTYPUT},
            Registry::{RegConnectRegistryW, HKEY_LOCAL_MACHINE},
            Variant::VARIANT,
        },
    },
};
use windows_core::{HSTRING, PWSTR};
use windows_registry::{Key, Type};

use crate::{
    error::{ComArgumentErrorType, Error, OleError, Result},
    ole_initialized,
    olemethoddata::{ole_methods_from_typeinfo, OleMethodData},
    types::OleClassNames,
    util::{create_instance, get_class_id},
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
    /// Returns a new OLE Automation object.
    /// The first argument `svr_name` specifies the OLE Automation server and should be a GUID (CLSID) or PROGID.
    ///
    pub fn new<S: AsRef<str>>(svr_name: S, host: Option<S>, license: Option<S>) -> Result<Self> {
        ole_initialized();
        let svr_name = svr_name.as_ref();
        if let Some(host) = host {
            let host = host.as_ref();
            return ole_create_dcom(svr_name, host /*, others*/);
        }

        /* get CLSID from OLE server name */
        let clsid = get_class_id(svr_name)?;

        let result = match license {
            None => {
                /* get IDispatch interface */
                create_instance(&clsid)
            }
            Some(license) => {
                let class_factory: IClassFactory2 = unsafe {
                    CoGetClassObject(&clsid, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, None)?
                };
                let license = license.as_ref();
                let bstrkey = BSTR::from(license);
                unsafe { class_factory.CreateInstanceLic(None, None, &bstrkey) }
            }
        };

        if let Err(error) = result {
            return Err(OleError::runtime(
                error,
                format!("failed to create WIN32OLE object from `{svr_name}`"),
            )
            .into());
        }
        let dispatch = result.unwrap();

        Ok(OleData { dispatch })
    }

    /* TODO: will have to figure out if the extra arguments are really necessary for the Rust
       implementation. The signature for the Ruby function is `fole_s_connect(int argc, VALUE *argv, VALUE self)`
       and then inside the function there is `rb_scan_args(argc, argv, "1*", &svr_name, &others);`
       The `svr_name` is a mandatory argument and it's extracted from `argv`. There is a splatted argument that goes
       into `others`. `self` in the function signature is passed to `create_win32ole_object(self, pDispatch, argc, argv)`
       which allows the creation of the oledata object. I don't see any use of `argc` and `argv` in this particular
       call to `create_win32ole_object`.
    */
    pub fn connect<S: AsRef<str>>(svr_name: S /*, VALUE *argv, VALUE self*/) -> Result<Self> {
        ole_initialized();
        let svr_name = svr_name.as_ref();

        /* get CLSID from OLE server name */
        let clsid = get_class_id(svr_name);
        if clsid.is_err() {
            return ole_bind_obj(svr_name);
        }

        let clsid = clsid.unwrap();

        let mut unknown = None;
        let result = unsafe { GetActiveObject(&clsid, None, &mut unknown) };
        if let Err(error) = result {
            return Err(
                OleError::runtime(error, format!("OLE server `{svr_name}` not running")).into(),
            );
        }
        let unknown = unknown.unwrap();
        let dispatch: windows::core::Result<IDispatch> = unknown.cast();
        if let Err(error) = dispatch {
            return Err(OleError::runtime(
                error,
                format!("failed to create WIN32OLE server `{svr_name}`"),
            )
            .into());
        }
        let dispatch = dispatch.unwrap();

        Ok(OleData { dispatch })
    }
    pub fn get_ids_of_names<H: Into<HSTRING> + Copy>(&self, names: &[H]) -> Result<Vec<i32>> {
        let namelen = names.len();
        let mut wnames = vec![PCWSTR::null(); namelen];
        for i in 0..namelen {
            let a = &names[i].into();
            wnames[i] = PCWSTR(a.as_ptr());
        }
        let mut dispids = vec![0; namelen];

        unsafe {
            self.dispatch.GetIDsOfNames(
                &GUID::zeroed(),
                wnames.as_ptr(),
                wnames.len() as u32,
                GetUserDefaultLCID(),
                dispids.as_mut_ptr()
            )
        }?;

        Ok(dispids)
    }
    pub fn responds_to<H: Into<HSTRING>>(&self, method: H) -> bool {
        let method = method.into();
        let methods = [PCWSTR(method.as_ptr())];
        let mut dispids = vec![0; 1];

        unsafe {
            self.dispatch
                .GetIDsOfNames(
                    &GUID::zeroed(),
                    methods.as_ptr(),
                    1,
                    GetUserDefaultLCID(),
                    dispids.as_mut_ptr(),
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
        let str_iid = str_iid.into();
        let iid = match unsafe { CLSIDFromString(&str_iid) } {
            Ok(guid) => guid,
            Err(error) => {
                return Err(OleError::runtime(error, format!("invalid iid: `{}`", str_iid.display())).into())
            }
        };
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
    let (ty, _) = unsafe { key.raw_get_info(&subkey)? };
    let subkey = if subkey.as_ref().is_null() {
        "".to_string()
    } else {
        unsafe { subkey.as_ref().to_string().unwrap() }
    };
    let data = HSTRING::from(key.get_string(&subkey)?);
    if ty == Type::ExpandString {
        let len = unsafe { ExpandEnvironmentStringsW(&data, None) };
        let mut expanded_data = vec![0; len as usize + 1];
        unsafe { ExpandEnvironmentStringsW(&data, Some(&mut expanded_data)) };
        let expanded_data_string = String::from_utf16_lossy(&expanded_data);
        return Ok(expanded_data_string);
    }
    Ok(String::try_from(data)?)
}

fn ole_bind_obj<H: Into<HSTRING>>(
    moniker: H, /*int argc, VALUE *argv, VALUE self*/
) -> Result<OleData> {
    ole_initialized();
    let buf = moniker.into();

    let mut eaten = 0;

    let bind_ctx = unsafe { CreateBindCtx(0) };
    if let Err(error) = bind_ctx {
        return Err(OleError::runtime(error, "failed to create bind context").into());
    }
    let bind_ctx = bind_ctx.unwrap();
    let mut moniker = None;

    let result = unsafe { MkParseDisplayName(&bind_ctx, &buf, &mut eaten, &mut moniker) };
    if let Err(error) = result {
        return Err(
            OleError::runtime(error, "failed to parse display name of moniker `{buf}`").into(),
        );
    }
    let moniker = moniker.unwrap();

    let result: windows::core::Result<IDispatch> = unsafe { moniker.BindToObject(&bind_ctx, None) };
    if let Err(error) = result {
        return Err(OleError::runtime(error, "failed to bind moniker `buf`").into());
    }
    let dispatch = result.unwrap();
    Ok(OleData { dispatch })
}

fn clsid_from_remote<H: Into<HSTRING>, S: AsRef<str>>(host: H, com: S) -> Result<GUID> {
    let host = host.into();
    let hlm = ptr::null_mut();
    let result = unsafe { RegConnectRegistryW(&host, HKEY_LOCAL_MACHINE, hlm) };
    if result != ERROR_SUCCESS {
        return Err(result.into());
    };
    let mut subkey = String::from("SOFTWARE\\Classes\\");
    subkey.push_str(com.as_ref());
    subkey.push_str("\\CLSID");
    let hlm = unsafe { Key::from_raw((*hlm).0) };
    let result = hlm.open(subkey);
    if let Err(error) = result {
        Err(error.into())
    } else {
        let hpid = result.unwrap();
        let result = hpid.get_string("");
        if let Ok(value) = result {
            let type_ = hpid.get_type("");
            if let Ok(type_) = type_ {
                if type_ == Type::String {
                    let value_hstring = HSTRING::from(value);
                    match unsafe { CLSIDFromString(&value_hstring) } {
                        Ok(guid) => Ok(guid),
                        Err(error) => Err(OleError::runtime(
                            error,
                            format!("unknown OLE server: `{}`", value_hstring.display()),
                        )
                        .into()),
                    }
                } else {
                    unreachable!()
                }
            } else {
                Err(type_.unwrap_err().into())
            }
        } else {
            Err(result.unwrap_err().into())
        }
    }
}

fn ole_create_dcom<S: AsRef<str>>(ole: S, host: S) -> Result<OleData> {
    let host = host.as_ref();
    let ole = ole.as_ref();

    let clsctx = CLSCTX_REMOTE_SERVER;
    let ole_hstring = HSTRING::from(ole);
    let clsid = match unsafe { CLSIDFromProgID(&ole_hstring) } {
        Ok(clsid) => Ok(clsid),
        Err(_) => match clsid_from_remote(host, ole) {
            Ok(clsid) => Ok(clsid),
            Err(_) => unsafe { CLSIDFromString(&ole_hstring) },
        },
    };

    if let Err(error) = clsid {
        return Err(OleError::runtime(error, format!("unknown OLE server: `{ole}`")).into());
    }

    let clsid = clsid.unwrap();
    let mut host_vec = OsStr::new(host)
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let serverinfo = COSERVERINFO {
        pwszName: PWSTR::from_raw(host_vec.as_mut_ptr()),
        ..Default::default()
    };
    let multi_qi = MULTI_QI {
        pIID: &IDispatch::IID,
        ..Default::default()
    };
    let mut multi_qi_arr = vec![multi_qi; 1];
    let result =
        unsafe { CoCreateInstanceEx(&clsid, None, clsctx, Some(&serverinfo), &mut multi_qi_arr) };
    if let Err(error) = result {
        return Err(OleError::runtime(
            error,
            format!("failed to create DCOM server `{ole}` in `{host}`"),
        )
        .into());
    }

    let multi_qi = multi_qi_arr.pop().unwrap();
    Ok(OleData {
        dispatch: multi_qi.pItf.as_ref().unwrap().cast::<IDispatch>().unwrap(),
    })
}

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
        let obj = super::OleData::new("Scripting.Dictionary", None, None);
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

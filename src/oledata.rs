use std::{ffi::OsStr, ptr};

use windows::{
    core::{Interface, Vtable, BSTR, GUID, PCWSTR},
    Win32::{
        Globalization::GetUserDefaultLCID,
        System::Com::{IDispatch, ITypeInfo, ITypeLib},
    },
};

use crate::{
    error::{Error, Result},
    olemethoddata::{ole_methods_from_typeinfo, OleMethodData},
    util::{
        conv::ToWide,
        ole::{create_com_object, get_class_id},
    },
    OleTypeData, OleTypeLibData,
};

pub struct OleData {
    dispatch: IDispatch,
}
impl OleData {
    pub fn new<S: AsRef<OsStr>>(prog_id: S) -> Result<Self> {
        Ok(OleData {
            dispatch: create_com_object(prog_id)?,
        })
    }
    pub fn get_ids_of_names<S: AsRef<OsStr> + Copy>(&self, names: &[S]) -> Result<Vec<i32>> {
        let namelen = names.len();
        let mut wnames = vec![PCWSTR::null(); namelen];
        for i in 0..namelen {
            let a = names[i].to_wide_null();
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
    pub fn responds_to<S: AsRef<OsStr>>(&self, method: S) -> bool {
        let method = method.to_wide_null();
        let methods = vec![PCWSTR(method.as_ptr())];
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
    pub fn ole_type(&self) -> Result<OleTypeData> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };
        OleTypeData::from_itypeinfo(&typeinfo)
    }
    pub fn ole_typelib(&self) -> Result<OleTypeLibData> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };
        OleTypeLibData::from_itypeinfo(&typeinfo)
    }
    pub fn ole_methods(&self, mask: i32) -> Result<Vec<OleMethodData>> {
        let mut methods = vec![];

        let typeinfo = self.typeinfo_from_ole()?;
        methods.extend(ole_methods_from_typeinfo(&typeinfo, mask)?);
        Ok(methods)
    }
    fn typeinfo_from_ole(&self) -> Result<ITypeInfo> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };

        let mut bstrname = BSTR::default();
        unsafe { typeinfo.GetDocumentation(-1, Some(&mut bstrname), None, ptr::null_mut(), None)? };
        let type_ = bstrname;
        let mut typelib: Option<ITypeLib> = None;
        let mut i = 0;
        unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut i)? };
        let typelib = typelib.unwrap();
        let count = unsafe { typelib.GetTypeInfoCount() };

        let mut ret_type_info = None;
        for i in 0..count {
            let mut bstrname = BSTR::default();
            let result = unsafe {
                typelib.GetDocumentation(i as i32, Some(&mut bstrname), None, ptr::null_mut(), None)
            };
            if result.is_ok() && bstrname == type_ {
                let result = unsafe { typelib.GetTypeInfo(i) };
                if let Ok(ret_type) = result {
                    ret_type_info = Some(ret_type);
                    break;
                }
            }
        }
        Ok(ret_type_info.unwrap())
    }
    pub fn ole_query_interface<S: AsRef<OsStr>>(&self, str_iid: S) -> Result<OleData> {
        let iid = get_class_id(str_iid)?;
        let mut dispatch_interface = ptr::null();
        println!("dispatch_interface is {dispatch_interface:p}");
        let result = unsafe { self.dispatch.query(&iid, &mut dispatch_interface) };
        let result = result.ok();
        println!("result is {result:?}, dispatch_interface is {dispatch_interface:p}",);
        if let Err(error) = result {
            Err(error.into())
        } else {
            let dispatch: IDispatch =
                unsafe { <IDispatch as Vtable>::from_raw(dispatch_interface as *mut _) };
            println!("This worked");
            Ok(OleData { dispatch })
        }
    }
    pub fn ole_method_help<S: AsRef<OsStr>>(&self, cmdname: S) -> Result<OleMethodData> {
        let typeinfo = self.typeinfo_from_ole();
        let Ok(typeinfo) = typeinfo else {
            return Err(Error::Custom(format!("failed to get ITypeInfo: {}", typeinfo.err().unwrap())));
        };
        let obj = OleMethodData::from_typeinfo(&typeinfo, &cmdname)?;

        if let Some(obj) = obj {
            Ok(obj)
        } else {
            Err(Error::Custom(format!(
                "not found {}",
                cmdname.as_ref().to_str().unwrap()
            )))
        }
    }
}

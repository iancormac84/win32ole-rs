use std::{ffi::OsStr, ptr};

use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Globalization::GetUserDefaultLCID,
        System::Com::{IDispatch, ITypeInfo, ITypeLib},
    },
};

use crate::{
    error::Result,
    olemethoddata::{ole_methods_from_typeinfo, OleMethodData},
    util::{conv::ToWide, ole::create_com_object},
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
    pub fn typelib(&self) -> Result<ITypeLib> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };
        let mut tlib: Option<ITypeLib> = None;
        let mut index = 0;
        unsafe { typeinfo.GetContainingTypeLib(&mut tlib, &mut index)? }

        Ok(tlib.unwrap())
    }
    pub fn ole_methods(&self, mask: i32) -> Result<Vec<OleMethodData>> {
        let mut methods = vec![];

        let typeinfo = self.typeinfo_from_ole()?;
        if let Some(typeinfo) = typeinfo {
            methods.extend(ole_methods_from_typeinfo(&typeinfo, mask)?);
        }
        Ok(methods)
    }
    fn typeinfo_from_ole(&self) -> Result<Option<ITypeInfo>> {
        let typeinfo = unsafe { self.dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };

        let mut bstrname = BSTR::default();
        unsafe { typeinfo.GetDocumentation(-1, Some(&mut bstrname), None, ptr::null_mut(), None)? };
        let type_ = bstrname;
        let mut typelib: Option<ITypeLib> = None;
        let mut i = 0;
        unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut i)? };
        let typelib = typelib.unwrap();
        let count = unsafe { typelib.GetTypeInfoCount() };
        let mut i = 0;
        let ret_typeinfo = loop {
            if i == count {
                break None;
            }
            let mut bstrname = BSTR::default();
            let result = unsafe {
                typelib.GetDocumentation(i as i32, Some(&mut bstrname), None, ptr::null_mut(), None)
            };
            if result.is_ok() && bstrname == type_ {
                let result = unsafe { typelib.GetTypeInfo(i) };
                if let Ok(ret_typeinfo) = result {
                    break Some(ret_typeinfo);
                }
            }
            i += 1;
        };
        Ok(ret_typeinfo)
    }
}

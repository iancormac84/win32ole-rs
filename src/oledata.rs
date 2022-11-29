use std::{ffi::OsStr, ptr};

use windows::{
    core::{BSTR, GUID, PWSTR},
    Win32::{
        Globalization::GetUserDefaultLCID,
        System::Com::{IDispatch, ITypeInfo, ITypeLib},
    },
};

use crate::{
    error::Result,
    olemethoddata::{ole_methods_from_typeinfo, OleMethodData},
    util::{create_com_object, to_u16s},
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
        let mut wnames = vec![PWSTR::null(); namelen];
        for i in 0..namelen {
            let mut a = to_u16s(names[i])?;
            wnames[i] = PWSTR(a.as_mut_ptr());
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
        let Ok(mut method) = to_u16s(method) else {
            return false;
        };
        let methods = vec![PWSTR(method.as_mut_ptr())];
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

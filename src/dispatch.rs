use std::ffi::OsStr;

use crate::{
    error::{ComArgumentErrorType, Error, Result},
    ToWide,
};
use windows::{
    core::{GUID, PCWSTR, VARIANT},
    Win32::{
        Foundation::{DISP_E_EXCEPTION, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH},
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{
                IDispatch, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
                DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::DISPID_PROPERTYPUT,
        },
    },
};

pub trait IDispatchExt {
    fn get(&self, name: &str) -> Result<VARIANT>;
    fn put(&self, name: &str, value: &mut VARIANT) -> Result<()>;
    fn call(&self, name: &str, args: Vec<VARIANT>) -> Result<VARIANT>;
}

fn invoke<S: AsRef<OsStr>>(
    obj: &IDispatch,
    name: S,
    dp: &mut DISPPARAMS,
    flags: DISPATCH_FLAGS,
) -> Result<VARIANT> {
    let name = PCWSTR::from_raw(name.as_ref().to_wide_null().as_ptr());
    let mut id = 0i32;
    unsafe {
        obj.GetIDsOfNames(&GUID::zeroed(), &name, 1, GetUserDefaultLCID(), &mut id)?;
    }

    let mut excep = EXCEPINFO::default();
    let mut arg_err = 0;
    let mut result = VARIANT::default();

    let res = unsafe {
        obj.Invoke(
            id,
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

impl IDispatchExt for IDispatch {
    /// Get a property from a COM object
    ///
    /// Note: consider using the [`get!`] macro
    fn get(&self, name: &str) -> Result<VARIANT> {
        let mut dp = DISPPARAMS::default();
        invoke(self, name, &mut dp, DISPATCH_PROPERTYGET)
    }

    /// Set a property on a COM object
    ///
    /// Note: consider using the [`put!`] macro
    fn put(&self, name: &str, value: &mut VARIANT) -> Result<()> {
        let mut dp = DISPPARAMS {
            cArgs: 1,
            rgvarg: value,
            cNamedArgs: 1,
            ..Default::default()
        };
        let mut id = DISPID_PROPERTYPUT;
        dp.rgdispidNamedArgs = &mut id as *mut _;
        invoke(self, name, &mut dp, DISPATCH_PROPERTYPUT)?;
        Ok(())
    }

    /// Call a method on a COM object
    ///
    fn call(&self, name: &str, args: Vec<VARIANT>) -> Result<VARIANT> {
        let mut dp = DISPPARAMS::default();
        let args: Vec<VARIANT> = args.into_iter().rev().collect();
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_ptr() as *mut _;
        invoke(self, name, &mut dp, DISPATCH_METHOD)
    }
}

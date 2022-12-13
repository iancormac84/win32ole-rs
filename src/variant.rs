use std::{
    ffi::{CStr, CString, OsStr, OsString},
    mem::ManuallyDrop,
};
use windows::{
    core::{IUnknown, BSTR, HRESULT, PSTR},
    Win32::{
        Foundation::{CHAR, VARIANT_BOOL, VARIANT_FALSE, VARIANT_TRUE},
        System::{
            Com::{
                IDispatch, CY, SAFEARRAY, VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
                VT_ARRAY, VT_BOOL, VT_BSTR, VT_BYREF, VT_CY, VT_DATE, VT_DECIMAL, VT_DISPATCH,
                VT_EMPTY, VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_LPSTR, VT_NULL, VT_R4,
                VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN, VT_VARIANT,
            },
            Ole::{SafeArrayCreateVector, SafeArrayPutElement},
        },
    },
};

use crate::{error::Error, Result, ToWide};

const VT_PUI1: VARENUM = VARENUM(VT_BYREF.0 | VT_UI1.0);
const VT_PI2: VARENUM = VARENUM(VT_BYREF.0 | VT_I2.0);
const VT_PI4: VARENUM = VARENUM(VT_BYREF.0 | VT_I4.0);
const VT_PI8: VARENUM = VARENUM(VT_BYREF.0 | VT_I8.0);
const VT_PUI8: VARENUM = VARENUM(VT_BYREF.0 | VT_UI8.0);
const VT_PR4: VARENUM = VARENUM(VT_BYREF.0 | VT_R4.0);
const VT_PR8: VARENUM = VARENUM(VT_BYREF.0 | VT_R8.0);
const VT_PBOOL: VARENUM = VARENUM(VT_BYREF.0 | VT_BOOL.0);
const VT_PERROR: VARENUM = VARENUM(VT_BYREF.0 | VT_ERROR.0);
const VT_PCY: VARENUM = VARENUM(VT_BYREF.0 | VT_CY.0);
const VT_PDATE: VARENUM = VARENUM(VT_BYREF.0 | VT_DATE.0);
const VT_PBSTR: VARENUM = VARENUM(VT_BYREF.0 | VT_BSTR.0);
const VT_PUNKNOWN: VARENUM = VARENUM(VT_BYREF.0 | VT_UNKNOWN.0);
const VT_PDISPATCH: VARENUM = VARENUM(VT_BYREF.0 | VT_DISPATCH.0);
const VT_PARRAY: VARENUM = VARENUM(VT_BYREF.0 | VT_ARRAY.0);
const VT_PDECIMAL: VARENUM = VARENUM(VT_BYREF.0 | VT_DECIMAL.0);
const VT_PI1: VARENUM = VARENUM(VT_BYREF.0 | VT_I1.0);
const VT_PUI2: VARENUM = VARENUM(VT_BYREF.0 | VT_UI2.0);
const VT_PUI4: VARENUM = VARENUM(VT_BYREF.0 | VT_UI4.0);
const VT_PINT: VARENUM = VARENUM(VT_BYREF.0 | VT_INT.0);
const VT_PUINT: VARENUM = VARENUM(VT_BYREF.0 | VT_UINT.0);

pub(crate) struct VariantFactory(VARENUM, VARIANT_0_0_0);

impl From<VariantFactory> for VARIANT {
    fn from(factory: VariantFactory) -> Self {
        let VariantFactory(vt, value) = factory;
        Self {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: value,
                }),
            },
        }
    }
}

/*pub union VARIANT_0_0_0 {
    //pub llVal: i64,
    //pub lVal: i32,
    //pub bVal: u8,
    //pub iVal: i16,
    //pub fltVal: f32,
    //pub dblVal: f64,
    //pub boolVal: super::super::Foundation::VARIANT_BOOL,
    pub __OBSOLETE__VARIANT_BOOL: super::super::Foundation::VARIANT_BOOL,
    //pub scode: i32,
    //pub cyVal: CY,
    pub date: f64,
    //pub bstrVal: ::std::mem::ManuallyDrop<::windows::core::BSTR>,
    //pub punkVal: ::std::mem::ManuallyDrop<::core::option::Option<::windows::core::IUnknown>>,
    //pub pdispVal: ::std::mem::ManuallyDrop<::core::option::Option<IDispatch>>,
    pub parray: *mut SAFEARRAY,
    pub pbVal: *mut u8,
    pub piVal: *mut i16,
    pub plVal: *mut i32,
    pub pllVal: *mut i64,
    pub pfltVal: *mut f32,
    pub pdblVal: *mut f64,
    pub pboolVal: *mut super::super::Foundation::VARIANT_BOOL,
    pub __OBSOLETE__VARIANT_PBOOL: *mut super::super::Foundation::VARIANT_BOOL,
    pub pscode: *mut i32,
    pub pcyVal: *mut CY,
    pub pdate: *mut f64,
    pub pbstrVal: *mut ::windows::core::BSTR,
    pub ppunkVal: *mut ::core::option::Option<::windows::core::IUnknown>,
    pub ppdispVal: *mut ::core::option::Option<IDispatch>,
    pub pparray: *mut *mut SAFEARRAY,
    pub pvarVal: *mut VARIANT,
    pub byref: *mut ::core::ffi::c_void,
    //pub cVal: super::super::Foundation::CHAR,
    //pub uiVal: u16,
    //pub ulVal: u32,
    //pub ullVal: u64,
    //pub intVal: i32,
    //pub uintVal: u32,
    pub pdecVal: *mut super::super::Foundation::DECIMAL,
    //pub pcVal: ::windows::core::PSTR,
    pub puiVal: *mut u16,
    pub pulVal: *mut u32,
    pub pullVal: *mut u64,
    pub pintVal: *mut i32,
    pub puintVal: *mut u32,
    pub Anonymous: ::std::mem::ManuallyDrop<VARIANT_0_0_0_0>,
}*/

impl From<i64> for VariantFactory {
    fn from(value: i64) -> Self {
        Self(VT_I8, VARIANT_0_0_0 { llVal: value })
    }
}

impl From<i32> for VariantFactory {
    fn from(value: i32) -> Self {
        Self(VT_I4, VARIANT_0_0_0 { lVal: value })
    }
}

impl From<u8> for VariantFactory {
    fn from(value: u8) -> Self {
        Self(VT_UI1, VARIANT_0_0_0 { bVal: value })
    }
}

impl From<i16> for VariantFactory {
    fn from(value: i16) -> Self {
        Self(VT_I2, VARIANT_0_0_0 { iVal: value })
    }
}

impl From<f32> for VariantFactory {
    fn from(value: f32) -> Self {
        Self(VT_R4, VARIANT_0_0_0 { fltVal: value })
    }
}

impl From<f64> for VariantFactory {
    fn from(value: f64) -> Self {
        Self(VT_R8, VARIANT_0_0_0 { dblVal: value })
    }
}

impl From<bool> for VariantFactory {
    fn from(value: bool) -> Self {
        Self(
            VT_BOOL,
            VARIANT_0_0_0 {
                boolVal: if value { VARIANT_TRUE } else { VARIANT_FALSE },
            },
        )
    }
}

#[allow(non_snake_case)]
impl From<HRESULT> for VariantFactory {
    fn from(value: HRESULT) -> Self {
        Self(VT_ERROR, VARIANT_0_0_0 { scode: value.0 })
    }
}

impl From<CY> for VariantFactory {
    fn from(value: CY) -> Self {
        Self(VT_CY, VARIANT_0_0_0 { cyVal: value })
    }
}

#[allow(non_snake_case)]
impl From<(VARENUM, f64)> for VariantFactory {
    fn from(value: (VARENUM, f64)) -> Self {
        match value.0 .0 {
            7 => Self(VT_DATE, VARIANT_0_0_0 { date: value.1 }), //VT_DATE
            3 => Self(VT_R8, VARIANT_0_0_0 { dblVal: value.1 }), //VT_R8
            _ => panic!(
                "Incompatible VARENUM type {:?} coupled with variant type represented as f64.",
                value.0
            ),
        }
    }
}

impl From<&str> for VariantFactory {
    fn from(value: &str) -> Self {
        let value: BSTR = value.into();
        value.into()
    }
}

impl From<String> for VariantFactory {
    fn from(value: String) -> Self {
        let value: BSTR = value.into();
        value.into()
    }
}

impl From<&OsStr> for VariantFactory {
    fn from(value: &OsStr) -> Self {
        let value_wide = value.to_wide_null();
        let value_bstr = BSTR::from_wide(&value_wide);
        value_bstr.into()
    }
}

impl From<OsString> for VariantFactory {
    fn from(value: OsString) -> Self {
        let value_wide = value.to_wide_null();
        let value_bstr = BSTR::from_wide(&value_wide);
        value_bstr.into()
    }
}

impl TryFrom<&CStr> for VariantFactory {
    type Error = Error;

    fn try_from(value: &CStr) -> Result<Self> {
        Ok(value.to_str()?.into())
    }
}

impl TryFrom<CString> for VariantFactory {
    type Error = Error;

    fn try_from(value: CString) -> Result<Self> {
        Ok(value.into_string()?.into())
    }
}

impl From<BSTR> for VariantFactory {
    fn from(value: BSTR) -> Self {
        Self(
            VT_BSTR,
            VARIANT_0_0_0 {
                bstrVal: ManuallyDrop::new(value),
            },
        )
    }
}

impl From<IUnknown> for VariantFactory {
    fn from(value: IUnknown) -> Self {
        Self(
            VT_UNKNOWN,
            VARIANT_0_0_0 {
                punkVal: ManuallyDrop::new(Some(value)),
            },
        )
    }
}

impl From<IDispatch> for VariantFactory {
    fn from(value: IDispatch) -> Self {
        Self(
            VT_DISPATCH,
            VARIANT_0_0_0 {
                pdispVal: ManuallyDrop::new(Some(value)),
            },
        )
    }
}

impl From<*mut SAFEARRAY> for VariantFactory {
    fn from(value: *mut SAFEARRAY) -> Self {
        Self(VT_ARRAY, VARIANT_0_0_0 { parray: value })
    }
}

impl From<*mut u8> for VariantFactory {
    fn from(value: *mut u8) -> Self {
        Self(VT_PUI1, VARIANT_0_0_0 { pbVal: value })
    }
}

impl From<*mut i16> for VariantFactory {
    fn from(value: *mut i16) -> Self {
        Self(VT_PI2, VARIANT_0_0_0 { piVal: value })
    }
}

impl From<*mut i32> for VariantFactory {
    fn from(value: *mut i32) -> Self {
        Self(VT_PI4, VARIANT_0_0_0 { plVal: value })
    }
}

impl From<*mut i64> for VariantFactory {
    fn from(value: *mut i64) -> Self {
        Self(VT_PI8, VARIANT_0_0_0 { pllVal: value })
    }
}

impl From<*mut f32> for VariantFactory {
    fn from(value: *mut f32) -> Self {
        Self(VT_PR4, VARIANT_0_0_0 { pfltVal: value })
    }
}

impl From<(VARENUM, *mut f64)> for VariantFactory {
    fn from(value: (VARENUM, *mut f64)) -> Self {
        match value.0 .0 {
            16391 => Self(VT_PDATE, VARIANT_0_0_0 { pdate: value.1 }),
            16389 => Self(VT_PR8, VARIANT_0_0_0 { pdblVal: value.1 }),
            _ => panic!(
                "Incompatible VARENUM type {:?} coupled with variant type represented as *mut f64.",
                value.0
            ),
        }
    }
}

impl From<*mut VARIANT_BOOL> for VariantFactory {
    fn from(value: *mut VARIANT_BOOL) -> Self {
        Self(VT_PBOOL, VARIANT_0_0_0 { pboolVal: value })
    }
}

impl From<*mut HRESULT> for VariantFactory {
    fn from(value: *mut HRESULT) -> Self {
        Self(
            VT_PERROR,
            VARIANT_0_0_0 {
                pscode: &mut unsafe { (*value).0 },
            },
        )
    }
}

impl From<*mut BSTR> for VariantFactory {
    fn from(value: *mut BSTR) -> Self {
        Self(VT_PBSTR, VARIANT_0_0_0 { pbstrVal: value })
    }
}

impl From<*mut Option<IUnknown>> for VariantFactory {
    fn from(value: *mut Option<IUnknown>) -> Self {
        Self(VT_PUNKNOWN, VARIANT_0_0_0 { ppunkVal: value })
    }
}

impl From<*mut Option<IDispatch>> for VariantFactory {
    fn from(value: *mut Option<IDispatch>) -> Self {
        Self(VT_PDISPATCH, VARIANT_0_0_0 { ppdispVal: value })
    }
}

impl From<*mut *mut SAFEARRAY> for VariantFactory {
    fn from(value: *mut *mut SAFEARRAY) -> Self {
        Self(VT_PARRAY, VARIANT_0_0_0 { pparray: value })
    }
}

impl From<*mut CY> for VariantFactory {
    fn from(value: *mut CY) -> Self {
        Self(VT_PCY, VARIANT_0_0_0 { pcyVal: value })
    }
}

impl From<*mut VARIANT> for VariantFactory {
    fn from(value: *mut VARIANT) -> Self {
        Self(VT_VARIANT, VARIANT_0_0_0 { pvarVal: value })
    }
}

impl From<*mut ::core::ffi::c_void> for VariantFactory {
    fn from(value: *mut ::core::ffi::c_void) -> Self {
        Self(VT_BYREF, VARIANT_0_0_0 { byref: value })
    }
}

impl From<CHAR> for VariantFactory {
    fn from(value: CHAR) -> Self {
        Self(VT_UI1, VARIANT_0_0_0 { cVal: value })
    }
}

impl From<u16> for VariantFactory {
    fn from(value: u16) -> Self {
        Self(VT_UI2, VARIANT_0_0_0 { uiVal: value })
    }
}

impl From<u32> for VariantFactory {
    fn from(value: u32) -> Self {
        Self(VT_UI4, VARIANT_0_0_0 { ulVal: value })
    }
}

impl From<u64> for VariantFactory {
    fn from(value: u64) -> Self {
        Self(VT_UI4, VARIANT_0_0_0 { ullVal: value })
    }
}

impl From<PSTR> for VariantFactory {
    fn from(value: PSTR) -> Self {
        Self(VT_LPSTR, VARIANT_0_0_0 { pcVal: value })
    }
}

impl From<*mut u16> for VariantFactory {
    fn from(value: *mut u16) -> Self {
        Self(VT_PUI2, VARIANT_0_0_0 { puiVal: value })
    }
}

impl From<*mut u32> for VariantFactory {
    fn from(value: *mut u32) -> Self {
        Self(VT_PUI4, VARIANT_0_0_0 { pulVal: value })
    }
}

impl From<*mut u64> for VariantFactory {
    fn from(value: *mut u64) -> Self {
        Self(VT_PUI8, VARIANT_0_0_0 { pullVal: value })
    }
}

fn safe_array_from_primitive_slice<T: Copy>(vt: VARENUM, slice: &[T]) -> Result<*mut SAFEARRAY> {
    let v0 = vt.0;
    if v0 & VT_ARRAY.0 == 1 || v0 & VT_BYREF.0 == 1 || v0 == VT_EMPTY.0 || v0 == VT_NULL.0 {
        return Err(Error::Generic(
            "Invalid slice contents for creation of SAFEARRAY",
        ));
    }
    let sa = unsafe { SafeArrayCreateVector(vt, 0, slice.len().try_into()?) };
    if sa.is_null() {
        return Err(Error::Generic("SAFEARRAY allocation failed"));
    }
    for (i, item) in slice.iter().enumerate() {
        let i: i32 = i.try_into()?;
        unsafe { SafeArrayPutElement(&*sa, &i, (item as *const T) as *const _) }?;
    }
    Ok(sa)
}

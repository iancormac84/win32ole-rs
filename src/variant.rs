use std::{mem::ManuallyDrop};
use thiserror::Error;
use windows::{
    core::{IUnknown, BSTR, HRESULT},
    Win32::{
        Foundation::{VARIANT_BOOL, CHAR},
        System::{Com::{
            IDispatch, SAFEARRAY, VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
            VT_ARRAY, VT_BOOL, VT_BSTR, VT_BYREF, VT_CY, VT_DATE, VT_DECIMAL, VT_DISPATCH,
            VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4,
            VT_UI8, VT_UINT, VT_UNKNOWN, VT_VARIANT, CY, VARIANT_0_0_0_0, VT_EMPTY, VT_NULL,
        }, Ole::{SafeArrayCreateVector, SafeArrayPutElement}},
    },
};

use crate::{Error, Result};

/// Encapsulates the ways converting from a `VARIANT` can fail.
#[derive(Copy, Clone, Debug, Error)]
pub enum FromVariantError {
    /// `VARIANT` pointer during conversion was null
    #[error("VARIANT pointer is null")]
    VariantPtrNull,
    /// Unknown VT for
    #[error("VARIANT cannot be this vartype: {0:p}")]
    UnknownVarType(u16),
}

/// Helper type for the OLE/COM+ type DATE
#[derive(Debug, Clone, Copy, PartialOrd, PartialEq)]
pub struct Date(f64); //DATE <--> F64

impl AsRef<f64> for Date {
    fn as_ref(&self) -> &f64 {
        &self.0
    }
}
impl From<f64> for Date {
    fn from(i: f64) -> Self {
        Date(i)
    }
}
impl<'f> From<&'f f64> for Date {
    fn from(i: &'f f64) -> Self {
        Date(*i)
    }
}
impl<'f> From<&'f mut f64> for Date {
    fn from(i: &'f mut f64) -> Self {
        Date(*i)
    }
}
impl From<Date> for f64 {
    fn from(o: Date) -> Self {
        o.0
    }
}
impl<'f> From<&'f Date> for f64 {
    fn from(o: &'f Date) -> Self {
        o.0
    }
}
impl<'f> From<&'f mut Date> for f64 {
    fn from(o: &'f mut Date) -> Self {
        o.0
    }
}
impl TryFrom<f64> for Date {
    type Error = FromVariantError;

    /// Does not return any errors.
    fn try_from(val: f64) -> Result<Self, FromVariantError> {
        Ok(Date::from(val))
    }
}
impl TryFrom<Date, IntoVariantError> for f64 {
    /// Does not return any errors.
    fn try_from(val: Date) -> Result<Self, IntoVariantError> {
        Ok(f64::from(val))
    }
}
impl<'c> TryFrom<&'c f64> for Date {
    type Error = FromVariantError;

    /// Does not return any errors.
    fn try_from(val: &'c f64) -> Result<Self, FromVariantError> {
        Ok(Date::from(val))
    }
}
impl<'c> TryFrom<&'c Date, IntoVariantError> for f64 {
    /// Does not return any errors.
    fn try_from(val: &'c Date) -> Result<Self, IntoVariantError> {
        Ok(f64::from(val))
    }
}
impl<'c> TryFrom<&'c mut f64> for Date {
    type Error = FromVariantError;

    /// Does not return any errors.
    fn try_from(val: &'c mut f64) -> Result<Self, FromVariantError> {
        Ok(Date::from(val))
    }
}
impl<'c> TryFrom<&'c mut Date, IntoVariantError> for f64 {
    /// Does not return any errors.
    fn try_from(val: &'c mut Date) -> Result<Self, IntoVariantError> {
        Ok(f64::from(val))
    }
}
impl TryFrom<f64, SafeArrayError> for Date {
    /// Does not return any errors.
    fn try_from(val: f64) -> Result<Self, SafeArrayError> {
        Ok(Date::from(val))
    }
}
impl TryFrom<Date, SafeArrayError> for f64 {
    /// Does not return any errors.
    fn try_from(val: Date) -> Result<Self, SafeArrayError> {
        Ok(f64::from(val))
    }
}
impl TryFrom<f64, ElementError> for Date {
    /// Does not return any errors.
    fn try_from(val: f64) -> Result<Self, ElementError> {
        Ok(Date::from(val))
    }
}
impl TryFrom<Date, ElementError> for f64 {
    /// Does not return any errors.
    fn try_from(val: Date) -> Result<Self, ElementError> {
        Ok(f64::from(val))
    }
}
impl TryFrom<Box<Date>> for *mut f64 {
    fn try_from(b: Box<Date>) -> Result<Self, IntoVariantError> {
        let b = *b;
        let inner = f64::from(b);
        Ok(Box::into_raw(Box::new(inner)))
    }
}
impl TryFrom<*mut f64> for Box<Date> {
    fn try_from(inner: *mut f64) -> Result<Self, FromVariantError> {
        if inner.is_null() {
            return Err(FromVariantError::VariantPtrNull);
        }
        let inner = unsafe { *inner };
        let wrapper = Date::from(inner);
        Ok(Box::new(wrapper))
    }
}

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
                    vt: VARENUM(vt.0 as u16),
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: value,
                }),
            },
        }
    }
}

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

const VARIANT_FALSE: i16 = 0i16;
const VARIANT_TRUE: i16 = -1i16;

impl From<bool> for VariantFactory {
    fn from(value: bool) -> Self {
        Self(
            VT_BOOL,
            VARIANT_0_0_0 {
                boolVal: if value {
                    VARIANT_BOOL(VARIANT_TRUE)
                } else {
                    VARIANT_BOOL(VARIANT_FALSE)
                },
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
        Self(VT_PERROR, VARIANT_0_0_0 { pscode: &mut unsafe { (*value).0 } })
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

impl From<VARIANT_0_0_0_0> for VariantFactory {
    fn from(value: VARIANT_0_0_0_0) -> Self {
        Self((), ())
    }
}

fn safe_array_from_primitive_slice<T: Copy>(vt: VARENUM, slice: &[T]) -> Result<*mut SAFEARRAY> {
    let v0 = vt.0;
    if v0 & VT_ARRAY.0 == 1 || v0 & VT_BYREF == 1 || v0 == VT_EMPTY || v0 == VT_NULL {
        return Err(Error::Generic("Invalid slice contents for creation of SAFEARRAY"));
    }
    let sa =
        unsafe { SafeArrayCreateVector(VARENUM(vt.0 as u16), 0, slice.len().try_into().unwrap()) };
    if sa.is_null() {
        return Err(Error::Generic("SAFEARRAY allocation failed"));
    }
    for (i, item) in slice.iter().enumerate() {
        let i: i32 = i.try_into().unwrap();
        unsafe { SafeArrayPutElement(&*sa, &i, (item as *const T) as *const _) }.unwrap();
    }
    Ok(sa)
}

/*pub union VARIANT_0_0_0 {
    pub intVal: i32,
    pub pdecVal: *mut super::super::Foundation::DECIMAL,
    pub pcVal: ::windows::core::PSTR,
    pub puintVal: *mut u32,
    pub Anonymous: ::core::mem::ManuallyDrop<VARIANT_0_0_0_0>,
}*/

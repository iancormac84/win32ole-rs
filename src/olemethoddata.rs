use crate::{
    error::{Error, Result},
    oleparamdata::OleParamData,
    util::{conv::ToWide, ole::ole_typedesc2val},
    OleTypeData,
};
use std::{ffi::OsStr, ptr};
use windows::{
    core::{BSTR, PCWSTR},
    Win32::System::Com::{
        ITypeInfo, IMPLTYPEFLAGS, IMPLTYPEFLAG_FSOURCE, INVOKEKIND, INVOKE_FUNC,
        INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_COCLASS, VARENUM,
    },
};

#[derive(Debug)]
pub struct OleMethodData {
    owner_typeinfo: Option<ITypeInfo>,
    pub typeinfo: ITypeInfo,
    name: String,
    pub index: u32,
}

impl OleMethodData {
    pub fn new<S: AsRef<OsStr>>(ole_type: &OleTypeData, name: S) -> Result<Option<OleMethodData>> {
        OleMethodData::from_typeinfo(ole_type.typeinfo(), name)
    }
    pub fn from_typeinfo<S: AsRef<OsStr>>(
        typeinfo: &ITypeInfo,
        name: S,
    ) -> Result<Option<OleMethodData>> {
        let type_attr = unsafe { typeinfo.GetTypeAttr()? };
        let mut method = ole_method_sub(None, typeinfo, &name)?;
        if method.is_some() {
            return Ok(method);
        }
        for i in 0..unsafe { (*type_attr).cImplTypes } {
            if method.is_some() {
                break;
            }
            let result = unsafe { typeinfo.GetRefTypeOfImplType(i as u32) };
            let Ok(href) = result else {
                continue;
            };
            let result = unsafe { typeinfo.GetRefTypeInfo(href) };
            let Ok(ref_typeinfo) = result else {
                continue;
            };
            method = ole_method_sub(Some(typeinfo), &ref_typeinfo, &name)?;
        }
        unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(method)
    }
    fn docinfo_from_type(
        &self,
        name: Option<*mut BSTR>,
        helpstr: Option<*mut BSTR>,
        helpcontext: *mut u32,
        helpfile: Option<*mut BSTR>,
    ) -> Result<()> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        unsafe {
            self.typeinfo.GetDocumentation(
                (*funcdesc).memid,
                name,
                helpstr,
                helpcontext,
                helpfile,
            )?
        };
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(())
    }
    pub fn helpstring(&self) -> Result<String> {
        let mut helpstring = BSTR::default();
        self.docinfo_from_type(None, Some(&mut helpstring), ptr::null_mut(), None)?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        self.docinfo_from_type(None, None, ptr::null_mut(), Some(&mut helpfile))?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self) -> Result<u32> {
        let mut helpcontext = 0;
        self.docinfo_from_type(None, None, &mut helpcontext, None)?;
        Ok(helpcontext)
    }
    pub fn dispid(&self) -> Result<i32> {
        let res = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let dispid = unsafe { (*res).memid };
        unsafe { self.typeinfo.ReleaseFuncDesc(res) };
        Ok(dispid)
    }
    pub fn return_type(&self) -> Result<String> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let type_ = ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { (*funcdesc).elemdescFunc.tdesc }),
            None,
        );
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(type_)
    }
    pub fn return_vtype(&self) -> Result<VARENUM> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let vvt = unsafe { (*funcdesc).elemdescFunc.tdesc.vt };
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(vvt)
    }
    pub fn return_type_detail(&self) -> Result<Vec<String>> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let mut type_details = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { (*funcdesc).elemdescFunc.tdesc }),
            Some(&mut type_details),
        );
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(type_details)
    }
    pub fn invkind(&self) -> Result<INVOKEKIND> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let invkind = unsafe { (*funcdesc).invkind };
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(invkind)
    }
    pub fn invoke_kind(&self) -> Result<String> {
        let invkind = self.invkind()?;
        if invkind.0 & INVOKE_PROPERTYGET.0 != 0 && invkind.0 & INVOKE_PROPERTYPUT.0 != 0 {
            Ok("PROPERTY".into())
        } else if invkind.0 & INVOKE_PROPERTYGET.0 != 0 {
            Ok("PROPERTYGET".into())
        } else if invkind.0 & INVOKE_PROPERTYPUT.0 != 0 {
            Ok("PROPERTYPUT".into())
        } else if invkind.0 & INVOKE_PROPERTYPUTREF.0 != 0 {
            Ok("PROPERTYPUTREF".into())
        } else if invkind.0 & INVOKE_FUNC.0 != 0 {
            Ok("FUNC".into())
        } else {
            Ok("UNKNOWN".into())
        }
    }
    pub fn is_event(&self) -> bool {
        if self.owner_typeinfo.is_none() {
            return false;
        }
        let result = unsafe { self.owner_typeinfo.as_ref().unwrap().GetTypeAttr() };
        if result.is_err() {
            return false;
        }
        let type_attr = result.unwrap();
        if unsafe { (*type_attr).typekind } != TKIND_COCLASS {
            unsafe {
                self.owner_typeinfo
                    .as_ref()
                    .unwrap()
                    .ReleaseTypeAttr(type_attr)
            };
            return false;
        }
        let mut event = false;
        for i in 0..unsafe { (*type_attr).cImplTypes } {
            let result = unsafe {
                self.owner_typeinfo
                    .as_ref()
                    .unwrap()
                    .GetImplTypeFlags(i as u32)
            };
            if result.is_err() {
                continue;
            }

            let flags = result.unwrap();

            if flags & IMPLTYPEFLAG_FSOURCE != IMPLTYPEFLAGS(0) {
                let result = unsafe {
                    self.owner_typeinfo
                        .as_ref()
                        .unwrap()
                        .GetRefTypeOfImplType(i as u32)
                };
                let Ok(href) = result else {
                    continue;
                };
                let result = unsafe { self.owner_typeinfo.as_ref().unwrap().GetRefTypeInfo(href) };
                let Ok(ref_type_info) = result else {
                    continue;
                };
                let result = unsafe { ref_type_info.GetFuncDesc(self.index) };
                let Ok(funcdesc) = result else {
                    continue;
                };
                let mut bstrname = BSTR::default();
                let result = unsafe {
                    ref_type_info.GetDocumentation(
                        (*funcdesc).memid,
                        Some(&mut bstrname),
                        None,
                        ptr::null_mut(),
                        None,
                    )
                };
                if result.is_err() {
                    unsafe { ref_type_info.ReleaseFuncDesc(funcdesc) };
                    continue;
                }

                unsafe { ref_type_info.ReleaseFuncDesc(funcdesc) };
                if self.name == bstrname {
                    event = true;
                    break;
                }
            }
        }
        unsafe {
            self.owner_typeinfo
                .as_ref()
                .unwrap()
                .ReleaseTypeAttr(type_attr)
        };
        event
    }
    pub fn name(&self) -> &str {
        &self.name[..]
    }
    pub fn params(&self) -> Result<Vec<OleParamData>> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index) }?;

        let mut len = 0;
        let mut rgbstrnames = vec![BSTR::default(); unsafe { (*funcdesc).cParams } as usize + 1];
        let result = unsafe {
            self.typeinfo.GetNames(
                (*funcdesc).memid,
                rgbstrnames.as_mut_ptr(),
                (*funcdesc).cParams as u32 + 1,
                &mut len,
            )
        };
        if let Err(error) = result {
            unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
            return Err(Error::Custom(format!(
                "ITypeInfo::GetNames call failed: {error}"
            )));
        }
        let mut params = vec![];

        if unsafe { (*funcdesc).cParams } > 0 {
            for i in 1..len {
                let param = OleParamData::make(
                    self,
                    self.index,
                    i - 1,
                    rgbstrnames[i as usize].to_string(),
                );
                params.push(param);
            }
        }
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(params)
    }
    pub fn offset_vtbl(&self) -> Result<i16> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index) }?;
        let offset_vtbl = unsafe { (*funcdesc).oVft };
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(offset_vtbl)
    }
}

pub(crate) fn ole_methods_from_typeinfo(
    typeinfo: &ITypeInfo,
    mask: i32,
) -> Result<Vec<OleMethodData>> {
    let type_attr = unsafe { typeinfo.GetTypeAttr()? };
    let mut methods = vec![];
    ole_methods_sub(None, typeinfo, &mut methods, mask)?;
    for i in 0..unsafe { (*type_attr).cImplTypes } {
        let reftype = unsafe { typeinfo.GetRefTypeOfImplType(i as u32) };
        if let Ok(reftype) = reftype {
            let reftype_info = unsafe { typeinfo.GetRefTypeInfo(reftype) };
            if let Ok(reftype_info) = reftype_info {
                ole_methods_sub(Some(typeinfo), &reftype_info, &mut methods, mask)?;
            } else {
                continue;
            }
        } else {
            continue;
        }
    }
    unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
    Ok(methods)
}
fn ole_method_sub<S: AsRef<OsStr>>(
    owner_typeinfo: Option<&ITypeInfo>,
    typeinfo: &ITypeInfo,
    name: &S,
) -> Result<Option<OleMethodData>> {
    let type_attr = unsafe { (*typeinfo).GetTypeAttr()? };
    let fname = name.to_wide_null();
    let fname_pcwstr = PCWSTR::from_raw(fname.as_ptr());

    let mut i = 0;
    let num_funcs = unsafe { (*type_attr).cFuncs };
    let mut method = None;
    loop {
        if i == num_funcs {
            break;
        }
        let result = unsafe { (*typeinfo).GetFuncDesc(i as u32) };
        let Ok(funcdesc) = result else {
            continue;
        };

        let mut bstrname = BSTR::default();
        let result = unsafe {
            (*typeinfo).GetDocumentation(
                (*funcdesc).memid,
                Some(&mut bstrname),
                None,
                ptr::null_mut(),
                None,
            )
        };
        if result.is_err() {
            unsafe { (*typeinfo).ReleaseFuncDesc(funcdesc) };
            continue;
        }
        if unsafe { fname_pcwstr.as_wide() } == bstrname.as_wide() {
            method = Some(OleMethodData {
                owner_typeinfo: owner_typeinfo.cloned(),
                typeinfo: typeinfo.clone(),
                name: bstrname.to_string(),
                index: i as u32,
            });
        }
        unsafe { (*typeinfo).ReleaseFuncDesc(funcdesc) };
        i += 1;
    }
    unsafe { (*typeinfo).ReleaseTypeAttr(type_attr) };
    Ok(method)
}
fn ole_methods_sub(
    owner_typeinfo: Option<&ITypeInfo>,
    typeinfo: &ITypeInfo,
    methods: &mut Vec<OleMethodData>,
    mask: i32,
) -> Result<()> {
    let type_attr_ptr = unsafe { (*typeinfo).GetTypeAttr()? };
    for i in 0..unsafe { (*type_attr_ptr).cFuncs } {
        let res = unsafe { (*typeinfo).GetFuncDesc(i as u32) };
        match res {
            Err(_) => continue,
            Ok(func_desc_ptr) => {
                let mut bstrname = BSTR::default();
                let res = unsafe {
                    (*typeinfo).GetDocumentation(
                        (*func_desc_ptr).memid,
                        Some(&mut bstrname),
                        None,
                        ptr::null_mut(),
                        None,
                    )
                };
                if res.is_err() {
                    unsafe { (*typeinfo).ReleaseFuncDesc(func_desc_ptr) };
                    continue;
                }
                if unsafe { (*func_desc_ptr).invkind.0 } & mask != 0 {
                    methods.push(OleMethodData {
                        owner_typeinfo: owner_typeinfo.cloned(),
                        typeinfo: typeinfo.clone(),
                        name: bstrname.to_string(),
                        index: i as u32,
                    });
                }
                unsafe { (*typeinfo).ReleaseFuncDesc(func_desc_ptr) };
            }
        }
    }
    unsafe { (*typeinfo).ReleaseTypeAttr(type_attr_ptr) };
    Ok(())
}

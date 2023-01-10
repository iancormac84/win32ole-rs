use crate::{
    error::Result,
    oleparamdata::OleParamData,
    util::{conv::ToWide, ole::ole_typedesc2val},
    OleTypeData,
};
use std::{
    ffi::OsStr,
    ptr::{self, NonNull},
};
use windows::{
    core::{BSTR, PCWSTR},
    Win32::System::Com::{
        ITypeInfo, FUNCDESC, IMPLTYPEFLAGS, IMPLTYPEFLAG_FSOURCE, INVOKEKIND, INVOKE_FUNC,
        INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_COCLASS, VARENUM,
    },
};

#[derive(Debug)]
pub struct OleMethodData {
    owner_typeinfo: Option<ITypeInfo>,
    typeinfo: ITypeInfo,
    name: String,
    index: u32,
    func_desc: NonNull<FUNCDESC>,
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
        unsafe {
            self.typeinfo.GetDocumentation(
                self.func_desc.as_ref().memid,
                name,
                helpstr,
                helpcontext,
                helpfile,
            )?
        };
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
        Ok(unsafe { self.func_desc.as_ref().memid })
    }
    pub fn return_type(&self) -> Result<String> {
        Ok(ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { self.func_desc.as_ref().elemdescFunc.tdesc }),
            None,
        ))
    }
    pub fn return_vtype(&self) -> Result<VARENUM> {
        Ok(unsafe { self.func_desc.as_ref().elemdescFunc.tdesc.vt })
    }
    pub fn return_type_detail(&self) -> Result<Vec<String>> {
        let mut type_details = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { self.func_desc.as_ref().elemdescFunc.tdesc }),
            Some(&mut type_details),
        );
        Ok(type_details)
    }
    pub fn invkind(&self) -> Result<INVOKEKIND> {
        Ok(unsafe { self.func_desc.as_ref().invkind })
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
    pub fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn index(&self) -> u32 {
        self.index
    }
    pub fn params(&self) -> Vec<Result<OleParamData>> {
        let mut len = 0;
        let cmaxnames = unsafe { self.func_desc.as_ref().cParams } + 1;
        let mut rgbstrnames = vec![BSTR::default(); cmaxnames as usize];
        let result = unsafe {
            self.typeinfo.GetNames(
                self.func_desc.as_ref().memid,
                rgbstrnames.as_mut_ptr(),
                cmaxnames as u32,
                &mut len,
            )
        };
        if result.is_err() {
            return vec![];
        }
        let mut params = vec![];

        if unsafe { self.func_desc.as_ref().cParams } > 0 {
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
        params
    }
    pub fn offset_vtbl(&self) -> Result<i16> {
        Ok(unsafe { self.func_desc.as_ref().oVft })
    }
}

impl Drop for OleMethodData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseFuncDesc(self.func_desc.as_ptr()) };
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
            let func_desc = NonNull::new(funcdesc);
            if let Some(func_desc) = func_desc {
                method = Some(OleMethodData {
                    owner_typeinfo: owner_typeinfo.cloned(),
                    typeinfo: typeinfo.clone(),
                    name: bstrname.to_string(),
                    index: i as u32,
                    func_desc,
                });
            }
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
            Ok(funcdesc) => {
                let mut bstrname = BSTR::default();
                let res = unsafe {
                    (*typeinfo).GetDocumentation(
                        (*funcdesc).memid,
                        Some(&mut bstrname),
                        None,
                        ptr::null_mut(),
                        None,
                    )
                };
                if res.is_err() {
                    unsafe { (*typeinfo).ReleaseFuncDesc(funcdesc) };
                    continue;
                }
                if unsafe { (*funcdesc).invkind.0 } & mask != 0 {
                    let func_desc = NonNull::new(funcdesc);
                    if let Some(func_desc) = func_desc {
                        methods.push(OleMethodData {
                            owner_typeinfo: owner_typeinfo.cloned(),
                            typeinfo: typeinfo.clone(),
                            name: bstrname.to_string(),
                            index: i as u32,
                            func_desc,
                        });
                    }
                }
                unsafe { (*typeinfo).ReleaseFuncDesc(funcdesc) };
            }
        }
    }
    unsafe { (*typeinfo).ReleaseTypeAttr(type_attr_ptr) };
    Ok(())
}

use crate::{
    error::{Error, Result},
    oleparamdata::OleParamData,
    util::{
        conv::ToWide,
        ole::{TypeRef, ValueDescription},
    },
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
        INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_COCLASS, TYPEATTR,
        TYPEDESC, VARENUM,
    },
};

#[derive(Debug)]
pub struct OleMethodData {
    owner_typeinfo: Option<ITypeInfo>,
    owner_type_attr: Option<NonNull<TYPEATTR>>,
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
    pub fn dispid(&self) -> i32 {
        unsafe { self.func_desc.as_ref().memid }
    }
    pub fn return_type(&self) -> String {
        self.ole_typedesc2val(None)
    }
    pub fn return_vtype(&self) -> VARENUM {
        unsafe { self.func_desc.as_ref().elemdescFunc.tdesc.vt }
    }
    pub fn return_type_detail(&self) -> Vec<String> {
        let mut type_details = vec![];
        self.ole_typedesc2val(Some(&mut type_details));
        type_details
    }
    pub fn invkind(&self) -> INVOKEKIND {
        unsafe { self.func_desc.as_ref().invkind }
    }
    pub fn invoke_kind(&self) -> &str {
        let invkind = self.invkind();
        if invkind.0 & INVOKE_PROPERTYGET.0 != 0 && invkind.0 & INVOKE_PROPERTYPUT.0 != 0 {
            "PROPERTY"
        } else if invkind.0 & INVOKE_PROPERTYGET.0 != 0 {
            "PROPERTYGET"
        } else if invkind.0 & INVOKE_PROPERTYPUT.0 != 0 {
            "PROPERTYPUT"
        } else if invkind.0 & INVOKE_PROPERTYPUTREF.0 != 0 {
            "PROPERTYPUTREF"
        } else if invkind.0 & INVOKE_FUNC.0 != 0 {
            "FUNC"
        } else {
            "UNKNOWN"
        }
    }
    pub fn is_event(&self) -> bool {
        if self.owner_typeinfo.is_none() {
            return false;
        }
        if self.owner_type_attr.is_none() {
            return false;
        }
        if unsafe { self.owner_type_attr.unwrap().as_ref().typekind } != TKIND_COCLASS {
            return false;
        }
        let mut event = false;
        for i in 0..unsafe { self.owner_type_attr.unwrap().as_ref().cImplTypes } {
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
        event
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn index(&self) -> u32 {
        self.index
    }
    pub fn params(&self) -> Vec<Result<OleParamData>> {
        let mut len = 0;
        let cparams = unsafe { self.func_desc.as_ref().cParams };
        let cmaxnames = cparams + 1;
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

        if cparams > 0 {
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

impl TypeRef for OleMethodData {
    fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    fn typedesc(&self) -> &TYPEDESC {
        let func_desc_ref = unsafe { self.func_desc.as_ref() };
        &func_desc_ref.elemdescFunc.tdesc
    }
}

impl ValueDescription for OleMethodData {}

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
                let owner_type_attr = if let Some(owner_typeinfo) = owner_typeinfo {
                    let type_attr = unsafe { owner_typeinfo.GetTypeAttr()? };
                    let type_attr = NonNull::new(type_attr).unwrap();
                    Some(type_attr)
                } else {
                    None
                };
                method = Some(OleMethodData {
                    owner_typeinfo: owner_typeinfo.cloned(),
                    owner_type_attr,
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
                        let owner_type_attr = if let Some(owner_typeinfo) = owner_typeinfo {
                            let type_attr = unsafe { owner_typeinfo.GetTypeAttr()? };
                            let type_attr = NonNull::new(type_attr).unwrap();
                            Some(type_attr)
                        } else {
                            None
                        };
                        methods.push(OleMethodData {
                            owner_typeinfo: owner_typeinfo.cloned(),
                            owner_type_attr,
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

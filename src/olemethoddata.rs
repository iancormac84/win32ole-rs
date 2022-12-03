use crate::{
    error::Result,
    util::{ole_typedesc2val, ToWide},
    OleTypeData,
};
use std::{ffi::OsStr, ptr};
use windows::{
    core::{BSTR, PCWSTR},
    Win32::System::Com::{ITypeInfo, VARENUM},
};

#[derive(Debug)]
pub struct OleMethodData {
    owner_typeinfo: Option<ITypeInfo>,
    typeinfo: ITypeInfo,
    name: String,
    index: u32,
}

impl OleMethodData {
    pub fn new<S: AsRef<OsStr>>(ole_type: &OleTypeData, name: S) -> Result<Option<OleMethodData>> {
        OleMethodData::from_typeinfo(&ole_type.typeinfo, name)
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
            println!("{i}");
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
        method_index: u32,
        name: Option<*mut BSTR>,
        helpstr: Option<*mut BSTR>,
        helpcontext: *mut u32,
        helpfile: Option<*mut BSTR>,
    ) -> Result<()> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(method_index)? };
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
    pub fn helpstring(&self, method_index: u32) -> Result<String> {
        let mut helpstring = BSTR::default();
        self.docinfo_from_type(
            method_index,
            None,
            Some(&mut helpstring),
            ptr::null_mut(),
            None,
        )?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self, method_index: u32) -> Result<String> {
        let mut helpfile = BSTR::default();
        self.docinfo_from_type(
            method_index,
            None,
            None,
            ptr::null_mut(),
            Some(&mut helpfile),
        )?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self, method_index: u32) -> Result<u32> {
        let mut helpcontext = 0;
        self.docinfo_from_type(method_index, None, None, &mut helpcontext, None)?;
        Ok(helpcontext)
    }
    pub fn dispid(&self, method_index: u32) -> Result<i32> {
        let res = unsafe { self.typeinfo.GetFuncDesc(method_index)? };
        let dispid = unsafe { (*res).memid };
        unsafe { self.typeinfo.ReleaseFuncDesc(res) };
        Ok(dispid)
    }
    pub fn return_type(&self, method_index: u32) -> Result<String> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(method_index)? };
        let type_ = ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { (*funcdesc).elemdescFunc.tdesc }),
            None,
        );
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(type_)
    }
    pub fn return_vtype(&self, method_index: u32) -> Result<VARENUM> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(method_index)? };
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
    pub fn name(&self) -> &str {
        &self.name[..]
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
    println!("Inside ole_method_sub");
    let type_attr = unsafe { (*typeinfo).GetTypeAttr()? };
    let fname = name.to_wide_null();
    let fname_pcwstr = PCWSTR::from_raw(fname.as_ptr());

    let mut i = 0;
    let num_funcs = unsafe { (*type_attr).cFuncs };
    let mut method = None;
    loop {
        if i == num_funcs {
            println!("Gonna break now!");
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
        println!("bstrname is {bstrname}, and i is {i}");
        if unsafe { fname_pcwstr.as_wide() } == bstrname.as_wide() {
            method = Some(OleMethodData {
                owner_typeinfo: if let Some(typ) = owner_typeinfo {
                    Some(typ.clone())
                } else {
                    None
                },
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
                        owner_typeinfo: if let Some(typ) = owner_typeinfo {
                            Some(typ.clone())
                        } else {
                            None
                        },
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

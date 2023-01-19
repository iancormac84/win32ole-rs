use crate::{
    error::Result,
    oleparamdata::OleParamData,
    types::{Methods, ReferencedTypes},
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
        ITypeInfo, FUNCDESC, FUNCKIND, INVOKEKIND, INVOKE_FUNC, INVOKE_PROPERTYGET,
        INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_COCLASS, TYPEATTR, TYPEDESC, VARENUM,
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
        OleMethodData::from_typeinfo(ole_type.typeinfo().clone(), name)
    }
    pub fn from_typeinfo<S: AsRef<OsStr>>(
        typeinfo: ITypeInfo,
        name: S,
    ) -> Result<Option<OleMethodData>> {
        let type_attr = unsafe { typeinfo.GetTypeAttr()? };
        let method = OleMethodData::maybe_find_and_create(None, &typeinfo, &name)?;
        if method.is_some() {
            return Ok(method);
        }
        let referenced_types = ReferencedTypes::new(&typeinfo, unsafe { &*type_attr }, 0);
        for referenced_type in referenced_types.filter_map(|t| t.ok()) {
            let method = OleMethodData::maybe_find_and_create(
                Some(&typeinfo),
                referenced_type.typeinfo(),
                &name,
            );
            if let Ok(method) = method {
                if method.is_some() {
                    return Ok(method);
                }
            }
        }

        Ok(None)
    }
    fn maybe_find_and_create<S: AsRef<OsStr>>(
        owner_typeinfo: Option<&ITypeInfo>,
        typeinfo: &ITypeInfo,
        name: &S,
    ) -> Result<Option<OleMethodData>> {
        let methods = Methods::new(typeinfo)?;

        let fname = name.to_wide_null();
        let fname_pcwstr = PCWSTR::from_raw(fname.as_ptr());

        for (i, method) in methods.enumerate() {
            if let Ok(method) = method {
                if unsafe { fname_pcwstr.as_wide() } == method.name().as_wide() {
                    let (typeinfo, func_desc, bstrname) = method.deconstruct();

                    let owner_type_attr = if let Some(owner_typeinfo) = owner_typeinfo {
                        let type_attr = unsafe { owner_typeinfo.GetTypeAttr()? };
                        let type_attr = NonNull::new(type_attr).unwrap();
                        Some(type_attr)
                    } else {
                        None
                    };
                    return Ok(Some(OleMethodData {
                        owner_typeinfo: owner_typeinfo.cloned(),
                        owner_type_attr,
                        typeinfo,
                        name: bstrname.to_string(),
                        index: i as u32,
                        func_desc,
                    }));
                }
            }
        }

        Ok(None)
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
    pub fn get_documentation(&self) -> Result<(String, String, u32, String)> {
        let mut strname = BSTR::default();
        let mut strdocstring = BSTR::default();
        let mut whelpcontext = 0;
        let mut strhelpfile = BSTR::default();
        self.docinfo_from_type(
            Some(&mut strname),
            Some(&mut strdocstring),
            &mut whelpcontext,
            Some(&mut strhelpfile),
        )?;
        Ok((
            String::try_from(strname)?,
            String::try_from(strdocstring)?,
            whelpcontext,
            String::try_from(strhelpfile)?,
        ))
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
    pub fn return_type_typedesc(&self) -> &TYPEDESC {
        unsafe { &self.func_desc.as_ref().elemdescFunc.tdesc }
    }
    pub fn return_vtype(&self) -> VARENUM {
        unsafe { self.func_desc.as_ref().elemdescFunc.tdesc.vt }
    }
    pub fn return_type_detail(&self) -> Vec<String> {
        let mut type_details = vec![];
        self.ole_typedesc2val(Some(&mut type_details));
        type_details
    }
    pub fn funckind(&self) -> FUNCKIND {
        unsafe { self.func_desc.as_ref().funckind }
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
        let referenced_types = ReferencedTypes::new(
            self.owner_typeinfo.as_ref().unwrap(),
            unsafe { self.owner_type_attr.unwrap().as_ref() },
            self.index,
        );
        for referenced_type in referenced_types.filter_map(|t| t.ok()) {
            if referenced_type.is_source() {
                let name = referenced_type.name();
                let Ok(name) = name else {
                        continue;
                    };
                if name == self.name {
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
    pub fn event_interface(&self) -> Result<Option<String>> {
        if self.is_event() {
            let mut name = BSTR::default();
            self.docinfo_from_type(Some(&mut name), None, ptr::null_mut(), None)?;
            return Ok(Some(name.to_string()));
        }
        Ok(None)
    }
    pub fn size_params(&self) -> i16 {
        unsafe { self.func_desc.as_ref().cParams }
    }
    pub fn size_opt_params(&self) -> i16 {
        unsafe { self.func_desc.as_ref().cParamsOpt }
    }
    pub fn desc(&self) -> &FUNCDESC {
        unsafe { self.func_desc.as_ref() }
    }
}

impl Drop for OleMethodData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseFuncDesc(self.func_desc.as_ptr()) };
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

pub(crate) fn ole_methods_from_typeinfo(
    typeinfo: ITypeInfo,
    mask: i32,
) -> Result<Vec<OleMethodData>> {
    let type_attr = unsafe { typeinfo.GetTypeAttr()? };
    let mut methods = vec![];
    ole_methods_sub(None, &typeinfo, &mut methods, mask)?;
    let referenced_types = ReferencedTypes::new(&typeinfo, unsafe { &*type_attr }, 0);
    for referenced_type in referenced_types.filter_map(|t| t.ok()) {
        ole_methods_sub(
            Some(&typeinfo),
            referenced_type.typeinfo(),
            &mut methods,
            mask,
        )?;
    }
    unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
    Ok(methods)
}

fn ole_methods_sub(
    owner_typeinfo: Option<&ITypeInfo>,
    typeinfo: &ITypeInfo,
    methods: &mut Vec<OleMethodData>,
    mask: i32,
) -> Result<()> {
    let methods_iter = Methods::new(typeinfo)?;
    for (i, method) in methods_iter.enumerate() {
        if let Ok(method) = method {
            if method.invkind_matches(mask) {
                let owner_type_attr = if let Some(owner_typeinfo) = owner_typeinfo {
                    let type_attr = unsafe { owner_typeinfo.GetTypeAttr()? };
                    let type_attr = NonNull::new(type_attr).unwrap();
                    Some(type_attr)
                } else {
                    None
                };
                let (typeinfo, func_desc, bstrname) = method.deconstruct();
                methods.push(OleMethodData {
                    owner_typeinfo: owner_typeinfo.cloned(),
                    owner_type_attr,
                    typeinfo,
                    name: bstrname.to_string(),
                    index: i as u32,
                    func_desc,
                });
            }
        }
    }
    Ok(())
}

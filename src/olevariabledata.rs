use std::ptr::NonNull;

use windows::Win32::System::Com::{
    ITypeInfo, TYPEDESC, VARDESC, VARFLAG_FHIDDEN, VARFLAG_FNONBROWSABLE, VARFLAG_FRESTRICTED,
    VARKIND, VAR_CONST, VAR_DISPATCH, VAR_PERINSTANCE, VAR_STATIC,
};

use crate::{
    error::Result,
    util::ole::{TypeRef, ValueDescription},
};

pub struct OleVariableData<'a> {
    typeinfo: &'a ITypeInfo,
    name: String,
    var_desc: NonNull<VARDESC>,
}

impl<'a> OleVariableData<'a> {
    pub fn new<S: AsRef<str>>(
        typeinfo: &ITypeInfo,
        index: u32,
        name: S,
    ) -> Result<OleVariableData> {
        let var_desc = unsafe { typeinfo.GetVarDesc(index)? };
        let var_desc = NonNull::new(var_desc).unwrap();
        Ok(OleVariableData {
            typeinfo,
            name: name.as_ref().into(),
            var_desc,
        })
    }
    pub fn make<S: AsRef<str>>(
        typeinfo: &ITypeInfo,
        name: S,
        var_desc: NonNull<VARDESC>,
    ) -> OleVariableData {
        OleVariableData {
            typeinfo,
            name: name.as_ref().into(),
            var_desc,
        }
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn ole_type(&self) -> String {
        self.ole_typedesc2val(None)
    }
    pub fn ole_type_detail(&self) -> Vec<String> {
        let mut typedetails = vec![];
        self.ole_typedesc2val(Some(&mut typedetails));
        typedetails
    }
    //pub fn value(&self)
    pub fn visible(&self) -> bool {
        let visible = unsafe { (self.var_desc.as_ref()).wVarFlags.0 }
            & (VARFLAG_FHIDDEN.0 | VARFLAG_FRESTRICTED.0 | VARFLAG_FNONBROWSABLE.0)
            == 0;
        visible
    }
    pub fn variable_kind(&self) -> &str {
        match unsafe { (self.var_desc.as_ref()).varkind } {
            VAR_PERINSTANCE => "PERINSTANCE",
            VAR_STATIC => "STATIC",
            VAR_CONST => "CONSTANT",
            VAR_DISPATCH => "DISPATCH",
            _ => "UNKNOWN",
        }
    }
    pub fn varkind(&self) -> VARKIND {
        unsafe { (self.var_desc.as_ref()).varkind }
    }
}

impl<'a> Drop for OleVariableData<'a> {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseVarDesc(self.var_desc.as_ptr()) };
    }
}

impl<'a> TypeRef for OleVariableData<'a> {
    fn typeinfo(&self) -> &ITypeInfo {
        self.typeinfo
    }
    fn typedesc(&self) -> &TYPEDESC {
        unsafe { &((self.var_desc.as_ref()).elemdescVar.tdesc) }
    }
}

impl<'a> ValueDescription for OleVariableData<'a> {}

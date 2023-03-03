use std::ptr::{self, NonNull};

use windows::{
    core::BSTR,
    Win32::System::Com::{
        ITypeInfo, TYPEDESC, VARDESC, VARFLAG_FHIDDEN, VARFLAG_FNONBROWSABLE, VARFLAG_FRESTRICTED,
        VARIANT, VARKIND, VAR_CONST, VAR_DISPATCH, VAR_PERINSTANCE, VAR_STATIC,
    },
};

use crate::{
    error::Result,
    util::ole::{TypeRef, ValueDescription},
};

pub struct OleVariableData {
    typeinfo: ITypeInfo,
    name: String,
    var_desc: NonNull<VARDESC>,
}

impl OleVariableData {
    pub fn new<S: AsRef<str>>(
        typeinfo: &ITypeInfo,
        index: u32,
        name: S,
    ) -> Result<OleVariableData> {
        let var_desc = unsafe { typeinfo.GetVarDesc(index)? };
        let var_desc = NonNull::new(var_desc).unwrap();
        Ok(OleVariableData {
            typeinfo: typeinfo.clone(),
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
            typeinfo: typeinfo.clone(),
            name: name.as_ref().into(),
            var_desc,
        }
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn variant(&self) -> *mut VARIANT {
        unsafe { self.var_desc.as_ref().Anonymous.lpvarValue }
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
        match self.varkind() {
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
    pub fn member_id(&self) -> i32 {
        unsafe { self.var_desc.as_ref().memid }
    }
    fn docinfo(
        &self,
        name: Option<*mut BSTR>,
        helpstr: Option<*mut BSTR>,
        helpcontext: *mut u32,
        helpfile: Option<*mut BSTR>,
    ) -> Result<()> {
        unsafe {
            self.typeinfo.GetDocumentation(
                self.var_desc.as_ref().memid,
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
        self.docinfo(
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
        self.docinfo(None, Some(&mut helpstring), ptr::null_mut(), None)?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        self.docinfo(None, None, ptr::null_mut(), Some(&mut helpfile))?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self) -> Result<u32> {
        let mut helpcontext = 0;
        self.docinfo(None, None, &mut helpcontext, None)?;
        Ok(helpcontext)
    }
}

impl Drop for OleVariableData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseVarDesc(self.var_desc.as_ptr()) };
    }
}

impl TypeRef for OleVariableData {
    fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    fn typedesc(&self) -> &TYPEDESC {
        unsafe { &((self.var_desc.as_ref()).elemdescVar.tdesc) }
    }
}

impl ValueDescription for OleVariableData {}

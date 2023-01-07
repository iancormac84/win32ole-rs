use windows::Win32::System::Com::{
    ITypeInfo, VARFLAG_FHIDDEN, VARFLAG_FNONBROWSABLE, VARFLAG_FRESTRICTED, VARKIND, VAR_CONST,
    VAR_DISPATCH, VAR_PERINSTANCE, VAR_STATIC,
};

use crate::{error::Result, util::ole::ole_typedesc2val};

pub struct OleVariableData {
    typeinfo: ITypeInfo,
    index: u32,
    name: String,
}

impl OleVariableData {
    pub fn new<S: AsRef<str>>(typeinfo: &ITypeInfo, index: u32, name: S) -> OleVariableData {
        OleVariableData {
            typeinfo: typeinfo.clone(),
            index,
            name: name.as_ref().into(),
        }
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn ole_type(&self) -> Result<String> {
        let vardesc = unsafe { self.typeinfo.GetVarDesc(self.index) }?;
        let type_ = ole_typedesc2val(
            &self.typeinfo,
            unsafe { &((*vardesc).elemdescVar.tdesc) },
            None,
        );
        unsafe { self.typeinfo.ReleaseVarDesc(vardesc) };
        Ok(type_)
    }
    pub fn ole_type_detail(&self) -> Result<Vec<String>> {
        let vardesc = unsafe { self.typeinfo.GetVarDesc(self.index) }?;
        let mut typedetails = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            unsafe { &((*vardesc).elemdescVar.tdesc) },
            Some(&mut typedetails),
        );
        unsafe { self.typeinfo.ReleaseVarDesc(vardesc) };
        Ok(typedetails)
    }
    //pub fn value(&self)
    pub fn visible(&self) -> bool {
        let vardesc = unsafe { self.typeinfo.GetVarDesc(self.index) };
        let Ok(vardesc) = vardesc else {
            return false;
        };
        let visible = unsafe { (*vardesc).wVarFlags.0 }
            & (VARFLAG_FHIDDEN.0 | VARFLAG_FRESTRICTED.0 | VARFLAG_FNONBROWSABLE.0)
            == 0;
        unsafe { self.typeinfo.ReleaseVarDesc(vardesc) };
        visible
    }
    pub fn variable_kind(&self) -> &str {
        let vardesc = unsafe { self.typeinfo.GetVarDesc(self.index) };
        let Ok(vardesc) = vardesc else {
            return "UNKNOWN";
        };
        let kind = match unsafe { (*vardesc).varkind } {
            VAR_PERINSTANCE => "PERINSTANCE",
            VAR_STATIC => "STATIC",
            VAR_CONST => "CONSTANT",
            VAR_DISPATCH => "DISPATCH",
            _ => "UNKNOWN",
        };
        unsafe { self.typeinfo.ReleaseVarDesc(vardesc) };
        kind
    }
    pub fn varkind(&self) -> Result<VARKIND> {
        let vardesc = unsafe { self.typeinfo.GetVarDesc(self.index) }?;
        let kind = unsafe { (*vardesc).varkind };
        unsafe { self.typeinfo.ReleaseVarDesc(vardesc) };
        Ok(kind)
    }
}

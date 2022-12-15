use crate::{
    error::{Error, Result},
    olemethoddata::ole_methods_from_typeinfo,
    oletypelibdata::typelib_file,
    util::{
        conv::ToWide,
        ole::{ole_initialized, ole_typedesc2val, ole_docinfo_from_type},
    },
    OleMethodData,
};
use std::{ffi::OsStr, ptr};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::System::{
        Com::{
            IDispatch, ITypeInfo, ITypeLib, ProgIDFromCLSID, INVOKE_FUNC, INVOKE_PROPERTYGET,
            INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_ALIAS, TYPEKIND, IMPLTYPEFLAGS, IMPLTYPEFLAG_FSOURCE, IMPLTYPEFLAG_FDEFAULT,
        },
        Ole::{LoadTypeLibEx, REGKIND_NONE, TYPEFLAG_FHIDDEN, TYPEFLAG_FRESTRICTED},
    },
};

//TODO: Remove dispatch member variable possibly by making initialized IDispatch'es global
pub struct OleTypeData {
    pub dispatch: Option<IDispatch>,
    pub typeinfo: ITypeInfo,
    pub name: String,
}

impl OleTypeData {
    pub fn from_typelib_and_oleclass<S: AsRef<OsStr>>(
        typelib: S,
        oleclass: S,
    ) -> Result<OleTypeData> {
        ole_initialized();
        let file = typelib_file(&typelib)?;
        let file_vec = file.to_wide_null();
        let typelib_iface =
            unsafe { LoadTypeLibEx(PCWSTR::from_raw(file_vec.as_ptr()), REGKIND_NONE)? };
        let maybe_typedata = oleclass_from_typelib(&typelib_iface, &oleclass);
        if let Some(typedata) = maybe_typedata {
            Ok(typedata)
        } else {
            return Err(Error::Custom(format!(
                "`{}` not found in `{}`",
                oleclass.as_ref().to_str().unwrap(),
                typelib.as_ref().to_str().unwrap()
            )));
        }
    }
    pub fn from_itypeinfo(typeinfo: &ITypeInfo) -> Result<OleTypeData> {
        let mut index = 0;
        let mut typelib = None;
        unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut index) }?;
        let typelib = typelib.unwrap();
        let mut bstr = BSTR::default();
        unsafe {
            typelib.GetDocumentation(index as i32, Some(&mut bstr), None, ptr::null_mut(), None)
        }?;
        let typedata = OleTypeData {
            dispatch: None,
            typeinfo: typeinfo.clone(),
            name: bstr.to_string(),
        };
        Ok(typedata)
    }
    pub fn helpstring(&self) -> Result<String> {
        let mut helpstring = BSTR::default();
        ole_docinfo_from_type(&self.typeinfo, Some(&mut helpstring), None, ptr::null_mut(), None)?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        ole_docinfo_from_type(&self.typeinfo, None, None, ptr::null_mut(), Some(&mut helpfile))?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self) -> Result<u32> {
        let mut helpcontext = 0;
        ole_docinfo_from_type(&self.typeinfo, None, None, &mut helpcontext, None)?;
        Ok(helpcontext)
    }
    pub fn major_version(&self) -> Result<u16> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };
        let ver = unsafe { (*type_attr).wMajorVerNum };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(ver)
    }
    pub fn minor_version(&self) -> Result<u16> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };
        let ver = unsafe { (*type_attr).wMinorVerNum };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(ver)
    }
    pub fn typekind(&self) -> Result<TYPEKIND> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };
        let typekind = unsafe { (*type_attr).typekind };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(typekind)
    }
    #[allow(non_snake_case, unused_variables)]
    pub fn ole_type(&self) -> Result<&str> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };

        let kind = unsafe { (*type_attr).typekind.0 };
        let type_ = match kind {
            0 => "Enum",
            1 => "Record",
            2 => "Module",
            3 => "Interface",
            4 => "Dispatch",
            5 => "Class",
            6 => "Alias",
            7 => "Union",
            8 => "Max",
            _ => panic!("TYPEKIND({kind}) has no WINAPI raw representation"),
        };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(type_)
    }
    pub fn guid(&self) -> Result<GUID> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };
        let guid = unsafe { (*type_attr).guid };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(guid)
    }
    pub fn progid(&self) -> Result<String> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr()? };
        let result = unsafe { ProgIDFromCLSID(&(*type_attr).guid)? };
        let progid = unsafe { result.to_string().unwrap() };
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Ok(progid)
    }
    pub fn visible(&self) -> bool {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr() };
        let Ok(type_attr) = type_attr else {
            return true;
        };
        let typeflags = unsafe { (*type_attr).wTypeFlags };
        let visible = typeflags & (TYPEFLAG_FHIDDEN.0 | TYPEFLAG_FRESTRICTED.0) as u16 == 0;
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        visible
    }
    pub fn variables(&self) -> Result<Vec<String>> {
        let type_attr_ptr = unsafe { self.typeinfo.GetTypeAttr()? };
        let mut variables = vec![];
        for i in 0..unsafe { (*type_attr_ptr).cVars } {
            let var_desc_ptr = unsafe { self.typeinfo.GetVarDesc(i as u32)? };
            let mut len = 0;
            let mut rgbstrnames = BSTR::default();
            let res = unsafe {
                self.typeinfo
                    .GetNames((*var_desc_ptr).memid, &mut rgbstrnames, 1, &mut len)
            };
            if res.is_err() || len == 0 || rgbstrnames.is_empty() {
                continue;
            }
            variables.push(String::try_from(rgbstrnames)?);
            unsafe { self.typeinfo.ReleaseVarDesc(var_desc_ptr) };
        }
        Ok(variables)
    }
    pub fn src_type(&self) -> Option<String> {
        let type_attr = unsafe { self.typeinfo.GetTypeAttr() };
        if type_attr.is_err() {
            return None;
        };
        let type_attr = type_attr.unwrap();
        if unsafe { (*type_attr).typekind } != TKIND_ALIAS {
            unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
            return None;
        }
        let alias = ole_typedesc2val(&self.typeinfo, &(unsafe { (*type_attr).tdescAlias }), None);
        unsafe { self.typeinfo.ReleaseTypeAttr(type_attr) };
        Some(alias)
    }
    pub fn methods(&self) -> Result<Vec<OleMethodData>> {
        ole_methods_from_typeinfo(
            &self.typeinfo,
            INVOKE_FUNC.0 | INVOKE_PROPERTYGET.0 | INVOKE_PROPERTYPUT.0 | INVOKE_PROPERTYPUTREF.0,
        )
    }
    pub fn implemented_ole_types(&self) -> Result<Vec<OleTypeData>> {
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAGS(0))
    }
    pub fn source_ole_types(&self) -> Result<Vec<OleTypeData>> {
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAG_FSOURCE)
    }
    pub fn default_event_sources(&self) -> Result<Vec<OleTypeData>> {
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAG_FSOURCE|IMPLTYPEFLAG_FDEFAULT)
    }
    pub fn default_ole_types(&self) -> Result<Vec<OleTypeData>> {
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAG_FDEFAULT)
    }
}

fn oleclass_from_typelib<P: AsRef<OsStr>>(typelib: &ITypeLib, oleclass: P) -> Option<OleTypeData> {
    let mut found = false;
    let mut typedata: Option<OleTypeData> = None;

    let count = unsafe { typelib.GetTypeInfoCount() };
    let mut i = 0;
    while i < count && !found {
        let typeinfo = unsafe { typelib.GetTypeInfo(i) };
        let Ok(typeinfo) = typeinfo else {
            continue;
        };
        let mut bstrname = BSTR::default();
        let result = unsafe {
            typelib.GetDocumentation(i as i32, Some(&mut bstrname), None, ptr::null_mut(), None)
        };
        if result.is_err() {
            continue;
        }
        let oleclass = oleclass.as_ref();
        if oleclass.to_str().unwrap() == bstrname {
            typedata = Some(OleTypeData {
                dispatch: None,
                typeinfo,
                name: bstrname.to_string(),
            });
            found = true;
        }
        i += 1;
    }
    typedata
}

fn ole_type_impl_ole_types(typeinfo: &ITypeInfo, implflags: IMPLTYPEFLAGS) -> Result<Vec<OleTypeData>> {
    let mut types = vec![];
    let type_attr = unsafe { typeinfo.GetTypeAttr() }?;
    
    for i in 0..unsafe { (*type_attr).cImplTypes } {
        let flags = unsafe { typeinfo.GetImplTypeFlags(i as u32) };
        let Ok(flags) = flags else {
            continue;
        };

        let href = unsafe { typeinfo.GetRefTypeOfImplType(i as u32) };
        let Ok(href) = href else {
            continue;
        };
        let ref_type_info = unsafe { typeinfo.GetRefTypeInfo(href) };
        let Ok(ref_type_info) = ref_type_info else {
            continue;
        };

        if (flags & implflags) == implflags {
            let type_ = OleTypeData::from_itypeinfo(&ref_type_info);
            if let Ok(type_) = type_ {
                types.push(type_);
            }
        }
    }
    unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
    Ok(types)
}
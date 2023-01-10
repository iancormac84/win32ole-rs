use crate::{
    error::{Error, OleError, Result},
    olemethoddata::ole_methods_from_typeinfo,
    oletypelibdata::typelib_file,
    olevariabledata::OleVariableData,
    types::{OleClassNames, TypeInfos},
    util::{
        conv::ToWide,
        ole::{ole_docinfo_from_type, ole_initialized, ole_typedesc2val},
    },
    OleMethodData,
};
use std::{
    ffi::OsStr,
    iter::zip,
    ptr::{self, NonNull},
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::System::{
        Com::{
            IDispatch, ITypeInfo, ITypeLib, ProgIDFromCLSID, IMPLTYPEFLAGS, IMPLTYPEFLAG_FDEFAULT,
            IMPLTYPEFLAG_FSOURCE, INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
            INVOKE_PROPERTYPUTREF, TKIND_ALIAS, TYPEATTR, TYPEKIND,
        },
        Ole::{LoadTypeLibEx, REGKIND_NONE, TYPEFLAG_FHIDDEN, TYPEFLAG_FRESTRICTED},
    },
};

//TODO: Remove dispatch member variable possibly by making initialized IDispatch'es global
pub struct OleTypeData {
    dispatch: Option<IDispatch>,
    typeinfo: ITypeInfo,
    name: String,
    type_attr: NonNull<TYPEATTR>,
}

impl OleTypeData {
    pub fn new<S: AsRef<OsStr>>(typelib: S, oleclass: S) -> Result<OleTypeData> {
        ole_initialized();
        let file = typelib_file(&typelib)?;
        let file_vec = file.to_wide_null();
        let typelib_iface =
            unsafe { LoadTypeLibEx(PCWSTR::from_raw(file_vec.as_ptr()), REGKIND_NONE)? };
        let maybe_typedata = oleclass_from_typelib(&typelib_iface, &oleclass)?;
        match maybe_typedata {
            Some(typedata) => Ok(typedata),
            None => Err(Error::Custom(format!(
                "`{}` not found in `{}`",
                oleclass.as_ref().to_str().unwrap(),
                typelib.as_ref().to_str().unwrap()
            ))),
        }
    }
    pub fn make<S: AsRef<str>>(
        dispatch: Option<IDispatch>,
        typeinfo: ITypeInfo,
        name: S,
    ) -> Result<OleTypeData> {
        let type_attr = unsafe { typeinfo.GetTypeAttr() }?;
        let type_attr = NonNull::new(type_attr).unwrap();

        Ok(OleTypeData {
            dispatch,
            typeinfo,
            name: name.as_ref().to_string(),
            type_attr,
        })
    }
    pub fn helpstring(&self) -> Result<String> {
        let mut helpstring = BSTR::default();
        ole_docinfo_from_type(
            &self.typeinfo,
            Some(&mut helpstring),
            None,
            ptr::null_mut(),
            None,
        )?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        ole_docinfo_from_type(
            &self.typeinfo,
            None,
            None,
            ptr::null_mut(),
            Some(&mut helpfile),
        )?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self) -> Result<u32> {
        let mut helpcontext = 0;
        ole_docinfo_from_type(&self.typeinfo, None, None, &mut helpcontext, None)?;
        Ok(helpcontext)
    }
    pub fn major_version(&self) -> u16 {
        unsafe { self.type_attr.as_ref().wMajorVerNum }
    }
    pub fn minor_version(&self) -> u16 {
        unsafe { self.type_attr.as_ref().wMinorVerNum }
    }
    pub fn typekind(&self) -> TYPEKIND {
        unsafe { self.type_attr.as_ref().typekind }
    }
    #[allow(non_snake_case, unused_variables)]
    pub fn ole_type(&self) -> &str {
        let kind = unsafe { self.type_attr.as_ref().typekind.0 };
        match kind {
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
        }
    }
    pub fn guid(&self) -> GUID {
        unsafe { self.type_attr.as_ref().guid }
    }
    pub fn progid(&self) -> Result<String> {
        let result = unsafe { ProgIDFromCLSID(&self.guid())? };
        Ok(unsafe { result.to_string()? })
    }
    pub fn visible(&self) -> bool {
        let typeflags = unsafe { self.type_attr.as_ref().wTypeFlags };
        typeflags & (TYPEFLAG_FHIDDEN.0 | TYPEFLAG_FRESTRICTED.0) as u16 == 0
    }
    pub fn variables(&self) -> Result<Vec<OleVariableData>> {
        let mut variables = vec![];
        for i in 0..unsafe { self.type_attr.as_ref().cVars } {
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
            let index = i as u32;
            let name = String::try_from(rgbstrnames)?;
            variables.push(OleVariableData::new(&self.typeinfo, index, name));
            unsafe { self.typeinfo.ReleaseVarDesc(var_desc_ptr) };
        }
        Ok(variables)
    }
    pub fn src_type(&self) -> Option<String> {
        if unsafe { self.type_attr.as_ref().typekind } != TKIND_ALIAS {
            return None;
        }
        Some(ole_typedesc2val(
            &self.typeinfo,
            &(unsafe { self.type_attr.as_ref().tdescAlias }),
            None,
        ))
    }
    pub fn ole_methods(&self) -> Result<Vec<OleMethodData>> {
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
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAG_FSOURCE | IMPLTYPEFLAG_FDEFAULT)
    }
    pub fn default_ole_types(&self) -> Result<Vec<OleTypeData>> {
        ole_type_impl_ole_types(&self.typeinfo, IMPLTYPEFLAG_FDEFAULT)
    }
    pub fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    pub fn name(&self) -> &str {
        &self.name
    }
}

impl TryFrom<&ITypeInfo> for OleTypeData {
    type Error = Error;

    fn try_from(typeinfo: &ITypeInfo) -> Result<OleTypeData> {
        let mut index = 0;
        let mut typelib = None;
        let result = unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut index) };
        if let Err(error) = result {
            return Err(OleError::interface(
                error,
                "failed to GetContainingTypeLib from ITypeInfo",
            )
            .into());
        }
        let typelib = typelib.unwrap();
        let mut bstr = BSTR::default();
        let result = unsafe {
            typelib.GetDocumentation(index as i32, Some(&mut bstr), None, ptr::null_mut(), None)
        };
        if let Err(error) = result {
            return Err(
                OleError::interface(error, "failed to GetDocumentation from ITypeLib").into(),
            );
        }
        let type_attr = unsafe { typeinfo.GetTypeAttr()? };
        let type_attr = NonNull::new(type_attr).unwrap();

        Ok(OleTypeData {
            dispatch: None,
            typeinfo: typeinfo.clone(),
            name: bstr.to_string(),
            type_attr,
        })
    }
}

fn oleclass_from_typelib<P: AsRef<OsStr>>(
    typelib: &ITypeLib,
    oleclass: P,
) -> Result<Option<OleTypeData>> {
    let typeinfos = TypeInfos::from(typelib);
    let ole_class_names = OleClassNames::from(typelib);
    let iter_pair = zip(typeinfos, ole_class_names);
    for (typeinfo, ole_class_name) in iter_pair {
        let Ok(typeinfo) = typeinfo else {
            continue;
        };
        let Ok(ole_class_name) = ole_class_name else {
            continue;
        };
        if ole_class_name == oleclass.as_ref().to_str().unwrap() {
            let type_attr = unsafe { typeinfo.GetTypeAttr()? };
            let type_attr = NonNull::new(type_attr).unwrap();

            return Ok(Some(OleTypeData {
                dispatch: None,
                typeinfo,
                name: ole_class_name,
                type_attr,
            }));
        }
    }
    Ok(None)
}

fn ole_type_impl_ole_types(
    typeinfo: &ITypeInfo,
    implflags: IMPLTYPEFLAGS,
) -> Result<Vec<OleTypeData>> {
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
            let type_ = OleTypeData::try_from(&ref_type_info);
            if let Ok(type_) = type_ {
                types.push(type_);
            }
        }
    }
    unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
    Ok(types)
}

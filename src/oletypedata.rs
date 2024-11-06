use crate::{
    error::{Error, OleError, Result},
    olemethoddata::ole_methods_from_typeinfo,
    oletypelibdata::typelib_file,
    olevariabledata::OleVariableData,
    types::{OleClassNames, ReferencedTypes, TypeInfos, Variables},
    util::{
        conv::ToWide,
        ole::{ole_docinfo, ole_initialized, ole_typedesc2val},
    },
    OleMethodData,
};
use std::{
    iter::zip,
    ptr::{self, NonNull},
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::System::{
        Com::{
            ITypeInfo, ITypeLib, ProgIDFromCLSID, IMPLTYPEFLAGS, IMPLTYPEFLAG_FDEFAULT,
            IMPLTYPEFLAG_FSOURCE, INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
            INVOKE_PROPERTYPUTREF, TKIND_ALIAS, TYPEATTR, TYPEKIND,
        },
        Ole::{LoadTypeLibEx, REGKIND_NONE, TYPEFLAG_FHIDDEN, TYPEFLAG_FRESTRICTED},
    },
};

pub struct OleTypeData {
    typeinfo: ITypeInfo,
    name: String,
    type_attr: NonNull<TYPEATTR>,
}

impl OleTypeData {
    pub fn new<S: AsRef<str>>(typelib: S, oleclass: S) -> Result<OleTypeData> {
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
                oleclass.as_ref(),
                typelib.as_ref()
            ))),
        }
    }
    pub fn make<S: AsRef<str>>(typeinfo: ITypeInfo, name: S) -> Result<OleTypeData> {
        let type_attr = unsafe { typeinfo.GetTypeAttr() }?;
        let type_attr = NonNull::new(type_attr).unwrap();

        Ok(OleTypeData {
            typeinfo,
            name: name.as_ref().to_string(),
            type_attr,
        })
    }
    pub fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    pub fn attribs(&self) -> &TYPEATTR {
        unsafe { self.type_attr.as_ref() }
    }
    pub fn get_documentation(&self) -> Result<(String, String, u32, String)> {
        let mut strname = BSTR::default();
        let mut strdocstring = BSTR::default();
        let mut whelpcontext = 0;
        let mut strhelpfile = BSTR::default();
        ole_docinfo(
            &self.typeinfo,
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
        ole_docinfo(
            &self.typeinfo,
            None,
            Some(&mut helpstring),
            ptr::null_mut(),
            None,
        )?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        ole_docinfo(
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
        ole_docinfo(&self.typeinfo, None, None, &mut helpcontext, None)?;
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
    pub fn variables(&self) -> Vec<Result<OleVariableData>> {
        let vars = Variables::new(self.typeinfo(), self.attribs());
        vars.collect()
    }
    pub fn src_type(&self) -> Option<String> {
        if unsafe { self.type_attr.as_ref().typekind } != TKIND_ALIAS {
            return None;
        }
        let type_attr_ref = unsafe { self.type_attr.as_ref() };

        Some(ole_typedesc2val(
            &self.typeinfo,
            &type_attr_ref.tdescAlias,
            None,
        ))
    }
    pub fn ole_methods(&self) -> Result<Vec<OleMethodData>> {
        ole_methods_from_typeinfo(
            self.typeinfo.clone(),
            INVOKE_FUNC.0 | INVOKE_PROPERTYGET.0 | INVOKE_PROPERTYPUT.0 | INVOKE_PROPERTYPUTREF.0,
        )
    }
    fn ole_type_impl_ole_types(&self, implflags: IMPLTYPEFLAGS) -> Result<Vec<OleTypeData>> {
        let mut types = vec![];

        let referenced_types = ReferencedTypes::from_type(self);
        for referenced_type in referenced_types.filter_map(|t| t.ok()) {
            if referenced_type.matches(implflags) {
                let type_ = OleTypeData::try_from(referenced_type.into_typeinfo());
                if let Ok(type_) = type_ {
                    types.push(type_);
                }
            }
        }

        Ok(types)
    }
    pub fn implemented_ole_types(&self) -> Result<Vec<OleTypeData>> {
        self.ole_type_impl_ole_types(IMPLTYPEFLAGS(0))
    }
    pub fn source_ole_types(&self) -> Result<Vec<OleTypeData>> {
        self.ole_type_impl_ole_types(IMPLTYPEFLAG_FSOURCE)
    }
    pub fn default_event_sources(&self) -> Result<Vec<OleTypeData>> {
        self.ole_type_impl_ole_types(IMPLTYPEFLAG_FSOURCE | IMPLTYPEFLAG_FDEFAULT)
    }
    pub fn default_ole_types(&self) -> Result<Vec<OleTypeData>> {
        self.ole_type_impl_ole_types(IMPLTYPEFLAG_FDEFAULT)
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn get_ref_type_info(&self, ref_type: u32) -> Result<OleTypeData> {
        let ref_type_info = unsafe { self.typeinfo.GetRefTypeInfo(ref_type)? };

        OleTypeData::try_from(ref_type_info)
    }
    pub fn get_interface_of_dispinterface(&self) -> Result<OleTypeData> {
        let ref_type = unsafe { self.typeinfo.GetRefTypeOfImplType((-1i32) as u32)? };
        let typeinfo = unsafe { self.typeinfo.GetRefTypeInfo(ref_type)? };
        OleTypeData::try_from(typeinfo)
    }
    pub fn num_impl_types(&self) -> u16 {
        unsafe { self.type_attr.as_ref().cImplTypes }
    }
    pub fn num_funcs(&self) -> u16 {
        unsafe { self.type_attr.as_ref().cFuncs }
    }
    pub fn num_variables(&self) -> u16 {
        unsafe { self.type_attr.as_ref().cVars }
    }
}

impl Drop for OleTypeData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseTypeAttr(self.type_attr.as_ptr()) };
    }
}

impl TryFrom<ITypeInfo> for OleTypeData {
    type Error = Error;

    fn try_from(typeinfo: ITypeInfo) -> Result<OleTypeData> {
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
            typeinfo,
            name: bstr.to_string(),
            type_attr,
        })
    }
}

fn oleclass_from_typelib<P: AsRef<str>>(
    typelib: &ITypeLib,
    oleclass: P,
) -> Result<Option<OleTypeData>> {
    let typeinfos = TypeInfos::from(typelib);
    let ole_class_names = OleClassNames::from(typelib);
    let iter_pair = zip(typeinfos, ole_class_names);
    for (typeinfo, name) in iter_pair {
        let Ok(typeinfo) = typeinfo else {
            continue;
        };
        let Ok(name) = name else {
            continue;
        };

        if name == oleclass.as_ref() {
            let type_attr = unsafe { typeinfo.GetTypeAttr()? };
            let type_attr = NonNull::new(type_attr).unwrap();

            return Ok(Some(OleTypeData {
                typeinfo,
                name,
                type_attr,
            }));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use windows::core::GUID;

    #[test]
    fn test_initialize() {
        let ole_type = super::OleTypeData::new("Microsoft Shell Controls And Automation", "Shell");
        assert!(ole_type.is_ok());
        let ole_type = ole_type.unwrap();
        assert_eq!(ole_type.name, "Shell");
        assert_eq!(ole_type.ole_type(), "Class");
        assert_eq!(
            ole_type.guid(),
            GUID::try_from("13709620-C279-11CE-A49E-444553540000").unwrap()
        );
        assert_eq!(ole_type.progid().unwrap(), "Shell.Application.1");
        assert!(ole_type.visible());
        assert_eq!(ole_type.major_version(), 0);
        assert_eq!(ole_type.minor_version(), 0);
        assert_eq!(ole_type.typekind().0, 5);
        assert_eq!(
            ole_type.helpstring().unwrap(),
            "Shell Object Type Information"
        );
        assert_eq!(ole_type.src_type(), None);
        assert_eq!(ole_type.helpfile().unwrap(), "");
        assert_eq!(ole_type.helpcontext().unwrap(), 0);
        assert!(ole_type.variables().is_empty());
        assert!(ole_type
            .ole_methods()
            .unwrap()
            .iter()
            .any(|m| m.name() == "NameSpace"));

        let ole_type2 = super::OleTypeData::new("{13709620-C279-11CE-A49E-444553540000}", "Shell");
        assert!(ole_type2.is_ok());
        let ole_type2 = ole_type2.unwrap();
        assert_eq!(ole_type.name, ole_type2.name);
        assert_eq!(ole_type.ole_type(), ole_type2.ole_type());
        assert_eq!(ole_type.guid(), ole_type2.guid());
        assert_eq!(ole_type.progid().unwrap(), ole_type2.progid().unwrap());
        assert_eq!(ole_type.visible(), ole_type2.visible());
        assert_eq!(ole_type.major_version(), ole_type2.major_version());
        assert_eq!(ole_type.minor_version(), ole_type2.minor_version());
        assert_eq!(ole_type.typekind(), ole_type2.typekind());
        assert_eq!(
            ole_type.helpstring().unwrap(),
            ole_type2.helpstring().unwrap()
        );
        assert_eq!(ole_type.src_type(), ole_type2.src_type());
        assert_eq!(ole_type.helpfile().unwrap(), ole_type2.helpfile().unwrap());
        assert_eq!(
            ole_type.helpcontext().unwrap(),
            ole_type2.helpcontext().unwrap()
        );
        assert_eq!(ole_type.variables().len(), ole_type2.variables().len());
        assert_eq!(
            ole_type.ole_methods().unwrap()[0].name(),
            ole_type2.ole_methods().unwrap()[0].name()
        );
        //assert_eq!(ole_type.typelib().name, ole_type2.ole_typelib.name);
        assert_eq!(
            ole_type.implemented_ole_types().unwrap().len(),
            ole_type2.implemented_ole_types().unwrap().len()
        );
    }
}

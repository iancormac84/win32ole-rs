use crate::{
    error::{Error, Result},
    olemethoddata::ole_methods_from_typeinfo,
    oletypelibdata::typelib_file,
    util::{get_class_id, ole_initialized, ole_typedesc2val, ToWide},
    OleMethodData,
};
use std::{ffi::OsStr, ptr};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{
                CoCreateInstance, IDispatch, ITypeInfo, ITypeLib, ProgIDFromCLSID,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, INVOKE_FUNC, INVOKE_PROPERTYGET,
                INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_ALIAS, TYPEKIND,
            },
            Ole::{LoadTypeLibEx, REGKIND_NONE, TYPEFLAG_FHIDDEN, TYPEFLAG_FRESTRICTED},
        },
    },
};

pub struct OleTypeData {
    dispatch: Option<IDispatch>,
    pub typeinfo: ITypeInfo,
}

impl OleTypeData {
    pub fn from_prog_id<S: AsRef<OsStr>>(prog_id: S) -> Result<OleTypeData> {
        ole_initialized();
        let app_clsid = get_class_id(prog_id)?;
        let flags = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;
        let dispatch: IDispatch = unsafe { CoCreateInstance(&app_clsid, None, flags)? };
        let typeinfo = unsafe { dispatch.GetTypeInfo(0, GetUserDefaultLCID())? };
        Ok(OleTypeData {
            dispatch: Some(dispatch),
            typeinfo,
        })
    }
    pub fn from_typelib_and_oleclass<S: AsRef<OsStr>>(
        typelib: S,
        oleclass: S,
    ) -> Result<OleTypeData> {
        let typelib = typelib.as_ref();
        let typelib_vec = typelib.to_wide_null();
        let typelib_pcwstr = PCWSTR::from_raw(typelib_vec.as_ptr());
        let oleclass = oleclass.as_ref();
        let oleclass_vec = oleclass.to_wide_null();
        let oleclass_pcwstr = PCWSTR::from_raw(oleclass_vec.as_ptr());
        ole_initialized();
        let mut file = typelib_file(typelib_pcwstr)?;
        if file.is_null() {
            file = typelib_pcwstr;
        }
        let typelib_iface = unsafe { LoadTypeLibEx(file, REGKIND_NONE)? };
        let maybe_typedata = oleclass_from_typelib(&typelib_iface, oleclass_pcwstr);
        if let Some(typedata) = maybe_typedata {
            Ok(typedata)
        } else {
            return Err(Error::Custom(format!(
                "`{}` not found in `{}`",
                unsafe { oleclass_pcwstr.display() },
                unsafe { typelib_pcwstr.display() }
            )));
        }
    }
    fn ole_docinfo_from_type(
        &self,
        name: Option<*mut BSTR>,
        helpstr: Option<*mut BSTR>,
        helpcontext: *mut u32,
        helpfile: Option<*mut BSTR>,
    ) -> Result<()> {
        let mut tlib: Option<ITypeLib> = None;
        let mut index = 0;
        unsafe { self.typeinfo.GetContainingTypeLib(&mut tlib, &mut index)? };
        let tlib = tlib.unwrap();
        unsafe { tlib.GetDocumentation(index as i32, name, helpstr, helpcontext, helpfile)? };
        Ok(())
    }
    pub fn helpstring(&self) -> Result<String> {
        let mut helpstring = BSTR::default();
        self.ole_docinfo_from_type(None, Some(&mut helpstring), ptr::null_mut(), None)?;
        Ok(String::try_from(helpstring)?)
    }
    pub fn helpfile(&self) -> Result<String> {
        let mut helpfile = BSTR::default();
        self.ole_docinfo_from_type(None, None, ptr::null_mut(), Some(&mut helpfile))?;
        Ok(String::try_from(helpfile)?)
    }
    pub fn helpcontext(&self) -> Result<u32> {
        let mut helpcontext = 0;
        self.ole_docinfo_from_type(None, None, &mut helpcontext, None)?;
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
            _ => panic!("TYPEKIND({}) has no WINAPI raw representation", kind),
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
}

fn oleclass_from_typelib(typelib: &ITypeLib, oleclass: PCWSTR) -> Option<OleTypeData> {
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
        if unsafe { oleclass.as_wide() } == bstrname.as_wide() {
            typedata = Some(OleTypeData {
                dispatch: None,
                typeinfo,
            });
            found = true;
        }
        i += 1;
    }
    typedata
}

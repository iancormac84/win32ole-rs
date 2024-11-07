use crate::{
    error::Result,
    oleparamdata::OleParamData,
    types::{Methods, ReferencedTypes},
    util::{
        conv::ToWide,
        ole::{ole_docinfo_from_type, ole_typedesc2val},
    },
    OleTypeData,
};
use std::{
    ffi::OsStr,
    ops::Deref,
    ptr::{self, NonNull},
};
use windows::{
    core::{BSTR, PCWSTR},
    Win32::System::{
        Com::{
            ITypeInfo, FUNCDESC, FUNCKIND, INVOKEKIND, INVOKE_FUNC, INVOKE_PROPERTYGET,
            INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, TKIND_COCLASS, TYPEATTR, TYPEDESC,
        },
        Variant::VARENUM,
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
        let method = OleMethodData::maybe_find_and_create(None, &typeinfo, &name)?;
        println!("In OleMethodData::from_typeinfo, There were no errors in the first call to OleMethodData::maybe_find_and_create");
        if method.is_some() {
            println!("In OleMethodData::from_typeinfo, we already found the method.");
            return Ok(method);
        }
        println!(
            "But we didn't find the method...so...in OleMethodData::from_typeinfo, we about to find the TYPEATTR for {:?}, which we will then pass to ReferencedTypes to do a deeper search",
            name.as_ref()
        );
        let type_attr = unsafe { typeinfo.GetTypeAttr()? };
        println!("In OleMethodData::from_typeinfo, we successfully found the TYPEATTR, and now we have to call ReferencedTypes::new with the typeinfo and the type_attr");
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

    // This is pretty much the same as the Ruby implementation's ole_method_sub function.
    fn maybe_find_and_create<S: AsRef<OsStr>>(
        owner_typeinfo: Option<&ITypeInfo>,
        typeinfo: &ITypeInfo,
        name: &S,
    ) -> Result<Option<OleMethodData>> {
        let methods = Methods::new(typeinfo)?;

        println!(
            "OleMethodData::maybe_find_and_create, We b looking for {:?}",
            name.as_ref()
        );
        let fname = name.to_wide_null();
        let fname_pcwstr = PCWSTR::from_raw(fname.as_ptr());

        for method in methods {
            if let Ok(method) = method {
                if unsafe { fname_pcwstr.as_wide() } == method.name().deref() {
                    let (typeinfo, func_desc, bstrname, method_index) = method.deconstruct();
                    println!("Found method name {} for method {method_index}", bstrname.to_string());

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
                        index: method_index as u32,
                        func_desc,
                    }));
                }
            }
        }

        Ok(None)
    }
    pub fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
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
    pub fn dispid(&self) -> i32 {
        unsafe { self.func_desc.as_ref().memid }
    }
    pub fn return_type(&self) -> String {
        ole_typedesc2val(&self.typeinfo, self.return_type_desc(), None)
    }
    pub fn return_type_desc(&self) -> &TYPEDESC {
        unsafe { &self.func_desc.as_ref().elemdescFunc.tdesc }
    }
    pub fn return_vtype(&self) -> VARENUM {
        unsafe { self.func_desc.as_ref().elemdescFunc.tdesc.vt }
    }
    pub fn return_type_detail(&self) -> Vec<String> {
        let mut type_details = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            self.return_type_desc(),
            Some(&mut type_details),
        );
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
        let cparams = unsafe { self.func_desc.as_ref().cParams };
        println!("cparams is {cparams}");
        let cmaxnames = cparams as u32 + 1;
        let mut bstrs = Vec::with_capacity(cmaxnames as usize);
        let mut len = 0;
        let result = unsafe {
            self.typeinfo
                .GetNames(self.func_desc.as_ref().memid, &mut bstrs, &mut len)
        };
        println!("len is {len}");
        if result.is_err() {
            println!("result is error: {result:?}");
            return vec![];
        }
        let mut params = vec![];

        println!("We are just about to compare cparams to 0");
        if cparams > 0 {
            for i in 1..bstrs.len() as u32 {
                let param =
                    OleParamData::make(self, self.index, i - 1, bstrs[i as usize].to_string());
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
            ole_docinfo_from_type(&self.typeinfo, Some(&mut name), None, ptr::null_mut(), None)?;
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
    pub fn get_ref_type_info(&self, href: u32) -> Result<ITypeInfo> {
        Ok(unsafe { self.typeinfo.GetRefTypeInfo(href)? })
    }
}

impl Drop for OleMethodData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseFuncDesc(self.func_desc.as_ptr()) };
    }
}

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
    for method in methods_iter {
        if let Ok(method) = method {
            if method.invkind_matches(mask) {
                let owner_type_attr = if let Some(owner_typeinfo) = owner_typeinfo {
                    let type_attr = unsafe { owner_typeinfo.GetTypeAttr()? };
                    let type_attr = NonNull::new(type_attr).unwrap();
                    Some(type_attr)
                } else {
                    None
                };
                let (typeinfo, func_desc, bstrname, method_index) = method.deconstruct();
                methods.push(OleMethodData {
                    owner_typeinfo: owner_typeinfo.cloned(),
                    owner_type_attr,
                    typeinfo,
                    name: bstrname.to_string(),
                    index: method_index as u32,
                    func_desc,
                });
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use windows::Win32::System::{Com::INVOKEKIND, Variant::VARENUM};

    #[test]
    fn test_win32ole_method() {
        println!("1");
        let ole_type = super::OleTypeData::new("Microsoft Shell Controls And Automation", "Shell");
        assert!(ole_type.is_ok());
        let ole_type = ole_type.unwrap();

        println!("2");
        let m_open = super::OleMethodData::new(&ole_type, "Open");
        assert!(m_open.is_ok());
        let m_open = m_open.unwrap();
        assert!(m_open.is_some());
        let m_open = m_open.unwrap();

        println!("3");
        let m_namespace = super::OleMethodData::new(&ole_type, "NameSpace");
        assert!(m_namespace.is_ok());
        let m_namespace = m_namespace.unwrap();
        assert!(m_namespace.is_some());
        let m_namespace = m_namespace.unwrap();

        println!("4");
        let m_parent = super::OleMethodData::new(&ole_type, "Parent");
        assert!(m_parent.is_ok());
        let m_parent = m_parent.unwrap();
        assert!(m_parent.is_some());
        let m_parent = m_parent.unwrap();

        println!("5");
        let m_invoke = super::OleMethodData::new(&ole_type, "Invoke");
        assert!(m_invoke.is_ok());
        let m_invoke = m_invoke.unwrap();
        assert!(m_invoke.is_some());
        let m_invoke = m_invoke.unwrap();

        println!("6");
        let m_browse_for_folder = super::OleMethodData::new(&ole_type, "BrowseForFolder");
        assert!(m_browse_for_folder.is_ok());
        let m_browse_for_folder = m_browse_for_folder.unwrap();
        assert!(m_browse_for_folder.is_some());
        let m_browse_for_folder = m_browse_for_folder.unwrap();

        println!("7");
        let ole_type1 = super::OleTypeData::new("Microsoft Scripting Runtime", "File");
        assert!(ole_type1.is_ok());
        let ole_type1 = ole_type1.unwrap();

        println!("8");
        let m_file_name = super::OleMethodData::new(&ole_type1, "Name");
        assert!(m_file_name.is_ok());
        let m_file_name = m_file_name.unwrap();
        assert!(m_file_name.is_some());
        let m_file_name = m_file_name.unwrap();

        println!("9");
        assert_eq!(m_open.name(), "Open");

        println!("10");
        println!("{}", m_open.return_type());
        assert_eq!(m_open.return_type(), "VOID");
        println!("10a");
        println!("{}", m_namespace.return_type());
        assert_eq!(m_namespace.return_type(), "Folder");

        println!("11");
        assert_eq!(m_open.return_vtype(), VARENUM(24));
        assert_eq!(m_namespace.return_vtype(), VARENUM(26));

        println!("12");
        assert_eq!(m_open.return_type_detail(), ["VOID"]);
        assert_eq!(
            m_namespace.return_type_detail(),
            ["PTR", "USERDEFINED", "Folder"]
        );

        println!("13");
        assert_eq!(m_open.invoke_kind(), "FUNC");
        assert_eq!(m_namespace.invoke_kind(), "FUNC");
        assert_eq!(m_parent.invoke_kind(), "PROPERTYGET");

        println!("14");
        assert_eq!(m_namespace.invkind(), INVOKEKIND(1));
        assert_eq!(m_parent.invkind(), INVOKEKIND(2));

        println!("15");
        assert!(m_namespace.helpstring().is_ok());
        assert_eq!(
            m_namespace.helpstring().unwrap(),
            "Get special folder from ShellSpecialFolderConstants"
        );

        println!("16");
        assert!(m_namespace.helpfile().is_ok());
        assert_eq!(m_namespace.helpfile().unwrap(), "");

        println!("17");
        assert!(m_namespace.helpcontext().is_ok());
        assert_eq!(m_namespace.helpcontext().unwrap(), 0);
        assert!(m_file_name.helpcontext().is_ok());
        assert_eq!(m_file_name.helpcontext().unwrap(), 2181996);

        println!("18");
        assert_eq!(m_namespace.dispid(), 1610743810);

        assert!(m_invoke.offset_vtbl().is_ok());
        assert_eq!(m_invoke.offset_vtbl().unwrap(), 48);

        println!("19");
        assert_eq!(m_open.size_params(), 1);
        assert_eq!(m_browse_for_folder.size_params(), 4);

        println!("20");
        assert_eq!(m_open.size_opt_params(), 0);
        assert_eq!(m_browse_for_folder.size_opt_params(), 1);

        println!("21");
        assert_eq!(m_browse_for_folder.params().len(), 4);
        assert!(m_browse_for_folder.params().iter().all(|p| p.is_ok()));

        println!("22");
    }

    #[test]
    fn test_win32ole_method_event() {
        let ole_type = super::OleTypeData::new("System Monitor Control", "SystemMonitor");
        assert!(ole_type.is_ok());
        let ole_type = ole_type.unwrap();
        let on_dbl_click = super::OleMethodData::new(&ole_type, "OnDblClick");
        assert!(on_dbl_click.is_ok());
        let on_dbl_click = on_dbl_click.unwrap();
        assert!(on_dbl_click.is_some());
        let on_dbl_click = on_dbl_click.unwrap();
        let ole_type1 = super::OleTypeData::new("Microsoft Shell Controls And Automation", "Shell");
        assert!(ole_type1.is_ok());
        let ole_type1 = ole_type1.unwrap();
        let namespace = super::OleMethodData::new(&ole_type1, "NameSpace");
        assert!(namespace.is_ok());
        let namespace = namespace.unwrap();
        assert!(namespace.is_some());
        let namespace = namespace.unwrap();
        let namespace_ei = namespace.event_interface();
        assert!(namespace_ei.is_ok());
        let namespace_ei = namespace_ei.unwrap();
        assert!(namespace_ei.is_none());
        assert!(on_dbl_click.is_event());
        let on_dbl_click_ei = on_dbl_click.event_interface();
        assert!(on_dbl_click_ei.is_ok());
        let on_dbl_click_ei = on_dbl_click_ei.unwrap();
        assert!(on_dbl_click_ei.is_some());
        let on_dbl_click_ei = on_dbl_click_ei.unwrap();
        assert_eq!(on_dbl_click_ei, "DISystemMonitorEvents");
    }
}

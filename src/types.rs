use std::ptr::{self, NonNull};

use windows::{
    core::BSTR,
    Win32::System::Com::{ITypeInfo, ITypeLib, ELEMDESC, FUNCDESC, TYPEATTR, VARDESC},
};

pub struct TypeAttr {
    typeinfo: ITypeInfo,
    name: String,
    type_attr: NonNull<TYPEATTR>,
}

impl TypeAttr {
    pub fn new(typeinfo: ITypeInfo) -> std::result::Result<TypeAttr, windows::core::Error> {
        let type_attr = unsafe { typeinfo.GetTypeAttr() }?;
        let mut name = BSTR::default();
        unsafe { typeinfo.GetDocumentation(-1, Some(&mut name), None, ptr::null_mut(), None) }?;
        let name = name.to_string();
        println!("name is {name}");
        Ok(TypeAttr {
            typeinfo,
            name,
            type_attr: NonNull::new(type_attr).unwrap(),
        })
    }
    pub fn name(&self) -> &str {
        &self.name
    }
}

impl PartialEq for TypeAttr {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
    }
}

impl std::ops::Deref for TypeAttr {
    type Target = TYPEATTR;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.type_attr.as_ptr() }
    }
}

impl Drop for TypeAttr {
    fn drop(&mut self) {
        unsafe {
            self.typeinfo.ReleaseTypeAttr(self.type_attr.as_ptr());
        }
    }
}

pub struct TypeInfos<'a> {
    typelib: &'a ITypeLib,
    count: u32,
    index: u32,
}

impl<'a> From<&'a ITypeLib> for TypeInfos<'a> {
    fn from(typelib: &'a ITypeLib) -> Self {
        TypeInfos {
            typelib,
            count: unsafe { typelib.GetTypeInfoCount() },
            index: 0,
        }
    }
}

impl<'a> Iterator for TypeInfos<'a> {
    type Item = std::result::Result<ITypeInfo, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.count {
            return None;
        }

        let result = unsafe { self.typelib.GetTypeInfo(self.index) };
        self.index += 1;
        Some(result)
    }
}

pub struct OleClassNames<'a> {
    typelib: &'a ITypeLib,
    count: u32,
    index: u32,
}

impl<'a> From<&'a ITypeLib> for OleClassNames<'a> {
    fn from(typelib: &'a ITypeLib) -> Self {
        OleClassNames {
            typelib,
            count: unsafe { typelib.GetTypeInfoCount() },
            index: 0,
        }
    }
}

impl<'a> Iterator for OleClassNames<'a> {
    type Item = std::result::Result<String, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.count {
            return None;
        }

        let mut bstrname = BSTR::default();
        let result = unsafe {
            self.typelib.GetDocumentation(
                self.index as i32,
                Some(&mut bstrname),
                None,
                ptr::null_mut(),
                None,
            )
        };
        self.index += 1;
        if let Err(error) = result {
            Some(Err(error))
        } else {
            Some(Ok(bstrname.to_string()))
        }
    }
}

pub struct TypeAttrs<'a> {
    typeinfos: &'a mut TypeInfos<'a>,
}

impl<'a> From<&'a mut TypeInfos<'a>> for TypeAttrs<'a> {
    fn from(typeinfos: &'a mut TypeInfos<'a>) -> Self {
        TypeAttrs { typeinfos }
    }
}

impl<'a> Iterator for TypeAttrs<'a> {
    type Item = std::result::Result<TypeAttr, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        let item = self.typeinfos.next();
        if let Some(result) = item {
            if let Ok(typeinfo) = result {
                Some(TypeAttr::new(typeinfo))
            } else {
                Some(Err(result.unwrap_err()))
            }
        } else {
            None
        }
    }
}

pub struct Var<'a> {
    type_attr: &'a TypeAttr,
    index: u32,
    name: String,
    var_desc: std::ptr::NonNull<VARDESC>,
}

impl<'a> Drop for Var<'a> {
    fn drop(&mut self) {
        unsafe {
            self.type_attr
                .typeinfo
                .ReleaseVarDesc(self.var_desc.as_ptr())
        };
    }
}

impl<'a> Var<'a> {
    pub fn new(
        type_attr: &'a TypeAttr,
        index: u32,
    ) -> std::result::Result<Var, windows::core::Error> {
        let var_desc = unsafe { type_attr.typeinfo.GetVarDesc(index) }?;
        let mut rgbstrnames = BSTR::default();
        let mut pcnames = 0;
        unsafe {
            type_attr
                .typeinfo
                .GetNames((*var_desc).memid, &mut rgbstrnames, 1, &mut pcnames)
        }?;
        Ok(Var {
            type_attr,
            index,
            name: rgbstrnames.to_string(),
            var_desc: std::ptr::NonNull::new(var_desc).unwrap(),
        })
    }
    pub fn name(&self) -> &str {
        &self.name
    }
}

impl<'a> PartialEq for Var<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name && self.index == other.index
    }
}

pub struct IterVars<'a> {
    type_attr: &'a TypeAttr,
    count: u16,
    index: u16,
}

impl<'a> From<&'a TypeAttr> for IterVars<'a> {
    fn from(type_attr: &'a TypeAttr) -> Self {
        IterVars {
            type_attr,
            count: type_attr.cVars,
            index: 0,
        }
    }
}

impl<'a> Iterator for IterVars<'a> {
    type Item = std::result::Result<Var<'a>, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.count {
            return None;
        }

        let result = Var::new(self.type_attr, self.index as u32);
        self.index += 1;
        Some(result)
    }
}

pub struct Method<'a> {
    type_attr: &'a TypeAttr,
    index: u32,
    name: String,
    func_desc: std::ptr::NonNull<FUNCDESC>,
    params: Vec<Param>,
}

impl<'a> Method<'a> {
    pub fn new(
        type_attr: &'a TypeAttr,
        index: u32,
    ) -> std::result::Result<Method, windows::core::Error> {
        let func_desc = unsafe { type_attr.typeinfo.GetFuncDesc(index) }?;
        let mut len = 0;
        let mut rgbstrnames = vec![BSTR::default(); unsafe { (*func_desc).cParams } as usize + 1];
        unsafe {
            type_attr.typeinfo.GetNames(
                (*func_desc).memid,
                rgbstrnames.as_mut_ptr(),
                (*func_desc).cParams as u32 + 1,
                &mut len,
            )
        }?;
        let name = rgbstrnames[0].to_string();
        let mut params = vec![];

        if unsafe { (*func_desc).cParams } > 0 {
            for i in 1..len {
                let elem_desc = unsafe { (*func_desc).lprgelemdescParam.offset(i as isize) };
                let elem_desc = std::ptr::NonNull::new(elem_desc).unwrap();
                let param = Param {
                    name: rgbstrnames[i as usize].to_string(),
                    elem_desc,
                };
                params.push(param);
            }
        }
        Ok(Method {
            type_attr,
            index,
            name,
            func_desc: std::ptr::NonNull::new(func_desc).unwrap(),
            params,
        })
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn params(&self) -> &[Param] {
        &self.params
    }
}

impl<'a> PartialEq for Method<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name && self.index == other.index
    }
}

impl<'a> Drop for Method<'a> {
    fn drop(&mut self) {
        unsafe {
            self.type_attr
                .typeinfo
                .ReleaseFuncDesc(self.func_desc.as_ptr())
        };
    }
}

pub struct IterMethods<'a> {
    type_attr: &'a TypeAttr,
    count: u16,
    index: u16,
}

impl<'a> From<&'a TypeAttr> for IterMethods<'a> {
    fn from(type_attr: &'a TypeAttr) -> Self {
        IterMethods {
            type_attr,
            count: type_attr.cFuncs,
            index: 0,
        }
    }
}

impl<'a> Iterator for IterMethods<'a> {
    type Item = std::result::Result<Method<'a>, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.count {
            return None;
        }

        let result = Method::new(self.type_attr, self.index as u32);
        self.index += 1;
        Some(result)
    }
}

pub struct Param {
    name: String,
    elem_desc: std::ptr::NonNull<ELEMDESC>,
}

impl Param {
    pub fn name(&self) -> &str {
        &self.name
    }
}

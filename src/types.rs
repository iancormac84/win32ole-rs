use std::ptr::{self, NonNull};

use windows::{
    core::BSTR,
    Win32::System::Com::{
        ITypeInfo, ITypeLib, FUNCDESC, IMPLTYPEFLAGS, IMPLTYPEFLAG_FSOURCE, TYPEATTR,
    },
};

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
        if self.index == self.count {
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
        if self.index == self.count {
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

#[derive(Debug)]
pub struct TypeImplDesc {
    typeinfo: ITypeInfo,
    index: u32,
    impl_type_flags: IMPLTYPEFLAGS,
}
impl TypeImplDesc {
    pub fn new(typeinfo: ITypeInfo, index: u32, impl_type_flags: IMPLTYPEFLAGS) -> Self {
        TypeImplDesc {
            typeinfo,
            index,
            impl_type_flags,
        }
    }
    pub fn typeinfo(&self) -> &ITypeInfo {
        &self.typeinfo
    }
    pub fn into_typeinfo(self) -> ITypeInfo {
        self.typeinfo
    }
    pub fn is_source(&self) -> bool {
        self.impl_type_flags & IMPLTYPEFLAG_FSOURCE != IMPLTYPEFLAGS(0)
    }
    pub fn matches(&self, flags: IMPLTYPEFLAGS) -> bool {
        (self.impl_type_flags & flags) == flags
    }
    pub fn name(&self) -> windows::core::Result<String> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index)? };
        let mut bstrname = BSTR::default();
        let result = unsafe {
            self.typeinfo.GetDocumentation(
                (*funcdesc).memid,
                Some(&mut bstrname),
                None,
                ptr::null_mut(),
                None,
            )
        };
        if result.is_err() {
            unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
            return Err(result.unwrap_err());
        }
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(bstrname.to_string())
    }
}

pub struct ReferencedTypes<'a> {
    typeinfo: &'a ITypeInfo,
    count: u16,
    index: u16,
    method_index: u32,
}

impl<'a> ReferencedTypes<'a> {
    pub fn new(typeinfo: &'a ITypeInfo, attributes: &TYPEATTR, method_index: u32) -> Self {
        ReferencedTypes {
            typeinfo,
            count: attributes.cImplTypes,
            index: 0,
            method_index,
        }
    }
}

impl<'a> Iterator for ReferencedTypes<'a> {
    type Item = Result<TypeImplDesc, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index == self.count {
            return None;
        }

        unsafe {
            let impl_type_flags = self.typeinfo.GetImplTypeFlags(self.index as u32);
            let Ok(impl_type_flags) = impl_type_flags else {
                return Some(Err(impl_type_flags.unwrap_err()));
            };
            let ref_type = self.typeinfo.GetRefTypeOfImplType(self.index as u32);
            let Ok(ref_type) = ref_type else {
                return Some(Err(ref_type.unwrap_err()));
            };
            let ref_type_info = self.typeinfo.GetRefTypeInfo(ref_type);
            let Ok(ref_type_info) = ref_type_info else {
                return Some(Err(ref_type_info.unwrap_err()));
            };

            self.index += 1;

            Some(Ok(TypeImplDesc::new(
                ref_type_info,
                self.method_index,
                impl_type_flags,
            )))
        }
    }
}

pub struct Method {
    typeinfo: ITypeInfo,
    func_desc: NonNull<FUNCDESC>,
    bstrname: BSTR,
}

impl Method {
    pub fn name(&self) -> &BSTR {
        &self.bstrname
    }
    pub fn deconstruct(self) -> (ITypeInfo, NonNull<FUNCDESC>, BSTR) {
        (self.typeinfo, self.func_desc, self.bstrname)
    }
}

pub struct Methods<'a> {
    typeinfo: &'a ITypeInfo,
    type_attr: NonNull<TYPEATTR>,
    count: u16,
    index: u16,
}

impl<'a> Methods<'a> {
    pub fn new(typeinfo: &'a ITypeInfo) -> windows::core::Result<Self> {
        let type_attr = unsafe { typeinfo.GetTypeAttr()? };
        let type_attr = NonNull::new(type_attr).unwrap();
        let count = unsafe { type_attr.as_ref().cFuncs };
        Ok(Methods {
            typeinfo,
            type_attr,
            count,
            index: 0,
        })
    }
}

impl<'a> Iterator for Methods<'a> {
    type Item = Result<Method, windows::core::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index == self.count {
            return None;
        }

        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index as u32) };
        let Ok(funcdesc) = funcdesc else {
            return Some(Err(funcdesc.unwrap_err()));
        };
        let mut bstrname = BSTR::default();
        let result = unsafe {
            self.typeinfo.GetDocumentation(
                (*funcdesc).memid,
                Some(&mut bstrname),
                None,
                ptr::null_mut(),
                None,
            )
        };
        if result.is_err() {
            unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
            return Some(Err(result.unwrap_err()));
        }
        self.index += 1;
        Some(Ok(Method {
            typeinfo: self.typeinfo.clone(),
            func_desc: NonNull::new(funcdesc).unwrap(),
            bstrname,
        }))
    }
}

impl<'a> Drop for Methods<'a> {
    fn drop(&mut self) {
        unsafe {
            self.typeinfo.ReleaseTypeAttr(self.type_attr.as_ptr());
        }
    }
}
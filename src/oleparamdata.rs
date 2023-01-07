use std::ptr::NonNull;

use windows::Win32::System::{
    Com::{ITypeInfo, FUNCDESC},
    Ole::{PARAMFLAG_FIN, PARAMFLAG_FOPT, PARAMFLAG_FOUT},
};

use crate::{
    error::{Error, Result},
    util::ole::ole_typedesc2val,
    OleMethodData,
};

pub struct OleParamData<'a> {
    typeinfo: &'a ITypeInfo,
    method_index: u32,
    index: u32,
    name: String,
    func_desc: NonNull<FUNCDESC>,
}

impl<'a> OleParamData<'a> {
    pub fn new(olemethod: &OleMethodData, n: u32) -> Result<OleParamData> {
        oleparam_ole_param_from_index(olemethod.typeinfo(), olemethod.index(), n as i32)
    }
    pub fn make(
        olemethod: &OleMethodData,
        method_index: u32,
        index: u32,
        name: String,
    ) -> Result<OleParamData> {
        let typeinfo = olemethod.typeinfo();
        let func_desc = unsafe { typeinfo.GetFuncDesc(method_index) }?;
        let func_desc = NonNull::new(func_desc).unwrap();

        Ok(OleParamData {
            typeinfo,
            method_index,
            index,
            name,
            func_desc,
        })
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn method_index(&self) -> u32 {
        self.method_index
    }
    pub fn index(&self) -> u32 {
        self.index
    }
    pub fn ole_type(&self) -> Result<String> {
        let typ = ole_typedesc2val(
            &self.typeinfo,
            unsafe {
                &(*(self.func_desc.as_ref())
                    .lprgelemdescParam
                    .offset(self.index as isize))
                .tdesc
            },
            None,
        );
        Ok(typ)
    }
    pub fn ole_type_detail(&self) -> Result<Vec<String>> {
        let mut typedetails = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            unsafe {
                &(*(self.func_desc.as_ref())
                    .lprgelemdescParam
                    .offset(self.index as isize))
                .tdesc
            },
            Some(&mut typedetails),
        );
        Ok(typedetails)
    }
    fn ole_param_flag_mask(&self, mask: u16) -> bool {
        let ret = unsafe {
            &(*(self.func_desc.as_ref())
                .lprgelemdescParam
                .offset(self.index as isize))
            .Anonymous
            .paramdesc
            .wParamFlags
            .0
        } & mask
            != 0;
        ret
    }
    pub fn input(&self) -> bool {
        self.ole_param_flag_mask(PARAMFLAG_FIN.0)
    }
    pub fn output(&self) -> bool {
        self.ole_param_flag_mask(PARAMFLAG_FOUT.0)
    }
    pub fn optional(&self) -> bool {
        self.ole_param_flag_mask(PARAMFLAG_FOPT.0)
    }
    pub fn retval(&self) -> bool {
        self.ole_param_flag_mask(PARAMFLAG_FOPT.0)
    }
    /*pub fn default_val<T>(&self) -> Option<T> {
        let mask = PARAMFLAGS(PARAMFLAG_FOPT.0 | PARAMFLAG_FHASDEFAULT.0);
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index) };
        let funcdesc = if let Ok(funcdesc) = funcdesc {
            funcdesc
        } else {
            return None;
        };
        let elemdesc = unsafe { (*funcdesc).lprgelemdescParam.offset(self.index as isize) };
        let paramflags = unsafe { (*elemdesc).Anonymous.paramdesc.wParamFlags };
        let mut defval = None;
        if paramflags & mask == mask {
            let paramdescex = unsafe { (*elemdesc).Anonymous.paramdesc.pparamdescex };
            defval = ole_variant2val(unsafe { &(*paramdescex).varDefaultValue });
        }
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        defval
    }*/
}

impl<'a> Drop for OleParamData<'a> {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseFuncDesc(self.func_desc.as_ptr()) };
    }
}

fn oleparam_ole_param_from_index(
    typeinfo: &ITypeInfo,
    method_index: u32,
    param_index: i32,
) -> Result<OleParamData> {
    let func_desc = unsafe { typeinfo.GetFuncDesc(method_index) }?;
    let func_desc = NonNull::new(func_desc).unwrap();

    let mut len = 0;
    let cmaxnames = unsafe { func_desc.as_ref() }.cParams + 1;
    let mut bstrs = Vec::with_capacity(cmaxnames as usize);
    let result = unsafe {
        typeinfo.GetNames(
            func_desc.as_ref().memid,
            bstrs.as_mut_ptr(),
            cmaxnames as u32,
            &mut len,
        )
    };
    if let Err(error) = result {
        unsafe { typeinfo.ReleaseFuncDesc(func_desc.as_ptr()) };
        return Err(Error::Custom(format!(
            "ITypeInfo::GetNames call failed: {error}"
        )));
    }
    bstrs.remove(0);
    if param_index < 1 || len <= param_index as u32 {
        unsafe { typeinfo.ReleaseFuncDesc(func_desc.as_ptr()) };
        return Err(Error::Custom(format!("index of param must be in 1..{len}")));
    }

    Ok(OleParamData {
        typeinfo,
        method_index,
        index: param_index as u32 - 1,
        name: bstrs[param_index as usize].to_string(),
        func_desc,
    })
}

use windows::Win32::System::{
    Com::ITypeInfo,
    Ole::{PARAMFLAG_FIN, PARAMFLAG_FOPT, PARAMFLAG_FOUT},
};

use crate::{
    error::{Error, Result},
    util::ole::ole_typedesc2val,
    OleMethodData,
};

pub struct OleParamData {
    typeinfo: ITypeInfo,
    method_index: u32,
    index: u32,
    name: String,
}

impl OleParamData {
    pub fn new(olemethod: &OleMethodData, n: u32) -> Result<OleParamData> {
        oleparam_ole_param_from_index(&olemethod.typeinfo, olemethod.index, n as i32)
    }
    pub fn make(
        olemethod: &OleMethodData,
        method_index: u32,
        index: u32,
        name: String,
    ) -> OleParamData {
        OleParamData {
            typeinfo: olemethod.typeinfo.clone(),
            method_index,
            index,
            name,
        }
    }
    pub fn name(&self) -> &str {
        &self.name[..]
    }
    pub fn ole_type(&self) -> Result<String> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.method_index) }?;
        let typ = ole_typedesc2val(
            &self.typeinfo,
            unsafe { &(*(*funcdesc).lprgelemdescParam.offset(self.index as isize)).tdesc },
            None,
        );
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(typ)
    }
    pub fn ole_type_detail(&self) -> Result<Vec<String>> {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.method_index) }?;
        let mut typedetails = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            unsafe { &(*(*funcdesc).lprgelemdescParam.offset(self.index as isize)).tdesc },
            Some(&mut typedetails),
        );
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
        Ok(typedetails)
    }
    fn ole_param_flag_mask(&self, mask: u16) -> bool {
        let funcdesc = unsafe { self.typeinfo.GetFuncDesc(self.index) };
        let funcdesc = if let Ok(funcdesc) = funcdesc {
            funcdesc
        } else {
            return false;
        };
        let ret = unsafe {
            &(*(*funcdesc).lprgelemdescParam.offset(self.index as isize))
                .Anonymous
                .paramdesc
                .wParamFlags
                .0
        } & mask
            != 0;
        unsafe { self.typeinfo.ReleaseFuncDesc(funcdesc) };
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

fn oleparam_ole_param_from_index(
    typeinfo: &ITypeInfo,
    method_index: u32,
    param_index: i32,
) -> Result<OleParamData> {
    let funcdesc = unsafe { typeinfo.GetFuncDesc(method_index) }?;

    let mut len = 0;
    let mut bstrs = Vec::with_capacity(unsafe { (*funcdesc).cParams } as usize + 1);
    let result = unsafe {
        typeinfo.GetNames(
            (*funcdesc).memid,
            bstrs.as_mut_ptr(),
            (*funcdesc).cParams as u32 + 1,
            &mut len,
        )
    };
    if let Err(error) = result {
        unsafe { typeinfo.ReleaseFuncDesc(funcdesc) };
        return Err(Error::Custom(format!(
            "ITypeInfo::GetNames call failed: {error}"
        )));
    }
    bstrs.remove(0);
    if param_index < 1 || len <= param_index as u32 {
        unsafe { typeinfo.ReleaseFuncDesc(funcdesc) };
        return Err(Error::Custom(format!("index of param must be in 1..{len}")));
    }

    let param = OleParamData {
        typeinfo: typeinfo.clone(),
        method_index,
        index: param_index as u32 - 1,
        name: bstrs[param_index as usize].to_string(),
    };

    unsafe { typeinfo.ReleaseFuncDesc(funcdesc) };
    Ok(param)
}

use std::ptr::NonNull;

use windows::{
    core::BSTR,
    Win32::System::{
        Com::{ITypeInfo, ELEMDESC, FUNCDESC},
        Ole::{PARAMFLAGS, PARAMFLAG_FHASDEFAULT, PARAMFLAG_FIN, PARAMFLAG_FOPT, PARAMFLAG_FOUT, PARAMFLAG_FRETVAL},
    },
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
    func_desc: NonNull<FUNCDESC>,
}

impl OleParamData {
    pub fn new(olemethod: OleMethodData, n: u32) -> Result<OleParamData> {
        oleparam_ole_param_from_index(olemethod.typeinfo(), olemethod.index(), n as i32)
    }
    pub fn make(
        olemethod: &OleMethodData,
        method_index: u32,
        index: u32,
        name: String,
    ) -> Result<OleParamData> {
        let typeinfo = olemethod.typeinfo().clone();
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
        Ok(ole_typedesc2val(
            &self.typeinfo,
            unsafe {
                &(*(self.func_desc.as_ref())
                    .lprgelemdescParam
                    .offset(self.index as isize))
                .tdesc
            },
            None,
        ))
    }
    pub fn ole_type_detail(&self) -> Result<Vec<String>> {
        let mut typedetails = vec![];
        ole_typedesc2val(
            &self.typeinfo,
            &(unsafe {
                *(self.func_desc.as_ref())
                    .lprgelemdescParam
                    .offset(self.index as isize)
            })
            .tdesc,
            Some(&mut typedetails),
        );
        Ok(typedetails)
    }
    pub fn param_flags(&self) -> PARAMFLAGS {
        unsafe {
            (*(self.func_desc.as_ref())
                .lprgelemdescParam
                .offset(self.index as isize))
            .Anonymous
            .paramdesc
            .wParamFlags
        }
    }
    fn ole_param_flag_mask(&self, mask: u16) -> bool {
        let paramflags = self.param_flags();
        paramflags & PARAMFLAGS(mask) != PARAMFLAGS(0)
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
        self.ole_param_flag_mask(PARAMFLAG_FRETVAL.0)
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
    pub fn elem_desc(&self) -> &ELEMDESC {
        unsafe {
            &*self
                .func_desc
                .as_ref()
                .lprgelemdescParam
                .offset(self.index as isize)
        }
    }
}

impl Drop for OleParamData {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseFuncDesc(self.func_desc.as_ptr()) };
    }
}

fn oleparam_ole_param_from_index(
    typeinfo: &ITypeInfo,
    method_index: u32,
    param_index: i32,
) -> Result<OleParamData> {
    let typeinfo = typeinfo.clone();
    let func_desc = unsafe { typeinfo.GetFuncDesc(method_index) }?;
    let func_desc = NonNull::new(func_desc).unwrap();

    let cmaxnames = unsafe { func_desc.as_ref() }.cParams as u32 + 1;
    let mut bstrs = vec![BSTR::default(); cmaxnames as usize];
    let mut len = 0;
    let result = unsafe { typeinfo.GetNames(func_desc.as_ref().memid, &mut bstrs, &mut len) };
    if let Err(error) = result {
        unsafe { typeinfo.ReleaseFuncDesc(func_desc.as_ptr()) };
        return Err(Error::Custom(format!(
            "ITypeInfo::GetNames call failed: {error}"
        )));
    }
    if param_index < 1 || len <= param_index as u32 {
        unsafe { typeinfo.ReleaseFuncDesc(func_desc.as_ptr()) };
        return Err(Error::Custom(format!(
            "index of param must be in the range 1..{}",
            bstrs.len()
        )));
    }

    let name = bstrs[param_index as usize].to_string();
    Ok(OleParamData {
        typeinfo,
        method_index,
        index: param_index as u32 - 1,
        name,
        func_desc,
    })
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_win32ole_param() {
        let ole_type =
            crate::OleTypeData::new("Microsoft Shell Controls And Automation", "ShellLinkObject");
        assert!(ole_type.is_ok());
        let ole_type = ole_type.unwrap();
        let m_geticonlocation = super::OleMethodData::new(&ole_type, "GetIconLocation");
        assert!(m_geticonlocation.is_ok());
        let m_geticonlocation = m_geticonlocation.unwrap();
        assert!(m_geticonlocation.is_some());
        let m_geticonlocation = m_geticonlocation.unwrap();
        let m_geticonlocation_params = m_geticonlocation.params();
        let m_geticonlocation_param = &m_geticonlocation_params[0];
        assert!(m_geticonlocation_param.is_ok());

        let ole_type1 = crate::OleTypeData::new("Microsoft HTML Object Library", "FontNames");
        assert!(ole_type1.is_ok());
        let ole_type1 = ole_type1.unwrap();
        let m_count = super::OleMethodData::new(&ole_type1, "Count");
        assert!(m_count.is_ok());
        let m_count = m_count.unwrap();
        assert!(m_count.is_some());
        let m_count = m_count.unwrap();
        let m_count_params = m_count.params();
        let m_count_param = &m_count_params[0];
        assert!(m_count_param.is_ok());

        let ole_type2 = crate::OleTypeData::new("Microsoft Scripting Runtime", "FileSystemObject");
        assert!(ole_type2.is_ok());
        let ole_type2 = ole_type2.unwrap();
        let m_copyfile = super::OleMethodData::new(&ole_type2, "CopyFile");
        assert!(m_copyfile.is_ok());
        let m_copyfile = m_copyfile.unwrap();
        assert!(m_copyfile.is_some());
        let m_copyfile = m_copyfile.unwrap();
        let m_copyfile_params = m_copyfile.params();
        let param_source = &m_copyfile_params[0];
        assert!(param_source.is_ok());
        let param_overwritefiles = &m_copyfile_params[2];
        assert!(param_overwritefiles.is_ok());

        let ole_type3 = crate::OleTypeData::new("Microsoft Scripting Runtime", "Dictionary");
        assert!(ole_type3.is_ok());
        let ole_type3 = ole_type3.unwrap();
        let m_add = super::OleMethodData::new(&ole_type3, "Add");
        assert!(m_add.is_ok());
        let m_add = m_add.unwrap();
        assert!(m_add.is_some());
        let m_add = m_add.unwrap();
        let m_add_params = m_add.params();
        let param_key = &m_add_params[0];
        assert!(param_key.is_ok());

        let param = super::OleParamData::new(m_copyfile, 3);
        assert!(param.is_ok());
        let param = param.unwrap();
        assert_eq!(param.name(), "OverWriteFiles");
        //assert_eq!(WIN32OLE::Param, param.class());
        //assert_eq!(true, param.default());

        assert_eq!(param_source.as_ref().unwrap().name(), "Source");
        assert_eq!(param_key.as_ref().unwrap().name(), "Key");

        let param_source_ole_type = param_source.as_ref().unwrap().ole_type();
        assert!(param_source_ole_type.is_ok());
        let param_source_ole_type = param_source_ole_type.unwrap();
        assert_eq!(param_source_ole_type, "BSTR");
        let param_key_ole_type = param_key.as_ref().unwrap().ole_type();
        assert!(param_key_ole_type.is_ok());
        let param_key_ole_type = param_key_ole_type.unwrap();
        assert_eq!(param_key_ole_type, "VARIANT");

        let param_source_ole_type_detail = param_source.as_ref().unwrap().ole_type_detail();
        assert!(param_source_ole_type_detail.is_ok());
        let param_source_ole_type_detail = param_source_ole_type_detail.unwrap();
        assert_eq!(param_source_ole_type_detail, ["BSTR"]);
        let param_key_ole_type_detail = param_key.as_ref().unwrap().ole_type_detail();
        assert!(param_key_ole_type_detail.is_ok());
        let param_key_ole_type_detail = param_key_ole_type_detail.unwrap();
        assert_eq!(param_key_ole_type_detail, ["PTR", "VARIANT"]);

        let param_source_input = param_source.as_ref().unwrap().input();
        assert_eq!(param_source_input, true);
        let m_geticonlocation_param_input = m_geticonlocation_param.as_ref().unwrap().input();
        assert_eq!(m_geticonlocation_param_input, false);

        let param_source_output = param_source.as_ref().unwrap().output();
        assert_eq!(param_source_output, false);
        let m_geticonlocation_param_output = m_geticonlocation_param.as_ref().unwrap().output();
        assert_eq!(m_geticonlocation_param_output, true);

        let param_source_optional = param_source.as_ref().unwrap().optional();
        assert_eq!(param_source_optional, false);
        let param_overwritefiles_optional = param_overwritefiles.as_ref().unwrap().optional();
        assert_eq!(param_overwritefiles_optional, true);

        let param_source_retval = param_source.as_ref().unwrap().retval();
        assert_eq!(param_source_retval, false);
        let m_count_param_retval = m_count_param.as_ref().unwrap().retval();
        assert_eq!(m_count_param_retval, true);

        /*assert_eq!(param_source.default, nil);
        assert_eq!(param_overwritefiles.default, true);*/
    }
}

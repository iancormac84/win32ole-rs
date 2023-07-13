use clap::{arg, command, Parser};
use win32ole::{error::Error, ole_initialized, types::TypeInfos, OleTypeData, TypeRef};
use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::{TYPE_E_CANTLOADLIBRARY, TYPE_E_ELEMENTNOTFOUND},
        System::{
            Com::{
                FUNC_DISPATCH, FUNC_STATIC, INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
                INVOKE_PROPERTYPUTREF, SAFEARRAYBOUND, TKIND_ALIAS, TKIND_COCLASS, TKIND_DISPATCH,
                TKIND_ENUM, TKIND_INTERFACE, TKIND_MODULE, TKIND_RECORD, TKIND_UNION, TYPEDESC,
            },
            Ole::{LoadTypeLibEx, PARAMFLAGS, PARAMFLAG_FIN, PARAMFLAG_FOUT, REGKIND_NONE},
            Variant::{
                VARENUM, VT_BOOL, VT_BSTR, VT_BYREF, VT_CARRAY, VT_CY, VT_DATE, VT_DECIMAL,
                VT_DISPATCH, VT_ERROR, VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_LPSTR,
                VT_LPWSTR, VT_PTR, VT_R4, VT_R8, VT_SAFEARRAY, VT_UI1, VT_UI2, VT_UI4, VT_UI8,
                VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT, VT_VOID,
            },
        },
    },
};

/// The result of running [`build`]
#[derive(Debug)]
pub struct BuildResult {
    /// The number of referenced types that could not be found and were replaced with `__missing_type__`
    pub num_missing_types: usize,

    /// The number of types that could not be found
    pub num_types_not_found: usize,

    /// The number of dispinterfaces that were skipped because the `emit_dispinterfaces` parameter of [`build`] was false
    pub skipped_dispinterfaces: Vec<String>,

    /// The number of dual interfaces whose dispinterface half was skipped
    pub skipped_dispinterface_of_dual_interfaces: Vec<String>,
}

/// Parses the typelib (or DLL with embedded typelib resource) at the given path and emits bindings to the given writer.
pub fn build<W>(
    filename: &std::path::Path,
    emit_dispinterfaces: bool,
    mut out: W,
) -> Result<BuildResult, Error>
where
    W: std::io::Write,
{
    let mut build_result = BuildResult {
        num_missing_types: 0,
        num_types_not_found: 0,
        skipped_dispinterfaces: vec![],
        skipped_dispinterface_of_dual_interfaces: vec![],
    };

    let filename = os_str_to_wstring(filename.as_os_str());

    ole_initialized();
    unsafe {
        let typelib = LoadTypeLibEx(PCWSTR::from_raw(filename.as_ptr()), REGKIND_NONE)?;

        let typeinfos = TypeInfos::from(&typelib);

        for typeinfo in typeinfos {
            let typeinfo = match typeinfo {
                Ok(typeinfo) => OleTypeData::try_from(typeinfo)?,
                Err(error) => {
                    if error == windows::core::Error::from(TYPE_E_CANTLOADLIBRARY) {
                        build_result.num_types_not_found += 1;
                        continue;
                    } else {
                        return Err(error.into());
                    }
                }
            };

            let typeinfo = if typeinfo.attribs().typekind == TKIND_DISPATCH {
                // Get dispinterface half of this interface if it's a dual interface
                // TODO: Also emit codegen for dispinterface side?
                match typeinfo.get_interface_of_dispinterface() {
                    Ok(disp_type_info) => {
                        build_result
                            .skipped_dispinterface_of_dual_interfaces
                            .push(typeinfo.name().to_string());
                        disp_type_info
                    }
                    Err(error) => match error {
                        Error::Windows(ref winerror) => {
                            if winerror == &windows::core::Error::from(TYPE_E_ELEMENTNOTFOUND) {
                                typeinfo // Not a dual interface
                            } else {
                                return Err(error);
                            }
                        }
                        _ => return Err(error),
                    },
                }
            } else {
                typeinfo
            };

            let attributes = typeinfo.attribs();
            let type_name = typeinfo.name();

            match attributes.typekind {
                TKIND_ENUM => {
                    let type_name = type_name.replace("tag", "");
                    write!(out, "pub struct {type_name}(pub ")?;

                    for (count, member) in typeinfo.variables().into_iter().enumerate() {
                        let member = member?;
                        let value = member.variant();
                        let wkt_str = well_known_type_to_string((*value).Anonymous.Anonymous.vt);
                        if count == 0 {
                            writeln!(out, "{});", wkt_str)?;
                        }
                        let real_value = match (*value).Anonymous.Anonymous.vt {
                            VT_I4 => (*value).Anonymous.Anonymous.Anonymous.lVal,
                            _ => unreachable!(),
                        };

                        write!(
                            out,
                            "pub const {}: {type_name} = {type_name}({real_value}{wkt_str});\n",
                            member.name()
                        )?;
                    }

                    writeln!(out)?;
                }

                TKIND_RECORD => {
                    let type_name = type_name.replace("tag", "");
                    writeln!(out, "#[repr(C)]\npub struct {type_name} {{")?;

                    let mut debug_str = format!("impl ::core::fmt::Debug for {type_name} {{\n    fn fmt(&self, f: &mut ::core::fmt::Formatter<'_>) -> ::core::fmt::Result {{\n        f.debug_struct({type_name:?})");
                    for field in typeinfo.variables() {
                        let field = field?;
                        let type_string = type_to_string(
                            field.typedesc(),
                            PARAMFLAG_FOUT,
                            &typeinfo,
                            &mut build_result,
                        )?;
                        let field_name = sanitize_reserved(field.name());
                        writeln!(out, "    pub {field_name}: {type_string},")?;
                        let f = format!(".field({field_name:?}, &self.{field_name})");
                        debug_str.push_str(&f);
                    }
                    debug_str.push_str(".finish()\n    }\n}\n");

                    writeln!(out, "}}")?;
                    writeln!(out, "impl ::core::marker::Copy for {type_name} {{}}\nimpl ::core::clone::Clone for {type_name} {{\n    fn clone(&self) -> Self {{\n        *self\n    }}\n}}\n{debug_str}unsafe impl ::windows::core::Abi for {type_name} {{\n    type Abi = Self;\n}}")?;
                    writeln!(out)?;
                }

                TKIND_MODULE => {
                    for function in typeinfo.ole_methods()? {
                        let function_desc = function.desc();

                        assert_eq!(function_desc.funckind, FUNC_STATIC);

                        let function_name = function.name();

                        writeln!(out, r#"extern "system" pub fn {function_name}("#)?;

                        for param in function.params() {
                            let param = param?;
                            let param_desc = param.typedesc();
                            let param_name = sanitize_reserved(param.name());
                            let type_string = type_to_string(
                                param_desc,
                                param.param_flags(),
                                &typeinfo,
                                &mut build_result,
                            )?;
                            writeln!(out, "    {param_name}: {type_string},")?;
                        }

                        let type_string = type_to_string(
                            &function_desc.elemdescFunc.tdesc,
                            PARAMFLAG_FOUT,
                            &typeinfo,
                            &mut build_result,
                        )?;
                        writeln!(out, ") -> {type_string},")?;
                        writeln!(out)?;
                    }

                    writeln!(out)?;
                }

                TKIND_INTERFACE => {
                    writeln!(out, "unsafe impl ::windows::core::Interface for {type_name} {{\n    const IID: ::windows::core::GUID = ::windows::core::GUID::from_u128(0x{:08x}_{:04x}_{:04x}_{:02x}{:02x}_{:02x}{:02x}{:02x}{:02x}{:02x}{:02x});\n}}",
						attributes.guid.data1, attributes.guid.data2, attributes.guid.data3,
						attributes.guid.data4[0], attributes.guid.data4[1], attributes.guid.data4[2], attributes.guid.data4[3],
						attributes.guid.data4[4], attributes.guid.data4[5], attributes.guid.data4[6], attributes.guid.data4[7])?;
                    write!(out, "interface {type_name}({type_name}Vtbl)")?;

                    let mut have_parents = false;
                    let mut parents_vtbl_size = 0;

                    for parent in typeinfo.implemented_ole_types()? {
                        let parent_name = parent.name();

                        if have_parents {
                            write!(out, ", {parent_name}({parent_name}Vtbl)")?;
                        } else {
                            write!(out, ": {parent_name}({parent_name}Vtbl)")?;
                        }
                        have_parents = true;

                        parents_vtbl_size += parent.attribs().cbSizeVft;
                    }

                    writeln!(out, " {{")?;

                    for function in typeinfo.ole_methods()? {
                        let function_desc = function.desc();

                        if (function_desc.oVft as u16) < parents_vtbl_size {
                            // Inherited from ancestors
                            continue;
                        }

                        assert_ne!(function_desc.funckind, FUNC_STATIC);
                        assert_ne!(function_desc.funckind, FUNC_DISPATCH);

                        let function_name = function.name();

                        match function_desc.invkind {
                            INVOKE_FUNC => {
                                writeln!(out, "    fn {function_name}(")?;

                                for param in function.params() {
                                    let param = param?;
                                    let param_desc = param.elem_desc();
                                    let param_name = sanitize_reserved(param.name());
                                    let type_string = type_to_string(
                                        &param_desc.tdesc,
                                        param.param_flags(),
                                        &typeinfo,
                                        &mut build_result,
                                    )?;
                                    writeln!(out, "        {param_name}: {type_string},")?;
                                }

                                let type_string = type_to_string(
                                    &function_desc.elemdescFunc.tdesc,
                                    PARAMFLAG_FOUT,
                                    &typeinfo,
                                    &mut build_result,
                                )?;
                                writeln!(out, "    ) -> {type_string},")?;
                            }

                            INVOKE_PROPERTYGET => {
                                writeln!(out, "    fn get_{function_name}(")?;

                                let mut explicit_ret_val = false;

                                for param in function.params() {
                                    let param = param?;
                                    let param_desc = param.elem_desc();
                                    writeln!(
                                        out,
                                        "        {}: {},",
                                        sanitize_reserved(param.name()),
                                        type_to_string(
                                            &param_desc.tdesc,
                                            param.param_flags(),
                                            &typeinfo,
                                            &mut build_result
                                        )?
                                    )?;

                                    if param.retval() {
                                        assert_eq!(function_desc.elemdescFunc.tdesc.vt, VT_HRESULT);
                                        explicit_ret_val = true;
                                    }
                                }

                                if explicit_ret_val {
                                    assert_eq!(function_desc.elemdescFunc.tdesc.vt, VT_HRESULT);
                                    writeln!(
                                        out,
                                        "    ) -> {},",
                                        type_to_string(
                                            &function_desc.elemdescFunc.tdesc,
                                            PARAMFLAG_FOUT,
                                            &typeinfo,
                                            &mut build_result
                                        )?
                                    )?;
                                } else {
                                    writeln!(
                                        out,
                                        "        value: *mut {},",
                                        type_to_string(
                                            &function_desc.elemdescFunc.tdesc,
                                            PARAMFLAG_FOUT,
                                            &typeinfo,
                                            &mut build_result
                                        )?
                                    )?;
                                    writeln!(
                                        out,
                                        "    ) -> {},",
                                        well_known_type_to_string(VT_HRESULT)
                                    )?;
                                }
                            }

                            INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF => {
                                writeln!(
                                    out,
                                    "    fn {}{}(",
                                    match function_desc.invkind {
                                        INVOKE_PROPERTYPUT => "put_",
                                        INVOKE_PROPERTYPUTREF => "putref_",
                                        _ => unreachable!(),
                                    },
                                    function_name
                                )?;

                                for param in function.params() {
                                    let param = param?;
                                    let param_desc = param.elem_desc();
                                    writeln!(
                                        out,
                                        "        {}: {},",
                                        sanitize_reserved(param.name()),
                                        type_to_string(
                                            &param_desc.tdesc,
                                            param.param_flags(),
                                            &typeinfo,
                                            &mut build_result
                                        )?
                                    )?;
                                }

                                writeln!(
                                    out,
                                    "    ) -> {},",
                                    type_to_string(
                                        &function_desc.elemdescFunc.tdesc,
                                        PARAMFLAG_FOUT,
                                        &typeinfo,
                                        &mut build_result
                                    )?
                                )?;
                            }

                            _ => unreachable!(),
                        }
                    }

                    for property in typeinfo.variables() {
                        let property = property?;

                        // Synthesize get_() and put_() functions for each property.

                        let property_name = sanitize_reserved(property.name());

                        writeln!(out, "    fn get_{property_name}(")?;
                        writeln!(
                            out,
                            "        value: *mut {},",
                            type_to_string(
                                property.typedesc(),
                                PARAMFLAG_FOUT,
                                &typeinfo,
                                &mut build_result
                            )?
                        )?;
                        writeln!(out, "    ) -> {},", well_known_type_to_string(VT_HRESULT))?;
                        writeln!(out, "    fn put_{property_name}(")?;
                        writeln!(
                            out,
                            "        value: {},",
                            type_to_string(
                                property.typedesc(),
                                PARAMFLAG_FIN,
                                &typeinfo,
                                &mut build_result
                            )?
                        )?;
                        writeln!(out, "    ) -> {},", well_known_type_to_string(VT_HRESULT))?;
                    }

                    writeln!(out, "}}}}")?;
                    writeln!(out)?;
                }

                TKIND_DISPATCH => {
                    if !emit_dispinterfaces {
                        build_result
                            .skipped_dispinterfaces
                            .push(typeinfo.name().to_string());
                        continue;
                    }

                    writeln!(out, "unsafe impl ::windows::core::Interface for {type_name} {{\n    const IID: ::windows::core::GUID = ::windows::core::GUID::from_u128(0x{:08x}_{:04x}_{:04x}_{:02x}{:02x}_{:02x}{:02x}{:02x}{:02x}{:02x}{:02x});\n}}",
						attributes.guid.data1, attributes.guid.data2, attributes.guid.data3,
						attributes.guid.data4[0], attributes.guid.data4[1], attributes.guid.data4[2], attributes.guid.data4[3],
						attributes.guid.data4[4], attributes.guid.data4[5], attributes.guid.data4[6], attributes.guid.data4[7])?;
                    writeln!(
                        out,
                        "interface {type_name}({type_name}Vtbl): IDispatch(IDispatchVtbl) {{"
                    )?;
                    writeln!(out, "}}}}")?;

                    {
                        let parents = typeinfo.implemented_ole_types()?;
                        let mut parents_iter = parents.iter();
                        if let Some(parent) = parents_iter.next() {
                            let parent_name = parent.name();
                            assert_eq!(parent_name.to_string(), "IDispatch");
                            assert_eq!(
                                parent.attribs().cbSizeVft as usize,
                                7 * std::mem::size_of::<usize>()
                            ); // 3 from IUnknown + 4 from IDispatch
                        } else {
                            unreachable!();
                        }

                        assert!(parents_iter.next().is_none());
                    }

                    writeln!(out)?;
                    writeln!(out, "impl {type_name} {{")?;

                    // IFaxServerNotify2 lists QueryInterface, etc
                    let has_inherited_functions = typeinfo
                        .ole_methods()?
                        .iter()
                        .any(|function| function.desc().oVft > 0);

                    for function in typeinfo.ole_methods()? {
                        println!("function name is {}", function.name());
                        let function_desc = function.desc();

                        assert_eq!(function_desc.funckind, FUNC_DISPATCH);

                        if has_inherited_functions
                            && (function_desc.oVft as usize) < 7 * std::mem::size_of::<usize>()
                        {
                            continue;
                        }

                        let function_name = function.name();
                        let params: Vec<_> = function
                            .params()
                            .into_iter()
                            .filter_map(|param| param.ok())
                            .filter(|param| !param.retval())
                            .collect();

                        writeln!(
                            out,
                            "    pub unsafe fn {}{}(",
                            match function_desc.invkind {
                                INVOKE_FUNC => "",
                                INVOKE_PROPERTYGET => "get_",
                                INVOKE_PROPERTYPUT => "put_",
                                INVOKE_PROPERTYPUTREF => "putref_",
                                _ => unreachable!(),
                            },
                            function_name
                        )?;

                        writeln!(out, "        &self,")?;

                        for param in &params {
                            let param_desc = param.elem_desc();
                            writeln!(
                                out,
                                "        {}: {},",
                                sanitize_reserved(param.name()),
                                type_to_string(
                                    &param_desc.tdesc,
                                    param.param_flags(),
                                    &typeinfo,
                                    &mut build_result
                                )?
                            )?;
                        }

                        writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;

                        if !params.is_empty() {
                            writeln!(out, "        let mut args: [VARIANT; {}] = [", params.len())?;

                            for param in params.into_iter().rev() {
                                let param_desc = param.elem_desc();
                                if !param.retval() {
                                    let (vt, mutator) = vartype_mutator(
                                        &param_desc.tdesc,
                                        &sanitize_reserved(param.name()),
                                        &typeinfo,
                                    );
                                    writeln!(out, "            {{ let mut v = VARIANT::default(); (*v).Anonymous.Anonymous.vt = VARENUM({}); (*v){}; v }},", vt.0, mutator)?;
                                }
                            }

                            writeln!(out, "        ];")?;
                            writeln!(out)?;
                        }

                        if function_desc.invkind == INVOKE_PROPERTYPUT
                            || function_desc.invkind == INVOKE_PROPERTYPUTREF
                        {
                            writeln!(out, "        let disp_id_put = DISPID_PROPERTYPUT;")?;
                            writeln!(out)?;
                        }

                        writeln!(out, "        let mut result = VARIANT::default();")?;
                        writeln!(out)?;
                        writeln!(
                            out,
                            "        let mut exception_info = EXCEPINFO::default();"
                        )?;
                        writeln!(out)?;
                        writeln!(out, "        let mut error_arg = 0;")?;
                        writeln!(out)?;
                        writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
                        writeln!(
                            out,
                            "            rgvarg: {},",
                            if function_desc.cParams > 0 {
                                "args.as_mut_ptr()"
                            } else {
                                "::core::ptr::null_mut()"
                            }
                        )?;
                        writeln!(
                            out,
                            "            rgdispidNamedArgs: {},",
                            match function_desc.invkind {
                                INVOKE_FUNC | INVOKE_PROPERTYGET => "::core::ptr::null_mut()",
                                INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF => "&disp_id_put",
                                _ => unreachable!(),
                            }
                        )?;
                        writeln!(out, "            cArgs: {},", function_desc.cParams)?;
                        writeln!(
                            out,
                            "            cNamedArgs: {},",
                            match function_desc.invkind {
                                INVOKE_FUNC | INVOKE_PROPERTYGET => "0",
                                INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF => "1",
                                _ => unreachable!(),
                            }
                        )?;
                        writeln!(out, "        }};")?;
                        writeln!(out)?;
                        writeln!(out, "        let hr = IDispatch::Invoke(")?;
                        writeln!(out, "            self,")?;
                        writeln!(
                            out,
                            "            /* dispIdMember */ {},",
                            function_desc.memid
                        )?;
                        writeln!(out, "            /* riid */ &IID_NULL,")?;
                        writeln!(out, "            /* lcid */ 0,")?;
                        writeln!(
                            out,
                            "            /* wFlags */ {},",
                            match function_desc.invkind {
                                INVOKE_FUNC => "DISPATCH_METHOD",
                                INVOKE_PROPERTYGET => "DISPATCH_PROPERTYGET",
                                INVOKE_PROPERTYPUT => "DISPATCH_PROPERTYPUT",
                                INVOKE_PROPERTYPUTREF => "DISPATCH_PROPERTYPUTREF",
                                _ => unreachable!(),
                            }
                        )?;
                        writeln!(out, "            /* pDispParams */ &disp_params,")?;
                        writeln!(out, "            /* pVarResult */ Some(&mut result),")?;
                        writeln!(
                            out,
                            "            /* pExcepInfo */ Some(&mut exception_info),"
                        )?;
                        writeln!(out, "            /* puArgErr */ Some(&mut error_arg),")?;
                        writeln!(out, "        );")?;
                        writeln!(out)?;
                        writeln!(out, "        (hr, result, exception_info, error_arg)")?;
                        writeln!(out, "    }}")?;
                        writeln!(out)?;
                    }

                    for property in typeinfo.variables() {
                        let property = property?;

                        // Synthesize get_() and put_() functions for each property.

                        let property_name = sanitize_reserved(property.name());
                        let type_ = property.typedesc();

                        writeln!(out, "    pub unsafe fn get_{property_name}(")?;
                        writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;
                        writeln!(out, "        let mut result = VARIANT::default();")?;
                        writeln!(out)?;
                        writeln!(
                            out,
                            "        let mut exception_info = EXCEPINFO::default();"
                        )?;
                        writeln!(out)?;
                        writeln!(out, "        let mut error_arg = 0;")?;
                        writeln!(out)?;
                        writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
                        writeln!(out, "            rgvarg: ::core::ptr::null_mut(),")?;
                        writeln!(
                            out,
                            "            rgdispidNamedArgs: ::core::ptr::null_mut(),"
                        )?;
                        writeln!(out, "            cArgs: 0,")?;
                        writeln!(out, "            cNamedArgs: 0,")?;
                        writeln!(out, "        }};")?;
                        writeln!(out)?;
                        writeln!(out, "        let hr = IDispatch::Invoke(")?;
                        writeln!(out, "            self,")?;
                        writeln!(
                            out,
                            "            /* dispIdMember */ {},",
                            property.member_id()
                        )?;
                        writeln!(out, "            /* riid */ &IID_NULL,")?;
                        writeln!(out, "            /* lcid */ 0,")?;
                        writeln!(out, "            /* wFlags */ DISPATCH_PROPERTYGET,")?;
                        writeln!(out, "            /* pDispParams */ &disp_params,")?;
                        writeln!(out, "            /* pVarResult */ Some(&mut result),")?;
                        writeln!(
                            out,
                            "            /* pExcepInfo */ Some(&mut exception_info),"
                        )?;
                        writeln!(out, "            /* puArgErr */ Some(&mut error_arg),")?;
                        writeln!(out, "        );")?;
                        writeln!(out)?;
                        writeln!(out, "        (hr, result, exception_info, error_arg)")?;
                        writeln!(out, "    }}")?;
                        writeln!(out)?;
                        writeln!(out, "    pub unsafe fn put_{property_name}(")?;
                        writeln!(
                            out,
                            "        value: {},",
                            type_to_string(
                                property.typedesc(),
                                PARAMFLAG_FIN,
                                &typeinfo,
                                &mut build_result
                            )?
                        )?;
                        writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;
                        writeln!(out, "        let mut args: [VARIANT; 1] = [")?;
                        let (vt, mutator) = vartype_mutator(type_, "value", &typeinfo);
                        writeln!(out, "            {{ let mut v = VARIANT::default(); (*v).Anonymous.Anonymous.vt = VARENUM({}); (*v){}; v }},", vt.0, mutator)?;
                        writeln!(out, "        ];")?;
                        writeln!(out)?;
                        writeln!(out, "        let mut result = VARIANT::default();")?;
                        writeln!(out)?;
                        writeln!(
                            out,
                            "        let mut exception_info = EXCEPINFO::default();"
                        )?;
                        writeln!(out)?;
                        writeln!(out, "        let mut error_arg = 0;")?;
                        writeln!(out)?;
                        writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
                        writeln!(out, "            rgvarg: args.as_mut_ptr(),")?;
                        writeln!(
                            out,
                            "            rgdispidNamedArgs: ::core::ptr::null_mut(),"
                        )?; // TODO: PROPERTYPUT needs named args?
                        writeln!(out, "            cArgs: 1,")?;
                        writeln!(out, "            cNamedArgs: 0,")?;
                        writeln!(out, "        }};")?;
                        writeln!(out)?;
                        writeln!(out, "        let hr = IDispatch::Invoke(")?;
                        writeln!(out, "            self,")?;
                        writeln!(
                            out,
                            "            /* dispIdMember */ {},",
                            property.member_id()
                        )?;
                        writeln!(out, "            /* riid */ &IID_NULL,")?;
                        writeln!(out, "            /* lcid */ 0,")?;
                        writeln!(out, "            /* wFlags */ DISPATCH_PROPERTYPUT,")?;
                        writeln!(out, "            /* pDispParams */ &disp_params,")?;
                        writeln!(out, "            /* pVarResult */ Some(&mut result),")?;
                        writeln!(
                            out,
                            "            /* pExcepInfo */ Some(&mut exception_info),"
                        )?;
                        writeln!(out, "            /* puArgErr */ Some(&mut error_arg),")?;
                        writeln!(out, "        );")?;
                        writeln!(out)?;
                        // TODO: VariantClear() on args
                        writeln!(out, "        (hr, result, exception_info, error_arg)")?;
                        writeln!(out, "    }}")?;
                        writeln!(out)?;
                    }

                    writeln!(out, "}}")?;
                    writeln!(out)?;
                }

                TKIND_COCLASS => {
                    for parent in typeinfo.implemented_ole_types()? {
                        let parent_name = parent.name();
                        writeln!(out, "// Implements {parent_name}")?;
                    }

                    writeln!(out, "unsafe impl ::windows::core::Interface for {type_name} {{\n    const IID: ::windows::core::GUID = ::windows::core::GUID::from_u128(0x{:08x}_{:04x}_{:04x}_{:02x}{:02x}_{:02x}{:02x}{:02x}{:02x}{:02x}{:02x});\n}}",
						attributes.guid.data1, attributes.guid.data2, attributes.guid.data3,
						attributes.guid.data4[0], attributes.guid.data4[1], attributes.guid.data4[2], attributes.guid.data4[3],
						attributes.guid.data4[4], attributes.guid.data4[5], attributes.guid.data4[6], attributes.guid.data4[7])?;
                    writeln!(out, "class {type_name}; }}")?;
                    writeln!(out)?;
                }

                TKIND_ALIAS => {
                    let type_string = type_to_string(
                        &attributes.tdescAlias,
                        PARAMFLAG_FOUT,
                        &typeinfo,
                        &mut build_result,
                    )?;
                    writeln!(out, "pub type {type_name} = {type_string};")?;
                    writeln!(out)?;
                }

                TKIND_UNION => {
                    let alignment = match attributes.cbAlignment {
                        4 => "u32",
                        8 => "u64",
                        _ => unreachable!(),
                    };

                    let num_aligned_elements =
                        (attributes.cbSizeInstance + attributes.cbAlignment as u32 - 1)
                            / attributes.cbAlignment as u32;
                    assert!(num_aligned_elements > 0);
                    let wrapped_type = match num_aligned_elements {
                        1 => alignment.to_string(),
                        _ => format!("[{alignment}; {num_aligned_elements}]"),
                    };

                    writeln!(out, "UNION2!{{union {type_name} {{")?;
                    writeln!(out, "    {wrapped_type},")?;

                    for field in typeinfo.variables() {
                        let field = field?;

                        let field_name = sanitize_reserved(field.name());
                        writeln!(
                            out,
                            "    {} {}_mut: {},",
                            field_name,
                            field_name,
                            type_to_string(
                                field.typedesc(),
                                PARAMFLAG_FOUT,
                                &typeinfo,
                                &mut build_result
                            )?
                        )?;
                    }

                    writeln!(out, "}}}}")?;
                    writeln!(out)?;
                }

                _ => unreachable!(),
            }
        }
    }

    Ok(build_result)
}

fn os_str_to_wstring(s: &std::ffi::OsStr) -> Vec<u16> {
    let result = std::os::windows::ffi::OsStrExt::encode_wide(s);
    let mut result: Vec<_> = result.collect();
    result.push(0);
    result
}

fn sanitize_reserved(s: &str) -> String {
    let s = s.to_string();
    match s.as_ref() {
        "impl" => "impl_".to_string(),
        "type" => "type_".to_string(),
        _ => s,
    }
}

fn type_to_string(
    type_: &TYPEDESC,
    param_flags: PARAMFLAGS,
    typeinfo: &OleTypeData,
    build_result: &mut BuildResult,
) -> Result<String, Error> {
    match type_.vt {
        VT_PTR => {
            if (param_flags & PARAMFLAG_FIN) == PARAMFLAG_FIN
                && (param_flags & PARAMFLAG_FOUT) == PARAMFLAGS(0)
            {
                // [in] => *const
                type_to_string(
                    unsafe { &*type_.Anonymous.lptdesc },
                    param_flags,
                    typeinfo,
                    build_result,
                )
                .map(|type_name| format!("*const {type_name}"))
            } else {
                // [in, out] => *mut
                // [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
                type_to_string(
                    unsafe { &*type_.Anonymous.lptdesc },
                    param_flags,
                    typeinfo,
                    build_result,
                )
                .map(|type_name| format!("*mut {type_name}"))
            }
        }

        VT_CARRAY => {
            let num_dimensions = unsafe { (*(type_.Anonymous.lpadesc)).cDims as usize };
            let dimensions: &[SAFEARRAYBOUND] = unsafe {
                std::slice::from_raw_parts(
                    (*(type_.Anonymous.lpadesc)).rgbounds.as_ptr(),
                    num_dimensions,
                )
            };

            let mut type_name = type_to_string(
                unsafe { &(*(type_.Anonymous.lpadesc)).tdescElem },
                param_flags,
                typeinfo,
                build_result,
            )?;

            for dimension in dimensions {
                type_name = format!("[{}; {}]", type_name, dimension.cElements);
            }

            Ok(type_name)
        }

        VT_USERDEFINED => match typeinfo
            .get_ref_type_info(unsafe { type_.Anonymous.hreftype })
            .map(|ref_type_info| ref_type_info.name().to_string())
        {
            Ok(ref_type_name) => Ok(ref_type_name),
            Err(error) => match error {
                Error::Windows(ref winerror) => {
                    if winerror == &windows::core::Error::from(TYPE_E_CANTLOADLIBRARY) {
                        build_result.num_types_not_found += 1;
                        Ok("__missing_type__".to_string())
                    } else {
                        Err(error)
                    }
                }
                _ => Err(error),
            },
        },

        _ => Ok(well_known_type_to_string(type_.vt).to_string()),
    }
}

fn well_known_type_to_string(vt: VARENUM) -> &'static str {
    match vt {
        VT_I2 => "i16",
        VT_I4 => "i32",
        VT_R4 => "f32",
        VT_R8 => "f64",
        VT_CY => "CY",
        VT_DATE => "f64",
        VT_BSTR => "::windows::core::PWSTR",
        VT_DISPATCH => "IDispatch",
        VT_ERROR => "i32",
        VT_BOOL => "VARIANT_BOOL",
        VT_VARIANT => "VARIANT",
        VT_UNKNOWN => "::windows::core::IUnknown",
        VT_DECIMAL => "DECIMAL",
        VT_I1 => "i8",
        VT_UI1 => "u8",
        VT_UI2 => "u16",
        VT_UI4 => "u32",
        VT_I8 => "i64",
        VT_UI8 => "u64",
        VT_INT => "i32",
        VT_UINT => "u32",
        VT_VOID => "c_void",
        VT_HRESULT => "::windows::core::HRESULT",
        VT_SAFEARRAY => "SAFEARRAY",
        VT_LPSTR => "::windows::core::PSTR",
        VT_LPWSTR => "::windows::core::PWSTR",
        _ => unreachable!(),
    }
}

fn vartype_mutator(
    type_: &TYPEDESC,
    param_name: &str,
    typeinfo: &OleTypeData,
) -> (VARENUM, String) {
    match type_.vt {
        vt @ VT_I2 => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.iVal = {param_name}"),
        ),
        vt @ VT_I4 => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.lVal = {param_name}"),
        ),
        vt @ VT_CY => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.cyVal = {param_name}"),
        ),
        vt @ VT_BSTR => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.bstrVal = {param_name}"),
        ),
        vt @ VT_DISPATCH => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.pdispVal = {param_name}"),
        ),
        vt @ VT_ERROR => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.scode = {param_name}"),
        ),
        vt @ VT_BOOL => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.boolVal = {param_name}"),
        ),
        vt @ VT_VARIANT => (vt, format!(" = *(&{param_name} as *const _ as *mut _)")),
        vt @ VT_UNKNOWN => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.punkVal = {param_name}"),
        ),
        vt @ VT_UI2 => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.uiVal = {param_name}"),
        ),
        vt @ VT_UI4 => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.ulVal = {param_name}"),
        ),
        vt @ VT_INT => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.intVal = {param_name}"),
        ),
        vt @ VT_UINT => (
            vt,
            format!(".Anonymous.Anonymous.Anonymous.uintVal = {param_name}"),
        ),
        VT_PTR => {
            let pointee_vt = unsafe { (*type_.Anonymous.lptdesc).vt };
            match pointee_vt {
                VT_I4 => (
                    VARENUM(pointee_vt.0 | VT_BYREF.0),
                    format!(".Anonymous.Anonymous.Anonymous.plVal = {param_name}"),
                ),
                VT_BSTR => (
                    VARENUM(pointee_vt.0 | VT_BYREF.0),
                    format!(".Anonymous.Anonymous.Anonymous.pbstrVal = {param_name}"),
                ),
                VT_DISPATCH => (
                    VARENUM(pointee_vt.0 | VT_BYREF.0),
                    format!(".Anonymous.Anonymous.Anonymous.ppdispVal = {param_name}"),
                ),
                VT_BOOL => (
                    VARENUM(pointee_vt.0 | VT_BYREF.0),
                    format!(".Anonymous.Anonymous.Anonymous.pboolVal = {param_name}"),
                ),
                VT_VARIANT => (
                    VARENUM(pointee_vt.0 | VT_BYREF.0),
                    format!(".Anonymous.Anonymous.Anonymous.pvarval = {param_name}"),
                ),
                VT_USERDEFINED => (
                    VT_DISPATCH,
                    format!(".Anonymous.Anonymous.Anonymous.pdispVal = {param_name}"),
                ),
                _ => unreachable!(),
            }
        }
        VT_USERDEFINED => {
            let ref_type = typeinfo
                .get_ref_type_info(unsafe { type_.Anonymous.hreftype })
                .unwrap();
            let size = ref_type.attribs().cbSizeInstance;
            match size {
                4 => (
                    VT_I4,
                    format!(".Anonymous.Anonymous.Anonymous.lVal = {param_name}"),
                ), // enum
                _ => unreachable!(),
            }
        }
        _ => unreachable!(),
    }
}

/// Capture typelib path and emit Rust code to bind to the interfaces defined in the typelib. Optionally emit code for DISPINTERFACES
#[derive(Parser)]
#[command(name = "Options")]
struct Options {
    /// path of typelib
    filename: std::path::PathBuf,
    /// emit code for DISPINTERFACEs (experimental)
    #[arg(long)]
    emit_dispinterfaces: bool,
}

fn main() {
    let args = Options::parse();

    let build_result = {
        let stdout = std::io::stdout();
        build(&args.filename, args.emit_dispinterfaces, stdout.lock()).unwrap()
    };

    if build_result.num_missing_types > 0 {
        eprintln!(
            "{} referenced types could not be found and were replaced with `__missing_type__`",
            build_result.num_missing_types
        );
    }

    if build_result.num_types_not_found > 0 {
        eprintln!(
            "{} types could not be found",
            build_result.num_types_not_found
        );
    }

    for skipped_dispinterface in build_result.skipped_dispinterfaces {
        eprintln!(
            "Dispinterface {skipped_dispinterface} was skipped because --emit-dispinterfaces was not specified"
        );
    }

    for skipped_dispinterface in build_result.skipped_dispinterface_of_dual_interfaces {
        eprintln!("Dispinterface half of dual interface {skipped_dispinterface} was skipped");
    }
}

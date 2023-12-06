use std::{io, sync::LazyLock};
use crate::error::Result;
use winreg::{RegKey, enums::{HKEY_LOCAL_MACHINE, HKEY_CLASSES_ROOT}};

pub mod error;
mod oledata;
//mod oleeventdata;
mod olemethoddata;
mod oleparamdata;
mod oletypedata;
mod oletypelibdata;
mod olevariabledata;
pub mod types;
mod util;
//mod variant;

pub use {
    oledata::OleData,
    olemethoddata::OleMethodData,
    oleparamdata::OleParamData,
    oletypedata::OleTypeData,
    oletypelibdata::{oletypelib_from_guid, OleTypeLibData},
    olevariabledata::OleVariableData,
    util::{
        conv::ToWide,
        ole::{init_runtime, ole_initialized, TypeRef},
    },
};

static G_RUNNING_NANO: LazyLock<bool> = LazyLock::new(|| {
    let hsubkey = RegKey::predef(HKEY_LOCAL_MACHINE)
        .open_subkey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Server\\ServerLevels");
    if let Ok(hsubkey) = hsubkey {
        let result: io::Result<String> = hsubkey.get_value("NanoServer");
        if result.is_ok() {
            return true;
        }
    }
    false
});

pub fn progids() -> Result<Vec<String>> {
    let hclsids = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("CLSID")?;
    let mut progids = vec![];

    for clsid_or_error in hclsids.enum_keys() {
        let clsid = clsid_or_error?;
        let hclsid = hclsids.open_subkey(&clsid);
        if let Ok(hclsid) = hclsid {
            match hclsid.open_subkey("ProgID") {
                Ok(prog_id_key) => {
                    let val: io::Result<String> = prog_id_key.get_value("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val: io::Result<String> = hclsid.get_value("ProgID");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
            }
            match hclsid.open_subkey("VersionIndependentProgID") {
                Ok(version_independent_prog_id_key) => {
                    let val: io::Result<String> = version_independent_prog_id_key.get_value("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val: io::Result<String> = hclsid.get_value("VersionIndependentProgID");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
            }
        } else {
            continue;
        }
    }
    Ok(progids)
}

pub fn typelibs() -> Result<Vec<Result<OleTypeLibData>>> {
    let htypelib = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("TypeLib")?;
    let mut typelibs = vec![];

    for guid_or_error in htypelib.enum_keys() {
        let guid = guid_or_error?;
        let hguid = htypelib.open_subkey(&guid);
        if let Ok(hguid) = hguid {
            for version_or_error in hguid.enum_keys() {
                let version = version_or_error?;
                let hversion = hguid.open_subkey(&version);
                if let Ok(hversion) = hversion {
                    let name: io::Result<String> = hversion.get_value("");
                    let name = if let Ok(name) = name {
                        Ok(name)
                    } else {
                        hversion.get_value(&version)
                    };
                    if let Ok(name) = name {
                        let typelib = oletypelib_from_guid(&guid, &version);
                        if let Ok(typelib) = typelib {
                            typelibs.push(OleTypeLibData::make(typelib, name));
                        }
                    }
                }
            }
        } else {
            continue;
        }
    }

    Ok(typelibs)
}

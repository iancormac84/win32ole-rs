use crate::error::Result;
use util::RegKey;
use windows::Win32::System::Registry::HKEY_CLASSES_ROOT;

pub mod error;
mod oledata;
mod olemethoddata;
mod oleparam;
mod oletypedata;
mod oletypelibdata;
mod util;
//mod variant;

pub use {
    oledata::OleData,
    olemethoddata::OleMethodData,
    oletypedata::OleTypeData,
    oletypelibdata::{oletypelib_from_guid, OleTypeLibData},
    util::conv::ToWide,
};

pub fn progids() -> Result<Vec<String>> {
    let hclsids = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("CLSID")?;
    let mut progids = vec![];

    for clsid_or_error in hclsids.enum_keys() {
        let clsid = clsid_or_error?;
        let hclsid = hclsids.open_subkey(&clsid);
        if let Ok(hclsid) = hclsid {
            match hclsid.open_subkey("ProgID") {
                Ok(prog_id_key) => {
                    let val: Result<String> = prog_id_key.get_value("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val: Result<String> = hclsid.get_value("ProgID");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
            }
            match hclsid.open_subkey("VersionIndependentProgID") {
                Ok(version_independent_prog_id_key) => {
                    let val: Result<String> = version_independent_prog_id_key.get_value("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val: Result<String> = hclsid.get_value("VersionIndependentProgID");
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

pub fn typelibs() -> Result<Vec<OleTypeLibData>> {
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
                    let name = if let Ok(name) = hversion.get_value("") {
                        Ok(name)
                    } else {
                        hversion.get_value(&version)
                    };
                    if let Ok(name) = name {
                        let typelib = oletypelib_from_guid(&guid, &version);
                        if let Ok(typelib) = typelib {
                            typelibs.push(OleTypeLibData { typelib, name });
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

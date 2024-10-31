use std::sync::LazyLock;
use crate::error::Result;
use windows_registry::{LOCAL_MACHINE, CLASSES_ROOT};

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
    let hsubkey = LOCAL_MACHINE
        .open("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Server\\ServerLevels");
    if let Ok(hsubkey) = hsubkey {
        let result = hsubkey.get_string("NanoServer");
        if result.is_ok() {
            return true;
        }
    }
    false
});

pub fn progids() -> Result<Vec<String>> {
    let hclsids = CLASSES_ROOT.open("CLSID")?;
    let mut progids = vec![];

    let clsid_iter = hclsids.keys()?;
    for clsid in clsid_iter {
        let hclsid = hclsids.open(&clsid);
        if let Ok(hclsid) = hclsid {
            match hclsid.open("ProgID") {
                Ok(prog_id_key) => {
                    let val = prog_id_key.get_string("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val = hclsid.get_string("ProgID");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
            }
            match hclsid.open("VersionIndependentProgID") {
                Ok(version_independent_prog_id_key) => {
                    let val = version_independent_prog_id_key.get_string("");
                    if let Ok(val) = val {
                        progids.push(val);
                    }
                }
                Err(_error) => {
                    let val = hclsid.get_string("VersionIndependentProgID");
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
    let htypelib = CLASSES_ROOT.open("TypeLib")?;
    let mut typelibs = vec![];

    let guid_iter = htypelib.keys()?;
    for guid in guid_iter {
        let hguid = htypelib.open(&guid);
        if let Ok(hguid) = hguid {
            let version_iter = hguid.keys()?;
            for version in version_iter {
                let hversion = hguid.open(&version);
                if let Ok(hversion) = hversion {
                    let name = hversion.get_string("");
                    let name = if let Ok(name) = name {
                        Ok(name)
                    } else {
                        hversion.get_string(&version)
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

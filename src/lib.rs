use util::{reg_enum_key, reg_get_val2_string, reg_open_key};
use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::ERROR_SUCCESS,
        System::Registry::{RegCloseKey, HKEY, HKEY_CLASSES_ROOT},
    },
};

pub mod error;
mod oledata;
mod olemethoddata;
mod oleparam;
mod oletypedata;
mod oletypelibdata;
mod util;
//mod variant;

pub use {oledata::OleData, olemethoddata::OleMethodData, oletypedata::OleTypeData, util::ToWide};

pub fn progids() -> Vec<String> {
    let mut hclsids = HKEY::default();
    let mut hclsid = HKEY::default();
    let mut progids = vec![];

    let clsid = "CLSID".to_wide_null();
    let clsid_pcwstr = PCWSTR::from_raw(clsid.as_ptr());
    let progid = "ProgID".to_wide_null();
    let progid_pcwstr = PCWSTR::from_raw(progid.as_ptr());
    let version_independent_progid = "VersionIndependentProgID".to_wide_null();
    let version_independent_progid_pcwstr = PCWSTR::from_raw(version_independent_progid.as_ptr());

    let mut result = reg_open_key(HKEY_CLASSES_ROOT, clsid_pcwstr, &mut hclsids);
    if result != ERROR_SUCCESS {
        return progids;
    }

    let mut i = 0;
    loop {
        let clsid = reg_enum_key(hclsids, i);

        if clsid.is_null() {
            break;
        }

        result = reg_open_key(hclsids, clsid, &mut hclsid);
        if result != ERROR_SUCCESS {
            println!("result is {:?}", result);
            i += 1;
            continue;
        }
        let v = reg_get_val2_string(hclsid, progid_pcwstr);
        if let Some(v) = v {
            progids.push(v);
        }
        let v = reg_get_val2_string(hclsid, version_independent_progid_pcwstr);
        if let Some(v) = v {
            progids.push(v);
        }
        unsafe { RegCloseKey(hclsid) };
        i += 1;
    }
    unsafe { RegCloseKey(hclsids) };
    progids
}

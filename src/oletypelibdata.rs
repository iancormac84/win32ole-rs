use crate::{
    error::Result,
    util::{reg_enum_key, reg_get_val, reg_get_val2, ToWide},
};
use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::ERROR_SUCCESS,
        System::{
            Com::ITypeLib,
            Environment::ExpandEnvironmentStringsW,
            Registry::{RegCloseKey, RegOpenKeyExW, HKEY, HKEY_CLASSES_ROOT, KEY_READ},
        },
    },
};

pub struct OleTypeLibData {
    typelib: ITypeLib,
}

fn typelib_file_from_typelib(ole: PCWSTR) -> Result<PCWSTR> {
    let mut htypelib = HKEY::default();
    let mut hclsid = HKEY::default();
    let mut hversion = HKEY::default();
    let mut hlang = HKEY::default();
    let mut file = PCWSTR::null();
    let mut i = 0;
    let mut j = 0;
    let mut k = 0;
    let mut found = false;
    let name = PCWSTR("TypeLib".to_wide_null().as_ptr());
    let result = unsafe { RegOpenKeyExW(HKEY_CLASSES_ROOT, name, 0, KEY_READ, &mut htypelib) };
    if result != ERROR_SUCCESS {
        return Err(windows::core::Error::from(result).into());
    }

    while !found {
        let clsid = reg_enum_key(htypelib, i);
        if clsid.is_null() {
            break;
        }
        let result = unsafe { RegOpenKeyExW(htypelib, clsid, 0, KEY_READ, &mut hclsid) };
        if result != ERROR_SUCCESS {
            continue;
        }
        let mut fver = 0f64;
        while !found {
            let ver = reg_enum_key(hclsid, j);
            if ver.is_null() {
                break;
            };
            let result = unsafe { RegOpenKeyExW(hclsid, ver, 0, KEY_READ, &mut hversion) };
            let verdbl = unsafe { ver.to_string().unwrap().parse().unwrap() };
            if result != ERROR_SUCCESS || fver > verdbl {
                continue;
            }
            fver = verdbl;
            let typelib = reg_get_val(hversion, None);
            if typelib.is_null() {
                continue;
            }
            if typelib == ole {
                while !found {
                    let lang = reg_enum_key(hversion, k);
                    if lang.is_null() {
                        break;
                    }
                    let result = unsafe { RegOpenKeyExW(hversion, lang, 0, KEY_READ, &mut hlang) };
                    if result == ERROR_SUCCESS {
                        file = reg_get_typelib_file_path(hlang);
                        if !file.is_null() {
                            found = true;
                        }
                        unsafe { RegCloseKey(hlang) };
                    }
                    k += 1;
                }
            }
            unsafe { RegCloseKey(hversion) };
            j += 1;
        }
        unsafe { RegCloseKey(hclsid) };
        i += 1;
    }
    unsafe { RegCloseKey(htypelib) };
    Ok(file)
}

fn reg_get_typelib_file_path(hkey: HKEY) -> PCWSTR {
    let win64_vec = "win64".to_wide_null();
    let win64 = PCWSTR::from_raw(win64_vec.as_ptr());
    let path = reg_get_val2(hkey, win64);
    if !path.is_null() {
        return path;
    }
    let win32_vec = "win32".to_wide_null();
    let win32 = PCWSTR::from_raw(win32_vec.as_ptr());
    let path = reg_get_val2(hkey, win32);
    if !path.is_null() {
        return path;
    }
    let win16_vec = "win16".to_wide_null();
    let win16 = PCWSTR::from_raw(win16_vec.as_ptr());
    reg_get_val2(hkey, win16)
}

fn typelib_file_from_clsid(ole: PCWSTR) -> Result<PCWSTR> {
    let clsid_vec = "CLSID".to_wide_null();
    let name = PCWSTR(clsid_vec.as_ptr());
    let mut hroot = HKEY::default();
    let result = unsafe { RegOpenKeyExW(HKEY_CLASSES_ROOT, name, 0, KEY_READ, &mut hroot) };
    if result != ERROR_SUCCESS {
        return Err(windows::core::Error::from(result).into());
    }
    let mut hclsid = HKEY::default();
    let result = unsafe { RegOpenKeyExW(hroot, ole, 0, KEY_READ, &mut hclsid) };
    if result != ERROR_SUCCESS {
        unsafe { RegCloseKey(hroot) };
        return Err(windows::core::Error::from(result).into());
    }
    let inproc_server32_vec = "InprocServer32".to_wide_null();
    let inproc_server32 = PCWSTR::from_raw(inproc_server32_vec.as_ptr());
    let mut typelib = reg_get_val2(hclsid, inproc_server32);
    unsafe {
        RegCloseKey(hroot);
        RegCloseKey(hclsid);
    }
    if !typelib.is_null() {
        let len = unsafe { ExpandEnvironmentStringsW(typelib, None) };
        let mut path = vec![0; len as usize];
        unsafe { ExpandEnvironmentStringsW(typelib, Some(&mut path)) };
        typelib.0 = path.as_ptr();
    }
    Ok(typelib)
}

pub(crate) fn typelib_file(ole: PCWSTR) -> Result<PCWSTR> {
    let file = typelib_file_from_clsid(ole)?;
    if !file.is_null() {
        return Ok(file);
    }
    typelib_file_from_typelib(ole)
}

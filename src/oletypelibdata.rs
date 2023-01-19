use std::{
    ffi::OsStr,
    iter::zip,
    path::PathBuf,
    ptr::{self, NonNull},
};

use crate::{
    error::{Error, OleError, Result},
    types::{OleClassNames, TypeInfos},
    util::{
        conv::{os_string_from_ptr, ToWide},
        RegKey,
    },
    OleTypeData,
};
use windows::{
    core::{BSTR, GUID, PCWSTR},
    Win32::{
        Foundation::E_UNEXPECTED,
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{ITypeInfo, ITypeLib, SYSKIND, TLIBATTR},
            Environment::ExpandEnvironmentStringsW,
            Ole::{
                LoadTypeLibEx, QueryPathOfRegTypeLib, LIBFLAG_FHIDDEN, LIBFLAG_FRESTRICTED,
                REGKIND_NONE,
            },
            Registry::HKEY_CLASSES_ROOT,
        },
    },
};

fn isdigit(c: char) -> bool {
    c.is_ascii_digit()
}

fn atof(s: &str) -> f64 {
    // This function stolen from either Rolf Neugebauer or Andrew Tolmach.
    // Probably Rolf.
    let mut a = 0.0;
    let mut e: i32 = 0;

    let mut cur_idx = 0;
    for (idx, c) in s.chars().enumerate() {
        cur_idx = idx;
        if isdigit(c) {
            a = a * 10.0 + (c as u32 - '0' as u32) as f64;
        } else {
            break;
        }
    }

    if &s[cur_idx..=cur_idx] == "." {
        cur_idx += 1;
        let n = cur_idx;
        for (idx, c) in s[n..].chars().enumerate() {
            cur_idx = idx;
            if isdigit(c) {
                a = a * 10.0 + (c as u32 - '0' as u32) as f64;
                e -= 1;
            } else {
                break;
            }
        }
    }
    if &s[cur_idx..=cur_idx] == "e" || &s[cur_idx..=cur_idx] == "E" {
        let mut sign: i8 = 1;
        let mut i = 0;
        cur_idx += 1;
        if &s[cur_idx..=cur_idx] == "+" {
            cur_idx += 1;
        } else if &s[cur_idx..=cur_idx] == "-" {
            cur_idx += 1;
            sign = -1;
        }
        let n = cur_idx;
        for c in s[n..].chars() {
            if isdigit(c) {
                i = i * 10 + (c as u32 - '0' as u32);
            }
        }

        e += i as i32 * sign as i32;
    }

    while e > 0 {
        a *= 10.0;
        e -= 1;
    }

    while e < 0 {
        a *= 0.1;
        e += 1;
    }
    a
}

pub struct OleTypeLibData {
    typelib: ITypeLib,
    name: String,
    tlib_attr: NonNull<TLIBATTR>,
}

impl OleTypeLibData {
    pub fn new1<S: AsRef<str>>(typelib_str: S) -> Result<OleTypeLibData> {
        let mut typelibdata = oletypelib_search_registry(&typelib_str);
        if typelibdata.is_err() {
            typelibdata = oletypelib_search_registry2([typelib_str.as_ref(), "", ""]);
        } else {
            return typelibdata;
        }
        if typelibdata.is_err() {
            let typelib_str = typelib_str.as_ref();
            let typelib_vec = typelib_str.to_wide_null();
            let typelib_pcwstr = PCWSTR::from_raw(typelib_vec.as_ptr());
            let typelib = unsafe { LoadTypeLibEx(typelib_pcwstr, REGKIND_NONE) };
            if let Ok(typelib) = typelib {
                let name = name_from_typelib(&typelib);
                let tlib_attr = unsafe { typelib.GetLibAttr() }?;
                let tlib_attr = NonNull::new(tlib_attr).unwrap();
                Ok(OleTypeLibData {
                    typelib,
                    name: name.unwrap_or(String::new()),
                    tlib_attr,
                })
            } else {
                Err(Error::Custom(format!(
                    "type library `{typelib_str}` not found",
                )))
            }
        } else {
            typelibdata
        }
    }
    pub fn new2<S: AsRef<str>>(typelib_str: S, version: f64) -> Result<OleTypeLibData> {
        let mut typelibdata = oletypelib_search_registry(&typelib_str);
        if typelibdata.is_err() {
            let version_str = version.to_string();
            typelibdata = oletypelib_search_registry2([typelib_str.as_ref(), &version_str, ""]);
        } else {
            return typelibdata;
        }
        if typelibdata.is_err() {
            let typelib_str = typelib_str.as_ref();
            let typelib_vec = typelib_str.to_wide_null();
            let typelib_pcwstr = PCWSTR::from_raw(typelib_vec.as_ptr());
            let typelib = unsafe { LoadTypeLibEx(typelib_pcwstr, REGKIND_NONE) };
            if let Ok(typelib) = typelib {
                let name = name_from_typelib(&typelib);
                let tlib_attr = unsafe { typelib.GetLibAttr() }?;
                let tlib_attr = NonNull::new(tlib_attr).unwrap();
                Ok(OleTypeLibData {
                    typelib,
                    name: name.unwrap_or(String::new()),
                    tlib_attr,
                })
            } else {
                Err(Error::Custom(format!(
                    "type library `{typelib_str}` not found",
                )))
            }
        } else {
            typelibdata
        }
    }
    pub fn new3<S: AsRef<str>>(typelib_str: S, major: S, minor: S) -> Result<OleTypeLibData> {
        let mut typelibdata = oletypelib_search_registry(&typelib_str);
        if typelibdata.is_err() {
            typelibdata =
                oletypelib_search_registry2([typelib_str.as_ref(), major.as_ref(), minor.as_ref()]);
        } else {
            return typelibdata;
        }
        if typelibdata.is_err() {
            let typelib_str = typelib_str.as_ref();
            let typelib_vec = typelib_str.to_wide_null();
            let typelib_pcwstr = PCWSTR::from_raw(typelib_vec.as_ptr());
            let typelib = unsafe { LoadTypeLibEx(typelib_pcwstr, REGKIND_NONE) };
            if let Ok(typelib) = typelib {
                let name = name_from_typelib(&typelib);
                let tlib_attr = unsafe { typelib.GetLibAttr() }?;
                let tlib_attr = NonNull::new(tlib_attr).unwrap();
                Ok(OleTypeLibData {
                    typelib,
                    name: name.unwrap_or(String::new()),
                    tlib_attr,
                })
            } else {
                Err(Error::Custom(format!(
                    "type library `{typelib_str}` not found",
                )))
            }
        } else {
            typelibdata
        }
    }
    pub fn make<S: AsRef<str>>(typelib: ITypeLib, name: S) -> Result<OleTypeLibData> {
        let tlib_attr = unsafe { typelib.GetLibAttr() }?;
        let tlib_attr = NonNull::new(tlib_attr).unwrap();
        Ok(OleTypeLibData {
            typelib,
            name: name.as_ref().to_string(),
            tlib_attr,
        })
    }
    pub fn guid(&self) -> GUID {
        unsafe { self.tlib_attr.as_ref().guid }
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn library_name(&self) -> Result<String> {
        library_name_from_typelib(&self.typelib)
    }
    pub fn version(&self) -> Result<f64> {
        let major = unsafe { self.tlib_attr.as_ref().wMajorVerNum };
        let minor = unsafe { self.tlib_attr.as_ref().wMinorVerNum };
        let version = format!("{major}.{minor}");
        match version.parse() {
            Ok(version) => Ok(version),
            Err(error) => Err(error.into()),
        }
    }
    pub fn major_version(&self) -> u16 {
        unsafe { self.tlib_attr.as_ref().wMajorVerNum }
    }
    pub fn minor_version(&self) -> u16 {
        unsafe { self.tlib_attr.as_ref().wMinorVerNum }
    }
    pub fn path(&self) -> Result<PathBuf> {
        let result = unsafe {
            QueryPathOfRegTypeLib(
                &self.tlib_attr.as_ref().guid,
                self.tlib_attr.as_ref().wMajorVerNum,
                self.tlib_attr.as_ref().wMinorVerNum,
                GetUserDefaultLCID(),
            )
        };
        if let Err(error) = result {
            return Err(OleError::runtime(error, "failed to QueryPathOfRegTypeTypeLib").into());
        }

        let bstr = result.unwrap();
        let path = unsafe { os_string_from_ptr(bstr) };
        Ok(path.into())
    }
    pub fn visible(&self) -> bool {
        let lib_flags = self.lib_flags();

        lib_flags == 0
            || lib_flags & LIBFLAG_FRESTRICTED.0 as u16 != 0
            || lib_flags & LIBFLAG_FHIDDEN.0 as u16 != 0
    }
    pub fn ole_types(&self) -> Vec<Result<OleTypeData>> {
        ole_types_from_typelib(&self.typelib)
    }
    pub fn lib_flags(&self) -> u16 {
        unsafe { self.tlib_attr.as_ref().wLibFlags }
    }
    pub fn lcid(&self) -> u32 {
        unsafe { self.tlib_attr.as_ref().lcid }
    }
    pub fn syskind(&self) -> SYSKIND {
        unsafe { self.tlib_attr.as_ref().syskind }
    }
    pub fn num_type_entries(&self) -> u32 {
        unsafe { self.typelib.GetTypeInfoCount() }
    }
}

impl TryFrom<&ITypeInfo> for OleTypeLibData {
    type Error = Error;

    fn try_from(typeinfo: &ITypeInfo) -> Result<OleTypeLibData> {
        let mut typelib = None;
        let mut index = 0;
        unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut index) }?;
        let typelib = typelib.unwrap();
        let name = library_name_from_typelib(&typelib)?;
        let tlib_attr = unsafe { typelib.GetLibAttr() }?;
        let tlib_attr = NonNull::new(tlib_attr).unwrap();
        Ok(OleTypeLibData {
            typelib,
            name,
            tlib_attr,
        })
    }
}

fn typelib_file_from_typelib<P: AsRef<OsStr>>(ole: P) -> Result<PathBuf> {
    let htypelib = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("TypeLib")?;
    let mut found = false;
    let mut file = None;

    for clsid_or_error in htypelib.enum_keys() {
        if found {
            break;
        }
        let clsid = clsid_or_error?;

        let hclsid = htypelib.open_subkey(clsid);
        if let Ok(hclsid) = hclsid {
            let mut fver = 0f64;
            for version_or_error in hclsid.enum_keys() {
                if found {
                    break;
                }
                let version = version_or_error?;
                let hversion = hclsid.open_subkey(&version);
                if hversion.is_err() || fver > atof(&version) {
                    continue;
                }
                let hversion = hversion?;
                fver = atof(&version);
                let typelib = hversion.get_value("");
                if typelib.is_err() {
                    continue;
                } else {
                    let typelib = typelib?;
                    let ole = ole.as_ref();
                    if typelib == ole.to_str().unwrap() {
                        for lang_or_error in hversion.enum_keys() {
                            if found {
                                break;
                            }
                            let lang = lang_or_error?;
                            let hlang = hversion.open_subkey(lang);
                            if let Ok(hlang) = hlang {
                                file = reg_get_typelib_file_path(hlang);
                                if let Some(ref file) = file {
                                    found = file.is_ok();
                                }
                            }
                        }
                    }
                }
            }
        } else {
            continue;
        }
    }
    file.unwrap()
}

fn reg_get_typelib_file_path(hkey: RegKey) -> Option<Result<PathBuf>> {
    let hwin64 = hkey.open_subkey("win64");
    if let Ok(hwin64) = hwin64 {
        let path = hwin64.get_value("");
        if let Ok(path) = path {
            return Some(Ok(PathBuf::from(path)));
        }
    }

    let hwin32 = hkey.open_subkey("win32");
    if let Ok(hwin32) = hwin32 {
        let path = hwin32.get_value("");
        if let Ok(path) = path {
            return Some(Ok(PathBuf::from(path)));
        }
    }

    let hwin16 = hkey.open_subkey("win16");
    if let Ok(hwin16) = hwin16 {
        let path = hwin16.get_value("");
        if let Ok(path) = path {
            return Some(Ok(PathBuf::from(path)));
        }
    }
    None
}

fn typelib_file_from_clsid<P: AsRef<OsStr>>(ole: P) -> Result<PathBuf> {
    let hroot = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("CLSID")?;

    let hclsid = hroot.open_subkey(ole)?;
    let htypelib = hclsid.open_subkey("InprocServer32");
    let typelib = if let Ok(htypelib) = htypelib {
        htypelib.get_value("")
    } else {
        hclsid.get_value("InprocServer32")
    };
    match typelib {
        Ok(typelib) => {
            let typelib_pcwstr = PCWSTR::from_raw(typelib.to_wide_null().as_ptr());
            let len = unsafe { ExpandEnvironmentStringsW(typelib_pcwstr, None) };
            let mut path = vec![0; len as usize + 1];
            unsafe { ExpandEnvironmentStringsW(typelib_pcwstr, Some(&mut path)) };
            let path = PathBuf::from(unsafe { typelib_pcwstr.to_string()? });
            Ok(path)
        }
        Err(error) => Err(error),
    }
}

pub(crate) fn typelib_file<P: AsRef<OsStr>>(ole: P) -> Result<PathBuf> {
    let file = typelib_file_from_clsid(&ole);
    match file {
        Ok(file) => Ok(file),
        Err(_) => typelib_file_from_typelib(&ole),
    }
}

pub fn oletypelib_path(guid: &str, version: &str) -> Option<Result<PathBuf>> {
    let key = format!(r"TypeLib\{guid}\{version}");
    let hkey = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey(key);
    if let Ok(hkey) = hkey {
        let mut iter = hkey.enum_keys();
        loop {
            match iter.next() {
                None => {
                    break None;
                }
                Some(lang_or_error) => {
                    if let Ok(lang) = lang_or_error {
                        let hlang = hkey.open_subkey(lang);
                        if let Ok(hlang) = hlang {
                            return reg_get_typelib_file_path(hlang);
                        }
                    }
                }
            }
        }
    } else {
        None
    }
}

pub fn oletypelib_from_guid(guid: &str, version: &str) -> Result<ITypeLib> {
    let path = oletypelib_path(guid, version);
    let Some(path) = path else {
        return Err(windows::core::Error::from(E_UNEXPECTED).into());
    };
    let path = path?;
    let result =
        unsafe { LoadTypeLibEx(PCWSTR::from_raw(path.to_wide_null().as_ptr()), REGKIND_NONE) };
    match result {
        Ok(typelib) => Ok(typelib),
        Err(error) => Err(error.into()),
    }
}

fn oletypelib_search_registry<S: AsRef<str>>(typelib_str: S) -> Result<OleTypeLibData> {
    let mut found = false;
    let mut maybe_oletypelibdata = None;
    let htypelib = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("TypeLib")?;

    for guid_or_error in htypelib.enum_keys() {
        if found {
            break;
        }
        let Ok(guid) = guid_or_error else {
            continue;
        };
        let hguid = htypelib.open_subkey(&guid);
        let Ok(hguid) = hguid else {
            continue;
        };
        for version_or_error in hguid.enum_keys() {
            if found {
                break;
            }
            let Ok(version) = version_or_error else {
                continue;
            };
            let hversion = hguid.open_subkey(&version);
            let Ok(hversion) = hversion else {
                continue;
            };
            let tlib = hversion.get_value("");
            let Ok(tlib) = tlib else {
                continue;
            };
            if typelib_str.as_ref() == tlib {
                let typelib = oletypelib_from_guid(&guid, &version);
                if let Ok(typelib) = typelib {
                    let name = name_from_typelib(&typelib);
                    let tlib_attr = unsafe { typelib.GetLibAttr() }?;
                    let tlib_attr = NonNull::new(tlib_attr).unwrap();
                    maybe_oletypelibdata = Some(OleTypeLibData {
                        typelib,
                        name: name.unwrap_or(String::new()),
                        tlib_attr,
                    });
                    found = true;
                }
            }
        }
    }
    if let Some(typelibdata) = maybe_oletypelibdata {
        Ok(typelibdata)
    } else {
        Err(Error::Custom(format!(
            "type library `{}` was not found",
            typelib_str.as_ref()
        )))
    }
}

fn oletypelib_search_registry2(args: [&str; 3]) -> Result<OleTypeLibData> {
    let mut maybe_oletypelibdata = None;
    let guid = args[0];
    let version_str = make_version_str(args[1], args[2]);

    let htypelib = RegKey::predef(HKEY_CLASSES_ROOT).open_subkey("TypeLib")?;

    let hguid = htypelib.open_subkey(guid)?;

    let mut typelib_str = String::new();
    let mut version = String::new();
    if let Some(ref version_str) = version_str {
        let hversion = hguid.open_subkey(version_str);
        if let Ok(hversion) = hversion {
            let tlib = hversion.get_value("");
            if let Ok(tlib) = tlib {
                typelib_str = tlib;
                version = version_str.to_string();
            }
        }
    } else {
        let mut fver = 0.0;
        for ver_or_error in hguid.enum_keys() {
            let Ok(ver) = ver_or_error else {
                break;
            };
            let hversion = hguid.open_subkey(&ver);
            let Ok(hversion) = hversion else {
                continue;
            };
            let tlib = hversion.get_value("");
            let Ok(tlib) = tlib else {
                continue;
            };

            if fver < atof(&ver) {
                fver = atof(&ver);
                version = ver;
                typelib_str = tlib;
            }
        }
    }
    if !typelib_str.is_empty() {
        let typelib = oletypelib_from_guid(guid, &version);
        if let Ok(typelib) = typelib {
            let name = name_from_typelib(&typelib);
            let tlib_attr = unsafe { typelib.GetLibAttr() }?;
            let tlib_attr = NonNull::new(tlib_attr).unwrap();
            maybe_oletypelibdata = Some(OleTypeLibData {
                typelib,
                name: name.unwrap_or(String::new()),
                tlib_attr,
            });
        }
    }
    if let Some(typelibdata) = maybe_oletypelibdata {
        Ok(typelibdata)
    } else {
        let ver_desc = if let Some(version_str) = version_str {
            format!("version {version_str}")
        } else {
            "".to_string()
        };
        Err(Error::Custom(format!(
            "type library `{typelib_str}` {ver_desc} was not found"
        )))
    }
}

fn make_version_str(major: &str, minor: &str) -> Option<String> {
    if major.is_empty() {
        return None;
    }
    let mut version_str = major.to_string();
    if !minor.is_empty() {
        version_str.push('.');
        version_str.push_str(minor);
    }
    Some(version_str)
}

fn name_from_typelib(typelib: &ITypeLib) -> Result<String> {
    let mut bstrname = BSTR::default();
    unsafe { typelib.GetDocumentation(-1, None, Some(&mut bstrname), ptr::null_mut(), None) }?;
    Ok(bstrname.to_string())
}

fn library_name_from_typelib(typelib: &ITypeLib) -> Result<String> {
    let mut bstrname = BSTR::default();
    unsafe { typelib.GetDocumentation(-1, Some(&mut bstrname), None, ptr::null_mut(), None) }?;
    Ok(bstrname.to_string())
}

fn ole_types_from_typelib(typelib: &ITypeLib) -> Vec<Result<OleTypeData>> {
    let ole_class_names = OleClassNames::from(typelib);
    let typeinfos = TypeInfos::from(typelib);
    let iter_pair = zip(ole_class_names, typeinfos);
    let mut classes = vec![];
    for (ole_class_name, typeinfo) in iter_pair {
        let Ok(ole_class_name) = ole_class_name else {
            continue;
        };
        let Ok(typeinfo) = typeinfo else {
            continue;
        };
        let oletype = OleTypeData::make(typeinfo, ole_class_name);
        classes.push(oletype);
    }
    classes
}

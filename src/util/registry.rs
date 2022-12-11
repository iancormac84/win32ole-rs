use crate::{error::Result, ToWide};
use std::{
    ffi::{OsStr, OsString},
    fmt,
    os::windows::prelude::OsStringExt,
    slice,
};
use windows::{
    core::{PCWSTR, PWSTR},
    Win32::{
        Foundation::{ERROR_BAD_FILE_TYPE, ERROR_INVALID_BLOCK, ERROR_SUCCESS, WIN32_ERROR},
        System::{
            Environment::ExpandEnvironmentStringsW,
            Registry::{
                RegCloseKey, RegEnumKeyExW, RegOpenKeyExW, RegQueryValueExW, HKEY,
                HKEY_CLASSES_ROOT, KEY_READ, REG_DWORD, REG_EXPAND_SZ, REG_MULTI_SZ, REG_QWORD,
                REG_SAM_FLAGS, REG_SZ, REG_VALUE_TYPE,
            },
        },
    },
};

/// A trait for types that can be loaded from registry values.
///
/// **NOTE:** Uses `from_utf16_lossy` when converting to `String`.
///
/// **NOTE:** When converting to `String` or `OsString`, trailing `NULL` characters are trimmed
/// and line separating `NULL` characters in `REG_MULTI_SZ` are replaced by `\n`
/// effectively representing the value as a multiline string.
/// When converting to `Vec<String>` or `Vec<OsString>` `NULL` is used as a strings separator.
pub trait FromRegValue: Sized {
    fn from_reg_value(val: &RegValue) -> Result<Self>;
}

impl FromRegValue for String {
    fn from_reg_value(val: &RegValue) -> Result<String> {
        match val.vtype {
            REG_SZ | REG_EXPAND_SZ | REG_MULTI_SZ => {
                let words = unsafe {
                    #[allow(clippy::cast_ptr_alignment)]
                    slice::from_raw_parts(val.bytes.as_ptr() as *const u16, val.bytes.len() / 2)
                };
                let mut s = if val.vtype == REG_EXPAND_SZ {
                    let words_pcwstr = PCWSTR::from_raw(words.as_ptr());
                    let len = unsafe { ExpandEnvironmentStringsW(words_pcwstr, None) };
                    let mut buf = vec![0; len as usize + 1];
                    unsafe { ExpandEnvironmentStringsW(words_pcwstr, Some(&mut buf)) };
                    String::from_utf16_lossy(&buf[..])
                } else {
                    String::from_utf16_lossy(words)
                };
                while s.ends_with('\u{0}') {
                    s.pop();
                }
                if val.vtype == REG_MULTI_SZ {
                    return Ok(s.replace('\u{0}', "\n"));
                }
                Ok(s)
            }
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

impl FromRegValue for Vec<String> {
    fn from_reg_value(val: &RegValue) -> Result<Vec<String>> {
        match val.vtype {
            REG_MULTI_SZ => {
                let words = unsafe {
                    slice::from_raw_parts(val.bytes.as_ptr() as *const u16, val.bytes.len() / 2)
                };
                let mut s = String::from_utf16_lossy(words);
                while s.ends_with('\u{0}') {
                    s.pop();
                }
                let v: Vec<String> = s.split('\u{0}').map(|x| x.to_owned()).collect();
                Ok(v)
            }
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

impl FromRegValue for OsString {
    fn from_reg_value(val: &RegValue) -> Result<OsString> {
        match val.vtype {
            REG_SZ | REG_EXPAND_SZ | REG_MULTI_SZ => {
                let mut words = unsafe {
                    #[allow(clippy::cast_ptr_alignment)]
                    slice::from_raw_parts(val.bytes.as_ptr() as *const u16, val.bytes.len() / 2)
                };
                while let Some(0) = words.last() {
                    words = &words[0..words.len() - 1];
                }
                let s = OsString::from_wide(words);
                Ok(s)
            }
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

impl FromRegValue for Vec<OsString> {
    fn from_reg_value(val: &RegValue) -> Result<Vec<OsString>> {
        match val.vtype {
            REG_MULTI_SZ => {
                let mut words = unsafe {
                    slice::from_raw_parts(val.bytes.as_ptr() as *const u16, val.bytes.len() / 2)
                };
                while let Some(0) = words.last() {
                    words = &words[0..words.len() - 1];
                }
                let v: Vec<OsString> = words
                    .split(|ch| *ch == 0u16)
                    .map(OsString::from_wide)
                    .collect();
                Ok(v)
            }
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

impl FromRegValue for u32 {
    fn from_reg_value(val: &RegValue) -> Result<u32> {
        match val.vtype {
            #[allow(clippy::cast_ptr_alignment)]
            REG_DWORD => Ok(unsafe { *(val.bytes.as_ptr() as *const u32) }),
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

impl FromRegValue for u64 {
    fn from_reg_value(val: &RegValue) -> Result<u64> {
        match val.vtype {
            #[allow(clippy::cast_ptr_alignment)]
            REG_QWORD => Ok(unsafe { *(val.bytes.as_ptr() as *const u64) }),
            _ => Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into()),
        }
    }
}

/// Raw registry value
#[derive(PartialEq)]
pub struct RegValue {
    pub bytes: Vec<u8>,
    pub vtype: REG_VALUE_TYPE,
}

macro_rules! format_reg_value {
    ($e:expr => $t:ident) => {
        match $t::from_reg_value($e) {
            Ok(val) => format!("{:?}", val),
            Err(_) => return Err(fmt::Error),
        }
    };
}

impl fmt::Display for RegValue {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let f_val = match self.vtype.0 {
            1 | 2 | 7 => format_reg_value!(self => String),
            4 => format_reg_value!(self => u32),
            11 => format_reg_value!(self => u64),
            _ => format!("{:?}", self.bytes), //TODO: implement more types
        };
        write!(f, "{f_val}")
    }
}

impl fmt::Debug for RegValue {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "RegValue({:?}: {})", self.vtype, self)
    }
}

unsafe impl Send for RegKey {}

/// Handle of opened registry key
#[derive(Debug)]
pub struct RegKey {
    hkey: HKEY,
}

impl RegKey {
    /// Open one of predefined keys:
    ///
    /// * `HKEY_CLASSES_ROOT`
    /// * `HKEY_CURRENT_USER`
    /// * `HKEY_LOCAL_MACHINE`
    /// * `HKEY_USERS`
    /// * `HKEY_PERFORMANCE_DATA`
    /// * `HKEY_PERFORMANCE_TEXT`
    /// * `HKEY_PERFORMANCE_NLSTEXT`
    /// * `HKEY_CURRENT_CONFIG`
    /// * `HKEY_DYN_DATA`
    /// * `HKEY_CURRENT_USER_LOCAL_SETTINGS`
    ///
    pub const fn predef(hkey: HKEY) -> RegKey {
        RegKey { hkey }
    }

    /// Open subkey with `KEY_READ` permissions.
    /// Will open another handle to itself if `path` is an empty string.
    /// To open with different permissions use `open_subkey_with_flags`.
    /// You can also use `create_subkey` to open with `KEY_ALL_ACCESS` permissions.
    ///
    pub fn open_subkey<P: AsRef<OsStr>>(&self, path: P) -> Result<RegKey> {
        self.open_subkey_with_flags(path, KEY_READ)
    }

    /// Open subkey with desired permissions.
    /// Will open another handle to itself if `path` is an empty string.
    ///
    pub fn open_subkey_with_flags<P: AsRef<OsStr>>(
        &self,
        path: P,
        perms: REG_SAM_FLAGS,
    ) -> Result<RegKey> {
        let c_path = path.to_wide_null();
        let mut new_hkey = HKEY::default();
        match unsafe { RegOpenKeyExW(self.hkey, PCWSTR(c_path.as_ptr()), 0, perms, &mut new_hkey) }
        {
            ERROR_SUCCESS => Ok(RegKey { hkey: new_hkey }),
            err => Err(windows::core::Error::from(err).into()),
        }
    }

    /// Return an iterator over subkeys names.
    ///
    pub const fn enum_keys(&self) -> EnumKeys {
        EnumKeys {
            key: self,
            index: 0,
        }
    }

    /// Get a value from registry and seamlessly convert it to the specified rust type
    /// with `FromRegValue` implemented (currently `String`, `u32` and `u64`).
    /// Will get the `Default` value if `name` is an empty string.
    ///
    pub fn get_value<N: AsRef<OsStr>>(&self, name: N) -> Result<String> {
        match self.get_raw_value(name) {
            Ok(ref val) => String::from_reg_value(val),
            Err(err) => Err(err),
        }
    }

    /// Get raw bytes from registry value.
    /// Will get the `Default` value if `name` is an empty string.
    ///
    pub fn get_raw_value<N: AsRef<OsStr>>(&self, name: N) -> Result<RegValue> {
        let c_name = name.to_wide_null();
        let mut buf_len = 2048;
        let mut buf_type = REG_VALUE_TYPE(0);
        let mut buf: Vec<u8> = Vec::with_capacity(buf_len as usize);
        loop {
            match unsafe {
                RegQueryValueExW(
                    self.hkey,
                    PCWSTR(c_name.as_ptr()),
                    None,
                    Some(&mut buf_type),
                    Some(buf.as_mut_ptr()),
                    Some(&mut buf_len),
                )
                .0
            } {
                0 => {
                    // ERROR_SUCCESS
                    unsafe {
                        buf.set_len(buf_len as usize);
                    }
                    // minimal check before transmute to RegType
                    if buf_type.0 > REG_QWORD.0 {
                        return Err(windows::core::Error::from(ERROR_BAD_FILE_TYPE).into());
                    }
                    return Ok(RegValue {
                        bytes: buf,
                        vtype: buf_type,
                    });
                }
                234 => {
                    // ERROR_MORE_DATA
                    buf.reserve(buf_len as usize);
                }
                err => return Err(windows::core::Error::from(WIN32_ERROR(err)).into()),
            }
        }
    }

    fn close_(&mut self) -> Result<()> {
        // don't try to close predefined keys
        if self.hkey.0 >= HKEY_CLASSES_ROOT.0 {
            return Ok(());
        };
        let result = unsafe { RegCloseKey(self.hkey) };
        match result.0 {
            0 => Ok(()),
            _ => Err(windows::core::Error::from(result).into()),
        }
    }

    fn enum_key(&self, index: u32) -> Option<Result<String>> {
        let mut name_len = 2048;
        #[allow(clippy::unnecessary_cast)]
        let mut name = [0; 2048];
        match unsafe {
            RegEnumKeyExW(
                self.hkey,
                index,
                PWSTR(name.as_mut_ptr()),
                &mut name_len,
                None,
                PWSTR::null(),
                None,
                None,
            )
            .0
        } {
            0 => match String::from_utf16(&name[..name_len as usize]) {
                // ERROR_SUCCESS
                Ok(s) => Some(Ok(s)),
                Err(_) => Some(Err(windows::core::Error::from(ERROR_INVALID_BLOCK).into())),
            },
            259 => None, // ERROR_NO_MORE_ITEMS
            err => Some(Err(windows::core::Error::from(WIN32_ERROR(err)).into())),
        }
    }
}

impl Drop for RegKey {
    fn drop(&mut self) {
        self.close_().unwrap_or(());
    }
}

/// Iterator over subkeys names
pub struct EnumKeys<'key> {
    key: &'key RegKey,
    index: u32,
}

impl<'key> Iterator for EnumKeys<'key> {
    type Item = Result<String>;

    fn next(&mut self) -> Option<Result<String>> {
        match self.key.enum_key(self.index) {
            v @ Some(_) => {
                self.index += 1;
                v
            }
            e @ None => e,
        }
    }

    fn nth(&mut self, n: usize) -> Option<Self::Item> {
        self.index += n as u32;
        self.next()
    }
}

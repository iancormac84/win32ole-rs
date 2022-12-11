use std::{ffi::OsStr, os::windows::prelude::OsStrExt};

pub trait ToWide {
    fn to_wide(&self) -> Vec<u16>;
    fn to_wide_null(&self) -> Vec<u16>;
}

impl<T> ToWide for T
where
    T: AsRef<OsStr>,
{
    fn to_wide(&self) -> Vec<u16> {
        self.as_ref().encode_wide().collect()
    }
    fn to_wide_null(&self) -> Vec<u16> {
        self.as_ref().encode_wide().chain(Some(0)).collect()
    }
}

/*pub fn to_u16s<S: AsRef<OsStr>>(s: S) -> Result<Vec<u16>> {
    let mut maybe_result: Vec<u16> = s.to_wide();
    if maybe_result.iter().any(|&u| u == 0) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "strings passed to WinAPI cannot contain NULs",
        )
        .into());
    }
    maybe_result.push(0);
    Ok(maybe_result)
}*/

use std::{fmt, io, string};

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
pub enum Error {
    Io(io::Error),
    Windows(windows::core::Error),
    Utf16(string::FromUtf16Error),
    Generic(&'static str),
    Custom(String),
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(self)
    }
}

impl From<io::Error> for Error {
    fn from(err: io::Error) -> Error {
        Error::Io(err)
    }
}

impl From<windows::core::Error> for Error {
    fn from(err: windows::core::Error) -> Error {
        Error::Windows(err)
    }
}

impl From<std::string::FromUtf16Error> for Error {
    fn from(err: std::string::FromUtf16Error) -> Error {
        Error::Utf16(err)
    }
}

impl fmt::Display for Error {
    fn fmt(&self, fmt: &mut fmt::Formatter) -> fmt::Result {
        use Error::*;
        match self {
            Io(ref err) => err.fmt(fmt),
            Windows(ref err) => err.fmt(fmt),
            Utf16(ref err) => err.fmt(fmt),
            Generic(ref err) => err.fmt(fmt),
            Custom(ref err) => err.fmt(fmt),
        }
    }
}

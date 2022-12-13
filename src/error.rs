use std::{
    ffi::IntoStringError,
    fmt, io,
    num::{ParseFloatError, TryFromIntError},
    str::Utf8Error,
    string::FromUtf16Error,
};

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
pub enum Error {
    Io(io::Error),
    Windows(windows::core::Error),
    Utf8(Utf8Error),
    Utf16(FromUtf16Error),
    ParseFloat(ParseFloatError),
    FromInt(TryFromIntError),
    FromVariant(FromVariantError),
    IntoString(IntoStringError),
    Generic(&'static str),
    Custom(String),
}

/// Encapsulates the ways converting from a `VARIANT` can fail.
#[derive(Copy, Clone, Debug)]
pub enum FromVariantError {
    /// `VARIANT` pointer during conversion was null
    VariantPtrNull,
    /// Unknown VT for
    UnknownVarType(u16),
}

impl fmt::Display for FromVariantError {
    fn fmt(&self, fmt: &mut fmt::Formatter) -> fmt::Result {
        match self {
            FromVariantError::VariantPtrNull => write!(fmt, "VARIANT pointer is null"),
            FromVariantError::UnknownVarType(e) => {
                write!(fmt, "VARIANT cannot be this vartype: {e}")
            }
        }
    }
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

impl From<Utf8Error> for Error {
    fn from(err: Utf8Error) -> Error {
        Error::Utf8(err)
    }
}

impl From<FromUtf16Error> for Error {
    fn from(err: FromUtf16Error) -> Error {
        Error::Utf16(err)
    }
}

impl From<ParseFloatError> for Error {
    fn from(err: ParseFloatError) -> Error {
        Error::ParseFloat(err)
    }
}

impl From<TryFromIntError> for Error {
    fn from(err: TryFromIntError) -> Error {
        Error::FromInt(err)
    }
}

impl From<FromVariantError> for Error {
    fn from(err: FromVariantError) -> Error {
        Error::FromVariant(err)
    }
}

impl From<IntoStringError> for Error {
    fn from(err: IntoStringError) -> Self {
        Error::IntoString(err)
    }
}

impl fmt::Display for Error {
    fn fmt(&self, fmt: &mut fmt::Formatter) -> fmt::Result {
        use Error::*;
        match self {
            Io(ref err) => err.fmt(fmt),
            Windows(ref err) => err.fmt(fmt),
            Utf8(ref err) => err.fmt(fmt),
            Utf16(ref err) => err.fmt(fmt),
            ParseFloat(ref err) => err.fmt(fmt),
            FromInt(ref err) => err.fmt(fmt),
            FromVariant(ref err) => err.fmt(fmt),
            IntoString(ref err) => err.fmt(fmt),
            Generic(ref err) => err.fmt(fmt),
            Custom(ref err) => err.fmt(fmt),
        }
    }
}

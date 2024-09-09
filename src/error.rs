use std::{
    ffi::IntoStringError,
    fmt, io,
    num::{ParseFloatError, TryFromIntError},
    str::Utf8Error,
    string::FromUtf16Error,
};

use windows::{
    core::HRESULT,
    Win32::{Foundation::WIN32_ERROR, System::Com::EXCEPINFO},
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
    IntoString(IntoStringError),
    Generic(&'static str),
    Custom(String),
    Ole(OleError),
    Exception(EXCEPINFO),
    IDispatchArgument {
        error_type: ComArgumentErrorType,
        arg_err: u32,
    },
}

#[derive(Debug)]
pub enum OleErrorType {
    Runtime,
    QueryInterface,
}

impl fmt::Display for OleErrorType {
    fn fmt(&self, fmt: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            OleErrorType::Runtime => write!(fmt, "Win32OleRuntimeError"),
            OleErrorType::QueryInterface => write!(fmt, "Win32OleQueryInterfaceError"),
        }
    }
}

#[derive(Debug)]
pub struct OleError {
    error_type: OleErrorType,
    hresult: HRESULT,
    context_message: String,
}

impl OleError {
    pub fn new<S: AsRef<str>, H: Into<HRESULT>>(
        error_type: OleErrorType,
        hresult: H,
        context_message: S,
    ) -> OleError {
        OleError {
            error_type,
            hresult: hresult.into(),
            context_message: context_message.as_ref().into(),
        }
    }
    pub fn runtime<S: AsRef<str>, H: Into<HRESULT>>(hresult: H, context_message: S) -> OleError {
        OleError::new(OleErrorType::Runtime, hresult, context_message)
    }
    pub fn interface<S: AsRef<str>, H: Into<HRESULT>>(hresult: H, context_message: S) -> OleError {
        OleError::new(OleErrorType::QueryInterface, hresult, context_message)
    }
}

impl fmt::Display for OleError {
    fn fmt(&self, fmt: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            fmt,
            "{}: {}. HRESULT error code: {} ({})",
            self.error_type,
            self.context_message,
            self.hresult,
            self.hresult.message()
        )
    }
}

impl From<OleError> for Error {
    fn from(err: OleError) -> Error {
        Error::Ole(err)
    }
}

#[derive(Debug)]
pub enum ComArgumentErrorType {
    TypeMismatch,
    ParameterNotFound,
}

impl fmt::Display for ComArgumentErrorType {
    fn fmt(&self, fmt: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self {
            ComArgumentErrorType::TypeMismatch => write!(
                fmt,
                "The value's type does not match the expected type for the parameter"
            ),
            ComArgumentErrorType::ParameterNotFound => {
                write!(fmt, "A required parameter was missing")
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

impl From<IntoStringError> for Error {
    fn from(err: IntoStringError) -> Self {
        Error::IntoString(err)
    }
}

impl From<WIN32_ERROR> for Error {
    fn from(err: WIN32_ERROR) -> Self {
        Error::Windows(HRESULT::from_win32(err.0).into())
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
            IntoString(ref err) => err.fmt(fmt),
            Generic(ref err) => err.fmt(fmt),
            Custom(ref err) => err.fmt(fmt),
            Ole(ref err) => err.fmt(fmt),
            Exception(excepinfo) => writeln!(fmt, "{}", ole_excepinfo2msg(excepinfo)),
            IDispatchArgument {
                error_type,
                arg_err,
            } => writeln!(
                fmt,
                "COM argument error {error_type} for argument {arg_err}"
            ),
        }
    }
}

fn ole_excepinfo2msg(excepinfo: &EXCEPINFO) -> String {
    let mut excepinfo = excepinfo.clone();
    if let Some(func) = excepinfo.pfnDeferredFillIn {
        let _ = unsafe { func(&mut excepinfo) };
    }

    let s = &excepinfo.bstrSource;
    let source = if !s.is_empty() {
        s.to_string()
    } else {
        String::new()
    };
    let d = &excepinfo.bstrDescription;
    let description = if !d.is_empty() {
        d.to_string()
    } else {
        String::new()
    };
    let mut msg = if excepinfo.wCode == 0 {
        format!("\n    OLE error code: {} in ", excepinfo.scode)
    } else {
        format!("\n    OLE error code: {} in ", excepinfo.wCode)
    };

    if !source.is_empty() {
        msg.push_str(&source);
    } else {
        msg.push_str("<Unknown>");
    }
    msg.push_str("\n      ");
    if !description.is_empty() {
        msg.push_str(&description);
    } else {
        msg.push_str("<No Description>");
    }

    let _ = excepinfo.bstrSource;
    let _ = excepinfo.bstrDescription;

    msg
}

use windows::{core::BSTR, Win32::System::Com::DISPPARAMS};

pub struct OleParam {
    dp: DISPPARAMS,
    named_args: Vec<BSTR>,
}

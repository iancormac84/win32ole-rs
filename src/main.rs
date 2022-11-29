use win32ole::{OleTypeData, OleMethodData};

fn main() {
    /*let typelib = "Microsoft Excel 15.0 Object Library".to_wide_null();
    let typelib_pcwstr = PCWSTR::from_raw(typelib.as_ptr());
    let oleclass = "Application".to_wide_null();
    let oleclass_pcwstr = PCWSTR::from_raw(oleclass.as_ptr());*/
    let excel = OleTypeData::from_prog_id("Excel.Application").unwrap();
    let method = OleMethodData::new(&excel, "Workbooks").unwrap().unwrap();
    let detail = method.return_type_detail().unwrap();
    println!("detail is {:?}", detail);
}

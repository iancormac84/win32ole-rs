use win32ole::OleTypeData;

fn main() {
    let excel = OleTypeData::from_typelib_and_oleclass(
        "Microsoft Excel 16.0 Object Library",
        "Application",
    )
    .unwrap();
    println!("excel.guid() is {:?}", excel.guid());
}

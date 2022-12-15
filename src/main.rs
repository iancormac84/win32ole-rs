fn main() {
    let ie = win32ole::OleTypeData::from_typelib_and_oleclass("Microsoft Internet Controls", "InternetExplorer").unwrap();
    let types = ie.default_ole_types().unwrap();
    for type_ in types {
        println!("{}", type_.name);
    }
}

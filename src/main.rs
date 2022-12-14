fn main() {
    let ie = win32ole::OleData::new("InternetExplorer.Application").unwrap();
    let ie_web_app = ie
        .ole_query_interface("{0002DF05-0000-0000-C000-000000000046}")
        .unwrap();
    println!("Here");
}

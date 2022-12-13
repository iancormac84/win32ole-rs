use win32ole::OleTypeLibData;

fn main() {
    let shell = OleTypeLibData::new1("C:\\Windows\\SYSTEM32\\USER32.DLL").unwrap();
    println!("{}", shell.name());
}

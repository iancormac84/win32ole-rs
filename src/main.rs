fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ole_type =
        win32ole::OleTypeData::new("{13709620-C279-11CE-A49E-444553540000}", "Shell").unwrap();
    let methods = ole_type.ole_methods().unwrap();
    for method in methods {
        println!("{}", method.name());
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ole_type = win32ole::OleTypeData::new("Microsoft Shell Controls And Automation", "Shell").unwrap();
    let methods = ole_type.ole_methods().unwrap();
    for method in methods {
        println!("{}", method.name());
    }
    Ok(())
}

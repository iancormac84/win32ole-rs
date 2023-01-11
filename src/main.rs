fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel = win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "Worksheet")?;
    let ole_types = excel.implemented_ole_types()?;
    let methods = excel.ole_methods()?;
    println!("methods.len() is {}", methods.len());
    for method in methods {
        println!("    {}", method.name());
    }
    for ole_type in ole_types {
        println!("{}", ole_type.name());
        let ole_type_methods = ole_type.ole_methods()?;
        println!("ole_type_methods.len() is {}", ole_type_methods.len());
        for ole_type_method in ole_type_methods {
            println!("    {}", ole_type_method.name());
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel = win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "Worksheet")?;
    let ole_types = excel.implemented_ole_types()?;
    let methods = excel.ole_methods()?;
    println!("methods.len() is {}", methods.len());
    for (idx, method) in methods.iter().enumerate() {
        println!("    {}) {}", idx + 1, method.name());
    }
    for (idx, ole_type) in ole_types.iter().enumerate() {
        println!("Implemented OLE type {}) {}", idx + 1, ole_type.name());
        let ole_type_methods = ole_type.ole_methods()?;
        for (idx1, ole_type_method) in ole_type_methods.iter().enumerate() {
            println!(
                "OLE type method {})    {}",
                idx1 + 1,
                ole_type_method.name()
            );
        }
    }
    Ok(())
}

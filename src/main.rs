fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel= win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "Workbook").unwrap();
    let ole_types = excel.implemented_ole_types()?;
    for ole_type in ole_types {
        println!("{}", ole_type.name());
        let methods = ole_type.ole_methods()?;
        for method in methods {
            println!("\t{}", method.name());
            let params = method.params();
            for param in params {
                let param = param?;
                println!("\t\t{}", param.name());
            }
        }
        let variables = ole_type.variables()?;
        for variable in variables {
            println!("\tVAR: {}", variable.name());
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel_worksheet =
        win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "Worksheet")?;
    let ole_types = excel_worksheet.implemented_ole_types()?;
    println!("ole_types.len() is {}", ole_types.len());
    let methods = excel_worksheet.ole_methods()?;
    println!("methods.len() is {}", methods.len());
    for (idx, method) in methods.iter().enumerate() {
        let method_name = method.name();
        println!("    OLE type method {}) {method_name}", idx + 1);
        let params = method.params();
        for param in params {
            match param {
                Ok(param) => println!("        {method_name} parameter: {}", param.name()),
                Err(error) => println!("        Error: {error}"),
            }
        }
    }
    for (idx, ole_type) in ole_types.iter().enumerate() {
        println!("Implemented OLE type {}) {}", idx + 1, ole_type.name());
        let ole_type_methods = ole_type.ole_methods()?;
        for (idx1, ole_type_method) in ole_type_methods.iter().enumerate() {
            let ole_type_method_name = ole_type_method.name();
            println!("OLE type method {})    {ole_type_method_name}", idx1 + 1);
            let params = ole_type_method.params();
            for param in params {
                match param {
                    Ok(param) => {
                        println!("        {ole_type_method_name} parameter: {}", param.name())
                    }
                    Err(error) => println!("        Error: {error}"),
                }
            }
        }
    }
    let excel_xl_sheet_type =
        win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "XlSheetType")?;
    let methods = excel_xl_sheet_type.ole_methods()?;
    println!("methods.len() is {}", methods.len());
    for (idx, method) in methods.iter().enumerate() {
        let method_name = method.name();
        println!("    OLE type method {}) {method_name}", idx + 1);
        let params = method.params();
        for param in params {
            match param {
                Ok(param) => println!("        {method_name} parameter: {}", param.name()),
                Err(error) => println!("        Error: {error}"),
            }
        }
    }
    let variables = excel_xl_sheet_type.variables();
    for (idx, variable) in variables.iter().enumerate() {
        match variable {
            Ok(variable) => println!("    OLE type variable {}) {}", idx + 1, variable.name()),
            Err(error) => println!("    OLE type variable error {}, {error}", idx + 1),
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel =
        win32ole::OleTypeData::new("Microsoft Excel 16.0 Object Library", "Workbook").unwrap();
    let methods = excel.ole_methods()?;
    println!("methods length is {}", methods.len());
    /*for ole_type in ole_types {
        println!("{}", ole_type.name());
        let methods = ole_type.ole_methods()?;
        println!("methods len is {}", methods.len());*/
        for method in methods {
            println!("\t{}", method.name());
            /*let params = method.params();
            println!("params len is {}", params.len());
            for param in params {
                let param = param?;
                println!("\t\t{}", param.name());
            }*/
        }
        /*let variables = ole_type.variables()?;
        println!("variables len is {}", variables.len());
        for variable in variables {
            println!("\tVAR: {}", variable.name());
        }
    }*/
    Ok(())
}

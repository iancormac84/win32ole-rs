fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel = win32ole::OleData::new("Excel.Application")?;
    let excel_tlib = excel.ole_typelib()?;
    let path = excel_tlib.path()?;
    println!("{}", path.display());
    let types = excel_tlib.ole_types();
    for type_ in types {
        match type_ {
            Ok(type_) => println!("{}", type_.name()),
            Err(error) => println!("Error was {error}"),
        }
    }

    Ok(())
}

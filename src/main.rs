fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel = win32ole::OleData::new("Excel.Application")?;
    let excel_tlib = excel.ole_typelib()?;
    let path = excel_tlib.path()?;
    println!("{}", path.display());

    Ok(())
}

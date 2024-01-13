fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel = win32ole::OleData::new("Excel.Application").unwrap();
    Ok(())
}

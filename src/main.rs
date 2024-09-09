use windows::core::VARIANT;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let excel_app = win32ole::OleData::new("Excel.Application")?;
    let mut vt_true = VARIANT::from(true);

    let vt = excel_app.get("Visible")?;
    println!("Visible: {}", vt);

    excel_app.put("Visible",&mut vt_true)?;

    let vt = excel_app.get("Visible")?;
    println!("Visible: {}", vt);

    std::thread::sleep(std::time::Duration::from_secs(25));
    Ok(())
}

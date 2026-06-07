fn main() -> Result<(), Box<dyn std::error::Error>> {
    let word =
        win32ole::OleData::new("Word.Application", None, None).unwrap();
    let methods = word.ole_methods().unwrap();
    for method in methods.iter() {
        println!("{}", method.name());
    }
    let visible_property = word.get("Visible").unwrap();
    println!("visible_property is {visible_property}");
    let vt = visible_property.vt();
    println!("vt is {vt:?}");
    //visible_property.Anonymous.Anonymous.
    Ok(())
}

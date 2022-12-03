use win32ole::OleTypeData;

fn main() {
    let excel = OleTypeData::from_prog_id("Excel.Application");
    match excel {
        Ok(excel) => {
            let methods = excel.methods().unwrap();
            println!("methods has {} structs", methods.len());
            for method in methods {
                println!("method name is {}", method.name());
            }
        }
        Err(error) => println!("{error}"),
    }
}

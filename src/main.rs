fn main() -> Result<(), Box<dyn std::error::Error>> {
    let wmplayer = win32ole::OleTypeLibData::new1("Windows Media Player").unwrap();
    let types = wmplayer.ole_types();
    for t in types {
        match t {
            Ok(t) => {
                println!("{}", t.name());
                let methods = t.ole_methods()?;
                for method in methods {
                    println!("  {}", method.name());
                }
                let variables = t.variables();
                for variable in variables {
                    let variable = variable?;
                    println!("      {}", variable.name());
                }
            }
            Err(error) => println!("Got error {error}"),
        }
    }
    Ok(())
}

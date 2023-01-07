use std::{
    fs::File,
    io::{BufWriter, Write},
};

fn main() {
    let excel = win32ole::OleTypeLibData::new1("Microsoft Excel 16.0 Object Library").unwrap();
    let ole_types = excel.ole_types();
    for ole_type in ole_types {
        println!("{}", ole_type.name());
    }
    /*let mut iterator_method_record =
        BufWriter::new(File::create(r"C:\workspace\win32ole\iterator_method_record.txt").unwrap());
    let mut ruby_method_record =
        BufWriter::new(File::create(r"C:\workspace\win32ole\ruby_method_record.txt").unwrap());
    let tlib = win32ole::OleTypeLibData::new1("Microsoft Excel 16.0 Object Library").unwrap();
    let mut iter_typeinfo = win32ole::types::TypeInfos::from(&tlib.typelib);
    let iter_typeattr = win32ole::types::TypeAttrs::from(&mut iter_typeinfo);
    for type_attr in iter_typeattr {
        let type_attr = type_attr.unwrap();
        writeln!(&mut iterator_method_record, "{}", type_attr.name()).unwrap();
        let iter_methods = win32ole::types::IterMethods::from(&type_attr);
        for method in iter_methods {
            let method = method.unwrap();
            writeln!(&mut iterator_method_record, "\t{}", method.name()).unwrap();
            let params = method.params();
            for param in params {
                writeln!(&mut iterator_method_record, "\t\t{}", param.name()).unwrap();
            }
        }
        let iter_vars = win32ole::types::IterVars::from(&type_attr);
        for var in iter_vars {
            let var = var.unwrap();
            writeln!(&mut iterator_method_record, "\tVAR: {}", var.name()).unwrap();
        }
    }
    let ole_types = tlib.ole_types().unwrap();
    for ole_type in ole_types {
        writeln!(&mut ruby_method_record, "{}", ole_type.name).unwrap();
        let methods = ole_type.ole_methods().unwrap();
        for method in methods {
            writeln!(&mut ruby_method_record, "\t{}", method.name()).unwrap();
            let params = method.params().unwrap();
            for param in params {
                writeln!(&mut ruby_method_record, "\t\t{}", param.name()).unwrap();
            }
        }
        let variables = ole_type.variables().unwrap();
        for variable in variables {
            writeln!(&mut ruby_method_record, "\tVAR: {}", variable.name()).unwrap();
        }
    }*/
}

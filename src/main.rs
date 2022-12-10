fn main() {
    let progids = win32ole::progids();
    println!("progids length is {}", progids.len());
    for progid in progids {
        println!("{progid}");
    }
}

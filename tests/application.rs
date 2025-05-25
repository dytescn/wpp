
#[cfg(test)]
mod tests {
    use wpssdk::ffi::info;
    #[test]
    fn it_works() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        let version = info::get_version();
        println!("---{:?}----",version);
        // app.
    }
    #[test]
    fn act_doc_name() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        let version = info::get_document_name();
        println!("---{:?}----",version);
        // app.
    }
}

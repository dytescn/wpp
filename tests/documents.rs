
#[cfg(test)]
mod tests {
    use wpp::ffi::document;
    #[test]
    fn it_works() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        let version = document::get_wps_docs_sum();
        println!("---{:?}----",version);
        // app.
    }

}

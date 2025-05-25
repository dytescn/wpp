
#[cfg(test)]
mod tests {
    use wpssdk::ffi::app;
    use wpssdk::ffi::info;
    #[test]
    fn it_works_doapp() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        // let version = info::get_version();
        // println!("---{:?}----",version);
        // app.
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\222.pptx";
        let _ = app::do_open(src_path);
        // let src = "c::"
    }
    #[test]
    fn it_works_layer_count() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        // let version = info::get_version();
        // println!("---{:?}----",version);
        // app.
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\222.pptx";
        let _ = info::get_count();
        // let src = "c::"
    }
    #[test]
    fn it_do_slide_export() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        // let version = info::get_version();
        // println!("---{:?}----",version);
        // app.
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\222.png";
        let _ = info::do_export(src_path,"png");
        // let src = "c::"
    }
    #[test]
    fn it_do_open_ppt() {
        // do_quit();
        // let version = get_version();
        // println!("{:?}",version);
        // let version = info::get_version();
        // println!("---{:?}----",version);
        // app.
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\123123.pptx";
        let _ = app::do_open(src_path);
        // let src = "c::"
    }
}


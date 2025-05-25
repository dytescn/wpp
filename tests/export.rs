#[cfg(test)]
mod tests {
    use wpssdk::ffi::export_cover::wpp_export_cover;
    use wpssdk::ffi::export_preview::wpp_export_preview_sum;
    use wpssdk::ffi::export_preview::wpp_export_preview;
    use wpssdk::ffi::export_file::export_file;
    #[test]
    fn do_export_cover() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\222.png";
        let _ = wpp_export_cover(src_path,"png");
    }
    
    #[test]
    fn do_export_preview_sum() {
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\222.png";
        let sum = wpp_export_preview_sum();
        println!("sum::{:?}",sum);
    }

    #[test]
    fn do_export_preview() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\";
        let sum = wpp_export_preview(src_path,"png");
        println!("sum::{:?}",sum);
    }
    
    #[test]
    fn do_export_fullname() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\123456.pptx";
        let sum = export_file(src_path);
        println!("sum::{:?}",sum);
    }
}
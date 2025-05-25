use windows::Win32::System::Com;
use crate::sdk::application::WppApplication;
// 保存封面图
pub fn export_file(file_src:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
        let app_res = WppApplication::new();
        if app_res.is_none() {
            return None;
        }
        // 保存封面
        let app =  app_res.unwrap();
        let act_doc = app.get_activepresentation().unwrap();
        let docs = app.get_presentations().unwrap();
        let doc_fullname = act_doc.get_fullname();
        println!("{:?}",doc_fullname);
        // 处理新创建的文件
        let path = std::path::Path::new(&doc_fullname);
        if path.extension().is_none() {
            act_doc.do_saveas(&file_src);
            return Some(());     
        }
        // 转移文件
        let res = std::fs::copy(doc_fullname, file_src);
        if res.is_err() {
            println!("{:?}",res.err());
            return None;
        }
        let _ = act_doc.do_close();
        let _ = docs.do_open(file_src);
        let _ = Com::CoUninitialize();
        return  Some(());
    }
}
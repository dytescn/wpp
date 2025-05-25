use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::WppApplication;

pub fn wpp_export_cover(file_path:&str,file_type:&str)->Option<()>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
    }
    let app_res = WppApplication::new();
    if app_res.is_none() {
        return None;
    }
    let app =  app_res.unwrap();
    let act_doc = app.get_activepresentation().unwrap();
    let layers = act_doc.get_layers().unwrap();
    let act_layer = layers.do_item(1).unwrap();
    let res = act_layer.do_export(file_path, file_type);
    if res.is_none() {
        return None;
    }
    unsafe{
        Com::CoUninitialize();
    } 
    return Some(());
}
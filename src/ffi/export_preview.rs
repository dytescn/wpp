use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::WppApplication;

pub fn wpp_export_preview(file_path:&str,file_type:&str)->Option<()>{
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
    let count = layers.get_count().unwrap();

    for item_index in 1..=count {
        let preview_src = format!("{}{}_preview.png",file_path,item_index);
        let act_layer = layers.do_item(item_index as i32).unwrap();
        let _ = act_layer.do_export(&preview_src, file_type);        
    }
    unsafe{
        Com::CoUninitialize();
    } 
    return Some(());
}

pub fn wpp_export_preview_sum()->Option<usize>{
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
    let layer_count = layers.get_count();
    unsafe{
        Com::CoUninitialize();
    } 
    return layer_count;
}
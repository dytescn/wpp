use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::WppApplication;
pub fn get_version()->Option<String>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
    }
    let app_res = WppApplication::new();
    if app_res.is_none() {
        return None;
    }
    let app =  app_res.unwrap();
    let ver = app.get_version();
    unsafe{
        Com::CoUninitialize();
    } 
    return Some(ver);
}


pub fn get_document_name()->Option<String>{
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
    let name =  act_doc.get_name();
    println!("{:?}",name);
    let docs = app.get_presentations().unwrap();
    let sum =  docs.get_count();
    println!("{:?}",sum);
    unsafe{
        Com::CoUninitialize();
    } 
    return Some("1".to_string());
}



pub fn get_count()->Option<usize>{
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
    let count = layers.get_count();
    println!("{:?}",count);

    return Some(0);
}


pub fn do_export(file_path:&str,file_type:&str)->Option<()>{
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
    let _ = act_layer.do_export(file_path, file_type);
    // let count = layers.get_count();
    // println!("{:?}",count);

    return Some(());
}
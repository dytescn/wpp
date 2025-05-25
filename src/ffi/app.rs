use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::core;
use crate::sdk::application::WppApplication;

pub fn do_open(file_src:&str)->Option<()>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
    }
    let app:WppApplication;
    let app_res = WppApplication::new();
    if app_res.is_none() {
        // return None;
        return None;
    }    
    app =  app_res.unwrap();
    let docs = app.get_presentations().unwrap();
    let res =docs.do_open(file_src);
    println!("{:?}",res);

    unsafe{
        Com::CoUninitialize();
    } 
    return Some(());
}


// 获取coreldraw
pub fn get_wps_version_info(ver:&str)->Option<u32>{
    unsafe{
        // 初始化系统
        let com_res = Com::CoInitialize(None);
        if com_res.is_err() {
            println!("com init error");
            return None;
        }
        let id_name = "kwpp.application.".to_string() + ver; // Can't use + with two &str
        let lpsz = core::HSTRING::from(id_name);
        let rclsid_res = Com::CLSIDFromProgID(&lpsz);
        if rclsid_res.is_err(){
            return None;
        }
        let ver_n = ver.parse::<u32>().unwrap();
        Com::CoUninitialize();
        return Some(ver_n);
    }
}

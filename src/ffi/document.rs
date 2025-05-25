use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::WppApplication;
// 获取coreldraw
pub fn get_wps_docs_sum()->Option<usize>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
    }
    let app_res = WppApplication::new();
    if app_res.is_none() {
        return None;
    }
    let app =  app_res.unwrap();
    let all_docs_res = app.get_presentations();
    if all_docs_res.is_none(){
        return None;
    }
    let all_docs = all_docs_res.unwrap();
    let layer_count = all_docs.get_count();
    unsafe{
        Com::CoUninitialize();
    } 
    return layer_count;
}

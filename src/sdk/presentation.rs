use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;

use super::slides::WppSlides;

pub struct WpsPresentation {
    disp:ComObject
}

impl WpsPresentation {
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    // 属性相关
    
    pub fn get_name(&self)->String{
        let app_vers = self.disp.get_property("name").expect("got err");
        app_vers.to_string().log_error("got error").unwrap()
    }
    
    pub fn set_name(){

    }

    pub fn get_fullname(&self)->String{
        let app_vers = self.disp.get_property("fullname").expect("got err");
        app_vers.to_string().log_error("got error").unwrap()
    }

    pub fn set_fullname(&self,path_src:&str)->Option<()>{
        let mut params = Vec::new();
        let vart_src = VARIANT::from_str(path_src);
        params.push(vart_src);
        let app_info = self.disp.set_property("fullname",params);
        if app_info.is_err() {
            println!("{:?}",app_info.err());
            return None;
        }
        return Some(());
    }

    pub fn get_layers(&self)->Option<WppSlides>{
        let active_doc_res = self.disp.get_property("slides");
        if active_doc_res.is_err() {
            return None;
        }
        let active_doc = active_doc_res.unwrap();
        let disp_res = active_doc.to_idispatch();
        if disp_res.is_err(){
            return None;
        }
        let disp = disp_res.unwrap();
        let ivg_doc = WppSlides::new(disp.clone());
        return Some(ivg_doc);
    }
    // 方法相关
    pub fn do_save(&self)->Option<()>{
        let mut params = Vec::new();
        let res_data = self.disp.invoke_method("save", params);
        match res_data {
            Ok(data)=>{
                return Some(());
            }
            Err(err)=>{
                return None;
            }
        }
    }
    // 方法相关
    pub fn do_close(&self)->Option<()>{
        let mut params = Vec::new();
        let res_data = self.disp.invoke_method("close", params);
        match res_data {
            Ok(data)=>{
                return Some(());
            }
            Err(err)=>{
                return None;
            }
        }
    }

    pub fn do_saveas(&self,path_src:&str)->Option<()>{
        let mut params = Vec::new();
        let vart_src = VARIANT::from_str(path_src);
        params.push(vart_src);
        let res_data = self.disp.invoke_method("saveas", params);
        match res_data {
            Ok(data)=>{
                return Some(());
            }
            Err(err)=>{
                return None;
            }
        }
    }
}
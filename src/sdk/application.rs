use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
use windows::Win32::System::Variant::VARIANT;
use crate::sdk::presentation::WpsPresentation;
use crate::sdk::presentations::WpsPresentations;

pub struct WppApplication {
    disp:ComObject
}

impl WppApplication {
    pub fn new()-> Option<Self> {
       let res_data=  ComObject::new_from_app("kwpp.Application");
       if res_data.is_none() {
           return None;
       }
       let data = res_data.unwrap();
       return Some(Self{
            disp:data
        });
    }
    pub fn new_name()-> Option<Self> {
        let res_data=  ComObject::new_from_name("kwpp.Application",GUID::zeroed());
        if res_data.is_none() {
            return None;
        }
        let data = res_data.unwrap();
        return Some(Self{
             disp:data
         });
     }
    // 属性相关
    pub fn get_visiable(){

    }
    pub fn set_visiable(&self,para:bool)->Option<()>{
        let mut params = Vec::new();
        let vart_src = VARIANT::from_bool(para);
        params.push(vart_src);
        let app_info = self.disp.set_property("visiable",params);
        if app_info.is_err() {
            return None;
        }
        return Some(());
    }
    
    pub fn get_Presentations(){

    }

    pub fn get_activepresentation(&self)->Option<WpsPresentation>{
        let active_doc_res = self.disp.get_property("ActivePresentation");
        if active_doc_res.is_err() {
            println!("res111:{:?}",active_doc_res.err());
            return None;
        }
        let active_doc = active_doc_res.unwrap();
        let disp_res = active_doc.to_idispatch();
        if disp_res.is_err(){
            println!("res222:{:?}",disp_res.err());
            return None;
        }
        let disp = disp_res.unwrap();
        let ivg_doc = WpsPresentation::new(disp.clone());
        return Some(ivg_doc);
    }

    pub fn get_presentations(&self)->Option<WpsPresentations>{
        let active_doc_res = self.disp.get_property("Presentations");
        if active_doc_res.is_err() {
            println!("res111:{:?}",active_doc_res.err());
            return None;
        }
        let active_doc = active_doc_res.unwrap();
        let disp_res = active_doc.to_idispatch();
        if disp_res.is_err(){
            println!("res222:{:?}",disp_res.err());
            return None;
        }
        let disp = disp_res.unwrap();
        let ivg_doc = WpsPresentations::new(disp.clone());
        return Some(ivg_doc);
    }
    // 方法相关
    pub fn get_version(&self)->String{
        let app_vers = self.disp.get_property("version").expect("got err");
        app_vers.to_string().log_error("got error").unwrap()
    }

    pub fn do_quit(&self)->Option<()>{
        
        return Some(());
    }
}
use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;

use super::slide::WppSlide;

pub struct WppSlides {
    disp:ComObject
}

impl WppSlides {
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    // 属性相关
    pub fn get_count(&self)->Option<usize>{
        let cunt_res = self.disp.get_property("count");
        if cunt_res.is_err(){
            return None;
        }
        let count =cunt_res.unwrap();
        let count_sum = count.to_u8().unwrap();
        return Some(count_sum as usize);
        // app_vers.to_string().log_error("got error").unwrap()
    }

    // 方法

    pub fn do_item(&self,index:i32)->Option<WppSlide>{
        let mut params = Vec::new();
        let range = VARIANT::from_i32(index);
        params.push(range);

        let res_data = self.disp.invoke_method("item", params);
        if res_data.is_err() {
            return None;
        }
        let data = res_data.unwrap();
        let disp_res = data.to_idispatch();
        if disp_res.is_err(){
            return None;
        }
        let disp = disp_res.unwrap();
        let ivg_doc = WppSlide::new(disp.clone());
        return Some(ivg_doc);
    }
    pub fn do_open(&self,path_src:&str)->Option<()>{
        let mut params = Vec::new();
        let vart_src = VARIANT::from_str(path_src);
        params.push(vart_src);
        let vart_read = VARIANT::from_i32(0);
        params.push(vart_read);
        // let vart_united = VARIANT::from_i32(0);
        // params.push(vart_united);
        // let vart_w = VARIANT::from_i32(1);
        // params.push(vart_w);



        let res_data = self.disp.invoke_method("open", params);
        if res_data.is_err() {
            println!("{:?}",res_data.err());
            return None;
        }
        return Some(());
    }
    pub fn do_add(){

    } 
}
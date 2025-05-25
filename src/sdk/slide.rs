use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;

pub struct WppSlide{
    disp:ComObject
}

impl WppSlide{
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
    pub fn do_export(&self,path_src:&str,file_type:&str)->Option<()>{
        let mut params = Vec::new();
        let vart_src = VARIANT::from_str(path_src);
        params.push(vart_src);
        let vart_read = VARIANT::from_str(file_type);
        params.push(vart_read);
        let res_data = self.disp.invoke_method("export", params);
        if res_data.is_err() {
            println!("{:?}",res_data.err());
            return None;
        }
        return Some(());
    }

    pub fn do_add(){

    } 
}
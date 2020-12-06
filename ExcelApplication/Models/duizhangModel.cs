using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Models
{
    public class duizhangModel
    {
        //送货日期 
        public DateTime songhuoriqi { get; set; }
        //送货单号  
        public string songhuodanhao { get; set; }
        //供应商 
        public string gongyinshang { get; set; }
        //物料类别 
        public string wuliaoleibie { get; set; }
        //物料名称 
        public string wuliaomingcheng { get; set; }
        //宽CM纸纹
        public int guigekuan { get; set; }
        //高CM 
        public int guigegao { get; set; }
        //数量 
        public int shuliang { get; set; }
        //单价 
        public double danjia { get; set; }
        //金额 
        public double jine { get; set; }
        //备注
        public string beizu { get; set; }

    }
}

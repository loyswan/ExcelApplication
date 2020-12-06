using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Models
{
    public class caigoujinduModel
    {
        public caigoujinduModel()
        {
            _rukuxinxiModels = new List<rukuxinxiModel>();
        }

        //供应商
        public string gongyinshang { get; set; }
        //开单日期   
        public DateTime kaidanriqi { get; set; }
        //采购单号 
        public string caigoudanhao { get; set; }
        //物料类别 
        public string wuliaoleibie { get; set; }
        //克重 
        public int kezhong { get; set; }
        //物料名称
        public string wuliaomingcheng { get; set; }
        //规格尺寸 宽
        public int guigekuan { get; set; }
        //规格尺寸 高
        public int guigegao { get; set; }
        // 订单需求张数 
        public int xuqiuzhangshu { get; set; }
        //采购数/重量(kg) 
        public int caigoushu { get; set; }
        //单价
        public double danjia { get; set; }
        //金额  
        public double jine { get; set; }
        //订单交期
        public DateTime dingdanjiaoqi { get; set; }
        //工单号 
        public string gongdanhao { get; set; }
        //客户 
        public string kehu { get; set; }
        //订单数量    
        public int dingdanshu { get; set; }
        //备注 
        public string beizu { get; set; }
        //来料交期    
        public DateTime lailiaojiaoqi { get; set; }
        //准确交期 
        public DateTime zhunquejiaoqi { get; set; }
        //提出需求日期 
        public DateTime xuqiuriqi { get; set; }
        //供应商回复日期 
        public DateTime gongyinshangriqi { get; set; }
        //供应商备注   
        public string gongyinshangbeizhu { get; set; }
        //欠到数/重量 
        public int qiandaoshu { get; set; }


        //List<rukuxinxiModel>  //入库信息列表
        private List<rukuxinxiModel> _rukuxinxiModels;

        public void AddrukuxinxiModel(rukuxinxiModel model)
        {
            _rukuxinxiModels.Add(model);
        }

        public List<rukuxinxiModel> GetrukuxinxiModels => _rukuxinxiModels;
    }

    public class rukuxinxiModel //入库信息
    {
        public rukuxinxiModel(int rukuxuhao, DateTime songhuoriqi,
            string songhuodanhao, int songhuoshu)
        {
            this.rukuxuhao = rukuxuhao;
            this.songhuoriqi = songhuoriqi;
            this.songhuodanhao = songhuodanhao;
            this.songhuoshu = songhuoshu;
        }

        //入库序号        
        public int rukuxuhao { get; set; }
        //送货日期 
        public DateTime songhuoriqi { get; set; }
        //送货单号 
        public string songhuodanhao { get; set; }
        //送货数/重量 
        public int songhuoshu { get; set; }
    }
}

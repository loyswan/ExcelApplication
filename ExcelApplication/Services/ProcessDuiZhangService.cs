using ExcelApplication.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Services
{
    public class ProcessDuiZhangService
    {
        public ProcessDuiZhangService()
        {
            lists = new List<duizhangModel>();
        }

        public ProcessDuiZhangService(List<caigoujinduModel> caigoujindulist) : this()
        {
            this.Caigoujindulist = caigoujindulist;
        }

        public List<caigoujinduModel> Caigoujindulist { get; set; }

        private List<duizhangModel> lists;

        public List<duizhangModel> GetList()
        {
            if (this.Caigoujindulist == null || this.Caigoujindulist.Count == 0)
            {
                return new List<duizhangModel>();
            }
            lists.Clear();
            this.Caigoujindulist.ForEach(c =>
            {
                c.GetrukuxinxiModels.ForEach(ruku =>
                {
                    duizhangModel model = new duizhangModel();
                    model.songhuoriqi = ruku.songhuoriqi;
                    model.songhuodanhao = ruku.songhuodanhao;
                    model.gongyinshang = c.gongyinshang;
                    model.wuliaoleibie = c.wuliaoleibie;
                    model.wuliaomingcheng = c.wuliaomingcheng;
                    model.guigekuan = c.guigekuan;
                    model.guigegao = c.guigegao;
                    model.shuliang = ruku.songhuoshu;
                    model.danjia = c.danjia;
                    model.jine = c.guigegao == 0
                        ? model.shuliang * model.danjia
                        : model.guigekuan * model.guigegao * 0.0001 * model.shuliang * model.danjia;
                    model.beizu = c.gongyinshangbeizhu;

                    this.lists.Add(model);
                });
            });

            return lists;
        }


        public void WriteFile(string filepath, List<duizhangModel> list = null)
        {
            list = list ?? this.lists;
            if (list == null) return;
            WriteExcelFileService service = new WriteExcelFileService(filepath);
            service.WriteDuizhang(list, "明细对账单");

        }
    }
}

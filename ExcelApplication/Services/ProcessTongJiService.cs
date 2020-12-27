using ExcelApplication.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Services
{
    public class ProcessTongJiService
    {
        public ProcessTongJiService()
        {
            lists = new List<tongjiModel>();
        }

        public ProcessTongJiService(List<duizhangModel> duizhanglist) : this()
        {
            this.Duizhanglist = duizhanglist;
        }

        public List<duizhangModel> Duizhanglist { get; set; }

        private List<tongjiModel> lists;

        /*
        public List<tongjiModel> GetList()
        {
            if (this.Duizhanglist == null || this.Duizhanglist.Count == 0)
            {
                return new List<tongjiModel>();
            }
            lists.Clear();

            List<tongjiModel> tongjilist = Duizhanglist
                .GroupBy(m => new { m.gongyinshang,m.wuliaoleibie,m.wuliaomingcheng })
                .Select(tongji =>
                {
                    tongjiModel model = new tongjiModel();
                    model.gongyinshang = tongji.Key.gongyinshang;
                    model.wuliaoleibie = tongji.Key.wuliaoleibie;
                    model.wuliaoming = tongji.Key.wuliaomingcheng;
                    model.pairs = tongji.GroupBy(n => n.songhuoriqi.ToString("yyyy-MM"))
                        .Select<IGrouping<string, duizhangModel>, MonthData>(
                            md => new MonthData(md.Key, md.Sum(x => x.shuliang), md.Sum(x => x.jine))
                        ).ToList();
                    return model;
                }).ToList();

            return tongjilist;
        }*/

        public List<tongjiModel> GetList()
        {
            if (this.Duizhanglist == null || this.Duizhanglist.Count == 0)
            {
                return new List<tongjiModel>();
            }
            lists.Clear();

            List<tongjiModel> tongjilist = Duizhanglist
                .GroupBy(m =>  m.gongyinshang )
                .Select(tongji =>
                {
                    tongjiModel model = new tongjiModel();
                    model.gongyinshang = tongji.Key;
                    model.pairs = tongji.GroupBy(n => n.songhuoriqi.ToString("yyyy-MM"))
                        .Select<IGrouping<string, duizhangModel>, MonthData>(
                            md => new MonthData(md.Key, md.Sum(x => x.shuliang), md.Sum(x => x.jine))
                        ).ToList();
                    return model;
                }).ToList();

            return tongjilist;
        }

        public void WriteFile(string filepath, List<tongjiModel> list = null)
        {
            list = list ?? this.lists;
            if (list == null) return;
            WriteExcelFileService service = new WriteExcelFileService(filepath);
            service.WriteTongJi(list, "供应商账单月统计");

        }
    }
}

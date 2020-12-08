using ExcelApplication.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Services
{
    public class SearchDuiZhangService
    {
        public SearchDuiZhangService()
        {
            searchlists = new List<duizhangModel>();
        }

        public SearchDuiZhangService(List<duizhangModel> duizhanglist) : this()
        {
            this.Duizhanglist = duizhanglist;
        }

        public List<duizhangModel> Duizhanglist { get; set; }

        private List<duizhangModel> searchlists;

        public List<duizhangModel> Search(DateTime startdate, DateTime enddate, string gongyinshangname = "All")
        {
            //分类统计
            return this.searchlists = this.Duizhanglist.Where(info => info.songhuoriqi >= startdate && info.songhuoriqi <= enddate)
                .OrderBy(d => d.gongyinshang).ToList();


            //if (string.IsNullOrEmpty(gongyinshangname) || gongyinshangname == "All")
            //{
            //    return dateinfos.ToList();
            //}
            //else
            //{
            //    return dateinfos.Where(x => x.gongyinshang == gongyinshangname).ToList();
            //}
        }

        public void WriteFile(string filepath, List<duizhangModel> list = null)
        {
            list = list ?? this.searchlists;
            if (list == null) return;
            WriteExcelFileService service = new WriteExcelFileService(filepath);
            service.WriteDuizhang(list, "输出对账单");


        }

       
    }
}

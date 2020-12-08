using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Models
{
    public class tongjiModel
    {
        //供应商
        public string gongyinshang { get; set; }
        //物料类别
        public string wuliaoleibie { get; set; }
        //物料名称
        public string wuliaoming { get; set; }

        ////01月货款	02月货款	03月货款	04月货款	05月货款	06月货款	07月货款	08月货款	09月货款	10月货款	11月货款	12月货款
        //月份 字符串 ”yyyy-MM“
        //金额 月度求和金额
        public List<MonthData> pairs = new List<MonthData>();

        //合计
        //public double hejijine => pairs.Sum(x => x.Value);

        //public void Add(string month, double jine) {
        //    if (pairs.Any(p=>p.Key==month))
        //    {
        //        pairs.Remove(pairs.First(p => p.Key == month));
        //    }
        //    pairs.Add(new KeyValuePair<string, double>(month, jine));
        //}

        public List<string> GetMonthString()
        {
            var month = pairs.GroupBy(p => p.MonthString).Select(g => g.First().MonthString).ToList();
            ////List<string> month = new List<string>();

            //if (pairs.Count > 0)
            //{
            //    foreach (var p in pairs)
            //    {
            //        if (!month.Contains(p.Key))
            //            month.Add(p.Key);
            //    }
            //}
            return month;
        }

    }

    public struct MonthData {
        public MonthData(string monthString, double number, double money)
        {
            MonthString = monthString ?? throw new ArgumentNullException(nameof(monthString));
            Number = number;
            Money = money;
        }

        public string MonthString { get; set; }
        public double Number { get; set; }
        public double Money { get; set; }

    }

}

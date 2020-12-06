using ExcelApplication.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApplication.Services
{
    public class ReadExcelFileService
    {
        public string FilePath { get; set; }
        public FileInfo CurrentFileInfo => new FileInfo(this.FilePath);

        public ReadExcelFileService(string filepath)
        {
            this.FilePath = filepath;
        }


        /// <summary>
        /// 返回读取的采购进度信息
        /// </summary>
        /// <param name="worksheetname"></param>
        /// <returns></returns>
        public List<caigoujinduModel> GetList(string worksheetname)
        {
            if (!this.CurrentFileInfo.Exists)
            {
                Console.WriteLine($"Error: {CurrentFileInfo.FullName}文件不存在。");
                return null;
            }

            var lists = new List<caigoujinduModel>();

            using (ExcelPackage package = new ExcelPackage(CurrentFileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == worksheetname)
                    ?? package.Workbook.Worksheets[0];

                //取得表格第一列最后一有效行 行号
                var endrow = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, 1]
                    .Last(c => c.Value != null).End.Row;

                //表头2行 格式行1行 低于3行视为无数据，第四行起为有效数据行
                if (endrow <= 3) return lists;

                //逐行读取数据
                Enumerable.Range(4, endrow - 4 + 1).ToList().ForEach(row =>
                    {
                        var model = new caigoujinduModel();
                        model.gongyinshang = worksheet.Cells[row, 1].Text;
                        model.kaidanriqi = worksheet.Cells[row, 2].GetValue<DateTime>();
                        model.caigoudanhao = worksheet.Cells[row, 3].Text;
                        model.wuliaoleibie = worksheet.Cells[row, 4].Text;
                        model.kezhong = worksheet.Cells[row, 5].GetValue<int>();
                        model.wuliaomingcheng = worksheet.Cells[row, 6].Text;
                        model.guigekuan = worksheet.Cells[row, 7].GetValue<int>();
                        model.guigegao = worksheet.Cells[row, 8].GetValue<int>();
                        model.xuqiuzhangshu = worksheet.Cells[row, 9].GetValue<int>();
                        model.caigoushu = worksheet.Cells[row, 10].GetValue<int>();
                        model.danjia = worksheet.Cells[row, 11].GetValue<double>();
                        model.jine = worksheet.Cells[row, 12].GetValue<double>();
                        model.dingdanjiaoqi = worksheet.Cells[row, 13].GetValue<DateTime>();
                        model.gongdanhao = worksheet.Cells[row, 14].Text;
                        model.kehu = worksheet.Cells[row, 15].Text;

                        model.dingdanshu = worksheet.Cells[row, 16].GetValue<int>();
                        model.beizu = worksheet.Cells[row, 17].Text;
                        model.lailiaojiaoqi = worksheet.Cells[row, 18].GetValue<DateTime>();
                        model.zhunquejiaoqi = worksheet.Cells[row, 19].GetValue<DateTime>();
                        model.xuqiuriqi = worksheet.Cells[row, 20].GetValue<DateTime>();
                        model.gongyinshangriqi = worksheet.Cells[row, 21].GetValue<DateTime>();
                        model.gongyinshangbeizhu = worksheet.Cells[row, 22].Text;
                        model.qiandaoshu = worksheet.Cells[row, 23].GetValue<int>();
                        //取得当前行的最后一列有效列号
                        var endcol = worksheet.Cells[row, 23, row, worksheet.Dimension.End.Column]
                                .First(cell => cell.Offset(0, 1).Value == null).End.Column;
                        if (endcol > 24)
                        {
                            //添加入库信息
                            for (int col = 24; col < endcol; col += 3)
                            {
                                model.AddrukuxinxiModel(new rukuxinxiModel(
                                    rukuxuhao: model.GetrukuxinxiModels.Count + 1,
                                    songhuoriqi: worksheet.Cells[row, col].GetValue<DateTime>(),
                                    songhuodanhao: worksheet.Cells[row, col + 1].Text,
                                    songhuoshu: worksheet.Cells[row, col + 2].GetValue<int>()
                                ));
                            }
                        }
                        //添加至列表
                        lists.Add(model);
                    });
            }
            return lists;
        }
    }
}

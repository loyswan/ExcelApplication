using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelApplication.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelApplication.Services
{
    public class WriteExcelFileService
    {
        public string FilePath { get; set; }
        public FileInfo CurrentFileInfo => new FileInfo(this.FilePath);

        public WriteExcelFileService(string filepath)
        {
            this.FilePath = filepath;
        }

        /// <summary>
        /// 写入对账明细  刷新
        /// </summary>
        /// <param name="worksheetname"></param>
        /// <returns></returns>
        public string WriteDuizhang(List<duizhangModel> list, string worksheetname)
        {
            if (!this.CurrentFileInfo.Exists)
            {
                Console.WriteLine($"Error: {CurrentFileInfo.FullName}文件不存在。");
                return $"Error: {CurrentFileInfo.FullName}文件不存在。";
            }

            using (ExcelPackage package = new ExcelPackage(CurrentFileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == worksheetname)
                    ?? package.Workbook.Worksheets.Add(worksheetname);

                worksheet.Cells.Clear();

                //填写标题行 字体加粗
                string[] tabletitlestring = { "送货日期", "送货单号", "供应商", "物料类别", "物料名称", "宽CM纸纹", "高CM", "数量", "单价", "金额", "备注" };
                for (int col = 0; col < tabletitlestring.Length; col++)
                {
                    var rng = worksheet.Cells[1, col + 1];
                    rng.Value = tabletitlestring[col];
                    rng.Style.Font.Bold = true;
                }
                //填写内容
                if (list.Count == 0) //没有内容直接保存退出
                {
                    package.Save();
                    return $"Error: 没有符合条件的数据！";
                }
                int row = 2;
                list.ForEach(duizhang =>
                {
                    worksheet.Cells[row, 1].Value = duizhang.songhuoriqi;
                    worksheet.Cells[row, 2].Value = duizhang.songhuodanhao;
                    worksheet.Cells[row, 3].Value = duizhang.gongyinshang;
                    worksheet.Cells[row, 4].Value = duizhang.wuliaoleibie;
                    worksheet.Cells[row, 5].Value = duizhang.wuliaomingcheng;
                    worksheet.Cells[row, 6].Value = duizhang.guigekuan;
                    worksheet.Cells[row, 7].Value = duizhang.guigegao;
                    worksheet.Cells[row, 8].Value = duizhang.shuliang;
                    worksheet.Cells[row, 9].Value = duizhang.danjia;
                    //worksheet.Cells[row, 10].Value = duizhang.jine;
                    worksheet.Cells[row, 11].Value = duizhang.beizu;
                    row++;
                });
                //填写公式
                worksheet.Cells[2, 10, list.Count + 1, 10].Formula = "=IF(G2=0,H2*I2,F2*G2*0.0001*H2*I2)";
                //单元格格式
                worksheet.Cells[2, 1, list.Count + 1, 1].Style.Numberformat.Format = "yyyy/m/d";
                worksheet.Cells[2, 9, list.Count + 1, 10].Style.Numberformat.Format = "0.00";
                //绘制边框
                using (var r = worksheet.Cells[1, 1, list.Count + 1, 11])
                {
                    r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    r.Style.Border.Top.Color.SetColor(Color.Black);
                    r.Style.Border.Bottom.Color.SetColor(Color.Black);
                    r.Style.Border.Left.Color.SetColor(Color.Black);
                    r.Style.Border.Right.Color.SetColor(Color.Black);

                    //自适应列宽
                    r.AutoFitColumns();
                    //居中
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                package.Save();
                return package.File.FullName;

            }
        }

        public string WriteJindu(List<caigoujinduModel> list, string worksheetname)
        {
            if (!this.CurrentFileInfo.Exists)
            {
                Console.WriteLine($"Error: {CurrentFileInfo.FullName}文件不存在。");
                return $"Error: {CurrentFileInfo.FullName}文件不存在。";
            }

            using (ExcelPackage package = new ExcelPackage(CurrentFileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == worksheetname)
                    ?? package.Workbook.Worksheets.Add(worksheetname);

                //worksheet.Cells.Clear();
                var endrow = worksheet.Dimension.End.Row;
                if (endrow >= 4)
                {
                    ExcelRange rng = worksheet.Cells[4, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                    rng.Clear();
                }

                //填写内容  第四行开始填写（1、2行标题，3行格式行）
                int row = 4;
                list.ForEach(c =>
                {
                    worksheet.Cells[row, 1].Value = c.gongyinshang;
                    worksheet.Cells[row, 2].Value = c.kaidanriqi;
                    worksheet.Cells[row, 3].Value = c.caigoudanhao;
                    worksheet.Cells[row, 4].Value = c.wuliaoleibie;
                    worksheet.Cells[row, 5].Value = c.kezhong;
                    worksheet.Cells[row, 6].Value = c.wuliaomingcheng;
                    worksheet.Cells[row, 7].Value = c.guigekuan;
                    worksheet.Cells[row, 8].Value = c.guigegao;
                    worksheet.Cells[row, 9].Value = c.xuqiuzhangshu;
                    worksheet.Cells[row, 10].Value = c.caigoushu;
                    worksheet.Cells[row, 11].Value = c.danjia;
                    worksheet.Cells[row, 12].Value = c.jine;
                    worksheet.Cells[row, 13].Value = c.dingdanjiaoqi;
                    worksheet.Cells[row, 14].Value = c.gongdanhao;
                    worksheet.Cells[row, 15].Value = c.kehu;
                    worksheet.Cells[row, 16].Value = c.dingdanshu;
                    worksheet.Cells[row, 17].Value = c.beizu;
                    worksheet.Cells[row, 18].Value = c.lailiaojiaoqi;
                    worksheet.Cells[row, 19].Value = c.zhunquejiaoqi;
                    worksheet.Cells[row, 20].Value = c.xuqiuriqi;
                    worksheet.Cells[row, 21].Value = c.gongyinshangriqi;
                    worksheet.Cells[row, 22].Value = c.gongyinshangbeizhu;
                    worksheet.Cells[row, 23].Value = c.qiandaoshu;
                    //循环填写入库信息
                    var ruku = c.GetrukuxinxiModels;
                    ruku.ForEach(r =>
                    {
                        worksheet.Cells[row, 24 - 3 + (r.rukuxuhao) * 3].Value = r.songhuoriqi;
                        worksheet.Cells[row, 24 - 2 + (r.rukuxuhao) * 3].Value = r.songhuodanhao;
                        worksheet.Cells[row, 24 - 1 + (r.rukuxuhao) * 3].Value = r.songhuoshu;
                    });
                    row++;
                });

                //格式刷 每列格式进行复制
                Enumerable.Range(1, 41).ToList().ForEach(col =>
                {
                    worksheet.Cells[3, col, worksheet.Dimension.End.Row, col].StyleID = worksheet.Cells[3, col].StyleID;
                });

                //保存工作表
                package.Save();
                return package.File.FullName;

            }
        }

        public string WriteTongJi(List<tongjiModel> list, string worksheetname)
        {
            if (!this.CurrentFileInfo.Exists)
            {
                Console.WriteLine($"Error: {CurrentFileInfo.FullName}文件不存在。");
                return $"Error: {CurrentFileInfo.FullName}文件不存在。";
            }

            using (ExcelPackage package = new ExcelPackage(CurrentFileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == worksheetname)
                    ?? package.Workbook.Worksheets.Add(worksheetname);

                worksheet.Cells.Clear();
                //var endrow = worksheet.Dimension.End.Row;
                //if (endrow >= 3) //第三行开始填写 1行标题、2行表头
                //{
                //    ExcelRange rng = worksheet.Cells[3, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                //    rng.Clear();
                //}
                //获取所有年月字符串
                List<string> monthstr = new List<string>();
                list.ForEach(tj => monthstr.AddRange(tj.GetMonthString()));
#if NET452
                monthstr = monthstr.Distinct().ToList().OrderBy(m => Convert.ToDateTime(m)).ToList();
#else
                monthstr = monthstr.ToHashSet().ToList().OrderBy(m => Convert.ToDateTime(m)).ToList();
#endif


                //填写表头  字体加粗
                List<string> header = new List<string>();

                header.Add("序号");
                header.Add("供应商");
                header.AddRange(monthstr);
                header.Add("年度合计");
                header.Add("备注");
                int col = 1; //第一列写入标题
                header.ForEach(n => worksheet.Cells[1, col++].Value = n);
                worksheet.Cells[1, 1, 1, col - 1].Style.Font.Bold = true;

                //填写内容  第二行开始填写（1行标题）
                int row = 2;
                list.ForEach(t =>
                {
                    worksheet.Cells[row, 1].Value = row - 1;
                    worksheet.Cells[row, 2].Value = t.gongyinshang;
                    //按年月 填充对应金额
                    t.pairs.ForEach(p => {
                        int index = monthstr.IndexOf(p.Key);
                        worksheet.Cells[row, index + 3].Value = p.Value;
                    });
                    //合计金额
                    worksheet.Cells[row, monthstr.Count + 3].Value = t.hejijine;

                    row++;
                });

                //绘制边框
                using (var r = worksheet.Cells[1, 1, row, monthstr.Count + 4])
                {
                    r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    r.Style.Border.Top.Color.SetColor(Color.Black);
                    r.Style.Border.Bottom.Color.SetColor(Color.Black);
                    r.Style.Border.Left.Color.SetColor(Color.Black);
                    r.Style.Border.Right.Color.SetColor(Color.Black);

                    //自适应列宽
                    r.AutoFitColumns();
                    //居中
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                //保存工作表
                package.Save();
                return package.File.FullName;

            }
        }
    }
}

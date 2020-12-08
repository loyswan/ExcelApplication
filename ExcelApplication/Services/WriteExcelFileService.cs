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
        public string WriteDuizhang(List<duizhangModel> list, string worksheetname, string tableTitle = "")
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

                var endrow = (worksheet.Dimension?.End.Row).GetValueOrDefault();
                if (endrow >= 4)
                {
                    //清除4行以后的内容
                    worksheet.Cells[4, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                        .Clear();
                }
                else
                {
                    //新表格填充标题行
                    string tableTitleString = "供应商对账明细表";
                    string tableTitle1String = "        制表人：区晓欣                入库员： 区晓欣                   领导审核：      ";
                    worksheet.Cells[1, 1].Value = tableTitle + tableTitleString;
                    worksheet.Cells[2, 1].Value = tableTitle1String;
                    worksheet.Cells[1, 1].Style.Font.Size = 22;
                    worksheet.Cells[2, 1].Style.Font.Size = 12;
                    //填写标题行 字体加粗
                    string[] tableHeaderString = { "送货日期", "送货单号", "供应商", "物料类别", "物料名称", "宽CM纸纹", "高CM", "数量", "单价", "金额", "备注" };
                    for (int col = 0; col < tableHeaderString.Length; col++)
                    {
                        var rng = worksheet.Cells[3, col + 1];
                        rng.Value = tableHeaderString[col];
                    }

                    worksheet.Cells[1, 1, 1, tableHeaderString.Length].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    worksheet.Cells[2, 1, 2, tableHeaderString.Length].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    using (var headercell = worksheet.Cells[3, 1, 3, tableHeaderString.Length])
                    {
                        headercell.Style.Font.Bold = true;
                        headercell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        headercell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        headercell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        headercell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                }


                //填写内容
                if (list.Count == 0) //没有内容直接保存退出
                {
                    package.Save();
                    return $"Error: 没有符合条件的数据！";
                }
                int row = 4;  //第四行开始写数据
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

                    worksheet.Row(row).Height = 21;//行高 24
                    row++;
                });
                //填写公式 第四行开始写数据
                worksheet.Cells[4, 10, list.Count + 3, 10].Formula = "=IF(G4=0,H4*I4,F4*G4*0.0001*H4*I4)";
                //单元格格式
                worksheet.Cells[4, 1, list.Count + 3, 1].Style.Numberformat.Format = "yyyy/m/d";
                worksheet.Cells[4, 9, list.Count + 3, 10].Style.Numberformat.Format = "0.00";
                //绘制边框
                using (var r = worksheet.Cells[4, 1, list.Count + 3, 11])
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
                var endrow = (worksheet.Dimension?.End.Row).GetValueOrDefault();
                if (endrow >= 4)
                {
                    //清除4行以后的内容
                    worksheet.Cells[4, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                        .Clear();
                }
                else
                {
                    //新表格填充标题行
                    package.Dispose();//此处直接报错
                    return $"Error: 工作簿中无{worksheetname}工作表不存在。";
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
                    worksheet.Row(row).Height = 24;//行高 24
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

                //填写标题行   1、2行标题
                var endrow = (worksheet.Dimension?.End.Row).GetValueOrDefault();
                if (endrow >= 3)
                {
                    //清除3行以后的内容   
                    worksheet.Cells[3, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                        .Clear();
                }
                else
                {
                    //新表格填充标题行
                    string tableTitleString = "供应商月份应付账款表";
                    string tableTitle1String = "制表人：区晓欣";
                    string tableTitle2String = "审核：";
                    string tableTitle3String = "领导审核：";

                    worksheet.Cells[1, 1].Value = tableTitleString;
                    worksheet.Cells[2, 1].Value = tableTitle1String;
                    worksheet.Cells[2, 3].Value = tableTitle2String;
                    worksheet.Cells[2, 5].Value = tableTitle3String;
                    worksheet.Cells[1, 1].Style.Font.Size = 22;
                    worksheet.Cells[2, 1, 2, 5].Style.Font.Size = 12;
                }

                //填写表头行 字体加粗
                string[] tableHeader1String = { "序号", "供应商", "物料类别", "物料名称" };
                string[] tableHeader2String = { "是否含税", "备注" };

                int col = 1;
                //填写第一部分表头
                tableHeader1String.ToList().ForEach(t =>
                {
                    var rng = worksheet.Cells[3, col, 4, col];
                    rng.Merge = true;
                    rng.Value = t;
                    col++;
                });
                //获取所有年月字符串 表头中间部分
                List<string> monthstr = new List<string>();
                list.ForEach(tj => monthstr.AddRange(tj.GetMonthString()));
                monthstr = monthstr.Distinct().ToList().OrderBy(m => Convert.ToDateTime(m)).ToList();
                monthstr.ForEach(m =>
                {
                    var rng = worksheet.Cells[3, col, 3, col + 1];
                    rng.Merge = true;
                    rng.Value = m;
                    worksheet.Cells[4, col].Value = "数量";
                    worksheet.Cells[4, col + 1].Value = "金额";
                    col += 2;
                });
                //填写第二部分表头
                tableHeader2String.ToList().ForEach(t =>
                {
                    var rng = worksheet.Cells[3, col, 4, col];
                    rng.Merge = true;
                    rng.Value = t;
                    col++;
                });
                //表头格式化
                worksheet.Cells[1, 1, 1, col - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;

                using (var headercell = worksheet.Cells[3, 1, 4, col - 1])
                {
                    headercell.Style.Font.Bold = true;
                    headercell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headercell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    headercell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    headercell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    headercell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    headercell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                //填写内容  第5行开始填写
                int row = 5;
                list.ForEach(t =>
                {
                    worksheet.Cells[row, 1].Value = row - 4; //序号
                    worksheet.Cells[row, 2].Value = t.gongyinshang;
                    worksheet.Cells[row, 3].Value = t.wuliaoleibie;
                    worksheet.Cells[row, 4].Value = t.wuliaoming;
                    //按年月 填充对应数量及金额
                    t.pairs.ForEach(p =>
                    {
                        int index = monthstr.IndexOf(p.MonthString);
                        worksheet.Cells[row, index * 2 + tableHeader1String.Length + 1].Value = p.Number;
                        worksheet.Cells[row, index * 2 + tableHeader1String.Length + 2].Value = p.Money;
                    });

                    worksheet.Row(row).Height = 21;//行高 
                    row++;
                });

                //绘制边框  第五行开始
                using (var r = worksheet.Cells[5, 1, row - 1, monthstr.Count * 2 + tableHeader1String.Length + tableHeader2String.Length])
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

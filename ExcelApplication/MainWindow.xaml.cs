using ExcelApplication.Models;
using ExcelApplication.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            //注册Epplus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

#if NET5
            MainView.Title = MainView.Title + " NET 5";
#elif NET452
            MainView.Title = MainView.Title + " NET 4.5.2";
#endif
        }

        private ReadExcelFileService readExcelFileService;
        private WriteExcelFileService writeExcelFileService;
        private ProcessDuiZhangService processDuiZhangService;
        private ProcessTongJiService processTongJiService;
        private SearchDuiZhangService searchDuiZhangService;

        //全局数据
        private List<caigoujinduModel> caigouAlllists;
        private List<duizhangModel> duizhanglists;
        private List<tongjiModel> tongjilists;
        //操作文件
        string filepath = string.Empty;

        private void Card_Drop(object sender, DragEventArgs e)
        {
            string msg = "拖动文件到此处";
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                msg = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

                if (System.IO.Path.GetExtension(msg) == ".xlsx")
                {
                    tbPath.Text = msg;
                    filepath = msg;
                    ClearMessage();

                    //重新拖入文件后，清除原来的数据
                    caigouAlllists = null;
                    duizhanglists = null;
                    tongjilists = null;
                }
                else
                {
                    tbPath.Text = "请拖入电子表格（扩展名 .xlsx )";
                    filepath = string.Empty;
                }
            }
        }




        private void readFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(filepath))
            {
                ShowMessageError("请拖入文件后使用");
                return;
            }
            readExcelFileService = new ReadExcelFileService(filepath);
            var caigoulists = readExcelFileService.GetList("采购进度表"); //读取采购进度表
            var caigouOklists = readExcelFileService.GetList("采购进度表(已完成)"); //读取采购进度表(已完成)

            caigouAlllists = caigoulists.Concat(caigouOklists).ToList();
            ShowMessageInfo("读取完成！");
        }


        private void caigouButton_Click(object sender, RoutedEventArgs e)
        {
            if (caigouAlllists == null || caigouAlllists.Count == 0)
            {
                ShowMessageError("请先读取文件");
                return;
            }
            //如果有已完成的数据，就将完成数据转移并 刷新数据
            if (caigouAlllists.Any(x => x.qiandaoshu <= 0))//有完成数据
            {
                var caigoulists = new List<caigoujinduModel>();
                var caigouOklists = new List<caigoujinduModel>();

                caigouAlllists.ForEach(x =>
                {
                    if (x.qiandaoshu > 0)
                        caigoulists.Add(x);
                    else
                        caigouOklists.Add(x);
                });

                //var caigoulists = caigouAlllists.Where(c => c.qiandaoshu > 0).ToList();
                //var caigouOklists = caigouAlllists.Where(c => c.qiandaoshu <= 0).ToList();
                writeExcelFileService = writeExcelFileService ?? new WriteExcelFileService(filepath);
                var caigoustr = writeExcelFileService.WriteJindu(caigoulists, "采购进度表");
                var caigouokstr = writeExcelFileService.WriteJindu(caigouOklists, "采购进度表(已完成)");
                if (caigoustr.StartsWith("Error") || caigouokstr.StartsWith("Error"))
                    ShowMessageError(caigoustr + "\\n" + caigouokstr);
                else
                    ShowMessageInfo("分类采购进度完成！");
            }
            else
            {
                ShowMessageInfo("没有已完成数据，无需分类。");
            }
        }


        private void duizhangButton_Click(object sender, RoutedEventArgs e)
        {
            if (caigouAlllists == null || caigouAlllists.Count == 0)
            {
                ShowMessageError("请先读取文件");
                return;
            }
            processDuiZhangService = new ProcessDuiZhangService(caigouAlllists);
            duizhanglists = processDuiZhangService.GetList();

            //将数据写入表格
            if (string.IsNullOrEmpty(filepath))
            {
                ShowMessageError("请拖入文件后使用");
                return;
            }
            writeExcelFileService = writeExcelFileService ?? new WriteExcelFileService(filepath);
            var infostr = writeExcelFileService.WriteDuizhang(duizhanglists, "明细对账单");
            if (infostr.StartsWith("Error"))
                ShowMessageError(infostr);
            else
                ShowMessageInfo("刷新对账明细完成！");
        }

        private void tongjiButton_Click(object sender, RoutedEventArgs e)
        {
            if (duizhanglists == null || duizhanglists.Count == 0)
            {
                ShowMessageError("请先刷新对账信息文件");
                return;
            }
            processTongJiService = new ProcessTongJiService(duizhanglists);
            tongjilists = processTongJiService.GetList();

            //将数据写入表格
            if (string.IsNullOrEmpty(filepath))
            {
                ShowMessageError("请拖入文件后使用");
                return;
            }
            writeExcelFileService = writeExcelFileService ?? new WriteExcelFileService(filepath);
            var infostr = writeExcelFileService.WriteTongJi(tongjilists, "供应商账单月统计");
            if (infostr.StartsWith("Error"))
                ShowMessageError(infostr);
            else
                ShowMessageInfo("刷新月度统计完成！");
        }

        private void searchButton_Click(object sender, RoutedEventArgs e)
        {
            //2020-11-24至2020-12-4材料对账明细表
            DateTime startdate = StartDatePicker.SelectedDate ?? DateTime.Today;
            DateTime enddate = EndDatePicker.SelectedDate ?? DateTime.Today;
            if (startdate > enddate)
            {
                ShowMessageError("结束日期不能早于起始日期");
                return;
            }
            string tableTitle = $"{ startdate.ToString("yyyy年M月d日") }至{ enddate.ToString("yyyy年M月d日")}";
            ShowMessageInfo($"当前日期范围：{tableTitle}");

            if (duizhanglists == null || duizhanglists.Count == 0)
            {
                ShowMessageError("请先刷新对账信息文件");
                return;
            }
            searchDuiZhangService = searchDuiZhangService ?? new SearchDuiZhangService();
            searchDuiZhangService.Duizhanglist = duizhanglists;

            var list = searchDuiZhangService.Search(startdate, enddate);
            writeExcelFileService = writeExcelFileService ?? new WriteExcelFileService(filepath);
            var searchstr = writeExcelFileService.WriteDuizhang(list, "输出对账单", tableTitle);

            if (searchstr.StartsWith("Error"))
                ShowMessageError(searchstr);
            else
                ShowMessageInfo("输出对账单完成！");
        }


        //==============================================
        private void ShowMessageError(string msg)
        {
            tbErrorMessage.Text = msg;
            tbErrorMessage.Foreground = new SolidColorBrush(Colors.Red);
            tbErrorMessage.Visibility = Visibility.Visible;
        }
        private void ShowMessageInfo(string msg)
        {
            tbErrorMessage.Text = msg;
            tbErrorMessage.Foreground = new SolidColorBrush(Colors.Green);
            tbErrorMessage.Visibility = Visibility.Visible;
        }
        private void ClearMessage()
        {
            tbErrorMessage.Text = "";
            tbErrorMessage.Visibility = Visibility.Collapsed;
        }
    }
}

using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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

namespace DCTonam_cnf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Thread napdulieu;
        List<dulieu> kho01;
        List<dulieu> kho02;
        List<dulieu> kho05;
        List<dulieu> dc02_01;
        List<dulieu> dc05_01;

        public MainWindow()
        {
            InitializeComponent();
             kho01 = new List<dulieu>();
             kho02 = new List<dulieu>();
             kho05 = new List<dulieu>();
        }
        public List<string> tenfiles { get; set; }
        public string getMakho(string nguon)
        {
            string pat = @"\d{2}$";
            Match ma = Regex.Match(nguon, pat);
            return ma.Value.ToString();
        }
        //check control kho hang
        public void checkControl(string makho, string ngayxuatton)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (makho == "01")
                {
                    checkKho01.Visibility = Visibility.Visible;
                    tenkho01.Text = "Kho layout 01 _ Ngày: " + ngayxuatton; 
                }
                else if (makho == "02")
                {
                    checkKho02.Visibility = Visibility.Visible;
                    tenkho02.Text = "Kho stock 02 _ Ngày: " + ngayxuatton;
                }
                else if (makho == "05")
                {
                    checkKho05.Visibility = Visibility.Visible;
                    tenkho05.Text = "Kho trung chuyển 05 _ Ngày: " + ngayxuatton;
                }
            }));
        }
        //load and unload napdulieu control
        public void fc_loadNapdulieu(bool load)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (load)
                {
                    loadNapdulieu.Visibility = Visibility.Visible;
                    btnXuly.IsEnabled = false;
                }
                else
                {
                    loadNapdulieu.Visibility = Visibility.Hidden;
                    btnXuly.IsEnabled = true;
                }
            }));
        }
        private void btnChonFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog chonfile = new OpenFileDialog();
            chonfile.Filter = "Mời các anh chọn file excel (*.xlsx)|*.xlsx";
            chonfile.Multiselect = true;
            if (chonfile.ShowDialog() == true)
            {
                txtTenFile.Clear();
                tenfiles = new List<string>();
                int sofile = 0;
                foreach (string fileitem in chonfile.FileNames)
                {
                    sofile++;
                    tenfiles.Add(fileitem);
                    txtTenFile.Text += "- File " +sofile+": " + fileitem + "\n";
                    btnNapData.IsEnabled = true;
                }
            }
        }

        private void btnNapData_Click(object sender, RoutedEventArgs e)
        {
            napdulieu = new Thread(fc_napdulieu);
            napdulieu.IsBackground = true;
            napdulieu.Start();
            btnNapData.IsEnabled = false; 
        }
        public void fc_napdulieu()
        {
            string pat_1 = @"Kho CNF\d{6}";
            string pat_2 = @"\d{2}/\d{2}/\d{4}";
            kho01.Clear();
            kho02.Clear();
            kho05.Clear();
            fc_loadNapdulieu(true);
            foreach (string filechon in tenfiles)
            {
                using (var wb = new ExcelPackage(new System.IO.FileInfo(filechon)))
                {
                    using (var ws = wb.Workbook.Worksheets[1])
                    {
                        string oA5 = ws.Cells[5, 1].Value.ToString();
                        if (Regex.IsMatch(oA5, pat_1))
                        {
                            List<dulieu> dltam = new List<dulieu>();
                            Match laytenkho = Regex.Match(oA5, pat_1);
                            Match layngay = Regex.Match(oA5, pat_2);
                            string makho = getMakho(laytenkho.Value);
                            int dongcuoi = ws.Dimension.End.Row;
                            for (int i = 9; i < (dongcuoi); i++)
                            {
                                if (ws.Cells[i,19].Value.ToString() == "0" || ws.Cells[i, 5].Value == null)
                                {
                                    continue;
                                }
                                dltam.Add(new dulieu(ws.Cells[i, 5].Value.ToString(), Convert.ToInt32(ws.Cells[i, 19].Value.ToString())));
                            }
                            if (makho == "01")
                            {
                                kho01 = dltam;
                            }
                            else if (makho == "02")
                            {
                                kho02 = dltam;
                            }
                            else if (makho == "05")
                            {
                                kho05 = dltam;
                            }
                            checkControl(makho, layngay.Value);
                        }
                    }
                }
            }
            fc_loadNapdulieu(false);
        }
        private void btnXuly_Click(object sender, RoutedEventArgs e)
        {
            dc02_01 = new List<dulieu>();
            var kq = kho01.Where(m => m.soluong < 0);
            foreach (var item in kq)
            {
                Console.WriteLine(item.masp + " _ " + item.soluong);
            }
            var kq2 = from a in (from k in kho02
                                 where k.soluong > 0
                                 select k)
                     join b in (from x in kho01
                                where x.soluong < 0
                                select x)
                     on a.masp equals b.masp
                     select new
                     {
                         masp = a.masp,
                         soluong2 = a.soluong,
                         soluong1 = b.soluong
                     };
            Console.WriteLine("\nsoluong luc loc voi kho 02\n");
            foreach (var item in kq2)
            {
                Console.WriteLine(item.masp + " _ " + item.soluong1 + " _ " + item.soluong2);
            }
            List<dulieu> kho011 = new List<dulieu>();
            foreach (var rs in kq2)
            {
                if ((rs.soluong1*(-1)) < rs.soluong2)
                {
                    dc02_01.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho011.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc02_01.Add(new dulieu(rs.masp, rs.soluong2));
                    kho011.Add(new dulieu(rs.masp,rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho011)
            {
                var itemToChange = kq.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            Console.WriteLine("\nSo lieu sau \n");                  
            foreach (var item in kq)
            {
                Console.WriteLine(item.masp + " _ " + item.soluong);
            }
        }
    }
}

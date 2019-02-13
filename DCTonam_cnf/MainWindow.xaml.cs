using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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
        List<dulieu> dcnhap_01;
        List<dulieu> dc01_02;
        List<dulieu> dc05_02;
        List<dulieu> dcnhap_02;
        List<dulieu> dc01_05;
        List<dulieu> dc02_05;
        List<dulieu> dcnhap_05;
        List<List<dulieu>> luu_datatach = new List<List<dulieu>>();
        string ngaydieuchinh = string.Empty;
        string folder_copy = string.Empty;
        string duongdangoc = System.AppDomain.CurrentDomain.BaseDirectory + @"\File_Xuat_Excel";
        string duongdanfile = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
             kho01 = new List<dulieu>();
             kho02 = new List<dulieu>();
             kho05 = new List<dulieu>();
            Directory.CreateDirectory(duongdangoc);
            
        }
        public List<string> tenfiles { get; set; }
        public string getMakho(string nguon)
        {
            string kq = string.Empty;
            string pat = @"\d{2}$";
            string pat3 = @"\d{6}$";
            if (Regex.IsMatch(nguon, pat3))
            {
                Match ma = Regex.Match(nguon, pat);
                kq = ma.Value.ToString();
            }
            else
            {
                kq = "05";
            }
            return kq;
        }
        //check control kho hang
        public void checkControl(string makho, string ngayxuatton, string tenkho)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (makho == "01")
                {
                    checkKho01.IsChecked = true;
                    tenkho01.Text = "Kho layout 01 _ "+ tenkho +" _Ngày: " + ngayxuatton; 
                }
                else if (makho == "02")
                {
                    checkKho02.IsChecked = true;
                    tenkho02.Text = "Kho stock 02 _ "+ tenkho + " _Ngày: " + ngayxuatton;
                }
                else if (makho == "05")
                {
                    checkKho05.IsChecked = true;
                    tenkho05.Text = "Kho trung chuyển 05 _" + tenkho + " _Ngày: " + ngayxuatton;
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
            card_thongbao.Visibility = Visibility.Hidden;
            btnXuatexcel.Visibility = Visibility.Hidden;
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
                    btnXuly.IsEnabled = false;
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
            string pat_1 = @"Kho CNF(\d{6}|\d{4})";
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
                            for (int i = 9; i < (dongcuoi + 1); i++)
                            {
                                if (ws.Cells[i, 19].Value == null || ws.Cells[i,19].Value.ToString() == "0" ||  ws.Cells[i, 5].Value == null)
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
                            checkControl(makho, layngay.Value, laytenkho.Value);
                            ngaydieuchinh = layngay.Value;
                        }
                    }
                }
            }
            fc_loadNapdulieu(false);
        }
        private void btnXuly_Click(object sender, RoutedEventArgs e)
        {
            dc02_01 = new List<dulieu>();
            dc05_01 = new List<dulieu>();
            dcnhap_01 = new List<dulieu>();
            dc01_02 = new List<dulieu>();
            dc05_02 = new List<dulieu>();
            dcnhap_02 = new List<dulieu>();
            dc02_05 = new List<dulieu>();
            dc01_05 = new List<dulieu>();
            dcnhap_05 = new List<dulieu>();

            #region dc ton am kho 01
            var kq1 = kho01.Where(m => m.soluong < 0);
            var kq1_2 = from a in (from k in kho02
                                 where k.soluong > 0
                                 select k)
                      join b in kq1
                      on a.masp equals b.masp
                      select new
                      {
                          masp = a.masp,
                          soluong2 = a.soluong,
                          soluong1 = b.soluong
                      };
            List<dulieu> kho012 = new List<dulieu>();
            //dc tu 02 vao 01
            foreach (var rs in kq1_2)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc02_01.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho012.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc02_01.Add(new dulieu(rs.masp, rs.soluong2));
                    kho012.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho012)
            {
                var itemToChange = kq1.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            // dc tu 05 vao 01
            var kq1_5 = from a in (from k in kho05
                                  where k.soluong > 0
                                  select k)
                       join b in kq1
                       on a.masp equals b.masp
                       select new
                       {
                           masp = a.masp,
                           soluong2 = a.soluong,
                           soluong1 = b.soluong
                       };
            List<dulieu> kho015 = new List<dulieu>();
            foreach (var rs in kq1_5)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc05_01.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho015.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc05_01.Add(new dulieu(rs.masp, rs.soluong2));
                    kho015.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho015)
            {
                var itemToChange = kq1.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            foreach (var item in kq1)
            {
                dcnhap_01.Add(new dulieu(item.masp, item.soluong*(-1)));
            }
            #endregion
            #region dc ton am kho 02
            var kq2 = kho02.Where(m => m.soluong < 0);
            var kq2_1 = from a in (from k in kho01
                                   where k.soluong > 0
                                   select k)
                        join b in kq2
                        on a.masp equals b.masp
                        select new
                        {
                            masp = a.masp,
                            soluong2 = a.soluong,
                            soluong1 = b.soluong
                        };
            List<dulieu> kho021 = new List<dulieu>();
            //dc tu 01 vao 02
            foreach (var rs in kq2_1)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc01_02.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho021.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc01_02.Add(new dulieu(rs.masp, rs.soluong2));
                    kho021.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho021)
            {
                var itemToChange = kq2.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            // dc tu 05 vao 02
            var kq2_5 = from a in (from k in kho05
                                   where k.soluong > 0
                                   select k)
                        join b in kq2
                        on a.masp equals b.masp
                        select new
                        {
                            masp = a.masp,
                            soluong2 = a.soluong,
                            soluong1 = b.soluong
                        };
            List<dulieu> kho025 = new List<dulieu>();
            foreach (var rs in kq2_5)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc05_02.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho025.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc05_02.Add(new dulieu(rs.masp, rs.soluong2));
                    kho025.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho025)
            {
                var itemToChange = kq2.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            //dc nhap kho 02
            foreach (var item in kq2)
            {
                dcnhap_02.Add(new dulieu(item.masp, item.soluong * (-1)));
            }
            #endregion
            #region dc ton am kho 05
            var kq5 = kho05.Where(m => m.soluong < 0);
            var kq5_1 = from a in (from k in kho01
                                   where k.soluong > 0
                                   select k)
                        join b in kq5
                        on a.masp equals b.masp
                        select new
                        {
                            masp = a.masp,
                            soluong2 = a.soluong,
                            soluong1 = b.soluong
                        };
            List<dulieu> kho051 = new List<dulieu>();
            //dc tu 01 vao 05
            foreach (var rs in kq5_1)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc01_05.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho051.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc01_05.Add(new dulieu(rs.masp, rs.soluong2));
                    kho051.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho051)
            {
                var itemToChange = kq5.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            // dc tu 02 vao 05
            var kq5_2 = from a in (from k in kho02
                                   where k.soluong > 0
                                   select k)
                        join b in kq5
                        on a.masp equals b.masp
                        select new
                        {
                            masp = a.masp,
                            soluong2 = a.soluong,
                            soluong1 = b.soluong
                        };
            List<dulieu> kho052 = new List<dulieu>();
            foreach (var rs in kq2_5)
            {
                if ((rs.soluong1 * (-1)) <= rs.soluong2)
                {
                    dc02_05.Add(new dulieu(rs.masp, rs.soluong1 * (-1)));
                    kho052.Add(new dulieu(rs.masp, 0));
                }
                else if ((rs.soluong1 * (-1)) > rs.soluong2)
                {
                    dc02_05.Add(new dulieu(rs.masp, rs.soluong2));
                    kho052.Add(new dulieu(rs.masp, rs.soluong1 + rs.soluong2));
                }
            }
            foreach (var x in kho052)
            {
                var itemToChange = kq5.First(d => d.masp == x.masp).soluong = x.soluong;
            }
            //dc nhap kho 05
            foreach (var item in kq5)
            {
                dcnhap_05.Add(new dulieu(item.masp, item.soluong * (-1)));
            }
            #endregion
            btnXuatexcel.Visibility = Visibility.Visible;
        }
        private void btnXuatexcel_Click(object sender, RoutedEventArgs e)
        {
            ngaydieuchinh = ngaydieuchinh.Replace("/", "-");
            Directory.CreateDirectory(duongdangoc + @"\" + ngaydieuchinh);
            folder_copy = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\DC_ton_am";
            Directory.CreateDirectory(folder_copy + @"\" + ngaydieuchinh);

            xuatexcel(dc02_01, "02_01", ngaydieuchinh);
            xuatexcel(dc05_01, "05_01", ngaydieuchinh);
            xuatexcel(dcnhap_01, "nhap_01", ngaydieuchinh);
            xuatexcel(dc01_02, "01_02", ngaydieuchinh);
            xuatexcel(dc05_02, "05_02", ngaydieuchinh);
            xuatexcel(dcnhap_02, "nhap_02", ngaydieuchinh);
            xuatexcel(dc01_05, "01_05", ngaydieuchinh);
            xuatexcel(dc02_05, "02_05", ngaydieuchinh);
            xuatexcel(dcnhap_05, "nhap_05", ngaydieuchinh);
            
            card_thongbao.Visibility = Visibility.Visible;
            tbnoidungthongbao.Text = "Done !!! .Vừa xuất file tại đường dẫn: " + folder_copy;
        }
        //ham xuat excel
        public void xuatexcel(List<dulieu> data, string tenfileout, string ngay)
        {
            if (!data.Any())
            {
                return;
            }
            dequy_tachdata(data, 0); // sau khi goi ham nay se co du lieu tai bien 'luu_datatach'
            int vs = 0;
            foreach (var item in luu_datatach)
            {
                vs++;
                using (ExcelPackage ex = new ExcelPackage())
                {
                    using (ExcelWorksheet ws = ex.Workbook.Worksheets.Add("dc_ton_am_" + tenfileout + "_vs" + vs))
                    {
                        if (!item.Any())
                        {
                            return;
                        }
                        
                        ws.Cells["A2"].LoadFromCollection(item);
                        ws.Column(1).AutoFit();
                        int tongma = 0;
                        int dongcuoi = ws.Dimension.End.Row;
                        for (int i = 2; i <= dongcuoi; i++)
                        {
                            tongma = tongma + Convert.ToInt32(ws.Cells[i, 2].Value.ToString());
                        }
                        duongdanfile = @"\" + ngaydieuchinh + @"\" + "DC_Ton_Am_" + tenfileout + "= " + tongma + "sp_ngay_" + ngay + "_vs" + vs + ".xlsx";
                        if (File.Exists(duongdangoc + duongdanfile))
                        {
                            File.Delete(duongdangoc + duongdanfile);
                        }
                        ex.SaveAs(new FileInfo(duongdangoc + duongdanfile));
                        File.Copy(duongdangoc + duongdanfile, folder_copy + duongdanfile, true);
                    }
                }
            }
            luu_datatach.Clear();
        }
        //dequy de tach data khi so item > 297 ma
        
        public void dequy_tachdata(List<dulieu> data, int ts)
        {
            List<dulieu> hh_1 = new List<dulieu>();
            int ts_1 = 0;
            int ts_2 = ts;
            if (data.Count() < 297)
            {
                for (int i = 0; i < data.Count(); i++)
                {
                    hh_1.Add(new dulieu(data[i].masp, data[i].soluong));
                }
                luu_datatach.Add(hh_1);
                return;
            }
            for (int i = ts; i < data.Count(); i++)
            {
                
                ts_1++;
                hh_1.Add(new dulieu(data[i].masp, data[i].soluong));
                if (ts_1 > 297)
                {
                    ts_2 = i + 1;
                    break;
                }
            }
            luu_datatach.Add(hh_1);
            if (ts_1 > 297)
            {
                dequy_tachdata(data, ts_2);
            }
        }
        
    }
}

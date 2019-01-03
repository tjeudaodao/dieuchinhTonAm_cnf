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
        public MainWindow()
        {
            InitializeComponent();
        }
        public List<string> tenfiles { get; set; }
        private void btnChonFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog chonfile = new OpenFileDialog();
            chonfile.Filter = "Mời các anh chọn file excel (*.xlsx)|*.xlsx";
            chonfile.Multiselect = true;
            if (chonfile.ShowDialog() == true)
            {
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
            foreach (string filechon in tenfiles)
            {
                using (var wb = new ExcelPackage(new System.IO.FileInfo(filechon)))
                {
                    using (var ws = wb.Workbook.Worksheets[1])
                    {
                        string oA5 = ws.Cells[5, 1].Value.ToString();
                        if (Regex.IsMatch(oA5, pat_1))
                        {
                            
                            Match laytenkho = Regex.Match(oA5, pat_1);
                            Match layngay = Regex.Match(oA5, pat_2);
                            MessageBox.Show(laytenkho.Value + "\n" + layngay.Value);
                        }
                    }
                }
            }
        }
        private void btnXuly_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(" van chay ");
        }
    }
}

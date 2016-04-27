﻿using MEIKCount.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
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

namespace MEIKCount
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string meikFolder = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        private ObservableCollection<Patient> patientCollection;
        private Dictionary<string, int> countDict = new Dictionary<string, int>();
        private IList<Patient> patientList=new List<Patient>();
        //private int times = 0;

        #region public Properties
        public ObservableCollection<Patient> PatientCollection
        {
            get
            {
                return this.patientCollection;
            }
            set
            {
                this.patientCollection = value;                
            }
        }
       
        #endregion

        public MainWindow()
        {
            if (File.Exists(meikFolder + "\\Language.CHN"))
            {
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
                
            }
            InitializeComponent();
            CountScreeningTimes();            
            

        }

        private void CountScreeningTimes()
        {
            //string meikPath = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
            //string meikPath = @"C:\Users\CampRay\Desktop\15010700103 - Silverio, Katherine V";
            //string meikPath = @"C:\Users\CampRay\Desktop\14102500104 - Pangilinan, Grace L";
            if (string.IsNullOrEmpty(meikFolder))
            {
                meikFolder = @"C:\Program Files (x86)\MEIK 5.6";
            }
            //string meikPath = @"C:\Program Files (x86)\MEIK 5.6";
            
            try
            {
                ListFiles(new DirectoryInfo(meikFolder)); 
                
                //List<Patient> patientList = new List<Patient>();
                //foreach (KeyValuePair<string, int> kvp in countDict)
                //{
                //    Patient patient = new Patient();
                //    patient.Code = kvp.Key.Split(";".ToCharArray())[0];
                //    patient.Name = kvp.Key.Split(";".ToCharArray())[1];
                //    patient.Times = kvp.Value;
                //    patientList.Add(patient);                    
                //}
                patientCollection = new ObservableCollection<Patient>(patientList);
                this.patientGrid.ItemsSource = patientCollection;
                this.txtTimes.Text = patientList.Count + "";
                this.patientTimes.Text = countDict.Count + "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to count the screening times. Exception: " + ex.Message);
            }
        }

        private void ListFiles(FileSystemInfo info)
        {
            if (!info.Exists) return;
            DirectoryInfo dir = info as DirectoryInfo;
            //不是目录 
            if (dir == null) return;
            if (!dir.Name.Equals("NORM"))
            {
                FileSystemInfo[] files = dir.GetFileSystemInfos();
                for (int i = 0; i < files.Length; i++)
                {
                    FileInfo file = files[i] as FileInfo;
                    //是文件 
                    if (file != null)
                    {
                        if (".tdb".Equals(file.Extension, StringComparison.OrdinalIgnoreCase))
                        {
                            FileStream fsRead = new FileStream(file.FullName, FileMode.Open);
                            byte[] nameBytes = new byte[105];
                            byte[] codeBytes = new byte[11];
                            byte[] descBytes = new byte[200];
                            fsRead.Seek(12, SeekOrigin.Begin);
                            fsRead.Read(nameBytes, 0, nameBytes.Count());
                            fsRead.Seek(117, SeekOrigin.Begin);
                            fsRead.Read(codeBytes, 0, codeBytes.Count());
                            fsRead.Seek(129, SeekOrigin.Begin);
                            fsRead.Read(descBytes, 0, descBytes.Count());
                            fsRead.Close();
                            string name = System.Text.Encoding.ASCII.GetString(nameBytes);
                            name = name.Split("\0".ToCharArray())[0];
                            string code = System.Text.Encoding.ASCII.GetString(codeBytes);
                            string desc = System.Text.Encoding.ASCII.GetString(descBytes);
                            desc = desc.Split("\0".ToCharArray())[0];
                            var patient = new Patient();
                            patient.Code = code;
                            patient.Name = name;
                            patient.Desc = desc;
                            patient.ScreenDate = file.LastWriteTime.ToString();
                            patientList.Add(patient);
                            var key = code + ";" + name;
                            if (countDict.ContainsKey(key))
                            {
                                countDict[key] = countDict[key] + 1;
                            }
                            else
                            {
                                countDict.Add(key, 1);
                            }
                            //times++;
                        }
                    }
                    //对于子目录，进行递归调用 
                    else
                    {
                        ListFiles(files[i]);
                    }
                }
            }
        }

        private void export_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "xls|*.xls" };
            if (dlg.ShowDialog(this) == true)
            {                
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var books = (Microsoft.Office.Interop.Excel.Workbooks)excelApp.Workbooks;
                var book = (Microsoft.Office.Interop.Excel._Workbook)(books.Add(System.Type.Missing));
                var sheets = (Microsoft.Office.Interop.Excel.Sheets)book.Worksheets;
                var sheet = (Microsoft.Office.Interop.Excel._Worksheet)(sheets.get_Item(1));


                Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("A1", "A100");
                range.ColumnWidth = 15;
                range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("B1", "B100");
                range.ColumnWidth = 30;
                range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("C1", "C100");
                range.ColumnWidth = 20;
                range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("D1", "D100");
                range.ColumnWidth = 100;
                sheet.Cells[1, 1] = App.Current.FindResource("MainWin3").ToString();
                sheet.Cells[1, 2] = App.Current.FindResource("MainWin4").ToString();
                sheet.Cells[1, 3] = App.Current.FindResource("MainWin6").ToString();
                sheet.Cells[1, 4] = App.Current.FindResource("MainWin5").ToString();
                for (int i = 0; i < patientList.Count; i++)
                {
                    Patient item = patientList[i];
                    sheet.Cells[i + 2, 1] = item.Code;
                    sheet.Cells[i + 2, 2] = item.Name;
                    sheet.Cells[i + 2, 3] = item.ScreenDate;
                    sheet.Cells[i + 2, 4] = item.Desc;
                }
                book.SaveAs(dlg.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close();
                excelApp.Quit();
            }            

        } 


    }
}

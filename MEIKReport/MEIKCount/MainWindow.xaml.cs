using MEIKCount.Common;
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
        private ObservableCollection<Patient> patientCollection;
        private Dictionary<string, int> countDict = new Dictionary<string, int>();
        private int times = 0;

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
            InitializeComponent();
            CountScreeningTimes();            
            

        }

        private void CountScreeningTimes()
        {
            string meikPath = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
            //string meikPath = @"C:\Users\CampRay\Desktop\15010700103 - Silverio, Katherine V";
            //string meikPath = @"C:\Users\CampRay\Desktop\14102500104 - Pangilinan, Grace L";
            //string meikPath = @"C:\Program Files (x86)\MEIK 5.6";
            
            try
            {
                ListFiles(new DirectoryInfo(meikPath)); 
                
                List<Patient> patientList = new List<Patient>();
                foreach (KeyValuePair<string, int> kvp in countDict)
                {
                    Patient patient = new Patient();
                    patient.Code = kvp.Key.Split(";".ToCharArray())[0];
                    patient.Name = kvp.Key.Split(";".ToCharArray())[1];
                    patient.Times = kvp.Value;
                    patientList.Add(patient);                    
                }
                patientCollection = new ObservableCollection<Patient>(patientList);
                this.patientGrid.ItemsSource = patientCollection;                
                this.txtTimes.Text = times+"";
                this.patientTimes.Text = patientList.Count+"";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to count the screening times. Exception: " + ex.Message);
            }
        }

        public void ListFiles(FileSystemInfo info)
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
                            fsRead.Seek(12, SeekOrigin.Begin);
                            fsRead.Read(nameBytes, 0, nameBytes.Count());
                            fsRead.Seek(117, SeekOrigin.Begin);
                            fsRead.Read(codeBytes, 0, codeBytes.Count());
                            fsRead.Close();
                            string name = System.Text.Encoding.ASCII.GetString(nameBytes);
                            name = name.Split("\0".ToCharArray())[0];
                            string code = System.Text.Encoding.ASCII.GetString(codeBytes);
                            var key = code + ";" + name;
                            if (countDict.ContainsKey(key))
                            {
                                countDict[key] = countDict[key] + 1;
                            }
                            else
                            {
                                countDict.Add(key, 1);
                            }
                            times++;
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


    }
}

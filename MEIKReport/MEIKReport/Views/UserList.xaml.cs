using MEIKReport.Common;
using MEIKReport.Model;
using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Data;
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
using System.Xml;

namespace MEIKReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class UserList : Window
    {
        public UserList()
        {
            InitializeComponent();            
        }        

        private void ExaminationReport_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            //var name=selectedUser.GetAttribute("Name");
            // View Examination Report
            ExaminationReportPage examinationReportPage = new ExaminationReportPage(selectedUser);
            //examinationReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            examinationReportPage.Owner = this;
            examinationReportPage.ShowDialog();     
        }

        private void SummaryReport_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            //var name = selectedUser.GetAttribute("Name");
            SummaryReportPage summaryReportPage = new SummaryReportPage(selectedUser);
            ////summaryReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            summaryReportPage.Owner = this;
            summaryReportPage.ShowDialog();
        }

        private void ShowMainWindow(object sender, EventArgs e)
        {
            this.Show();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //this.Owner.Show();
            this.Owner.Visibility=Visibility.Visible;
        }

        private void btnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();            
            folderBrowserDialog.ShowDialog();
            
            string folderName = folderBrowserDialog.SelectedPath;
            App.dataFolder = folderName;
            txtFolderPath.Text = folderName;
            CollectionViewSource customerSource = (CollectionViewSource)this.FindResource("CustomerSource");
            HashSet<Person> set=new HashSet<Person>();
            //遍历指定文件夹下所有文件
            DirectoryInfo theFolder = new DirectoryInfo(folderName);
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            List<DirectoryInfo> list=dirInfo.ToList();
            list.Add(theFolder);
            dirInfo = list.ToArray();
            foreach (DirectoryInfo NextFolder in dirInfo)
            {                
                FileInfo[] fileInfo = NextFolder.GetFiles();
                //遍历文件
                foreach (FileInfo NextFile in fileInfo)
                {
                    if (".crd".Equals(NextFile.Extension, StringComparison.OrdinalIgnoreCase))
                    {
                        string surname = OperateIniFile.ReadIniData("Personal data", "surname", "", NextFile.FullName);
                        string address = OperateIniFile.ReadIniData("Personal data", "address", "", NextFile.FullName);
                        string birthdate = OperateIniFile.ReadIniData("Personal data", "birth date", "", NextFile.FullName);
                        string birthmonth = OperateIniFile.ReadIniData("Personal data", "birth month", "", NextFile.FullName);
                        string birthyear = OperateIniFile.ReadIniData("Personal data", "birth year", "", NextFile.FullName);
                        string registrationdate = OperateIniFile.ReadIniData("Personal data", "registration date", "", NextFile.FullName);
                        string registrationmonth = OperateIniFile.ReadIniData("Personal data", "registration month", "", NextFile.FullName);
                        string registrationyear = OperateIniFile.ReadIniData("Personal data", "registration year", "", NextFile.FullName);
                        Person person = new Person();
                        person.Code = NextFile.Name.Substring(0, NextFile.Name.Length-4);
                        person.SurName = surname;
                        person.Address = address;
                        person.Birthday = birthmonth + "/" + birthdate + "/" + birthyear;
                        person.Regdate = registrationmonth + "/" + registrationdate + "/" + registrationyear;
                        if (!string.IsNullOrEmpty(person.Birthday))
                        {                            
                            int m_Y1 = DateTime.Parse(person.Birthday).Year;
                            int m_Y2 = DateTime.Now.Year;                            
                            person.Age = m_Y2 - m_Y1;
                        }
                        set.Add(person);
                    }
                }
            }
                     
            customerSource.Source = set;
            if (set.Count > 0)
            {
                reportButtonPanel.Visibility = Visibility.Visible;
            }
            else
            {
                reportButtonPanel.Visibility = Visibility.Hidden;
            }
        }

        private void exitReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Owner.Show();
        }
        
        
    }
}

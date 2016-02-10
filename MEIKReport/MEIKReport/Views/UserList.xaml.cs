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
        private ReportSettingModel reportSettingModel = new ReportSettingModel();
        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "Data";
        public UserList()
        {
            InitializeComponent();
            LoadInitConfig();
        }        

        private void ExaminationReport_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
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
            this.Visibility = Visibility.Hidden;
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
            try
            {
                App.opendWin = null;
                IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
                //如果主窗体存在
                if (mainWinHwnd != IntPtr.Zero)
                {
                    int WM_SYSCOMMAND = 0x0112;
                    int SC_CLOSE = 0xF060;
                    Win32Api.SendMessage(mainWinHwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Owner.Visibility = Visibility.Visible;                
            }
        }

        private void btnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            try { 
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
                    FileInfo[] fileInfo = null;
                    try
                    {
                        fileInfo = NextFolder.GetFiles();
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
                                person.Code = NextFile.Name.Substring(0, NextFile.Name.Length - 4);
                                person.SurName = surname;
                                person.Address = address;
                                person.Birthday = birthmonth + "/" + birthdate + "/" + birthyear;
                                person.Regdate = registrationmonth + "/" + registrationdate + "/" + registrationyear;
                                person.ArchiveFolder = folderName;
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
                    catch (Exception){ }
                
                }
                     
                customerSource.Source = set;
                if (set.Count > 0)
                {
                    reportButtonPanel.Visibility = Visibility.Visible;
                    emailButton.Visibility = Visibility.Visible;
                }
                else
                {
                    reportButtonPanel.Visibility = Visibility.Hidden;
                    emailButton.Visibility = Visibility.Hidden;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void exitReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Owner.Show();
        }

        private void EmailButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(reportSettingModel.MailAddress) || string.IsNullOrEmpty(reportSettingModel.MailHost)
                || string.IsNullOrEmpty(reportSettingModel.MailUsername) || string.IsNullOrEmpty(reportSettingModel.MailPwd))
            {
                MessageBox.Show("You don't configure the mail account or configured the mail account incorrectly.");
                return;
            }
            string toMail="";
            MessageBoxResult result = MessageBox.Show("Do you want to email the report files to " + toMail + "?", "Email the report", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var selectedUser = this.CodeListBox.SelectedItem as Person;
                    string senderServerIp = reportSettingModel.MailHost;
                    string toMailAddress = toMail;
                    string fromMailAddress = reportSettingModel.MailAddress;
                    string subjectInfo =reportSettingModel.MailSubject+ " ("+selectedUser.SurName +")";
                    string bodyInfo = "";
                    string mailUsername = reportSettingModel.MailUsername;
                    string mailPassword = reportSettingModel.MailPwd;
                    string mailPort = reportSettingModel.MailPort+"";
                    string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
                    string pdfFile=dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_LF.pdf";
                    string attachPath = dataFile + ";" + pdfFile;
                    bool isSsl = reportSettingModel.MailSsl;
                    EmailHelper email = new EmailHelper(senderServerIp, toMailAddress, fromMailAddress, subjectInfo, bodyInfo, mailUsername, mailPassword, mailPort, isSsl, false);
                    email.AddAttachments(attachPath);
                    email.Send();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        private void LoadInitConfig()
        {
            try
            {                
                string doctorNames = OperateIniFile.ReadIniData("Report", "Doctor Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                if (!string.IsNullOrEmpty(doctorNames))
                {
                    var doctorList = doctorNames.Split(';').ToList<string>();
                    doctorList.ForEach(item => reportSettingModel.DoctorNames.Add(item));
                }
                string techNames = OperateIniFile.ReadIniData("Report", "Technician Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                if (!string.IsNullOrEmpty(doctorNames))
                {
                    var techList = techNames.Split(';').ToList<string>();
                    techList.ForEach(item => reportSettingModel.TechNames.Add(item));
                }
                reportSettingModel.PrintPaper = OperateIniFile.ReadIniData("Report", "Print Paper", "Letter", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");                
                reportSettingModel.MailAddress = OperateIniFile.ReadIniData("Mail", "Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailHost = OperateIniFile.ReadIniData("Mail", "Mail Host", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailPort = Convert.ToInt32(OperateIniFile.ReadIniData("Mail", "Mail Port", "25", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                reportSettingModel.MailUsername = OperateIniFile.ReadIniData("Mail", "Mail Username", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailPwd = OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load the report setting. Exception: " + ex.Message);
            }            
        }
    }
}

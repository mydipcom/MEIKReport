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
    public partial class ReportSettingPage : Window
    {
        
        public ReportSettingPage()
        {
            InitializeComponent();
            LoadInitConfig();
            if ("Letter".Equals(App.reportSettingModel.PrintPaper, StringComparison.OrdinalIgnoreCase))
            {
                radLetter.IsChecked = true;
                radA4.IsChecked = false;
            }
            else
            {
                radLetter.IsChecked = false;
                radA4.IsChecked = true;
            }
            this.tabSetting.DataContext = App.reportSettingModel;
        }

        

        private void Window_Closed(object sender, EventArgs e)
        {
            this.Owner.Visibility = Visibility.Visible;
        }

        private void btnReportSettingSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(txtMailPwd.Password))
                {
                    App.reportSettingModel.MailPwd = txtMailPwd.Password;
                }
                OperateIniFile.WriteIniData("Base", "MEIK base", App.reportSettingModel.MeikBase, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Use Default Signature", App.reportSettingModel.UseDefaultSignature.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Doctor Names List", string.Join(";", App.reportSettingModel.DoctorNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Technician Names List", string.Join(";", App.reportSettingModel.TechNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Print Paper", App.reportSettingModel.PrintPaper, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "My Mail Address", App.reportSettingModel.MailAddress, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "To Mail Address", App.reportSettingModel.ToMailAddress, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Subject", App.reportSettingModel.MailSubject, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Content", App.reportSettingModel.MailSubject, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Host", App.reportSettingModel.MailHost, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Port", App.reportSettingModel.MailPort.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Username", App.reportSettingModel.MailUsername, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                string mailPwd = App.reportSettingModel.MailPwd;
                if (!string.IsNullOrEmpty(mailPwd))
                {
                    mailPwd = SecurityTools.EncryptText(mailPwd);
                }
                OperateIniFile.WriteIniData("Mail", "Mail Password", mailPwd, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail SSL", App.reportSettingModel.MailSsl.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");

                OperateIniFile.WriteIniData("Device", "Device No", App.reportSettingModel.DeviceNo.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Device", "Device Type", App.reportSettingModel.DeviceType.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");                
                MessageBox.Show(App.Current.FindResource("Message_14").ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(App.Current.FindResource("Message_13").ToString()+" "+ex.Message);
            }
        }

        private void btnReportSettingClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Owner.Visibility = Visibility.Visible;
        }

        private void tabMail_GotFocus(object sender, RoutedEventArgs e)
        {
            var title = Application.Current.FindResource("SettingEmailTitle").ToString();
            var desc = Application.Current.FindResource("SettingEmailDesc").ToString();
            labSettingTitle.Content = title;
            labSettingDesc.Content = desc;
        }

        private void tabReport_GotFocus(object sender, RoutedEventArgs e)
        {
            var title = Application.Current.FindResource("SettingReportTitle").ToString();
            var desc = Application.Current.FindResource("SettingReportDesc").ToString();
            labSettingTitle.Content = title;
            labSettingDesc.Content = desc;
        }

        private void btnAddDoctor_Click(object sender, RoutedEventArgs e)
        {
            AddNamePage addNamePage = new AddNamePage(App.reportSettingModel.DoctorNames);            
            addNamePage.ShowDialog();
        }

        private void btnDelDoctor_Click(object sender, RoutedEventArgs e)
        {
            if (this.listDoctorName.SelectedIndex != -1)
            {
                App.reportSettingModel.DoctorNames.RemoveAt(this.listDoctorName.SelectedIndex);
            }
        }

        private void btnAddTech_Click(object sender, RoutedEventArgs e)
        {
            AddNamePage addNamePage = new AddNamePage(App.reportSettingModel.TechNames);            
            addNamePage.ShowDialog();

        }

        private void btnDelTech_Click(object sender, RoutedEventArgs e)
        {
            if (this.listTechnicianName.SelectedIndex != -1)
            {
                App.reportSettingModel.TechNames.RemoveAt(this.listTechnicianName.SelectedIndex);
            }
        }

        private void LoadInitConfig()
        {
            try
            {
                if (App.reportSettingModel == null)
                {
                    App.reportSettingModel = new ReportSettingModel();
                    App.reportSettingModel.MeikBase = OperateIniFile.ReadIniData("Base", "MEIK base", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    string doctorNames = OperateIniFile.ReadIniData("Report", "Doctor Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(doctorNames))
                    {
                        var doctorList = doctorNames.Split(';').ToList<string>();
                        doctorList.ForEach(item => App.reportSettingModel.DoctorNames.Add(item));
                    }
                    string techNames = OperateIniFile.ReadIniData("Report", "Technician Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(doctorNames))
                    {
                        var techList = techNames.Split(';').ToList<string>();
                        techList.ForEach(item => App.reportSettingModel.TechNames.Add(item));
                    }
                    App.reportSettingModel.PrintPaper = OperateIniFile.ReadIniData("Report", "Print Paper", "Letter", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailAddress = OperateIniFile.ReadIniData("Mail", "My Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.ToMailAddress = OperateIniFile.ReadIniData("Mail", "To Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailSubject = OperateIniFile.ReadIniData("Mail", "Mail Subject", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailBody = OperateIniFile.ReadIniData("Mail", "Mail Content", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailHost = OperateIniFile.ReadIniData("Mail", "Mail Host", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailPort = Convert.ToInt32(OperateIniFile.ReadIniData("Mail", "Mail Port", "25", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.MailUsername = OperateIniFile.ReadIniData("Mail", "Mail Username", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    string mailPwd=OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if(string.IsNullOrEmpty(mailPwd)){
                        App.reportSettingModel.MailPwd = SecurityTools.DecryptText(mailPwd);
                    }                    
                    App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.DeviceNo = OperateIniFile.ReadIniData("Device", "Device No", "000", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.DeviceType = Convert.ToInt32(OperateIniFile.ReadIniData("Device", "Device Type", "1", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(App.Current.FindResource("Message_9").ToString()+" " + ex.Message);
            }
        }
    }
}

﻿using MEIKReport.Common;
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
        private ReportSettingModel reportSettingModel = new ReportSettingModel();
        public ReportSettingPage()
        {
            InitializeComponent();
            LoadInitConfig();            
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
                if ("Letter".Equals(reportSettingModel.PrintPaper, StringComparison.OrdinalIgnoreCase))
                {
                    radLetter.IsChecked = true;
                    radA4.IsChecked = false;
                }
                else
                {
                    radLetter.IsChecked = false;
                    radA4.IsChecked = true;
                }
                reportSettingModel.MailAddress = OperateIniFile.ReadIniData("Mail", "My Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.ToMailAddress = OperateIniFile.ReadIniData("Mail", "To Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailSubject = OperateIniFile.ReadIniData("Mail", "Mail Subject", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailHost = OperateIniFile.ReadIniData("Mail", "Mail Host", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailPort = Convert.ToInt32(OperateIniFile.ReadIniData("Mail", "Mail Port", "25", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                reportSettingModel.MailUsername = OperateIniFile.ReadIniData("Mail", "Mail Username", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailPwd = OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                this.tabSetting.DataContext = reportSettingModel;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load the report setting. Exception: " + ex.Message);
            }
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
                    this.reportSettingModel.MailPwd = txtMailPwd.Password;
                }
                OperateIniFile.WriteIniData("Report", "Doctor Names List", string.Join(";", this.reportSettingModel.DoctorNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Technician Names List", string.Join(";", this.reportSettingModel.TechNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Print Paper", this.reportSettingModel.PrintPaper, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "My Mail Address", this.reportSettingModel.MailAddress, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "To Mail Address", this.reportSettingModel.ToMailAddress, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Subject", this.reportSettingModel.MailSubject, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Host", this.reportSettingModel.MailHost, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Port", this.reportSettingModel.MailPort.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Username", this.reportSettingModel.MailUsername, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail Password", this.reportSettingModel.MailPwd, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Mail", "Mail SSL", this.reportSettingModel.MailSsl.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                MessageBox.Show("Saved the report setting successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save the report setting. Exception: "+ex.Message);
            }
        }

        private void btnReportSettingClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Owner.Visibility = Visibility.Visible;
        }

        private void tabMail_GotFocus(object sender, RoutedEventArgs e)
        {
            labSettingTitle.Content = "Default Mail Setting";
            labSettingDesc.Content = "There are some initial mail parameters, you can change that at any time.";
        }

        private void tabReport_GotFocus(object sender, RoutedEventArgs e)
        {
            labSettingTitle.Content = "Default Report Setting";
            labSettingDesc.Content = "There are some initial report parameters, you can change that at any time.";
        }

        private void btnAddDoctor_Click(object sender, RoutedEventArgs e)
        {
            AddNamePage addNamePage = new AddNamePage(this.reportSettingModel.DoctorNames);            
            addNamePage.ShowDialog();
        }

        private void btnDelDoctor_Click(object sender, RoutedEventArgs e)
        {
            if (this.listDoctorName.SelectedIndex != -1)
            {
                this.reportSettingModel.DoctorNames.RemoveAt(this.listDoctorName.SelectedIndex);
            }
        }

        private void btnAddTech_Click(object sender, RoutedEventArgs e)
        {
            AddNamePage addNamePage = new AddNamePage(this.reportSettingModel.TechNames);            
            addNamePage.ShowDialog();

        }

        private void btnDelTech_Click(object sender, RoutedEventArgs e)
        {
            if (this.listTechnicianName.SelectedIndex != -1)
            {
                this.reportSettingModel.TechNames.RemoveAt(this.listTechnicianName.SelectedIndex);
            }
        }       
                      
    }
}

using MEIKReport.Common;
using MEIKReport.Model;
using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
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
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);
        private delegate void ProgressBarGridDelegate(DependencyProperty dp, Object value); 
        public ReportSettingPage()
        {
            InitializeComponent();
            //LoadInitConfig();
            //if ("Letter".Equals(App.reportSettingModel.PrintPaper, StringComparison.OrdinalIgnoreCase))
            //{
            //    radLetter.IsChecked = true;
            //    radA4.IsChecked = false;
            //}
            //else
            //{
            //    radLetter.IsChecked = false;
            //    radA4.IsChecked = true;
            //}
            //if (App.reportSettingModel.DeviceType == 1)
            //{                
            //    optDoctor.Visibility = Visibility.Collapsed;                  
            //    optTech.Visibility = Visibility.Visible;
            //}
            //else if (App.reportSettingModel.DeviceType == 2)
            //{
            //    optDoctor.Visibility = Visibility.Visible;
            //    optTech.Visibility = Visibility.Collapsed;
            //}
            //else
            //{
            //    grpUserProfile.Visibility = Visibility.Collapsed;
            //}
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
                if (!string.IsNullOrEmpty(txtFTPPwd.Password))
                {
                    App.reportSettingModel.FtpPwd = txtFTPPwd.Password;
                }

                if (!string.IsNullOrEmpty(txtMailPwd.Password))
                {
                    App.reportSettingModel.MailPwd = txtMailPwd.Password;
                }
                OperateIniFile.WriteIniData("Base", "MEIK base", App.reportSettingModel.MeikBase, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Use Default Signature", App.reportSettingModel.UseDefaultSignature.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                User[] doctorUsers=App.reportSettingModel.DoctorNames.ToArray<User>();
                List<string> doctorsArr = new List<string>();
                foreach (var item in doctorUsers)
                {
                    doctorsArr.Add(item.Name + "|" + item.License);
                }
                OperateIniFile.WriteIniData("Report", "Doctor Names List", string.Join(";", doctorsArr.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");

                User[] techUsers = App.reportSettingModel.TechNames.ToArray<User>();
                List<string> techArr = new List<string>();
                foreach (var item in techUsers)
                {
                    techArr.Add(item.Name + "|" + item.License);
                }
                OperateIniFile.WriteIniData("Report", "Technician Names List", string.Join(";", techArr.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                //OperateIniFile.WriteIniData("Report", "Doctor Names List", string.Join(";", App.reportSettingModel.DoctorNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                //OperateIniFile.WriteIniData("Report", "Technician Names List", string.Join(";", App.reportSettingModel.TechNames.ToArray()), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");

                OperateIniFile.WriteIniData("Report", "Hide Doctor Signature", App.reportSettingModel.NoShowDoctorSignature.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("Report", "Hide Technician Signature", App.reportSettingModel.NoShowTechSignature.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");

                OperateIniFile.WriteIniData("FTP", "FTP Path", App.reportSettingModel.FtpPath, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                OperateIniFile.WriteIniData("FTP", "FTP User", App.reportSettingModel.FtpUser, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                string ftpPwd = App.reportSettingModel.FtpPwd;
                if (!string.IsNullOrEmpty(ftpPwd))
                {
                    ftpPwd = SecurityTools.EncryptText(ftpPwd);
                }
                OperateIniFile.WriteIniData("FTP", "FTP Password", ftpPwd, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                
                OperateIniFile.WriteIniData("Report", "Print Paper", App.reportSettingModel.PrintPaper.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
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
                MessageBox.Show(this, App.Current.FindResource("Message_14").ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_13").ToString() + " " + ex.Message);
            }
        }

        private void btnReportSettingClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Owner.Visibility = Visibility.Visible;
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

        //private void LoadInitConfig()
        //{
        //    //加载ini文件内容
        //    try
        //    {
        //        if (App.reportSettingModel == null)
        //        {
        //            App.reportSettingModel = new ReportSettingModel();
        //            App.reportSettingModel.MeikBase = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\MEIKData", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.Version = OperateIniFile.ReadIniData("Base", "Version", "1.0.0", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.HasNewer = Convert.ToBoolean(OperateIniFile.ReadIniData("Base", "Newer", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            string doctorNames = OperateIniFile.ReadIniData("Report", "Doctor Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");                    
        //            if (!string.IsNullOrEmpty(doctorNames))
        //            {                        
        //                var doctorList = doctorNames.Split(';').ToList<string>();
        //                //doctorList.ForEach(item => App.reportSettingModel.DoctorNames.Add(item));
        //                foreach (var item in doctorList)
        //                {
        //                    User doctorUser=new User();
        //                    string[] arr = item.Split('|');
        //                    doctorUser.Name = arr[0];
        //                    doctorUser.License = arr[1];
        //                    App.reportSettingModel.DoctorNames.Add(doctorUser);
        //                }                        
        //            }
        //            string techNames = OperateIniFile.ReadIniData("Report", "Technician Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");                    
        //            if (!string.IsNullOrEmpty(techNames))
        //            {
        //                var techList = techNames.Split(';').ToList<string>();
        //                //techList.ForEach(item => App.reportSettingModel.TechNames.Add(item));
        //                foreach (var item in techList)
        //                {
        //                    User techUser = new User();
        //                    string[] arr = item.Split('|');
        //                    techUser.Name = arr[0];
        //                    techUser.License = arr[1];
        //                    App.reportSettingModel.TechNames.Add(techUser);
        //                }
        //            }

        //            App.reportSettingModel.NoShowDoctorSignature = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Hide Doctor Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            App.reportSettingModel.NoShowTechSignature = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Hide Technician Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            App.reportSettingModel.FtpPath = OperateIniFile.ReadIniData("FTP", "FTP Path", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.FtpUser = OperateIniFile.ReadIniData("FTP", "FTP User", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            string ftpPwd = OperateIniFile.ReadIniData("FTP", "FTP Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            if (!string.IsNullOrEmpty(ftpPwd))
        //            {
        //                App.reportSettingModel.FtpPwd = SecurityTools.DecryptText(ftpPwd);
        //            }
                    
        //            App.reportSettingModel.PrintPaper = (PageSize)Enum.Parse(typeof(PageSize),OperateIniFile.ReadIniData("Report", "Print Paper", "Letter", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"),true);
        //            App.reportSettingModel.MailAddress = OperateIniFile.ReadIniData("Mail", "My Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.ToMailAddress = OperateIniFile.ReadIniData("Mail", "To Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.MailSubject = OperateIniFile.ReadIniData("Mail", "Mail Subject", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.MailBody = OperateIniFile.ReadIniData("Mail", "Mail Content", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.MailHost = OperateIniFile.ReadIniData("Mail", "Mail Host", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.MailPort = Convert.ToInt32(OperateIniFile.ReadIniData("Mail", "Mail Port", "25", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            App.reportSettingModel.MailUsername = OperateIniFile.ReadIniData("Mail", "Mail Username", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            string mailPwd=OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            if(!string.IsNullOrEmpty(mailPwd)){
        //                App.reportSettingModel.MailPwd = SecurityTools.DecryptText(mailPwd);
        //            }                    
        //            App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //            App.reportSettingModel.DeviceNo = OperateIniFile.ReadIniData("Device", "Device No", "000", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //            App.reportSettingModel.DeviceType = Convert.ToInt32(OperateIniFile.ReadIniData("Device", "Device Type", "1", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
        //        }                
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(this, App.Current.FindResource("Message_9").ToString() + " " + ex.Message);
        //    }
                        
        //}
        

        /// <summary>
        /// 立即进行版本更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdateNow_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_46").ToString(), "Update Now", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                //定义委托代理
                ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
                //使用系统代理方式显示进度条面板
                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(5) });
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
                
                try
                {
                    string ftpPath = App.reportSettingModel.FtpPath.Replace("home", "MeikUpdate");
                    //查询FTP上所有版本文件列表要下载的文件列表
                    var fileList = FtpHelper.Instance.GetFileList(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, ftpPath);                    
                    fileList.Reverse();
                    foreach (var setupFileName in fileList)
                    {
                        if (setupFileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            var verStr = setupFileName.ToLower().Replace(".exe", "").Replace("meiksetup.", "");
                            string currentVer=App.reportSettingModel.Version;
                            if (string.Compare(verStr, currentVer,StringComparison.OrdinalIgnoreCase) > 0)
                            {
                                //下載新的升級安裝包
                                FtpWebRequest reqFTP;
                                try
                                {
                                    FileStream outputStream = new FileStream(App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + setupFileName, FileMode.Create);
                                    reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpPath + setupFileName));
                                    reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;
                                    reqFTP.UseBinary = true;
                                    reqFTP.Credentials = new NetworkCredential(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd);
                                    reqFTP.UsePassive = false;
                                    FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                                    Stream ftpStream = response.GetResponseStream();
                                    long cl = response.ContentLength;
                                    int bufferSize = 2048;
                                    double moveSize = 100d / (cl / 2048);
                                    int x = 1;
                                    int readCount;
                                    byte[] buffer = new byte[bufferSize];
                                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                                    while (readCount > 0)
                                    {
                                        outputStream.Write(buffer, 0, readCount);
                                        readCount = ftpStream.Read(buffer, 0, bufferSize);
                                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(moveSize * x) });
                                        x++;
                                    }
                                    ftpStream.Close();
                                    outputStream.Close();
                                    response.Close();
                                }
                                catch (Exception ex2)
                                {
                                    throw ex2;
                                }

                                //FtpHelper.Instance.Download(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, ftpPath, App.reportSettingModel.DataBaseFolder, setupFileName);
                                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                                MessageBox.Show(this, App.Current.FindResource("Message_48").ToString());

                                try
                                {
                                    //启动外部程序
                                    Process setupProc = Process.Start(App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + setupFileName);
                                    if (setupProc != null)
                                    {                                       
                                        //proc.WaitForExit();//等待外部程序退出后才能往下执行
                                        setupProc.WaitForInputIdle();                                        
                                        File.Copy(System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini", System.AppDomain.CurrentDomain.BaseDirectory + "Config_bak.ini", true);                                        
                                        OperateIniFile.WriteIniData("Base", "Version", verStr, System.AppDomain.CurrentDomain.BaseDirectory + "Config_bak.ini");                                                                       
                                        //关闭当前程序
                                        App.Current.Shutdown();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(App.Current.FindResource("Message_51").ToString() + " " + ex.Message);
                                }


                            }
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_47").ToString() + " " + ex.Message);
                }
                finally
                {
                    Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                }
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var radioButton = sender as RadioButton;
            if (radioButton!=null&&!radioButton.DataContext.Equals(App.reportSettingModel.DeviceType.ToString()))
            {
                int targetDeviceType = Convert.ToInt32(radioButton.DataContext);
                int deviceType = App.reportSettingModel.DeviceType;
                if (radioButton.DataContext.Equals("1"))
                {
                    PasswordPage passwordPage = new PasswordPage(1);
                    passwordPage.ShowDialog();                    
                }
                else if(radioButton.DataContext.Equals("2")){
                    PasswordPage passwordPage = new PasswordPage(2);
                    passwordPage.ShowDialog();
                }
                else{
                    PasswordPage passwordPage = new PasswordPage(3);
                    passwordPage.ShowDialog();
                }
                if (targetDeviceType != App.reportSettingModel.DeviceType)
                {

                    if (this.radradDeviceType1.DataContext.Equals(deviceType.ToString()))
                    {
                        this.radradDeviceType1.IsChecked = true;
                    }
                    else if (this.radradDeviceType2.DataContext.Equals(deviceType.ToString()))
                    {
                        this.radradDeviceType2.IsChecked = true;
                    }
                    else
                    {
                        this.radradDeviceType3.IsChecked = true;
                    }

                }
                else
                {
                    var userList = this.Owner as UserList;
                    if (App.reportSettingModel.DeviceType==1)
                    {
                        userList.btnScreening.Visibility = Visibility.Visible;
                        userList.btnDiagnostics.Visibility = Visibility.Visible;
                        userList.btnRecords.Visibility = Visibility.Visible;
                        userList.btnReceivePdf.Visibility = Visibility.Visible;
                        userList.btnReceive.Visibility = Visibility.Collapsed;
                        userList.btnNewArchive.Visibility = Visibility.Visible;
                        userList.btnExaminationReport.Visibility = Visibility.Collapsed;
                        userList.sendDataButton.Visibility = Visibility.Visible;
                        userList.sendReportButton.Visibility = Visibility.Collapsed;
                    }
                    else if (App.reportSettingModel.DeviceType == 2)
                    {
                        userList.btnScreening.Visibility = Visibility.Collapsed;
                        userList.btnDiagnostics.Visibility = Visibility.Collapsed;
                        userList.btnRecords.Visibility = Visibility.Collapsed;
                        userList.btnReceive.Visibility = Visibility.Visible;
                        userList.btnReceivePdf.Visibility = Visibility.Collapsed;
                        userList.btnNewArchive.Visibility = Visibility.Collapsed;
                        userList.btnExaminationReport.Visibility = Visibility.Visible;
                        userList.sendDataButton.Visibility = Visibility.Collapsed;
                        userList.sendReportButton.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        userList.btnScreening.Visibility = Visibility.Visible;
                        userList.btnDiagnostics.Visibility = Visibility.Visible;
                        userList.btnRecords.Visibility = Visibility.Visible;
                        userList.btnReceive.Visibility = Visibility.Visible;
                        userList.btnReceivePdf.Visibility = Visibility.Visible;
                        userList.btnNewArchive.Visibility = Visibility.Visible;
                        userList.btnExaminationReport.Visibility = Visibility.Visible;
                        userList.sendDataButton.Visibility = Visibility.Visible;
                        userList.sendReportButton.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void tabSetting_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems!=null&&e.AddedItems.Count > 0)
            {
                var tabItem=e.AddedItems[0] as TabItem;
                if (tabItem!=null&&"tabVersion".Equals(tabItem.Name))
                {
                    if (this.radradDeviceType1.DataContext.Equals(App.reportSettingModel.DeviceType.ToString()))
                    {
                        this.radradDeviceType1.IsChecked = true;
                    }
                    else if (this.radradDeviceType2.DataContext.Equals(App.reportSettingModel.DeviceType.ToString()))
                    {
                        this.radradDeviceType2.IsChecked = true;
                    }
                    else
                    {
                        this.radradDeviceType3.IsChecked = true;
                    }
            
                    //定义委托代理
                    ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
                    UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
                    //使用系统代理方式显示进度条面板
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(10) });
                    Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(20) });
                    List<string> fileList = new List<string>();
                    try
                    {
                        string ftpPath = App.reportSettingModel.FtpPath.Replace("home", "MeikUpdate");
                        //查询FTP上所有版本文件列表要下载的文件列表
                        fileList = FtpHelper.Instance.GetFileList(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, ftpPath);
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(90) });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_47").ToString() + " " + ex.Message);
                        Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                        return;
                    }
                    try
                    {
                        fileList.Reverse();
                        foreach (var setupFileName in fileList)
                        {
                            if (setupFileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                            {
                                var verStr = setupFileName.ToLower().Replace(".exe", "").Replace("meiksetup.", "");
                                string currentVer = App.reportSettingModel.Version;
                                if (string.Compare(verStr, currentVer,StringComparison.OrdinalIgnoreCase) > 0)
                                {                                                                
                                    labVerCheckInfo.Content = App.Current.FindResource("SettingVersionCheckInfo2").ToString();
                                    btnUpdateNow.IsEnabled = true;                            
                                }
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        MessageBox.Show(this, ex1.Message);
                    }
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                    Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
            
                }
            }
        }
    }
}

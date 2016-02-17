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
                                Person person = new Person();
                                person.Code = NextFile.Name.Substring(0, NextFile.Name.Length - 4);
                                //Personal Data
                                person.SurName = OperateIniFile.ReadIniData("Personal data", "surname", "", NextFile.FullName);
                                person.GivenName = OperateIniFile.ReadIniData("Personal data", "given name", "", NextFile.FullName);
                                person.OtherName = OperateIniFile.ReadIniData("Personal data", "other name", "", NextFile.FullName);
                                person.Address = OperateIniFile.ReadIniData("Personal data", "address", "", NextFile.FullName);
                                person.BirthDate = OperateIniFile.ReadIniData("Personal data", "birth date", "", NextFile.FullName);
                                person.BirthMonth = OperateIniFile.ReadIniData("Personal data", "birth month", "", NextFile.FullName);
                                person.BirthYear = OperateIniFile.ReadIniData("Personal data", "birth year", "", NextFile.FullName);
                                person.RegDate = OperateIniFile.ReadIniData("Personal data", "registration date", "", NextFile.FullName);
                                person.RegMonth = OperateIniFile.ReadIniData("Personal data", "registration month", "", NextFile.FullName);
                                person.RegYear = OperateIniFile.ReadIniData("Personal data", "registration year", "", NextFile.FullName);
                                                                                                
                                person.Birthday = person.BirthMonth + "/" + person.BirthDate + "/" + person.BirthYear;
                                //person.Regdate = registrationmonth + "/" + registrationdate + "/" + registrationyear;
                                person.ArchiveFolder = folderName;
                                if (!string.IsNullOrEmpty(person.Birthday))
                                {
                                    int m_Y1 = DateTime.Parse(person.Birthday).Year;
                                    int m_Y2 = DateTime.Now.Year;
                                    person.Age = m_Y2 - m_Y1;
                                }

                                //Complaints
                                person.Pain =Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "pain", "0", NextFile.FullName)));
                                person.Colostrum = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "colostrum", "0", NextFile.FullName)));
                                person.SerousDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "serous discharge", "0", NextFile.FullName)));
                                person.BloodDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "blood discharge", "0", NextFile.FullName)));
                                person.Other = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "other", "0", NextFile.FullName)));
                                person.Pregnancy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "pregnancy", "0", NextFile.FullName)));
                                person.Lactation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "lactation", "0", NextFile.FullName)));
                                person.OtherDesc = OperateIniFile.ReadIniData("Complaints", "other description", "", NextFile.FullName);
                                person.PregnancyTerm = OperateIniFile.ReadIniData("Complaints", "pregnancy term", "", NextFile.FullName);
                                person.OtherDesc = person.OtherDesc.Replace(";;;;", "\r\n");


                                person.MenstrualCycleDisorder = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "menstrual cycle disorder", "0", NextFile.FullName)));
                                person.Postmenopause = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "postmenopause", "0", NextFile.FullName)));
                                person.HormonalContraceptives = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "hormonal contraceptives", "0", NextFile.FullName)));
                                person.MenstrualCycleDisorderDesc = OperateIniFile.ReadIniData("Menses", "menstrual cycle disorder description", "", NextFile.FullName);
                                person.PostmenopauseDesc = OperateIniFile.ReadIniData("Menses", "postmenopause description", "", NextFile.FullName);
                                person.HormonalContraceptivesBrandName = OperateIniFile.ReadIniData("Menses", "hormonal contraceptives brand name", "", NextFile.FullName);
                                person.HormonalContraceptivesPeriod = OperateIniFile.ReadIniData("Menses", "hormonal contraceptives period", "", NextFile.FullName);


                                person.Adiposity = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "adiposity", "0", NextFile.FullName)));
                                person.EssentialHypertension = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "essential hypertension", "0", NextFile.FullName)));
                                person.Diabetes = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "diabetes", "0", NextFile.FullName)));
                                person.ThyroidGlandDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "thyroid gland diseases", "0", NextFile.FullName)));
                                person.SomaticOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "other", "0", NextFile.FullName)));
                                person.EssentialHypertensionDesc = OperateIniFile.ReadIniData("Somatic", "essential hypertension description", "", NextFile.FullName);
                                person.DiabetesDesc = OperateIniFile.ReadIniData("Somatic", "diabetes description", "", NextFile.FullName);
                                person.ThyroidGlandDiseasesDesc = OperateIniFile.ReadIniData("Somatic", "thyroid gland diseases description", "", NextFile.FullName);
                                person.SomaticOtherDesc = OperateIniFile.ReadIniData("Somatic", "other description", "", NextFile.FullName);


                                person.Infertility = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "infertility", "0", NextFile.FullName)));
                                person.OvaryDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "ovary diseases", "0", NextFile.FullName)));
                                person.OvaryCyst = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "ovary cyst", "0", NextFile.FullName)));
                                person.OvaryCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "ovary cancer", "0", NextFile.FullName)));
                                person.OvaryEndometriosis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "ovary endometriosis", "0", NextFile.FullName)));
                                person.OvaryOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "ovary other", "0", NextFile.FullName)));
                                person.UterusDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "uterus diseases", "0", NextFile.FullName)));
                                person.UterusMyoma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "uterus myoma", "0", NextFile.FullName)));
                                person.UterusCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "uterus cancer", "0", NextFile.FullName)));
                                person.UterusEndometriosis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "uterus endometriosis", "0", NextFile.FullName)));
                                person.UterusOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "uterus other", "0", NextFile.FullName)));
                                person.GynecologicOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Gynecologic", "other", "0", NextFile.FullName)));
                                person.InfertilityDesc = OperateIniFile.ReadIniData("Gynecologic", "infertility-description", "", NextFile.FullName);
                                person.OvaryOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "ovary other description", "", NextFile.FullName);
                                person.UterusOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "uterus other description", "", NextFile.FullName);
                                person.GynecologicOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "other description", "", NextFile.FullName);


                                person.Abortions = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Obstetric", "abortions", "0", NextFile.FullName)));
                                person.Births = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Obstetric", "births", "0", NextFile.FullName)));
                                person.AbortionsNumber = OperateIniFile.ReadIniData("Obstetric", "abortions number", "", NextFile.FullName);
                                person.BirthsNumber = OperateIniFile.ReadIniData("Obstetric", "births number", "", NextFile.FullName);


                                person.Infertility = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation till 1 month", "0", NextFile.FullName)));
                                person.OvaryDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation till 1 year", "0", NextFile.FullName)));
                                person.OvaryCyst = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation over 1 year", "0", NextFile.FullName)));


                                person.Trauma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "trauma", "0", NextFile.FullName)));
                                person.Mastitis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "mastitis", "0", NextFile.FullName)));
                                person.FibrousCysticMastopathy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "fibrous- cystic mastopathy", "0", NextFile.FullName)));
                                person.Cysts = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "cysts", "0", NextFile.FullName)));
                                person.Cancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "cancer", "0", NextFile.FullName)));
                                person.DiseasesOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "other", "0", NextFile.FullName)));
                                person.TraumaDesc = OperateIniFile.ReadIniData("Diseases", "trauma description", "", NextFile.FullName);
                                person.MastitisDesc = OperateIniFile.ReadIniData("Diseases", "mastitis description", "", NextFile.FullName);
                                person.FibrousCysticMastopathyDesc = OperateIniFile.ReadIniData("Diseases", "fibrous- cystic mastopathy description", "", NextFile.FullName);
                                person.CystsDesc = OperateIniFile.ReadIniData("Diseases", "cysts descriptuin", "", NextFile.FullName);
                                person.CancerDesc = OperateIniFile.ReadIniData("Diseases", "cancer description", "", NextFile.FullName);
                                person.DiseasesOtherDesc = OperateIniFile.ReadIniData("Diseases", "other description", "", NextFile.FullName);


                                person.PalpationDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Palpation", "diffuse", "0", NextFile.FullName)));
                                person.PalpationFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Palpation", "focal", "0", NextFile.FullName)));
                                person.PalpationDesc = OperateIniFile.ReadIniData("Palpation", "description", "", NextFile.FullName);

                                person.UltrasoundDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Ultrasound", "diffuse", "0", NextFile.FullName)));
                                person.UltrasoundFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Ultrasound", "focal", "0", NextFile.FullName)));
                                person.UltrasoundnDesc = OperateIniFile.ReadIniData("Ultrasound", "description", "", NextFile.FullName);

                                person.MammographyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Mammography", "diffuse", "0", NextFile.FullName)));
                                person.MammographyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Mammography", "focal", "0", NextFile.FullName)));
                                person.MammographyDesc = OperateIniFile.ReadIniData("Mammography", "description", "", NextFile.FullName);

                                person.BiopsyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "diffuse", "0", NextFile.FullName)));
                                person.BiopsyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "focal", "0", NextFile.FullName)));
                                person.BiopsyCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "cancer", "0", NextFile.FullName)));
                                person.BiopsyProliferation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "proliferation", "0", NextFile.FullName)));
                                person.BiopsyDysplasia = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "dysplasia", "0", NextFile.FullName)));
                                person.BiopsyIntraductalPapilloma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "intraductal papilloma", "0", NextFile.FullName)));
                                person.BiopsyFibroadenoma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "fibroadenoma", "0", NextFile.FullName)));
                                person.BiopsyOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "other", "0", NextFile.FullName)));
                                person.BiopsyOtherDesc = OperateIniFile.ReadIniData("Biopsy", "other description", "", NextFile.FullName);
                                
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
            if (string.IsNullOrEmpty(App.reportSettingModel.MailAddress) || string.IsNullOrEmpty(App.reportSettingModel.MailHost)
                || string.IsNullOrEmpty(App.reportSettingModel.MailUsername) || string.IsNullOrEmpty(App.reportSettingModel.MailPwd)
                || string.IsNullOrEmpty(App.reportSettingModel.ToMailAddress))
            {
                MessageBox.Show("You don't configure the mail account or configured the mail account incorrectly.");
                return;
            }
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
            if (!File.Exists(dataFile))
            {
                MessageBox.Show("The user's report has not been created, please fill and save the report first.");
                return;
            }
                        
            string toMail = App.reportSettingModel.ToMailAddress;
            MessageBoxResult result = MessageBox.Show("Are you sure to email all of the reports to " + toMail + "?", "Email Reports", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    
                    string senderServerIp = App.reportSettingModel.MailHost;
                    string toMailAddress = toMail;
                    string fromMailAddress = App.reportSettingModel.MailAddress;
                    string subjectInfo = App.reportSettingModel.MailSubject + " (" + selectedUser.SurName + ")";
                    string bodyInfo = App.reportSettingModel.MailBody;
                    string mailUsername = App.reportSettingModel.MailUsername;
                    string mailPassword = App.reportSettingModel.MailPwd;
                    string mailPort = App.reportSettingModel.MailPort + "";
                    string lfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_LF.pdf";
                    string sfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_SF.pdf";                    
                    string attachPath = dataFile + ";" + lfPdfFile + ";" + sfPdfFile;
                    bool isSsl = App.reportSettingModel.MailSsl;
                    EmailHelper email = new EmailHelper(senderServerIp, toMailAddress, fromMailAddress, subjectInfo, bodyInfo, mailUsername, mailPassword, mailPort, isSsl, false);
                    email.AddAttachments(attachPath);
                    email.Send();
                    MessageBox.Show("Sent the emial successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to send email, Exception: "+ex.Message);
                }
            }
        }

        private void LoadInitConfig()
        {
            try
            {
                if (App.reportSettingModel == null)
                {
                    App.reportSettingModel = new ReportSettingModel();
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
                    App.reportSettingModel.MailPwd = OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load the report setting. Exception: " + ex.Message);
            }
        }
    }
}

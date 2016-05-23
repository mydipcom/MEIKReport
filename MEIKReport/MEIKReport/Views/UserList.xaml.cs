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
        private string deviceNo = OperateIniFile.ReadIniData("Device", "Device No", "000", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        private string meikFolder = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");        
        protected MouseHook mouseHook = new MouseHook();
        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "Data";
        private IList<Patient> patientList = new List<Patient>();
        public UserList()
        {
            InitializeComponent();            
            listLang.SelectedIndex = App.local.Equals("en-US") ? 0 : App.local.Equals("zh-HK") ? 1 : 1;
            if (!string.IsNullOrEmpty(App.dataFolder))
            {
                loadArchiveFolder(App.dataFolder);                
                string meikiniFile = meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                OperateIniFile.WriteIniData("Base", "Patients base", App.dataFolder, meikiniFile);
            }
            
            LoadInitConfig();
            labDeviceNo.Content = App.reportSettingModel.DeviceNo;

            if (App.reportSettingModel.DeviceType == 1 || App.reportSettingModel.DeviceType == 3)
            {
                btnRecords.Visibility = Visibility.Visible;
            }
            //if (App.reportSettingModel.DeviceType == 3)
            //{
            //    btnSummaryReport.Visibility = Visibility.Visible;
            //}
            mouseHook.MouseUp += new System.Windows.Forms.MouseEventHandler(mouseHook_MouseUp);            
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
                MessageBox.Show(this, ex.Message);
            }
            finally
            {
                //this.Owner.Visibility = Visibility.Visible;                
                var owner = this.Owner as MainWindow;
                owner.exitMeik();
            }
        }

        private void exitReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            //this.Owner.Show();
            var owner = this.Owner as MainWindow;
            owner.exitMeik();
        }

        private void btnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult res = folderBrowserDialog.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                string folderName = folderBrowserDialog.SelectedPath;
                App.dataFolder = folderName;
                loadArchiveFolder(folderName); 
            }        
        }


        public void loadArchiveFolder(string folderName)
        {
            try
            {                
                txtFolderPath.Text = folderName;
                CollectionViewSource customerSource = (CollectionViewSource)this.FindResource("CustomerSource");
                HashSet<Person> set = new HashSet<Person>();
                //遍历指定文件夹下所有文件
                HandleFolder(folderName,ref set);

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

                string meikiniFile = meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                var selectItem = this.CodeListBox.SelectedItem as Person;
                if (selectItem != null)
                {                                        
                    //修改原始MEIK程序中的患者档案目录，让原始MEIK程序运行后直接打开此患者档案
                    OperateIniFile.WriteIniData("Base", "Patients base", selectItem.ArchiveFolder, meikiniFile);
                }
                else
                {
                    //修改原始MEIK程序中的患者档案目录，让原始MEIK程序运行后直接打开此患者档案
                    OperateIniFile.WriteIniData("Base", "Patients base", folderName, meikiniFile);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void HandleFolder(string folderName,ref HashSet<Person> set)
        {
            //遍历指定文件夹下所有文件
            DirectoryInfo theFolder = new DirectoryInfo(folderName);             
            try
            {
                FileInfo[] fileInfo = theFolder.GetFiles();
                //遍历文件
                foreach (FileInfo NextFile in fileInfo)
                {
                    if (".crd".Equals(NextFile.Extension, StringComparison.OrdinalIgnoreCase))
                    {
                        Person person = new Person();
                        //person.ArchiveFolder = folderName;
                        person.ArchiveFolder = theFolder.FullName;
                        person.CrdFilePath = NextFile.FullName;

                        person.Code = NextFile.Name.Substring(0, NextFile.Name.Length - 4);

                        person.TechName = OperateIniFile.ReadIniData("Report", "Technician Name", "", NextFile.FullName);
                        person.TechLicense=OperateIniFile.ReadIniData("Report", "Technician License", "", NextFile.FullName);                        

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
                        
                        if (!string.IsNullOrEmpty(person.Birthday))
                        {
                            int m_Y1 = DateTime.Parse(person.Birthday).Year;
                            int m_Y2 = DateTime.Now.Year;
                            person.Age = m_Y2 - m_Y1;
                        }

                        //Complaints
                        person.Pain = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "pain", "0", NextFile.FullName)));
                        person.Colostrum = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "colostrum", "0", NextFile.FullName)));
                        person.SerousDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "serous discharge", "0", NextFile.FullName)));
                        person.BloodDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "blood discharge", "0", NextFile.FullName)));
                        person.Other = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "other", "0", NextFile.FullName)));
                        person.Pregnancy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "pregnancy", "0", NextFile.FullName)));
                        person.Lactation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Complaints", "lactation", "0", NextFile.FullName)));
                        person.OtherDesc = OperateIniFile.ReadIniData("Complaints", "other description", "", NextFile.FullName);
                        person.PregnancyTerm = OperateIniFile.ReadIniData("Complaints", "pregnancy term", "", NextFile.FullName);
                        person.OtherDesc = person.OtherDesc.Replace(";;", "\r\n");
                        person.PregnancyTerm = person.PregnancyTerm.Replace(";;", "\r\n");

                        person.MenstrualCycleDisorder = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "menstrual cycle disorder", "0", NextFile.FullName)));
                        person.Postmenopause = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "postmenopause", "0", NextFile.FullName)));
                        person.HormonalContraceptives = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Menses", "hormonal contraceptives", "0", NextFile.FullName)));
                        person.MenstrualCycleDisorderDesc = OperateIniFile.ReadIniData("Menses", "menstrual cycle disorder description", "", NextFile.FullName);
                        person.MenstrualCycleDisorderDesc = person.MenstrualCycleDisorderDesc.Replace(";;", "\r\n");
                        person.PostmenopauseDesc = OperateIniFile.ReadIniData("Menses", "postmenopause description", "", NextFile.FullName);
                        person.PostmenopauseDesc = person.PostmenopauseDesc.Replace(";;", "\r\n");
                        person.HormonalContraceptivesBrandName = OperateIniFile.ReadIniData("Menses", "hormonal contraceptives brand name", "", NextFile.FullName);
                        person.HormonalContraceptivesBrandName = person.HormonalContraceptivesBrandName.Replace(";;", "\r\n");
                        person.HormonalContraceptivesPeriod = OperateIniFile.ReadIniData("Menses", "hormonal contraceptives period", "", NextFile.FullName);
                        person.HormonalContraceptivesPeriod = person.HormonalContraceptivesPeriod.Replace(";;", "\r\n");


                        person.Adiposity = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "adiposity", "0", NextFile.FullName)));
                        person.EssentialHypertension = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "essential hypertension", "0", NextFile.FullName)));
                        person.Diabetes = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "diabetes", "0", NextFile.FullName)));
                        person.ThyroidGlandDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "thyroid gland diseases", "0", NextFile.FullName)));
                        person.SomaticOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Somatic", "other", "0", NextFile.FullName)));
                        person.EssentialHypertensionDesc = OperateIniFile.ReadIniData("Somatic", "essential hypertension description", "", NextFile.FullName);
                        person.EssentialHypertensionDesc = person.EssentialHypertensionDesc.Replace(";;", "\r\n");
                        person.DiabetesDesc = OperateIniFile.ReadIniData("Somatic", "diabetes description", "", NextFile.FullName);
                        person.DiabetesDesc = person.DiabetesDesc.Replace(";;", "\r\n");
                        person.ThyroidGlandDiseasesDesc = OperateIniFile.ReadIniData("Somatic", "thyroid gland diseases description", "", NextFile.FullName);
                        person.ThyroidGlandDiseasesDesc = person.ThyroidGlandDiseasesDesc.Replace(";;", "\r\n");
                        person.SomaticOtherDesc = OperateIniFile.ReadIniData("Somatic", "other description", "", NextFile.FullName);
                        person.SomaticOtherDesc = person.SomaticOtherDesc.Replace(";;", "\r\n");


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
                        person.InfertilityDesc = person.InfertilityDesc.Replace(";;", "\r\n");
                        person.OvaryOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "ovary other description", "", NextFile.FullName);
                        person.OvaryOtherDesc = person.OvaryOtherDesc.Replace(";;", "\r\n");
                        person.UterusOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "uterus other description", "", NextFile.FullName);
                        person.UterusOtherDesc = person.UterusOtherDesc.Replace(";;", "\r\n");
                        person.GynecologicOtherDesc = OperateIniFile.ReadIniData("Gynecologic", "other description", "", NextFile.FullName);
                        person.GynecologicOtherDesc = person.GynecologicOtherDesc.Replace(";;", "\r\n");


                        person.Abortions = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Obstetric", "abortions", "0", NextFile.FullName)));
                        person.Births = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Obstetric", "births", "0", NextFile.FullName)));
                        person.AbortionsNumber = OperateIniFile.ReadIniData("Obstetric", "abortions number", "", NextFile.FullName);
                        person.BirthsNumber = OperateIniFile.ReadIniData("Obstetric", "births number", "", NextFile.FullName);


                        person.LactationTill1Month = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation till 1 month", "0", NextFile.FullName)));
                        person.LactationTill1Year = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation till 1 year", "0", NextFile.FullName)));
                        person.LactationOver1Year = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Lactation", "lactation over 1 year", "0", NextFile.FullName)));


                        person.Trauma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "trauma", "0", NextFile.FullName)));
                        person.Mastitis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "mastitis", "0", NextFile.FullName)));
                        person.FibrousCysticMastopathy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "fibrous- cystic mastopathy", "0", NextFile.FullName)));
                        person.Cysts = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "cysts", "0", NextFile.FullName)));
                        person.Cancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "cancer", "0", NextFile.FullName)));
                        person.DiseasesOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Diseases", "other", "0", NextFile.FullName)));
                        person.TraumaDesc = OperateIniFile.ReadIniData("Diseases", "trauma description", "", NextFile.FullName);
                        person.TraumaDesc = person.TraumaDesc.Replace(";;", "\r\n");
                        person.MastitisDesc = OperateIniFile.ReadIniData("Diseases", "mastitis description", "", NextFile.FullName);
                        person.MastitisDesc = person.MastitisDesc.Replace(";;", "\r\n");
                        person.FibrousCysticMastopathyDesc = OperateIniFile.ReadIniData("Diseases", "fibrous- cystic mastopathy description", "", NextFile.FullName);
                        person.FibrousCysticMastopathyDesc = person.FibrousCysticMastopathyDesc.Replace(";;", "\r\n");
                        person.CystsDesc = OperateIniFile.ReadIniData("Diseases", "cysts descriptuin", "", NextFile.FullName);
                        person.CystsDesc = person.CystsDesc.Replace(";;", "\r\n");
                        person.CancerDesc = OperateIniFile.ReadIniData("Diseases", "cancer description", "", NextFile.FullName);
                        person.CancerDesc = person.CancerDesc.Replace(";;", "\r\n");
                        person.DiseasesOtherDesc = OperateIniFile.ReadIniData("Diseases", "other description", "", NextFile.FullName);
                        person.DiseasesOtherDesc = person.DiseasesOtherDesc.Replace(";;", "\r\n");


                        person.PalpationDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Palpation", "diffuse", "0", NextFile.FullName)));
                        person.PalpationFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Palpation", "focal", "0", NextFile.FullName)));
                        person.PalpationDesc = OperateIniFile.ReadIniData("Palpation", "description", "", NextFile.FullName);
                        person.PalpationDesc = person.PalpationDesc.Replace(";;", "\r\n");

                        person.UltrasoundDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Ultrasound", "diffuse", "0", NextFile.FullName)));
                        person.UltrasoundFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Ultrasound", "focal", "0", NextFile.FullName)));
                        person.UltrasoundnDesc = OperateIniFile.ReadIniData("Ultrasound", "description", "", NextFile.FullName);
                        person.UltrasoundnDesc = person.UltrasoundnDesc.Replace(";;", "\r\n");

                        person.MammographyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Mammography", "diffuse", "0", NextFile.FullName)));
                        person.MammographyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Mammography", "focal", "0", NextFile.FullName)));
                        person.MammographyDesc = OperateIniFile.ReadIniData("Mammography", "description", "", NextFile.FullName);
                        person.MammographyDesc = person.MammographyDesc.Replace(";;", "\r\n");

                        person.BiopsyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "diffuse", "0", NextFile.FullName)));
                        person.BiopsyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "focal", "0", NextFile.FullName)));
                        person.BiopsyCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "cancer", "0", NextFile.FullName)));
                        person.BiopsyProliferation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "proliferation", "0", NextFile.FullName)));
                        person.BiopsyDysplasia = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "dysplasia", "0", NextFile.FullName)));
                        person.BiopsyIntraductalPapilloma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "intraductal papilloma", "0", NextFile.FullName)));
                        person.BiopsyFibroadenoma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "fibroadenoma", "0", NextFile.FullName)));
                        person.BiopsyOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Biopsy", "other", "0", NextFile.FullName)));
                        person.BiopsyOtherDesc = OperateIniFile.ReadIniData("Biopsy", "other description", "", NextFile.FullName);
                        person.BiopsyOtherDesc = person.BiopsyOtherDesc.Replace(";;", "\r\n");

                        set.Add(person);
                    }
                }
            }
            catch (Exception) { }

            //Handle folder
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();            
            foreach (DirectoryInfo subFolder in dirInfo)
            {
                HandleFolder(subFolder.FullName, ref set);
            }
        }        

        private void EmailButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(App.reportSettingModel.MailAddress) || string.IsNullOrEmpty(App.reportSettingModel.MailHost)
                || string.IsNullOrEmpty(App.reportSettingModel.MailUsername) || string.IsNullOrEmpty(App.reportSettingModel.MailPwd)
                || string.IsNullOrEmpty(App.reportSettingModel.ToMailAddress))
            {
                MessageBox.Show(this, App.Current.FindResource("Message_15").ToString());
                return;
            }
            string deviceNo = App.reportSettingModel.DeviceNo;
            int deviceType = App.reportSettingModel.DeviceType;
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            //string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
            string dataFile=selectedUser.ArchiveFolder+ System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";

            if (!File.Exists(dataFile) && deviceType == 2)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_16").ToString());
                return;
            }
                        
            string toMail = App.reportSettingModel.ToMailAddress;
            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_17").ToString(), "Email Reports", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                //如果是Technician那么判断报告是否已经填写Techinican名称
                if (deviceType == 1)
                {
                    SelectTechnicianPage SelectTechnicianPage = new SelectTechnicianPage(App.reportSettingModel.TechNames, selectedUser);
                    SelectTechnicianPage.Owner=this;
                    SelectTechnicianPage.ShowDialog();
                    if (!string.IsNullOrEmpty(selectedUser.TechName))
                    {
                        try
                        {
                            OperateIniFile.WriteIniData("Report", "Technician Name", selectedUser.TechName, selectedUser.CrdFilePath);
                            OperateIniFile.WriteIniData("Report", "Technician License", selectedUser.TechLicense, selectedUser.CrdFilePath);
                        }
                        catch (Exception exec)
                        {
                            //如果不能写入ini文件
                            FileHelper.SetFolderPower(selectedUser.ArchiveFolder, "Everyone", "FullControl");
                            FileHelper.SetFolderPower(selectedUser.ArchiveFolder, "Users", "FullControl");
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                try
                {
                    string zipFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_" + deviceNo + ".zip";
                    ZipTools ZipTools = new ZipTools(selectedUser.ArchiveFolder);
                    try
                    {
                        ZipTools.ZipFolder(zipFile);
                    }
                    catch (Exception ex1)
                    {
                        MessageBox.Show(this, string.Format(App.Current.FindResource("Message_21").ToString() + " " + ex1.Message, selectedUser.ArchiveFolder));
                        return;
                    }
                    string senderServerIp = App.reportSettingModel.MailHost;
                    string toMailAddress = toMail;
                    string fromMailAddress = App.reportSettingModel.MailAddress;
                    string subjectInfo = App.reportSettingModel.MailSubject + " (" + selectedUser.SurName + ")";
                    string bodyInfo = App.reportSettingModel.MailBody;
                    string mailUsername = App.reportSettingModel.MailUsername;
                    string mailPassword = App.reportSettingModel.MailPwd;
                    string mailPort = App.reportSettingModel.MailPort + "";
                    //string lfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_LF.pdf";
                    //string sfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "_SF.pdf";                    
                    //string attachPath = dataFile + ";" + lfPdfFile + ";" + sfPdfFile;
                    string attachPath = zipFile;
                    bool isSsl = App.reportSettingModel.MailSsl;
                    EmailHelper email = new EmailHelper(senderServerIp, toMailAddress, fromMailAddress, subjectInfo, bodyInfo, mailUsername, mailPassword, mailPort, isSsl, false);
                    email.AddAttachments(attachPath);                    
                    email.Send();

                    //Send email with screening count records
                    if (deviceType == 1)
                    {
                        string excelFile = dataFolder + System.IO.Path.DirectorySeparatorChar + deviceNo + "_Count.xls";
                        try
                        {
                            if (File.Exists(excelFile))
                            {
                                try
                                {
                                    File.Delete(excelFile);
                                }
                                catch (Exception ex2) { }
                            }
                            exportExcel(excelFile);
                            subjectInfo = App.Current.FindResource("MailSubject").ToString() + " (" + deviceNo + ")";
                            bodyInfo = "";
                            EmailHelper emailCount = new EmailHelper(senderServerIp, toMailAddress, fromMailAddress, subjectInfo, bodyInfo, mailUsername, mailPassword, mailPort, isSsl, false);
                            emailCount.AddAttachments(excelFile);
                            emailCount.Send();
                        }
                        catch (Exception ex1) { }                        
                    }
                    MessageBox.Show(this, App.Current.FindResource("Message_18").ToString());                    
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_19").ToString() + " " + ex.Message);
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
                    App.reportSettingModel.UseDefaultSignature = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    string doctorNames = OperateIniFile.ReadIniData("Report", "Doctor Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(doctorNames))
                    {
                        var doctorList = doctorNames.Split(';').ToList<string>();
                        //doctorList.ForEach(item => App.reportSettingModel.DoctorNames.Add(item));
                        foreach (var item in doctorList)
                        {
                            User doctorUser = new User();
                            string[] arr = item.Split('|');
                            doctorUser.Name = arr[0];
                            doctorUser.License = arr[1];
                            App.reportSettingModel.DoctorNames.Add(doctorUser);
                        }
                    }
                    string techNames = OperateIniFile.ReadIniData("Report", "Technician Names List", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(techNames))
                    {
                        var techList = techNames.Split(';').ToList<string>();
                        //techList.ForEach(item => App.reportSettingModel.TechNames.Add(item));
                        foreach (var item in techList)
                        {
                            User techUser = new User();
                            string[] arr = item.Split('|');
                            techUser.Name = arr[0];
                            techUser.License = arr[1];
                            App.reportSettingModel.TechNames.Add(techUser);
                        }
                    }
                    App.reportSettingModel.PrintPaper = OperateIniFile.ReadIniData("Report", "Print Paper", "Letter", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailAddress = OperateIniFile.ReadIniData("Mail", "My Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.ToMailAddress = OperateIniFile.ReadIniData("Mail", "To Mail Address", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailSubject = OperateIniFile.ReadIniData("Mail", "Mail Subject", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailBody = OperateIniFile.ReadIniData("Mail", "Mail Content", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailHost = OperateIniFile.ReadIniData("Mail", "Mail Host", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.MailPort = Convert.ToInt32(OperateIniFile.ReadIniData("Mail", "Mail Port", "25", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.MailUsername = OperateIniFile.ReadIniData("Mail", "Mail Username", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    string mailPwd = OperateIniFile.ReadIniData("Mail", "Mail Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(mailPwd))
                    {
                        App.reportSettingModel.MailPwd = SecurityTools.DecryptText(mailPwd);
                    }   
                    App.reportSettingModel.MailSsl = Convert.ToBoolean(OperateIniFile.ReadIniData("Mail", "Mail SSL", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.DeviceNo = OperateIniFile.ReadIniData("Device", "Device No", "000", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.DeviceType = Convert.ToInt32(OperateIniFile.ReadIniData("Device", "Device Type", "1", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_20").ToString() + " " + ex.Message);
            }
        }
        
        /// <summary>
        /// Changed the icon for MRN list item.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CodeListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                var selectItem = (Person)e.AddedItems[0];
                selectItem.Icon = "/Images/id_card_ok.png";                
                string meikiniFile = meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                OperateIniFile.WriteIniData("Base", "Patients base", selectItem.ArchiveFolder, meikiniFile);
            }
            if (e.RemovedItems.Count > 0)
            {
                var lostItem = (Person)e.RemovedItems[0];
                lostItem.Icon = "/Images/id_card.png";
            }
        }

        private void exportExcel(string excelFile)
        {
            var selectedUser = this.CodeListBox.SelectedItem as Person;            
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var books = (Microsoft.Office.Interop.Excel.Workbooks)excelApp.Workbooks;
            var book = (Microsoft.Office.Interop.Excel._Workbook)(books.Add(System.Type.Missing));
            var sheets = (Microsoft.Office.Interop.Excel.Sheets)book.Worksheets;
            var sheet = (Microsoft.Office.Interop.Excel._Worksheet)(sheets.get_Item(1));
                        
            ListFiles(new DirectoryInfo(meikFolder));
            

            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("A1", "A100");
            range.ColumnWidth = 15;
            range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("B1", "B100");
            range.ColumnWidth = 30;
            range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("C1", "C100");
            range.ColumnWidth = 20;
            range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("D1", "D100");
            range.ColumnWidth = 100;
            sheet.Cells[1, 1] = App.Current.FindResource("Excel1").ToString();
            sheet.Cells[1, 2] = App.Current.FindResource("Excel2").ToString();
            sheet.Cells[1, 3] = App.Current.FindResource("Excel4").ToString();
            sheet.Cells[1, 4] = App.Current.FindResource("Excel3").ToString();
            for (int i = 0; i < patientList.Count; i++)
            {
                Patient item = patientList[i];
                sheet.Cells[i + 2, 1] = item.Code;
                sheet.Cells[i + 2, 2] = item.Name;
                sheet.Cells[i + 2, 3] = item.ScreenDate;
                sheet.Cells[i + 2, 4] = item.Desc;
            }
            book.SaveAs(excelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            book.Close();
            excelApp.Quit();

        }

        private void ListFiles(FileSystemInfo info)
        {
            if (!info.Exists) return;
            IFormatProvider culture = new System.Globalization.CultureInfo("en-US", true);
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
                            DateTime fileTime=file.LastWriteTime;
                            fileTime = fileTime.AddMonths(-1);
                            string dateStr = fileTime.Month+"/1/"+fileTime.Year;
                            DateTime beginDate = DateTime.Parse(dateStr,culture, System.Globalization.DateTimeStyles.NoCurrentDateDefault);
                            if (beginDate < fileTime)
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
                                patient.ScreenDate = fileTime.ToString();

                                patientList.Add(patient);
                            }
                            
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

        private void btnScreening_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IntPtr screeningBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strScreening);
                Win32Api.SendMessage(screeningBtnHwnd, Win32Api.WM_CLICK, 0, 0);                
                StartMouseHook();
                this.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "System Exception: " + ex.Message);
            }
        }

        private void btnDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            try
            {                
                IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strDiagnostics);
                Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                StartMouseHook();
                this.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "System Exception: " + ex.Message);
            }
        }

        private void btnRecords_Click(object sender, RoutedEventArgs e)
        {
            RecordsWindow recordsWindow = new RecordsWindow();
            recordsWindow.Owner = this;
            recordsWindow.ShowDialog();
        }

        private void btnSetup_Click(object sender, RoutedEventArgs e)
        {            
            ReportSettingPage reportSettingPage = new ReportSettingPage();
            reportSettingPage.Owner = this;
            reportSettingPage.ShowDialog();
        }

        /// <summary>
        /// 启用鼠标钩子
        /// </summary>
        public void StartMouseHook()
        {
            //启动鼠标钩子            
            mouseHook.Start();
        }
        /// <summary>
        /// 停止鼠标钩子
        /// </summary>
        public void StopMouseHook()
        {
            mouseHook.Stop();
        }

        /// <summary>
        /// 鼠标按下的钩子回调方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void mouseHook_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                IntPtr exitButtonHandle = Win32Api.WindowFromPoint(e.X, e.Y);
                IntPtr winHandle = Win32Api.GetParent(exitButtonHandle);
                var owner = this.Owner as MainWindow;
                if (Win32Api.GetParent(winHandle) == owner.AppProc.MainWindowHandle)
                {
                    StringBuilder winText = new StringBuilder(512);
                    Win32Api.GetWindowText(exitButtonHandle, winText, winText.Capacity);
                    if (App.strExit.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        if (App.opendWin != null)
                        {
                            App.opendWin.Visibility = Visibility.Visible;
                            App.opendWin.WindowState = WindowState.Maximized;                            
                        }
                        else
                        {
                            this.Visibility = Visibility.Visible;
                        }
                        this.StopMouseHook();

                        string meikiniFile = meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                        App.dataFolder=OperateIniFile.ReadIniData("Base", "Patients base", "", meikiniFile);
                        txtFolderPath.Text = App.dataFolder;
                        loadArchiveFolder(txtFolderPath.Text);
                    }
                }
            }
        }
        

        private void btnNewArchive_Click(object sender, RoutedEventArgs e)
        {
            AddFolderPage folderPage = new AddFolderPage();
            folderPage.Owner = this;
            folderPage.ShowDialog();
        }

        private void listLang_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var eleObj = listLang.SelectedItem as XmlElement;
            var local = eleObj.GetAttribute("Flag");
            if ("zh-HK".Equals(local))
            {
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
                
            }
            else if ("zh-CN".Equals(local))
            {
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });
                
            }
            else
            {
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });
                
            }
            OperateIniFile.WriteIniData("Base", "Language", local, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;

                //Personal Data
                person.SurName = this.txtName.Text;
                OperateIniFile.WriteIniData("Personal data", "surname", this.txtName.Text, person.CrdFilePath);
                person.GivenName = this.txtGivenName.Text;
                OperateIniFile.WriteIniData("Personal data", "given name", this.txtGivenName.Text, person.CrdFilePath);
                person.OtherName = this.txtOtherName.Text;
                OperateIniFile.WriteIniData("Personal data", "other name", this.txtOtherName.Text, person.CrdFilePath);
                person.Address = this.txtAddress.Text;
                OperateIniFile.WriteIniData("Personal data", "address", this.txtAddress.Text, person.CrdFilePath);
                person.BirthDate = this.txtBirthDate.Text;
                OperateIniFile.WriteIniData("Personal data", "birth date", this.txtBirthDate.Text, person.CrdFilePath);
                person.BirthMonth = this.txtBirthMonth.Text;
                OperateIniFile.WriteIniData("Personal data", "birth month", this.txtBirthMonth.Text, person.CrdFilePath);
                person.BirthYear = this.txtBirthYear.Text;
                OperateIniFile.WriteIniData("Personal data", "birth year", this.txtBirthYear.Text, person.CrdFilePath);
                person.RegDate = this.txtRegDate.Text;
                OperateIniFile.WriteIniData("Personal data", "registration date", this.txtRegDate.Text, person.CrdFilePath);
                person.RegMonth = this.txttRegMonth.Text;
                OperateIniFile.WriteIniData("Personal data", "registration month", this.txttRegMonth.Text, person.CrdFilePath);
                person.RegYear = this.txttRegYear.Text;
                OperateIniFile.WriteIniData("Personal data", "registration year", this.txttRegYear.Text, person.CrdFilePath);

                person.Birthday = person.BirthMonth + "/" + person.BirthDate + "/" + person.BirthYear;
                if (!string.IsNullOrEmpty(person.Birthday))
                {
                    int m_Y1 = DateTime.Parse(person.Birthday).Year;
                    int m_Y2 = DateTime.Now.Year;
                    person.Age = m_Y2 - m_Y1;
                }

                //Complaints
                person.Pain = this.checkPain.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "pain", this.checkPain.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Colostrum = this.checkColostrum.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "colostrum", this.checkColostrum.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.SerousDischarge = this.checkSerousDischarge.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "serous discharge", this.checkSerousDischarge.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BloodDischarge = this.checkBloodDischarge.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "blood discharge", this.checkBloodDischarge.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Other = this.checkOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "other", this.checkOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Pregnancy = this.checkPregnancy.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "pregnancy", this.checkPregnancy.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Lactation = this.checkLactation.IsChecked.Value;
                OperateIniFile.WriteIniData("Complaints", "lactation", this.checkLactation.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.OtherDesc = this.txtOtherDesc.Text;
                var otherDesc = this.txtOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Complaints", "other description", otherDesc, person.CrdFilePath);
                person.PregnancyTerm = this.txtPregnancyTerm.Text;
                OperateIniFile.WriteIniData("Complaints", "pregnancy term", this.txtPregnancyTerm.Text, person.CrdFilePath);

                //Anamnesis
                person.MenstrualCycleDisorder = this.checkMenstrualCycleDisorder.IsChecked.Value;
                OperateIniFile.WriteIniData("Menses", "menstrual cycle disorder", this.checkMenstrualCycleDisorder.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Postmenopause = this.checkPostmenopause.IsChecked.Value;
                OperateIniFile.WriteIniData("Menses", "postmenopause", this.checkPostmenopause.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.HormonalContraceptives = this.checkHormonalContraceptives.IsChecked.Value;
                OperateIniFile.WriteIniData("Menses", "hormonal contraceptives", this.checkHormonalContraceptives.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.MenstrualCycleDisorderDesc = this.txtMenstrualCycleDisorderDesc.Text;
                var menstrualCycleDisorderDesc = this.txtMenstrualCycleDisorderDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Menses", "menstrual cycle disorder description", menstrualCycleDisorderDesc, person.CrdFilePath);


                person.PostmenopauseDesc = this.txtPostmenopauseDesc.Text;
                var postmenopauseDesc = this.txtPostmenopauseDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Menses", "postmenopause description", postmenopauseDesc, person.CrdFilePath);
                person.HormonalContraceptivesBrandName = this.txtHormonalContraceptivesBrandName.Text;
                var hormonalContraceptivesBrandName = this.txtHormonalContraceptivesBrandName.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Menses", "hormonal contraceptives brand name", hormonalContraceptivesBrandName, person.CrdFilePath);
                person.HormonalContraceptivesPeriod = this.txtHormonalContraceptivesPeriod.Text;
                var hormonalContraceptivesPeriod = this.txtHormonalContraceptivesPeriod.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Menses", "hormonal contraceptives period", hormonalContraceptivesPeriod, person.CrdFilePath);

                person.Adiposity = this.checkAdiposity.IsChecked.Value;
                OperateIniFile.WriteIniData("Somatic", "adiposity", this.checkAdiposity.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.EssentialHypertension = this.checkEssentialHypertension.IsChecked.Value;
                OperateIniFile.WriteIniData("Somatic", "essential hypertension", this.checkEssentialHypertension.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Diabetes = this.checkDiabetes.IsChecked.Value;
                OperateIniFile.WriteIniData("Somatic", "diabetes", this.checkDiabetes.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.ThyroidGlandDiseases = this.checkThyroidGlandDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData("Somatic", "thyroid gland diseases", this.checkThyroidGlandDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.SomaticOther = this.checkSomaticOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Somatic", "other", this.checkSomaticOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.EssentialHypertensionDesc = this.txtEssentialHypertensionDesc.Text;
                var essentialHypertensionDesc = this.txtEssentialHypertensionDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Somatic", "essential hypertension description", essentialHypertensionDesc, person.CrdFilePath);
                person.DiabetesDesc = this.txtDiabetesDesc.Text;
                var diabetesDesc = this.txtDiabetesDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Somatic", "diabetes description", diabetesDesc, person.CrdFilePath);
                person.ThyroidGlandDiseasesDesc = this.txtThyroidGlandDiseasesDesc.Text;
                var thyroidGlandDiseasesDesc = this.txtThyroidGlandDiseasesDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Somatic", "thyroid gland diseases description", thyroidGlandDiseasesDesc, person.CrdFilePath);
                person.SomaticOtherDesc = this.txtSomaticOtherDesc.Text;
                var somaticOtherDesc = this.txtSomaticOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Somatic", "other description", somaticOtherDesc, person.CrdFilePath);


                person.Infertility = this.checkInfertility.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "infertility", this.checkInfertility.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryDiseases = this.checkOvaryDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "ovary diseases", this.checkOvaryDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryCyst = this.checkOvaryCyst.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "ovary cyst", this.checkOvaryCyst.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryCancer = this.checkOvaryCancer.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "ovary cancer", this.checkOvaryCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryEndometriosis = this.checkOvaryEndometriosis.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "ovary endometriosis", this.checkOvaryEndometriosis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryOther = this.checkOvaryOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "ovary other", this.checkOvaryOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusDiseases = this.checkUterusDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "uterus diseases", this.checkUterusDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusMyoma = this.checkUterusMyoma.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "uterus myoma", this.checkUterusMyoma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusCancer = this.checkUterusCancer.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "uterus cancer", this.checkUterusCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusEndometriosis = this.checkUterusEndometriosis.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "uterus endometriosis", this.checkUterusEndometriosis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusOther = this.checkUterusOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "uterus other", this.checkUterusOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.GynecologicOther = this.checkGynecologicOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Gynecologic", "other", this.checkGynecologicOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.InfertilityDesc = this.txtInfertility.Text;
                var infertilityDesc = this.txtInfertility.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Gynecologic", "infertility-description", infertilityDesc, person.CrdFilePath);
                person.OvaryOtherDesc = this.txtOvaryOtherDesc.Text;
                var ovaryOtherDesc = this.txtOvaryOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Gynecologic", "ovary other description", ovaryOtherDesc, person.CrdFilePath);
                person.UterusOtherDesc = this.txtUterusOtherDesc.Text;
                var uterusOtherDesc = this.txtUterusOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Gynecologic", "uterus other description", uterusOtherDesc, person.CrdFilePath);
                person.GynecologicOtherDesc = this.txtGynecologicOtherDesc.Text;
                var gynecologicOtherDesc = this.txtGynecologicOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Gynecologic", "other description", gynecologicOtherDesc, person.CrdFilePath);


                person.Abortions = this.checkAbortions.IsChecked.Value;
                OperateIniFile.WriteIniData("Obstetric", "abortions", this.checkAbortions.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Births = this.checkBirths.IsChecked.Value;
                OperateIniFile.WriteIniData("Obstetric", "births", this.checkBirths.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.AbortionsNumber = this.txtAbortionsNumber.Text;
                OperateIniFile.WriteIniData("Obstetric", "abortions number", this.txtAbortionsNumber.Text, person.CrdFilePath);
                person.BirthsNumber = this.txtBirthsNumber.Text;
                OperateIniFile.WriteIniData("Obstetric", "births number", this.txtBirthsNumber.Text, person.CrdFilePath);


                person.LactationTill1Month = this.checkLactationTill1Month.IsChecked.Value;
                OperateIniFile.WriteIniData("Lactation", "lactation till 1 month", this.checkLactationTill1Month.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.LactationTill1Year = this.checkLactationTill1Year.IsChecked.Value;
                OperateIniFile.WriteIniData("Lactation", "lactation till 1 year", this.checkLactationTill1Year.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.LactationOver1Year = this.checkLactationOver1Year.IsChecked.Value;
                OperateIniFile.WriteIniData("Lactation", "lactation over 1 year", this.checkLactationOver1Year.IsChecked.Value ? "1" : "0", person.CrdFilePath);


                person.Trauma = this.checkTrauma.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "trauma", this.checkTrauma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Mastitis = this.checkMastitis.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "mastitis", this.checkMastitis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FibrousCysticMastopathy = this.checkFibrousCysticMastopathy.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "fibrous- cystic mastopathy", this.checkFibrousCysticMastopathy.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Cysts = this.checkCysts.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "cysts", this.checkCysts.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Cancer = this.checkCancer.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "cancer", this.checkCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.DiseasesOther = this.checkDiseasesOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Diseases", "other", this.checkDiseasesOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.TraumaDesc = this.txtTraumaDesc.Text;
                var traumaDesc = this.txtTraumaDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "trauma description", traumaDesc, person.CrdFilePath);
                person.MastitisDesc = this.txtMastitisDesc.Text;
                var mastitisDesc = this.txtMastitisDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "mastitis description", mastitisDesc, person.CrdFilePath);
                person.FibrousCysticMastopathyDesc = this.txtFibrousCysticMastopathyDesc.Text;
                var fibrousCysticMastopathyDesc = this.txtFibrousCysticMastopathyDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "fibrous- cystic mastopathy description", fibrousCysticMastopathyDesc, person.CrdFilePath);
                person.CystsDesc = this.txtCystsDesc.Text;
                var cystsDesc = this.txtCystsDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "cysts descriptuin", cystsDesc, person.CrdFilePath);
                person.CancerDesc = this.txtCancerDesc.Text;
                var cancerDesc = this.txtCancerDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "cancer description", cancerDesc, person.CrdFilePath);
                person.DiseasesOtherDesc = this.txtDiseasesOtherDesc.Text;
                var diseasesOtherDesc = this.txtDiseasesOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Diseases", "other description", diseasesOtherDesc, person.CrdFilePath);


                person.PalpationDiffuse = this.checkPalpationDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData("Palpation", "diffuse", this.checkPalpationDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PalpationFocal = this.checkPalpationFocal.IsChecked.Value;
                OperateIniFile.WriteIniData("Palpation", "focal", this.checkPalpationFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PalpationDesc = this.txtPalpationDesc.Text;
                var palpationDesc = this.txtPalpationDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Palpation", "description", palpationDesc, person.CrdFilePath);

                person.UltrasoundDiffuse = this.checkUltrasoundDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData("Ultrasound", "diffuse", this.checkUltrasoundDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UltrasoundFocal = this.checkUltrasoundFocal.IsChecked.Value;
                OperateIniFile.WriteIniData("Ultrasound", "focal", this.checkUltrasoundFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UltrasoundnDesc = this.txtUltrasoundnDesc.Text;
                var ultrasoundnDesc = this.txtUltrasoundnDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Ultrasound", "description", ultrasoundnDesc, person.CrdFilePath);

                person.MammographyDiffuse = this.checkMammographyDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData("Mammography", "diffuse", this.checkMammographyDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MammographyFocal = this.checkMammographyFocal.IsChecked.Value;
                OperateIniFile.WriteIniData("Mammography", "focal", this.checkMammographyFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MammographyDesc = this.txtMammographyDesc.Text;
                var mammographyDesc = this.txtMammographyDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Mammography", "description", mammographyDesc, person.CrdFilePath);

                person.BiopsyDiffuse = this.checkBiopsyDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "diffuse", this.checkBiopsyDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyFocal = this.checkBiopsyFocal.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "focal", this.checkBiopsyFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyCancer = this.checkBiopsyCancer.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "cancer", this.checkBiopsyCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyProliferation = this.checkBiopsyProliferation.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "proliferation", this.checkBiopsyProliferation.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyDysplasia = this.checkBiopsyDysplasia.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "dysplasia", this.checkBiopsyDysplasia.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyIntraductalPapilloma = this.checkBiopsyIntraductalPapilloma.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "intraductal papilloma", this.checkBiopsyIntraductalPapilloma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyFibroadenoma = this.checkBiopsyFibroadenoma.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "fibroadenoma", this.checkBiopsyFibroadenoma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyOther = this.checkBiopsyOther.IsChecked.Value;
                OperateIniFile.WriteIniData("Biopsy", "other", this.checkBiopsyOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyOtherDesc = this.txtBiopsyOtherDesc.Text;
                var biopsyOtherDesc = this.txtBiopsyOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData("Biopsy", "other description", biopsyOtherDesc, person.CrdFilePath);

                MessageBox.Show(this, App.Current.FindResource("Message_30").ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_31").ToString()+ex.Message);
            }
        }       
    }
}

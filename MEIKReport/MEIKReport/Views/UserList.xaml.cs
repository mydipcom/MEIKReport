using MEIKReport.Common;
using MEIKReport.Model;
using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);
        private delegate void ProgressBarGridDelegate(DependencyProperty dp, Object value);  
        private string deviceNo = OperateIniFile.ReadIniData("Device", "Device No", "000", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //private string meikFolder = OperateIniFile.ReadIniData("Base", "MEIK base", "C:\\Program Files (x86)\\MEIK 5.6", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");        
        protected MouseHook mouseHook = new MouseHook();
        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "Data";

        private ExaminationReportPage examinationReportPage = null;
        private IList<Patient> patientList = new List<Patient>();
        public UserList()
        {
            if (!Directory.Exists(dataFolder))
            {                
                try
                {
                    Directory.CreateDirectory(dataFolder);
                    FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                    FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                }
                catch (Exception) { }
            }
            InitializeComponent();
            languageUS.Foreground = App.local.Equals("en-US") ? new SolidColorBrush(Color.FromRgb(0x40, 0xc4, 0xe4)) : Brushes.White;
            languageHK.Foreground = App.local.Equals("zh-CN") ? new SolidColorBrush(Color.FromRgb(0x40, 0xc4, 0xe4)) : Brushes.White; 
            //listLang.SelectedIndex = App.local.Equals("en-US") ? 0 : App.local.Equals("zh-HK") ? 1 : 1;
            
            //加载初始化配置
            LoadInitConfig();
            string month = DateTime.Now.ToShortDateString();
            //修改原始MEIK程序中的档案改变日期，让原始MEIK程序运行时跨月份打开程序不会出现提示对话框
            OperateIniFile.WriteIniData("Base", "Archive change date", month, App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini");
            //建立当天文件夹          
            string dayFolder = App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + DateTime.Now.ToString("MM_yyyy") + System.IO.Path.DirectorySeparatorChar + DateTime.Now.ToString("dd");
            if (!Directory.Exists(dayFolder))
            {
                Directory.CreateDirectory(dayFolder);
            }
            App.dataFolder = dayFolder;
            //加载当前档案目录数据
            if (!string.IsNullOrEmpty(App.dataFolder))
            {
                loadArchiveFolder(App.dataFolder);
            }
            //显示设备编号
            labDeviceNo.Content = App.reportSettingModel.DeviceNo;
            //根据设备类型显示功能模式
            if (App.reportSettingModel.DeviceType == 1 )
            {
                btnScreening.Visibility = Visibility.Visible;
                btnDiagnostics.Visibility = Visibility.Visible;
                btnRecords.Visibility = Visibility.Visible;
                btnReceivePdf.Visibility = Visibility.Visible;
                btnReceive.Visibility = Visibility.Collapsed;
                btnNewArchive.Visibility = Visibility.Visible;
                btnExaminationReport.Visibility = Visibility.Collapsed;
                sendDataButton.Visibility = Visibility.Visible;
                sendReportButton.Visibility = Visibility.Collapsed;
                
            }            
            else if (App.reportSettingModel.DeviceType == 3)
            {
                btnScreening.Visibility = Visibility.Visible;
                btnDiagnostics.Visibility = Visibility.Visible;
                btnRecords.Visibility = Visibility.Visible;
                btnReceive.Visibility = Visibility.Visible;
                btnReceivePdf.Visibility = Visibility.Visible;
                btnNewArchive.Visibility = Visibility.Visible;
                btnExaminationReport.Visibility = Visibility.Visible;
                sendDataButton.Visibility = Visibility.Visible;
                sendReportButton.Visibility = Visibility.Visible;
            }
            else
            {
                btnScreening.Visibility = Visibility.Collapsed;
                btnDiagnostics.Visibility = Visibility.Collapsed;
                btnRecords.Visibility = Visibility.Collapsed;
                btnReceive.Visibility = Visibility.Visible;
                btnReceivePdf.Visibility = Visibility.Collapsed;
                btnNewArchive.Visibility = Visibility.Collapsed;
                btnExaminationReport.Visibility = Visibility.Visible;
                sendDataButton.Visibility = Visibility.Collapsed;
                sendReportButton.Visibility = Visibility.Visible;
                //醫生模式隱藏個人信息編輯功能
                btnPersonal.IsEnabled = false;
                btnFamily.IsEnabled = false;
                btnComplaints.IsEnabled = false;
                btnMestrual.IsEnabled = false;
                btnObstetric.IsEnabled = false;
                btnAnamnesis.IsEnabled = false;
                btnExaminations.IsEnabled = false;
                btnVisual.IsEnabled = false;
            }            
            
            mouseHook.MouseUp += new System.Windows.Forms.MouseEventHandler(mouseHook_MouseUp);

            ////加载序列化的Screening统计数据
            //try
            //{
            //    string countFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "sc.data";
            //    if (File.Exists(countFile))
            //    {
            //        App.countDictionary = SerializeUtilities.Desrialize<SortedDictionary<string, List<long>>>(countFile);
            //    }
            //}
            //catch (Exception) { }
            
            this.progressBarGrid.DataContext = App.reportSettingModel;
        }        

        private void ExaminationReport_Click(object sender, RoutedEventArgs e)
        {
            //this.Visibility = Visibility.Hidden;
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            
            //var name=selectedUser.GetAttribute("Name");
            // View Examination Report
            examinationReportPage = new ExaminationReportPage(selectedUser);
            //examinationReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            examinationReportPage.Owner = this;
            examinationReportPage.Show();     
        }

        private void SummaryReport_Click(object sender, RoutedEventArgs e)
        {
            //this.Visibility = Visibility.Hidden;
            var selectedUser = this.CodeListBox.SelectedItem as Person;
            //var name = selectedUser.GetAttribute("Name");
            SummaryReportPage summaryReportPage = new SummaryReportPage(selectedUser);
            ////summaryReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            summaryReportPage.Owner = this;
            summaryReportPage.Show();
        }

        private void ShowMainWindow(object sender, EventArgs e)
        {
            this.Show();
        }

        private void Window_Closed(object sender, EventArgs e)
        {                                    
            //this.Owner.Visibility = Visibility.Visible;                
            var owner = this.Owner as MainWindow;
            owner.exitMeik();            
        }

        private void exitReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();            
        }

        private void btnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            //System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            //folderBrowserDialog.SelectedPath = txtFolderPath.Text;
            //System.Windows.Forms.DialogResult res = folderBrowserDialog.ShowDialog();
            //if (res == System.Windows.Forms.DialogResult.OK)
            //{
            //    string folderName = folderBrowserDialog.SelectedPath;
            //    loadArchiveFolder(folderName);
            //}   
            OpenFolderPage ofp = new OpenFolderPage();
            ofp.SelectedPath = txtFolderPath.Text;
            ofp.Owner = this;
            ofp.ShowDialog();
        }


        public void loadArchiveFolder(string folderName)
        {
            try
            {
                //设定系统当前选择的档案文件夹
                App.dataFolder = folderName;          
                txtFolderPath.Text = folderName;
                CollectionViewSource customerSource = (CollectionViewSource)this.FindResource("CustomerSource");
                HashSet<Person> set = new HashSet<Person>();
                //遍历指定文件夹下所有文件
                HandleFolder(folderName,ref set);

                customerSource.Source = set;
                if (set.Count > 0)
                {
                    reportButtonPanel.Visibility = Visibility.Visible;                    
                }
                else
                {
                    reportButtonPanel.Visibility = Visibility.Hidden;                    
                }

                string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
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
                        try
                        {
                            Person person = new Person();
                            //person.ArchiveFolder = folderName;
                            person.ArchiveFolder = theFolder.FullName;
                            person.CrdFilePath = NextFile.FullName;

                            person.Code = NextFile.Name.Substring(0, NextFile.Name.Length - 4);
                            if (App.uploadedCodeList.Contains(person.Code))
                            {
                                person.Uploaded = Visibility.Visible.ToString();
                            }
                            else
                            {
                                person.Uploaded = Visibility.Collapsed.ToString();
                            }

                            bool isZH_CN = false;
                            string local = OperateIniFile.ReadIniData("Report", "language", "", NextFile.FullName);
                            if (!string.IsNullOrEmpty(local))
                            {
                                if ("zh-CN".Equals(local))
                                {
                                    isZH_CN = true;
                                }
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(OperateIniFile.ReadIniData("Personal data", "birth year", "", person.CrdFilePath)))
                                {
                                    isZH_CN = true;
                                }
                            }

                            string personalData = isZH_CN ? "个人信息" : "Personal data";
                            string complaints = isZH_CN ? "病情描述" : "Complaints";
                            string menses = isZH_CN ? "月经" : "Menses";
                            string somatic = isZH_CN ? "慢性病" : "Somatic";
                            string gynecologic = isZH_CN ? "妇科疾病" : "Gynecologic";
                            string obstetric = isZH_CN ? "产科" : "Obstetric";
                            string lactation = isZH_CN ? "哺乳期" : "Lactation";
                            string diseases = isZH_CN ? "乳腺疾病" : "Diseases";
                            string palpation = isZH_CN ? "触诊" : "Palpation";
                            string ultrasound = isZH_CN ? "超声" : "Ultrasound";
                            string mammography = isZH_CN ? "钼靶" : "Mammography";
                            string biopsy = isZH_CN ? "活检" : "Biopsy";
                            string histology = isZH_CN ? "组织学" : "Histology";                            

                            //Personal Data
                            person.ClientNumber = OperateIniFile.ReadIniData(personalData, "clientnumber", "", NextFile.FullName);
                            person.SurName = OperateIniFile.ReadIniData(personalData, "surname", "", NextFile.FullName);
                            person.GivenName = OperateIniFile.ReadIniData(personalData, "given name", "", NextFile.FullName);
                            person.OtherName = OperateIniFile.ReadIniData(personalData, "other name", "", NextFile.FullName);
                            person.Address = OperateIniFile.ReadIniData(personalData, "address", "", NextFile.FullName);
                            person.Height = OperateIniFile.ReadIniData(personalData, "height", "", NextFile.FullName);
                            person.Weight = OperateIniFile.ReadIniData(personalData, "weight", "", NextFile.FullName);
                            person.Mobile = OperateIniFile.ReadIniData(personalData, "mobile", "", NextFile.FullName);
                            person.Email = OperateIniFile.ReadIniData(personalData, "email", "", NextFile.FullName);
                            person.ReportLanguage = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(personalData, "is english", "1", NextFile.FullName)));
                            
                            person.BirthDate = OperateIniFile.ReadIniData(personalData, "birth date", "", NextFile.FullName);
                            person.BirthMonth = OperateIniFile.ReadIniData(personalData, "birth month", "", NextFile.FullName);
                            person.BirthYear = OperateIniFile.ReadIniData(personalData, "birth year", "", NextFile.FullName);
                            person.RegDate = OperateIniFile.ReadIniData(personalData, "registration date", "", NextFile.FullName);
                            person.RegMonth = OperateIniFile.ReadIniData(personalData, "registration month", "", NextFile.FullName);
                            person.RegYear = OperateIniFile.ReadIniData(personalData, "registration year", "", NextFile.FullName);
                            person.Remark = OperateIniFile.ReadIniData(personalData, "remark", "", NextFile.FullName);
                            person.Remark = person.Remark.Replace(";;", "\r\n");

                            person.ScreenVenue = OperateIniFile.ReadIniData("Report", "Screen Venue", "", NextFile.FullName);
                            person.TechName = OperateIniFile.ReadIniData("Report", "Technician Name", "", NextFile.FullName);
                            person.TechLicense = OperateIniFile.ReadIniData("Report", "Technician License", "", NextFile.FullName);
                            try
                            {
                                if (!string.IsNullOrEmpty(person.BirthYear))
                                {
                                    //person.BirthMonth = string.IsNullOrEmpty(person.BirthMonth) ? "1" : person.BirthMonth;
                                    //person.BirthDate = string.IsNullOrEmpty(person.BirthDate) ? "1" : person.BirthDate;
                                    //person.Birthday = person.BirthMonth + "/" + person.BirthDate + "/" + person.BirthYear;
                                    ////person.Regdate = registrationmonth + "/" + registrationdate + "/" + registrationyear;                                
                                    //if (!string.IsNullOrEmpty(person.Birthday))
                                    //{
                                    //    int m_Y1 = DateTime.ParseExact(person.Birthday, "M/d/yyyy", System.Globalization.CultureInfo.CurrentCulture).Year;
                                    //    int m_Y2 = DateTime.Now.Year;
                                    //    person.Age = m_Y2 - m_Y1;
                                    //}
                                    int m_Y1 = Convert.ToInt32(person.BirthYear);
                                    int m_Y2 = DateTime.Now.Year;
                                    person.Age = m_Y2 - m_Y1;
                                }
                                
                            }
                            catch(Exception){ }

                            try
                            {
                                //Family History
                                person.FamilyBreastCancer1 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyBreastCancer1", "0", NextFile.FullName)));
                                person.FamilyBreastCancer2 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyBreastCancer2", "0", NextFile.FullName)));
                                person.FamilyBreastCancer3 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyBreastCancer3", "0", NextFile.FullName)));
                                person.FamilyUterineCancer1 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyUterineCancer1", "0", NextFile.FullName)));
                                person.FamilyUterineCancer2 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyUterineCancer2", "0", NextFile.FullName)));
                                person.FamilyUterineCancer3 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyUterineCancer3", "0", NextFile.FullName)));
                                person.FamilyCervicalCancer1 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyCervicalCancer1", "0", NextFile.FullName)));
                                person.FamilyCervicalCancer2 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyCervicalCancer2", "0", NextFile.FullName)));
                                person.FamilyCervicalCancer3 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyCervicalCancer3", "0", NextFile.FullName)));
                                person.FamilyOvarianCancer1 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyOvarianCancer1", "0", NextFile.FullName)));
                                person.FamilyOvarianCancer2 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyOvarianCancer2", "0", NextFile.FullName)));
                                person.FamilyOvarianCancer3 = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Family History", "FamilyOvarianCancer3", "0", NextFile.FullName)));                                

                            }
                            catch { }


                            try {
                                //Complaints
                                person.PalpableLumps = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "palpable lumps", "0", NextFile.FullName)));
                                if (person.PalpableLumps)
                                {
                                    person.LeftPosition = Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "left position", "0", NextFile.FullName));
                                    person.RightPosition = Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "right position", "0", NextFile.FullName));
                                }
                                                                                                
                                person.Pain = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "pain", "0", NextFile.FullName)));
                                if(person.Pain){
                                    person.Degree = Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "degree", "0", NextFile.FullName));
                                }

                                person.Colostrum = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "colostrum", "0", NextFile.FullName)));
                                person.SerousDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "serous discharge", "0", NextFile.FullName)));
                                person.BloodDischarge = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "blood discharge", "0", NextFile.FullName)));
                                person.Other = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "other", "0", NextFile.FullName)));
                                person.Pregnancy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "pregnancy", "0", NextFile.FullName)));
                                person.Lactation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "lactation", "0", NextFile.FullName)));
                                person.OtherDesc = OperateIniFile.ReadIniData(complaints, "other description", "", NextFile.FullName);
                                person.PregnancyTerm = OperateIniFile.ReadIniData(complaints, "pregnancy term", "", NextFile.FullName);
                                person.OtherDesc = person.OtherDesc.Replace(";;", "\r\n");                                
                                person.LactationTerm = OperateIniFile.ReadIniData(complaints, "lactation term", "", NextFile.FullName);

                                person.BreastImplants = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "BreastImplants", "0", NextFile.FullName)));
                                person.BreastImplantsLeft = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "BreastImplantsLeft", "0", NextFile.FullName)));
                                person.BreastImplantsRight = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "BreastImplantsRight", "0", NextFile.FullName)));
                                person.MaterialsGel = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "MaterialsGel", "0", NextFile.FullName)));
                                person.MaterialsFat = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "MaterialsFat", "0", NextFile.FullName)));
                                person.MaterialsOthers = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(complaints, "MaterialsOthers", "0", NextFile.FullName)));

                                person.BreastImplantsLeftYear = OperateIniFile.ReadIniData(complaints, "BreastImplantsLeftYear", "", NextFile.FullName);
                                person.BreastImplantsRightYear = OperateIniFile.ReadIniData(complaints, "BreastImplantsRightYear", "", NextFile.FullName);
                                
                            }
                            catch (Exception) { }

                            try {
                                person.DateLastMenstruation = OperateIniFile.ReadIniData(menses, "DateLastMenstruation", "", NextFile.FullName);                                
                                person.MenstrualCycleDisorder = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(menses, "menstrual cycle disorder", "0", NextFile.FullName)));
                                person.Postmenopause = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(menses, "postmenopause", "0", NextFile.FullName)));
                                person.HormonalContraceptives = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(menses, "hormonal contraceptives", "0", NextFile.FullName)));
                                person.PostmenopauseYear = OperateIniFile.ReadIniData(menses, "postmenopause year", "", NextFile.FullName);
                                person.MenstrualCycleDisorderDesc = OperateIniFile.ReadIniData(menses, "menstrual cycle disorder description", "", NextFile.FullName);
                                person.MenstrualCycleDisorderDesc = person.MenstrualCycleDisorderDesc.Replace(";;", "\r\n");
                                person.PostmenopauseDesc = OperateIniFile.ReadIniData(menses, "postmenopause description", "", NextFile.FullName);
                                person.PostmenopauseDesc = person.PostmenopauseDesc.Replace(";;", "\r\n");
                                person.HormonalContraceptivesBrandName = OperateIniFile.ReadIniData(menses, "hormonal contraceptives brand name", "", NextFile.FullName);
                                person.HormonalContraceptivesBrandName = person.HormonalContraceptivesBrandName.Replace(";;", "\r\n");
                                person.HormonalContraceptivesPeriod = OperateIniFile.ReadIniData(menses, "hormonal contraceptives period", "", NextFile.FullName);
                                person.HormonalContraceptivesPeriod = person.HormonalContraceptivesPeriod.Replace(";;", "\r\n");
                                person.MensesStatus = person.MenstrualCycleDisorder || person.Postmenopause || person.HormonalContraceptives ? true : false;
                            }
                            catch (Exception) { }

                            try { 
                                person.Adiposity = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(somatic, "adiposity", "0", NextFile.FullName)));
                                person.EssentialHypertension = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(somatic, "essential hypertension", "0", NextFile.FullName)));
                                person.Diabetes = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(somatic, "diabetes", "0", NextFile.FullName)));
                                person.ThyroidGlandDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(somatic, "thyroid gland diseases", "0", NextFile.FullName)));
                                person.SomaticOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(somatic, "other", "0", NextFile.FullName)));
                                person.EssentialHypertensionDesc = OperateIniFile.ReadIniData(somatic, "essential hypertension description", "", NextFile.FullName);
                                person.EssentialHypertensionDesc = person.EssentialHypertensionDesc.Replace(";;", "\r\n");
                                person.DiabetesDesc = OperateIniFile.ReadIniData(somatic, "diabetes description", "", NextFile.FullName);
                                person.DiabetesDesc = person.DiabetesDesc.Replace(";;", "\r\n");
                                person.ThyroidGlandDiseasesDesc = OperateIniFile.ReadIniData(somatic, "thyroid gland diseases description", "", NextFile.FullName);
                                person.ThyroidGlandDiseasesDesc = person.ThyroidGlandDiseasesDesc.Replace(";;", "\r\n");
                                person.SomaticOtherDesc = OperateIniFile.ReadIniData(somatic, "other description", "", NextFile.FullName);
                                person.SomaticOtherDesc = person.SomaticOtherDesc.Replace(";;", "\r\n");
                                person.SomaticStatus = person.Adiposity || person.EssentialHypertension || person.Diabetes || person.ThyroidGlandDiseases || person.SomaticOther ? true : false;
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Infertility = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "infertility", "0", NextFile.FullName)));
                                person.OvaryDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "ovary diseases", "0", NextFile.FullName)));
                                person.OvaryCyst = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "ovary cyst", "0", NextFile.FullName)));
                                person.OvaryCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "ovary cancer", "0", NextFile.FullName)));
                                person.OvaryEndometriosis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "ovary endometriosis", "0", NextFile.FullName)));
                                person.OvaryOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "ovary other", "0", NextFile.FullName)));
                                person.UterusDiseases = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "uterus diseases", "0", NextFile.FullName)));
                                person.UterusMyoma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "uterus myoma", "0", NextFile.FullName)));
                                person.UterusCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "uterus cancer", "0", NextFile.FullName)));
                                person.UterusEndometriosis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "uterus endometriosis", "0", NextFile.FullName)));
                                person.UterusOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "uterus other", "0", NextFile.FullName)));
                                person.GynecologicOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(gynecologic, "other", "0", NextFile.FullName)));
                                person.InfertilityDesc = OperateIniFile.ReadIniData(gynecologic, "infertility-description", "", NextFile.FullName);
                                person.InfertilityDesc = person.InfertilityDesc.Replace(";;", "\r\n");
                                person.OvaryOtherDesc = OperateIniFile.ReadIniData(gynecologic, "ovary other description", "", NextFile.FullName);
                                person.OvaryOtherDesc = person.OvaryOtherDesc.Replace(";;", "\r\n");
                                person.UterusOtherDesc = OperateIniFile.ReadIniData(gynecologic, "uterus other description", "", NextFile.FullName);
                                person.UterusOtherDesc = person.UterusOtherDesc.Replace(";;", "\r\n");
                                person.GynecologicOtherDesc = OperateIniFile.ReadIniData(gynecologic, "other description", "", NextFile.FullName);
                                person.GynecologicOtherDesc = person.GynecologicOtherDesc.Replace(";;", "\r\n");
                                person.GynecologicStatus = person.Infertility || person.OvaryDiseases || person.OvaryCyst || person.OvaryCancer 
                                    || person.OvaryEndometriosis || person.OvaryOther || person.UterusDiseases || person.UterusMyoma
                                    || person.UterusCancer || person.UterusEndometriosis || person.UterusOther || person.GynecologicOther ? true : false;
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Abortions = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(obstetric, "abortions", "0", NextFile.FullName)));
                                person.Births = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(obstetric, "births", "0", NextFile.FullName)));
                                person.AbortionsNumber = OperateIniFile.ReadIniData(obstetric, "abortions number", "", NextFile.FullName);
                                person.BirthsNumber = OperateIniFile.ReadIniData(obstetric, "births number", "", NextFile.FullName);
                                person.ObstetricStatus = person.Abortions || person.Births ? true : false;
                        
                            }
                            catch(Exception){ }

                            try { 
                                person.LactationTill1Month = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(lactation, "lactation till 1 month", "0", NextFile.FullName)));
                                person.LactationTill1Year = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(lactation, "lactation till 1 year", "0", NextFile.FullName)));
                                person.LactationOver1Year = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(lactation, "lactation over 1 year", "0", NextFile.FullName)));
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Trauma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "trauma", "0", NextFile.FullName)));
                                person.Mastitis = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "mastitis", "0", NextFile.FullName)));
                                person.FibrousCysticMastopathy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "fibrous- cystic mastopathy", "0", NextFile.FullName)));
                                person.Cysts = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "cysts", "0", NextFile.FullName)));
                                person.Cancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "cancer", "0", NextFile.FullName)));
                                person.DiseasesOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(diseases, "other", "0", NextFile.FullName)));
                                person.TraumaDesc = OperateIniFile.ReadIniData(diseases, "trauma description", "", NextFile.FullName);
                                person.TraumaDesc = person.TraumaDesc.Replace(";;", "\r\n");
                                person.MastitisDesc = OperateIniFile.ReadIniData(diseases, "mastitis description", "", NextFile.FullName);
                                person.MastitisDesc = person.MastitisDesc.Replace(";;", "\r\n");
                                person.FibrousCysticMastopathyDesc = OperateIniFile.ReadIniData(diseases, "fibrous- cystic mastopathy description", "", NextFile.FullName);
                                person.FibrousCysticMastopathyDesc = person.FibrousCysticMastopathyDesc.Replace(";;", "\r\n");
                                person.CystsDesc = OperateIniFile.ReadIniData(diseases, "cysts descriptuin", "", NextFile.FullName);
                                person.CystsDesc = person.CystsDesc.Replace(";;", "\r\n");
                                person.CancerDesc = OperateIniFile.ReadIniData(diseases, "cancer description", "", NextFile.FullName);
                                person.CancerDesc = person.CancerDesc.Replace(";;", "\r\n");
                                person.DiseasesOtherDesc = OperateIniFile.ReadIniData(diseases, "other description", "", NextFile.FullName);
                                person.DiseasesOtherDesc = person.DiseasesOtherDesc.Replace(";;", "\r\n");
                                person.DiseasesStatus = person.Trauma || person.Mastitis || person.FibrousCysticMastopathy || person.Cysts || person.Cancer || person.DiseasesOther ? true : false;
                            }
                            catch (Exception) { }

                            try {

                                person.Palpation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(palpation, "palpation", "0", NextFile.FullName)));
                                person.PalationYear = OperateIniFile.ReadIniData(palpation, "palpation year", "", NextFile.FullName);                                

                                person.PalpationDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(palpation, "diffuse", "0", NextFile.FullName)));
                                person.PalpationFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(palpation, "focal", "0", NextFile.FullName)));
                                person.PalpationDesc = OperateIniFile.ReadIniData(palpation, "description", "", NextFile.FullName);
                                person.PalpationDesc = person.PalpationDesc.Replace(";;", "\r\n");
                                person.PalpationStatus = person.PalpationDiffuse || person.PalpationFocal ? true : false;
                                //person.PalpationStatus = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(palpation, "palpation status", "0", NextFile.FullName)));
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Ultrasound = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(ultrasound, "ultrasound", "0", NextFile.FullName)));
                                person.UltrasoundYear = OperateIniFile.ReadIniData(ultrasound, "ultrasound year", "", NextFile.FullName);                                
                                person.UltrasoundDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(ultrasound, "diffuse", "0", NextFile.FullName)));
                                person.UltrasoundFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(ultrasound, "focal", "0", NextFile.FullName)));
                                person.UltrasoundnDesc = OperateIniFile.ReadIniData(ultrasound, "description", "", NextFile.FullName);
                                person.UltrasoundnDesc = person.UltrasoundnDesc.Replace(";;", "\r\n");
                                person.UltrasoundStatus = person.UltrasoundDiffuse || person.UltrasoundFocal ? true : false;
                                //person.UltrasoundStatus = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(ultrasound, "ultrasound status", "0", NextFile.FullName)));
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Mammography = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(mammography, "mammography", "0", NextFile.FullName)));
                                person.MammographyYear = OperateIniFile.ReadIniData(mammography, "mammography year", "", NextFile.FullName);  
                                person.MammographyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(mammography, "diffuse", "0", NextFile.FullName)));
                                person.MammographyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(mammography, "focal", "0", NextFile.FullName)));
                                person.MammographyDesc = OperateIniFile.ReadIniData(mammography, "description", "", NextFile.FullName);
                                person.MammographyDesc = person.MammographyDesc.Replace(";;", "\r\n");
                                person.MammographyStatus = person.MammographyDiffuse || person.MammographyFocal ? true : false;
                                //person.MammographyStatus = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(mammography, "mammography status", "0", NextFile.FullName)));
                            }
                            catch (Exception) { }

                            try
                            {
                                person.Biopsy = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "biopsy", "0", NextFile.FullName)));
                                person.BiopsyYear = OperateIniFile.ReadIniData(biopsy, "biopsy year", "", NextFile.FullName); 
                                person.BiopsyDiffuse = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "diffuse", "0", NextFile.FullName)));
                                person.BiopsyFocal = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "focal", "0", NextFile.FullName)));
                                person.BiopsyCancer = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "cancer", "0", NextFile.FullName)));
                                person.BiopsyProliferation = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "proliferation", "0", NextFile.FullName)));
                                person.BiopsyDysplasia = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "dysplasia", "0", NextFile.FullName)));
                                person.BiopsyIntraductalPapilloma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "intraductal papilloma", "0", NextFile.FullName)));
                                person.BiopsyFibroadenoma = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "fibroadenoma", "0", NextFile.FullName)));
                                person.BiopsyOther = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "other", "0", NextFile.FullName)));
                                person.BiopsyOtherDesc = OperateIniFile.ReadIniData(biopsy, "other description", "", NextFile.FullName);
                                person.BiopsyOtherDesc = person.BiopsyOtherDesc.Replace(";;", "\r\n");
                                person.BiopsyStatus = person.BiopsyDiffuse || person.BiopsyFocal || person.BiopsyCancer || person.BiopsyProliferation
                                    || person.BiopsyDysplasia || person.BiopsyIntraductalPapilloma || person.BiopsyFibroadenoma || person.BiopsyOther ? true : false;
                                //person.BiopsyStatus = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData(biopsy, "biopsy status", "0", NextFile.FullName)));
                            }
                            catch (Exception) { }

                            try
                            {
                                person.RedSwollen = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "red swollen", "0", NextFile.FullName)));
                                person.Palpable = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "palpable", "0", NextFile.FullName)));
                                person.Serous = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "serous discharge", "0", NextFile.FullName)));
                                person.Wounds = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "wounds", "0", NextFile.FullName)));
                                person.Scars = Convert.ToBoolean(Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "scars", "0", NextFile.FullName)));
                                person.RedSwollenLeft = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "red swollen left segment", "0", NextFile.FullName));
                                person.RedSwollenRight = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "red swollen right segment", "0", NextFile.FullName));

                                person.PalpableLeft = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "palpable left segment", "0", NextFile.FullName));
                                person.PalpableRight = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "palpable right segment", "0", NextFile.FullName));

                                person.SerousLeft = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "serous left segment", "0", NextFile.FullName));
                                person.SerousRight = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "serous right segment", "0", NextFile.FullName));

                                person.WoundsLeft = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "wounds left segment", "0", NextFile.FullName));
                                person.WoundsRight = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "wounds right segment", "0", NextFile.FullName));

                                person.ScarsLeft = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "scars left segment", "0", NextFile.FullName));
                                person.ScarsRight = Convert.ToInt32(OperateIniFile.ReadIniData("Visual", "scars right segment", "0", NextFile.FullName));
                            }
                            catch { }

                            set.Add(person);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, string.Format(App.Current.FindResource("Message_32").ToString() + ex.Message, NextFile.FullName));
                        }
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

        /// <summary>
        /// Technician发送数据事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SendDataButton_Click(object sender, RoutedEventArgs e)
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
            var selectedUserList = this.CodeListBox.SelectedItems;
            if (selectedUserList == null || selectedUserList.Count == 0)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_35").ToString());
                return;
            }

            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_17").ToString(), "Send Data", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                //判断报告是否已经填写Techinican名称
                SelectTechnicianPage SelectTechnicianPage = new SelectTechnicianPage(App.reportSettingModel.TechNames);
                SelectTechnicianPage.Owner = this;
                SelectTechnicianPage.ShowDialog();
                if (string.IsNullOrEmpty(App.reportSettingModel.ReportTechName))
                {
                    return;
                }

                //隐藏已MRN区域显示的上传过的图标
                foreach (Person selectedUser in selectedUserList)
                {
                    selectedUser.Uploaded = Visibility.Collapsed.ToString();
                }
                //定义委托代理
                ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
                //使用系统代理方式显示进度条面板
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
                try
                {
                    List<string> errMsg = new List<string>();
                    int n = 0;
                    double groupValue = 100 / selectedUserList.Count;
                    double sizeValue = groupValue / 10;
                    //循环处理选择的每一个患者档案
                    foreach (Person selectedUser in selectedUserList)
                    {
                        //string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
                        string dataFile = selectedUser.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
                        //如果是医生模式但当前的患者档案中却没有报告数据，则不能发送数据
                        if (!File.Exists(dataFile) && deviceType == 2)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + App.Current.FindResource("Message_16").ToString());
                            continue;
                        }

                        //把检测员名称和证书写入.crd文件
                        try
                        {
                            OperateIniFile.WriteIniData("Report", "Technician Name", selectedUser.TechName, selectedUser.CrdFilePath);
                            OperateIniFile.WriteIniData("Report", "Technician License", selectedUser.TechLicense, selectedUser.CrdFilePath);
                            OperateIniFile.WriteIniData("Report", "Screen Venue", App.reportSettingModel.ScreenVenue, selectedUser.CrdFilePath);
                            //保存讀取crd數據的語言類型
                            OperateIniFile.WriteIniData("Report", "language", App.local, selectedUser.CrdFilePath);
                        }
                        catch (Exception exe)
                        {
                            //如果不能写入ini文件
                            FileHelper.SetFolderPower(selectedUser.ArchiveFolder, "Everyone", "FullControl");
                            FileHelper.SetFolderPower(selectedUser.ArchiveFolder, "Users", "FullControl");
                        }

                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue) });

                        //打包数据文件，并上传到FPT服务
                        string zipFile = dataFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + "-" + selectedUser.SurName;
                        zipFile = zipFile + (string.IsNullOrEmpty(selectedUser.GivenName) ? "" : "," + selectedUser.GivenName) + (string.IsNullOrEmpty(selectedUser.OtherName) ? "" : " " + selectedUser.OtherName) + ".zip";

                        try
                        {
                            //ZipTools.Instance.ZipFolder(selectedUser.ArchiveFolder, zipFile, 1);
                            ZipTools.Instance.ZipFiles(selectedUser.ArchiveFolder, zipFile, 1);
                            Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue * 5) });
                        }
                        catch (Exception ex1)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + string.Format(App.Current.FindResource("Message_21").ToString() + " " + ex1.Message, selectedUser.ArchiveFolder));
                            continue;
                        }

                        //上传到FTP服务器
                        try
                        {
                            FtpHelper.Instance.Upload(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + "Send/", zipFile);
                            //上传成功后保存上传状态，以便让用户知道这个患者今天已经成功提交过数据
                            if (!App.uploadedCodeList.Contains(selectedUser.Code))
                            {
                                App.uploadedCodeList.Add(selectedUser.Code);
                            }
                            selectedUser.Uploaded = Visibility.Visible.ToString();
                            try
                            {
                                File.Delete(zipFile);
                            }
                            catch { }
                        }
                        catch (Exception ex2)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + App.Current.FindResource("Message_52").ToString() + " " + ex2.Message);
                            continue;
                        }
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue * 9) });

                        ////如果是检测员模式，则每天要上传一次扫描记录
                        //string currentDate = System.DateTime.Now.ToString("yyyyMMdd");
                        //if (App.countDictionary.Count > 0 && !currentDate.Equals(App.reportSettingModel.RecordDate))
                        //{
                        //    string excelFile = dataFolder + System.IO.Path.DirectorySeparatorChar + deviceNo + "_Count.xls";
                        //    try
                        //    {
                        //        if (File.Exists(excelFile))
                        //        {
                        //            try
                        //            {
                        //                File.Delete(excelFile);
                        //            }
                        //            catch { }
                        //        }
                        //        exportExcel(excelFile);
                        //        if (!FtpHelper.Instance.FolderExist(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "ScreeningRecords"))
                        //        {
                        //            FtpHelper.Instance.MakeDir(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "ScreeningRecords");
                        //        }
                        //        FtpHelper.Instance.Upload(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + "/ScreeningRecords/", excelFile);
                        //        App.reportSettingModel.RecordDate = currentDate;
                        //        OperateIniFile.WriteIniData("Data", "Record Date", currentDate, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                        //        try
                        //        {
                        //            File.Delete(excelFile);
                        //        }
                        //        catch { }
                        //    }
                        //    catch { }
                        //}
                        n++;
                    }

                    try
                    {
                        //发送通知邮件  
                        SendEmail((selectedUserList.Count - errMsg.Count), true);
                    }
                    catch (Exception ex3)
                    {
                        errMsg.Add(App.Current.FindResource("Message_19").ToString() + " " + ex3.Message);
                    }

                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                    //显示没有成功发送数据的错误消息
                    if (errMsg.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(string.Format(App.Current.FindResource("Message_36").ToString(), errMsg.Count));
                        foreach (var err in errMsg)
                        {
                            sb.Append("\r\n" + err);
                        }
                        MessageBox.Show(this, sb.ToString());
                    }
                    else
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_18").ToString());
                    }
                }
                finally
                {
                    Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                }
            }
        }

        private void SendReportButton_Click(object sender, RoutedEventArgs e)
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
            var selectedUserList = this.CodeListBox.SelectedItems;
            if (selectedUserList == null || selectedUserList.Count==0)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_35").ToString());
                return;
            }
                                                
            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_17").ToString(), "Send Report", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {                
                foreach (Person selectedUser in selectedUserList)
                {
                    selectedUser.Uploaded = Visibility.Collapsed.ToString();
                }
                ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
                //使用系统代理方式显示进度条面板
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
                try
                {
                    List<string> errMsg = new List<string>();
                    int n = 0;
                    double groupValue = 100 / selectedUserList.Count;
                    double sizeValue = groupValue / 10;
                    //循环处理选择的每一个患者档案
                    foreach (Person selectedUser in selectedUserList)
                    {                        
                        string dataFile = selectedUser.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + selectedUser.Code + ".dat";
                        //如果是医生模式但当前的患者档案中却没有报告数据，则不能发送数据
                        if (!File.Exists(dataFile) && deviceType == 2)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + App.Current.FindResource("Message_16").ToString());
                            continue;
                        }                        

                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue) });

                        //打包数据文件，并上传到FPT服务
                        string zipFile = dataFolder + System.IO.Path.DirectorySeparatorChar + "R_" + selectedUser.Code + "-" + selectedUser.SurName;                        
                        zipFile = zipFile + (string.IsNullOrEmpty(selectedUser.GivenName) ? "" : "," + selectedUser.GivenName) + (string.IsNullOrEmpty(selectedUser.OtherName) ? "" : " " + selectedUser.OtherName) + ".zip";                        

                        try
                        {
                            //ZipTools.Instance.ZipFolder(selectedUser.ArchiveFolder, zipFile, deviceType);
                            ZipTools.Instance.ZipFiles(selectedUser.ArchiveFolder, zipFile,2);
                            Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue * 5) });
                        }
                        catch (Exception ex1)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + string.Format(App.Current.FindResource("Message_21").ToString() + " " + ex1.Message, selectedUser.ArchiveFolder));
                            continue;
                        }

                        //上传到FTP服务器
                        try
                        {
                            FtpHelper.Instance.Upload(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath+"Send/", zipFile);
                            //上传成功后保存上传状态，以便让用户知道这个患者今天已经成功提交过数据
                            if (!App.uploadedCodeList.Contains(selectedUser.Code))
                            {
                                App.uploadedCodeList.Add(selectedUser.Code);
                            }
                            selectedUser.Uploaded = Visibility.Visible.ToString();
                            try
                            {
                                File.Delete(zipFile);
                            }
                            catch { }
                        }
                        catch (Exception ex2)
                        {
                            errMsg.Add(selectedUser.Code + " :: " + App.Current.FindResource("Message_52").ToString() + " " + ex2.Message);
                            continue;
                        }
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue * 9) });
                        
                        n++;
                    }

                    try
                    {
                        //发送通知邮件  
                        SendEmail((selectedUserList.Count - errMsg.Count), true);
                    }
                    catch (Exception ex3)
                    {
                        errMsg.Add(App.Current.FindResource("Message_19").ToString() + " " + ex3.Message);                        
                    }
                    //完成进度条
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });

                    //显示没有成功发送数据的错误消息
                    if (errMsg.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(string.Format(App.Current.FindResource("Message_36").ToString(), errMsg.Count));
                        foreach (var err in errMsg)
                        {
                            sb.Append("\r\n" + err);
                        }
                        MessageBox.Show(this, sb.ToString());
                    }
                    else
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_18").ToString());
                    }
                }
                finally
                {
                    //关闭进度条蒙板
                    Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                }
                
            }
        }

        /// <summary>
        /// 上传数据成功后发送的通知邮件
        /// </summary>
        /// <param name="patientNum"></param>
        /// <param name="isReport"></param>
        private void SendEmail(int patientNum,bool isReport=false)
        {
            //发送通知邮件                
            try
            {
                if (patientNum > 0)
                {
                    string senderServerIp = App.reportSettingModel.MailHost;
                    string toMailAddress = App.reportSettingModel.ToMailAddress;
                    string fromMailAddress = App.reportSettingModel.MailAddress;
                    string subjectInfo = isReport ? (App.reportSettingModel.MailSubject + " (" + App.reportSettingModel.FtpUser + ")") : (App.reportSettingModel.MailSubject + " (" + App.reportSettingModel.DeviceNo + ")");
                    string bodyInfo = string.Format(App.reportSettingModel.MailBody, patientNum);
                    string mailUsername = App.reportSettingModel.MailUsername;
                    string mailPassword = App.reportSettingModel.MailPwd;
                    string mailPort = App.reportSettingModel.MailPort + "";                    
                    bool isSsl = App.reportSettingModel.MailSsl;
                    EmailHelper email = new EmailHelper(senderServerIp, toMailAddress, fromMailAddress, subjectInfo, bodyInfo, mailUsername, mailPassword, mailPort, isSsl, false);                    
                    email.Send();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_19").ToString() + " " + ex.Message);
            }
        }

        private void LoadInitConfig()
        {
            try
            {
                if (File.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "Config_bak.ini"))
                {
                    try
                    {
                        File.Copy(System.AppDomain.CurrentDomain.BaseDirectory + "Config_bak.ini", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini", true);
                        File.Delete(System.AppDomain.CurrentDomain.BaseDirectory + "Config_bak.ini");
                    }
                    catch { }
                }

                if (App.reportSettingModel == null)
                {
                    App.reportSettingModel = new ReportSettingModel();
                    App.reportSettingModel.DataBaseFolder = OperateIniFile.ReadIniData("Base", "Data base", "C:\\MEIKData", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.Version = OperateIniFile.ReadIniData("Base", "Version", "1.0.0", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    
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
                    App.reportSettingModel.ScreenVenue = OperateIniFile.ReadIniData("Report", "Screen Venue", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.NoShowDoctorSignature = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Hide Doctor Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.NoShowTechSignature = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Hide Technician Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    App.reportSettingModel.FtpPath = OperateIniFile.ReadIniData("FTP", "FTP Path", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    App.reportSettingModel.FtpUser = OperateIniFile.ReadIniData("FTP", "FTP User", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    string ftpPwd = OperateIniFile.ReadIniData("FTP", "FTP Password", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    if (!string.IsNullOrEmpty(ftpPwd))
                    {
                        App.reportSettingModel.FtpPwd = SecurityTools.DecryptText(ftpPwd);
                    }

                    string pagesize = OperateIniFile.ReadIniData("Report", "Print Paper", "Letter", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    pagesize = string.IsNullOrEmpty(pagesize) ? "Letter" : pagesize;
                    App.reportSettingModel.PrintPaper = (PageSize)Enum.Parse(typeof(PageSize), pagesize,true);
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
                    //加载操作模式
                    string filePath = System.AppDomain.CurrentDomain.BaseDirectory + "mode.dat";
                    if (File.Exists(filePath))
                    {
                        string modeStr = SecurityTools.DecryptTextFromFile(filePath);
                        if (string.IsNullOrEmpty(modeStr))
                        {
                            modeStr = "Technician";
                        }
                        App.reportSettingModel.DeviceType = (int)Enum.Parse(typeof(DeviceType), modeStr, true);
                    }
                    else
                    {
                        string deviceType = OperateIniFile.ReadIniData("Device", "Device Type", "1", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                        if (string.IsNullOrEmpty(deviceType))
                        {
                            App.reportSettingModel.DeviceType = 1;
                        }
                        else
                        {
                            App.reportSettingModel.DeviceType = Convert.ToInt32(deviceType);
                        }
                        if (App.reportSettingModel.DeviceType == 1)
                        {
                            SecurityTools.EncryptTextToFile("Technician", filePath);
                        }
                        else if (App.reportSettingModel.DeviceType == 3)
                        {
                            SecurityTools.EncryptTextToFile("Admin", filePath);
                        }
                        else
                        {
                            SecurityTools.EncryptTextToFile("Doctor", filePath);
                        }
                    }
                    App.reportSettingModel.RecordDate = OperateIniFile.ReadIniData("Data", "Record Date", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                    
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
            try
            {
                if (e.AddedItems.Count > 0)
                {
                    var selectItems = e.AddedItems;
                    foreach (Person item in selectItems)
                    {
                        item.Icon = "/Images/id_card_ok.png";
                    }
                    var selectItem = e.AddedItems[0] as Person;
                    string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                    OperateIniFile.WriteIniData("Base", "Patients base", selectItem.ArchiveFolder, meikiniFile);

                    if (selectItem.ReportLanguage)
                    {
                        radChinese.IsChecked = false;
                        radEnglish.IsChecked = true;
                    }
                    else
                    {
                        radChinese.IsChecked = true;
                        radEnglish.IsChecked = false;
                    }
                    if (selectItem.PalpableLumps)
                    {
                        leftClock1.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock1.Tag)) > 0 ? 1 : 0;
                        leftClock2.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock2.Tag)) > 0 ? 1 : 0;
                        leftClock3.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock3.Tag)) > 0 ? 1 : 0;
                        leftClock4.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock4.Tag)) > 0 ? 1 : 0;
                        leftClock5.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock5.Tag)) > 0 ? 1 : 0;
                        leftClock6.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock6.Tag)) > 0 ? 1 : 0;
                        leftClock7.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock7.Tag)) > 0 ? 1 : 0;
                        leftClock8.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock8.Tag)) > 0 ? 1 : 0;
                        leftClock9.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock9.Tag)) > 0 ? 1 : 0;
                        leftClock10.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock10.Tag)) > 0 ? 1 : 0;
                        leftClock11.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock11.Tag)) > 0 ? 1 : 0;
                        leftClock12.Opacity = (selectItem.LeftPosition & Convert.ToInt32(leftClock12.Tag)) > 0 ? 1 : 0;

                        rightClock1.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock1.Tag)) > 0 ? 1 : 0;
                        rightClock2.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock2.Tag)) > 0 ? 1 : 0;
                        rightClock3.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock3.Tag)) > 0 ? 1 : 0;
                        rightClock4.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock4.Tag)) > 0 ? 1 : 0;
                        rightClock5.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock5.Tag)) > 0 ? 1 : 0;
                        rightClock6.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock6.Tag)) > 0 ? 1 : 0;
                        rightClock7.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock7.Tag)) > 0 ? 1 : 0;
                        rightClock8.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock8.Tag)) > 0 ? 1 : 0;
                        rightClock9.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock9.Tag)) > 0 ? 1 : 0;
                        rightClock10.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock10.Tag)) > 0 ? 1 : 0;
                        rightClock11.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock11.Tag)) > 0 ? 1 : 0;
                        rightClock12.Opacity = (selectItem.RightPosition & Convert.ToInt32(rightClock12.Tag)) > 0 ? 1 : 0;

                    }
                    else
                    {
                        leftClock1.Opacity = 0;
                        leftClock2.Opacity = 0;
                        leftClock3.Opacity = 0;
                        leftClock4.Opacity = 0;
                        leftClock5.Opacity = 0;
                        leftClock6.Opacity = 0;
                        leftClock7.Opacity = 0;
                        leftClock8.Opacity = 0;
                        leftClock9.Opacity = 0;
                        leftClock9.Opacity = 0;
                        leftClock10.Opacity = 0;
                        leftClock11.Opacity = 0;
                        leftClock12.Opacity = 0;

                        rightClock1.Opacity = 0;
                        rightClock2.Opacity = 0;
                        rightClock3.Opacity = 0;
                        rightClock4.Opacity = 0;
                        rightClock5.Opacity = 0;
                        rightClock6.Opacity = 0;
                        rightClock7.Opacity = 0;
                        rightClock8.Opacity = 0;
                        rightClock9.Opacity = 0;
                        rightClock10.Opacity = 0;
                        rightClock11.Opacity = 0;
                        rightClock12.Opacity = 0;
                    }

                    if (selectItem.Pain)
                    {
                        degree1.IsChecked = (selectItem.Degree == Convert.ToInt32(degree1.Tag)) ? true : false;
                        degree2.IsChecked = (selectItem.Degree == Convert.ToInt32(degree2.Tag)) ? true : false;
                        degree3.IsChecked = (selectItem.Degree == Convert.ToInt32(degree3.Tag)) ? true : false;
                        degree4.IsChecked = (selectItem.Degree == Convert.ToInt32(degree4.Tag)) ? true : false;
                        degree5.IsChecked = (selectItem.Degree == Convert.ToInt32(degree5.Tag)) ? true : false;
                        degree6.IsChecked = (selectItem.Degree == Convert.ToInt32(degree6.Tag)) ? true : false;
                        degree7.IsChecked = (selectItem.Degree == Convert.ToInt32(degree7.Tag)) ? true : false;
                        degree8.IsChecked = (selectItem.Degree == Convert.ToInt32(degree8.Tag)) ? true : false;
                        degree9.IsChecked = (selectItem.Degree == Convert.ToInt32(degree9.Tag)) ? true : false;
                        degree10.IsChecked = (selectItem.Degree == Convert.ToInt32(degree10.Tag)) ? true : false;
                        degree1.IsEnabled = true;
                        degree2.IsEnabled = true;
                        degree3.IsEnabled = true;
                        degree4.IsEnabled = true;
                        degree5.IsEnabled = true;
                        degree6.IsEnabled = true;
                        degree7.IsEnabled = true;
                        degree8.IsEnabled = true;
                        degree9.IsEnabled = true;
                        degree10.IsEnabled = true;
                    }
                    else
                    {
                        degree1.IsChecked = false;
                        degree2.IsChecked = false;
                        degree3.IsChecked = false;
                        degree4.IsChecked = false;
                        degree5.IsChecked = false;
                        degree6.IsChecked = false;
                        degree7.IsChecked = false;
                        degree8.IsChecked = false;
                        degree9.IsChecked = false;
                        degree10.IsChecked = false;
                    }

                    if (selectItem.RedSwollen)
                    {
                        redSwollenLeft.IsEnabled=true;
                        redSwollenRight.IsEnabled = true;
                    }
                    else
                    {
                        redSwollenLeft.IsEnabled = false;
                        redSwollenRight.IsEnabled = false;
                    }

                    if (selectItem.Palpable)
                    {
                        palpableLeft.IsEnabled = true;
                        palpableRight.IsEnabled = true;
                    }
                    else
                    {
                        palpableLeft.IsEnabled = false;
                        palpableRight.IsEnabled = false;
                    }

                    if (selectItem.Serous)
                    {
                        serousLeft.IsEnabled = true;
                        serousRight.IsEnabled = true;
                    }
                    else
                    {
                        serousLeft.IsEnabled = false;
                        serousRight.IsEnabled = false;
                    }

                    if (selectItem.Wounds)
                    {
                        woundsLeft.IsEnabled = true;
                        woundsRight.IsEnabled = true;
                    }
                    else
                    {
                        woundsLeft.IsEnabled = false;
                        woundsRight.IsEnabled = false;
                    }

                    if (selectItem.Scars)
                    {
                        scarsLeft.IsEnabled = true;
                        scarsRight.IsEnabled = true;
                    }
                    else
                    {
                        scarsLeft.IsEnabled = false;
                        scarsRight.IsEnabled = false;
                    }
                }
                if (e.RemovedItems.Count > 0)
                {
                    var lostItem = e.RemovedItems;
                    foreach (Person item in lostItem)
                    {
                        item.Icon = "/Images/id_card.png";
                    }
                }
            }
            catch { }
        }

        //private void exportExcel(string excelFile)
        //{
        //    var selectedUser = this.CodeListBox.SelectedItem as Person;            
        //    var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    var books = (Microsoft.Office.Interop.Excel.Workbooks)excelApp.Workbooks;
        //    var book = (Microsoft.Office.Interop.Excel._Workbook)(books.Add(System.Type.Missing));
        //    var sheets = (Microsoft.Office.Interop.Excel.Sheets)book.Worksheets;
        //    var sheet = (Microsoft.Office.Interop.Excel._Worksheet)(sheets.get_Item(1));
        //    sheet.Name = "TDB Records";
        //    Microsoft.Office.Interop.Excel._Worksheet screensheet;
        //    if (sheets.Count > 1)
        //    {
        //        screensheet = (Microsoft.Office.Interop.Excel._Worksheet)(sheets.get_Item(2));
        //    }
        //    else
        //    {
        //        screensheet = sheets.Add(System.Type.Missing, sheet, System.Type.Missing, System.Type.Missing);                
        //    }
        //    screensheet.Name = "Screening Records";
        //    patientList.Clear();
        //    ListFiles(new DirectoryInfo(App.reportSettingModel.DataBaseFolder));
            
        //    //tdb文件統計數據
        //    Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("A1", "A100");
        //    range.ColumnWidth = 15;
        //    range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("B1", "B100");
        //    range.ColumnWidth = 30;
        //    range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("C1", "C100");
        //    range.ColumnWidth = 20;
        //    range = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("D1", "D100");
        //    range.ColumnWidth = 100;
        //    sheet.Cells[1, 1] = App.Current.FindResource("Excel1").ToString();
        //    sheet.Cells[1, 2] = App.Current.FindResource("Excel2").ToString();
        //    sheet.Cells[1, 3] = App.Current.FindResource("Excel4").ToString();
        //    sheet.Cells[1, 4] = App.Current.FindResource("Excel3").ToString();
        //    for (int i = 0; i < patientList.Count; i++)
        //    {
        //        Patient item = patientList[i];
        //        sheet.Cells[i + 2, 1] = item.Code;
        //        sheet.Cells[i + 2, 2] = item.Name;
        //        sheet.Cells[i + 2, 3] = item.ScreenDate;
        //        sheet.Cells[i + 2, 4] = item.Desc;
        //    }

        //    //启动MEIK设备次数
        //    Microsoft.Office.Interop.Excel.Range range1 = (Microsoft.Office.Interop.Excel.Range)screensheet.get_Range("A1", "A100");
        //    range1.ColumnWidth = 100;
        //    range1 = (Microsoft.Office.Interop.Excel.Range)screensheet.get_Range("B1", "B100");
        //    range1.ColumnWidth = 30;
            
        //    screensheet.Cells[1, 1] = App.Current.FindResource("Excel5").ToString();
        //    screensheet.Cells[1, 2] = App.Current.FindResource("Excel6").ToString();
        //    int row = 1;
        //    foreach (KeyValuePair<string, List<long>> item in App.countDictionary)                            
        //    {
        //        screensheet.Cells[row + 1, 1] = item.Key;                
        //        foreach (var tick in item.Value)
        //        {
        //            DateTime screeningTime = new DateTime(tick);
        //            screensheet.Cells[row + 1, 2] = screeningTime.ToString("yyyy-MM-dd HH:mm:ss");                    
        //            row++; 
        //        }                            
        //    }

        //    book.SaveAs(excelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    book.Close();
        //    excelApp.Quit();
            
        //}

        private void ListFiles(FileSystemInfo info)
        {
            try
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
                            DateTime beginDate = DateTime.Now;
                            if (".tdb".Equals(file.Extension, StringComparison.OrdinalIgnoreCase))
                            {
                                DateTime fileTime = file.LastWriteTime;
                                beginDate = beginDate.AddMonths(-1);
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
                                    patient.ScreenDate = fileTime.ToString("yyyy-MM-dd HH:mm:ss");

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
            catch { }
        }

        private void btnScreening_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Visibility = Visibility.Hidden;
                IntPtr screeningBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strScreening);
                Win32Api.SendMessage(screeningBtnHwnd, Win32Api.WM_CLICK, 0, 0);                
                StartMouseHook();                
            }
            catch (Exception ex)
            {
                this.Visibility = Visibility.Visible;
                MessageBox.Show(this, "System Exception: " + ex.Message);
            }
        }

        private void btnDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Visibility = Visibility.Hidden;
                IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strDiagnostics);
                Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                StartMouseHook();                
            }
            catch (Exception ex)
            {
                this.Visibility = Visibility.Visible;
                MessageBox.Show(this, "System Exception: " + ex.Message);                
            }
        }

        /// <summary>
        /// 统计按钮点击处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRecords_Click(object sender, RoutedEventArgs e)
        {
            RecordsWindow recordsWindow = new RecordsWindow();
            recordsWindow.Owner = this;
            recordsWindow.ShowDialog();
        }

        /// <summary>
        /// 系统设置按钮点击处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetup_Click(object sender, RoutedEventArgs e)
        {            
            ReportSettingPage reportSettingPage = new ReportSettingPage();
            reportSettingPage.Owner = this;
            reportSettingPage.ShowDialog();
        }

        /// <summary>
        /// 医生模式接收服务中心分配的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReceive_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
            //使用系统代理方式显示进度条面板
            Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(0) });
            Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
            
            List<string> errMsg = new List<string>();
            try
            {
                string targetFolder = "Receive/";                
                //查询FTP上要下载的文件列表
                var fileArr = FtpHelper.Instance.GetFileList(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + targetFolder);
                //判断是否存在已下载的目录IsDownloaded
                if (!FtpHelper.Instance.FolderExist(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "IsDownloaded"))
                {
                    FtpHelper.Instance.MakeDir(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "IsDownloaded");
                }                               
                
                int n = 0;
                double groupValue = 100 / (fileArr.Count + 1);
                double sizeValue = groupValue / 2;

                foreach (var file in fileArr)
                {                                        
                    if (file.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) && !file.StartsWith("R_"))
                    {
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n) });
                        n++;
                        //创建下载目录路径
                        string folderpath = "";
                        try
                        {
                            string monthfolder = file.Substring(2, 2) + "_20" + file.Substring(0, 2);
                            folderpath = App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + monthfolder + System.IO.Path.DirectorySeparatorChar + file.Substring(4, 2);
                            if (!Directory.Exists(folderpath))
                            {
                                Directory.CreateDirectory(folderpath);
                            }
                        }
                        catch (Exception e1)
                        {
                            //MessageBox.Show(this, e1.Message);
                            errMsg.Add(file + " :: " + App.Current.FindResource("Message_39").ToString() + e1.Message);
                            continue;
                        }
                        //从FTP下载zip包
                        try
                        {
                            FtpHelper.Instance.Download(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + targetFolder, folderpath, file);
                            FtpHelper.Instance.MovieFile(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + targetFolder, file, "/home/IsDownloaded/" + file);
                        }
                        catch (Exception e2)
                        {
                            //MessageBox.Show(this, e2.Message);
                            errMsg.Add(file + " :: " + App.Current.FindResource("Message_40").ToString() + e2.Message);
                            continue;
                        }
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + sizeValue) });
                        //解压缩档案Zip包
                        try
                        {
                            //获取档案目录路径
                            string archiveFolder = folderpath + System.IO.Path.DirectorySeparatorChar + file.Replace(".zip", "");
                            if (!Directory.Exists(archiveFolder))
                            {
                                //创建患者档案目录
                                Directory.CreateDirectory(archiveFolder);
                            }
                            //ZipTools.Instance.UnzipToFolder(folderpath + System.IO.Path.DirectorySeparatorChar + file, archiveFolder);
                            ZipTools.Instance.UnZip(folderpath + System.IO.Path.DirectorySeparatorChar + file, archiveFolder);
                        }
                        catch (Exception e3)
                        {
                            //MessageBox.Show(this, e3.Message);
                            errMsg.Add(file + " :: " + App.Current.FindResource("Message_41").ToString() + e3.Message);
                            continue;
                        }
                    }
                }
                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                if (errMsg.Count == 0)
                {
                    if (n == 0)
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_43").ToString());
                    }
                    else
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_37").ToString());                        
                    }
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(App.Current.FindResource("Message_38").ToString());
                    foreach (var err in errMsg)
                    {
                        sb.Append("\r\n" + err);
                    }
                    MessageBox.Show(this, sb.ToString());
                }
                loadArchiveFolder(txtFolderPath.Text);
            }
            catch (Exception exe)
            {
                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                MessageBox.Show(this, App.Current.FindResource("Message_42").ToString() + exe.Message);                
            }                                    
        }

        /// <summary>
        /// Technician模式接收PDF報告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReceivePdf_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarGridDelegate progressBarGridDelegate = new ProgressBarGridDelegate(progressBarGrid.SetValue);
            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(uploadProgressBar.SetValue);
            //使用系统代理方式显示进度条面板
            Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(0) });
            Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Visible });
            
            List<string> errMsg = new List<string>();
            try
            {
                //查询FTP上要下载的文件列表
                var fileArr = FtpHelper.Instance.GetFileList(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + "Receive/");
                //判断是否存在已下载的目录IsDownloaded
                if (!FtpHelper.Instance.FolderExist(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "IsDownloaded"))
                {
                    FtpHelper.Instance.MakeDir(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath, "IsDownloaded");
                }

                int n = 0;
                double groupValue = 100 / (fileArr.Count + 1);                

                foreach (var file in fileArr)
                {
                    if (file.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n) });
                        n++;
                        //创建下载目录路径
                        string folderpath = "";
                        try
                        {
                            string monthfolder = file.Substring(2, 2) + "_20" + file.Substring(0, 2);
                            folderpath = App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + monthfolder + System.IO.Path.DirectorySeparatorChar + file.Substring(4, 2);
                            if (!Directory.Exists(folderpath))
                            {
                                Directory.CreateDirectory(folderpath);
                            }
                        }
                        catch (Exception e1)
                        {
                            //MessageBox.Show(this, e1.Message);
                            errMsg.Add(file + " :: " + App.Current.FindResource("Message_39").ToString() + e1.Message);
                            continue;
                        }
                        //从FTP下载PDF文件
                        try
                        {
                            FtpHelper.Instance.Download(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + "Receive/", folderpath, file);
                            Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(groupValue * n + groupValue/2) });
                            FtpHelper.Instance.MovieFile(App.reportSettingModel.FtpUser, App.reportSettingModel.FtpPwd, App.reportSettingModel.FtpPath + "Receive/", file, "/home/IsDownloaded/" + file);
                        }
                        catch (Exception e2)
                        {
                            //MessageBox.Show(this, e2.Message);
                            errMsg.Add(file + " :: " + App.Current.FindResource("Message_40").ToString() + e2.Message);
                            continue;
                        }                        
                        
                    }
                }
                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                if (errMsg.Count == 0)
                {
                    if (n == 0)
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_43").ToString());
                    }
                    else
                    {
                        MessageBox.Show(this, App.Current.FindResource("Message_37").ToString());
                    }
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(App.Current.FindResource("Message_38").ToString());
                    foreach (var err in errMsg)
                    {
                        sb.Append("\r\n" + err);
                    }
                    MessageBox.Show(this, sb.ToString());
                }
            }
            catch (Exception exe)
            {
                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.ProgressBar.ValueProperty, Convert.ToDouble(100) });
                Dispatcher.Invoke(progressBarGridDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { System.Windows.Controls.Grid.VisibilityProperty, Visibility.Collapsed });
                MessageBox.Show(this, App.Current.FindResource("Message_45").ToString() + exe.Message);
            }
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
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    IntPtr buttonHandle = Win32Api.WindowFromPoint(e.X, e.Y);
                    IntPtr winHandle = Win32Api.GetParent(buttonHandle);
                    var owner = this.Owner as MainWindow;
                    if (Win32Api.GetParent(winHandle) == owner.AppProc.MainWindowHandle)
                    {
                        StringBuilder winText = new StringBuilder(512);
                        Win32Api.GetWindowText(buttonHandle, winText, winText.Capacity);
                        if (App.strExit.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            this.StopMouseHook();
                            if (App.opendWin != null)
                            {
                                App.opendWin.Visibility = Visibility.Visible;
                                App.opendWin.WindowState = WindowState.Maximized;
                            }
                            else
                            {
                                this.Visibility = Visibility.Visible;
                            }

                            string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                            txtFolderPath.Text = OperateIniFile.ReadIniData("Base", "Patients base", "", meikiniFile);
                            loadArchiveFolder(txtFolderPath.Text);
                        }
                        //else if (App.strStart.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))
                        //{
                        //    string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                        //    var screenFolderPath = OperateIniFile.ReadIniData("Base", "Patients base", "", meikiniFile);
                        //    if(App.countDictionary.ContainsKey(screenFolderPath)){
                        //        List<long> ticks=App.countDictionary[screenFolderPath];
                        //        ticks.Add(DateTime.Now.Ticks);
                        //    }
                        //    else{
                        //        List<long> ticks=new List<long>();
                        //        ticks.Add(DateTime.Now.Ticks);
                        //        App.countDictionary.Add(screenFolderPath,ticks);
                        //    }                        
                        //    //序列化统计字典到文件
                        //    string countFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "sc.data";
                        //    SerializeUtilities.Serialize<SortedDictionary<string, List<long>>>(App.countDictionary, countFile);
                        //}
                    }
                }
            }
            catch { }
        }
        
        /// <summary>
        /// 添加新的档案文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewArchive_Click(object sender, RoutedEventArgs e)
        {
            AddFolderPage folderPage = new AddFolderPage();
            folderPage.Owner = this;
            folderPage.ShowDialog();
        }

        ///// <summary>
        ///// 选择语言(下拉列表方式)
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void listLang_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    var eleObj = listLang.SelectedItem as XmlElement;
        //    var local = eleObj.GetAttribute("Flag");
        //    App.local = local;
        //    if ("zh-HK".Equals(local))
        //    {
        //        App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });

        //    }
        //    else if ("zh-CN".Equals(local))
        //    {
        //        App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });

        //    }
        //    else
        //    {
        //        App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
        //        App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });

        //    }
        //    OperateIniFile.WriteIniData("Base", "Language", local, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        //}

        /// <summary>
        /// 选择语言
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Language_Click(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;            
            string local = "en-US";
            if ("languageHK".Equals(button.Name))
            {
                local = "zh-HK";
                this.languageHK.Foreground = new SolidColorBrush(Color.FromRgb(0x40, 0xc4, 0xe4));
                this.languageUS.Foreground = Brushes.White;
            }
            //else if ("languageCN".Equals(button.Name))
            //{
            //    local = "zh-CN";
            //    this.languageHK.Foreground = new SolidColorBrush(Color.FromRgb(0x40, 0xc4, 0xe4));
            //    this.languageUS.Foreground = Brushes.White;
            //}
            else
            {
                this.languageUS.Foreground = new SolidColorBrush(Color.FromRgb(0x40, 0xc4, 0xe4));
                this.languageHK.Foreground = Brushes.White;
            }
            App.local = local;
            if ("zh-HK".Equals(local))
            {
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });

            }
            //else if ("zh-CN".Equals(local))
            //{
            //    App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
            //    App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
            //    App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });

            //}
            else
            {
                App.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-CN.xaml", UriKind.RelativeOrAbsolute) });
                App.Current.Resources.MergedDictionaries.Remove(new ResourceDictionary() { Source = new Uri(@"/Resources/StringResource.zh-HK.xaml", UriKind.RelativeOrAbsolute) });

            }
            OperateIniFile.WriteIniData("Base", "Language", local, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        }  

        /// <summary>
        /// 保存患者病历卡
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person == null)
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_35").ToString());
                    return;
                }

                bool isZH_CN = false;
                string local = OperateIniFile.ReadIniData("Report", "language", "", person.CrdFilePath);
                if (!string.IsNullOrEmpty(local))
                {
                    if ("zh-CN".Equals(local))
                    {
                        isZH_CN = true;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(OperateIniFile.ReadIniData("Personal data", "birth year", "", person.CrdFilePath)))
                    {
                        isZH_CN = true;
                    }
                }

                string personalData = isZH_CN ? "个人信息" : "Personal data";
                string complaints = isZH_CN ? "病情描述" : "Complaints";
                string menses = isZH_CN ? "月经" : "Menses";
                string somatic = isZH_CN ? "慢性病" : "Somatic";
                string gynecologic = isZH_CN ? "妇科疾病" : "Gynecologic";
                string obstetric = isZH_CN ? "产科" : "Obstetric";
                string lactation = isZH_CN ? "哺乳期" : "Lactation";
                string diseases = isZH_CN ? "乳腺疾病" : "Diseases";
                string palpation = isZH_CN ? "触诊" : "Palpation";
                string ultrasound = isZH_CN ? "超声" : "Ultrasound";
                string mammography = isZH_CN ? "钼靶" : "Mammography";
                string biopsy = isZH_CN ? "活检" : "Biopsy";
                string histology = isZH_CN ? "组织学" : "Histology";                

                //Personal Data
                person.ClientNumber = this.txtClientNum.Text;
                OperateIniFile.WriteIniData(personalData, "clientnumber", this.txtClientNum.Text, person.CrdFilePath);
                person.SurName = this.txtName.Text;
                OperateIniFile.WriteIniData(personalData, "surname", this.txtName.Text, person.CrdFilePath);
                person.GivenName = this.txtGivenName.Text;
                OperateIniFile.WriteIniData(personalData, "given name", this.txtGivenName.Text, person.CrdFilePath);
                person.OtherName = this.txtOtherName.Text;
                OperateIniFile.WriteIniData(personalData, "other name", this.txtOtherName.Text, person.CrdFilePath);
                person.Address = this.txtAddress.Text;
                OperateIniFile.WriteIniData(personalData, "address", this.txtAddress.Text, person.CrdFilePath);
                person.Height = this.txtHeight.Text;
                OperateIniFile.WriteIniData(personalData, "height", this.txtHeight.Text, person.CrdFilePath);
                person.Weight = this.txtWeight.Text;
                OperateIniFile.WriteIniData(personalData, "weight", this.txtWeight.Text, person.CrdFilePath);
                person.Mobile = this.txtMobileNumber.Text;
                OperateIniFile.WriteIniData(personalData, "mobile", this.txtMobileNumber.Text, person.CrdFilePath);
                person.Email = this.txtEmail.Text;
                OperateIniFile.WriteIniData(personalData, "email", this.txtEmail.Text, person.CrdFilePath);
               person.ReportLanguage = this.radEnglish.IsChecked.Value;
                OperateIniFile.WriteIniData(personalData, "is english", this.radEnglish.IsChecked.Value?"1":"0", person.CrdFilePath);


                person.BirthDate = this.txtBirthDate.Text;
                OperateIniFile.WriteIniData(personalData, "birth date", this.txtBirthDate.Text, person.CrdFilePath);
                person.BirthMonth = this.txtBirthMonth.Text;
                OperateIniFile.WriteIniData(personalData, "birth month", this.txtBirthMonth.Text, person.CrdFilePath);
                person.BirthYear = this.txtBirthYear.Text;
                OperateIniFile.WriteIniData(personalData, "birth year", this.txtBirthYear.Text, person.CrdFilePath);
                person.RegDate = this.txtRegDate.Text;
                OperateIniFile.WriteIniData(personalData, "registration date", this.txtRegDate.Text, person.CrdFilePath);
                person.RegMonth = this.txttRegMonth.Text;
                OperateIniFile.WriteIniData(personalData, "registration month", this.txttRegMonth.Text, person.CrdFilePath);
                person.RegYear = this.txttRegYear.Text;
                OperateIniFile.WriteIniData(personalData, "registration year", this.txttRegYear.Text, person.CrdFilePath);
                person.Remark = this.txtRemark.Text;
                string remark = this.txtRemark.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(personalData, "remark", remark, person.CrdFilePath);
                try
                {
                    if (!string.IsNullOrEmpty(person.BirthYear))
                    {
                        //person.BirthMonth = string.IsNullOrEmpty(person.BirthMonth) ? "1" : person.BirthMonth;
                        //person.BirthDate = string.IsNullOrEmpty(person.BirthDate) ? "1" : person.BirthDate;
                        //person.Birthday = person.BirthMonth + "/" + person.BirthDate + "/" + person.BirthYear;
                        ////person.Regdate = registrationmonth + "/" + registrationdate + "/" + registrationyear;                                
                        //if (!string.IsNullOrEmpty(person.Birthday))
                        //{
                        //    int m_Y1 = DateTime.Parse(person.Birthday).Year;
                        //    int m_Y2 = DateTime.Now.Year;
                        //    person.Age = m_Y2 - m_Y1;
                        //}

                        int m_Y1 = Convert.ToInt32(person.BirthYear);
                        int m_Y2 = DateTime.Now.Year;
                        person.Age = m_Y2 - m_Y1;
                    }

                }
                catch (Exception) { }

                //Family History
                person.FamilyBreastCancer1 = this.checkBreastCancer1.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyBreastCancer1", this.checkBreastCancer1.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyBreastCancer2 = this.checkBreastCancer2.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyBreastCancer2", this.checkBreastCancer2.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyBreastCancer3 = this.checkBreastCancer3.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyBreastCancer3", this.checkBreastCancer3.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyUterineCancer1 = this.checkUterineCancer1.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyUterineCancer1", this.checkUterineCancer1.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyUterineCancer2 = this.checkUterineCancer2.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyUterineCancer2", this.checkUterineCancer2.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyUterineCancer3 = this.checkUterineCancer3.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyUterineCancer3", this.checkUterineCancer3.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyCervicalCancer1 = this.checkCervicalCancer1.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyCervicalCancer1", this.checkCervicalCancer1.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyCervicalCancer2 = this.checkCervicalCancer2.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyCervicalCancer2", this.checkCervicalCancer2.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyCervicalCancer3 = this.checkCervicalCancer3.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyCervicalCancer3", this.checkCervicalCancer3.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyOvarianCancer1 = this.checkOvarianCancer1.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyOvarianCancer1", this.checkOvarianCancer1.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyOvarianCancer2 = this.checkOvarianCancer2.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyOvarianCancer2", this.checkOvarianCancer2.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FamilyOvarianCancer3 = this.checkOvarianCancer3.IsChecked.Value;
                OperateIniFile.WriteIniData("Family History", "FamilyOvarianCancer3", this.checkOvarianCancer3.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                //Complaints
                person.PalpableLumps = this.checkPalpableLumps.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "palpable lumps", this.checkPalpableLumps.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                if(this.checkPalpableLumps.IsChecked.Value){
                    person.LeftPosition  += leftClock1.Opacity>0?Convert.ToInt32(leftClock1.Tag):0;
                    person.LeftPosition  += leftClock2.Opacity>0?Convert.ToInt32(leftClock2.Tag):0;
                    person.LeftPosition  += leftClock3.Opacity>0?Convert.ToInt32(leftClock3.Tag):0;
                    person.LeftPosition  += leftClock4.Opacity>0?Convert.ToInt32(leftClock4.Tag):0;
                    person.LeftPosition  += leftClock5.Opacity>0?Convert.ToInt32(leftClock5.Tag):0;
                    person.LeftPosition  += leftClock6.Opacity>0?Convert.ToInt32(leftClock6.Tag):0;
                    person.LeftPosition  += leftClock7.Opacity>0?Convert.ToInt32(leftClock7.Tag):0;
                    person.LeftPosition  += leftClock8.Opacity>0?Convert.ToInt32(leftClock8.Tag):0;
                    person.LeftPosition  += leftClock9.Opacity>0?Convert.ToInt32(leftClock9.Tag):0;
                    person.LeftPosition  += leftClock10.Opacity>0?Convert.ToInt32(leftClock10.Tag):0;
                    person.LeftPosition  += leftClock11.Opacity>0?Convert.ToInt32(leftClock11.Tag):0;
                    person.LeftPosition  += leftClock12.Opacity>0?Convert.ToInt32(leftClock12.Tag):0;
                    OperateIniFile.WriteIniData(complaints, "left position", person.LeftPosition.ToString(), person.CrdFilePath);
                    person.RightPosition += rightClock1.Opacity > 0 ? Convert.ToInt32(rightClock1.Tag) : 0;
                    person.RightPosition += rightClock2.Opacity > 0 ? Convert.ToInt32(rightClock2.Tag) : 0;
                    person.RightPosition += rightClock3.Opacity > 0 ? Convert.ToInt32(rightClock3.Tag) : 0;
                    person.RightPosition += rightClock4.Opacity > 0 ? Convert.ToInt32(rightClock4.Tag) : 0;
                    person.RightPosition += rightClock5.Opacity > 0 ? Convert.ToInt32(rightClock5.Tag) : 0;
                    person.RightPosition += rightClock6.Opacity > 0 ? Convert.ToInt32(rightClock6.Tag) : 0;
                    person.RightPosition += rightClock7.Opacity > 0 ? Convert.ToInt32(rightClock7.Tag) : 0;
                    person.RightPosition += rightClock8.Opacity > 0 ? Convert.ToInt32(rightClock8.Tag) : 0;
                    person.RightPosition += rightClock9.Opacity > 0 ? Convert.ToInt32(rightClock9.Tag) : 0;
                    person.RightPosition += rightClock10.Opacity > 0 ? Convert.ToInt32(rightClock10.Tag) : 0;
                    person.RightPosition += rightClock11.Opacity > 0 ? Convert.ToInt32(rightClock11.Tag) : 0;
                    person.RightPosition += rightClock12.Opacity > 0 ? Convert.ToInt32(rightClock12.Tag) : 0; 
                    OperateIniFile.WriteIniData(complaints, "right position", person.RightPosition.ToString(), person.CrdFilePath);
                }

                if (degree1.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree1.Tag);
                }
                if (degree2.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree2.Tag);
                }
                if (degree3.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree3.Tag);
                }
                if (degree4.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree4.Tag);
                }
                if (degree5.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree5.Tag);
                }
                if (degree6.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree6.Tag);
                }
                if (degree7.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree7.Tag);
                }
                if (degree8.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree8.Tag);
                }
                if (degree9.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree9.Tag);
                }
                if (degree10.IsChecked.Value)
                {
                    person.Degree = Convert.ToInt32(degree10.Tag);
                }
                OperateIniFile.WriteIniData(complaints, "degree", person.Degree.ToString(), person.CrdFilePath);

                person.Pain = this.checkPain.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "pain", this.checkPain.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Colostrum = this.checkColostrum.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "colostrum", this.checkColostrum.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.SerousDischarge = this.checkSerousDischarge.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "serous discharge", this.checkSerousDischarge.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BloodDischarge = this.checkBloodDischarge.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "blood discharge", this.checkBloodDischarge.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Other = this.checkOther.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "other", this.checkOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Pregnancy = this.checkPregnancy.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "pregnancy", this.checkPregnancy.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Lactation = this.checkLactation.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "lactation", this.checkLactation.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.LactationTerm = this.txtLactationTerm.Text;
                OperateIniFile.WriteIniData(complaints, "lactation term", this.txtLactationTerm.Text, person.CrdFilePath);
                person.OtherDesc = this.txtOtherDesc.Text;
                var otherDesc = this.txtOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(complaints, "other description", otherDesc, person.CrdFilePath);
                person.PregnancyTerm = this.txtPregnancyTerm.Text;
                OperateIniFile.WriteIniData(complaints, "pregnancy term", this.txtPregnancyTerm.Text, person.CrdFilePath);

                person.BreastImplants = this.checkBreastImplants.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "BreastImplants", this.checkBreastImplants.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BreastImplantsLeft = this.checkBreastImplantsLeft.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "BreastImplantsLeft", this.checkBreastImplantsLeft.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BreastImplantsRight = this.checkBreastImplantsRight.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "BreastImplantsRight", this.checkBreastImplantsRight.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MaterialsGel = this.checkMaterialsGel.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "MaterialsGel", this.checkMaterialsGel.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MaterialsFat = this.checkMaterialsFat.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "MaterialsFat", this.checkMaterialsFat.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MaterialsOthers = this.checkMaterialsOthers.IsChecked.Value;
                OperateIniFile.WriteIniData(complaints, "MaterialsOthers", this.checkMaterialsOthers.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.BreastImplantsLeftYear = this.txtBreastImplantsLeftYear.Text;
                OperateIniFile.WriteIniData(complaints, "BreastImplantsLeftYear", this.txtBreastImplantsLeftYear.Text, person.CrdFilePath);
                person.BreastImplantsRightYear = this.txtBreastImplantsRightYear.Text;
                OperateIniFile.WriteIniData(complaints, "BreastImplantsRightYear", this.txtBreastImplantsRightYear.Text, person.CrdFilePath);

                //Anamnesis
                person.DateLastMenstruation = this.dateLastMenstruation.Text;
                OperateIniFile.WriteIniData(menses, "DateLastMenstruation", this.dateLastMenstruation.Text, person.CrdFilePath);
                person.MenstrualCycleDisorder = this.checkMenstrualCycleDisorder.IsChecked.Value;
                OperateIniFile.WriteIniData(menses, "menstrual cycle disorder", this.checkMenstrualCycleDisorder.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Postmenopause = this.checkPostmenopause.IsChecked.Value;
                OperateIniFile.WriteIniData(menses, "postmenopause", this.checkPostmenopause.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PostmenopauseYear = this.txtPostmenopauseYear.Text;
                OperateIniFile.WriteIniData(menses, "postmenopause year", this.txtPostmenopauseYear.Text, person.CrdFilePath);

                person.HormonalContraceptives = this.checkHormonalContraceptives.IsChecked.Value;
                OperateIniFile.WriteIniData(menses, "hormonal contraceptives", this.checkHormonalContraceptives.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.MenstrualCycleDisorderDesc = this.txtMenstrualCycleDisorderDesc.Text;
                var menstrualCycleDisorderDesc = this.txtMenstrualCycleDisorderDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(menses, "menstrual cycle disorder description", menstrualCycleDisorderDesc, person.CrdFilePath);


                person.PostmenopauseDesc = this.txtPostmenopauseDesc.Text;
                var postmenopauseDesc = this.txtPostmenopauseDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(menses, "postmenopause description", postmenopauseDesc, person.CrdFilePath);
                person.HormonalContraceptivesBrandName = this.txtHormonalContraceptivesBrandName.Text;
                var hormonalContraceptivesBrandName = this.txtHormonalContraceptivesBrandName.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(menses, "hormonal contraceptives brand name", hormonalContraceptivesBrandName, person.CrdFilePath);
                person.HormonalContraceptivesPeriod = this.txtHormonalContraceptivesPeriod.Text;
                var hormonalContraceptivesPeriod = this.txtHormonalContraceptivesPeriod.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(menses, "hormonal contraceptives period", hormonalContraceptivesPeriod, person.CrdFilePath);

                person.Adiposity = this.checkAdiposity.IsChecked.Value;
                OperateIniFile.WriteIniData(somatic, "adiposity", this.checkAdiposity.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.EssentialHypertension = this.checkEssentialHypertension.IsChecked.Value;
                OperateIniFile.WriteIniData(somatic, "essential hypertension", this.checkEssentialHypertension.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Diabetes = this.checkDiabetes.IsChecked.Value;
                OperateIniFile.WriteIniData(somatic, "diabetes", this.checkDiabetes.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.ThyroidGlandDiseases = this.checkThyroidGlandDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData(somatic, "thyroid gland diseases", this.checkThyroidGlandDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.SomaticOther = this.checkSomaticOther.IsChecked.Value;
                OperateIniFile.WriteIniData(somatic, "other", this.checkSomaticOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.EssentialHypertensionDesc = this.txtEssentialHypertensionDesc.Text;
                var essentialHypertensionDesc = this.txtEssentialHypertensionDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(somatic, "essential hypertension description", essentialHypertensionDesc, person.CrdFilePath);
                person.DiabetesDesc = this.txtDiabetesDesc.Text;
                var diabetesDesc = this.txtDiabetesDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(somatic, "diabetes description", diabetesDesc, person.CrdFilePath);
                person.ThyroidGlandDiseasesDesc = this.txtThyroidGlandDiseasesDesc.Text;
                var thyroidGlandDiseasesDesc = this.txtThyroidGlandDiseasesDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(somatic, "thyroid gland diseases description", thyroidGlandDiseasesDesc, person.CrdFilePath);
                person.SomaticOtherDesc = this.txtSomaticOtherDesc.Text;
                var somaticOtherDesc = this.txtSomaticOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(somatic, "other description", somaticOtherDesc, person.CrdFilePath);


                person.Infertility = this.checkInfertility.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "infertility", this.checkInfertility.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryDiseases = this.checkOvaryDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "ovary diseases", this.checkOvaryDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryCyst = this.checkOvaryCyst.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "ovary cyst", this.checkOvaryCyst.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryCancer = this.checkOvaryCancer.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "ovary cancer", this.checkOvaryCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryEndometriosis = this.checkOvaryEndometriosis.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "ovary endometriosis", this.checkOvaryEndometriosis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.OvaryOther = this.checkOvaryOther.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "ovary other", this.checkOvaryOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusDiseases = this.checkUterusDiseases.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "uterus diseases", this.checkUterusDiseases.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusMyoma = this.checkUterusMyoma.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "uterus myoma", this.checkUterusMyoma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusCancer = this.checkUterusCancer.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "uterus cancer", this.checkUterusCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusEndometriosis = this.checkUterusEndometriosis.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "uterus endometriosis", this.checkUterusEndometriosis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UterusOther = this.checkUterusOther.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "uterus other", this.checkUterusOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.GynecologicOther = this.checkGynecologicOther.IsChecked.Value;
                OperateIniFile.WriteIniData(gynecologic, "other", this.checkGynecologicOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.InfertilityDesc = this.txtInfertility.Text;
                var infertilityDesc = this.txtInfertility.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(gynecologic, "infertility-description", infertilityDesc, person.CrdFilePath);
                person.OvaryOtherDesc = this.txtOvaryOtherDesc.Text;
                var ovaryOtherDesc = this.txtOvaryOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(gynecologic, "ovary other description", ovaryOtherDesc, person.CrdFilePath);
                person.UterusOtherDesc = this.txtUterusOtherDesc.Text;
                var uterusOtherDesc = this.txtUterusOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(gynecologic, "uterus other description", uterusOtherDesc, person.CrdFilePath);
                person.GynecologicOtherDesc = this.txtGynecologicOtherDesc.Text;
                var gynecologicOtherDesc = this.txtGynecologicOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(gynecologic, "other description", gynecologicOtherDesc, person.CrdFilePath);


                person.Abortions = this.checkAbortions.IsChecked.Value;
                OperateIniFile.WriteIniData(obstetric, "abortions", this.checkAbortions.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Births = this.checkBirths.IsChecked.Value;
                OperateIniFile.WriteIniData(obstetric, "births", this.checkBirths.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.AbortionsNumber = this.txtAbortionsNumber.Text;
                OperateIniFile.WriteIniData(obstetric, "abortions number", this.txtAbortionsNumber.Text, person.CrdFilePath);
                person.BirthsNumber = this.txtBirthsNumber.Text;
                OperateIniFile.WriteIniData(obstetric, "births number", this.txtBirthsNumber.Text, person.CrdFilePath);


                person.LactationTill1Month = this.checkLactationTill1Month.IsChecked.Value;
                OperateIniFile.WriteIniData(lactation, "lactation till 1 month", this.checkLactationTill1Month.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.LactationTill1Year = this.checkLactationTill1Year.IsChecked.Value;
                OperateIniFile.WriteIniData(lactation, "lactation till 1 year", this.checkLactationTill1Year.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.LactationOver1Year = this.checkLactationOver1Year.IsChecked.Value;
                OperateIniFile.WriteIniData(lactation, "lactation over 1 year", this.checkLactationOver1Year.IsChecked.Value ? "1" : "0", person.CrdFilePath);


                person.Trauma = this.checkTrauma.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "trauma", this.checkTrauma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Mastitis = this.checkMastitis.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "mastitis", this.checkMastitis.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.FibrousCysticMastopathy = this.checkFibrousCysticMastopathy.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "fibrous- cystic mastopathy", this.checkFibrousCysticMastopathy.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Cysts = this.checkCysts.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "cysts", this.checkCysts.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Cancer = this.checkCancer.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "cancer", this.checkCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.DiseasesOther = this.checkDiseasesOther.IsChecked.Value;
                OperateIniFile.WriteIniData(diseases, "other", this.checkDiseasesOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.TraumaDesc = this.txtTraumaDesc.Text;
                var traumaDesc = this.txtTraumaDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "trauma description", traumaDesc, person.CrdFilePath);
                person.MastitisDesc = this.txtMastitisDesc.Text;
                var mastitisDesc = this.txtMastitisDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "mastitis description", mastitisDesc, person.CrdFilePath);
                person.FibrousCysticMastopathyDesc = this.txtFibrousCysticMastopathyDesc.Text;
                var fibrousCysticMastopathyDesc = this.txtFibrousCysticMastopathyDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "fibrous- cystic mastopathy description", fibrousCysticMastopathyDesc, person.CrdFilePath);
                person.CystsDesc = this.txtCystsDesc.Text;
                var cystsDesc = this.txtCystsDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "cysts descriptuin", cystsDesc, person.CrdFilePath);
                person.CancerDesc = this.txtCancerDesc.Text;
                var cancerDesc = this.txtCancerDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "cancer description", cancerDesc, person.CrdFilePath);
                person.DiseasesOtherDesc = this.txtDiseasesOtherDesc.Text;
                var diseasesOtherDesc = this.txtDiseasesOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(diseases, "other description", diseasesOtherDesc, person.CrdFilePath);

                person.Palpation = this.chkPalpation.IsChecked.Value;
                OperateIniFile.WriteIniData(palpation, "palpation", this.chkPalpation.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PalationYear = this.txtPalationYear.Text;
                OperateIniFile.WriteIniData(palpation, "palpation year", this.txtPalationYear.Text, person.CrdFilePath);
                person.PalpationDiffuse = this.checkPalpationDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData(palpation, "diffuse", this.checkPalpationDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PalpationFocal = this.checkPalpationFocal.IsChecked.Value;
                OperateIniFile.WriteIniData(palpation, "focal", this.checkPalpationFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.PalpationDesc = this.txtPalpationDesc.Text;
                var palpationDesc = this.txtPalpationDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(palpation, "palpation status", this.radPalpationStatusAbnormal.IsChecked.Value ? "1" : "0", person.CrdFilePath);                

                person.Ultrasound = this.chkUltrasound.IsChecked.Value;
                OperateIniFile.WriteIniData(ultrasound, "ultrasound", this.chkUltrasound.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UltrasoundYear = this.txtUltrasoundYear.Text;
                OperateIniFile.WriteIniData(ultrasound, "ultrasound year", this.txtUltrasoundYear.Text, person.CrdFilePath);
                person.UltrasoundDiffuse = this.checkUltrasoundDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData(ultrasound, "diffuse", this.checkUltrasoundDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UltrasoundFocal = this.checkUltrasoundFocal.IsChecked.Value;
                OperateIniFile.WriteIniData(ultrasound, "focal", this.checkUltrasoundFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.UltrasoundnDesc = this.txtUltrasoundnDesc.Text;
                var ultrasoundnDesc = this.txtUltrasoundnDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(ultrasound, "description", ultrasoundnDesc, person.CrdFilePath);
                OperateIniFile.WriteIniData(ultrasound, "ultrasound status", this.radUltrasoundStatusAbnormal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                

                person.Mammography = this.chkMammography.IsChecked.Value;
                OperateIniFile.WriteIniData(mammography, "mammography", this.chkMammography.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MammographyYear = this.txtMammographyYear.Text;
                OperateIniFile.WriteIniData(mammography, "mammography year", this.txtMammographyYear.Text, person.CrdFilePath);
                person.MammographyDiffuse = this.checkMammographyDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData(mammography, "diffuse", this.checkMammographyDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MammographyFocal = this.checkMammographyFocal.IsChecked.Value;
                OperateIniFile.WriteIniData(mammography, "focal", this.checkMammographyFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.MammographyDesc = this.txtMammographyDesc.Text;
                var mammographyDesc = this.txtMammographyDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(mammography, "description", mammographyDesc, person.CrdFilePath);
                OperateIniFile.WriteIniData(mammography, "mammography status", this.radMammographyStatusAbnormal.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.Biopsy = this.chkBiopsy.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "biopsy", this.chkBiopsy.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyYear = this.txtBiopsyYear.Text;
                OperateIniFile.WriteIniData(biopsy, "biopsy year", this.txtBiopsyYear.Text, person.CrdFilePath);
                person.BiopsyDiffuse = this.checkBiopsyDiffuse.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "diffuse", this.checkBiopsyDiffuse.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyFocal = this.checkBiopsyFocal.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "focal", this.checkBiopsyFocal.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyCancer = this.checkBiopsyCancer.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "cancer", this.checkBiopsyCancer.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyProliferation = this.checkBiopsyProliferation.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "proliferation", this.checkBiopsyProliferation.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyDysplasia = this.checkBiopsyDysplasia.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "dysplasia", this.checkBiopsyDysplasia.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyIntraductalPapilloma = this.checkBiopsyIntraductalPapilloma.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "intraductal papilloma", this.checkBiopsyIntraductalPapilloma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyFibroadenoma = this.checkBiopsyFibroadenoma.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "fibroadenoma", this.checkBiopsyFibroadenoma.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyOther = this.checkBiopsyOther.IsChecked.Value;
                OperateIniFile.WriteIniData(biopsy, "other", this.checkBiopsyOther.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.BiopsyOtherDesc = this.txtBiopsyOtherDesc.Text;
                var biopsyOtherDesc = this.txtBiopsyOtherDesc.Text.Replace("\r\n", ";;");
                OperateIniFile.WriteIniData(biopsy, "other description", biopsyOtherDesc, person.CrdFilePath);
                OperateIniFile.WriteIniData(biopsy, "biopsy status", this.radBiopsyStatusAbnormal.IsChecked.Value ? "1" : "0", person.CrdFilePath);

                person.RedSwollen = this.chkRedSwollen.IsChecked.Value;
                OperateIniFile.WriteIniData("Visual", "red swollen", this.chkRedSwollen.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Palpable = this.chkPalpable.IsChecked.Value;
                OperateIniFile.WriteIniData("Visual", "palpable", this.chkPalpable.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Serous = this.chkSerous.IsChecked.Value;
                OperateIniFile.WriteIniData("Visual", "serous discharge", this.chkSerous.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Wounds = this.chkWounds.IsChecked.Value;
                OperateIniFile.WriteIniData("Visual", "wounds", this.chkWounds.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.Scars = this.chkScars.IsChecked.Value;
                OperateIniFile.WriteIniData("Visual", "scars", this.chkScars.IsChecked.Value ? "1" : "0", person.CrdFilePath);
                person.RedSwollenLeft = this.redSwollenLeft.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "red swollen left segment", this.redSwollenLeft.SelectedIndex.ToString(), person.CrdFilePath);
                person.RedSwollenRight = this.redSwollenRight.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "red swollen right segment", this.redSwollenRight.SelectedIndex.ToString(), person.CrdFilePath);

                person.PalpableLeft = this.palpableLeft.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "palpable left segment", this.palpableLeft.SelectedIndex.ToString(), person.CrdFilePath);
                person.PalpableRight = this.palpableRight.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "palpable right segment", this.palpableRight.SelectedIndex.ToString(), person.CrdFilePath);

                person.SerousLeft = this.serousLeft.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "serous left segment", this.serousLeft.SelectedIndex.ToString(), person.CrdFilePath);
                person.SerousRight = this.serousRight.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "serous right segment", this.serousRight.SelectedIndex.ToString(), person.CrdFilePath);

                person.WoundsLeft = this.woundsLeft.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "wounds left segment", this.woundsLeft.SelectedIndex.ToString(), person.CrdFilePath);
                person.WoundsRight = this.woundsRight.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "wounds right segment", this.woundsRight.SelectedIndex.ToString(), person.CrdFilePath);

                person.ScarsLeft = this.scarsLeft.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "scars left segment", this.scarsLeft.SelectedIndex.ToString(), person.CrdFilePath);
                person.ScarsRight = this.scarsRight.SelectedIndex;
                OperateIniFile.WriteIniData("Visual", "scars right segment", this.scarsRight.SelectedIndex.ToString(), person.CrdFilePath);

                MessageBox.Show(this, App.Current.FindResource("Message_30").ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_31").ToString()+ex.Message);
            }
        }       

        private void checkMenses_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.MensesStatus = true;
                    }
                    else
                    {
                        person.MensesStatus = this.checkMenstrualCycleDisorder.IsChecked.Value || this.checkPostmenopause.IsChecked.Value || this.checkHormonalContraceptives.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkSomatic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.SomaticStatus = true;
                    }
                    else
                    {
                        person.SomaticStatus = this.checkAdiposity.IsChecked.Value || this.checkEssentialHypertension.IsChecked.Value || this.checkDiabetes.IsChecked.Value || this.checkThyroidGlandDiseases.IsChecked.Value || this.checkSomaticOther.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkGynecologic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.GynecologicStatus = true;
                    }
                    else
                    {
                        person.GynecologicStatus = this.checkInfertility.IsChecked.Value || this.checkOvaryDiseases.IsChecked.Value || this.checkOvaryCyst.IsChecked.Value
                            || this.checkOvaryCancer.IsChecked.Value || this.checkOvaryEndometriosis.IsChecked.Value
                            || this.checkOvaryOther.IsChecked.Value || this.checkUterusDiseases.IsChecked.Value
                            || this.checkUterusMyoma.IsChecked.Value || this.checkUterusCancer.IsChecked.Value
                            || this.checkUterusEndometriosis.IsChecked.Value || this.checkUterusOther.IsChecked.Value
                            || this.checkGynecologicOther.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        } 

        private void checkObstetric_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.ObstetricStatus = true;
                    }
                    else
                    {
                        person.ObstetricStatus = this.checkAbortions.IsChecked.Value || this.checkBirths.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkDiseases_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.DiseasesStatus = true;
                    }
                    else
                    {
                        person.DiseasesStatus = this.checkTrauma.IsChecked.Value || this.checkMastitis.IsChecked.Value
                            || this.checkFibrousCysticMastopathy.IsChecked.Value || this.checkCysts.IsChecked.Value
                            || this.checkCancer.IsChecked.Value || this.checkDiseasesOther.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkPalpation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.PalpationStatus = true;
                    }
                    else
                    {
                        person.PalpationStatus = this.checkPalpationDiffuse.IsChecked.Value || this.checkPalpationFocal.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkUltrasound_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.UltrasoundStatus = true;
                    }
                    else
                    {
                        person.UltrasoundStatus = this.checkUltrasoundDiffuse.IsChecked.Value || this.checkUltrasoundFocal.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkMammography_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.MammographyStatus = true;
                    }
                    else
                    {
                        person.MammographyStatus = this.checkMammographyDiffuse.IsChecked.Value || this.checkMammographyFocal.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void checkBiopsy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var person = this.CodeListBox.SelectedItem as Person;
                if (person != null)
                {
                    CheckBox chk = (CheckBox)sender;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        person.BiopsyStatus = true;
                    }
                    else
                    {
                        person.BiopsyStatus = this.checkBiopsyDiffuse.IsChecked.Value || this.checkBiopsyFocal.IsChecked.Value
                            || this.checkBiopsyCancer.IsChecked.Value || this.checkBiopsyProliferation.IsChecked.Value
                            || this.checkBiopsyDysplasia.IsChecked.Value || this.checkBiopsyIntraductalPapilloma.IsChecked.Value
                            || this.checkBiopsyFibroadenoma.IsChecked.Value || this.checkBiopsyOther.IsChecked.Value ? true : false;
                    }
                }
            }
            catch { }
        }

        private void imgChoice_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (CodeListBox.SelectionMode.Equals(SelectionMode.Single))
                {
                    CodeListBox.SelectionMode = SelectionMode.Extended;
                    this.imgChoice.Source = new BitmapImage(new Uri("/Images/multiple_choice.png", UriKind.Relative));
                }
                else if (CodeListBox.SelectionMode.Equals(SelectionMode.Extended))
                {
                    CodeListBox.SelectionMode = SelectionMode.Single;
                    this.imgChoice.Source = new BitmapImage(new Uri("/Images/single_choice.png", UriKind.Relative));
                }
            }
            catch { }   
        }

        private bool isHighLight = false;
        private void Clock_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (checkPalpableLumps.IsChecked.Value)
                {
                    var clock = (Polygon)sender;
                    //var name = clock.Name;
                    //var brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x87, 0x36, 0x32));
                    //var defaultBrush = new SolidColorBrush(Color.FromArgb(0x00, 0x32, 0x87, 0x87));
                    if (!isHighLight)
                    {
                        clock.Opacity = 1;
                        isHighLight = true;
                    }
                    else
                    {
                        clock.Opacity = 0;
                        isHighLight = false;
                    }
                }
            }
            catch { }
        }
        
        private void Clock_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                if (checkPalpableLumps.IsChecked.Value)
                {
                    var clock = (Polygon)sender;
                    isHighLight = Convert.ToBoolean(clock.Opacity);
                    clock.Opacity = 1;
                }
            }
            catch { }
            
        }

        private void Clock_MouseLeave(object sender, MouseEventArgs e)
        {
            try
            {
                if (checkPalpableLumps.IsChecked.Value)
                {
                    var clock = (Polygon)sender;
                    if (!isHighLight)
                    {
                        clock.Opacity = 0;
                    }
                }
            }
            catch { }
        }       

        private void checkPalpableLumps_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var checkBox = (CheckBox)sender;
                if (!checkBox.IsChecked.Value)
                {
                    leftClock1.Opacity = 0;
                    leftClock2.Opacity = 0;
                    leftClock3.Opacity = 0;
                    leftClock4.Opacity = 0;
                    leftClock5.Opacity = 0;
                    leftClock6.Opacity = 0;
                    leftClock7.Opacity = 0;
                    leftClock8.Opacity = 0;
                    leftClock9.Opacity = 0;
                    leftClock9.Opacity = 0;
                    leftClock10.Opacity = 0;
                    leftClock11.Opacity = 0;
                    leftClock12.Opacity = 0;

                    rightClock1.Opacity = 0;
                    rightClock2.Opacity = 0;
                    rightClock3.Opacity = 0;
                    rightClock4.Opacity = 0;
                    rightClock5.Opacity = 0;
                    rightClock6.Opacity = 0;
                    rightClock7.Opacity = 0;
                    rightClock8.Opacity = 0;
                    rightClock9.Opacity = 0;
                    rightClock10.Opacity = 0;
                    rightClock11.Opacity = 0;
                    rightClock12.Opacity = 0;
                }
            }
            catch { }
            
        }

        private void checkPain_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var checkBox = (CheckBox)sender;
                if (!checkBox.IsChecked.Value)
                {
                    degree1.IsChecked = false;
                    degree2.IsChecked = false;
                    degree3.IsChecked = false;
                    degree4.IsChecked = false;
                    degree5.IsChecked = false;
                    degree6.IsChecked = false;
                    degree7.IsChecked = false;
                    degree8.IsChecked = false;
                    degree9.IsChecked = false;
                    degree10.IsChecked = false;
                    degree1.IsEnabled = false;
                    degree2.IsEnabled = false;
                    degree3.IsEnabled = false;
                    degree4.IsEnabled = false;
                    degree5.IsEnabled = false;
                    degree6.IsEnabled = false;
                    degree7.IsEnabled = false;
                    degree8.IsEnabled = false;
                    degree9.IsEnabled = false;
                    degree10.IsEnabled = false;
                }
                else
                {
                    degree1.IsEnabled = true;
                    degree2.IsEnabled = true;
                    degree3.IsEnabled = true;
                    degree4.IsEnabled = true;
                    degree5.IsEnabled = true;
                    degree6.IsEnabled = true;
                    degree7.IsEnabled = true;
                    degree8.IsEnabled = true;
                    degree9.IsEnabled = true;
                    degree10.IsEnabled = true;
                }
            }
            catch { }
        }

        private void txtClientNum_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var textBox=(TextBox)sender;
            if (textBox.Text.Length > 10)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
           
        }

        private void Number_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9]+");
                if (re.IsMatch(e.Text))
                {
                    var textBox = (TextBox)sender;
                    if (!string.IsNullOrEmpty(textBox.Text))
                    {
                        if (Convert.ToInt32(textBox.Text + e.Text) < 100)
                        {
                            e.Handled = false;
                        }
                        else
                        {
                            MessageBox.Show(this, App.Current.FindResource("Message_56").ToString());
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        e.Handled = false;
                    }

                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_56").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Year_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9]+");
                if (re.IsMatch(e.Text))
                {
                    var textBox = (TextBox)sender;
                    if (!string.IsNullOrEmpty(textBox.Text))
                    {
                        if (Convert.ToInt32(textBox.Text + e.Text) <= DateTime.Now.Year)
                        {
                            e.Handled = false;
                        }
                        else
                        {
                            MessageBox.Show(this, App.Current.FindResource("Message_59").ToString());
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        e.Handled = false;
                    }

                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_53").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Month_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9]+");
                if (re.IsMatch(e.Text))
                {
                    var textBox = (TextBox)sender;
                    if (!string.IsNullOrEmpty(textBox.Text))
                    {
                        if (Convert.ToInt32(textBox.Text + e.Text) <= DateTime.Now.Year)
                        {
                            e.Handled = false;
                        }
                        else
                        {
                            MessageBox.Show(this, App.Current.FindResource("Message_57").ToString());
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        e.Handled = false;
                    }

                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_53").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Day_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9]+");
                if (re.IsMatch(e.Text))
                {
                    var textBox = (TextBox)sender;
                    if (!string.IsNullOrEmpty(textBox.Text))
                    {
                        if (Convert.ToInt32(textBox.Text + e.Text) <= DateTime.Now.Year)
                        {
                            e.Handled = false;
                        }
                        else
                        {
                            MessageBox.Show(this, App.Current.FindResource("Message_58").ToString());
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        e.Handled = false;
                    }

                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_53").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Weight_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9]+");
                if (re.IsMatch(e.Text))
                {
                    var textBox = (TextBox)sender;
                    if (!string.IsNullOrEmpty(textBox.Text))
                    {
                        if (Convert.ToInt32(textBox.Text + e.Text) <500)
                        {
                            e.Handled = false;
                        }
                        else
                        {
                            MessageBox.Show(this, App.Current.FindResource("Message_60").ToString());
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        e.Handled = false;
                    }

                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_53").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Mobile_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Regex re = new Regex("[0-9+-]+");
                if (re.IsMatch(e.Text))
                {                    
                    e.Handled = false;                    
                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_61").ToString());
                    e.Handled = true;
                }
            }
            catch (Exception)
            {
                e.Handled = true;
            }

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (examinationReportPage != null&&examinationReportPage.IsActive)
            {                
                examinationReportPage.Close();
                
            }
        }

        private void chkRedSwollen_Click(object sender, RoutedEventArgs e)
        {
            var person = this.CodeListBox.SelectedItem as Person;
            if (person != null)
            {
                CheckBox chk = (CheckBox)sender;
                if (!chk.IsChecked.HasValue || !chk.IsChecked.Value)
                {                
                    redSwollenLeft.SelectedIndex = 0;
                    redSwollenRight.SelectedIndex = 0;
                    redSwollenLeft.IsEnabled = false;
                    redSwollenRight.IsEnabled = false;
                }
                else
                {
                    redSwollenLeft.IsEnabled = true;
                    redSwollenRight.IsEnabled = true;
                }
            }
        }

        private void chkPalpable_Click(object sender, RoutedEventArgs e)
        {
            var person = this.CodeListBox.SelectedItem as Person;
            if (person != null)
            {
                CheckBox chk = (CheckBox)sender;
                if (!chk.IsChecked.HasValue || !chk.IsChecked.Value)                
                {
                    palpableLeft.SelectedIndex = 0;
                    palpableRight.SelectedIndex = 0;
                    palpableLeft.IsEnabled = false;
                    palpableRight.IsEnabled = false;
                }
                else
                {
                    palpableLeft.IsEnabled = true;
                    palpableRight.IsEnabled = true;
                }
            }
        }

        private void chkSerous_Click(object sender, RoutedEventArgs e)
        {
            var person = this.CodeListBox.SelectedItem as Person;
            if (person != null)
            {
                CheckBox chk = (CheckBox)sender;
                if (!chk.IsChecked.HasValue || !chk.IsChecked.Value)
                {
                    serousLeft.SelectedIndex = 0;
                    serousRight.SelectedIndex = 0;
                    serousLeft.IsEnabled = false;
                    serousRight.IsEnabled = false;
                }
                else
                {
                    serousLeft.IsEnabled = true;
                    serousRight.IsEnabled = true;
                }
            }
        }

        private void chkWounds_Click(object sender, RoutedEventArgs e)
        {
            var person = this.CodeListBox.SelectedItem as Person;
            if (person != null)
            {
                CheckBox chk = (CheckBox)sender;
                if (!chk.IsChecked.HasValue || !chk.IsChecked.Value)
                {
                    woundsLeft.SelectedIndex = 0;
                    woundsRight.SelectedIndex = 0;
                    woundsLeft.IsEnabled = false;
                    woundsRight.IsEnabled = false;
                }
                else
                {
                    woundsLeft.IsEnabled = true;
                    woundsRight.IsEnabled = true;
                }
            }
        }

        private void chkScars_Click(object sender, RoutedEventArgs e)
        {
            var person = this.CodeListBox.SelectedItem as Person;
            if (person != null)
            {
                CheckBox chk = (CheckBox)sender;
                if (!chk.IsChecked.HasValue || !chk.IsChecked.Value)
                {
                    scarsLeft.SelectedIndex = 0;
                    scarsRight.SelectedIndex = 0;
                    scarsLeft.IsEnabled = false;
                    scarsRight.IsEnabled = false;
                }
                else
                {
                    scarsLeft.IsEnabled = true;
                    scarsRight.IsEnabled = true;
                }
            }
        }
    }
}

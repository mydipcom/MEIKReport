using MEIKReport.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Xml;

namespace MEIKReport.Views
{
    /// <summary>
    /// Examination Report Page 
    /// </summary>
    public partial class ExaminationReportPage : Window
    {
        private delegate void DoPrintMethod(PrintDialog pdlg, DocumentPaginator paginator);
        //public event CloseWindowHandler closeWindowEvent;
        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "Data";
        private Person person = null;
        private ShortFormReport shortFormReportModel = new ShortFormReport();
        protected MouseHook mouseHook = new MouseHook();

        public ExaminationReportPage()
        {
            InitializeComponent();
        }
        public ExaminationReportPage(object data): this()
        {            
            mouseHook.MouseUp += new System.Windows.Forms.MouseEventHandler(mouseHook_MouseUp);
            //mouseHook.MouseMove += new System.Windows.Forms.MouseEventHandler(mouseHook_MouseMove);
            try { 
                this.person = data as Person;
                if (this.person == null)
                {
                    MessageBox.Show(this, "Please select a patient.");
                    this.Close();
                }
                else
                {                    
                    string dataFile = FindUserReportData(person.ArchiveFolder);
                    if (string.IsNullOrEmpty(dataFile))
                    {
                        dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                    }
                    //string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                    if ("en-US".Equals(App.local))
                    {
                        dataScreenDate.FormatString = "MMMM d, yyyy";                        
                    }
                    else
                    {
                        dataScreenDate.FormatString = "yyyy年MM月dd日";                        
                    }

                    if (File.Exists(dataFile))
                    {                    
                        this.shortFormReportModel = SerializeUtilities.Desrialize<ShortFormReport>(dataFile);
                        //判断是否已经从MEIK生成的DOC文档中导入检查数据，如果之前没有，则查找是否已在本地生成DOC文档，导入数据
                        if (string.IsNullOrEmpty(shortFormReportModel.DataLeftAgeRelated) && string.IsNullOrEmpty(shortFormReportModel.DataRightAgeRelated))
                        {
                            string docFile = FindUserReportWord(person.ArchiveFolder);
                            if (!string.IsNullOrEmpty(docFile))
                            {                                
                                ShortFormReport shortFormReport = WordTools.ReadWordFile(docFile);
                                shortFormReportModel.DataMenstrualCycle = shortFormReport.DataMenstrualCycle;
                                shortFormReportModel.DataLeftMeanElectricalConductivity1 = shortFormReport.DataLeftMeanElectricalConductivity1;
                                shortFormReportModel.DataRightMeanElectricalConductivity1 = shortFormReport.DataRightMeanElectricalConductivity1;
                                shortFormReportModel.DataLeftMeanElectricalConductivity2 = shortFormReport.DataLeftMeanElectricalConductivity2;
                                shortFormReportModel.DataRightMeanElectricalConductivity2 = shortFormReport.DataRightMeanElectricalConductivity2;
                                shortFormReportModel.DataMeanElectricalConductivity3 =shortFormReport.DataMeanElectricalConductivity3;
                                shortFormReportModel.DataLeftMeanElectricalConductivity3 =shortFormReport.DataLeftMeanElectricalConductivity3;
                                shortFormReportModel.DataRightMeanElectricalConductivity3 = shortFormReport.DataRightMeanElectricalConductivity3;
                                shortFormReportModel.DataLeftComparativeElectricalConductivity1 =shortFormReport.DataLeftComparativeElectricalConductivity1;
                                shortFormReportModel.DataLeftComparativeElectricalConductivity2 =shortFormReport.DataLeftComparativeElectricalConductivity2;
                                shortFormReportModel.DataLeftComparativeElectricalConductivity3 =shortFormReport.DataLeftComparativeElectricalConductivity3;
                                shortFormReportModel.DataLeftDivergenceBetweenHistograms1 =shortFormReport.DataLeftDivergenceBetweenHistograms1;
                                shortFormReportModel.DataLeftDivergenceBetweenHistograms2 =shortFormReport.DataLeftDivergenceBetweenHistograms2;
                                shortFormReportModel.DataLeftDivergenceBetweenHistograms3 =shortFormReport.DataLeftDivergenceBetweenHistograms3;
                                shortFormReportModel.DataLeftPhaseElectricalConductivity =shortFormReport.DataLeftPhaseElectricalConductivity;
                                shortFormReportModel.DataRightPhaseElectricalConductivity =shortFormReport.DataRightPhaseElectricalConductivity;
                                shortFormReportModel.DataAgeElectricalConductivityReference =shortFormReport.DataAgeElectricalConductivityReference;
                                shortFormReportModel.DataLeftAgeElectricalConductivity =shortFormReport.DataLeftAgeElectricalConductivity;
                                shortFormReportModel.DataRightAgeElectricalConductivity =shortFormReport.DataRightAgeElectricalConductivity;
                                shortFormReportModel.DataExamConclusion =shortFormReport.DataExamConclusion;
                                shortFormReportModel.DataLeftMammaryGland =shortFormReport.DataLeftMammaryGland;
                                shortFormReportModel.DataLeftAgeRelated =shortFormReport.DataLeftAgeRelated;
                                shortFormReportModel.DataRightMammaryGland =shortFormReport.DataRightMammaryGland;
                                shortFormReportModel.DataRightAgeRelated =shortFormReport.DataRightAgeRelated;
                            }
                        }
                        //this.reportDataGrid.DataContext = this.shortFormReportModel;                                            
                        if (shortFormReportModel.DataSignImg != null)
                        {
                            this.dataSignImg.Source = ImageTools.GetBitmapImage(shortFormReportModel.DataSignImg);
                        }
                        if (shortFormReportModel.DataScreenShotImg != null)
                        {
                            this.btnScreenShot.Content = App.Current.FindResource("ReportContext_170").ToString();// "View Screenshot";
                            this.btnRemoveImg.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            this.btnScreenShot.Content = App.Current.FindResource("ReportContext_175").ToString();// "Capture Screen";
                            this.btnRemoveImg.Visibility = Visibility.Hidden;
                        }                                          
                        
                    }
                    else
                    {
                        string docFile = FindUserReportWord(person.ArchiveFolder);
                        if (!string.IsNullOrEmpty(docFile))
                        {
                            shortFormReportModel = WordTools.ReadWordFile(docFile);
                        }
                        shortFormReportModel.DataUserCode = person.Code;
                        shortFormReportModel.DataName = person.SurName;
                        shortFormReportModel.DataAge = person.Age + "";
                        shortFormReportModel.DataAddress = person.Address;                        
                        if ("en-US".Equals(App.local))
                        {
                            shortFormReportModel.DataScreenDate = DateTime.ParseExact("20" + person.Code.Substring(0, 6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture).ToString("MMMM d, yyyy");                            
                        }
                        else
                        {
                            shortFormReportModel.DataScreenDate = DateTime.ParseExact("20" + person.Code.Substring(0, 6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture).ToString("yyyy年MM月dd日");
                        }
                        

                        bool defaultSign = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                        if (defaultSign)
                        {
                            string imgFile = AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg";
                            if (File.Exists(imgFile))
                            {
                                this.dataSignImg.Source = ImageTools.GetBitmapImage(imgFile);
                                //dataScreenShotImg.Source = GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "/Images/BigIcon.png");
                            }
                        }
                    }
                    this.reportDataGrid.DataContext = this.shortFormReportModel;
                    //以下是添加处理操作员和医生的名字的选择项                    
                    User doctorUser = new User();
                    if (!string.IsNullOrEmpty(shortFormReportModel.DataDoctor))
                    {
                        doctorUser.Name = shortFormReportModel.DataDoctor;
                        doctorUser.License = shortFormReportModel.DataDoctorLicense;
                    }
                    else
                    {
                        if(!string.IsNullOrEmpty(this.person.DoctorName)){
                            doctorUser.Name = this.person.DoctorName;
                            doctorUser.License = this.person.DoctorLicense;
                            shortFormReportModel.DataDoctor = this.person.DoctorName;
                            shortFormReportModel.DataDoctorLicense = this.person.DoctorLicense;
                        }                                                
                    }
                    this.dataDoctor.ItemsSource = App.reportSettingModel.DoctorNames;
                    if(!string.IsNullOrEmpty(doctorUser.Name)){                        
                        for (int i = 0; i < App.reportSettingModel.DoctorNames.Count; i++)
			            {
                            var item=App.reportSettingModel.DoctorNames[i];
                            if (doctorUser.Name.Equals(item.Name) && (string.IsNullOrEmpty(doctorUser.License)==string.IsNullOrEmpty(item.License)||doctorUser.License==item.License))
                            {                                
                                this.dataDoctor.SelectedIndex = i;
                                break;
                            }
			            }
                        //如果没有找到匹配的用户
                        if (this.dataDoctor.SelectedIndex == -1)
                        {
                            App.reportSettingModel.DoctorNames.Add(doctorUser);
                            this.dataDoctor.SelectedIndex = App.reportSettingModel.DoctorNames.Count - 1;
                        }
                    }                                        

                    User techUser = new User();
                    if (!string.IsNullOrEmpty(shortFormReportModel.DataMeikTech))
                    {
                        techUser.Name = shortFormReportModel.DataMeikTech;
                        techUser.License = shortFormReportModel.DataTechLicense;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(this.person.TechName))
                        {
                            techUser.Name = this.person.TechName;
                            techUser.License = this.person.TechLicense;
                            shortFormReportModel.DataMeikTech = this.person.TechName;
                            shortFormReportModel.DataTechLicense = this.person.TechLicense;
                        }
                    }
                    this.dataMeikTech.ItemsSource = App.reportSettingModel.TechNames;
                    if (!string.IsNullOrEmpty(techUser.Name))
                    {
                        for (int i = 0; i < App.reportSettingModel.TechNames.Count; i++)
                        {
                            var item = App.reportSettingModel.TechNames[i];
                            if (techUser.Name.Equals(item.Name) && (string.IsNullOrEmpty(techUser.License) == string.IsNullOrEmpty(item.License) || techUser.License == item.License))
                            {
                                this.dataMeikTech.SelectedIndex = i;
                                break;
                            }
                        }
                        //如果没有找到匹配的用户
                        if (this.dataMeikTech.SelectedIndex == -1)
                        {
                            App.reportSettingModel.TechNames.Add(techUser);
                            this.dataMeikTech.SelectedIndex = App.reportSettingModel.TechNames.Count - 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            //App.opendWin = null;            
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
            catch (Exception)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Owner.Show();
            }
        }

        private void DoPrint(PrintDialog pdlg, DocumentPaginator paginator)
        {
            pdlg.PrintDocument(paginator, "Examination Report Document");
        }

        private void previewBtn_Click(object sender, RoutedEventArgs e)
        {
            try {
                LoadDataModel();
                var reportModel=CloneReportModel();
                
                PrintPreviewWindow previewWnd = new PrintPreviewWindow("Views/ExaminationReportDocument.xaml", true, reportModel);                
                previewWnd.Owner = this;
                previewWnd.ShowInTaskbar = false;
                previewWnd.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void printBtn_Click(object sender, RoutedEventArgs e)
        {
            try {
                LoadDataModel();
                PrintDialog pdlg = new PrintDialog();
                if (pdlg.ShowDialog() == true)
                {
                    FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender("Views/ExaminationReportDocument.xaml", shortFormReportModel);
                    FixedDocument fixedDoc = new FixedDocument();//创建一个文档
                    if ("Letter".Equals(App.reportSettingModel.PrintPaper, StringComparison.OrdinalIgnoreCase))
                    {                        
                        fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.5, 96 * 11);
                    }
                    else if ("A4".Equals(App.reportSettingModel.PrintPaper, StringComparison.OrdinalIgnoreCase))
                    {
                        fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.27, 96 * 11.69);
                    }

                    PageContent pageContent = new PageContent();
                    ((IAddChild)pageContent).AddChild(page);
                    fixedDoc.Pages.Add(pageContent);//将对象加入到当前文档中
                    Dispatcher.BeginInvoke(new DoPrintMethod(DoPrint), DispatcherPriority.ApplicationIdle, pdlg, fixedDoc.DocumentPaginator);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void btnOpenDiagn_Click(object sender, RoutedEventArgs e)
        {            
            try
            {
                IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
                //如果主窗体不存在
                if (mainWinHwnd == IntPtr.Zero)
                {
                    IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strDiagnostics);
                    Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                }                
                WinMinimized();
                //var userListWin = this.Owner as UserList;
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "System Exception: " + ex.Message);
            }
        }        

        private void save_Click(object sender, RoutedEventArgs e)
        {
            SaveReport(null);
        }

        private void btnSignature_Click(object sender, RoutedEventArgs e)
        {
            SignatureBox signBox = new SignatureBox();
            signBox.Owner = this;
            //signBox.callbackMethod = ShowSignature;
            signBox.ShowDialog();
        }

        /// <summary>
        /// 保存报表数据到文件
        /// </summary>
        private void SaveReport(string otherDataFolder)
        {
            try
            {
                //保存到患者档案目录
                string datafile = null;
                if (!string.IsNullOrEmpty(otherDataFolder))
                {
                    datafile = otherDataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                }
                else
                {
                    //保存到程序自身的Data目录
                    if (!Directory.Exists(dataFolder))
                    {
                        Directory.CreateDirectory(dataFolder);
                        FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                        FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                    }
                    datafile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                }                
                LoadDataModel();
                
                if (string.IsNullOrEmpty(shortFormReportModel.DataSignDate))
                {
                    if ("en-US".Equals(App.local)){
                        shortFormReportModel.DataSignDate = DateTime.Today.ToString("MMMM d, yyyy");
                    }
                    else{
                        shortFormReportModel.DataSignDate = DateTime.Today.ToString("yyyy年MM月dd日");                        
                    }
                }
                SerializeUtilities.Serialize<ShortFormReport>(shortFormReportModel, datafile);
                File.Copy(datafile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat", true);
                var reportModel = CloneReportModel();
                //Save PDF file
                string lfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_LF.pdf";
                string lfReportTempl = "Views/ExaminationReportDocument.xaml";
                ExportPDF(lfReportTempl, lfPdfFile, reportModel);
                File.Copy(lfPdfFile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_LF.pdf", true);
                string sfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_SF.pdf";
                string sfReportTempl = "Views/SummaryReportDocument.xaml";                
                if (shortFormReportModel.DataScreenShotImg != null)
                {
                    sfReportTempl = "Views/SummaryReportImageDocument.xaml";                                        
                }
                ExportPDF(sfReportTempl, sfPdfFile, reportModel);
                File.Copy(sfPdfFile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_SF.pdf", true);

                MessageBox.Show(this, App.Current.FindResource("Message_2").ToString());
            }
            catch (Exception ex)
            {
                FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                MessageBox.Show(this, App.Current.FindResource("Message_3").ToString() + ex.Message);
            }
        }

        /// <summary>
        /// 加载报表数据到模型对象
        /// </summary>
        private void LoadDataModel()
        {
            shortFormReportModel.DataUserCode = this.dataUserCode.Text;
            shortFormReportModel.DataAge = this.dataAge.Text;
            shortFormReportModel.DataBiRadsCategory = this.dataBiRadsCategory.SelectedIndex.ToString();
            shortFormReportModel.DataRecommendation = this.dataRecommendation.SelectedIndex.ToString();
            shortFormReportModel.DataComments = this.dataComments.Text;
            shortFormReportModel.DataConclusion = this.dataConclusion.SelectedIndex.ToString();
            shortFormReportModel.DataFurtherExam = this.dataFurtherExam.SelectedIndex.ToString();
            //shortFormReportModel.DataConclusion2 = this.dataConclusion2.Text;
            if (this.dataDoctor.SelectedItem != null)
            {
                var doctor = this.dataDoctor.SelectedItem as User;
                shortFormReportModel.DataDoctor = doctor.Name;
                shortFormReportModel.DataDoctorLicense = doctor.License;                
            }
            if (this.dataMeikTech.SelectedItem != null)
            {
                var technician = this.dataMeikTech.SelectedItem as User;
                shortFormReportModel.DataMeikTech = technician.Name;
                shortFormReportModel.DataTechLicense = technician.License;
            }
            shortFormReportModel.DataLeftLocation = this.dataLeftLocation.Text;
            shortFormReportModel.DataLeftSize = this.dataLeftSize.Text;
            
            shortFormReportModel.DataName = this.dataName.Text;
                        
            shortFormReportModel.DataRightLocation = this.dataRightLocation.Text;
            shortFormReportModel.DataRightSize = this.dataRightSize.Text;
            shortFormReportModel.DataScreenDate = this.dataScreenDate.Text;
            shortFormReportModel.DataScreenLocation = this.dataScreenLocation.Text;
            
            shortFormReportModel.DataAddress = this.dataAddress.Text;
            shortFormReportModel.DataGender = this.dataGender.SelectedIndex.ToString();
            shortFormReportModel.DataHealthCard = this.dataHealthCard.Text;
            shortFormReportModel.DataWeight = this.dataWeight.Text;
            shortFormReportModel.DataWeightUnit = this.dataWeightUnit.SelectedIndex.ToString();
            shortFormReportModel.DataMenstrualCycle = this.dataMenstrualCycle.SelectedIndex.ToString();
            shortFormReportModel.DataHormones = this.dataHormones.Text;
            shortFormReportModel.DataSkinAffections = this.dataSkinAffections.SelectedIndex.ToString();
            shortFormReportModel.DataPertinentHistory = this.dataPertinentHistory.SelectedIndex.ToString();
            shortFormReportModel.DataPertinentHistory1 = this.dataPertinentHistory1.Text;
            shortFormReportModel.DataMotherUltra = this.dataMotherUltra.SelectedIndex.ToString();
            shortFormReportModel.DataLeftBreast = this.dataLeftBreast.Text;
            shortFormReportModel.DataRightBreast = this.dataRightBreast.Text;
            shortFormReportModel.DataLeftPalpableMass = this.dataLeftPalpableMass.SelectedIndex.ToString();
            shortFormReportModel.DataRightPalpableMass = this.dataRightPalpableMass.SelectedIndex.ToString();
            shortFormReportModel.DataLeftChangesOfElectricalConductivity = this.dataLeftChangesOfElectricalConductivity.SelectedIndex.ToString();
            shortFormReportModel.DataRightChangesOfElectricalConductivity = this.dataRightChangesOfElectricalConductivity.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMammaryStruct = this.dataLeftMammaryStruct.SelectedIndex.ToString();
            shortFormReportModel.DataRightMammaryStruct = this.dataRightMammaryStruct.SelectedIndex.ToString();
            shortFormReportModel.DataLeftLactiferousSinusZone = this.dataLeftLactiferousSinusZone.SelectedIndex.ToString();
            shortFormReportModel.DataRightLactiferousSinusZone = this.dataRightLactiferousSinusZone.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMammaryContour = this.dataLeftMammaryContour.SelectedIndex.ToString();
            shortFormReportModel.DataRightMammaryContour = this.dataRightMammaryContour.SelectedIndex.ToString();

            shortFormReportModel.DataLeftNumber = this.dataLeftNumber.Text;
            shortFormReportModel.DataRightNumber = this.dataRightNumber.Text;
            shortFormReportModel.DataLeftShape = this.dataLeftShape.SelectedIndex.ToString();
            shortFormReportModel.DataRightShape = this.dataRightShape.SelectedIndex.ToString();
            shortFormReportModel.DataLeftContourAroundFocus = this.dataLeftContourAroundFocus.SelectedIndex.ToString();
            shortFormReportModel.DataRightContourAroundFocus = this.dataRightContourAroundFocus.SelectedIndex.ToString();
            shortFormReportModel.DataLeftInternalElectricalStructure = this.dataLeftInternalElectricalStructure.SelectedIndex.ToString();
            shortFormReportModel.DataRightInternalElectricalStructure = this.dataRightInternalElectricalStructure.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSurroundingTissues = this.dataLeftSurroundingTissues.SelectedIndex.ToString();
            shortFormReportModel.DataRightSurroundingTissues = this.dataRightSurroundingTissues.SelectedIndex.ToString();
            shortFormReportModel.DataLeftOncomarkerHighlightBenignChanges = this.dataLeftOncomarkerHighlightBenignChanges.SelectedIndex.ToString();
            shortFormReportModel.DataRightOncomarkerHighlightBenignChanges = this.dataRightOncomarkerHighlightBenignChanges.SelectedIndex.ToString();
            shortFormReportModel.DataLeftOncomarkerHighlightSuspiciousChanges = this.dataLeftOncomarkerHighlightSuspiciousChanges.SelectedIndex.ToString();
            shortFormReportModel.DataRightOncomarkerHighlightSuspiciousChanges = this.dataRightOncomarkerHighlightSuspiciousChanges.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMeanElectricalConductivity1 = this.dataLeftMeanElectricalConductivity1.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity1N1 = this.dataLeftMeanElectricalConductivity1N1.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity1N2 = this.dataLeftMeanElectricalConductivity1N2.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity1 = this.dataRightMeanElectricalConductivity1.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity1N1 = this.dataRightMeanElectricalConductivity1N1.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity1N2 = this.dataRightMeanElectricalConductivity1N2.Text;

            shortFormReportModel.DataLeftMeanElectricalConductivity2 = this.dataLeftMeanElectricalConductivity2.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity2N1 = this.dataLeftMeanElectricalConductivity2N1.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity2N2 = this.dataLeftMeanElectricalConductivity2N2.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity2 = this.dataRightMeanElectricalConductivity2.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity2N1 = this.dataRightMeanElectricalConductivity2N1.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity2N2 = this.dataRightMeanElectricalConductivity2N2.Text;

            shortFormReportModel.DataMeanElectricalConductivity3 = this.dataMeanElectricalConductivity3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMeanElectricalConductivity3 = this.dataLeftMeanElectricalConductivity3.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity3N1 = this.dataLeftMeanElectricalConductivity3N1.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity3N2 = this.dataLeftMeanElectricalConductivity3N2.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3 = this.dataRightMeanElectricalConductivity3.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3N1 = this.dataRightMeanElectricalConductivity3N1.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3N2 = this.dataRightMeanElectricalConductivity3N2.Text;

            //shortFormReportModel.DataComparativeElectricalConductivityReference1 = this.dataComparativeElectricalConductivityReference1.Text;
            shortFormReportModel.DataLeftComparativeElectricalConductivity1 = this.dataLeftComparativeElectricalConductivity1.Text;
            //shortFormReportModel.DataRightComparativeElectricalConductivity1 = this.dataRightComparativeElectricalConductivity1.Text;
            //shortFormReportModel.DataComparativeElectricalConductivityReference2 = this.dataComparativeElectricalConductivityReference2.Text;
            shortFormReportModel.DataLeftComparativeElectricalConductivity2 = this.dataLeftComparativeElectricalConductivity2.Text;
            //shortFormReportModel.DataRightComparativeElectricalConductivity2 = this.dataRightComparativeElectricalConductivity2.Text;
            shortFormReportModel.DataComparativeElectricalConductivity3 = this.dataComparativeElectricalConductivity3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftComparativeElectricalConductivity3 = this.dataLeftComparativeElectricalConductivity3.Text;
            //shortFormReportModel.DataRightComparativeElectricalConductivity3 = this.dataRightComparativeElectricalConductivity3.Text;

            //shortFormReportModel.DataDivergenceBetweenHistogramsReference1 = this.dataDivergenceBetweenHistogramsReference1.Text;
            shortFormReportModel.DataLeftDivergenceBetweenHistograms1 = this.dataLeftDivergenceBetweenHistograms1.Text;
            //shortFormReportModel.DataRightDivergenceBetweenHistograms1 = this.dataRightDivergenceBetweenHistograms1.Text;
            //shortFormReportModel.DataDivergenceBetweenHistogramsReference2 = this.dataDivergenceBetweenHistogramsReference2.Text;
            shortFormReportModel.DataLeftDivergenceBetweenHistograms2 = this.dataLeftDivergenceBetweenHistograms2.Text;
            //shortFormReportModel.DataRightDivergenceBetweenHistograms2 = this.dataRightDivergenceBetweenHistograms2.Text;
            shortFormReportModel.DataDivergenceBetweenHistograms3 = this.dataDivergenceBetweenHistograms3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftDivergenceBetweenHistograms3 = this.dataLeftDivergenceBetweenHistograms3.Text;
            //shortFormReportModel.DataRightDivergenceBetweenHistograms3 = this.dataRightDivergenceBetweenHistograms3.Text;

            shortFormReportModel.DataLeftComparisonWithNorm = this.dataLeftComparisonWithNorm.Text;
            shortFormReportModel.DataRightComparisonWithNorm = this.dataRightComparisonWithNorm.Text;

            //shortFormReportModel.DataPhaseElectricalConductivityReference = this.dataPhaseElectricalConductivityReference.Text;
            shortFormReportModel.DataLeftPhaseElectricalConductivity = this.dataLeftPhaseElectricalConductivity.Text;
            shortFormReportModel.DataRightPhaseElectricalConductivity = this.dataRightPhaseElectricalConductivity.Text;

            shortFormReportModel.DataAgeElectricalConductivityReference = this.dataAgeElectricalConductivityReference.Text;
            shortFormReportModel.DataLeftAgeElectricalConductivity = this.dataLeftAgeElectricalConductivity.SelectedIndex.ToString();
            shortFormReportModel.DataRightAgeElectricalConductivity = this.dataRightAgeElectricalConductivity.SelectedIndex.ToString();
            shortFormReportModel.DataAgeValueOfEC = this.dataAgeValueOfEC.SelectedIndex.ToString();

            shortFormReportModel.DataExamConclusion = this.dataExamConclusion.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMammaryGland = this.dataLeftMammaryGland.SelectedIndex.ToString();
            shortFormReportModel.DataLeftAgeRelated = this.dataLeftAgeRelated.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMeanECOfLesion = this.dataLeftMeanECOfLesion.Text;
            shortFormReportModel.DataLeftFindings = this.dataLeftFindings.SelectedIndex.ToString();
            shortFormReportModel.DataLeftMammaryGlandResult = this.dataLeftMammaryGlandResult.SelectedIndex.ToString();

            shortFormReportModel.DataRightMammaryGland = this.dataRightMammaryGland.SelectedIndex.ToString();
            shortFormReportModel.DataRightAgeRelated = this.dataRightAgeRelated.SelectedIndex.ToString();
            shortFormReportModel.DataRightMeanECOfLesion = this.dataRightMeanECOfLesion.Text;
            shortFormReportModel.DataRightFindings = this.dataRightFindings.SelectedIndex.ToString();
            shortFormReportModel.DataRightMammaryGlandResult = this.dataRightMammaryGlandResult.SelectedIndex.ToString();

            shortFormReportModel.DataTotalPts = this.dataTotalPts.SelectedIndex.ToString();
            shortFormReportModel.DataPoint = this.dataPoint.SelectedIndex.ToString();
            
            var signImg = this.dataSignImg.Source as BitmapImage;
            if (signImg != null)
            {
                var stream = signImg.StreamSource;
                stream.Position = 0;
                byte[] buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                stream.Flush();
                //stream.Close();
                shortFormReportModel.DataSignImg = buffer;
            }
            
            
        }

        /// <summary>
        /// 克隆数据模型对象
        /// </summary>
        private ShortFormReport CloneReportModel()
        {
            ShortFormReport reportModel = shortFormReportModel.Clone();            
            reportModel = shortFormReportModel.Clone();
            reportModel.DataGender = this.dataGender.Text;
            reportModel.DataBiRadsCategory = this.dataBiRadsCategory.Text;
            reportModel.DataRecommendation = this.dataRecommendation.Text;
            reportModel.DataConclusion = this.dataConclusion.Text;
            reportModel.DataFurtherExam = this.dataFurtherExam.Text;
            reportModel.DataWeightUnit = this.dataWeightUnit.Text;
            reportModel.DataMenstrualCycle = this.dataMenstrualCycle.Text;
            reportModel.DataSkinAffections = this.dataSkinAffections.Text;
            reportModel.DataPertinentHistory = this.dataPertinentHistory.Text;
            reportModel.DataMotherUltra = this.dataMotherUltra.Text;
            reportModel.DataLeftPalpableMass = this.dataLeftPalpableMass.Text;
            reportModel.DataRightPalpableMass = this.dataRightPalpableMass.Text;
            reportModel.DataLeftChangesOfElectricalConductivity = this.dataLeftChangesOfElectricalConductivity.Text;
            reportModel.DataRightChangesOfElectricalConductivity = this.dataRightChangesOfElectricalConductivity.Text;
            reportModel.DataLeftMammaryStruct = this.dataLeftMammaryStruct.Text;
            reportModel.DataRightMammaryStruct = this.dataRightMammaryStruct.Text;
            reportModel.DataLeftLactiferousSinusZone = this.dataLeftLactiferousSinusZone.Text;
            reportModel.DataRightLactiferousSinusZone = this.dataRightLactiferousSinusZone.Text;
            reportModel.DataLeftMammaryContour = this.dataLeftMammaryContour.Text;
            reportModel.DataRightMammaryContour = this.dataRightMammaryContour.Text;
            reportModel.DataLeftShape = this.dataLeftShape.Text;
            reportModel.DataRightShape = this.dataRightShape.Text;
            reportModel.DataLeftContourAroundFocus = this.dataLeftContourAroundFocus.Text;
            reportModel.DataRightContourAroundFocus = this.dataRightContourAroundFocus.Text;
            reportModel.DataLeftInternalElectricalStructure = this.dataLeftInternalElectricalStructure.Text;
            reportModel.DataRightInternalElectricalStructure = this.dataRightInternalElectricalStructure.Text;
            reportModel.DataLeftSurroundingTissues = this.dataLeftSurroundingTissues.Text;
            reportModel.DataRightSurroundingTissues = this.dataRightSurroundingTissues.Text;
            reportModel.DataLeftOncomarkerHighlightBenignChanges = this.dataLeftOncomarkerHighlightBenignChanges.Text;
            reportModel.DataRightOncomarkerHighlightBenignChanges = this.dataRightOncomarkerHighlightBenignChanges.Text;
            reportModel.DataLeftOncomarkerHighlightSuspiciousChanges = this.dataLeftOncomarkerHighlightSuspiciousChanges.Text;
            reportModel.DataRightOncomarkerHighlightSuspiciousChanges = this.dataRightOncomarkerHighlightSuspiciousChanges.Text;
            reportModel.DataMeanElectricalConductivity3 = this.dataMeanElectricalConductivity3.Text;
            reportModel.DataComparativeElectricalConductivity3 = this.dataComparativeElectricalConductivity3.Text;
            reportModel.DataDivergenceBetweenHistograms3 = this.dataDivergenceBetweenHistograms3.Text;
            reportModel.DataLeftAgeElectricalConductivity = this.dataLeftAgeElectricalConductivity.Text;
            reportModel.DataRightAgeElectricalConductivity = this.dataRightAgeElectricalConductivity.Text;
            reportModel.DataAgeValueOfEC = this.dataAgeValueOfEC.Text;
            reportModel.DataExamConclusion = this.dataExamConclusion.Text;
            reportModel.DataLeftMammaryGland = this.dataLeftMammaryGland.Text;
            reportModel.DataLeftAgeRelated = this.dataLeftAgeRelated.Text;
            reportModel.DataLeftFindings = this.dataLeftFindings.Text;
            reportModel.DataLeftMammaryGlandResult = this.dataLeftMammaryGlandResult.Text;
            reportModel.DataRightMammaryGland = this.dataRightMammaryGland.Text;
            reportModel.DataRightAgeRelated = this.dataRightAgeRelated.Text;
            reportModel.DataRightFindings = this.dataRightFindings.Text;
            reportModel.DataRightMammaryGlandResult = this.dataRightMammaryGlandResult.Text;
            reportModel.DataTotalPts = this.dataTotalPts.Text;
            reportModel.DataPoint = this.dataPoint.Text;

            return reportModel;            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_4").ToString(), "Save Report", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                SaveReport(null);
            }
        }

        public void ShowSignature(Object obj)
        {
            dataSignImg.Source = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "Signature" + System.IO.Path.DirectorySeparatorChar + "temp.jpg");
        }

        private void savePdfBtn_Click(object sender, RoutedEventArgs e)
        {
            try { 
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                    FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                    FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                }
                LoadDataModel();                
                string xpsFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                //string userTempPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                //string xpsFile = userTempPath + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                if (File.Exists(xpsFile))
                {
                    File.Delete(xpsFile);
                }
                var reportModel = CloneReportModel();
                FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender("Views/ExaminationReportDocument.xaml", reportModel);
                XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
                //将flow document写入基于内存的xps document中去
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
                writer.Write(page);
                xpsDocument.Close();
                var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "pdf|*.pdf" };
                if (dlg.ShowDialog(this) == true)
                {
                    //string pdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_LF.pdf";
                    //if (File.Exists(pdfFile))
                    //{
                    //    File.Delete(pdfFile);
                    //}
                    //PDFTools.SavePDFFile(xpsFile, pdfFile);
                    PDFTools.SavePDFFile(xpsFile, dlg.FileName);
                    MessageBox.Show(this, App.Current.FindResource("Message_5").ToString());
                }
            }
            catch (Exception ex)
            {
                FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                MessageBox.Show(this, ex.Message);
            }
        }

        private void ExportPDF(string reportTempl,string pdfFile,ShortFormReport reportModel)
        {
            string xpsFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
            //string userTempPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            //string xpsFile = userTempPath + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
            if (File.Exists(xpsFile))
            {
                File.Delete(xpsFile);
            }
            
            FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender(reportTempl, reportModel);
            
            XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
            //将flow document写入基于内存的xps document中去
            XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
            writer.Write(page);
            xpsDocument.Close();
            if (File.Exists(pdfFile))
            {
                File.Delete(pdfFile);
            }
            PDFTools.SavePDFFile(xpsFile, pdfFile);            
        }        

        private void saveAsBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            string folderName = folderBrowserDialog.SelectedPath;
            if (!string.IsNullOrEmpty(folderName))
            {
                SaveReport(folderName);
                MessageBox.Show(this, App.Current.FindResource("Message_6").ToString());
            }

        }

        private string FindUserReportData(string folderName)
        {
            //遍历指定文件夹下所有文件
            DirectoryInfo theFolder = new DirectoryInfo(folderName);
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            List<DirectoryInfo> list = dirInfo.ToList();
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
                        if ((person.Code + ".dat").Equals(NextFile.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            return NextFile.FullName;
                        }
                    }
                }
                catch (Exception) { }
            }
            return null;
        }

        private string FindUserReportWord(string folderName)
        {
            //遍历指定文件夹下所有文件
            DirectoryInfo theFolder = new DirectoryInfo(folderName);
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            List<DirectoryInfo> list = dirInfo.ToList();
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
                        if ((person.Code + ".doc").Equals(NextFile.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            return NextFile.FullName;
                        }
                    }
                }
                catch (Exception) { }
            }
            return null;
        }

        private void btnScreenShot_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (shortFormReportModel.DataScreenShotImg == null)
                {
                    this.WindowState = WindowState.Minimized;
                    App.opendWin = this;
                    ScreenCapture screenCaptureWin = new ScreenCapture(this.person);
                    screenCaptureWin.callbackMethod = LoadScreenShot;
                    screenCaptureWin.ShowDialog();
                }
                else
                {
                    ViewImagePage viewImagePage = new ViewImagePage(shortFormReportModel.DataScreenShotImg);
                    viewImagePage.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "System Exception: " + ex.Message);
            } 
        }

        private void LoadScreenShot(Object imgFileName)
        {
            var screenShotImg = ImageTools.GetBitmapImage(imgFileName as string);
            if (screenShotImg != null)
            {
                var stream = screenShotImg.StreamSource;
                stream.Position = 0;
                byte[] buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                stream.Flush();
                shortFormReportModel.DataScreenShotImg = buffer;
                this.btnScreenShot.Content = App.Current.FindResource("ReportContext_170").ToString();                
                this.btnRemoveImg.Visibility = Visibility.Visible;
            }
           
            //dataScreenShotImg.Source = ImageTools.GetBitmapImage(imgFileName as string);
        }

        private void btnRemoveImg_Click(object sender, RoutedEventArgs e)
        {
            this.btnScreenShot.Content = App.Current.FindResource("ReportContext_175").ToString();            
            shortFormReportModel.DataScreenShotImg = null;
            this.btnRemoveImg.Visibility = Visibility.Hidden;
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            //判断是否已经从MEIK生成的DOC文档中导入检查数据，如果之前没有，则查找是否已在本地生成DOC文档，导入数据            
            string docFile = FindUserReportWord(person.ArchiveFolder);
            if (!string.IsNullOrEmpty(docFile))
            {
                ShortFormReport shortFormReport = WordTools.ReadWordFile(docFile);
                shortFormReportModel.DataMenstrualCycle = shortFormReport.DataMenstrualCycle;
                shortFormReportModel.DataLeftMeanElectricalConductivity1 = shortFormReport.DataLeftMeanElectricalConductivity1;
                shortFormReportModel.DataRightMeanElectricalConductivity1 = shortFormReport.DataRightMeanElectricalConductivity1;
                shortFormReportModel.DataLeftMeanElectricalConductivity2 = shortFormReport.DataLeftMeanElectricalConductivity2;
                shortFormReportModel.DataRightMeanElectricalConductivity2 = shortFormReport.DataRightMeanElectricalConductivity2;
                shortFormReportModel.DataMeanElectricalConductivity3 = shortFormReport.DataMeanElectricalConductivity3;
                shortFormReportModel.DataLeftMeanElectricalConductivity3 = shortFormReport.DataLeftMeanElectricalConductivity3;
                shortFormReportModel.DataRightMeanElectricalConductivity3 = shortFormReport.DataRightMeanElectricalConductivity3;
                shortFormReportModel.DataLeftComparativeElectricalConductivity1 = shortFormReport.DataLeftComparativeElectricalConductivity1;
                shortFormReportModel.DataLeftComparativeElectricalConductivity2 = shortFormReport.DataLeftComparativeElectricalConductivity2;
                shortFormReportModel.DataLeftComparativeElectricalConductivity3 = shortFormReport.DataLeftComparativeElectricalConductivity3;
                shortFormReportModel.DataLeftDivergenceBetweenHistograms1 = shortFormReport.DataLeftDivergenceBetweenHistograms1;
                shortFormReportModel.DataLeftDivergenceBetweenHistograms2 = shortFormReport.DataLeftDivergenceBetweenHistograms2;
                shortFormReportModel.DataLeftDivergenceBetweenHistograms3 = shortFormReport.DataLeftDivergenceBetweenHistograms3;
                shortFormReportModel.DataLeftPhaseElectricalConductivity = shortFormReport.DataLeftPhaseElectricalConductivity;
                shortFormReportModel.DataRightPhaseElectricalConductivity = shortFormReport.DataRightPhaseElectricalConductivity;
                shortFormReportModel.DataAgeElectricalConductivityReference = shortFormReport.DataAgeElectricalConductivityReference;
                shortFormReportModel.DataLeftAgeElectricalConductivity = shortFormReport.DataLeftAgeElectricalConductivity;
                shortFormReportModel.DataRightAgeElectricalConductivity = shortFormReport.DataRightAgeElectricalConductivity;
                shortFormReportModel.DataExamConclusion = shortFormReport.DataExamConclusion;
                shortFormReportModel.DataLeftMammaryGland = shortFormReport.DataLeftMammaryGland;
                shortFormReportModel.DataLeftAgeRelated = shortFormReport.DataLeftAgeRelated;
                shortFormReportModel.DataRightMammaryGland = shortFormReport.DataRightMammaryGland;
                shortFormReportModel.DataRightAgeRelated = shortFormReport.DataRightAgeRelated;
                MessageBox.Show(this, App.Current.FindResource("Message_27").ToString());
            }
            else
            {
                MessageBox.Show(this,App.Current.FindResource("Message_26").ToString());
            }
        }


        /// <summary>
        /// 窗口大小状态变化时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_StateChanged(object sender, EventArgs e)
        {
            try
            {
                var mainWin = this.Owner.Owner as MainWindow;
                if (this.WindowState == WindowState.Maximized)
                {
                    this.StopMouseHook();
                }
                if (this.WindowState == WindowState.Minimized)
                {
                    WinMinimized();                    
                }
                else
                {
                    this.StopMouseHook();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 最小化窗体以便显示MEIK程序窗体
        /// </summary>
        private void WinMinimized()
        {
            App.opendWin = this;
            this.WindowState = WindowState.Minimized;
            //this.WindowStartupLocation = WindowStartupLocation.Manual;//设置可手动指定窗体位置                
            int left = (int)(System.Windows.SystemParameters.PrimaryScreenWidth - 156);
            IntPtr winHandle = new WindowInteropHelper(this).Handle;
            Win32Api.MoveWindow(winHandle, left, 0, 0, 0, false);
            IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
            //如果主窗体存在
            if (mainWinHwnd != IntPtr.Zero)
            {
                this.StartMouseHook();
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
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                IntPtr buttonHandle = Win32Api.WindowFromPoint(e.X, e.Y);
                IntPtr winHandle = Win32Api.GetParent(buttonHandle);
                var owner = this.Owner.Owner as MainWindow;
                if (Win32Api.GetParent(winHandle) == owner.AppProc.MainWindowHandle)
                {
                    StringBuilder winText = new StringBuilder(512);
                    Win32Api.GetWindowText(buttonHandle, winText, winText.Capacity);
                    if (App.strExit.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
                        //如果主窗体存在
                        if (mainWinHwnd != IntPtr.Zero)
                        {
                            int WM_SYSCOMMAND = 0x0112;
                            int SC_CLOSE = 0xF060;
                            Win32Api.SendMessage(mainWinHwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
                        }
                        //this.Visibility = Visibility.Visible;
                        this.WindowState = WindowState.Normal;
                        //this.StopMouseHook();
                    }
                    else if (App.strStart.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))//点击开始设备按钮
                    {
                        string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                        var screenFolderPath = OperateIniFile.ReadIniData("Base", "Patients base", "", meikiniFile);
                        if (App.countDictionary.ContainsKey(screenFolderPath))
                        {
                            List<long> ticks = App.countDictionary[screenFolderPath];
                            ticks.Add(DateTime.Now.Ticks);
                        }
                        else
                        {
                            List<long> ticks = new List<long>();
                            ticks.Add(DateTime.Now.Ticks);
                            App.countDictionary.Add(screenFolderPath, ticks);
                        } 
                        //序列化统计字典到文件
                        string countFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "sc.data";
                        SerializeUtilities.Serialize<SortedDictionary<string, List<long>>>(App.countDictionary, countFile);
                    }
                }
            }
        }

        /// <summary>
        /// 鼠标移动的钩子回调方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void mouseHook_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {                                     
            StringBuilder winText = new StringBuilder(512);
            try
            {
                IntPtr exitButtonHandle = Win32Api.WindowFromPoint(e.X, e.Y);  
                Win32Api.GetWindowText(exitButtonHandle, winText, winText.Capacity);
            }
            catch(Exception){}
            //如果鼠标移动到退出按钮上
            if (App.strExit.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                this.Visibility = Visibility.Hidden;
            }
            else
            {
                this.Visibility = Visibility.Visible;
            }                            
        }
               
    }
}

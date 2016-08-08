﻿using MEIKReport.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
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

                    if (File.Exists(dataFile))
                    {
                        try
                        {
                            this.shortFormReportModel = SerializeUtilities.Desrialize<ShortFormReport>(dataFile);
                        }
                        catch (Exception exec1) {                                                        
                        }                        
                                                                  
                        if (shortFormReportModel.DataSignImg != null)
                        {
                            this.dataSignImg.Source = ImageTools.GetBitmapImage(shortFormReportModel.DataSignImg);
                        }                                                             
                        
                    }
                    else
                    {                        
                        shortFormReportModel.DataUserCode = person.Code;                                                                        
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
                    if ("en-US".Equals(App.local))
                    {
                        shortFormReportModel.DataScreenDate = DateTime.ParseExact("20" + person.Code.Substring(0, 6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture).ToString("MM/dd/yyyy");
                    }
                    else
                    {
                        shortFormReportModel.DataScreenDate = DateTime.ParseExact("20" + person.Code.Substring(0, 6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture).ToString("yyyy年MM月dd日");
                    }
                    shortFormReportModel.DataClientNum = person.ClientNumber;
                    shortFormReportModel.DataName = person.SurName;
                    shortFormReportModel.DataAge = person.Age + "";
                    shortFormReportModel.DataAddress = person.Address;
                    shortFormReportModel.DataHeight = person.Height;
                    shortFormReportModel.DataWeight = person.Weight;
                    shortFormReportModel.DataMobile = person.Mobile;
                    shortFormReportModel.DataEmail = person.Email;
                    shortFormReportModel.DataScreenLocation = person.ScreenVenue;
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
                    //this.dataMeikTech.ItemsSource = App.reportSettingModel.TechNames;
                    //if (!string.IsNullOrEmpty(techUser.Name))
                    //{
                    //    for (int i = 0; i < App.reportSettingModel.TechNames.Count; i++)
                    //    {
                    //        var item = App.reportSettingModel.TechNames[i];
                    //        if (techUser.Name.Equals(item.Name) && (string.IsNullOrEmpty(techUser.License) == string.IsNullOrEmpty(item.License) || techUser.License == item.License))
                    //        {
                    //            this.dataMeikTech.SelectedIndex = i;
                    //            break;
                    //        }
                    //    }
                    //    //如果没有找到匹配的用户
                    //    if (this.dataMeikTech.SelectedIndex == -1)
                    //    {
                    //        App.reportSettingModel.TechNames.Add(techUser);
                    //        this.dataMeikTech.SelectedIndex = App.reportSettingModel.TechNames.Count - 1;
                    //    }
                    //}
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

                PrintPreviewWindow previewWnd = new PrintPreviewWindow("Views/ExaminationReportFlow.xaml", false, reportModel);                
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
                    FlowDocument page = (FlowDocument)PrintPreviewWindow.LoadFlowDocumentAndRender("Views/ExaminationReportFlow.xaml", shortFormReportModel);
                    //FixedDocument fixedDoc = new FixedDocument();//创建一个文档
                    //if ("Letter".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
                    //{
                    //    fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.5, 96 * 11);
                    //}
                    //else if ("A4".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
                    //{
                    //    fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.27, 96 * 11.69);
                    //}
                    
                    //PageContent pageContent = new PageContent(); 
                    //((IAddChild)pageContent).AddChild(page);
                    //fixedDoc.Pages.Add(pageContent);//将对象加入到当前文档中
                    //Dispatcher.BeginInvoke(new DoPrintMethod(DoPrint), DispatcherPriority.ApplicationIdle, pdlg, fixedDoc.DocumentPaginator);
                    Dispatcher.BeginInvoke(new DoPrintMethod(DoPrint), DispatcherPriority.ApplicationIdle, pdlg, ((IDocumentPaginatorSource)page).DocumentPaginator);
                    
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

        private void generatePoints()
        {
            //計算診斷分數
            int leftBreast1 = 0;
            int leftBreast2 = 0;
            int leftBreast3 = 0;
            int rightBreast1 = 0;
            int rightBreast2 = 0;
            int rightBreast3 = 0;
            bool isEmpty = false;
            int divergenceBetweenHistograms = 0;
            if (!string.IsNullOrEmpty(dataLeftDivergenceBetweenHistograms1.Text))
            {
                double val = Convert.ToDouble(dataLeftDivergenceBetweenHistograms1.Text);
                divergenceBetweenHistograms = val < 20 ? 0 : val < 30 ? 1 : val < 40 ? 2 : 3;
            }
            else if (!string.IsNullOrEmpty(dataLeftDivergenceBetweenHistograms2.Text))
            {
                double val = Convert.ToDouble(dataLeftDivergenceBetweenHistograms2.Text);
                divergenceBetweenHistograms = val < 20 ? 0 : val < 30 ? 1 : val < 40 ? 2 : 3;
            }
            else if (!string.IsNullOrEmpty(dataLeftDivergenceBetweenHistograms3.Text))
            {
                double val = Convert.ToDouble(dataLeftDivergenceBetweenHistograms3.Text);
                divergenceBetweenHistograms = val < 20 ? 0 : val < 30 ? 1 : val < 40 ? 2 : 3;
            }
            else
            {
                isEmpty = true;
            }

            if (dataLeftLocation.SelectedIndex > 0)
            {
                leftBreast1 += dataLeftShape.SelectedIndex > 2 ? 2 : 1;
                leftBreast1 += dataLeftInternalElectricalStructure.SelectedIndex > 1 ? dataLeftInternalElectricalStructure.SelectedIndex - 1 : 0;
                leftBreast1 += dataLeftSurroundingTissues.SelectedIndex > 1 ? (dataLeftSurroundingTissues.SelectedIndex > 3 ? 2 : 1) : 0;
                leftBreast1 += divergenceBetweenHistograms;
            }
            if (dataRightLocation.SelectedIndex > 0)
            {
                rightBreast1 += dataRightShape.SelectedIndex > 2 ? 2 : 1;
                rightBreast1 += dataRightInternalElectricalStructure.SelectedIndex > 1 ? dataRightInternalElectricalStructure.SelectedIndex - 1 : 0;
                rightBreast1 += dataRightSurroundingTissues.SelectedIndex > 1 ? (dataRightSurroundingTissues.SelectedIndex > 3 ? 2 : 1) : 0;
                rightBreast1 += divergenceBetweenHistograms;
            }

            if (dataLeftLocation2.SelectedIndex > 0)
            {
                leftBreast2 += dataLeftShape2.SelectedIndex > 2 ? 2 : 1;
                leftBreast2 += dataLeftInternalElectricalStructure2.SelectedIndex > 1 ? dataLeftInternalElectricalStructure2.SelectedIndex - 1 : 0;
                leftBreast2 += dataLeftSurroundingTissues2.SelectedIndex > 1 ? (dataLeftSurroundingTissues2.SelectedIndex > 3 ? 2 : 1) : 0;
                leftBreast2 += divergenceBetweenHistograms;
            }
            if (dataRightLocation2.SelectedIndex > 0)
            {
                rightBreast2 += dataRightShape2.SelectedIndex > 2 ? 2 : 1;
                rightBreast2 += dataRightInternalElectricalStructure2.SelectedIndex > 1 ? dataRightInternalElectricalStructure2.SelectedIndex - 1 : 0;
                rightBreast2 += dataRightSurroundingTissues2.SelectedIndex > 1 ? (dataRightSurroundingTissues2.SelectedIndex > 3 ? 2 : 1) : 0;
                rightBreast2 += divergenceBetweenHistograms;
            }

            if (dataLeftLocation3.SelectedIndex > 0)
            {
                leftBreast3 += dataLeftShape3.SelectedIndex > 2 ? 2 : 1;
                leftBreast3 += dataLeftInternalElectricalStructure3.SelectedIndex > 1 ? dataLeftInternalElectricalStructure3.SelectedIndex - 1 : 0;
                leftBreast3 += dataLeftSurroundingTissues3.SelectedIndex > 1 ? (dataLeftSurroundingTissues3.SelectedIndex > 3 ? 2 : 1) : 0;
                leftBreast3 += divergenceBetweenHistograms;
            }
            if (dataRightLocation3.SelectedIndex > 0)
            {
                rightBreast3 += dataRightShape3.SelectedIndex > 2 ? 2 : 1;
                rightBreast3 += dataRightInternalElectricalStructure3.SelectedIndex > 1 ? dataRightInternalElectricalStructure3.SelectedIndex - 1 : 0;
                rightBreast3 += dataRightSurroundingTissues3.SelectedIndex > 1 ? (dataRightSurroundingTissues3.SelectedIndex > 3 ? 2 : 1) : 0;
                rightBreast3 += divergenceBetweenHistograms;
            }

            List<int> list = new List<int>();
            list.Add(leftBreast1);
            list.Add(leftBreast2);
            list.Add(leftBreast3);
            list.Add(rightBreast1);
            list.Add(rightBreast2);
            list.Add(rightBreast3);
            list.Sort();
            string leftFlag = ""; 
            int maxLeftPoints=0;
            if (leftBreast1 >= leftBreast2 && leftBreast1 >= leftBreast3)
            {
                leftFlag = "L1";
                maxLeftPoints=leftBreast1;
            }
            else if (leftBreast2 >= leftBreast1 && leftBreast2 >= leftBreast3)
            {
                leftFlag = "L2";
                maxLeftPoints=leftBreast2;
            }
            else if (leftBreast3 >= leftBreast1 && leftBreast3 >= leftBreast2)
            {
                leftFlag = "L3";
                maxLeftPoints=leftBreast3;
            }             

            string rightFlag = "";
            int maxRightPoints=0;
            if (rightBreast1 >= rightBreast2 && rightBreast1 >= rightBreast3)
            {
                rightFlag = "R1";
                maxRightPoints=rightBreast1;
            }
            else if (rightBreast2 >= rightBreast1 && rightBreast2 >= rightBreast3)
            {
                rightFlag = "R2";
                maxRightPoints=rightBreast2;
            }
            else if (leftBreast3 >= rightBreast1 && rightBreast3 >= rightBreast2)
            {
                rightFlag = "R3";
                maxRightPoints=rightBreast3;
            }
             
            shortFormReportModel.DataLeftMaxFlag = leftFlag;
            shortFormReportModel.DataRightMaxFlag = rightFlag;            
            int maxPoints = list[5];
            if (!isEmpty)
            {                
                shortFormReportModel.DataTotalPts = maxPoints > 8 ? "8" : maxPoints == 0 ? "1" : maxPoints.ToString();
                this.dataTotalPts.SelectedIndex=Convert.ToInt32(shortFormReportModel.DataTotalPts);
                shortFormReportModel.DataPoint = maxPoints < 2 ? "1" : maxPoints < 4 ? "2" : maxPoints < 5 ? "3" : maxPoints < 8 ? "4" : "5";
                this.dataPoint.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataPoint);
                shortFormReportModel.DataBiRadsCategory = maxPoints < 2 ? "2" : maxPoints < 4 ? "3" : maxPoints < 5 ? "4" : maxPoints < 8 ? "5" : "6";
                this.dataBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataBiRadsCategory);

                shortFormReportModel.DataLeftTotalPts = maxLeftPoints > 8 ? "8" : maxLeftPoints == 0 ? "1" : maxLeftPoints.ToString();
                shortFormReportModel.DataRightTotalPts = maxRightPoints > 8 ? "8" : maxRightPoints == 0 ? "1" : maxRightPoints.ToString();
                this.dataLeftTotalPts.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftTotalPts);
                this.dataRightTotalPts.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightTotalPts);

                shortFormReportModel.DataLeftBiRadsCategory = maxLeftPoints < 2 ? "2" : maxLeftPoints < 4 ? "3" : maxLeftPoints < 5 ? "4" : maxLeftPoints < 8 ? "5" : "6";
                shortFormReportModel.DataRightBiRadsCategory = maxRightPoints < 2 ? "2" : maxRightPoints < 4 ? "3" : maxRightPoints < 5 ? "4" : maxRightPoints < 8 ? "5" : "6";
                this.dataLeftBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftBiRadsCategory);
                this.dataRightBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightBiRadsCategory);

            }
            else
            {
                shortFormReportModel.DataBiRadsCategory = "1";
                this.dataBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataBiRadsCategory);

                shortFormReportModel.DataLeftBiRadsCategory = "1";
                shortFormReportModel.DataRightBiRadsCategory = "1";
                this.dataLeftBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftBiRadsCategory);
                this.dataRightBiRadsCategory.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightBiRadsCategory);
            }
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
                generatePoints();
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

                ////生成PDF报告
                //var reportModel = CloneReportModel();
                ////Save PDF file
                //string lfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_LF.pdf";
                //string lfReportTempl = "Views/ExaminationReportDocument.xaml";
                //ExportPDF(lfReportTempl, lfPdfFile, reportModel);
                //File.Copy(lfPdfFile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_LF.pdf", true);
                //string sfPdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_SF.pdf";
                //string sfReportTempl = "Views/SummaryReportDocument.xaml";                
                //if (shortFormReportModel.DataScreenShotImg != null)
                //{
                //    sfReportTempl = "Views/SummaryReportImageDocument.xaml";                                        
                //}
                //ExportPDF(sfReportTempl, sfPdfFile, reportModel);
                //File.Copy(sfPdfFile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_SF.pdf", true);

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
            //shortFormReportModel.DataClientNum = this.dataClientNum.Text;
            //shortFormReportModel.DataUserCode = this.dataUserCode.Text;
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
            //if (this.dataMeikTech.SelectedItem != null)
            //{
            //    var technician = this.dataMeikTech.SelectedItem as User;
            //    shortFormReportModel.DataMeikTech = technician.Name;
            //    shortFormReportModel.DataTechLicense = technician.License;
            //}
                                    
            //shortFormReportModel.DataName = this.dataName.Text;
            
            shortFormReportModel.DataScreenDate = this.dataScreenDate.Text;
            //shortFormReportModel.DataScreenLocation = this.dataScreenLocation.Text;
            shortFormReportModel.DataScreenLocation = person.ScreenVenue;
            //shortFormReportModel.DataAddress = this.dataAddress.Text;
            shortFormReportModel.DataAddress = person.Address;

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
            
            shortFormReportModel.DataLeftBreastH = this.dataLeftBreastH.Text;
            shortFormReportModel.DataRightBreastH = this.dataRightBreastH.Text;
            shortFormReportModel.DataLeftBreastM = this.dataLeftBreastM.Text;
            shortFormReportModel.DataRightBreastM = this.dataRightBreastM.Text;
            shortFormReportModel.DataLeftBreast = this.dataLeftBreastH.Text + ":" + this.dataLeftBreastM.Text;
            shortFormReportModel.DataRightBreast = this.dataRightBreastH.Text + ":" + this.dataRightBreastM.Text; 

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

            shortFormReportModel.DataLeftNumber = this.dataLeftNumber.SelectedIndex.ToString();
            shortFormReportModel.DataRightNumber = this.dataRightNumber.SelectedIndex.ToString();

            shortFormReportModel.DataLeftLocation = this.dataLeftLocation.SelectedIndex.ToString();
            shortFormReportModel.DataRightLocation = this.dataRightLocation.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSizeLong = this.dataLeftSizeLong.Text;
            shortFormReportModel.DataLeftSizeWidth = this.dataLeftSizeWidth.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeLong) || !string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeWidth)) {
                shortFormReportModel.DataLeftSize = shortFormReportModel.DataLeftSizeLong + "x" + shortFormReportModel.DataLeftSizeWidth + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataRightSizeLong = this.dataRightSizeLong.Text;
            shortFormReportModel.DataRightSizeWidth = this.dataRightSizeWidth.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataRightSizeLong) || !string.IsNullOrEmpty(shortFormReportModel.DataRightSizeWidth))
            {
                shortFormReportModel.DataRightSize = shortFormReportModel.DataRightSizeLong + "x" + shortFormReportModel.DataRightSizeWidth + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataLeftShape = this.dataLeftShape.SelectedIndex.ToString();
            shortFormReportModel.DataRightShape = this.dataRightShape.SelectedIndex.ToString();
            shortFormReportModel.DataLeftContourAroundFocus = this.dataLeftContourAroundFocus.SelectedIndex.ToString();
            shortFormReportModel.DataRightContourAroundFocus = this.dataRightContourAroundFocus.SelectedIndex.ToString();
            shortFormReportModel.DataLeftInternalElectricalStructure = this.dataLeftInternalElectricalStructure.SelectedIndex.ToString();
            shortFormReportModel.DataRightInternalElectricalStructure = this.dataRightInternalElectricalStructure.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSurroundingTissues = this.dataLeftSurroundingTissues.SelectedIndex.ToString();
            shortFormReportModel.DataRightSurroundingTissues = this.dataRightSurroundingTissues.SelectedIndex.ToString();

            shortFormReportModel.DataLeftLocation2 = this.dataLeftLocation2.SelectedIndex.ToString();
            shortFormReportModel.DataRightLocation2 = this.dataRightLocation2.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSizeLong2 = this.dataLeftSizeLong2.Text;
            shortFormReportModel.DataLeftSizeWidth2 = this.dataLeftSizeWidth2.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeLong2) || !string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeWidth2))
            {
                shortFormReportModel.DataLeftSize2 = shortFormReportModel.DataLeftSizeLong2 + "x" + shortFormReportModel.DataLeftSizeWidth2 + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataRightSizeLong2 = this.dataRightSizeLong2.Text;
            shortFormReportModel.DataRightSizeWidth2 = this.dataRightSizeWidth2.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataRightSizeLong2) || !string.IsNullOrEmpty(shortFormReportModel.DataRightSizeWidth2))
            {
                shortFormReportModel.DataRightSize2 = shortFormReportModel.DataRightSizeLong2 + "x" + shortFormReportModel.DataRightSizeWidth2 + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataLeftShape2 = this.dataLeftShape2.SelectedIndex.ToString();
            shortFormReportModel.DataRightShape2 = this.dataRightShape2.SelectedIndex.ToString();
            shortFormReportModel.DataLeftContourAroundFocus2 = this.dataLeftContourAroundFocus2.SelectedIndex.ToString();
            shortFormReportModel.DataRightContourAroundFocus2 = this.dataRightContourAroundFocus2.SelectedIndex.ToString();
            shortFormReportModel.DataLeftInternalElectricalStructure2 = this.dataLeftInternalElectricalStructure2.SelectedIndex.ToString();
            shortFormReportModel.DataRightInternalElectricalStructure2 = this.dataRightInternalElectricalStructure2.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSurroundingTissues2 = this.dataLeftSurroundingTissues2.SelectedIndex.ToString();
            shortFormReportModel.DataRightSurroundingTissues2 = this.dataRightSurroundingTissues2.SelectedIndex.ToString();

            shortFormReportModel.DataLeftLocation3 = this.dataLeftLocation3.SelectedIndex.ToString();
            shortFormReportModel.DataRightLocation3 = this.dataRightLocation3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSizeLong3 = this.dataLeftSizeLong3.Text;
            shortFormReportModel.DataLeftSizeWidth3 = this.dataLeftSizeWidth3.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeLong3) || !string.IsNullOrEmpty(shortFormReportModel.DataLeftSizeWidth3))
            {
                shortFormReportModel.DataLeftSize3 = shortFormReportModel.DataLeftSizeLong3 + "x" + shortFormReportModel.DataLeftSizeWidth3 + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataRightSizeLong3 = this.dataRightSizeLong3.Text;
            shortFormReportModel.DataRightSizeWidth3 = this.dataRightSizeWidth3.Text;
            if (!string.IsNullOrEmpty(shortFormReportModel.DataRightSizeLong3) || !string.IsNullOrEmpty(shortFormReportModel.DataRightSizeWidth3))
            {
                shortFormReportModel.DataRightSize3 = shortFormReportModel.DataRightSizeLong3 + "x" + shortFormReportModel.DataRightSizeWidth3 + App.Current.FindResource("ReportContext_225").ToString();
            }
            shortFormReportModel.DataLeftShape3 = this.dataLeftShape3.SelectedIndex.ToString();
            shortFormReportModel.DataRightShape3 = this.dataRightShape3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftContourAroundFocus3 = this.dataLeftContourAroundFocus3.SelectedIndex.ToString();
            shortFormReportModel.DataRightContourAroundFocus3 = this.dataRightContourAroundFocus3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftInternalElectricalStructure3 = this.dataLeftInternalElectricalStructure3.SelectedIndex.ToString();
            shortFormReportModel.DataRightInternalElectricalStructure3 = this.dataRightInternalElectricalStructure3.SelectedIndex.ToString();
            shortFormReportModel.DataLeftSurroundingTissues3 = this.dataLeftSurroundingTissues3.SelectedIndex.ToString();
            shortFormReportModel.DataRightSurroundingTissues3 = this.dataRightSurroundingTissues3.SelectedIndex.ToString();

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
            reportModel.DataLeftBiRadsCategory = this.dataLeftBiRadsCategory.Text;
            reportModel.DataRightBiRadsCategory = this.dataRightBiRadsCategory.Text;
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
            reportModel.DataLeftNumber = this.dataLeftNumber.Text;
            reportModel.DataRightNumber = this.dataRightNumber.Text;

            reportModel.DataLeftLocation = this.dataLeftLocation.Text;
            reportModel.DataRightLocation = this.dataRightLocation.Text;
            reportModel.DataLeftShape = this.dataLeftShape.Text;
            reportModel.DataRightShape = this.dataRightShape.Text;
            reportModel.DataLeftContourAroundFocus = this.dataLeftContourAroundFocus.Text;
            reportModel.DataRightContourAroundFocus = this.dataRightContourAroundFocus.Text;
            reportModel.DataLeftInternalElectricalStructure = this.dataLeftInternalElectricalStructure.Text;
            reportModel.DataRightInternalElectricalStructure = this.dataRightInternalElectricalStructure.Text;
            reportModel.DataLeftSurroundingTissues = this.dataLeftSurroundingTissues.Text;
            reportModel.DataRightSurroundingTissues = this.dataRightSurroundingTissues.Text;
            reportModel.DataLeftLocation2 = this.dataLeftLocation2.Text;
            reportModel.DataRightLocation2 = this.dataRightLocation2.Text;
            reportModel.DataLeftShape2 = this.dataLeftShape2.Text;
            reportModel.DataRightShape2 = this.dataRightShape2.Text;
            reportModel.DataLeftContourAroundFocus2 = this.dataLeftContourAroundFocus2.Text;
            reportModel.DataRightContourAroundFocus2 = this.dataRightContourAroundFocus2.Text;
            reportModel.DataLeftInternalElectricalStructure2 = this.dataLeftInternalElectricalStructure2.Text;
            reportModel.DataRightInternalElectricalStructure2 = this.dataRightInternalElectricalStructure2.Text;
            reportModel.DataLeftSurroundingTissues2 = this.dataLeftSurroundingTissues2.Text;
            reportModel.DataRightSurroundingTissues2 = this.dataRightSurroundingTissues2.Text;
            reportModel.DataLeftLocation3 = this.dataLeftLocation3.Text;
            reportModel.DataRightLocation3 = this.dataRightLocation3.Text;
            reportModel.DataLeftShape3 = this.dataLeftShape3.Text;
            reportModel.DataRightShape3 = this.dataRightShape3.Text;
            reportModel.DataLeftContourAroundFocus3 = this.dataLeftContourAroundFocus3.Text;
            reportModel.DataRightContourAroundFocus3 = this.dataRightContourAroundFocus3.Text;
            reportModel.DataLeftInternalElectricalStructure3 = this.dataLeftInternalElectricalStructure3.Text;
            reportModel.DataRightInternalElectricalStructure3 = this.dataRightInternalElectricalStructure3.Text;
            reportModel.DataLeftSurroundingTissues3 = this.dataLeftSurroundingTissues3.Text;
            reportModel.DataRightSurroundingTissues3 = this.dataRightSurroundingTissues3.Text;

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
            reportModel.DataLeftTotalPts = this.dataLeftTotalPts.Text;
            reportModel.DataRightTotalPts = this.dataRightTotalPts.Text;
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
                //Clone生成全文本的报表数据对象模型
                var reportModel = CloneReportModel();
                //打开文件夹对话框，选择要保存的目录
                System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
                folderBrowserDialog.SelectedPath = this.person.ArchiveFolder;
                System.Windows.Forms.DialogResult res = folderBrowserDialog.ShowDialog();
                if (res == System.Windows.Forms.DialogResult.OK)
                {                    
                    string folderName = folderBrowserDialog.SelectedPath;
                    string strName = person.SurName + (string.IsNullOrEmpty(person.GivenName) ? "" : "," + person.GivenName) + (string.IsNullOrEmpty(person.OtherName) ? "" : " " + person.OtherName)+".pdf";  
                    //生成Examination报告的PDF文件
                    string lfPdfFile = folderName + System.IO.Path.DirectorySeparatorChar + person.Code + " LF - " + strName;
                    string lfReportTempl = "Views/ExaminationReportFlow.xaml";
                    ExportFlowDocumentPDF(lfReportTempl, lfPdfFile, reportModel);

                    //生成Summary报告的PDF文件
                    string sfPdfFile = folderName + System.IO.Path.DirectorySeparatorChar + person.Code + "SF - " + strName;
                    string sfReportTempl = "Views/SummaryReportDocument.xaml";
                    if (shortFormReportModel.DataScreenShotImg != null)
                    {
                        sfReportTempl = "Views/SummaryReportNuvoTekDocument.xaml";
                    }
                    ExportPDF(sfReportTempl, sfPdfFile, reportModel);                    

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

        /// <summary>
        ///针对FixedPage对象生成PDF 
        /// </summary>
        /// <param name="reportTempl"></param>
        /// <param name="pdfFile"></param>
        /// <param name="reportModel"></param>
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

        /// <summary>
        /// 针对FLowDocument对象生成PDF文件
        /// </summary>
        /// <param name="reportTempl"></param>
        /// <param name="pdfFile"></param>
        /// <param name="reportModel"></param>
        private void ExportFlowDocumentPDF(string reportTempl, string pdfFile, ShortFormReport reportModel)
        {
            string xpsFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
            if (File.Exists(xpsFile))
            {
                File.Delete(xpsFile);
            }

            FlowDocument page = (FlowDocument)PrintPreviewWindow.LoadFlowDocumentAndRender(reportTempl, reportModel);
            LoadDataForFlowDocument(page, reportModel);
            XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
            //将flow document写入基于内存的xps document中去
            XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
            writer.Write(((IDocumentPaginatorSource)page).DocumentPaginator);
            xpsDocument.Close();
            if (File.Exists(pdfFile))
            {
                File.Delete(pdfFile);
            }
            PDFTools.SavePDFFile(xpsFile, pdfFile);
        }

        private void LoadDataForFlowDocument(FlowDocument page, ShortFormReport reportModel)
        {
            if (reportModel != null)
            {
                var textBlock1 = page.FindName("dataClientNum") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataClientNum;
                }
                textBlock1 = page.FindName("dataUserCode") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataUserCode;
                }
                textBlock1 = page.FindName("dataScreenDate") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataScreenDate;
                }
                textBlock1 = page.FindName("dataName") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataName;
                }
                textBlock1 = page.FindName("dataAge") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataAge;
                }
                textBlock1 = page.FindName("dataHeight") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataHeight;
                }
                textBlock1 = page.FindName("dataWeight") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataWeight;
                }
                textBlock1 = page.FindName("dataScreenLocation") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataScreenLocation;
                }
                textBlock1 = page.FindName("dataMobile") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataMobile;
                }
                textBlock1 = page.FindName("dataEmail") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataEmail;
                }
                textBlock1 = page.FindName("dataLeftBreast") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftBreast;
                }
                textBlock1 = page.FindName("dataRightBreast") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightBreast;
                }
                textBlock1 = page.FindName("dataLeftPalpableMass") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftPalpableMass;
                }
                textBlock1 = page.FindName("dataRightPalpableMass") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightPalpableMass;
                }
                textBlock1 = page.FindName("dataLeftChangesOfElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftChangesOfElectricalConductivity;
                }
                textBlock1 = page.FindName("dataRightChangesOfElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightChangesOfElectricalConductivity;
                }
                textBlock1 = page.FindName("dataLeftMammaryStruct") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMammaryStruct;
                }
                textBlock1 = page.FindName("dataRightMammaryStruct") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMammaryStruct;
                }
                textBlock1 = page.FindName("dataLeftLactiferousSinusZone") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftLactiferousSinusZone;
                }
                textBlock1 = page.FindName("dataRightLactiferousSinusZone") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightLactiferousSinusZone;
                }
                textBlock1 = page.FindName("dataLeftMammaryContour") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMammaryContour;
                }
                textBlock1 = page.FindName("dataRightMammaryContour") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMammaryContour;
                }
                textBlock1 = page.FindName("dataLeftNumber") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftNumber;
                }
                textBlock1 = page.FindName("dataRightNumber") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightNumber;
                }
                textBlock1 = page.FindName("dataLeftLocation") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftLocation;
                }
                textBlock1 = page.FindName("dataRightLocation") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightLocation;
                }
                textBlock1 = page.FindName("dataLeftSize") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSize;
                }
                textBlock1 = page.FindName("dataRightSize") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSize;
                }
                textBlock1 = page.FindName("dataLeftShape") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftShape;
                }
                textBlock1 = page.FindName("dataRightShape") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightShape;
                }
                textBlock1 = page.FindName("dataLeftContourAroundFocus") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftContourAroundFocus;
                }
                textBlock1 = page.FindName("dataRightContourAroundFocus") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightContourAroundFocus;
                }
                textBlock1 = page.FindName("dataLeftInternalElectricalStructure") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftInternalElectricalStructure;
                }
                textBlock1 = page.FindName("dataRightInternalElectricalStructure") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightInternalElectricalStructure;
                }
                textBlock1 = page.FindName("dataLeftSurroundingTissues") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSurroundingTissues;
                }
                textBlock1 = page.FindName("dataRightSurroundingTissues") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSurroundingTissues;
                }
                textBlock1 = page.FindName("dataLeftLocation2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftLocation2;
                }
                textBlock1 = page.FindName("dataRightLocation2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightLocation2;
                }
                textBlock1 = page.FindName("dataLeftSize2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSize2;
                }
                textBlock1 = page.FindName("dataRightSize2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSize2;
                }
                textBlock1 = page.FindName("dataLeftShape2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftShape2;
                }
                textBlock1 = page.FindName("dataRightShape2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightShape2;
                }
                textBlock1 = page.FindName("dataLeftContourAroundFocus2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftContourAroundFocus2;
                }
                textBlock1 = page.FindName("dataRightContourAroundFocus2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightContourAroundFocus2;
                }
                textBlock1 = page.FindName("dataLeftInternalElectricalStructure2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftInternalElectricalStructure2;
                }
                textBlock1 = page.FindName("dataRightInternalElectricalStructure2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightInternalElectricalStructure2;
                }
                textBlock1 = page.FindName("dataLeftSurroundingTissues2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSurroundingTissues2;
                }
                textBlock1 = page.FindName("dataRightSurroundingTissues2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSurroundingTissues2;
                }
                textBlock1 = page.FindName("dataLeftLocation3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftLocation3;
                }
                textBlock1 = page.FindName("dataRightLocation3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightLocation3;
                }
                textBlock1 = page.FindName("dataLeftSize3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSize3;
                }
                textBlock1 = page.FindName("dataRightSize3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSize3;
                }
                textBlock1 = page.FindName("dataLeftShape3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftShape3;
                }
                textBlock1 = page.FindName("dataRightShape3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightShape3;
                }
                textBlock1 = page.FindName("dataLeftContourAroundFocus3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftContourAroundFocus3;
                }
                textBlock1 = page.FindName("dataRightContourAroundFocus3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightContourAroundFocus3;
                }
                textBlock1 = page.FindName("dataLeftInternalElectricalStructure3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftInternalElectricalStructure3;
                }
                textBlock1 = page.FindName("dataRightInternalElectricalStructure3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightInternalElectricalStructure3;
                }
                textBlock1 = page.FindName("dataLeftSurroundingTissues3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftSurroundingTissues3;
                }
                textBlock1 = page.FindName("dataRightSurroundingTissues3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightSurroundingTissues3;
                }
                textBlock1 = page.FindName("dataLeftOncomarkerHighlightBenignChanges") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftOncomarkerHighlightBenignChanges;
                }
                textBlock1 = page.FindName("dataRightOncomarkerHighlightBenignChanges") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightOncomarkerHighlightBenignChanges;
                }
                textBlock1 = page.FindName("dataLeftOncomarkerHighlightSuspiciousChanges") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftOncomarkerHighlightSuspiciousChanges;
                }
                textBlock1 = page.FindName("dataRightOncomarkerHighlightSuspiciousChanges") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightOncomarkerHighlightSuspiciousChanges;
                }
                textBlock1 = page.FindName("dataLeftMeanElectricalConductivity1") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMeanElectricalConductivity1;
                }
                textBlock1 = page.FindName("dataRightMeanElectricalConductivity1") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMeanElectricalConductivity1;
                }
                textBlock1 = page.FindName("dataLeftMeanElectricalConductivity2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMeanElectricalConductivity2;
                }
                textBlock1 = page.FindName("dataRightMeanElectricalConductivity2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMeanElectricalConductivity2;
                }
                textBlock1 = page.FindName("dataMeanElectricalConductivity3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataMeanElectricalConductivity3;
                }
                textBlock1 = page.FindName("dataLeftMeanElectricalConductivity3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMeanElectricalConductivity3;
                }
                textBlock1 = page.FindName("dataRightMeanElectricalConductivity3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMeanElectricalConductivity3;
                }
                textBlock1 = page.FindName("dataLeftComparativeElectricalConductivity1") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftComparativeElectricalConductivity1;
                }
                textBlock1 = page.FindName("dataLeftComparativeElectricalConductivity2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftComparativeElectricalConductivity2;
                }
                textBlock1 = page.FindName("dataComparativeElectricalConductivity3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataComparativeElectricalConductivity3;
                }
                textBlock1 = page.FindName("dataLeftComparativeElectricalConductivity3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftComparativeElectricalConductivity3;
                }
                textBlock1 = page.FindName("dataLeftDivergenceBetweenHistograms1") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftDivergenceBetweenHistograms1;
                }
                textBlock1 = page.FindName("dataLeftDivergenceBetweenHistograms2") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftDivergenceBetweenHistograms2;
                }
                textBlock1 = page.FindName("dataDivergenceBetweenHistograms3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataDivergenceBetweenHistograms3;
                }
                textBlock1 = page.FindName("dataLeftDivergenceBetweenHistograms3") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftDivergenceBetweenHistograms3;
                }
                textBlock1 = page.FindName("dataLeftComparisonWithNorm") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftComparisonWithNorm;
                }
                textBlock1 = page.FindName("dataRightComparisonWithNorm") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightComparisonWithNorm;
                }
                textBlock1 = page.FindName("dataLeftPhaseElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftPhaseElectricalConductivity;
                }
                textBlock1 = page.FindName("dataRightPhaseElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightPhaseElectricalConductivity;
                }
                textBlock1 = page.FindName("dataAgeElectricalConductivityReference") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataAgeElectricalConductivityReference;
                }
                textBlock1 = page.FindName("dataLeftAgeElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftAgeElectricalConductivity;
                }
                textBlock1 = page.FindName("dataRightAgeElectricalConductivity") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightAgeElectricalConductivity;
                }
                textBlock1 = page.FindName("dataExamConclusion") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataExamConclusion;
                }
                textBlock1 = page.FindName("dataLeftMammaryGland") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMammaryGland;
                }
                
                textBlock1 = page.FindName("dataLeftAgeRelated") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftAgeRelated;
                }
                
                textBlock1 = page.FindName("dataLeftMammaryGlandResult") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftMammaryGlandResult;
                }
                textBlock1 = page.FindName("dataRightMammaryGland") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMammaryGland;
                }
                
                textBlock1 = page.FindName("dataRightAgeRelated") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightAgeRelated;
                }
                
                textBlock1 = page.FindName("dataRightMammaryGlandResult") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightMammaryGlandResult;
                }
                textBlock1 = page.FindName("dataTotalPts") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataTotalPts;
                }
                textBlock1 = page.FindName("dataLeftTotalPts") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftTotalPts;
                }
                textBlock1 = page.FindName("dataRightTotalPts") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightTotalPts;
                }
                textBlock1 = page.FindName("dataPoint") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataPoint;
                }
                textBlock1 = page.FindName("dataBiRadsCategory") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataBiRadsCategory;
                }
                textBlock1 = page.FindName("dataLeftBiRadsCategory") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataLeftBiRadsCategory;
                }
                textBlock1 = page.FindName("dataRightBiRadsCategory") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRightBiRadsCategory;
                }
                textBlock1 = page.FindName("dataRecommendation") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataRecommendation;
                }
                textBlock1 = page.FindName("dataFurtherExam") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataFurtherExam;
                }
                textBlock1 = page.FindName("dataConclusion") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataConclusion;
                }
                textBlock1 = page.FindName("dataComments") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataComments;
                }
                textBlock1 = page.FindName("dataMeikTech") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataMeikTech;
                }
                textBlock1 = page.FindName("dataDoctor") as TextBlock;
                if (textBlock1 != null)
                {
                    textBlock1.Text = reportModel.DataDoctor;
                }                
            }
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
                this.WindowState = WindowState.Minimized;
                App.opendWin = this;
                ScreenCapture screenCaptureWin = new ScreenCapture(this.person);
                screenCaptureWin.callbackMethod = LoadScreenShot;
                screenCaptureWin.ShowDialog();               
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

                //this.btnScreenShot.Content = App.Current.FindResource("ReportContext_170").ToString();                
                //this.btnRemoveImg.Visibility = Visibility.Visible;
            }
           
            //dataScreenShotImg.Source = ImageTools.GetBitmapImage(imgFileName as string);
        }

        private void btnViewImg_Click(object sender, RoutedEventArgs e)
        {            
            if (shortFormReportModel.DataScreenShotImg == null)
            {
                MessageBox.Show(this, App.Current.FindResource("Message_34").ToString());
            }
            else
            {
                ViewImagePage viewImagePage = new ViewImagePage(shortFormReportModel.DataScreenShotImg);
                viewImagePage.ShowDialog();
            }
        }

        //private void btnRemoveImg_Click(object sender, RoutedEventArgs e)
        //{
        //    this.btnScreenShot.Content = App.Current.FindResource("ReportContext_175").ToString();            
        //    shortFormReportModel.DataScreenShotImg = null;
        //    this.btnRemoveImg.Visibility = Visibility.Hidden;
        //}

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            //判断是否已经从MEIK生成的DOC文档中导入检查数据，如果之前没有，则查找是否已在本地生成DOC文档，导入数据            
            string docFile = FindUserReportWord(person.ArchiveFolder);
            if (!string.IsNullOrEmpty(docFile) &&File.Exists(docFile))
            {
                ShortFormReport shortFormReport = WordTools.ReadWordFile(docFile);
                if (shortFormReport != null)
                {
                    try
                    {
                        /**因为ShortFormReport的属性没有添加依赖变更事件，所以这里为shortFormReportModel赋值不会影响页面显示效果。
                         * 不过即使为ShortFormReport的属性添加上依赖变更事件，但由于之前用户可能已经序列化过一些报表数据，
                         * 如果改变ShortFormReport对象的继承定义，会导致反序列化失败，所以只能暂时不处理，只在这里强制修改页面元素
                         * **/
                        shortFormReportModel.DataMenstrualCycle = shortFormReport.DataMenstrualCycle;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataMenstrualCycle))
                        {
                            this.dataMenstrualCycle.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataMenstrualCycle);
                        }
                        shortFormReportModel.DataLeftMeanElectricalConductivity1 = shortFormReport.DataLeftMeanElectricalConductivity1;
                        this.dataLeftMeanElectricalConductivity1.Text = shortFormReportModel.DataLeftMeanElectricalConductivity1;
                        shortFormReportModel.DataRightMeanElectricalConductivity1 = shortFormReport.DataRightMeanElectricalConductivity1;
                        this.dataRightMeanElectricalConductivity1.Text = shortFormReportModel.DataRightMeanElectricalConductivity1;
                        shortFormReportModel.DataLeftMeanElectricalConductivity2 = shortFormReport.DataLeftMeanElectricalConductivity2;
                        this.dataLeftMeanElectricalConductivity2.Text = shortFormReportModel.DataLeftMeanElectricalConductivity2;
                        shortFormReportModel.DataRightMeanElectricalConductivity2 = shortFormReport.DataRightMeanElectricalConductivity2;
                        this.dataRightMeanElectricalConductivity2.Text = shortFormReportModel.DataRightMeanElectricalConductivity2;
                        shortFormReportModel.DataMeanElectricalConductivity3 = shortFormReport.DataMeanElectricalConductivity3;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataMeanElectricalConductivity3))
                        {
                            this.dataMeanElectricalConductivity3.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataMeanElectricalConductivity3);
                        }
                        shortFormReportModel.DataLeftMeanElectricalConductivity3 = shortFormReport.DataLeftMeanElectricalConductivity3;
                        this.dataLeftMeanElectricalConductivity3.Text = shortFormReportModel.DataLeftMeanElectricalConductivity3;
                        shortFormReportModel.DataRightMeanElectricalConductivity3 = shortFormReport.DataRightMeanElectricalConductivity3;
                        this.dataRightMeanElectricalConductivity3.Text = shortFormReportModel.DataRightMeanElectricalConductivity3;
                        shortFormReportModel.DataLeftComparativeElectricalConductivity1 = shortFormReport.DataLeftComparativeElectricalConductivity1;
                        this.dataLeftComparativeElectricalConductivity1.Text = shortFormReportModel.DataLeftComparativeElectricalConductivity1;
                        shortFormReportModel.DataLeftComparativeElectricalConductivity2 = shortFormReport.DataLeftComparativeElectricalConductivity2;
                        this.dataLeftComparativeElectricalConductivity2.Text = shortFormReportModel.DataLeftComparativeElectricalConductivity2;
                        shortFormReportModel.DataLeftComparativeElectricalConductivity3 = shortFormReport.DataLeftComparativeElectricalConductivity3;
                        this.dataLeftComparativeElectricalConductivity3.Text = shortFormReportModel.DataLeftComparativeElectricalConductivity3;

                        shortFormReportModel.DataComparativeElectricalConductivity3 = shortFormReportModel.DataComparativeElectricalConductivity3;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataComparativeElectricalConductivity3))
                        {
                            this.dataComparativeElectricalConductivity3.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataComparativeElectricalConductivity3);
                        }

                        shortFormReportModel.DataDivergenceBetweenHistograms3 = shortFormReportModel.DataDivergenceBetweenHistograms3;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataDivergenceBetweenHistograms3))
                        {
                            this.dataDivergenceBetweenHistograms3.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataDivergenceBetweenHistograms3);
                        }

                        shortFormReportModel.DataLeftDivergenceBetweenHistograms1 = shortFormReport.DataLeftDivergenceBetweenHistograms1.Replace("%","");
                        this.dataLeftDivergenceBetweenHistograms1.Text = shortFormReportModel.DataLeftDivergenceBetweenHistograms1;
                        shortFormReportModel.DataLeftDivergenceBetweenHistograms2 = shortFormReport.DataLeftDivergenceBetweenHistograms2.Replace("%", ""); 
                        this.dataLeftDivergenceBetweenHistograms2.Text = shortFormReportModel.DataLeftDivergenceBetweenHistograms2;
                        shortFormReportModel.DataLeftDivergenceBetweenHistograms3 = shortFormReport.DataLeftDivergenceBetweenHistograms3.Replace("%", ""); 
                        this.dataLeftDivergenceBetweenHistograms3.Text = shortFormReportModel.DataLeftDivergenceBetweenHistograms3;

                        shortFormReportModel.DataLeftPhaseElectricalConductivity = shortFormReport.DataLeftPhaseElectricalConductivity;
                        this.dataLeftPhaseElectricalConductivity.Text = shortFormReportModel.DataLeftPhaseElectricalConductivity;
                        shortFormReportModel.DataRightPhaseElectricalConductivity = shortFormReport.DataRightPhaseElectricalConductivity;
                        this.dataRightPhaseElectricalConductivity.Text = shortFormReportModel.DataRightPhaseElectricalConductivity;
                        shortFormReportModel.DataAgeElectricalConductivityReference = shortFormReport.DataAgeElectricalConductivityReference;
                        this.dataAgeElectricalConductivityReference.Text = shortFormReportModel.DataAgeElectricalConductivityReference;
                        shortFormReportModel.DataLeftAgeElectricalConductivity = shortFormReport.DataLeftAgeElectricalConductivity;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftAgeElectricalConductivity))
                        {
                            this.dataLeftAgeElectricalConductivity.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftAgeElectricalConductivity);
                        }
                        shortFormReportModel.DataRightAgeElectricalConductivity = shortFormReport.DataRightAgeElectricalConductivity;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataRightAgeElectricalConductivity))
                        {
                            this.dataRightAgeElectricalConductivity.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightAgeElectricalConductivity);
                        }
                        shortFormReportModel.DataExamConclusion = shortFormReport.DataExamConclusion;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataExamConclusion))
                        {
                            this.dataExamConclusion.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataExamConclusion);
                        }
                        shortFormReportModel.DataLeftMammaryGland = shortFormReport.DataLeftMammaryGland;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftMammaryGland))
                        {
                            this.dataLeftMammaryGland.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftMammaryGland);
                        }
                        shortFormReportModel.DataLeftAgeRelated = shortFormReport.DataLeftAgeRelated;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataLeftAgeRelated))
                        {
                            this.dataLeftAgeRelated.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataLeftAgeRelated);
                        }
                        shortFormReportModel.DataRightMammaryGland = shortFormReport.DataRightMammaryGland;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataRightMammaryGland))
                        {
                            this.dataRightMammaryGland.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightMammaryGland);
                        }
                        shortFormReportModel.DataRightAgeRelated = shortFormReport.DataRightAgeRelated;
                        if (!string.IsNullOrEmpty(shortFormReportModel.DataRightAgeRelated))
                        {
                            this.dataRightAgeRelated.SelectedIndex = Convert.ToInt32(shortFormReportModel.DataRightAgeRelated);
                        }
                        MessageBox.Show(this, App.Current.FindResource("Message_27").ToString());
                        if (File.Exists(person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_protocal.doc"))
                        {
                            File.Delete(person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_protocal.doc");
                        }
                        //改名原始的Protocal文件名，解決不能重新生成一個新的Protocal文件的問題
                        File.Move(docFile, person.ArchiveFolder+System.IO.Path.DirectorySeparatorChar +person.Code+"_protocal.doc");
                    }
                    catch (Exception ex) {
                        MessageBox.Show(this, ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show(this, string.Format(App.Current.FindResource("Message_33").ToString(), docFile));
                }
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
            int left = (int)(System.Windows.SystemParameters.PrimaryScreenWidth - 216);
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
                    //else if (App.strStart.Equals(winText.ToString(), StringComparison.OrdinalIgnoreCase))//点击开始设备按钮
                    //{
                    //    string meikiniFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "MEIK.ini";
                    //    var screenFolderPath = OperateIniFile.ReadIniData("Base", "Patients base", "", meikiniFile);
                    //    if (App.countDictionary.ContainsKey(screenFolderPath))
                    //    {
                    //        List<long> ticks = App.countDictionary[screenFolderPath];
                    //        ticks.Add(DateTime.Now.Ticks);
                    //    }
                    //    else
                    //    {
                    //        List<long> ticks = new List<long>();
                    //        ticks.Add(DateTime.Now.Ticks);
                    //        App.countDictionary.Add(screenFolderPath, ticks);
                    //    } 
                    //    //序列化统计字典到文件
                    //    string countFile = App.meikFolder + System.IO.Path.DirectorySeparatorChar + "sc.data";
                    //    SerializeUtilities.Serialize<SortedDictionary<string, List<long>>>(App.countDictionary, countFile);
                    //}
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
        

        private void Number_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex re = new Regex("[^0-9.-]+");
            e.Handled = re.IsMatch(e.Text);
        }

        private void Hour_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex re = new Regex("[0-9.-]+");
            if (re.IsMatch(e.Text))
            {
                var textBox = (TextBox)sender;
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    if (Convert.ToInt32(textBox.Text + e.Text) < 24 && (textBox.Text.Length+1) <= 2)
                    {
                        e.Handled = false;
                    }
                    else
                    {
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
                e.Handled = true;
            }            
            
        }

        private void Minute_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex re = new Regex("[0-9.-]+");
            if (re.IsMatch(e.Text))
            {
                var textBox = (TextBox)sender;
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    if (Convert.ToInt32(textBox.Text+e.Text) < 60 && (textBox.Text.Length+1) <= 2)
                    {
                        e.Handled = false;
                    }
                    else
                    {
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
                e.Handled = true;
            }            
        }       
               
    }
}

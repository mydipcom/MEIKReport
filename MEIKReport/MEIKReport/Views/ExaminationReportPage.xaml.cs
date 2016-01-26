using MEIKReport.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
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
        public ExaminationReportPage()
        {
            InitializeComponent();
        }
        public ExaminationReportPage(object data): this()
        {
            try { 
                this.person = data as Person;
                if (this.person == null)
                {
                    MessageBox.Show("Please select a patient.");
                    this.Close();
                }
                else
                {
                    string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                    if (File.Exists(dataFile))
                    {                    
                        this.shortFormReportModel = SerializeUtilities.Desrialize<ShortFormReport>(dataFile);
                        this.reportDataGrid.DataContext = this.shortFormReportModel;                                            
                        if (shortFormReportModel.DataSignImg != null)
                        {
                            this.dataSignImg.Source = ImageTools.GetBitmapImage(shortFormReportModel.DataSignImg);
                        }     
                    }
                    else
                    {
                        shortFormReportModel.DataUserCode = person.Code;
                        shortFormReportModel.DataName = person.SurName;
                        shortFormReportModel.DataAge = person.Age + "";
                        shortFormReportModel.DataAddress = person.Address;
                        shortFormReportModel.DataScreenDate = DateTime.Parse(person.Regdate).ToLongDateString();
                        
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
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            try { 
                App.opendWin = null;
                IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
                //如果主窗体存在
                if (mainWinHwnd != IntPtr.Zero)
                {
                    int WM_SYSCOMMAND = 0x0112;
                    int SC_CLOSE = 0xF060;
                    Win32Api.SendMessage(mainWinHwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
                }
                this.Owner.Show();            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                PrintPreviewWindow previewWnd = new PrintPreviewWindow("Views/ExaminationReportDocument.xaml", true, shortFormReportModel);
                previewWnd.Owner = this;
                previewWnd.ShowInTaskbar = false;
                previewWnd.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                    fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.5, 96 * 11);

                    PageContent pageContent = new PageContent();
                    ((IAddChild)pageContent).AddChild(page);
                    fixedDoc.Pages.Add(pageContent);//将对象加入到当前文档中
                    Dispatcher.BeginInvoke(new DoPrintMethod(DoPrint), DispatcherPriority.ApplicationIdle, pdlg, fixedDoc.DocumentPaginator);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                    IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, "Diagnostics");
                    Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                }                
                WinMinimized();
                var mainWin = this.Owner.Owner as MainWindow;
                mainWin.StartMouseHook();

            }
            catch (Exception ex)
            {
                MessageBox.Show("System Exception: " + ex.Message);
            }
        }        

        private void save_Click(object sender, RoutedEventArgs e)
        {
            SaveReport();
        }

        private void btnSignature_Click(object sender, RoutedEventArgs e)
        {
            SignatureBox signBox = new SignatureBox();
            signBox.callbackMethod = ShowSignature;
            signBox.Show();
        }

        /// <summary>
        /// 保存报表数据到文件
        /// </summary>
        private void SaveReport()
        {
            try
            {                
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                    FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                    FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                }
                LoadDataModel();
                string datafile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                SerializeUtilities.Serialize<ShortFormReport>(shortFormReportModel, datafile);
                MessageBox.Show("Report is saved successfully.");
            }
            catch (Exception ex)
            {
                FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                MessageBox.Show("Failed to save the report. Error: " + ex.Message);
            }
        }

        /// <summary>
        /// 加载报表数据到模型对象
        /// </summary>
        private void LoadDataModel()
        {
            shortFormReportModel.DataUserCode = this.dataUserCode.Text;
            shortFormReportModel.DataAge = this.dataAge.Text;
            shortFormReportModel.DataBiRadsCategory = this.dataBiRadsCategory.Text;
            shortFormReportModel.DataComments = this.dataComments.Text;
            shortFormReportModel.DataConclusion = this.dataConclusion.Text;
            shortFormReportModel.DataConclusion2 = this.dataConclusion2.Text;
            shortFormReportModel.DataDoctor = this.dataDoctor.Text;
            shortFormReportModel.DataFurtherExam = this.dataFurtherExam.Text;
            
            shortFormReportModel.DataLeftLocation = this.dataLeftLocation.Text;
            shortFormReportModel.DataLeftSize = this.dataLeftSize.Text;
            shortFormReportModel.DataMeikTech = this.dataMeikTech.Text;
            shortFormReportModel.DataName = this.dataName.Text;
            shortFormReportModel.DataRecommendation = this.dataRecommendation.Text;
            
            shortFormReportModel.DataRightLocation = this.dataRightLocation.Text;
            shortFormReportModel.DataRightSize = this.dataRightSize.Text;
            shortFormReportModel.DataScreenDate = this.dataScreenDate.Text;
            shortFormReportModel.DataScreenLocation = this.dataScreenLocation.Text;
            
            shortFormReportModel.DataAddress = this.dataAddress.Text;
            shortFormReportModel.DataGender = this.dataGender.Text;
            shortFormReportModel.DataHealthCard = this.dataHealthCard.Text;
            shortFormReportModel.DataWeight = this.dataWeight.Text;
            shortFormReportModel.DataWeightUnit = this.dataWeightUnit.Text;
            shortFormReportModel.DataMenstrualCycle = this.dataMenstrualCycle.Text;
            shortFormReportModel.DataHormones = this.dataHormones.Text;
            shortFormReportModel.DataSkinAffections = this.dataSkinAffections.Text;
            shortFormReportModel.DataPertinentHistory = this.dataPertinentHistory.Text;
            shortFormReportModel.DataPertinentHistory1 = this.dataPertinentHistory1.Text;
            shortFormReportModel.DataMotherUltra = this.dataMotherUltra.Text;
            shortFormReportModel.DataLeftBreast = this.dataLeftBreast.Text;
            shortFormReportModel.DataRightBreast = this.dataRightBreast.Text;
            shortFormReportModel.DataLeftPalpableMass = this.dataLeftPalpableMass.Text;
            shortFormReportModel.DataRightPalpableMass = this.dataRightPalpableMass.Text;
            shortFormReportModel.DataLeftChangesOfElectricalConductivity = this.dataLeftChangesOfElectricalConductivity.Text;
            shortFormReportModel.DataRightChangesOfElectricalConductivity = this.dataRightChangesOfElectricalConductivity.Text;
            shortFormReportModel.DataLeftMammaryStruct = this.dataLeftMammaryStruct.Text;
            shortFormReportModel.DataRightMammaryStruct = this.dataRightMammaryStruct.Text;
            shortFormReportModel.DataLeftLactiferousSinusZone = this.dataLeftLactiferousSinusZone.Text;
            shortFormReportModel.DataRightLactiferousSinusZone = this.dataRightLactiferousSinusZone.Text;
            shortFormReportModel.DataLeftMammaryContour = this.dataLeftMammaryContour.Text;
            shortFormReportModel.DataRightMammaryContour = this.dataRightMammaryContour.Text;

            shortFormReportModel.DataLeftNumber = this.dataLeftNumber.Text;
            shortFormReportModel.DataRightNumber = this.dataRightNumber.Text;
            shortFormReportModel.DataLeftShape = this.dataLeftShape.Text;
            shortFormReportModel.DataRightShape = this.dataRightShape.Text;
            shortFormReportModel.DataLeftContourAroundFocus = this.dataLeftContourAroundFocus.Text;
            shortFormReportModel.DataRightContourAroundFocus = this.dataRightContourAroundFocus.Text;
            shortFormReportModel.DataLeftInternalElectricalStructure = this.dataLeftInternalElectricalStructure.Text;
            shortFormReportModel.DataRightInternalElectricalStructure = this.dataRightInternalElectricalStructure.Text;
            shortFormReportModel.DataLeftSurroundingTissues = this.dataLeftSurroundingTissues.Text;
            shortFormReportModel.DataRightSurroundingTissues = this.dataRightSurroundingTissues.Text;
            shortFormReportModel.DataLeftOncomarkerHighlightBenignChanges = this.dataLeftOncomarkerHighlightBenignChanges.Text;
            shortFormReportModel.DataRightOncomarkerHighlightBenignChanges = this.dataRightOncomarkerHighlightBenignChanges.Text;
            shortFormReportModel.DataLeftOncomarkerHighlightSuspiciousChanges = this.dataLeftOncomarkerHighlightSuspiciousChanges.Text;
            shortFormReportModel.DataRightOncomarkerHighlightSuspiciousChanges = this.dataRightOncomarkerHighlightSuspiciousChanges.Text;
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

            shortFormReportModel.DataMeanElectricalConductivity3 = this.dataMeanElectricalConductivity3.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity3 = this.dataLeftMeanElectricalConductivity3.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity3N1 = this.dataLeftMeanElectricalConductivity3N1.Text;
            shortFormReportModel.DataLeftMeanElectricalConductivity3N2 = this.dataLeftMeanElectricalConductivity3N2.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3 = this.dataRightMeanElectricalConductivity3.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3N1 = this.dataRightMeanElectricalConductivity3N1.Text;
            shortFormReportModel.DataRightMeanElectricalConductivity3N2 = this.dataRightMeanElectricalConductivity3N2.Text;

            shortFormReportModel.DataComparativeElectricalConductivityReference1 = this.dataComparativeElectricalConductivityReference1.Text;
            shortFormReportModel.DataLeftComparativeElectricalConductivity1 = this.dataLeftComparativeElectricalConductivity1.Text;
            shortFormReportModel.DataRightComparativeElectricalConductivity1 = this.dataRightComparativeElectricalConductivity1.Text;
            shortFormReportModel.DataComparativeElectricalConductivityReference2 = this.dataComparativeElectricalConductivityReference2.Text;
            shortFormReportModel.DataLeftComparativeElectricalConductivity2 = this.dataLeftComparativeElectricalConductivity2.Text;
            shortFormReportModel.DataRightComparativeElectricalConductivity2 = this.dataRightComparativeElectricalConductivity2.Text;
            shortFormReportModel.DataComparativeElectricalConductivity3 = this.dataComparativeElectricalConductivity3.Text;
            shortFormReportModel.DataLeftComparativeElectricalConductivity3 = this.dataLeftComparativeElectricalConductivity3.Text;
            shortFormReportModel.DataRightComparativeElectricalConductivity3 = this.dataRightComparativeElectricalConductivity3.Text;

            shortFormReportModel.DataDivergenceBetweenHistogramsReference1 = this.dataDivergenceBetweenHistogramsReference1.Text;
            shortFormReportModel.DataLeftDivergenceBetweenHistograms1 = this.dataLeftDivergenceBetweenHistograms1.Text;
            shortFormReportModel.DataRightDivergenceBetweenHistograms1 = this.dataRightDivergenceBetweenHistograms1.Text;
            shortFormReportModel.DataDivergenceBetweenHistogramsReference2 = this.dataDivergenceBetweenHistogramsReference2.Text;
            shortFormReportModel.DataLeftDivergenceBetweenHistograms2 = this.dataLeftDivergenceBetweenHistograms2.Text;
            shortFormReportModel.DataRightDivergenceBetweenHistograms2 = this.dataRightDivergenceBetweenHistograms2.Text;
            shortFormReportModel.DataDivergenceBetweenHistograms3 = this.dataDivergenceBetweenHistograms3.Text;
            shortFormReportModel.DataLeftDivergenceBetweenHistograms3 = this.dataLeftDivergenceBetweenHistograms3.Text;
            shortFormReportModel.DataRightDivergenceBetweenHistograms3 = this.dataRightDivergenceBetweenHistograms3.Text;

            shortFormReportModel.DataLeftComparisonWithNorm = this.dataLeftComparisonWithNorm.Text;
            shortFormReportModel.DataRightComparisonWithNorm = this.dataRightComparisonWithNorm.Text;

            shortFormReportModel.DataPhaseElectricalConductivityReference = this.dataPhaseElectricalConductivityReference.Text;
            shortFormReportModel.DataLeftPhaseElectricalConductivity = this.dataLeftPhaseElectricalConductivity.Text;
            shortFormReportModel.DataRightPhaseElectricalConductivity = this.dataRightPhaseElectricalConductivity.Text;

            shortFormReportModel.DataAgeElectricalConductivityReference = this.dataAgeElectricalConductivityReference.Text;
            shortFormReportModel.DataLeftAgeElectricalConductivity = this.dataLeftAgeElectricalConductivity.Text;
            shortFormReportModel.DataRightAgeElectricalConductivity = this.dataRightAgeElectricalConductivity.Text;
            shortFormReportModel.DataAgeValueOfEC = this.dataAgeValueOfEC.Text;

            shortFormReportModel.DataExamConclusion = this.dataExamConclusion.Text;
            shortFormReportModel.DataLeftMammaryGland = this.dataLeftMammaryGland.Text;
            shortFormReportModel.DataLeftAgeRelated = this.dataLeftAgeRelated.Text;
            shortFormReportModel.DataLeftMeanECOfLesion = this.dataLeftMeanECOfLesion.Text;
            shortFormReportModel.DataLeftFindings = this.dataLeftFindings.Text;
            shortFormReportModel.DataLeftMammaryGlandResult = this.dataLeftMammaryGlandResult.Text;

            shortFormReportModel.DataRightMammaryGland = this.dataRightMammaryGland.Text;
            shortFormReportModel.DataRightAgeRelated = this.dataRightAgeRelated.Text;
            shortFormReportModel.DataRightMeanECOfLesion = this.dataRightMeanECOfLesion.Text;
            shortFormReportModel.DataRightFindings = this.dataRightFindings.Text;
            shortFormReportModel.DataRightMammaryGlandResult = this.dataRightMammaryGlandResult.Text;

            shortFormReportModel.DataTotalPts = this.dataTotalPts.Text;
            shortFormReportModel.DataPoint = this.dataPoint.Text;
            
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

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Do you want to save the report?", "Save Report", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                SaveReport();
            }
        }

        private void ShowSignature(Object obj)
        {                     
            dataSignImg.Source = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg");
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
                //string xpsFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                string userTempPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                string xpsFile = userTempPath + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                if (File.Exists(xpsFile))
                {
                    File.Delete(xpsFile);
                }
                FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender("Views/ExaminationReportDocument.xaml", shortFormReportModel);
                XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
                //将flow document写入基于内存的xps document中去
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
                writer.Write(page);
                xpsDocument.Close();
                var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "pdf|*.pdf" };
                if (dlg.ShowDialog(this) == true)
                {
                    PDFTools.SavePDFFile(xpsFile, dlg.FileName);
                    MessageBox.Show("Exported the PDF file successfully.");
                }
            }
            catch (Exception ex)
            {
                FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 窗口大小状态变化时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_StateChanged(object sender, EventArgs e)
        {
            try { 
                var mainWin = this.Owner.Owner as MainWindow;
                if (this.WindowState == WindowState.Maximized)
                {
                    mainWin.StopMouseHook();
                }
                if (this.WindowState == WindowState.Minimized)
                {
                    WinMinimized();
                    IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
                    //如果主窗体不存在
                    if (mainWinHwnd != IntPtr.Zero)
                    {
                        mainWin.StartMouseHook();
                    }
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
        }
    }
}

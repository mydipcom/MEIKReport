using MEIKReport.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
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
    //定义委托
    //public delegate void CloseWindowHandler(object sender, EventArgs e);    
    /// <summary>
    /// Page1.xaml 的交互逻辑
    /// </summary>
    public partial class SummaryReportPage : Window
    {
        private delegate void DoPrintMethod(PrintDialog pdlg, DocumentPaginator paginator);
        //public event CloseWindowHandler closeWindowEvent;       

        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "/Data";
        private Person person = null;
        private ShortFormReport shortFormReportModel = new ShortFormReport();
        public SummaryReportPage()
        {
            InitializeComponent();            
        }
        public SummaryReportPage(object data): this()
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
                    string dataFile=dataFolder+"/"+person.Code+".dat";
                    if(File.Exists(dataFile)){
                        ////序列化xaml
                        //using (FileStream fs = new FileStream(dataFolder+"/"+person.Code+".dat", FileMode.Open))
                        //{                        
                        //    var scrollViewer = XamlReader.Load(fs) as ScrollViewer;                        
                        //    this.reportPage = scrollViewer;
                        //}
                        this.shortFormReportModel=SerializeUtilities.Desrialize<ShortFormReport>(dataFile);
                                        
                        if (shortFormReportModel.DataScreenShotImg != null)
                        {
                            this.dataScreenShotImg.Source = ImageTools.GetBitmapImage(shortFormReportModel.DataScreenShotImg);
                        }
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
                        shortFormReportModel.DataScreenDate = DateTime.Parse(person.Regdate).ToLongDateString();
                        shortFormReportModel.DataSignDate = DateTime.Today.ToLongDateString();
                        bool defaultSign = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                        if (defaultSign)
                        {
                            string imgFile = AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg";
                            if (File.Exists(imgFile))
                            {
                                dataSignImg.Source = ImageTools.GetBitmapImage(imgFile);
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
                //if (closeWindowEvent != null)
                //{
                //    closeWindowEvent(sender, e);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void DoPrint(PrintDialog pdlg, DocumentPaginator paginator)
        {
            pdlg.PrintDocument(paginator, "Summary Report Document");
        }

        private void previewBtn_Click(object sender, RoutedEventArgs e)
        {
            try { 
                LoadDataModel();
                string reportTempl = "Views/SummaryReportDocument.xaml";
                if (this.dataScreenShotImg.Source != null)
                {
                    reportTempl = "Views/SummaryReportImageDocument.xaml";
                }

                PrintPreviewWindow previewWnd = new PrintPreviewWindow(reportTempl, true, shortFormReportModel);
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
                    string reportTempl = "Views/SummaryReportDocument.xaml";
                    if (this.dataScreenShotImg.Source != null)
                    {
                        reportTempl = "Views/SummaryReportImageDocument.xaml";
                    }
                    FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender(reportTempl, shortFormReportModel);
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

        private void save_Click(object sender, RoutedEventArgs e)
        {
            SaveReport();          
            //using (FileStream fs = new FileStream(dataFolder+"/"+person.Code + ".dat", FileMode.Create))
            //{
            //    XamlWriter.Save(this.reportPage, fs);
            //    MessageBox.Show("Data is saved successfully.");
            //}
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
                    this.CreateFolder(dataFolder);
                }
                LoadDataModel();
                string datafile = dataFolder + "/" + person.Code + ".dat";
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
            shortFormReportModel.DataAddress = person.Address;
            shortFormReportModel.DataBiRadsCategory = this.dataBiRadsCategory.Text;            
            shortFormReportModel.DataComments = this.dataComments.Text;
            shortFormReportModel.DataConclusion = this.dataConclusion.Text;
            shortFormReportModel.DataConclusion2 = this.dataConclusion2.Text;
            shortFormReportModel.DataDoctor = this.dataDoctor.Text;
            shortFormReportModel.DataFurtherExam = this.dataFurtherExam.Text;
            shortFormReportModel.DataLeftFinding = this.dataLeftFinding.Text;
            shortFormReportModel.DataLeftLocation = this.dataLeftLocation.Text;
            shortFormReportModel.DataLeftSize = this.dataLeftSize.Text;
            shortFormReportModel.DataMeikTech = this.dataMeikTech.Text;
            shortFormReportModel.DataName = this.dataName.Text;
            shortFormReportModel.DataRecommendation = this.dataRecommendation.Text;
            shortFormReportModel.DataRightFinding = this.dataRightFinding.Text;
            shortFormReportModel.DataRightLocation = this.dataRightLocation.Text;
            shortFormReportModel.DataRightSize = this.dataRightSize.Text;
            shortFormReportModel.DataScreenDate = this.dataScreenDate.Text;
            shortFormReportModel.DataScreenLocation = this.dataScreenLocation.Text;
            var screenShotImg = this.dataScreenShotImg.Source as BitmapImage;
            if (screenShotImg != null)
            {
                var stream = screenShotImg.StreamSource;                
                stream.Position = 0;
                byte[] buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                stream.Flush();                
                shortFormReportModel.DataScreenShotImg = buffer;
            }

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
            shortFormReportModel.DataSignDate = this.dataSignDate.Text;
        }
        private void ShowSignature(Object obj)
        {           
            //this.inkCanvas.Strokes = (StrokeCollection)obj;
            ////scaleTrans.CenterX = this.inkCanvas.Width / 2;
            //scaleTrans.ScaleX = 0.5;
            //scaleTrans.ScaleY = 0.5;
            //this.inkCanvas.Strokes = ((InkCanvas)obj).Strokes;            
            dataSignImg.Source = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg");
        }            

        private void btnScreenShot_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, "Diagnostics");
                Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                this.WindowState = WindowState.Minimized;
                App.opendWin = this;
                ScreenCapture screenCaptureWin = new ScreenCapture(this.person);
                screenCaptureWin.callbackMethod = LoadScreenShot;
                screenCaptureWin.ShowDialog();                
            }
            catch (Exception ex)
            {
                MessageBox.Show("System Exception: " + ex.Message);
            }            
        }

        private void LoadScreenShot(Object imgFileName)
        {
            dataScreenShotImg.Source = ImageTools.GetBitmapImage(imgFileName as string);            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Do you want to save the report?", "Save Report", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                SaveReport();                
            }
        }

        private void btnRemoveImg_Click(object sender, RoutedEventArgs e)
        {
            this.dataScreenShotImg.Source = null;            
        }

        private void savePdfBtn_Click(object sender, RoutedEventArgs e)
        {
            try { 
                if (!Directory.Exists(dataFolder))
                {
                    this.CreateFolder(dataFolder);                    
                }
                string xpsFile = dataFolder + "/" + person.Code + ".xps";
                if (File.Exists(xpsFile)) {
                    File.Delete(xpsFile);
                }
                string reportTempl = "Views/SummaryReportDocument.xaml";
                if (this.dataScreenShotImg.Source != null)
                {
                    reportTempl = "Views/SummaryReportImageDocument.xaml";
                }
                FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender(reportTempl, shortFormReportModel);
                XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
                //将flow document写入基于内存的xps document中去
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);            
                writer.Write(page);            
                xpsDocument.Close();
                var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "pdf|*.pdf" };
                if (dlg.ShowDialog(this) == true)
                {
                    PDFTools.SavePDFFile(xpsFile, dlg.FileName);
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
        /// 打开MEIK程序的诊断窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                else
                {
                    //Win32Api.ShowWindow(mainWinHwnd, 1);
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

        ///// <summary>
        ///// 打开MEIK程序的扫描窗口
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void btnOpenScreen_Click(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {                
        //        IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
        //        //如果主窗体不存在
        //        if (mainWinHwnd == IntPtr.Zero)
        //        {
        //            IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, "Screening");
        //            Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
        //        }
        //        else
        //        {
        //            //Win32Api.ShowWindow(mainWinHwnd, 1);
        //        }
        //        WinMinimized();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("System Exception: " + ex.Message);
        //    }
        //}

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
            catch(Exception ex){
                MessageBox.Show(ex.Message);
            }
        }

        private void CreateFolder(string folderPath)
        {
            Directory.CreateDirectory(folderPath);
            FileHelper.SetFolderPower(folderPath, "Everyone", "FullControl");
            FileHelper.SetFolderPower(folderPath, "Users", "FullControl");
        }
        
    }
}

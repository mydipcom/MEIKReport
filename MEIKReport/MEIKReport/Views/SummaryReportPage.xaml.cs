using MEIKReport.Common;
using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading;
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

        private string dataFolder = AppDomain.CurrentDomain.BaseDirectory + "Data";
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
                    MessageBox.Show(this, App.Current.FindResource("Message_8").ToString());
                    this.Close();
                }
                else
                {
                    string dataFile = FindUserReportData(person.ArchiveFolder);
                    if (string.IsNullOrEmpty(dataFile))
                    {
                        dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
                    }
                    if ("en-US".Equals(App.local))
                    {
                        dataScreenDate.FormatString = "MMMM d, yyyy";
                        dataSignDate.FormatString = "MMMM d, yyyy";                        
                    }
                    else
                    {
                        dataScreenDate.FormatString = "yyyy年MM月dd日";
                        dataSignDate.FormatString = "yyyy年MM月dd日";
                    }
                    //string dataFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat";
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
                        shortFormReportModel.DataScreenDate = DateTime.Parse(person.RegMonth + "/" + person.RegDate + "/" + person.RegYear).ToLongDateString();
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
                    //以下是添加处理操作员和医生的名字的选择项                    
                    User doctorUser = new User();
                    if (!string.IsNullOrEmpty(shortFormReportModel.DataDoctor))
                    {
                        doctorUser.Name = shortFormReportModel.DataDoctor;
                        doctorUser.License = shortFormReportModel.DataDoctorLicense;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(this.person.DoctorName))
                        {
                            doctorUser.Name = this.person.DoctorName;
                            doctorUser.License = this.person.DoctorLicense;
                            shortFormReportModel.DataDoctor = this.person.DoctorName;
                            shortFormReportModel.DataDoctorLicense = this.person.DoctorLicense;
                        }
                    }
                    this.dataDoctor.ItemsSource = App.reportSettingModel.DoctorNames;
                    if (!string.IsNullOrEmpty(doctorUser.Name))
                    {
                        for (int i = 0; i < App.reportSettingModel.DoctorNames.Count; i++)
                        {
                            var item = App.reportSettingModel.DoctorNames[i];
                            if (doctorUser.Name.Equals(item.Name) && (string.IsNullOrEmpty(doctorUser.License) == string.IsNullOrEmpty(item.License) || doctorUser.License == item.License))
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
            App.opendWin = null;
            this.Owner.Show();
            //try { 
            //    App.opendWin = null;
            //    IntPtr mainWinHwnd = Win32Api.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "TfmMain", null);
            //    //如果主窗体存在
            //    if (mainWinHwnd != IntPtr.Zero)
            //    {
            //        int WM_SYSCOMMAND = 0x0112;
            //        int SC_CLOSE = 0xF060;
            //        Win32Api.SendMessage(mainWinHwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
            //    }
            //    this.Owner.Show();
            //    //if (closeWindowEvent != null)
            //    //{
            //    //    closeWindowEvent(sender, e);
            //    //}
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
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
                var reportModel = CloneReportModel();
                PrintPreviewWindow previewWnd = new PrintPreviewWindow(reportTempl, true, reportModel);                
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
                    string reportTempl = "Views/SummaryReportDocument.xaml";
                    if (this.dataScreenShotImg.Source != null)
                    {
                        reportTempl = "Views/SummaryReportImageDocument.xaml";
                    }
                    var reportModel = CloneReportModel();
                    FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender(reportTempl, reportModel);
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

        private void save_Click(object sender, RoutedEventArgs e)
        {
            SaveReport(null);          
            //using (FileStream fs = new FileStream(dataFolder+"/"+person.Code + ".dat", FileMode.Create))
            //{
            //    XamlWriter.Save(this.reportPage, fs);
            //    MessageBox.Show(this,"Data is saved successfully.");
            //}
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
                string datafile = null;
                if (!string.IsNullOrEmpty(otherDataFolder))
                {
                    datafile = otherDataFolder + "/" + person.Code + ".dat";
                }
                else 
                {
                    if (!Directory.Exists(dataFolder))
                    {
                        this.CreateFolder(dataFolder);
                    }
                    datafile = dataFolder + "/" + person.Code + ".dat";
                }
                
                LoadDataModel();
                var reportModel = CloneReportModel();
                SerializeUtilities.Serialize<ShortFormReport>(shortFormReportModel, datafile);
                File.Copy(datafile, person.ArchiveFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".dat", true);

                //Save PDF file
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
                MessageBox.Show(this, "Report is saved successfully.");
            }
            catch (Exception ex)
            {
                FileHelper.SetFolderPower(dataFolder, "Everyone", "FullControl");
                FileHelper.SetFolderPower(dataFolder, "Users", "FullControl");
                MessageBox.Show(this, App.Current.FindResource("Message_3").ToString() + " " + ex.Message);
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
            //shortFormReportModel.DataBiRadsCategory = this.dataBiRadsCategory.Text;            
            shortFormReportModel.DataComments = this.dataComments.Text;
            shortFormReportModel.DataConclusion = this.dataConclusion.SelectedIndex.ToString();
            shortFormReportModel.DataFurtherExam = this.dataFurtherExam.SelectedIndex.ToString();
            shortFormReportModel.DataRecommendation = this.dataRecommendation.SelectedIndex.ToString();
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
            //shortFormReportModel.DataLeftFinding = this.dataLeftFinding.Text;
            shortFormReportModel.DataLeftLocation = this.dataLeftLocation.Text;
            shortFormReportModel.DataLeftSize = this.dataLeftSize.Text;
            
            shortFormReportModel.DataName = this.dataName.Text;
            
            //shortFormReportModel.DataRightFinding = this.dataRightFinding.Text;
            shortFormReportModel.DataRightLocation = this.dataRightLocation.Text;
            shortFormReportModel.DataRightSize = this.dataRightSize.Text;
            shortFormReportModel.DataScreenDate = this.dataScreenDate.Text;
            shortFormReportModel.DataScreenLocation = this.dataScreenLocation.Text;
            shortFormReportModel.DataLeftMammaryGlandResult = this.dataLeftMammaryGlandResult.SelectedIndex.ToString();
            shortFormReportModel.DataRightMammaryGlandResult = this.dataRightMammaryGlandResult.SelectedIndex.ToString();
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

        /// <summary>
        /// 克隆数据模型对象
        /// </summary>
        private ShortFormReport CloneReportModel()
        {
            ShortFormReport reportModel = shortFormReportModel.Clone();
            reportModel = shortFormReportModel.Clone();
            reportModel.DataRecommendation = this.dataRecommendation.Text;
            reportModel.DataConclusion = this.dataConclusion.Text;
            reportModel.DataFurtherExam = this.dataFurtherExam.Text;
            reportModel.DataRightMammaryGlandResult = this.dataRightMammaryGlandResult.Text;            

            return reportModel;
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
                //IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, "Diagnostics");
                //Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
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
            dataScreenShotImg.Source = ImageTools.GetBitmapImage(imgFileName as string);            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(this, App.Current.FindResource("Message_4").ToString(), "Save Report", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                SaveReport(null);                
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
                LoadDataModel();
                //string tempPath=Path.GetTempPath().ToString();
                //string userTempPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                //string xpsFile = userTempPath + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                string xpsFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + ".xps";
                
                if (File.Exists(xpsFile))
                {
                    File.Delete(xpsFile);
                }
                string reportTempl = "Views/SummaryReportDocument.xaml";
                if (this.dataScreenShotImg.Source != null)
                {
                    reportTempl = "Views/SummaryReportImageDocument.xaml";
                }
                var reportModel = CloneReportModel();
                FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender(reportTempl, reportModel);
                XpsDocument xpsDocument = new XpsDocument(xpsFile, FileAccess.ReadWrite);
                //将flow document写入基于内存的xps document中去
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
                writer.Write(page);            
                xpsDocument.Close();
                
                //string xpsFile = "C:\\MEIKReport\\14102500104.xps";
                var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "pdf|*.pdf" };
                if (dlg.ShowDialog(this) == true)
                {
                    //string pdfFile = dataFolder + System.IO.Path.DirectorySeparatorChar + person.Code + "_SF.pdf";
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

        private void ExportPDF(string reportTempl, string pdfFile, ShortFormReport reportModel)
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
                    IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, App.strDiagnostics);
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
                MessageBox.Show(this, "System Exception: " + ex.Message);
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
        //        MessageBox.Show(this,"System Exception: " + ex.Message);
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
                MessageBox.Show(this, ex.Message);
            }
        }

        private void CreateFolder(string folderPath)
        {
            Directory.CreateDirectory(folderPath);
            FileHelper.SetFolderPower(folderPath, "Everyone", "FullControl");
            FileHelper.SetFolderPower(folderPath, "Users", "FullControl");
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
                        if ((person.Code+".dat").Equals(NextFile.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            return NextFile.FullName;
                        }
                    }
                }
                catch (Exception) { }
            }
            return null;
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
        
    }
}

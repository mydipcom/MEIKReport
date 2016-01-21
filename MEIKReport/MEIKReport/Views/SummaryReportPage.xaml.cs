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
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
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
                    this.reportDataGrid.DataContext = this.shortFormReportModel;
                    this.dataScreenDate.Text = shortFormReportModel.DataScreenDate;
                    this.dataSignDate.Text = shortFormReportModel.DataSignDate;
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
                    this.dataUserCode.Text = person.Code;
                    this.dataName.Text = person.SurName;
                    this.dataAge.Text = person.Age + "";
                    this.dataScreenDate.Value = DateTime.Parse(person.Regdate);
                    bool defaultSign = Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
                    if (defaultSign)
                    {
                        dataSignImg.Source = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg");
                        //dataScreenShotImg.Source = GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "/Images/BigIcon.png");
                    }
                    
                }
            }

            
        }

        private void Window_Closed(object sender, EventArgs e)
        {            
            this.Owner.Show();
            //if (closeWindowEvent != null)
            //{
            //    closeWindowEvent(sender, e);
            //}
        }
        private void DoPrint(PrintDialog pdlg, DocumentPaginator paginator)
        {
            pdlg.PrintDocument(paginator, "Summary Report Document");
        }

        private void previewBtn_Click(object sender, RoutedEventArgs e)
        {
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

        private void printBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadDataModel();
            PrintDialog pdlg = new PrintDialog();
            if (pdlg.ShowDialog() == true)
            {
                FixedPage page = (FixedPage)PrintPreviewWindow.LoadFixedDocumentAndRender("Views/SummaryReportDocument.xaml", shortFormReportModel);
                FixedDocument fixedDoc = new FixedDocument();//创建一个文档
                fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.5, 96 * 11);

                PageContent pageContent = new PageContent();
                ((IAddChild)pageContent).AddChild(page);
                fixedDoc.Pages.Add(pageContent);//将对象加入到当前文档中
                Dispatcher.BeginInvoke(new DoPrintMethod(DoPrint), DispatcherPriority.ApplicationIdle, pdlg, fixedDoc.DocumentPaginator);
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
                dataFolder = AppDomain.CurrentDomain.BaseDirectory + "/Data";
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
                LoadDataModel();
                string datafile = dataFolder + "/" + person.Code + ".dat";
                SerializeUtilities.Serialize<ShortFormReport>(shortFormReportModel, datafile);
                MessageBox.Show("Report is saved successfully.");
            }
            catch (Exception ex)
            {
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
      
        private void btnOpenDiagn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IntPtr diagnosticsBtnHwnd = Win32Api.FindWindowEx(App.splashWinHwnd, IntPtr.Zero, null, "Diagnostics");
                Win32Api.SendMessage(diagnosticsBtnHwnd, Win32Api.WM_CLICK, 0, 0);
                this.btnOpenDiagn.IsEnabled = false;
                this.WindowState = WindowState.Minimized;
                App.opendWin = this;
            }
            catch (Exception ex)
            {
                MessageBox.Show("System Exception: " + ex.Message);
            }
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
            this.btnRemoveImg.IsEnabled = true;
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
            this.btnRemoveImg.IsEnabled = false;
        }
        
    }
}

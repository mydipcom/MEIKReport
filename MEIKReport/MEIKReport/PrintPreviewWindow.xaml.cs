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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace MEIKReport
{
    /// <summary>
    /// PrintPreviewWindow.xaml 的交互逻辑
    /// </summary>
    public partial class PrintPreviewWindow : Window
    {
        private delegate void LoadXpsMethod();
        private readonly Object m_data;
        private readonly Object m_doc;
        private readonly bool isFixedPage;
        private MemoryStream ms = new MemoryStream();

        /// <summary>
        /// 使用FixedPage模板并加载渲染数据
        /// </summary>
        /// <param name="strTmplName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static FixedPage LoadFixedDocumentAndRender(string strTmplName, Object data)
        {
            FixedPage doc = (FixedPage)Application.LoadComponent(new Uri(strTmplName, UriKind.RelativeOrAbsolute));
            if ("Letter".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                doc.Width = 96 * 8.5;
                doc.Height = 96 * 11;
            }
            else if ("A4".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                doc.Width = 96 * 8.27;
                doc.Height = 96 * 11.69;                
            }
            doc.DataContext = data;
            //加载图片到文档中
            if (data.GetType() == typeof(ShortFormReport))
            {
                var shortFormReport = data as ShortFormReport;
                if (shortFormReport != null)
                {
                     
                    var topImage = doc.FindName("imgTitleLog") as Image;
                    if (topImage != null)
                    {
                        if (!App.reportSettingModel.DefaultLogo)
                        {
                            topImage.Source = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "logo.png");
                        }                        
                    }
                    var footTxt = doc.FindName("footTxt") as TextBlock;
                    if (footTxt != null)
                    {
                        footTxt.Text = App.reportSettingModel.CompanyAddress;
                    }
                    
                    var signImage = doc.FindName("dataSignImg") as Image;
                    if (signImage != null && shortFormReport.DataSignImg!=null)
                    {
                        signImage.Source = ImageTools.GetBitmapImage(shortFormReport.DataSignImg);
                    }
                    var screenShotImg = doc.FindName("dataScreenShotImg") as Image;
                    if (screenShotImg != null && shortFormReport.DataScreenShotImg != null)
                    {
                        screenShotImg.Source = ImageTools.GetBitmapImage(shortFormReport.DataScreenShotImg);
                    }
                    if (!App.reportSettingModel.ShowTechSignature)
                    {
                        var techSignPanel = doc.FindName("techSignPanel") as Panel;
                        if (techSignPanel != null)
                        {
                            techSignPanel.Visibility = Visibility.Collapsed;
                        }
                    }
                    if (!App.reportSettingModel.ShowDoctorSignature)
                    {
                        var doctorSignPanel = doc.FindName("doctorSignPanel") as Panel;
                        if (doctorSignPanel != null)
                        {
                            doctorSignPanel.Visibility = Visibility.Collapsed;
                        }
                        var doctorSignGrid = doc.FindName("doctorSignGrid") as Panel;
                        if (doctorSignGrid != null)
                        {
                            doctorSignGrid.Visibility = Visibility.Collapsed;
                        }
                    }

                }
            }            
            return doc;
        }
        /// <summary>
        /// 使用FlowDocument模板并加载渲染数据
        /// </summary>
        /// <param name="strTmplName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static FlowDocument LoadFlowDocumentAndRender(string strTmplName, Object data,string pagesize=null)
        {
            FlowDocument doc = (FlowDocument)Application.LoadComponent(new Uri(strTmplName, UriKind.RelativeOrAbsolute));
            if (string.IsNullOrEmpty(pagesize))
            {
                if ("Letter".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    //doc.PagePadding = new Thickness(20);
                    doc.PageWidth = 96 * 8.5;
                    doc.PageHeight = 96 * 11;
                    doc.ColumnWidth = 734;
                }
                else if ("A4".Equals(App.reportSettingModel.PrintPaper.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    //doc.PagePadding = new Thickness(20);
                    doc.PageWidth = 96 * 8.27;
                    doc.PageHeight = 96 * 11.69;
                    doc.ColumnWidth = 734;
                }
            }
            else
            {
                if ("Letter".Equals(pagesize, StringComparison.OrdinalIgnoreCase))
                {
                    //doc.PagePadding = new Thickness(20);
                    doc.PageWidth = 96 * 8.5;
                    doc.PageHeight = 96 * 11;
                    doc.ColumnWidth = 734;
                }
                else if ("A4".Equals(pagesize, StringComparison.OrdinalIgnoreCase))
                {
                    //doc.PagePadding = new Thickness(20);
                    doc.PageWidth = 96 * 8.27;
                    doc.PageHeight = 96 * 11.69;
                    doc.ColumnWidth = 734;
                }
            }
            doc.DataContext = data;            

            var shortFormReport = data as ShortFormReport;
            if (shortFormReport != null)
            {                
                var signImage = doc.FindName("dataSignImg") as Image;
                if (signImage != null && shortFormReport.DataSignImg != null)
                {
                    signImage.Source = ImageTools.GetBitmapImage(shortFormReport.DataSignImg);
                }
                var screenShotImg = doc.FindName("dataScreenShotImg") as Image;
                if (screenShotImg != null && shortFormReport.DataScreenShotImg != null)
                {
                    screenShotImg.Source = ImageTools.GetBitmapImage(shortFormReport.DataScreenShotImg);
                }
                if (!App.reportSettingModel.ShowTechSignature )
                {
                    var techSignPanel = doc.FindName("techSignPanel") as Panel;
                    if (techSignPanel != null)
                    {
                        techSignPanel.Visibility = Visibility.Collapsed;
                    }
                }
                if (!App.reportSettingModel.ShowDoctorSignature)
                {
                    var doctorSignPanel = doc.FindName("doctorSignPanel") as Panel;
                    if (doctorSignPanel != null)
                    {
                        doctorSignPanel.Visibility = Visibility.Collapsed;
                    }
                    var doctorSignGrid = doc.FindName("doctorSignGrid") as Panel;
                    if (doctorSignGrid != null)
                    {
                        doctorSignGrid.Visibility = Visibility.Collapsed;
                    }
                }

                //替换头部图片
                var titleRow = doc.FindName("titleImg") as TableRow;
                if (titleRow != null && !App.reportSettingModel.DefaultLogo)
                {
                    var brush = new ImageBrush();
                    brush.ImageSource = ImageTools.GetBitmapImage(AppDomain.CurrentDomain.BaseDirectory + "logo.png");
                    brush.AlignmentY = AlignmentY.Top;
                    brush.Stretch = Stretch.Uniform;
                    titleRow.Background = brush;
                }

            }
            return doc;
        }
        public PrintPreviewWindow()
        {
            InitializeComponent();            
        }
        public PrintPreviewWindow(string strTmplName, bool isFixedPage,Object data)
        {
            InitializeComponent();
            m_data = data;
            this.isFixedPage = isFixedPage;
            if (isFixedPage)
            {
                m_doc = LoadFixedDocumentAndRender(strTmplName, data);
            }
            else
            {
                m_doc = LoadFlowDocumentAndRender(strTmplName, data);
            }
            Dispatcher.BeginInvoke(new LoadXpsMethod(LoadXps), DispatcherPriority.ApplicationIdle);
        }        

        public void LoadXps()
        {
            //构造一个基于内存的xps document
            //MemoryStream ms = new MemoryStream();
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
            Uri DocumentUri = new Uri("pack://InMemoryDocument.xps");
            PackageStore.RemovePackage(DocumentUri);
            PackageStore.AddPackage(DocumentUri, package);
            XpsDocument xpsDocument = new XpsDocument(package, CompressionOption.Fast, DocumentUri.AbsoluteUri);

            //将flow document写入基于内存的xps document中去
            XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
            if (this.isFixedPage)
            {
                writer.Write((FixedPage)m_doc);

            }
            else
            {
                writer.Write(((IDocumentPaginatorSource)m_doc).DocumentPaginator);
            }

            //获取这个基于内存的xps document的fixed document
            docViewer.Document = xpsDocument.GetFixedDocumentSequence();
            //关闭基于内存的xps document
            xpsDocument.Close();
        }


        private void Window_Closed(object sender, EventArgs e)
        {
            if (this.ms != null)
            {
                this.ms.Close();
            }
        }
        
    }
}

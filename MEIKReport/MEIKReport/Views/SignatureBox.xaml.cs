﻿using MEIKReport.Common;
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
    public partial class SignatureBox : Window
    {
        public DelegateHelper.commonDelgateMethod callbackMethod;
        private bool defaultSign = false;
        private string signFolder = AppDomain.CurrentDomain.BaseDirectory + "/Signature";
        public SignatureBox()
        {
            InitializeComponent();
            defaultSign =Convert.ToBoolean(OperateIniFile.ReadIniData("Report", "Use Default Signature", "false", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini"));
            cbDefault.IsChecked = defaultSign;
            if (defaultSign)
            {
                using (var stream = System.IO.File.OpenRead(signFolder + "/default.strokes"))
                {
                    this.inkCanvas.Strokes = new System.Windows.Ink.StrokeCollection(stream);
                }
            }
        }                

        private void Window_Closed(object sender, EventArgs e)
        {
            //this.Owner.Show();
            this.Owner.Visibility=Visibility.Visible;
        }

        private void btnSaveSign_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog() { Filter = "strokes|*.strokes" };
            if (dlg.ShowDialog(this) == true)
            {
                using (var stream = System.IO.File.OpenWrite(dlg.FileName))
                {
                    this.inkCanvas.Strokes.Save(stream);
                }
            }

        }

        private void btnClearSignBox_Click(object sender, RoutedEventArgs e)
        {
            this.inkCanvas.Strokes.Clear();
        }

        private void btnLoadSign_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog() { Filter = "strokes|*.strokes" };
            if (dlg.ShowDialog(this) == true)
            {
                using (var stream = System.IO.File.OpenRead(dlg.FileName))
                {
                    this.inkCanvas.Strokes = new System.Windows.Ink.StrokeCollection(stream);
                }
            }

        }
        
        private void SaveBitmap(){
            try
            {
                using (var stream = System.IO.File.OpenWrite(AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.jpg"))
                {
                    Size size = new Size(this.inkCanvas.ActualWidth, this.inkCanvas.ActualHeight);
                    RenderTargetBitmap rt = new RenderTargetBitmap((int)size.Width, (int)size.Height, 96, 96, PixelFormats.Default);

                    rt.Render(this.inkCanvas);
                    var encoder = new JpegBitmapEncoder();
                    //因为有黑边所以裁切一下
                    CroppedBitmap CroppedBitmap = new CroppedBitmap(BitmapFrame.Create(rt), new Int32Rect(1, 1, 795, 155));
                    encoder.Frames.Add(BitmapFrame.Create(CroppedBitmap));
                    //encoder.Frames.Add(BitmapFrame.Create(rt));
                    encoder.Save(stream);
                    stream.Close();
                    
                }
            }
            catch(Exception ex){
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// 确认使用当前签名
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOkSignBox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(signFolder))
                {
                    Directory.CreateDirectory(signFolder);
                }
                if (this.cbDefault.IsChecked == true)
                {
                    using (var stream = System.IO.File.OpenWrite(AppDomain.CurrentDomain.BaseDirectory + "/Signature/default.strokes"))
                    {
                        this.inkCanvas.Strokes.Save(stream);
                    }
                }
                else
                {
                    using (var stream = System.IO.File.OpenWrite(AppDomain.CurrentDomain.BaseDirectory + "/Signature/temp.strokes"))
                    {
                        this.inkCanvas.Strokes.Save(stream);
                    }
                }
                SaveBitmap();
                if (defaultSign != this.cbDefault.IsChecked)
                {
                    OperateIniFile.WriteIniData("Report", "Use Default Signature", this.cbDefault.IsChecked.ToString(), System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
                }
                if (this.callbackMethod != null)
                {                    
                    this.callbackMethod(this.inkCanvas.Strokes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);                
            }
            finally
            {
                this.Close();
            }
            
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        
    }
}
using MEIKReport.Common;
using MEIKReport.Model;
using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    public partial class PasswordPage : Window
    {
        private int _deviceType;
        public PasswordPage(int deviceType)
        {
            InitializeComponent();
            this._deviceType = deviceType;
            this.txtPwd.Focus();
        }                        

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            string filePath = System.AppDomain.CurrentDomain.BaseDirectory + "mode.dat";
            if (_deviceType<3)
            {
               
                if ("nuvomed2016".Equals(this.txtPwd.Password))
                {
                    App.reportSettingModel.DeviceType = _deviceType;
                    if (_deviceType == 1)
                    {
                        SecurityTools.EncryptTextToFile("Technician", filePath);                        
                    }
                    else
                    {
                        SecurityTools.EncryptTextToFile("Doctor", filePath);                           
                    }
                    this.Close();
                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_49").ToString());
                    this.txtPwd.Focus();
                }
                
            }            
            else
            {
                if ("NuvoMED@2016".Equals(this.txtPwd.Password))
                {
                    App.reportSettingModel.DeviceType = _deviceType;
                    SecurityTools.EncryptTextToFile("Admin", filePath);                     
                    this.Close();
                }
                else
                {
                    MessageBox.Show(this, App.Current.FindResource("Message_49").ToString());
                    this.txtPwd.Focus();
                }
            }

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }                
        
    }
}

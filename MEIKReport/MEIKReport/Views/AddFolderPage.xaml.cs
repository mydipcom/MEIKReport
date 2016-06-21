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
    public partial class AddFolderPage : Window
    {
        private string lastCode = OperateIniFile.ReadIniData("Data", "Last Code", "", System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");
        
        public AddFolderPage()
        {
            InitializeComponent();
            this.txtLastName.Focus();
            string dateStr=DateTime.Now.ToString("yyMMdd");
            int num = 1;
            if (!string.IsNullOrEmpty(lastCode))
            {
                try
                {
                    num = Convert.ToInt32(lastCode.Substring(lastCode.Length - 2, 2));
                }
                catch (Exception)
                {
                    num = 50;
                }
                num = lastCode.StartsWith(dateStr) ? (num+1) : 1;
            }

            this.txtPatientCode.Text = dateStr + App.reportSettingModel.DeviceNo + num.ToString("00");            
        }                        

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtLastName.Text))
            {
                MessageBox.Show(this, App.Current.FindResource("Message_22").ToString());
                this.txtLastName.Focus();
            }              
            else
            {
                string dayFolder = App.reportSettingModel.DataBaseFolder + System.IO.Path.DirectorySeparatorChar + DateTime.Now.ToString("MM_yyyy") + System.IO.Path.DirectorySeparatorChar + DateTime.Now.ToString("dd");
                if (!Directory.Exists(dayFolder))
                {
                    Directory.CreateDirectory(dayFolder);                    
                }
                
                string patientFolder = null;
                try
                {
                    patientFolder = dayFolder + System.IO.Path.DirectorySeparatorChar + this.txtPatientCode.Text + "-" + this.txtLastName.Text;
                    if (!string.IsNullOrEmpty(this.txtFirstName.Text))
                    {
                        patientFolder = patientFolder + "," + this.txtFirstName.Text;
                    }                    
                    if (!string.IsNullOrEmpty(this.txtMiddleInitial.Text))
                    {
                        patientFolder = patientFolder + " " + this.txtMiddleInitial.Text;
                    }
                    if (!Directory.Exists(patientFolder))
                    {                        
                        Directory.CreateDirectory(patientFolder);                        
                        OperateIniFile.WriteIniData("Data", "Last Code", this.txtPatientCode.Text, System.AppDomain.CurrentDomain.BaseDirectory + "Config.ini");                                                                                              
                    }
                    else
                    {
                        MessageBox.Show(this, string.Format(App.Current.FindResource("Message_24").ToString(), patientFolder));                       
                    }
                    UserList userlistWin = this.Owner as UserList;
                    userlistWin.loadArchiveFolder(patientFolder);
                    try
                    {
                        Clipboard.SetText(this.txtPatientCode.Text);
                    }
                    catch (Exception) { }
                }
                catch (Exception ex)
                {
                    FileHelper.SetFolderPower(patientFolder, "Everyone", "FullControl");
                    FileHelper.SetFolderPower(patientFolder, "Users", "FullControl");
                    MessageBox.Show(this, App.Current.FindResource("Message_23").ToString());
                    MessageBox.Show(this, string.Format(App.Current.FindResource("Message_25").ToString(), patientFolder, ", Error: " + ex.Message));
                }
                finally
                {
                    this.Close();
                }
            }

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        //private string MatchFolder(string folderName,string code)
        //{            
        //    //遍历指定文件夹下所有文件
        //    DirectoryInfo theFolder = new DirectoryInfo(folderName);
        //    try
        //    {
        //        DirectoryInfo[] folderList = theFolder.GetDirectories();
        //        //遍历文件
        //        foreach (DirectoryInfo NextDir in folderList)
        //        {
        //            if (NextDir.Name.StartsWith(code))
        //            {
        //                return NextDir.FullName;
        //            }
        //        }
        //        return null;
        //    }
        //    catch (Exception) {
        //        return null;
        //    }

            
        //} 
        
    }
}

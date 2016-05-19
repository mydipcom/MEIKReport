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
    public partial class SelectTechnicianPage : Window
    {
        private ObservableCollection<User> _userList;
        private Person _selectUser;
        public SelectTechnicianPage(ObservableCollection<User> userList,Person selectUser)
        {
            InitializeComponent();
            this.listTechnicianName.Focus();
            this._userList = userList;
            this._selectUser = selectUser;
            this.listTechnicianName.ItemsSource = userList;
        }                        

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {

            if (this.listTechnicianName.SelectedIndex==-1)
            {
                MessageBox.Show(App.Current.FindResource("Message_28").ToString());
                this.listTechnicianName.Focus();                
            }
            else
            {
                var user = this.listTechnicianName.SelectedItem as User;
                _selectUser.TechName = user.Name;
                _selectUser.TechLicense = this.technicianLicense.Text;
                this.Close();
            }            

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }                
        
    }
}

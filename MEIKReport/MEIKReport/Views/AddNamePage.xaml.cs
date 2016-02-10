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
    public partial class AddNamePage : Window
    {        
        private ObservableCollection<string> nameList;
        public AddNamePage(ObservableCollection<string> names)
        {
            InitializeComponent();
            this.txtName.Focus();
            this.nameList = names;
        }                        

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtName.Text))
            {
                MessageBox.Show("Please input a name.");
                this.txtName.Focus();
            }
            else
            {
                if (nameList.Contains(this.txtName.Text))
                {
                    MessageBox.Show("This name already exists.");
                }
                else
                {                                        
                    nameList.Add(this.txtName.Text);                        
                    this.Close();                    
                }
            }

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }                
        
    }
}

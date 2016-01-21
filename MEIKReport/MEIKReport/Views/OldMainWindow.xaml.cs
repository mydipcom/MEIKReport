using MEIKReport.Views;
using System;
using System.Collections.Generic;
using System.Data;
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
    public partial class OldMainWindow : Window
    {
        public OldMainWindow()
        {
            InitializeComponent();
        }

        private void equipmentList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {            
           // userGrid.DataContext = equipmentCode.SelectedItem;
        }

        private void ExaminationReport_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            var selectedUser = this.userGrid.SelectedItem as XmlElement;
            var name=selectedUser.GetAttribute("Name");
            // View Examination Report
            ExaminationReportPage examinationReportPage = new ExaminationReportPage(this.userGrid.SelectedItem);
            //examinationReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            examinationReportPage.Owner = this;
            examinationReportPage.ShowDialog();            
        }

        private void SummaryReport_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            var selectedUser = userGrid.SelectedItem as XmlElement;
            var name = selectedUser.GetAttribute("Name");
            SummaryReportPage summaryReportPage = new SummaryReportPage(this.userGrid.SelectedItem);
            //summaryReportPage.closeWindowEvent += new CloseWindowHandler(ShowMainWindow);
            summaryReportPage.Owner = this;
            summaryReportPage.ShowDialog();
        }

        private void ShowMainWindow(object sender, EventArgs e)
        {
            this.Show();
        }
    }
}

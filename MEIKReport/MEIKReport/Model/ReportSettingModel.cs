using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Model
{
    public class ReportSettingModel : ViewModelBase
    {
        private ObservableCollection<string> doctorNames = new ObservableCollection<string>();

        public ObservableCollection<string> DoctorNames
        {
            get { return doctorNames; }
            set { doctorNames = value; OnPropertyChanged("DoctorNames"); }
        }
        private ObservableCollection<string> techNames = new ObservableCollection<string>();

        public ObservableCollection<string> TechNames
        {
            get { return techNames; }
            set { techNames = value; OnPropertyChanged("TechNames"); }
        }
        private string printPaper = "Letter";

        public string PrintPaper
        {
            get { return printPaper; }
            set { printPaper = value; OnPropertyChanged("PrintPaper"); }
        }
        private string mailAddress = null;

        public string MailAddress
        {
            get { return mailAddress; }
            set { mailAddress = value; OnPropertyChanged("MailAddress"); }
        }
        private string toMailAddress = null;

        public string ToMailAddress
        {
            get { return toMailAddress; }
            set { toMailAddress = value; OnPropertyChanged("ToMailAddress"); }
        }
        private string mailSubject = null;

        public string MailSubject
        {
            get { return mailSubject; }
            set { mailSubject = value; OnPropertyChanged("MailSubject"); }
        }

        private string mailBody = null;

        public string MailBody
        {
            get { return mailBody; }
            set { mailBody = value; OnPropertyChanged("MailBody"); }
        }

        private string mailHost = null;

        public string MailHost
        {
            get { return mailHost; }
            set { mailHost = value; OnPropertyChanged("MailHost"); }
        }
        private int mailPort = 25;

        public int MailPort
        {
            get { return mailPort; }
            set { mailPort = value; OnPropertyChanged("MailPort"); }
        }
        private string mailUsername = null;

        public string MailUsername
        {
            get { return mailUsername; }
            set { mailUsername = value; OnPropertyChanged("MailUsername"); }
        }
        private string mailPwd = null;

        public string MailPwd
        {
            get { return mailPwd; }
            set { mailPwd = value; OnPropertyChanged("MailPwd"); }
        }

        private bool mailSsl;

        public bool MailSsl
        {
            get { return mailSsl; }
            set { mailSsl = value; OnPropertyChanged("MailSsl"); }
        }
    }
}

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
        private ObservableCollection<User> doctorNames = new ObservableCollection<User>();

        public ObservableCollection<User> DoctorNames
        {
            get { return doctorNames; }
            set { doctorNames = value; OnPropertyChanged("DoctorNames"); }
        }
        private ObservableCollection<User> techNames = new ObservableCollection<User>();

        public ObservableCollection<User> TechNames
        {
            get { return techNames; }
            set { techNames = value; OnPropertyChanged("TechNames"); }
        }

        private string meikBase = @"C:\Program Files (x86)\MEIK 5.6";

        public string MeikBase
        {
            get { return meikBase; }
            set { meikBase = value; OnPropertyChanged("MeikBase"); }
        }

        private bool useDefaultSignature;

        public bool UseDefaultSignature
        {
            get { return useDefaultSignature; }
            set { useDefaultSignature = value; OnPropertyChanged("UseDefaultSignature"); }
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

        private string deviceNo = "000";

        public string DeviceNo
        {
            get { return deviceNo; }
            set { deviceNo = value; OnPropertyChanged("DeviceNo"); }
        }

        /// <summary>
        /// 设备使用类型：1 操作员使用设备， 2 医生使用设备
        /// </summary>
        private int deviceType=1;

        public int DeviceType
        {
            get { return deviceType; }
            set { deviceType = value; OnPropertyChanged("DeviceType"); }
        }
    }
}

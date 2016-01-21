using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Model
{
    public class Equipment : ViewModelBase
    {
        private int id;
        private string sn;
        private string location;
        private string address;

        #region Public Members
        public int Id { 
            get { return this.id; }
            set {
                this.id = value;
                OnPropertyChanged("Id");
            }
        }
        public string SN
        {
            get { return this.sn; }
            set
            {
                this.sn = value;
                OnPropertyChanged("SN");
            }
        }
        public string Location
        {
            get { return this.location; }
            set
            {
                this.location = value;
                OnPropertyChanged("Location");
            }
        }
        public string Address
        {
            get { return this.address; }
            set
            {
                this.address = value;
                OnPropertyChanged("Address");
            }
        }
        #endregion
    }
}

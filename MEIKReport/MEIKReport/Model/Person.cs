
namespace MEIKReport.Model
{
    public class Person : ViewModelBase
    {
        private int id;
        private int equipmentId;
        private string code;
        private string surName;
        private string gender;
        private int age;
        private string address;
        private string birthday;
        private string regdate;


        #region Public Members
        public int Id { 
            get { return this.id; }
            set {
                this.id = value;
                OnPropertyChanged("Id");
            }
        }
        public int EquipmentId
        {
            get { return this.equipmentId; }
            set
            {
                this.equipmentId = value;
                OnPropertyChanged("EquipmentId");
            }
        }
        public string Code
        {
            get { return this.code; }
            set
            {
                this.code = value;
                OnPropertyChanged("Code");
            }
        }
        public string SurName
        {
            get { return this.surName; }
            set
            {
                this.surName = value;
                OnPropertyChanged("SurName");
            }
        }
        public string Gender
        {
            get { return this.gender; }
            set
            {
                this.gender = value;
                OnPropertyChanged("Gender");
            }
        }
        public int Age
        {
            get { return this.age; }
            set
            {
                this.age = value;
                OnPropertyChanged("Age");
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
        public string Birthday
        {
            get { return this.birthday; }
            set
            {
                this.birthday = value;
                OnPropertyChanged("Birthday");
            }
        }
        public string Regdate
        {
            get { return this.regdate; }
            set
            {
                this.regdate = value;
                OnPropertyChanged("Regdate");
            }
        }
        #endregion
    }
}

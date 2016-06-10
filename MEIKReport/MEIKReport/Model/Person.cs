
namespace MEIKReport.Model
{
    public class Person : ViewModelBase
    {
        private string archiveFolder;
        private string crdFilePath;
        private int id;        
        private string code;
        private string surName;
        private string givenName;
        private string otherName;

        private string gender;
        private int age;
        private string address;
        private string birthday;
        private string birthDate;
        private string birthMonth;
        private string birthYear;


        private string regDate;
        private string regMonth;
        private string regYear;
        private string remark;

        private bool pain;

        private string icon = "/Images/id_card.png";

        public bool Pain
        {
            get { return pain; }
            set { pain = value; OnPropertyChanged("Pain"); }
        }
        private bool colostrum;

        public bool Colostrum
        {
            get { return colostrum; }
            set { colostrum = value; OnPropertyChanged("Colostrum"); }
        }
        private bool serousDischarge;

        public bool SerousDischarge
        {
            get { return serousDischarge; }
            set { serousDischarge = value; OnPropertyChanged("SerousDischarge"); }
        }
        private bool bloodDischarge;

        public bool BloodDischarge
        {
            get { return bloodDischarge; }
            set { bloodDischarge = value; OnPropertyChanged("BloodDischarge"); }
        }
        private bool other;

        public bool Other
        {
            get { return other; }
            set { other = value; OnPropertyChanged("Other"); }
        }
        private bool pregnancy;

        public bool Pregnancy
        {
            get { return pregnancy; }
            set { pregnancy = value; OnPropertyChanged("Pregnancy"); }
        }
        private bool lactation;

        public bool Lactation
        {
            get { return lactation; }
            set { lactation = value; OnPropertyChanged("Lactation"); }
        }
        private string pregnancyTerm;

        public string PregnancyTerm
        {
            get { return pregnancyTerm; }
            set { pregnancyTerm = value; OnPropertyChanged("PregnancyTerm"); }
        }

        private string otherDesc;

        public string OtherDesc
        {
            get { return otherDesc; }
            set { otherDesc = value; OnPropertyChanged("OtherDesc"); }
        }
        
        public string ArchiveFolder
        {
            get { return archiveFolder; }
            set { archiveFolder = value; }
        }

        #region Public Members
        public int Id { 
            get { return this.id; }
            set {
                this.id = value;
                OnPropertyChanged("Id");
            }
        }

        public string Icon
        {
            get { return this.icon; }
            set
            {
                this.icon = value;
                OnPropertyChanged("Icon");
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

        public string GivenName
        {
            get { return this.givenName; }
            set
            {
                this.givenName = value;
                OnPropertyChanged("GivenName");
            }
        }
        public string OtherName
        {
            get { return this.otherName; }
            set
            {
                this.otherName = value;
                OnPropertyChanged("OtherName");
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

        public string BirthDate
        {
            get { return this.birthDate; }
            set
            {
                this.birthDate = value;
                OnPropertyChanged("BirthDate");
            }
        }

        public string BirthMonth
        {
            get { return this.birthMonth; }
            set
            {
                this.birthMonth = value;
                OnPropertyChanged("BirthMonth");
            }
        }

        public string BirthYear
        {
            get { return this.birthYear; }
            set
            {
                this.birthYear = value;
                OnPropertyChanged("BirthYear");
            }
        }
        public string RegDate
        {
            get { return this.regDate; }
            set
            {
                this.regDate = value;
                OnPropertyChanged("RegDate");
            }
        }
        public string RegMonth
        {
            get { return this.regMonth; }
            set
            {
                this.regMonth = value;
                OnPropertyChanged("RegMonth");
            }
        }
        public string RegYear
        {
            get { return this.regYear; }
            set
            {
                this.regYear = value;
                OnPropertyChanged("RegYear");
            }
        }

        public string Remark
        {
            get { return this.remark; }
            set
            {
                this.remark = value;
                OnPropertyChanged("Remark");
            }
        }

        private bool menstrualCycleDisorder;

        public bool MenstrualCycleDisorder
        {
            get { return menstrualCycleDisorder; }
            set { menstrualCycleDisorder = value; OnPropertyChanged("MenstrualCycleDisorder"); }
        }
        private bool postmenopause;

        public bool Postmenopause
        {
            get { return postmenopause; }
            set { postmenopause = value; OnPropertyChanged("Postmenopause"); }
        }
        private bool hormonalContraceptives;

        public bool HormonalContraceptives
        {
            get { return hormonalContraceptives; }
            set { hormonalContraceptives = value; OnPropertyChanged("HormonalContraceptives"); }
        }
        private string menstrualCycleDisorderDesc;

        public string MenstrualCycleDisorderDesc
        {
            get { return menstrualCycleDisorderDesc; }
            set { menstrualCycleDisorderDesc = value; OnPropertyChanged("MenstrualCycleDisorderDesc"); }
        }
        private string postmenopauseDesc;

        public string PostmenopauseDesc
        {
            get { return postmenopauseDesc; }
            set { postmenopauseDesc = value; OnPropertyChanged("PostmenopauseDesc"); }
        }
        private string hormonalContraceptivesBrandName;

        public string HormonalContraceptivesBrandName
        {
            get { return hormonalContraceptivesBrandName; }
            set { hormonalContraceptivesBrandName = value; OnPropertyChanged("HormonalContraceptivesBrandName"); }
        }
        private string hormonalContraceptivesPeriod;

        public string HormonalContraceptivesPeriod
        {
            get { return hormonalContraceptivesPeriod; }
            set { hormonalContraceptivesPeriod = value; OnPropertyChanged("HormonalContraceptivesPeriod"); }
        }



        private bool adiposity;

        public bool Adiposity
        {
            get { return adiposity; }
            set { adiposity = value; OnPropertyChanged("Adiposity"); }
        }
        private bool essentialHypertension;

        public bool EssentialHypertension
        {
            get { return essentialHypertension; }
            set { essentialHypertension = value; OnPropertyChanged("EssentialHypertension"); }
        }
        private bool diabetes;

        public bool Diabetes
        {
            get { return diabetes; }
            set { diabetes = value; OnPropertyChanged("Diabetes"); }
        }
        private bool thyroidGlandDiseases;

        public bool ThyroidGlandDiseases
        {
            get { return thyroidGlandDiseases; }
            set { thyroidGlandDiseases = value; OnPropertyChanged("ThyroidGlandDiseases"); }
        }
        private bool somaticOther;

        public bool SomaticOther
        {
            get { return somaticOther; }
            set { somaticOther = value; OnPropertyChanged("SomaticOther"); }
        }
        private string essentialHypertensionDesc;

        public string EssentialHypertensionDesc
        {
            get { return essentialHypertensionDesc; }
            set { essentialHypertensionDesc = value; OnPropertyChanged("EssentialHypertensionDesc"); }
        }
        private string diabetesDesc;

        public string DiabetesDesc
        {
            get { return diabetesDesc; }
            set { diabetesDesc = value; OnPropertyChanged("DiabetesDesc"); }
        }
        private string thyroidGlandDiseasesDesc;

        public string ThyroidGlandDiseasesDesc
        {
            get { return thyroidGlandDiseasesDesc; }
            set { thyroidGlandDiseasesDesc = value; OnPropertyChanged("ThyroidGlandDiseasesDesc"); }
        }
        private string somaticOtherDesc;

        public string SomaticOtherDesc
        {
            get { return somaticOtherDesc; }
            set { somaticOtherDesc = value; OnPropertyChanged("SomaticOtherDesc"); }
        }


        private bool infertility;

        public bool Infertility
        {
            get { return infertility; }
            set { infertility = value; OnPropertyChanged("Infertility"); }
        }
        private bool ovaryDiseases;

        public bool OvaryDiseases
        {
            get { return ovaryDiseases; }
            set { ovaryDiseases = value; OnPropertyChanged("OvaryDiseases"); }
        }
        private bool ovaryCyst;

        public bool OvaryCyst
        {
            get { return ovaryCyst; }
            set { ovaryCyst = value; OnPropertyChanged("OvaryCyst"); }
        }
        private bool ovaryCancer;

        public bool OvaryCancer
        {
            get { return ovaryCancer; }
            set { ovaryCancer = value; OnPropertyChanged("OvaryCancer"); }
        }
        private bool ovaryEndometriosis;

        public bool OvaryEndometriosis
        {
            get { return ovaryEndometriosis; }
            set { ovaryEndometriosis = value; OnPropertyChanged("OvaryEndometriosis"); }
        }
        private bool ovaryOther;

        public bool OvaryOther
        {
            get { return ovaryOther; }
            set { ovaryOther = value; OnPropertyChanged("OvaryOther"); }
        }
        private bool uterusDiseases;

        public bool UterusDiseases
        {
            get { return uterusDiseases; }
            set { uterusDiseases = value; OnPropertyChanged("UterusDiseases"); }
        }
        private bool uterusMyoma;

        public bool UterusMyoma
        {
            get { return uterusMyoma; }
            set { uterusMyoma = value; OnPropertyChanged("UterusMyoma"); }
        }
        private bool uterusCancer;

        public bool UterusCancer
        {
            get { return uterusCancer; }
            set { uterusCancer = value; OnPropertyChanged("UterusCancer"); }
        }
        private bool uterusEndometriosis;

        public bool UterusEndometriosis
        {
            get { return uterusEndometriosis; }
            set { uterusEndometriosis = value; OnPropertyChanged("UterusEndometriosis"); }
        }
        private bool uterusOther;

        public bool UterusOther
        {
            get { return uterusOther; }
            set { uterusOther = value; OnPropertyChanged("UterusOther"); }
        }
        private bool gynecologicOther;

        public bool GynecologicOther
        {
            get { return gynecologicOther; }
            set { gynecologicOther = value; OnPropertyChanged("GynecologicOther"); }
        }

        private string infertilityDesc;

        public string InfertilityDesc
        {
            get { return infertilityDesc; }
            set { infertilityDesc = value; OnPropertyChanged("InfertilityDesc"); }
        }
        private string ovaryOtherDesc;

        public string OvaryOtherDesc
        {
            get { return ovaryOtherDesc; }
            set { ovaryOtherDesc = value; OnPropertyChanged("OvaryOtherDesc"); }
        }
        private string uterusOtherDesc;

        public string UterusOtherDesc
        {
            get { return uterusOtherDesc; }
            set { uterusOtherDesc = value; OnPropertyChanged("UterusOtherDesc"); }
        }
        private string gynecologicOtherDesc;

        public string GynecologicOtherDesc
        {
            get { return gynecologicOtherDesc; }
            set { gynecologicOtherDesc = value; OnPropertyChanged("GynecologicOtherDesc"); }
        }


        private bool abortions;

        public bool Abortions
        {
            get { return abortions; }
            set { abortions = value; OnPropertyChanged("Abortions"); }
        }
        private bool births;

        public bool Births
        {
            get { return births; }
            set { births = value; OnPropertyChanged("Births"); }
        }
        private string abortionsNumber;

        public string AbortionsNumber
        {
            get { return abortionsNumber; }
            set { abortionsNumber = value; OnPropertyChanged("AbortionsNumber"); }
        }
        private string birthsNumber;

        public string BirthsNumber
        {
            get { return birthsNumber; }
            set { birthsNumber = value; OnPropertyChanged("BirthsNumber"); }
        }


        private bool lactationTill1Month;

        public bool LactationTill1Month
        {
            get { return this.lactationTill1Month; }
            set { this.lactationTill1Month = value; OnPropertyChanged("LactationTill1Month"); }
        }
        private bool lactationTill1Year;

        public bool LactationTill1Year
        {
            get { return this.lactationTill1Year; }
            set { this.lactationTill1Year = value; OnPropertyChanged("LactationTill1Year"); }
        }
        private bool lactationOver1Year;

        public bool LactationOver1Year
        {
            get { return this.lactationOver1Year; }
            set { this.lactationOver1Year = value; OnPropertyChanged("LactationOver1Year"); }
        }



        private bool trauma;

        public bool Trauma
        {
            get { return trauma; }
            set { trauma = value; OnPropertyChanged("Trauma"); }
        }
        private bool mastitis;

        public bool Mastitis
        {
            get { return mastitis; }
            set { mastitis = value; OnPropertyChanged("Mastitis"); }
        }
        private bool fibrousCysticMastopathy;

        public bool FibrousCysticMastopathy
        {
            get { return fibrousCysticMastopathy; }
            set { fibrousCysticMastopathy = value; OnPropertyChanged("FibrousCysticMastopathy"); }
        }
        private bool cysts;

        public bool Cysts
        {
            get { return cysts; }
            set { cysts = value; OnPropertyChanged("Cysts"); }
        }
        private bool cancer;

        public bool Cancer
        {
            get { return cancer; }
            set { cancer = value; OnPropertyChanged("Cancer"); }
        }
        private bool diseasesOther;

        public bool DiseasesOther
        {
            get { return diseasesOther; }
            set { diseasesOther = value; OnPropertyChanged("DiseasesOther"); }
        }
        private string traumaDesc;

        public string TraumaDesc
        {
            get { return traumaDesc; }
            set { traumaDesc = value; OnPropertyChanged("TraumaDesc"); }
        }
        private string mastitisDesc;

        public string MastitisDesc
        {
            get { return mastitisDesc; }
            set { mastitisDesc = value; OnPropertyChanged("MastitisDesc"); }
        }
        private string fibrousCysticMastopathyDesc;

        public string FibrousCysticMastopathyDesc
        {
            get { return fibrousCysticMastopathyDesc; }
            set { fibrousCysticMastopathyDesc = value; OnPropertyChanged("FibrousCysticMastopathyDesc"); }
        }
        private string cystsDesc;

        public string CystsDesc
        {
            get { return cystsDesc; }
            set { cystsDesc = value; OnPropertyChanged("CystsDesc"); }
        }
        private string cancerDesc;

        public string CancerDesc
        {
            get { return cancerDesc; }
            set { cancerDesc = value; OnPropertyChanged("CancerDesc"); }
        }
        private string diseasesOtherDesc;

        public string DiseasesOtherDesc
        {
            get { return diseasesOtherDesc; }
            set { diseasesOtherDesc = value; OnPropertyChanged("DiseasesOtherDesc"); }
        }



        private bool palpationDiffuse;

        public bool PalpationDiffuse
        {
            get { return palpationDiffuse; }
            set { palpationDiffuse = value; }
        }
        private bool palpationFocal;

        public bool PalpationFocal
        {
            get { return palpationFocal; }
            set { palpationFocal = value; }
        }
        private bool ultrasoundDiffuse;

        public bool UltrasoundDiffuse
        {
            get { return ultrasoundDiffuse; }
            set { ultrasoundDiffuse = value; }
        }
        private bool ultrasoundFocal;

        public bool UltrasoundFocal
        {
            get { return ultrasoundFocal; }
            set { ultrasoundFocal = value; }
        }
        private bool mammographyDiffuse;

        public bool MammographyDiffuse
        {
            get { return mammographyDiffuse; }
            set { mammographyDiffuse = value; }
        }
        private bool mammographyFocal;

        public bool MammographyFocal
        {
            get { return mammographyFocal; }
            set { mammographyFocal = value; }
        }
        private bool biopsyDiffuse;

        public bool BiopsyDiffuse
        {
            get { return biopsyDiffuse; }
            set { biopsyDiffuse = value; }
        }
        private bool biopsyFocal;

        public bool BiopsyFocal
        {
            get { return biopsyFocal; }
            set { biopsyFocal = value; }
        }
        private bool biopsyCancer;
        
        public bool BiopsyCancer
        {
            get { return biopsyCancer; }
            set { biopsyCancer = value; }
        }
        private bool biopsyProliferation;

        public bool BiopsyProliferation
        {
            get { return biopsyProliferation; }
            set { biopsyProliferation = value; }
        }
        private bool biopsyDysplasia;

        public bool BiopsyDysplasia
        {
            get { return biopsyDysplasia; }
            set { biopsyDysplasia = value; }
        }
        private bool biopsyIntraductalPapilloma;

        public bool BiopsyIntraductalPapilloma
        {
            get { return biopsyIntraductalPapilloma; }
            set { biopsyIntraductalPapilloma = value; }
        }
        private bool biopsyFibroadenoma;

        public bool BiopsyFibroadenoma
        {
            get { return biopsyFibroadenoma; }
            set { biopsyFibroadenoma = value; }
        }
        private bool biopsyOther;

        public bool BiopsyOther
        {
            get { return biopsyOther; }
            set { biopsyOther = value; }
        }
        private string palpationDesc;

        public string PalpationDesc
        {
            get { return palpationDesc; }
            set { palpationDesc = value; }
        }
        private string ultrasoundnDesc;

        public string UltrasoundnDesc
        {
            get { return ultrasoundnDesc; }
            set { ultrasoundnDesc = value; }
        }
        private string mammographyDesc;

        public string MammographyDesc
        {
            get { return mammographyDesc; }
            set { mammographyDesc = value; }
        }
        private string biopsyOtherDesc;

        public string BiopsyOtherDesc
        {
            get { return biopsyOtherDesc; }
            set { biopsyOtherDesc = value; }
        }

        private string techName;
        public string TechName
        {
            get { return techName; }
            set { techName = value; }
        }
        private string techLicense;
        public string TechLicense
        {
            get { return techLicense; }
            set { techLicense = value; }
        }

        private string doctorName;
        public string DoctorName
        {
            get { return doctorName; }
            set { techName = value; }
        }
        private string doctorLicense;
        public string DoctorLicense
        {
            get { return doctorLicense; }
            set { doctorLicense = value; }
        }

        public string CrdFilePath
        {
            get { return crdFilePath; }
            set { crdFilePath = value; }
        }

        #endregion
    }
}

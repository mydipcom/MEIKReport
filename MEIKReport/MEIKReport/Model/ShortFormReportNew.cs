using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Model
{
    [Serializable]
    public class ShortFormReportNew : ViewModelBase,ICloneable
    {
        private string dataScreenDate;

        public string DataScreenDate
        {
            get { return dataScreenDate; }
            set { dataScreenDate = value; OnPropertyChanged("DataScreenDate"); }
        }
        private string dataScreenLocation;

        public string DataScreenLocation
        {
            get { return dataScreenLocation; }
            set { dataScreenLocation = value; OnPropertyChanged("DataScreenLocation"); }
        }
        private string dataUserCode;

        public string DataUserCode
        {
            get { return dataUserCode; }
            set { dataUserCode = value; OnPropertyChanged("DataUserCode"); }
        }
        private string dataName;

        public string DataName
        {
            get { return dataName; }
            set { dataName = value; OnPropertyChanged("DataName"); }
        }
        private string dataAge;

        public string DataAge
        {
            get { return dataAge; }
            set { dataAge = value; OnPropertyChanged("DataAge"); }
        }
        private string dataLeftFinding;

        public string DataLeftFinding
        {
            get { return dataLeftFinding; }
            set { dataLeftFinding = value; OnPropertyChanged("DataLeftFinding"); }
        }
        private string dataRightFinding;

        public string DataRightFinding
        {
            get { return dataRightFinding; }
            set { dataRightFinding = value; OnPropertyChanged("DataRightFinding"); }
        }
        private string dataLeftLocation;

        public string DataLeftLocation
        {
            get { return dataLeftLocation; }
            set { dataLeftLocation = value; OnPropertyChanged("DataLeftLocation"); }
        }
        private string dataRightLocation;

        public string DataRightLocation
        {
            get { return dataRightLocation; }
            set { dataRightLocation = value; OnPropertyChanged("DataRightLocation"); }
        }

        private string dataLeftSize;

        public string DataLeftSize
        {
            get { return dataLeftSize; }
            set { dataLeftSize = value; OnPropertyChanged("DataLeftSize"); }
        }
        private string dataRightSize;

        public string DataRightSize
        {
            get { return dataRightSize; }
            set { dataRightSize = value; OnPropertyChanged("DataRightSize"); }
        }
        private string dataBiRadsCategory;

        public string DataBiRadsCategory
        {
            get { return dataBiRadsCategory; }
            set { dataBiRadsCategory = value; OnPropertyChanged("DataBiRadsCategory"); }
        }
        private string dataRecommendation;

        public string DataRecommendation
        {
            get { return dataRecommendation; }
            set { dataRecommendation = value; OnPropertyChanged("DataRecommendation"); }
        }
        private string dataFurtherExam;

        public string DataFurtherExam
        {
            get { return dataFurtherExam; }
            set { dataFurtherExam = value; OnPropertyChanged("DataFurtherExam"); }
        }
        private string dataConclusion;

        public string DataConclusion
        {
            get { return dataConclusion; }
            set { dataConclusion = value; OnPropertyChanged("DataConclusion"); }
        }
        private string dataConclusion2;

        public string DataConclusion2
        {
            get { return dataConclusion2; }
            set { dataConclusion2 = value; OnPropertyChanged("DataConclusion2"); }
        }
        private string dataComments;

        public string DataComments
        {
            get { return dataComments; }
            set { dataComments = value; OnPropertyChanged("DataComments"); }
        }
        private string dataSignDate;

        public string DataSignDate
        {
            get { return dataSignDate; }
            set { dataSignDate = value; OnPropertyChanged("DataSignDate"); }
        }

        private byte[] dataSignImg;

        public byte[] DataSignImg
        {
            get { return dataSignImg; }
            set { dataSignImg = value; OnPropertyChanged("DataSignImg"); }
        }
        private string dataMeikTech;

        public string DataMeikTech
        {
            get { return dataMeikTech; }
            set { dataMeikTech = value; OnPropertyChanged("DataMeikTech"); }
        }
        private string dataTechLicense;

        public string DataTechLicense
        {
            get { return dataTechLicense; }
            set { dataTechLicense = value; OnPropertyChanged("DataTechLicense"); }
        }
        private string dataDoctor;

        public string DataDoctor
        {
            get { return dataDoctor; }
            set { dataDoctor = value; OnPropertyChanged("DataDoctor"); }
        }
        private string dataDoctorLicense;

        public string DataDoctorLicense
        {
            get { return dataDoctorLicense; }
            set { dataDoctorLicense = value; OnPropertyChanged("DataDoctorLicense"); }
        }
        //private string dataCategories;

        //public string DataCategories
        //{
        //    get { return dataCategories; }
        //    set { dataCategories = value; }
        //}
        private byte[] dataScreenShotImg;

        public byte[] DataScreenShotImg
        {
            get { return dataScreenShotImg; }
            set { dataScreenShotImg = value; OnPropertyChanged("DataScreenShotImg"); }
        }




        private string dataGender;

        public string DataGender
        {
            get { return dataGender; }
            set { dataGender = value; OnPropertyChanged("DataGender"); }
        }

        private string dataAddress;

        public string DataAddress
        {
            get { return dataAddress; }
            set { dataAddress = value; OnPropertyChanged("DataAddress"); }
        }

        private string dataHealthCard;

        public string DataHealthCard
        {
            get { return dataHealthCard; }
            set { dataHealthCard = value; OnPropertyChanged("DataHealthCard"); }
        }

        private string dataWeight;

        public string DataWeight
        {
            get { return dataWeight; }
            set { dataWeight = value; OnPropertyChanged("DataWeight"); }
        }

        private string dataWeightUnit;

        public string DataWeightUnit
        {
            get { return dataWeightUnit; }
            set { dataWeightUnit = value; OnPropertyChanged("DataWeightUnit"); }
        }

        private string dataMenstrualCycle;

        public string DataMenstrualCycle
        {
            get { return dataMenstrualCycle; }
            set { dataMenstrualCycle = value; OnPropertyChanged("DataMenstrualCycle"); }
        }

        private string dataHormones;

        public string DataHormones
        {
            get { return dataHormones; }
            set { dataHormones = value; OnPropertyChanged("DataHormones"); }
        }

        private string dataSkinAffections;

        public string DataSkinAffections
        {
            get { return dataSkinAffections; }
            set { dataSkinAffections = value; OnPropertyChanged("DataSkinAffections"); }
        }

        private string dataPertinentHistory;

        public string DataPertinentHistory
        {
            get { return dataPertinentHistory; }
            set { dataPertinentHistory = value; OnPropertyChanged("DataPertinentHistory"); }
        }

        private string dataPertinentHistory1;

        public string DataPertinentHistory1
        {
            get { return dataPertinentHistory1; }
            set { dataPertinentHistory1 = value; OnPropertyChanged("DataPertinentHistory1"); }
        }

        private string dataMotherUltra;

        public string DataMotherUltra
        {
            get { return dataMotherUltra; }
            set { dataMotherUltra = value; OnPropertyChanged("DataMotherUltra"); }
        }

        private string dataLeftBreast;

        public string DataLeftBreast
        {
            get { return dataLeftBreast; }
            set { dataLeftBreast = value; OnPropertyChanged("DataLeftBreast"); }
        }

        private string dataRightBreast;

        public string DataRightBreast
        {
            get { return dataRightBreast; }
            set { dataRightBreast = value; OnPropertyChanged("DataRightBreast"); }
        }

        private string dataLeftPalpableMass;

        public string DataLeftPalpableMass
        {
            get { return dataLeftPalpableMass; }
            set { dataLeftPalpableMass = value; OnPropertyChanged("DataLeftPalpableMass"); }
        }

        private string dataRightPalpableMass;

        public string DataRightPalpableMass
        {
            get { return dataRightPalpableMass; }
            set { dataRightPalpableMass = value; OnPropertyChanged("DataRightPalpableMass"); }
        }

        private string dataLeftChangesOfElectricalConductivity;

        public string DataLeftChangesOfElectricalConductivity
        {
            get { return dataLeftChangesOfElectricalConductivity; }
            set { dataLeftChangesOfElectricalConductivity = value; OnPropertyChanged("DataLeftChangesOfElectricalConductivity"); }
        }

        private string dataRightChangesOfElectricalConductivity;

        public string DataRightChangesOfElectricalConductivity
        {
            get { return dataRightChangesOfElectricalConductivity; }
            set { dataRightChangesOfElectricalConductivity = value; OnPropertyChanged("DataRightChangesOfElectricalConductivity"); }
        }

        private string dataLeftMammaryStruct;

        public string DataLeftMammaryStruct
        {
            get { return dataLeftMammaryStruct; }
            set { dataLeftMammaryStruct = value; OnPropertyChanged("DataLeftMammaryStruct"); }
        }

        private string dataRightMammaryStruct;

        public string DataRightMammaryStruct
        {
            get { return dataRightMammaryStruct; }
            set { dataRightMammaryStruct = value; OnPropertyChanged("DataRightMammaryStruct"); }
        }

        private string dataLeftLactiferousSinusZone;

        public string DataLeftLactiferousSinusZone
        {
            get { return dataLeftLactiferousSinusZone; }
            set { dataLeftLactiferousSinusZone = value; OnPropertyChanged("DataLeftLactiferousSinusZone"); }
        }

        private string dataRightLactiferousSinusZone;

        public string DataRightLactiferousSinusZone
        {
            get { return dataRightLactiferousSinusZone; }
            set { dataRightLactiferousSinusZone = value; OnPropertyChanged("DataRightLactiferousSinusZone"); }
        }

        private string dataLeftMammaryContour;

        public string DataLeftMammaryContour
        {
            get { return dataLeftMammaryContour; }
            set { dataLeftMammaryContour = value; OnPropertyChanged("DataLeftMammaryContour"); }
        }

        private string dataRightMammaryContour;

        public string DataRightMammaryContour
        {
            get { return dataRightMammaryContour; }
            set { dataRightMammaryContour = value; OnPropertyChanged("DataRightMammaryContour"); }
        }

        private string dataLeftNumber;

        public string DataLeftNumber
        {
            get { return dataLeftNumber; }
            set { dataLeftNumber = value; OnPropertyChanged("DataLeftNumber"); }
        }

        private string dataRightNumber;

        public string DataRightNumber
        {
            get { return dataRightNumber; }
            set { dataRightNumber = value; OnPropertyChanged("DataRightNumber"); }
        }

        private string dataLeftShape;

        public string DataLeftShape
        {
            get { return dataLeftShape; }
            set { dataLeftShape = value; OnPropertyChanged("DataLeftShape"); }
        }

        private string dataRightShape;

        public string DataRightShape
        {
            get { return dataRightShape; }
            set { dataRightShape = value; OnPropertyChanged("DataRightShape"); }
        }

        private string dataLeftContourAroundFocus;

        public string DataLeftContourAroundFocus
        {
            get { return dataLeftContourAroundFocus; }
            set { dataLeftContourAroundFocus = value; OnPropertyChanged("DataLeftContourAroundFocus"); }
        }

        private string dataRightContourAroundFocus;

        public string DataRightContourAroundFocus
        {
            get { return dataRightContourAroundFocus; }
            set { dataRightContourAroundFocus = value; OnPropertyChanged("DataRightContourAroundFocus"); }
        }

        private string dataLeftInternalElectricalStructure;

        public string DataLeftInternalElectricalStructure
        {
            get { return dataLeftInternalElectricalStructure; }
            set { dataLeftInternalElectricalStructure = value; OnPropertyChanged("DataLeftInternalElectricalStructure"); }
        }

        private string dataRightInternalElectricalStructure;

        public string DataRightInternalElectricalStructure
        {
            get { return dataRightInternalElectricalStructure; }
            set { dataRightInternalElectricalStructure = value; OnPropertyChanged("DataRightInternalElectricalStructure"); }
        }

        private string dataLeftSurroundingTissues;

        public string DataLeftSurroundingTissues
        {
            get { return dataLeftSurroundingTissues; }
            set { dataLeftSurroundingTissues = value; OnPropertyChanged("DataLeftSurroundingTissues"); }
        }

        private string dataRightSurroundingTissues;

        public string DataRightSurroundingTissues
        {
            get { return dataRightSurroundingTissues; }
            set { dataRightSurroundingTissues = value; OnPropertyChanged("DataRightSurroundingTissues"); }
        }

        private string dataLeftOncomarkerHighlightBenignChanges;

        public string DataLeftOncomarkerHighlightBenignChanges
        {
            get { return dataLeftOncomarkerHighlightBenignChanges; }
            set { dataLeftOncomarkerHighlightBenignChanges = value; OnPropertyChanged("DataLeftOncomarkerHighlightBenignChanges"); }
        }

        private string dataRightOncomarkerHighlightBenignChanges;

        public string DataRightOncomarkerHighlightBenignChanges
        {
            get { return dataRightOncomarkerHighlightBenignChanges; }
            set { dataRightOncomarkerHighlightBenignChanges = value; OnPropertyChanged("DataRightOncomarkerHighlightBenignChanges"); }
        }

        private string dataLeftOncomarkerHighlightSuspiciousChanges;

        public string DataLeftOncomarkerHighlightSuspiciousChanges
        {
            get { return dataLeftOncomarkerHighlightSuspiciousChanges; }
            set { dataLeftOncomarkerHighlightSuspiciousChanges = value; OnPropertyChanged("DataLeftOncomarkerHighlightSuspiciousChanges"); }
        }

        private string dataRightOncomarkerHighlightSuspiciousChanges;

        public string DataRightOncomarkerHighlightSuspiciousChanges
        {
            get { return dataRightOncomarkerHighlightSuspiciousChanges; }
            set { dataRightOncomarkerHighlightSuspiciousChanges = value; OnPropertyChanged("DataRightOncomarkerHighlightSuspiciousChanges"); }
        }

        private string dataLeftMeanElectricalConductivity1;

        public string DataLeftMeanElectricalConductivity1
        {
            get { return dataLeftMeanElectricalConductivity1; }
            set { dataLeftMeanElectricalConductivity1 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity1"); }
        }

        private string dataLeftMeanElectricalConductivity1N1;

        public string DataLeftMeanElectricalConductivity1N1
        {
            get { return dataLeftMeanElectricalConductivity1N1; }
            set { dataLeftMeanElectricalConductivity1N1 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity1N1"); }
        }

        private string dataLeftMeanElectricalConductivity1N2;

        public string DataLeftMeanElectricalConductivity1N2
        {
            get { return dataLeftMeanElectricalConductivity1N2; }
            set { dataLeftMeanElectricalConductivity1N2 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity1N2"); }
        }

        private string dataRightMeanElectricalConductivity1;

        public string DataRightMeanElectricalConductivity1
        {
            get { return dataRightMeanElectricalConductivity1; }
            set { dataRightMeanElectricalConductivity1 = value; OnPropertyChanged("DataRightMeanElectricalConductivity1"); }
        }

        private string dataRightMeanElectricalConductivity1N1;

        public string DataRightMeanElectricalConductivity1N1
        {
            get { return dataRightMeanElectricalConductivity1N1; }
            set { dataRightMeanElectricalConductivity1N1 = value; OnPropertyChanged("DataRightMeanElectricalConductivity1N1"); }
        }

        private string dataRightMeanElectricalConductivity1N2;

        public string DataRightMeanElectricalConductivity1N2
        {
            get { return dataRightMeanElectricalConductivity1N2; }
            set { dataRightMeanElectricalConductivity1N2 = value; OnPropertyChanged("DataRightMeanElectricalConductivity1N2"); }
        }

        private string dataLeftMeanElectricalConductivity2;

        public string DataLeftMeanElectricalConductivity2
        {
            get { return dataLeftMeanElectricalConductivity2; }
            set { dataLeftMeanElectricalConductivity2 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity2"); }
        }

        private string dataLeftMeanElectricalConductivity2N1;

        public string DataLeftMeanElectricalConductivity2N1
        {
            get { return dataLeftMeanElectricalConductivity2N1; }
            set { dataLeftMeanElectricalConductivity2N1 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity2N1"); }
        }

        private string dataLeftMeanElectricalConductivity2N2;

        public string DataLeftMeanElectricalConductivity2N2
        {
            get { return dataLeftMeanElectricalConductivity2N2; }
            set { dataLeftMeanElectricalConductivity2N2 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity2N2"); }
        }

        private string dataRightMeanElectricalConductivity2;

        public string DataRightMeanElectricalConductivity2
        {
            get { return dataRightMeanElectricalConductivity2; }
            set { dataRightMeanElectricalConductivity2 = value; OnPropertyChanged("DataRightMeanElectricalConductivity2"); }
        }

        private string dataRightMeanElectricalConductivity2N1;

        public string DataRightMeanElectricalConductivity2N1
        {
            get { return dataRightMeanElectricalConductivity2N1; }
            set { dataRightMeanElectricalConductivity2N1 = value; OnPropertyChanged("DataRightMeanElectricalConductivity2N1"); }
        }

        private string dataRightMeanElectricalConductivity2N2;

        public string DataRightMeanElectricalConductivity2N2
        {
            get { return dataRightMeanElectricalConductivity2N2; }
            set { dataRightMeanElectricalConductivity2N2 = value; OnPropertyChanged("DataRightMeanElectricalConductivity2N2"); }
        }

        private string dataMeanElectricalConductivity3;

        public string DataMeanElectricalConductivity3
        {
            get { return dataMeanElectricalConductivity3; }
            set { dataMeanElectricalConductivity3 = value; OnPropertyChanged("DataMeanElectricalConductivity3"); }
        }

        private string dataLeftMeanElectricalConductivity3;

        public string DataLeftMeanElectricalConductivity3
        {
            get { return dataLeftMeanElectricalConductivity3; }
            set { dataLeftMeanElectricalConductivity3 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity3"); }
        }

        private string dataLeftMeanElectricalConductivity3N1;

        public string DataLeftMeanElectricalConductivity3N1
        {
            get { return dataLeftMeanElectricalConductivity3N1; }
            set { dataLeftMeanElectricalConductivity3N1 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity3N1"); }
        }

        private string dataLeftMeanElectricalConductivity3N2;

        public string DataLeftMeanElectricalConductivity3N2
        {
            get { return dataLeftMeanElectricalConductivity3N2; }
            set { dataLeftMeanElectricalConductivity3N2 = value; OnPropertyChanged("DataLeftMeanElectricalConductivity3N2"); }
        }

        private string dataRightMeanElectricalConductivity3;

        public string DataRightMeanElectricalConductivity3
        {
            get { return dataRightMeanElectricalConductivity3; }
            set { dataRightMeanElectricalConductivity3 = value; OnPropertyChanged("DataRightMeanElectricalConductivity3"); }
        }

        private string dataRightMeanElectricalConductivity3N1;

        public string DataRightMeanElectricalConductivity3N1
        {
            get { return dataRightMeanElectricalConductivity3N1; }
            set { dataRightMeanElectricalConductivity3N1 = value; OnPropertyChanged("DataRightMeanElectricalConductivity3N1"); }
        }

        private string dataRightMeanElectricalConductivity3N2;

        public string DataRightMeanElectricalConductivity3N2
        {
            get { return dataRightMeanElectricalConductivity3N2; }
            set { dataRightMeanElectricalConductivity3N2 = value; OnPropertyChanged("DataRightMeanElectricalConductivity3N2"); }
        }

        private string dataComparativeElectricalConductivityReference1;

        public string DataComparativeElectricalConductivityReference1
        {
            get { return dataComparativeElectricalConductivityReference1; }
            set { dataComparativeElectricalConductivityReference1 = value; OnPropertyChanged("DataComparativeElectricalConductivityReference1"); }
        }

        private string dataLeftComparativeElectricalConductivity1;

        public string DataLeftComparativeElectricalConductivity1
        {
            get { return dataLeftComparativeElectricalConductivity1; }
            set { dataLeftComparativeElectricalConductivity1 = value; OnPropertyChanged("DataLeftComparativeElectricalConductivity1"); }
        }

        private string dataRightComparativeElectricalConductivity1;

        public string DataRightComparativeElectricalConductivity1
        {
            get { return dataRightComparativeElectricalConductivity1; }
            set { dataRightComparativeElectricalConductivity1 = value; OnPropertyChanged("DataRightComparativeElectricalConductivity1"); }
        }

        private string dataComparativeElectricalConductivityReference2;

        public string DataComparativeElectricalConductivityReference2
        {
            get { return dataComparativeElectricalConductivityReference2; }
            set { dataComparativeElectricalConductivityReference2 = value; OnPropertyChanged("DataComparativeElectricalConductivityReference2"); }
        }

        private string dataLeftComparativeElectricalConductivity2;

        public string DataLeftComparativeElectricalConductivity2
        {
            get { return dataLeftComparativeElectricalConductivity2; }
            set { dataLeftComparativeElectricalConductivity2 = value; OnPropertyChanged("DataLeftComparativeElectricalConductivity2"); }
        }

        private string dataRightComparativeElectricalConductivity2;

        public string DataRightComparativeElectricalConductivity2
        {
            get { return dataRightComparativeElectricalConductivity2; }
            set { dataRightComparativeElectricalConductivity2 = value; OnPropertyChanged("DataRightComparativeElectricalConductivity2"); }
        }

        private string dataComparativeElectricalConductivity3;

        public string DataComparativeElectricalConductivity3
        {
            get { return dataComparativeElectricalConductivity3; }
            set { dataComparativeElectricalConductivity3 = value; OnPropertyChanged("DataComparativeElectricalConductivity3"); }
        }

        private string dataLeftComparativeElectricalConductivity3;

        public string DataLeftComparativeElectricalConductivity3
        {
            get { return dataLeftComparativeElectricalConductivity3; }
            set { dataLeftComparativeElectricalConductivity3 = value; OnPropertyChanged("DataLeftComparativeElectricalConductivity3"); }
        }

        private string dataRightComparativeElectricalConductivity3;

        public string DataRightComparativeElectricalConductivity3
        {
            get { return dataRightComparativeElectricalConductivity3; }
            set { dataRightComparativeElectricalConductivity3 = value; OnPropertyChanged("DataRightComparativeElectricalConductivity3"); }
        }

        private string dataDivergenceBetweenHistogramsReference1;

        public string DataDivergenceBetweenHistogramsReference1
        {
            get { return dataDivergenceBetweenHistogramsReference1; }
            set { dataDivergenceBetweenHistogramsReference1 = value; OnPropertyChanged("DataDivergenceBetweenHistogramsReference1"); }
        }

        private string dataLeftDivergenceBetweenHistograms1;

        public string DataLeftDivergenceBetweenHistograms1
        {
            get { return dataLeftDivergenceBetweenHistograms1; }
            set { dataLeftDivergenceBetweenHistograms1 = value; OnPropertyChanged("DataLeftDivergenceBetweenHistograms1"); }
        }

        private string dataRightDivergenceBetweenHistograms1;

        public string DataRightDivergenceBetweenHistograms1
        {
            get { return dataRightDivergenceBetweenHistograms1; }
            set { dataRightDivergenceBetweenHistograms1 = value; OnPropertyChanged("DataRightDivergenceBetweenHistograms1"); }
        }

        private string dataDivergenceBetweenHistogramsReference2;

        public string DataDivergenceBetweenHistogramsReference2
        {
            get { return dataDivergenceBetweenHistogramsReference2; }
            set { dataDivergenceBetweenHistogramsReference2 = value; OnPropertyChanged("DataDivergenceBetweenHistogramsReference2"); }
        }

        private string dataLeftDivergenceBetweenHistograms2;

        public string DataLeftDivergenceBetweenHistograms2
        {
            get { return dataLeftDivergenceBetweenHistograms2; }
            set { dataLeftDivergenceBetweenHistograms2 = value; OnPropertyChanged("DataLeftDivergenceBetweenHistograms2"); }
        }

        private string dataRightDivergenceBetweenHistograms2;

        public string DataRightDivergenceBetweenHistograms2
        {
            get { return dataRightDivergenceBetweenHistograms2; }
            set { dataRightDivergenceBetweenHistograms2 = value; OnPropertyChanged("DataRightDivergenceBetweenHistograms2"); }
        }

        private string dataDivergenceBetweenHistograms3;

        public string DataDivergenceBetweenHistograms3
        {
            get { return dataDivergenceBetweenHistograms3; }
            set { dataDivergenceBetweenHistograms3 = value; OnPropertyChanged("DataDivergenceBetweenHistograms3"); }
        }

        private string dataLeftDivergenceBetweenHistograms3;

        public string DataLeftDivergenceBetweenHistograms3
        {
            get { return dataLeftDivergenceBetweenHistograms3; }
            set { dataLeftDivergenceBetweenHistograms3 = value; OnPropertyChanged("DataLeftDivergenceBetweenHistograms3"); }
        }

        private string dataRightDivergenceBetweenHistograms3;

        public string DataRightDivergenceBetweenHistograms3
        {
            get { return dataRightDivergenceBetweenHistograms3; }
            set { dataRightDivergenceBetweenHistograms3 = value; OnPropertyChanged("DataRightDivergenceBetweenHistograms3"); }
        }

        private string dataLeftComparisonWithNorm;

        public string DataLeftComparisonWithNorm
        {
            get { return dataLeftComparisonWithNorm; }
            set { dataLeftComparisonWithNorm = value; OnPropertyChanged("DataLeftComparisonWithNorm"); }
        }

        private string dataRightComparisonWithNorm;

        public string DataRightComparisonWithNorm
        {
            get { return dataRightComparisonWithNorm; }
            set { dataRightComparisonWithNorm = value; OnPropertyChanged("DataRightComparisonWithNorm"); }
        }

        private string dataPhaseElectricalConductivityReference;

        public string DataPhaseElectricalConductivityReference
        {
            get { return dataPhaseElectricalConductivityReference; }
            set { dataPhaseElectricalConductivityReference = value; OnPropertyChanged("DataPhaseElectricalConductivityReference"); }
        }

        private string dataLeftPhaseElectricalConductivity;

        public string DataLeftPhaseElectricalConductivity
        {
            get { return dataLeftPhaseElectricalConductivity; }
            set { dataLeftPhaseElectricalConductivity = value; OnPropertyChanged("DataLeftPhaseElectricalConductivity"); }
        }

        private string dataRightPhaseElectricalConductivity;

        public string DataRightPhaseElectricalConductivity
        {
            get { return dataRightPhaseElectricalConductivity; }
            set { dataRightPhaseElectricalConductivity = value; OnPropertyChanged("DataRightPhaseElectricalConductivity"); }
        }

        private string dataAgeElectricalConductivityReference;

        public string DataAgeElectricalConductivityReference
        {
            get { return dataAgeElectricalConductivityReference; }
            set { dataAgeElectricalConductivityReference = value; OnPropertyChanged("DataAgeElectricalConductivityReference"); }
        }

        private string dataLeftAgeElectricalConductivity;

        public string DataLeftAgeElectricalConductivity
        {
            get { return dataLeftAgeElectricalConductivity; }
            set { dataLeftAgeElectricalConductivity = value; OnPropertyChanged("DataLeftAgeElectricalConductivity"); }
        }

        private string dataRightAgeElectricalConductivity;

        public string DataRightAgeElectricalConductivity
        {
            get { return dataRightAgeElectricalConductivity; }
            set { dataRightAgeElectricalConductivity = value; OnPropertyChanged("DataRightAgeElectricalConductivity"); }
        }

        private string dataAgeValueOfEC;

        public string DataAgeValueOfEC
        {
            get { return dataAgeValueOfEC; }
            set { dataAgeValueOfEC = value; OnPropertyChanged("DataAgeValueOfEC"); }
        }

        private string dataExamConclusion;

        public string DataExamConclusion
        {
            get { return dataExamConclusion; }
            set { dataExamConclusion = value; OnPropertyChanged("DataExamConclusion"); }
        }

        private string dataLeftMammaryGland;

        public string DataLeftMammaryGland
        {
            get { return dataLeftMammaryGland; }
            set { dataLeftMammaryGland = value; OnPropertyChanged("DataLeftMammaryGland"); }
        }

        private string dataLeftAgeRelated;

        public string DataLeftAgeRelated
        {
            get { return dataLeftAgeRelated; }
            set { dataLeftAgeRelated = value; OnPropertyChanged("DataLeftAgeRelated"); }
        }

        private string dataLeftMeanECOfLesion;

        public string DataLeftMeanECOfLesion
        {
            get { return dataLeftMeanECOfLesion; }
            set { dataLeftMeanECOfLesion = value; OnPropertyChanged("DataLeftMeanECOfLesion"); }
        }

        private string dataLeftFindings;

        public string DataLeftFindings
        {
            get { return dataLeftFindings; }
            set { dataLeftFindings = value; OnPropertyChanged("DataLeftFindings"); }
        }

        private string dataLeftMammaryGlandResult;

        public string DataLeftMammaryGlandResult
        {
            get { return dataLeftMammaryGlandResult; }
            set { dataLeftMammaryGlandResult = value; OnPropertyChanged("DataLeftMammaryGlandResult"); }
        }

        private string dataRightMammaryGland;

        public string DataRightMammaryGland
        {
            get { return dataRightMammaryGland; }
            set { dataRightMammaryGland = value; OnPropertyChanged("DataRightMammaryGland"); }
        }

        private string dataRightAgeRelated;

        public string DataRightAgeRelated
        {
            get { return dataRightAgeRelated; }
            set { dataRightAgeRelated = value; OnPropertyChanged("DataRightAgeRelated"); }
        }

        private string dataRightMeanECOfLesion;

        public string DataRightMeanECOfLesion
        {
            get { return dataRightMeanECOfLesion; }
            set { dataRightMeanECOfLesion = value; OnPropertyChanged("DataRightMeanECOfLesion"); }
        }

        private string dataRightFindings;

        public string DataRightFindings
        {
            get { return dataRightFindings; }
            set { dataRightFindings = value; OnPropertyChanged("DataRightFindings"); }
        }

        private string dataRightMammaryGlandResult;

        public string DataRightMammaryGlandResult
        {
            get { return dataRightMammaryGlandResult; }
            set { dataRightMammaryGlandResult = value; OnPropertyChanged("DataRightMammaryGlandResult"); }
        }

        private string dataTotalPts;

        public string DataTotalPts
        {
            get { return dataTotalPts; }
            set { dataTotalPts = value; OnPropertyChanged("DataTotalPts"); }
        }

        private string dataPoint;

        public string DataPoint
        {
            get { return dataPoint; }
            set { dataPoint = value; OnPropertyChanged("DataPoint"); }
        }

        /// <summary>
        /// 克隆此对象
        /// </summary>
        /// <returns></returns>
        Object ICloneable.Clone()
        {
            return this.Clone();
        }
        /// <summary>
        /// 克隆此对象
        /// </summary>
        /// <returns></returns>
        public ShortFormReport Clone()
        {
            return (ShortFormReport)this.MemberwiseClone(); 
        }
    }
}

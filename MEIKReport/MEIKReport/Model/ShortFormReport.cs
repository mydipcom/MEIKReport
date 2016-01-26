using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Model
{
    [Serializable]
    public class ShortFormReport
    {
        private string dataScreenDate;

        public string DataScreenDate
        {
            get { return dataScreenDate; }
            set { dataScreenDate = value; }
        }
        private string dataScreenLocation;

        public string DataScreenLocation
        {
            get { return dataScreenLocation; }
            set { dataScreenLocation = value; }
        }
        private string dataUserCode;

        public string DataUserCode
        {
            get { return dataUserCode; }
            set { dataUserCode = value; }
        }
        private string dataName;

        public string DataName
        {
            get { return dataName; }
            set { dataName = value; }
        }
        private string dataAge;

        public string DataAge
        {
            get { return dataAge; }
            set { dataAge = value; }
        }
        private string dataLeftFinding;

        public string DataLeftFinding
        {
            get { return dataLeftFinding; }
            set { dataLeftFinding = value; }
        }
        private string dataRightFinding;

        public string DataRightFinding
        {
            get { return dataRightFinding; }
            set { dataRightFinding = value; }
        }
        private string dataLeftLocation;

        public string DataLeftLocation
        {
            get { return dataLeftLocation; }
            set { dataLeftLocation = value; }
        }
        private string dataRightLocation;

        public string DataRightLocation
        {
            get { return dataRightLocation; }
            set { dataRightLocation = value; }
        }

        private string dataLeftSize;

        public string DataLeftSize
        {
            get { return dataLeftSize; }
            set { dataLeftSize = value; }
        }
        private string dataRightSize;

        public string DataRightSize
        {
            get { return dataRightSize; }
            set { dataRightSize = value; }
        }
        private string dataBiRadsCategory;

        public string DataBiRadsCategory
        {
            get { return dataBiRadsCategory; }
            set { dataBiRadsCategory = value; }
        }
        private string dataRecommendation;

        public string DataRecommendation
        {
            get { return dataRecommendation; }
            set { dataRecommendation = value; }
        }
        private string dataFurtherExam;

        public string DataFurtherExam
        {
            get { return dataFurtherExam; }
            set { dataFurtherExam = value; }
        }
        private string dataConclusion;

        public string DataConclusion
        {
            get { return dataConclusion; }
            set { dataConclusion = value; }
        }
        private string dataConclusion2;

        public string DataConclusion2
        {
            get { return dataConclusion2; }
            set { dataConclusion2 = value; }
        }
        private string dataComments;

        public string DataComments
        {
            get { return dataComments; }
            set { dataComments = value; }
        }
        private string dataSignDate;

        public string DataSignDate
        {
            get { return dataSignDate; }
            set { dataSignDate = value; }
        }

        private byte[] dataSignImg;

        public byte[] DataSignImg
        {
            get { return dataSignImg; }
            set { dataSignImg = value; }
        }
        private string dataMeikTech;

        public string DataMeikTech
        {
            get { return dataMeikTech; }
            set { dataMeikTech = value; }
        }
        private string dataDoctor;

        public string DataDoctor
        {
            get { return dataDoctor; }
            set { dataDoctor = value; }
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
            set { dataScreenShotImg = value; }
        }




        private string dataGender;

        public string DataGender
        {
            get { return dataGender; }
            set { dataGender = value; }
        }

        private string dataAddress;

        public string DataAddress
        {
            get { return dataAddress; }
            set { dataAddress = value; }
        }

        private string dataHealthCard;

        public string DataHealthCard
        {
            get { return dataHealthCard; }
            set { dataHealthCard = value; }
        }

        private string dataWeight;

        public string DataWeight
        {
            get { return dataWeight; }
            set { dataWeight = value; }
        }

        private string dataWeightUnit;

        public string DataWeightUnit
        {
            get { return dataWeightUnit; }
            set { dataWeightUnit = value; }
        }

        private string dataMenstrualCycle;

        public string DataMenstrualCycle
        {
            get { return dataMenstrualCycle; }
            set { dataMenstrualCycle = value; }
        }

        private string dataHormones;

        public string DataHormones
        {
            get { return dataHormones; }
            set { dataHormones = value; }
        }

        private string dataSkinAffections;

        public string DataSkinAffections
        {
            get { return dataSkinAffections; }
            set { dataSkinAffections = value; }
        }

        private string dataPertinentHistory;

        public string DataPertinentHistory
        {
            get { return dataPertinentHistory; }
            set { dataPertinentHistory = value; }
        }

        private string dataPertinentHistory1;

        public string DataPertinentHistory1
        {
            get { return dataPertinentHistory1; }
            set { dataPertinentHistory1 = value; }
        }

        private string dataMotherUltra;

        public string DataMotherUltra
        {
            get { return dataMotherUltra; }
            set { dataMotherUltra = value; }
        }

        private string dataLeftBreast;

        public string DataLeftBreast
        {
            get { return dataLeftBreast; }
            set { dataLeftBreast = value; }
        }

        private string dataRightBreast;

        public string DataRightBreast
        {
            get { return dataRightBreast; }
            set { dataRightBreast = value; }
        }

        private string dataLeftPalpableMass;

        public string DataLeftPalpableMass
        {
            get { return dataLeftPalpableMass; }
            set { dataLeftPalpableMass = value; }
        }

        private string dataRightPalpableMass;

        public string DataRightPalpableMass
        {
            get { return dataRightPalpableMass; }
            set { dataRightPalpableMass = value; }
        }

        private string dataLeftChangesOfElectricalConductivity;

        public string DataLeftChangesOfElectricalConductivity
        {
            get { return dataLeftChangesOfElectricalConductivity; }
            set { dataLeftChangesOfElectricalConductivity = value; }
        }

        private string dataRightChangesOfElectricalConductivity;

        public string DataRightChangesOfElectricalConductivity
        {
            get { return dataRightChangesOfElectricalConductivity; }
            set { dataRightChangesOfElectricalConductivity = value; }
        }

        private string dataLeftMammaryStruct;

        public string DataLeftMammaryStruct
        {
            get { return dataLeftMammaryStruct; }
            set { dataLeftMammaryStruct = value; }
        }

        private string dataRightMammaryStruct;

        public string DataRightMammaryStruct
        {
            get { return dataRightMammaryStruct; }
            set { dataRightMammaryStruct = value; }
        }

        private string dataLeftLactiferousSinusZone;

        public string DataLeftLactiferousSinusZone
        {
            get { return dataLeftLactiferousSinusZone; }
            set { dataLeftLactiferousSinusZone = value; }
        }

        private string dataRightLactiferousSinusZone;

        public string DataRightLactiferousSinusZone
        {
            get { return dataRightLactiferousSinusZone; }
            set { dataRightLactiferousSinusZone = value; }
        }

        private string dataLeftMammaryContour;

        public string DataLeftMammaryContour
        {
            get { return dataLeftMammaryContour; }
            set { dataLeftMammaryContour = value; }
        }

        private string dataRightMammaryContour;

        public string DataRightMammaryContour
        {
            get { return dataRightMammaryContour; }
            set { dataRightMammaryContour = value; }
        }

        private string dataLeftNumber;

        public string DataLeftNumber
        {
            get { return dataLeftNumber; }
            set { dataLeftNumber = value; }
        }

        private string dataRightNumber;

        public string DataRightNumber
        {
            get { return dataRightNumber; }
            set { dataRightNumber = value; }
        }

        private string dataLeftShape;

        public string DataLeftShape
        {
            get { return dataLeftShape; }
            set { dataLeftShape = value; }
        }

        private string dataRightShape;

        public string DataRightShape
        {
            get { return dataRightShape; }
            set { dataRightShape = value; }
        }

        private string dataLeftContourAroundFocus;

        public string DataLeftContourAroundFocus
        {
            get { return dataLeftContourAroundFocus; }
            set { dataLeftContourAroundFocus = value; }
        }

        private string dataRightContourAroundFocus;

        public string DataRightContourAroundFocus
        {
            get { return dataRightContourAroundFocus; }
            set { dataRightContourAroundFocus = value; }
        }

        private string dataLeftInternalElectricalStructure;

        public string DataLeftInternalElectricalStructure
        {
            get { return dataLeftInternalElectricalStructure; }
            set { dataLeftInternalElectricalStructure = value; }
        }

        private string dataRightInternalElectricalStructure;

        public string DataRightInternalElectricalStructure
        {
            get { return dataRightInternalElectricalStructure; }
            set { dataRightInternalElectricalStructure = value; }
        }

        private string dataLeftSurroundingTissues;

        public string DataLeftSurroundingTissues
        {
            get { return dataLeftSurroundingTissues; }
            set { dataLeftSurroundingTissues = value; }
        }

        private string dataRightSurroundingTissues;

        public string DataRightSurroundingTissues
        {
            get { return dataRightSurroundingTissues; }
            set { dataRightSurroundingTissues = value; }
        }

        private string dataLeftOncomarkerHighlightBenignChanges;

        public string DataLeftOncomarkerHighlightBenignChanges
        {
            get { return dataLeftOncomarkerHighlightBenignChanges; }
            set { dataLeftOncomarkerHighlightBenignChanges = value; }
        }

        private string dataRightOncomarkerHighlightBenignChanges;

        public string DataRightOncomarkerHighlightBenignChanges
        {
            get { return dataRightOncomarkerHighlightBenignChanges; }
            set { dataRightOncomarkerHighlightBenignChanges = value; }
        }

        private string dataLeftOncomarkerHighlightSuspiciousChanges;

        public string DataLeftOncomarkerHighlightSuspiciousChanges
        {
            get { return dataLeftOncomarkerHighlightSuspiciousChanges; }
            set { dataLeftOncomarkerHighlightSuspiciousChanges = value; }
        }

        private string dataRightOncomarkerHighlightSuspiciousChanges;

        public string DataRightOncomarkerHighlightSuspiciousChanges
        {
            get { return dataRightOncomarkerHighlightSuspiciousChanges; }
            set { dataRightOncomarkerHighlightSuspiciousChanges = value; }
        }

        private string dataLeftMeanElectricalConductivity1;

        public string DataLeftMeanElectricalConductivity1
        {
            get { return dataLeftMeanElectricalConductivity1; }
            set { dataLeftMeanElectricalConductivity1 = value; }
        }

        private string dataLeftMeanElectricalConductivity1N1;

        public string DataLeftMeanElectricalConductivity1N1
        {
            get { return dataLeftMeanElectricalConductivity1N1; }
            set { dataLeftMeanElectricalConductivity1N1 = value; }
        }

        private string dataLeftMeanElectricalConductivity1N2;

        public string DataLeftMeanElectricalConductivity1N2
        {
            get { return dataLeftMeanElectricalConductivity1N2; }
            set { dataLeftMeanElectricalConductivity1N2 = value; }
        }

        private string dataRightMeanElectricalConductivity1;

        public string DataRightMeanElectricalConductivity1
        {
            get { return dataRightMeanElectricalConductivity1; }
            set { dataRightMeanElectricalConductivity1 = value; }
        }

        private string dataRightMeanElectricalConductivity1N1;

        public string DataRightMeanElectricalConductivity1N1
        {
            get { return dataRightMeanElectricalConductivity1N1; }
            set { dataRightMeanElectricalConductivity1N1 = value; }
        }

        private string dataRightMeanElectricalConductivity1N2;

        public string DataRightMeanElectricalConductivity1N2
        {
            get { return dataRightMeanElectricalConductivity1N2; }
            set { dataRightMeanElectricalConductivity1N2 = value; }
        }

        private string dataLeftMeanElectricalConductivity2;

        public string DataLeftMeanElectricalConductivity2
        {
            get { return dataLeftMeanElectricalConductivity2; }
            set { dataLeftMeanElectricalConductivity2 = value; }
        }

        private string dataLeftMeanElectricalConductivity2N1;

        public string DataLeftMeanElectricalConductivity2N1
        {
            get { return dataLeftMeanElectricalConductivity2N1; }
            set { dataLeftMeanElectricalConductivity2N1 = value; }
        }

        private string dataLeftMeanElectricalConductivity2N2;

        public string DataLeftMeanElectricalConductivity2N2
        {
            get { return dataLeftMeanElectricalConductivity2N2; }
            set { dataLeftMeanElectricalConductivity2N2 = value; }
        }

        private string dataRightMeanElectricalConductivity2;

        public string DataRightMeanElectricalConductivity2
        {
            get { return dataRightMeanElectricalConductivity2; }
            set { dataRightMeanElectricalConductivity2 = value; }
        }

        private string dataRightMeanElectricalConductivity2N1;

        public string DataRightMeanElectricalConductivity2N1
        {
            get { return dataRightMeanElectricalConductivity2N1; }
            set { dataRightMeanElectricalConductivity2N1 = value; }
        }

        private string dataRightMeanElectricalConductivity2N2;

        public string DataRightMeanElectricalConductivity2N2
        {
            get { return dataRightMeanElectricalConductivity2N2; }
            set { dataRightMeanElectricalConductivity2N2 = value; }
        }

        private string dataMeanElectricalConductivity3;

        public string DataMeanElectricalConductivity3
        {
            get { return dataMeanElectricalConductivity3; }
            set { dataMeanElectricalConductivity3 = value; }
        }

        private string dataLeftMeanElectricalConductivity3;

        public string DataLeftMeanElectricalConductivity3
        {
            get { return dataLeftMeanElectricalConductivity3; }
            set { dataLeftMeanElectricalConductivity3 = value; }
        }

        private string dataLeftMeanElectricalConductivity3N1;

        public string DataLeftMeanElectricalConductivity3N1
        {
            get { return dataLeftMeanElectricalConductivity3N1; }
            set { dataLeftMeanElectricalConductivity3N1 = value; }
        }

        private string dataLeftMeanElectricalConductivity3N2;

        public string DataLeftMeanElectricalConductivity3N2
        {
            get { return dataLeftMeanElectricalConductivity3N2; }
            set { dataLeftMeanElectricalConductivity3N2 = value; }
        }

        private string dataRightMeanElectricalConductivity3;

        public string DataRightMeanElectricalConductivity3
        {
            get { return dataRightMeanElectricalConductivity3; }
            set { dataRightMeanElectricalConductivity3 = value; }
        }

        private string dataRightMeanElectricalConductivity3N1;

        public string DataRightMeanElectricalConductivity3N1
        {
            get { return dataRightMeanElectricalConductivity3N1; }
            set { dataRightMeanElectricalConductivity3N1 = value; }
        }

        private string dataRightMeanElectricalConductivity3N2;

        public string DataRightMeanElectricalConductivity3N2
        {
            get { return dataRightMeanElectricalConductivity3N2; }
            set { dataRightMeanElectricalConductivity3N2 = value; }
        }

        private string dataComparativeElectricalConductivityReference1;

        public string DataComparativeElectricalConductivityReference1
        {
            get { return dataComparativeElectricalConductivityReference1; }
            set { dataComparativeElectricalConductivityReference1 = value; }
        }

        private string dataLeftComparativeElectricalConductivity1;

        public string DataLeftComparativeElectricalConductivity1
        {
            get { return dataLeftComparativeElectricalConductivity1; }
            set { dataLeftComparativeElectricalConductivity1 = value; }
        }

        private string dataRightComparativeElectricalConductivity1;

        public string DataRightComparativeElectricalConductivity1
        {
            get { return dataRightComparativeElectricalConductivity1; }
            set { dataRightComparativeElectricalConductivity1 = value; }
        }

        private string dataComparativeElectricalConductivityReference2;

        public string DataComparativeElectricalConductivityReference2
        {
            get { return dataComparativeElectricalConductivityReference2; }
            set { dataComparativeElectricalConductivityReference2 = value; }
        }

        private string dataLeftComparativeElectricalConductivity2;

        public string DataLeftComparativeElectricalConductivity2
        {
            get { return dataLeftComparativeElectricalConductivity2; }
            set { dataLeftComparativeElectricalConductivity2 = value; }
        }

        private string dataRightComparativeElectricalConductivity2;

        public string DataRightComparativeElectricalConductivity2
        {
            get { return dataRightComparativeElectricalConductivity2; }
            set { dataRightComparativeElectricalConductivity2 = value; }
        }

        private string dataComparativeElectricalConductivity3;

        public string DataComparativeElectricalConductivity3
        {
            get { return dataComparativeElectricalConductivity3; }
            set { dataComparativeElectricalConductivity3 = value; }
        }

        private string dataLeftComparativeElectricalConductivity3;

        public string DataLeftComparativeElectricalConductivity3
        {
            get { return dataLeftComparativeElectricalConductivity3; }
            set { dataLeftComparativeElectricalConductivity3 = value; }
        }

        private string dataRightComparativeElectricalConductivity3;

        public string DataRightComparativeElectricalConductivity3
        {
            get { return dataRightComparativeElectricalConductivity3; }
            set { dataRightComparativeElectricalConductivity3 = value; }
        }

        private string dataDivergenceBetweenHistogramsReference1;

        public string DataDivergenceBetweenHistogramsReference1
        {
            get { return dataDivergenceBetweenHistogramsReference1; }
            set { dataDivergenceBetweenHistogramsReference1 = value; }
        }

        private string dataLeftDivergenceBetweenHistograms1;

        public string DataLeftDivergenceBetweenHistograms1
        {
            get { return dataLeftDivergenceBetweenHistograms1; }
            set { dataLeftDivergenceBetweenHistograms1 = value; }
        }

        private string dataRightDivergenceBetweenHistograms1;

        public string DataRightDivergenceBetweenHistograms1
        {
            get { return dataRightDivergenceBetweenHistograms1; }
            set { dataRightDivergenceBetweenHistograms1 = value; }
        }

        private string dataDivergenceBetweenHistogramsReference2;

        public string DataDivergenceBetweenHistogramsReference2
        {
            get { return dataDivergenceBetweenHistogramsReference2; }
            set { dataDivergenceBetweenHistogramsReference2 = value; }
        }

        private string dataLeftDivergenceBetweenHistograms2;

        public string DataLeftDivergenceBetweenHistograms2
        {
            get { return dataLeftDivergenceBetweenHistograms2; }
            set { dataLeftDivergenceBetweenHistograms2 = value; }
        }

        private string dataRightDivergenceBetweenHistograms2;

        public string DataRightDivergenceBetweenHistograms2
        {
            get { return dataRightDivergenceBetweenHistograms2; }
            set { dataRightDivergenceBetweenHistograms2 = value; }
        }

        private string dataDivergenceBetweenHistograms3;

        public string DataDivergenceBetweenHistograms3
        {
            get { return dataDivergenceBetweenHistograms3; }
            set { dataDivergenceBetweenHistograms3 = value; }
        }

        private string dataLeftDivergenceBetweenHistograms3;

        public string DataLeftDivergenceBetweenHistograms3
        {
            get { return dataLeftDivergenceBetweenHistograms3; }
            set { dataLeftDivergenceBetweenHistograms3 = value; }
        }

        private string dataRightDivergenceBetweenHistograms3;

        public string DataRightDivergenceBetweenHistograms3
        {
            get { return dataRightDivergenceBetweenHistograms3; }
            set { dataRightDivergenceBetweenHistograms3 = value; }
        }

        private string dataLeftComparisonWithNorm;

        public string DataLeftComparisonWithNorm
        {
            get { return dataLeftComparisonWithNorm; }
            set { dataLeftComparisonWithNorm = value; }
        }

        private string dataRightComparisonWithNorm;

        public string DataRightComparisonWithNorm
        {
            get { return dataRightComparisonWithNorm; }
            set { dataRightComparisonWithNorm = value; }
        }

        private string dataPhaseElectricalConductivityReference;

        public string DataPhaseElectricalConductivityReference
        {
            get { return dataPhaseElectricalConductivityReference; }
            set { dataPhaseElectricalConductivityReference = value; }
        }

        private string dataLeftPhaseElectricalConductivity;

        public string DataLeftPhaseElectricalConductivity
        {
            get { return dataLeftPhaseElectricalConductivity; }
            set { dataLeftPhaseElectricalConductivity = value; }
        }

        private string dataRightPhaseElectricalConductivity;

        public string DataRightPhaseElectricalConductivity
        {
            get { return dataRightPhaseElectricalConductivity; }
            set { dataRightPhaseElectricalConductivity = value; }
        }

        private string dataAgeElectricalConductivityReference;

        public string DataAgeElectricalConductivityReference
        {
            get { return dataAgeElectricalConductivityReference; }
            set { dataAgeElectricalConductivityReference = value; }
        }

        private string dataLeftAgeElectricalConductivity;

        public string DataLeftAgeElectricalConductivity
        {
            get { return dataLeftAgeElectricalConductivity; }
            set { dataLeftAgeElectricalConductivity = value; }
        }

        private string dataRightAgeElectricalConductivity;

        public string DataRightAgeElectricalConductivity
        {
            get { return dataRightAgeElectricalConductivity; }
            set { dataRightAgeElectricalConductivity = value; }
        }

        private string dataAgeValueOfEC;

        public string DataAgeValueOfEC
        {
            get { return dataAgeValueOfEC; }
            set { dataAgeValueOfEC = value; }
        }

        private string dataExamConclusion;

        public string DataExamConclusion
        {
            get { return dataExamConclusion; }
            set { dataExamConclusion = value; }
        }

        private string dataLeftMammaryGland;

        public string DataLeftMammaryGland
        {
            get { return dataLeftMammaryGland; }
            set { dataLeftMammaryGland = value; }
        }

        private string dataLeftAgeRelated;

        public string DataLeftAgeRelated
        {
            get { return dataLeftAgeRelated; }
            set { dataLeftAgeRelated = value; }
        }

        private string dataLeftMeanECOfLesion;

        public string DataLeftMeanECOfLesion
        {
            get { return dataLeftMeanECOfLesion; }
            set { dataLeftMeanECOfLesion = value; }
        }

        private string dataLeftFindings;

        public string DataLeftFindings
        {
            get { return dataLeftFindings; }
            set { dataLeftFindings = value; }
        }

        private string dataLeftMammaryGlandResult;

        public string DataLeftMammaryGlandResult
        {
            get { return dataLeftMammaryGlandResult; }
            set { dataLeftMammaryGlandResult = value; }
        }

        private string dataRightMammaryGland;

        public string DataRightMammaryGland
        {
            get { return dataRightMammaryGland; }
            set { dataRightMammaryGland = value; }
        }

        private string dataRightAgeRelated;

        public string DataRightAgeRelated
        {
            get { return dataRightAgeRelated; }
            set { dataRightAgeRelated = value; }
        }

        private string dataRightMeanECOfLesion;

        public string DataRightMeanECOfLesion
        {
            get { return dataRightMeanECOfLesion; }
            set { dataRightMeanECOfLesion = value; }
        }

        private string dataRightFindings;

        public string DataRightFindings
        {
            get { return dataRightFindings; }
            set { dataRightFindings = value; }
        }

        private string dataRightMammaryGlandResult;

        public string DataRightMammaryGlandResult
        {
            get { return dataRightMammaryGlandResult; }
            set { dataRightMammaryGlandResult = value; }
        }

        private string dataTotalPts;

        public string DataTotalPts
        {
            get { return dataTotalPts; }
            set { dataTotalPts = value; }
        }

        private string dataPoint;

        public string DataPoint
        {
            get { return dataPoint; }
            set { dataPoint = value; }
        }


    }
}

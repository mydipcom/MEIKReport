using MEIKReport.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace MEIKReport.Common
{
    public class WordTools
    {
        public static ShortFormReport ReadWordFile(string wordFilePath)
        {
            if (!File.Exists(wordFilePath))
            {
                return null;
            }
            Word.Application app = new Word.Application();            
            object unknow = Type.Missing;
            if (app.Documents.Count > 0)
            {
                app.ActiveDocument.Close();
            }
            app.Visible = false;
            Object file = wordFilePath;
            Word.Document doc = null;
            
            ShortFormReport ShortFormReport = new ShortFormReport();
            try
            {
                doc = app.Documents.Open(ref file);
                string dataMenstrualCycle = doc.FormFields[11].Result;
                if (string.IsNullOrEmpty(dataMenstrualCycle))
                {
                    ShortFormReport.DataMenstrualCycle = "";
                }
                else if (dataMenstrualCycle.StartsWith("1 phase") || dataMenstrualCycle.StartsWith("phase 1") || dataMenstrualCycle.StartsWith("第一阶段") || dataMenstrualCycle.StartsWith("一期"))
                {
                    ShortFormReport.DataMenstrualCycle = "1";// App.Current.FindResource("ReportContext_15").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("2 phase") || dataMenstrualCycle.StartsWith("phase 2") || dataMenstrualCycle.StartsWith("第二阶段") || dataMenstrualCycle.StartsWith("二期"))
                {
                    ShortFormReport.DataMenstrualCycle = "2";//App.Current.FindResource("ReportContext_16").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("1 and 2 phase") || dataMenstrualCycle.StartsWith("第一和第二阶段"))
                {
                    ShortFormReport.DataMenstrualCycle = "3";//App.Current.FindResource("ReportContext_17").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("dysmenorrhea") || dataMenstrualCycle.StartsWith("痛经"))
                {
                    ShortFormReport.DataMenstrualCycle = "4";// App.Current.FindResource("ReportContext_18").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("missing") || dataMenstrualCycle.StartsWith("绝经期"))
                {
                    ShortFormReport.DataMenstrualCycle = "5";// App.Current.FindResource("ReportContext_19").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("pregnancy") || dataMenstrualCycle.StartsWith("孕期"))
                {
                    ShortFormReport.DataMenstrualCycle = "6";// App.Current.FindResource("ReportContext_20").ToString();
                }
                else if (dataMenstrualCycle.StartsWith("lactation") || dataMenstrualCycle.StartsWith("哺乳期"))
                {
                    ShortFormReport.DataMenstrualCycle = "7";// App.Current.FindResource("ReportContext_21").ToString();
                }
                //ShortFormReport.DataMenstrualCycle = doc.FormFields[11].Result;
                //ShortFormReport.DataLeftChangesOfElectricalConductivity = doc.FormFields[15].Result;
                //ShortFormReport.DataRightChangesOfElectricalConductivity = doc.FormFields[16].Result;
                //ShortFormReport.DataLeftMammaryStruct = doc.FormFields[17].Result;
                //ShortFormReport.DataRightMammaryStruct = doc.FormFields[18].Result;
                //ShortFormReport.DataLeftLactiferousSinusZone = doc.FormFields[19].Result;
                //ShortFormReport.DataRightLactiferousSinusZone = doc.FormFields[20].Result;
                //ShortFormReport.DataLeftMammaryContour = doc.FormFields[21].Result;
                //ShortFormReport.DataLeftMammaryContour = doc.FormFields[22].Result;

                //ShortFormReport.DataLeftLocation = doc.FormFields[23].Result;
                //ShortFormReport.DataRightLocation = doc.FormFields[24].Result;
                //ShortFormReport.DataLeftNumber = doc.FormFields[25].Result;
                //ShortFormReport.DataRightNumber = doc.FormFields[26].Result;
                //ShortFormReport.DataLeftSize = doc.FormFields[27].Result;
                //ShortFormReport.DataRightSize = doc.FormFields[28].Result;
                //ShortFormReport.DataLeftShape = doc.FormFields[29].Result;
                //ShortFormReport.DataRightShape = doc.FormFields[30].Result;
                //ShortFormReport.DataLeftContourAroundFocus = doc.FormFields[31].Result;
                //ShortFormReport.DataRightContourAroundFocus = doc.FormFields[32].Result;
                //ShortFormReport.DataLeftInternalElectricalStructure = doc.FormFields[33].Result;
                //ShortFormReport.DataRightInternalElectricalStructure = doc.FormFields[34].Result;
                //ShortFormReport.DataLeftSurroundingTissues = doc.FormFields[35].Result;
                //ShortFormReport.DataRightSurroundingTissues = doc.FormFields[36].Result;


                ShortFormReport.DataLeftMeanElectricalConductivity1 = doc.FormFields[37].Result;
                ShortFormReport.DataRightMeanElectricalConductivity1 = doc.FormFields[38].Result;
                ShortFormReport.DataLeftMeanElectricalConductivity2 = doc.FormFields[39].Result;
                ShortFormReport.DataRightMeanElectricalConductivity2 = doc.FormFields[40].Result;
                string dataMeanElectricalConductivity3 = doc.FormFields[41].Result;
                if (string.IsNullOrEmpty(dataMeanElectricalConductivity3))
                {
                    ShortFormReport.DataMeanElectricalConductivity3 = "0";//"";
                }
                else if (dataMeanElectricalConductivity3.StartsWith("postmenopause") || dataMeanElectricalConductivity3.StartsWith("绝经后期"))
                {
                    ShortFormReport.DataMeanElectricalConductivity3 = "1";// App.Current.FindResource("ReportContext_103").ToString();
                }
                else if (dataMeanElectricalConductivity3.StartsWith("pregnancy") || dataMeanElectricalConductivity3.StartsWith("妊娠"))
                {
                    ShortFormReport.DataMeanElectricalConductivity3 = "2";// App.Current.FindResource("ReportContext_104").ToString();
                }
                else if (dataMeanElectricalConductivity3.StartsWith("lactation") || dataMeanElectricalConductivity3.StartsWith("哺乳期"))
                {
                    ShortFormReport.DataMeanElectricalConductivity3 = "3";// App.Current.FindResource("ReportContext_105").ToString();
                }
                //ShortFormReport.DataMeanElectricalConductivity3 = doc.FormFields[41].Result;
                ShortFormReport.DataLeftMeanElectricalConductivity3 = doc.FormFields[42].Result;
                ShortFormReport.DataRightMeanElectricalConductivity3 = doc.FormFields[43].Result;

                ShortFormReport.DataLeftComparativeElectricalConductivity1 = doc.FormFields[44].Result;
                //ShortFormReport.DataRightComparativeElectricalConductivity1 = doc.FormFields[44].Result;
                ShortFormReport.DataLeftComparativeElectricalConductivity2 = doc.FormFields[45].Result;
                //ShortFormReport.DataRightComparativeElectricalConductivity2 = doc.FormFields[45].Result;
                ShortFormReport.DataLeftComparativeElectricalConductivity3 = doc.FormFields[46].Result;
                //ShortFormReport.DataRightComparativeElectricalConductivity3 = doc.FormFields[46].Result;
                ShortFormReport.DataLeftDivergenceBetweenHistograms1 = doc.FormFields[47].Result;
                //ShortFormReport.DataRightDivergenceBetweenHistograms1 = doc.FormFields[47].Result;
                ShortFormReport.DataLeftDivergenceBetweenHistograms2 = doc.FormFields[48].Result;
                //ShortFormReport.DataRightDivergenceBetweenHistograms2 = doc.FormFields[48].Result;
                ShortFormReport.DataLeftDivergenceBetweenHistograms3 = doc.FormFields[49].Result;
                //ShortFormReport.DataRightDivergenceBetweenHistograms3 = doc.FormFields[49].Result;

                ShortFormReport.DataLeftPhaseElectricalConductivity = doc.FormFields[54].Result;
                ShortFormReport.DataRightPhaseElectricalConductivity = doc.FormFields[55].Result;

                ShortFormReport.DataAgeElectricalConductivityReference = doc.FormFields[56].Result;

                string dataLeftAgeElectricalConductivity = doc.FormFields[57].Result;
                if (string.IsNullOrEmpty(dataLeftAgeElectricalConductivity))
                {
                    ShortFormReport.DataLeftAgeElectricalConductivity = "0";// "";
                }
                else if (dataLeftAgeElectricalConductivity.StartsWith("<5"))
                {
                    ShortFormReport.DataLeftAgeElectricalConductivity = "1";// App.Current.FindResource("ReportContext_111").ToString();
                }
                else if (dataLeftAgeElectricalConductivity.StartsWith(">95"))
                {
                    ShortFormReport.DataLeftAgeElectricalConductivity = "3";// App.Current.FindResource("ReportContext_113").ToString();
                }
                else 
                {
                    ShortFormReport.DataLeftAgeElectricalConductivity = "2";// App.Current.FindResource("ReportContext_112").ToString();
                }
                //ShortFormReport.DataLeftAgeElectricalConductivity = doc.FormFields[57].Result;
                string dataRightAgeElectricalConductivity = doc.FormFields[58].Result;
                if (string.IsNullOrEmpty(dataRightAgeElectricalConductivity))
                {
                    ShortFormReport.DataRightAgeElectricalConductivity = "0";// "";
                }
                else if (dataRightAgeElectricalConductivity.StartsWith("<5"))
                {
                    ShortFormReport.DataRightAgeElectricalConductivity = "1";// App.Current.FindResource("ReportContext_111").ToString();
                }
                else if (dataRightAgeElectricalConductivity.StartsWith(">95"))
                {
                    ShortFormReport.DataRightAgeElectricalConductivity = "3";// App.Current.FindResource("ReportContext_113").ToString();
                }
                else
                {
                    ShortFormReport.DataRightAgeElectricalConductivity = "2";// App.Current.FindResource("ReportContext_112").ToString();
                }
                //ShortFormReport.DataRightAgeElectricalConductivity = doc.FormFields[58].Result;
                string dataExamConclusion = doc.FormFields[59].Result;
                if (string.IsNullOrEmpty(dataExamConclusion))
                {
                    ShortFormReport.DataExamConclusion = "0";//"";
                }
                else if (dataExamConclusion.StartsWith("Pubertal Period") || dataExamConclusion.StartsWith("青春期"))
                {
                    ShortFormReport.DataExamConclusion = "1";// App.Current.FindResource("ReportContext_116").ToString();
                }
                else if (dataExamConclusion.StartsWith("Early childbearing age") || dataExamConclusion.StartsWith("育龄早期"))
                {
                    ShortFormReport.DataExamConclusion = "2";// App.Current.FindResource("ReportContext_117").ToString();
                }
                else if (dataExamConclusion.StartsWith("Childbearing age") || dataExamConclusion.StartsWith("育龄期"))
                {
                    ShortFormReport.DataExamConclusion = "3";// App.Current.FindResource("ReportContext_118").ToString();
                }
                else if (dataExamConclusion.StartsWith("Perimenopausal period") || dataExamConclusion.StartsWith("围绝经期"))
                {
                    ShortFormReport.DataExamConclusion = "4";// App.Current.FindResource("ReportContext_119").ToString();
                }
                else if (dataExamConclusion.StartsWith("Postmenopausal period") || dataExamConclusion.StartsWith("绝经期"))
                {
                    ShortFormReport.DataExamConclusion = "5";// App.Current.FindResource("ReportContext_120").ToString();
                }
                //ShortFormReport.DataExamConclusion = doc.FormFields[59].Result;
                string dataLeftMammaryGland=doc.FormFields[60].Result;
                if (string.IsNullOrEmpty(dataLeftMammaryGland))
                {
                    ShortFormReport.DataLeftMammaryGland = "0";// "";
                }
                else if (dataLeftMammaryGland.StartsWith("Ductal type") || dataLeftMammaryGland.StartsWith("导管型乳腺结构") || dataLeftMammaryGland.StartsWith("导管式结构"))
                {
                    ShortFormReport.DataLeftMammaryGland = "5";// App.Current.FindResource("ReportContext_126").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type with ductal component predominance") || dataLeftMammaryGland.StartsWith("混合型，导管型结构优势") || dataLeftMammaryGland.StartsWith("导管成分优先的"))
                {
                    ShortFormReport.DataLeftMammaryGland = "4";// App.Current.FindResource("ReportContext_125").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type of mammary gland structure") || dataLeftMammaryGland.StartsWith("Mixed type of structure") || dataLeftMammaryGland.StartsWith("混合型乳腺结构") || dataLeftMammaryGland.StartsWith("混合式结构"))
                {
                    ShortFormReport.DataLeftMammaryGland = "3";// App.Current.FindResource("ReportContext_124").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type with amorphous component predominance") || dataLeftMammaryGland.StartsWith("混合型，无定型结构优势") || ( dataLeftMammaryGland.Contains("无") && dataLeftMammaryGland.Contains("混合")))
                {
                    ShortFormReport.DataLeftMammaryGland = "2";// App.Current.FindResource("ReportContext_123").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Amorphous type") || dataLeftMammaryGland.StartsWith("无定型乳腺结构") || (dataLeftMammaryGland.Contains("无") && !dataLeftMammaryGland.Contains("混合")))
                {
                    ShortFormReport.DataLeftMammaryGland = "1";// App.Current.FindResource("ReportContext_122").ToString();
                }
                //ShortFormReport.DataLeftMammaryGland = doc.FormFields[60].Result;
                string dataLeftAgeRelated = doc.FormFields[61].Result;
                if (string.IsNullOrEmpty(dataLeftAgeRelated))
                {
                    ShortFormReport.DataLeftAgeRelated = "0";//"";
                }
                else if (dataLeftAgeRelated.StartsWith("<5"))
                {
                    ShortFormReport.DataLeftAgeRelated = "1";// App.Current.FindResource("ReportContext_111").ToString();
                }
                else if (dataLeftAgeRelated.StartsWith(">95"))
                {
                    ShortFormReport.DataLeftAgeRelated = "3";// App.Current.FindResource("ReportContext_113").ToString();
                }
                else
                {
                    ShortFormReport.DataLeftAgeRelated = "2";// App.Current.FindResource("ReportContext_112").ToString();
                }
                //ShortFormReport.DataLeftAgeRelated = doc.FormFields[61].Result;

                string dataRightMammaryGland = doc.FormFields[64].Result;
                if (string.IsNullOrEmpty(dataLeftMammaryGland))
                {
                    ShortFormReport.DataRightMammaryGland = "0";// "";
                }
                else if (dataLeftMammaryGland.StartsWith("Ductal type") || dataLeftMammaryGland.StartsWith("导管型乳腺结构") || dataLeftMammaryGland.StartsWith("导管式结构"))
                {
                    ShortFormReport.DataRightMammaryGland = "5";// App.Current.FindResource("ReportContext_126").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type with ductal component predominance") || dataLeftMammaryGland.StartsWith("混合型，导管型结构优势") || dataLeftMammaryGland.StartsWith("导管成分优先的"))
                {
                    ShortFormReport.DataRightMammaryGland = "4";// App.Current.FindResource("ReportContext_125").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type of mammary gland structure") || dataLeftMammaryGland.StartsWith("Mixed type of structure") || dataLeftMammaryGland.StartsWith("混合型乳腺结构") || dataLeftMammaryGland.StartsWith("混合式结构"))
                {
                    ShortFormReport.DataRightMammaryGland = "3";// App.Current.FindResource("ReportContext_124").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Mixed type with amorphous component predominance") || dataLeftMammaryGland.StartsWith("混合型，无定型结构优势") || (dataLeftMammaryGland.Contains("无") && dataLeftMammaryGland.Contains("混合")))
                {
                    ShortFormReport.DataRightMammaryGland = "2";// App.Current.FindResource("ReportContext_123").ToString();
                }
                else if (dataLeftMammaryGland.StartsWith("Amorphous type") || dataLeftMammaryGland.StartsWith("无定型乳腺结构") || (dataLeftMammaryGland.Contains("无") && !dataLeftMammaryGland.Contains("混合")))
                {
                    ShortFormReport.DataRightMammaryGland = "1";// App.Current.FindResource("ReportContext_122").ToString();
                }
                //ShortFormReport.DataLeftMammaryGland = doc.FormFields[60].Result;
                string dataRightAgeRelated = doc.FormFields[65].Result;
                if (string.IsNullOrEmpty(dataRightAgeRelated))
                {
                    ShortFormReport.DataRightAgeRelated = "0";// "";
                }
                else if (dataRightAgeRelated.StartsWith("<5"))
                {
                    ShortFormReport.DataRightAgeRelated = "1";// App.Current.FindResource("ReportContext_111").ToString();
                }
                else if (dataRightAgeRelated.StartsWith(">95"))
                {
                    ShortFormReport.DataRightAgeRelated = "3";// App.Current.FindResource("ReportContext_113").ToString();
                }
                else
                {
                    ShortFormReport.DataRightAgeRelated = "2";// App.Current.FindResource("ReportContext_112").ToString();
                }

                //ShortFormReport.DataRightMammaryGland = doc.FormFields[64].Result;
                //ShortFormReport.DataRightAgeRelated = doc.FormFields[65].Result;
                                

            }
            catch(Exception ex){}
            finally
            {
                Type wordType = app.GetType();
                try
                {
                    if (doc != null)
                    {
                        doc.Close();
                    }
                    if (app != null)
                    {
                        app.Quit();
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);
                        doc = null;
                        app = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    catch (Exception e) { }
                }
            }
            return ShortFormReport;
        }


    }
}

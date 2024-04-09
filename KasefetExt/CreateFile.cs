using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Common;
using DAL;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace KasefetExt
{
    public class CreateFile
    {
        private IDataLayer dal;
        public Sdg Sdg { get; set; }
        public bool Success { get; set; }
        public Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        public Microsoft.Office.Interop.Excel._Worksheet ExcelWorkSheet;
        private const string CODE_LAB = "607";
        private Product _firstProduct;
        public CreateFile(IDataLayer dal, Sdg sdg)
        {
            try
            {
                this.dal = dal;
                this.Sdg = sdg;
                Success = Create(sdg);
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile("error in CreateFile " + ex.Message, false);
                Logger.WriteLogFile(ex);
                throw;

            }
            finally
            {
                ExcelApp = null;

                ExcelWorkSheet = null;
            }
        }



        public bool Create(Sdg sdg)
        {
            SetFirstProduct(sdg);
            var q = (from sample in sdg.Samples
                     from aliquot in sample.Aliqouts
                     where
                     aliquot.Retest != "T" &&
                     aliquot.Status != "X"
                     && aliquot.TestTemplateEx.TestCode != null
                     select aliquot).ToList();

            if (q == null || q.Count < 1)
            {
                return false;
            }

            CreateExcel(q, columnStrings);

            SaveFile();
            return true;
        }

        private void SetFirstProduct(Sdg sdg)
        {
            if (sdg.Samples != null && sdg.Samples.Count > 0)
            {
                var sample = sdg.Samples.FirstOrDefault();
                if (sample != null) _firstProduct = sample.Product;
            }
        }

        readonly string[] columnStrings = { "קוד נקודה", "קוד בדיקה", "תאריך דיגום", "תוצאת בדיקה", "קוד קבוצה", "קוד מעבדה", "זמן דיגום", "מספר מעבדה/מספר דגימה", "סטטוס", "שם בודק" };

        public void CreateExcel(List<Aliquot> list, string[] columnStrings)
        {
            //    ExcelApp.DefaultSheetDirection = (int)XlDirection.xlToLeft;


            ExcelApp.Workbooks.Add();
            ExcelWorkSheet = ExcelApp.ActiveSheet;
            ExcelWorkSheet.DisplayRightToLeft = true;

            int row = 1;
            for (int i = 0; i < columnStrings.Length; i++)
            {
                ExcelWorkSheet.Cells[row, i + 1] = columnStrings[i];

            }
            foreach (Aliquot aliquot in list)
            {

                Sdg sdg = aliquot.Sample.Sdg;
                Sample sample = aliquot.Sample;
                int col = 1;
                row++;

                string reportedResult = aliquot.U_DEFAULT_VALUE;
                Result r = null;
                if (string.IsNullOrEmpty(reportedResult))
                {
                    //Get reported result
                    r = (from item in aliquot.Tests
                         from result in item.Results
                         where result.REPORTED == "T"
                         orderby result.ResultId ascending
                         select result).FirstOrDefault();

                    if (r != null) reportedResult = r.FormattedResult;
                }
                string groupCode = _firstProduct.U_KASEFET_POOL_GROUP_CODE;

                if (aliquot.ShortName == "Leg") //. בדיקת ליגיונלה (על פי SHORT NAME=Leg) עושה OVERIDE על הקודים הללו והקוד שלה הוא 4.
                {


                    if (_firstProduct.ProductId == 206 || _firstProduct.ProductId == 203)
                    {
                        groupCode = "4";
                    }
                    //עבור מי שתיה 321
                    else if (_firstProduct.ProductId == 202)
                    {
                        groupCode = "321";
                    }
                    if (r != null && r.FormattedResult != null)
                    {
                        var o1 = r.FormattedResult.IndexOf("(");
                        var o2 = r.FormattedResult.IndexOf(")");
                        if (o1 != -1 && o2 != -1)
                        {
                            var ss = r.FormattedResult.Substring(o1, (o2 + 1) - o1);
                            var a = r.FormattedResult.Replace(ss, "");
                            reportedResult = a;
                        }

                    }
                }

                ExcelWorkSheet.Cells[row, col++] = sample.PointCode;
                ExcelWorkSheet.Cells[row, col++] = aliquot.TestTemplateEx.TestCode;
                ExcelWorkSheet.Cells[row, col++] = sample.TextualSamplingTime;
                ExcelWorkSheet.Cells[row, col++] = reportedResult;
                ExcelWorkSheet.Cells[row, col++] = groupCode;
                ExcelWorkSheet.Cells[row, col++] = CODE_LAB;
                if (sample.TextualSamplingTime2 != null)
                    ExcelWorkSheet.Cells[row, col++] = sample.TextualSamplingTime2.Replace(":", "");
                else
                    col++;
                ExcelWorkSheet.Cells[row, col++] = sample.SampleId;
                ExcelWorkSheet.Cells[row, col++] = "Done";
                if (sdg.PerformedOperator != null)
                    ExcelWorkSheet.Cells[row, col] = sdg.PerformedOperator.Name;
            }
            Range c1 = ExcelWorkSheet.Cells[1, 1];
            Range c2 = ExcelWorkSheet.Cells[row, columnStrings.Length];
            var oRange = ExcelWorkSheet.get_Range(c1, c2);
            oRange.EntireColumn.ColumnWidth = 20;
        }


        public void SaveFile()
        {
            var kc = _firstProduct.U_KASEFET_CATEGORY ?? "";
            var fn = "M_" + kc + "_" + CODE_LAB + "_" + Sdg.Name + "_" + DateTime.Now.ToString("ddMMyyyyhhmmss");

            var ph = dal.GetPhraseByName("Location folders");

            var pe = ph.PhraseEntries.FirstOrDefault(p => p.PhraseDescription == "Safe (KASEFET)");
            try
            {


                if (!Directory.Exists(pe.PhraseName))
                {
                    Directory.CreateDirectory(pe.PhraseName);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex.Message);
           
            }
            ExcelWorkSheet.SaveAs(pe.PhraseName + fn);
            ExcelApp.Quit();


        }
    }
}

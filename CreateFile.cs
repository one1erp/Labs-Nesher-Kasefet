using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;
using DAL;
using Microsoft.Office.Interop.Excel;

namespace Kasefet
{
    public class CreateFile
    {
        private IDataLayer dal;
        public Sdg Sdg { get; set; }
        public Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        public Microsoft.Office.Interop.Excel._Worksheet ExcelWorkSheet;

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

                Logger.WriteLogFile(ex);
                throw;

            }
            finally
            {
                ExcelApp = null;

                ExcelWorkSheet = null;
            }
        }


        public bool Success { get; set; }
        public bool Create(Sdg sdg)
        {
            var q = (from sample in sdg.Samples
                     from aliquot in sample.Aliqouts
                     where
                     aliquot.U_CHARGE == "T" &&
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

        readonly string[] columnStrings = { "קוד נקודה", "קוד בדיקה", "תאריך דיגום", "תוצאת בדיקה", "קוד קבוצה", "קוד מעבדה", "זמן דיגום", "מספר מעבדה/מספר דגימה", "סטטוס", "שם בודק" };

        public void CreateExcel(List<Aliquot> list, string[] columnStrings)
        {


            ExcelApp.Workbooks.Add();
            ExcelWorkSheet = ExcelApp.ActiveSheet;
            ExcelWorkSheet.DisplayRightToLeft = true;
            ExcelWorkSheet.DisplayRightToLeft = false;
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

                //Get reported result
                var reportedResult = (from item in aliquot.Tests
                                      from result in item.Results
                                      where result.REPORTED == "T"
                                      select result.FormattedResult).FirstOrDefault();


                ExcelWorkSheet.Cells[row, col++] = sdg.U_SAMPLING_SITE;//קוד נקודה
                ExcelWorkSheet.Cells[row, col++] = aliquot.TestTemplateEx.TestCode;
                ExcelWorkSheet.Cells[row, col++] = sample.SampledOn;
                ExcelWorkSheet.Cells[row, col++] = reportedResult;
                ExcelWorkSheet.Cells[row, col++] = "4";
                ExcelWorkSheet.Cells[row, col++] = "604";
                ExcelWorkSheet.Cells[row, col++] = sdg.SamplingTime;
                ExcelWorkSheet.Cells[row, col++] = aliquot.SampleId;
                ExcelWorkSheet.Cells[row, col++] = "Done";
                ExcelWorkSheet.Cells[row, col] = sdg.SampledByOperator.Name;

            }
            Range c1 = ExcelWorkSheet.Cells[1, 1];
            Range c2 = ExcelWorkSheet.Cells[row, columnStrings.Length];
            var oRange = ExcelWorkSheet.get_Range(c1, c2);
            oRange.EntireColumn.ColumnWidth = 20;
        }


        public void SaveFile()
        {
            string wn = "KASEFET " + Sdg.LabInfo.Name;
            var ph = dal.GetPhraseByName("Location folders");
            var pe = ph.PhraseEntries.FirstOrDefault(p => p.PhraseDescription == "Safe (KASEFET)");


            ExcelWorkSheet.SaveAs(pe.PhraseName + wn + "-" + DateTime.Now.ToString("yyyyMMddHHmmss"));

            ExcelApp.Quit();


        }
    }
}

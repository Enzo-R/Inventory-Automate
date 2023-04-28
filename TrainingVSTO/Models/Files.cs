using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TrainingVSTO.Models
{
    public class Files
    {
        //classe responsavel por manipular arquivos e criar intancias do excel

        public static void OpenM7Model()
        {
            Application excelApp = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.getActiveApp();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(Models.Excel.PathToM7D);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        }

        public static void CreateM7D(string day)
        {
            string FileName = "inventárioDia" + day;

            // Use the current instence of Excel and open selected workbook
            Workbooks.SheetSelect("M7", Models.Excel.PathM7C);

            // Put the M7 data to a new file model
            Globals.ThisAddIn.getActiveWorksheet().Range["B5 : K20000"].Value = Excel.Data;

            //Globals.Worksheet.PrintPreview();
            //workbook.SaveAs(FileName);
            //workbook.SaveCopyAs(FileName);
            //var _ = worksheet.ExportAsFixedFormat().Columns[2, 10];
            Workbooks.releaseObject(Globals.ThisAddIn.getActiveWorksheet());
        }
    }
}

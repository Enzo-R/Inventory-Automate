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
            Workbook workbook = excelApp.Workbooks.Open(Models.Excel.PathToM7DModel);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        }

        public static void CreateM7D(string day)
        {
            string File = "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\M7 - STK " + day + ".xla";
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Workbook wb = Globals.ThisAddIn.getActiveWorkbook();


            // Use the current instence of Excel and open selected workbook
            Workbooks.SheetSelect("M7", Models.Excel.PathToM7Open);

            // Put the M7 data to a new file model
            currentSheet.Range["B5 : K20000"].Value = Excel.Data;
            currentSheet.Columns.AutoFit();

            //Formulas to get the final data

            //End
            wb.SaveAs(File);
            Workbooks.releaseObject(currentSheet);
        }
    }
}

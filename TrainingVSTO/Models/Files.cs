using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace TrainingVSTO.Models
{
    public class Files
    {
        //classe responsavel por manipular arquivos e criar intancias do excel

        public static void OpenM7Model()
        {
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(Models.Excel.PathToM7DModel);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        }

        public static void CreateM7D()
        {
            // Use the current instence of Excel and open selected workbook
            Workbooks.SheetSelect("M7", Models.Excel.PathToM7DOpen);

            string date = Models.Excel.date.ToString();
            date.Replace("/", ".");

            // Variables
            string PathToServer = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2022\Teste\M7 - STK " + date + ".xlsx";
            Workbook workbook = Globals.ThisAddIn.getActiveWorkbook();
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            // Put the M7 data to a new file model
            currentSheet.Range["A4"].PasteSpecial(XlPasteType.xlPasteAll);
            Models.Workbooks.UpFormulas();
            currentSheet.Columns.AutoFit();


            //End
            if (currentSheet.Cells != null)
            {
                try
                {
                    workbook.SaveAs(PathToServer);
                }
                catch (Exception)
                {
                    workbook
                    .SaveAs(@"C:\Users\Enzo\OneDrive\Área de Trabalho\Joyson\M7 - STK " + date + ".xlsx");
                }
                Workbooks.ReleaseObject(currentSheet);
                Clipboard.Clear();
            }

        }
    }
}

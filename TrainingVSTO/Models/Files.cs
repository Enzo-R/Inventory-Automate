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
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(Excel.PathToM7DOpen);
            Worksheet Sheet = workbook.Sheets["M7"];



            // Variables
            string date = Excel.date.ToString("d");
            string dateValidate = date.Replace("/", ".");
            string PathToServer = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2022\Teste"
                                                                        + @"\M7 - STK " + dateValidate + " -.xlsx";
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();


            // Put the M7 data to a new file model
            currentSheet.Range["A4"].PasteSpecial(XlPasteType.xlPasteAll);
            Workbooks.M7Formulas();
            currentSheet.Columns.AutoFit();

            //Create Power Pivot
            Workbooks.DynimicTable();


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
                    .SaveAs(@"C:\Users\EROLIVEIRA\OneDrive - Joyson Group\Área de Trabalho\Joyson\M7 - STK " + dateValidate + " -.xlsx");
                }
                finally
                {
                    Clipboard.Clear();
                    Workbooks.ReleaseObject(currentSheet);
                    Workbooks.ReleaseObject(excelApp);
                }
            }

        }

        public static void OpenNoDispSTK()
        {
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(Excel.PathToM7DOpen);
            Worksheet Sheet = workbook.Sheets["M7"];
        }
    }
}

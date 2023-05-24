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
            Sheet.Activate();
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();


            // Put the M7 data to a new file model
            currentSheet.Range["A4"].PasteSpecial(XlPasteType.xlPasteAll);
            Workbooks.M7Formulas();
            currentSheet.Columns.AutoFit();

            //Create Power Pivot
            Workbooks.DynimicTable();

            Clipboard.Clear();

        }

        public static void OpenNoDispSTK(string path, string sheet)
        {
            //Open file
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet currentSheet = workbook.Sheets[sheet];
            currentSheet.Activate();

            //Manipulating objects
            currentSheet.Columns["D:E"].Delete();
            Workbooks.Data(sheet, "A2:I2");
            workbook.Close(false);
            
            

            //Generate STK
            Workbooks.NoDisponible_();

            //End
            if (currentSheet.Cells != null)
            {
                try
                {
                    workbook.SaveAs(Excel.PathToServer);
                }
                catch (Exception)
                {
                    workbook
                    .SaveAs(@"C:\Users\EROLIVEIRA\OneDrive - Joyson Group\Área de Trabalho\Joyson\M7 - STK " + Excel.dateValidate + " -.xlsx");
                }
                finally
                {
                    Clipboard.Clear();
                    Workbooks.ReleaseObject(currentSheet);
                    Workbooks.ReleaseObject(excelApp);
                }
            }
        }
    }
}

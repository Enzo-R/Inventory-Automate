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
using Microsoft.Office.Tools.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TrainingVSTO.Models
{
    public class Files
    {
        //classe responsavel por manipular arquivos e criar intancias do excel

        public static Workbook OpenM7Model()
        {
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(Models.Excel.PathToM7DModel);
            return workbook;
        }

        public static void CreateM7D()
        {
            // Use the current instence of Excel and open selected workbook
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(Excel.PathToM7DOpen);
            Worksheet Sheet = workbook.Sheets["M7"];
            Sheet.Activate();


            // Put the M7 data to a new file model
            Sheet.Range["A4"].PasteSpecial(XlPasteType.xlPasteAll);
            Workbooks.M7Formulas();
            Sheet.Columns.AutoFit();

            //Create Power Pivot
            Workbooks.DynimicTable();

            Clipboard.Clear();
            Finals(Globals.ThisAddIn.getActiveWorkbook());

        }

        public static void OpenNoDispSTK(string path, string sheet)
        {
            //Open file
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet selectSheet = workbook.Sheets[sheet];
            selectSheet.Activate();

            //Manipulating objects
            selectSheet.Columns["D:E"].Delete();
            Workbooks.GetData(sheet, "A2:I2");

            //Generate STK
            workbook.Close(false);
            //Dialogs dialogs = Globals.ThisAddIn.getActiveWorkbook().Dialogs;

            //// Fechar todas as caixas de diálogo abertas
            //foreach (Dialog dialog in dialogs)
            //{
            //    dialog.Cancel();
            //    dialog.Dispose();
            //}
            Workbooks.SetData("A4", sheet);
            Workbooks.NoDispProcess();

            Finals(Globals.ThisAddIn.getActiveWorkbook());

        }

        public static void OpenFG(string path, string sheet)
        {
            //Open file
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet selectSheet = workbook.Sheets[sheet];
            selectSheet.Activate();

            //Manipulating objects
            Workbooks.GetData(sheet, "A2:I2");
            workbook.Close(false);
            Workbooks.SetData("A2", sheet);
            Workbooks.FG_expedicao();
            Finals(Globals.ThisAddIn.getActiveWorkbook());

        }

        public static void Finals(Workbook wb)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            //End
            if (currentSheet.Cells != null)
            {
                try
                {
                    wb.SaveAs(Excel.PathToServer);
                }
                catch (Exception)
                {
                    wb
                    .SaveAs(@"C:\Inventario\M7 - STK " + Excel.dateValidate + " -.xlsx");
                }
                finally
                {
                    Clipboard.Clear();
                    Workbooks.ReleaseObject(currentSheet);
                }
            }
        }
    }
}

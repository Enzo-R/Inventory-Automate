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

        public static void OpenM7()
        {
            string path = "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\AbreModelo7 - Rev1 - Copia.xlsm";
            Application excelApp = (Excel.Application)Globals.ThisAddIn.getActiveApp();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

        }

        public static void CreateM7D(string day)
        {
            string path = "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\modelo.xlsx";
            string FileName = "inventárioDia" + day;
            string m7 = "M7 EF";

            // Create a new instence of Excel and open selected workbook
            Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            // Put the M7 data to a new file model
            



            //Globals.Worksheet.PrintPreview();
            //workbook.SaveAs(FileName);
            //workbook.SaveCopyAs(FileName);
            //var _ = worksheet.ExportAsFixedFormat().Columns[2, 10];
        }
    }
}

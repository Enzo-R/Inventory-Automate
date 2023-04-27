using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TrainingVSTO
{
    public class Files
    {
        //classe responsavel por manipular arquivos e criar intancias do excel

        public static void CreateM7D()
        {
            string path = "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\modelo.xlsx";

            // Create a new instence of Excel
            Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            // Open a new workbook and sheets
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            
            //worksheet.Name = "M7";

            //Globals.Worksheet.PrintPreview();

            //wb.Save();
            //var _ = worksheet.ExportAsFixedFormat().Columns[2, 10];
        }
    }
}

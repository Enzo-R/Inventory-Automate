using System;
using System.Collections.Generic;
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
            // Create a new instence of Excel
            Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            // Create a new workboob and sheets
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "M7";

            //wb.Save();
            //var _ = worksheet.ExportAsFixedFormat().Columns[2, 10];
        }
    }
}

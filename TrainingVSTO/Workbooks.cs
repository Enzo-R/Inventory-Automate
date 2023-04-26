using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TrainingVSTO
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar objetos e intancias do excel

        public static void CriandoM7Diario()
        {
            // Criar uma nova instância do Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Workbook wb = new Workbook();
            //wb.Save();
            Worksheet worksheet = new Worksheet();
            worksheet.Name = "M7 Dayli";

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();




            //Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            //var _ = currentSheet.ExportAsFixedFormat().Columns[2, 10];



        }
        public static void clearWorksheet()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            string text = currentSheet.Cells.Text.ToString();

            if (text != string.Empty)
            {
                currentSheet.Columns.Rows.Clear();
            }
        }

        public static void ReadAndWriteArq(string path)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            var arq = File.ReadAllText(path);

            try
            {
                var cell = currentSheet.Range["A1"].Value2 = arq;
                cell.Columns.AutoFit();

            }
            catch (Exception ex)
            {

               MessageBox.Show(ex.Message);
            }
        }
    }
}

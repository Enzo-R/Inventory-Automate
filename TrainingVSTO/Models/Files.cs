using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            string File = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2022\Teste\M7 - STK " + day + ".xlsx";
            Range cols = currentSheet.Range["A4 : J4"];

            currentSheet.Cells[Excel.Data].Value = Excel.Data;
            //Models.Workbooks.VLookUp();
            currentSheet.Columns.AutoFit();
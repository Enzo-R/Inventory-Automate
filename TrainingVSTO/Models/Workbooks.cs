using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using TrainingVSTO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TrainingVSTO.Models
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar elementos dentro do Excel

        public static Worksheet SheetSelect(string sheet, string path)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet Sheet = workbook.Sheets[sheet];
            return Sheet;
        }
        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ocorreu um erro ao liberar o objeto do Excel: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        public static void ReadAndWriteArq(string path)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            var content = File.ReadAllText(path);
            Clipboard.SetText(content);

            Range col = currentSheet.Range["A:A"];
            col.PasteSpecial(XlPasteType.xlPasteAll);
            if (col.Value != null)
            {
                Clipboard.Clear();
            }
        }


        public static object Data(string sheet)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[sheet];
            Range cells = currentSheet.Range["B5 : K5"];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];

            return select.Value;
        }


        public static void VLookUp()
        {
            //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range f1;

            f1 = currentSheet.Range["K4:K14598"];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";

            FilterData();
        }


        public static void FilterData()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range dataRange = currentSheet.Range["$A$3:$P$15000"];
            Range colK = currentSheet.Range["K:K"];


            string[] filterCriteria = new string[] {
                "BENS CAPITAL EM PROCESSO",
                "DISPOSITIVOS TAKATA",
                "FERRAMENTAL PARA VENDA",
                "INSUMO MANUT.MAQ.EQUIP.",
                "INSUMOS AUTOMACAO",
                "INSUMOS FERRAMENTARIA",
                "MANUT.MAQ.PROD.FERRAM",
                "MAQUINAS PARA VENDA",
                "MAQUINAS TAKATA",
                "MATERIA PRIMA",
                "MATERIAIS AUXILIARES",
                "MATERIAIS EPI",
                "MATERIAIS PARA TESTE",
                "MATERIAL CONSTR CIVIL",
                "MATERIAL DE LIMPEZA",
                "MATERIAL ESCRITORIO",
                "MATERIAL INERTE",
                "MATERIAL INFORMATICA",
                "MATERIAL QUIMICO",
                "MERCADORIAS EM TRANSITO",
                "MERCADORIAS PARA REVENDA",
                "MOVEIS E UTENSILIOS",
                "PRODUTOS ACABADOS",
                "PRODUTOS SEMIACABADOS",
                "SUBPRODUTO",
                "USO CONSUMO MAQ.EQUIP."
            };

            if (colK.AutoFilter(11, "#N/D"))
            {
                dataRange.AutoFilter(4, filterCriteria, XlAutoFilterOperator.xlFilterValues);
                currentSheet.Range["$A$4:$P$15000"].SpecialCells(XlCellType.xlCellTypeVisible).Clear();
            }
            if (dataRange.AutoFilter(3, "TRM"))
            {
                //currentSheet.Range["$C$4:$P$15000"].
            }


            //if (string.IsNullOrEmpty())
            //{
            //    // Célula está vazia
            //    // Faça o que for necessário para tratar a célula vazia
            //}
            //if(dataRange.AutoFilter()) delete.rows

        }


        public static void GetCellsToSelect(String cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range cellSelect = currentSheet.Range[cell];
            Range sl = currentSheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
        }
    }
}

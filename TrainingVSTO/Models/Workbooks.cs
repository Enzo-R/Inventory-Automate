using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;

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
            currentSheet.Range["A:A"].PasteSpecial(XlPasteType.xlPasteAll);

        }
        public static object Data(string sheet)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[sheet];
            object dados = currentSheet.Range["B5 : K20000"].Value;
            return dados;
        }
        public static void VLookUp()
        {
            //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range f1;
            //forma de obter o tamanho das colunas.
            f1 = currentSheet.Range["K4:K14598"];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";

            FilterData();


        }
        public static void FilterData()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];

            string[] filterCriteria = new string[] {
                "BENS CAPITAL EM PROCESSO",
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
                "MERCADORIAS EM TRANSITO",
                "MERCADORIAS PARA REVENDA",
                "MOVEIS E UTENSILIOS",
                "PRODUTOS ACABADOS",
                "PRODUTOS SEMIACABADOS",
                "SUBPRODUTO",
                "USO CONSUMO MAQ.EQUIP."
            };
            Range dataRange = currentSheet.Range["$A$3:$P$14619"];

            if (currentSheet.Range["K:K"].AutoFilter(11, "#N/D"))
            {
                dataRange.AutoFilter(4, filterCriteria, XlAutoFilterOperator.xlFilterValues);
                currentSheet.Range["$A$4:$P$14619"].SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            }


        }
    }
}

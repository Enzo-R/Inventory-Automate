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


        public static void Data(string sheet)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[sheet];
            Range cells = currentSheet.Range["B5 : K5"];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];
            select.Copy();
        }


        public static void UpFormulas()
        {
            //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];

            Range range = GetCellsToSelect("B4");
            int rows = range.SpecialCells(XlCellType.xlCellTypeVisible).Count + 3;

            Range f1 = currentSheet.Range["K4:K" + rows];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";




            FilterData();

            Range f2 = GetCellsToSelect("M4:M" + rows);
            f2.Formula = @"=J4/$I$1";

            Range f3 = GetCellsToSelect("O4:O" + rows);
            f3.Formula = @"=J4/5,0758";


        }


        public static void FilterData()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range k3 = GetCellsToSelect("K3");

            string[] filterCriteria1 = new string[] {
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
            string[] filterCriteria2 = new string[]
            {
                "EMBALAGENS RETORNAVEIS",
                "MATERIAL TERCEIRO",
                "PRODUTOS EM ELABORACAO"
            };
            string[] filterCriteria3 = new string[]
            {
                "SUBPROD",
                "TERC",
                "="
            };

            //filtragem por classificação e descrição.
            if (k3.AutoFilter(11, "#N/D"))
            {
                Range d3 = GetCellsToSelect("D3");
                d3.AutoFilter(4, filterCriteria1, XlAutoFilterOperator.xlFilterValues);

                Range data = GetCellsToSelect("A4:K4");
                data.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

                d3.AutoFilter(4, filterCriteria2, XlAutoFilterOperator.xlFilterValues);

                Range k4 = GetCellsToSelect("K4");
                k4.Value = "Raw Material";
            }
            //deletando linhas vazias em classificação e descrição
            if (k3.AutoFilter(11, "="))
            {
                Range d3 = GetCellsToSelect("D3");
                d3.AutoFilter(4, "=");
                d3.SpecialCells(XlCellType.xlCellTypeBlanks).EntireRow.Delete();

            }

            currentSheet.AutoFilterMode = false;

            try
            {
                _ = currentSheet.EnableAutoFilter;
                currentSheet.Rows["3:3"].AutoFilter();
            }
            catch (Exception)
            {
                MessageBox.Show("Filtro não foi ativado");
            }

            //filtragem por subconta
            Range range = GetCellsToSelect("B4");
            int alt = range.Count+3;
            Range c4 = GetCellsToSelect("C4:C" + alt);
            c4.AutoFilter(3, filterCriteria3, XlAutoFilterOperator.xlFilterValues);
            c4.Value = "SW";

            c4.AutoFilter(3, "TRM");
            c4.Value= "ISS";

        }


        public static Range GetCellsToSelect(String cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range cellSelect = currentSheet.Range[cell];
            Range sl = currentSheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
            return sl;
        }
    }
}

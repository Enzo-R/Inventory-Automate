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

            FilterDataToM7();

            Range range2 = GetCellsToSelect("K4");
            int rowsCount = range2.Count + 3;

            Range f2 = currentSheet.Range["M4:M" + rowsCount];
            f2.Formula = @"=J4/$I$1";

            Range f3 = currentSheet.Range["O4:O" + rowsCount];
            f3.Formula = @"=J4/5.0758";


            currentSheet.Range["J2"].Formula = @"=SUBTOTAL(9,J4:J" + rowsCount + ")";
            currentSheet.Range["M2"].Formula = @"=SUBTOTAL(9,M4:J" + rowsCount + ")";
            currentSheet.Range["O2"].Formula = @"=SUBTOTAL(9,O4:J" + rowsCount + ")";

            //FilterDataToClient();

        }


        public static void FilterDataToM7()
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

            disableFilter();

            //filtragem por subconta
            Range range = GetCellsToSelect("B4");
            int alt = range.Count + 3;
            Range c4 = GetCellsToSelect("C4:C" + alt);
            c4.AutoFilter(3, filterCriteria3, XlAutoFilterOperator.xlFilterValues);
            c4.Value = "SW";

            c4.AutoFilter(3, "TRM");
            c4.Value = "ISS";

            disableFilter();

        }


        public static Range GetCellsToSelect(String cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range cellSelect = currentSheet.Range[cell];
            Range sl = currentSheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
            return sl;
        }


        public static void FilterDataToClient()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range f4 = GetCellsToSelect("F4");
            int i = f4.Count + 3;

            Range l4 =currentSheet.Range["L4:L"+i];
            Range n4 = currentSheet.Range["N4:N"+i];
            n4.Formula = @"=VLOOKUP(L4, Clientes!A:B,2,0)";


            if (f4.AutoFilter(6, "*TOY*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Toyota";
            }

            if (f4.AutoFilter(6, "*MAN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "MAN";
            }

            if (f4.AutoFilter(6, "*FCA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FCA";
            }

            if (f4.AutoFilter(6, "*FIAT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FIAT";
            }

            if (f4.AutoFilter(6, "*VW*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "VW";
            }

            if (f4.AutoFilter(6, "*Corsa*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "GM";
            }

            if (f4.AutoFilter(6, "*Nissa*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "NISSAN";
            }

            if (f4.AutoFilter(6, "*REN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "RENAULT";
            }

            if (f4.AutoFilter(6, "*HONDA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }

            if (f4.AutoFilter(6, "*MITSUB*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HPE";
            }

            if (f4.AutoFilter(6, "*RENAUL*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "RENAULT";
            }

        }


        public static void disableFilter()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
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
        }
    }
}

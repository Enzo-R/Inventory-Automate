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
using System.Net.Http;

namespace TrainingVSTO.Models
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar elementos dentro do Excel
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


        public static void Data(string sheet, string range)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[sheet];
            Range cells = currentSheet.Range[range];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];
            select.Copy();
        }


        public static void M7Formulas()
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

            int newCount = rowsCount - 3;
            currentSheet.Range["J2"].Formula = @"=SUBTOTAL(9,J4:J" + newCount + ")";
            currentSheet.Range["M2"].Formula = @"=SUBTOTAL(9,M4:M" + newCount + ")";
            currentSheet.Range["O2"].Formula = @"=SUBTOTAL(9,O4:O" + newCount + ")";

            FilterDataToClient();

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

            refreshFilter();

            //filtragem por subconta
            Range range = GetCellsToSelect("B4");
            int alt = range.Count + 3;
            Range c4 = GetCellsToSelect("C4:C" + alt);
            c4.AutoFilter(3, filterCriteria3, XlAutoFilterOperator.xlFilterValues);
            c4.Value = "SW";

            c4.AutoFilter(3, "TRM");
            c4.Value = "ISS";

            refreshFilter();

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

            Range l4 = currentSheet.Range["L4:L" + i];
            Range n4 = currentSheet.Range["N4:N" + i];
            n4.Formula = @"=VLOOKUP(L4, Clientes!A:B,2,0)";

            if (f4.AutoFilter(6, "*GM*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "GM";
            }

            if (f4.AutoFilter(6, "*PSA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "PSA";
                if (f4.AutoFilter(6, "*PEUGEO*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
                {
                    l4.Value = "PSA";
                }
            }

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
            if (f4.AutoFilter(6, "*FI AT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FIAT";
                if (f4.AutoFilter(6, "*F IAT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
                {
                    l4.Value = "FIAT";
                }
            }

            if (f4.AutoFilter(6, "*VW*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "VW";
            }
            if (f4.AutoFilter(6, "*V W*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "VW";
            }
            if (f4.AutoFilter(6, "*FOX*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "VW";
            }

            if (f4.AutoFilter(6, "*Corsa*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "GM";
            }

            if (f4.AutoFilter(6, "*Niss*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "NISSAN";
            }

            if (f4.AutoFilter(6, "*REN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "RENAULT";
            }

            if (f4.AutoFilter(6, "*HON*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }

            if (f4.AutoFilter(6, "*HYU*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Hyundai";

                if (f4.AutoFilter(6, "*HY UNDAI*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
                {
                    l4.Value = "Hyundai";

                }
            }


            if (f4.AutoFilter(6, "*MIT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HPE";
            }
            if (f4.AutoFilter(6, "*MITSUB*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HPE";
            }

            if (f4.AutoFilter(6, "*RENA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "RENAULT";
            }

            if (f4.AutoFilter(6, "*SCANI*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "SCANIA";
            }

            if (f4.AutoFilter(6, "*FORD*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FORD";
            }

            if (f4.AutoFilter(6, "*Faure*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Faurencia";
            }

            if (f4.AutoFilter(6, "*STELLA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "STELLANTIS";
                if (f4.AutoFilter(6, "*CIVI*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
                {
                    l4.Value = "STELLANTIS";
                }
            }

            if (f4.AutoFilter(6, "*COROLL*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Toyota";
            }

            refreshFilter();

        }


        public static void refreshFilter()
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


        public static void DynimicTable()
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range all = currentSheet.Range[GetCellsToSelect("A3"), GetCellsToSelect("A3").End[XlDirection.xlToRight]];

            Workbook workbook = Globals.ThisAddIn.getActiveWorkbook();
            Worksheet newSheet = workbook.Sheets.Add();
            newSheet.Name = "Pivot Table";

            //Get data for Pivot tabel
            PivotCache oPivotCache = workbook.PivotCaches().Add(XlPivotTableSourceType.xlDatabase, all);

            //Create Pivot table
            PivotCaches pch = workbook.PivotCaches();
            pch.Add(XlPivotTableSourceType.xlDatabase, all)
               .CreatePivotTable(newSheet.Cells[3, 1], "Pivot Table 1", Type.Missing, Type.Missing);

            //Manipulate pivot table object "pvt"
            PivotTable pvt = newSheet.PivotTables("Pivot Table 1");
            pvt.ShowDrillIndicators = true;

            //set filds for pivot table
            PivotField fld = pvt.PivotFields("Subconta");
            fld.Orientation = XlPivotFieldOrientation.xlRowField;

            fld = pvt.PivotFields("Quantidade");
            fld.Orientation = XlPivotFieldOrientation.xlDataField;
            fld.Position = 1;
            fld.NumberFormat = "#,##0";

            fld = pvt.PivotFields("Classificação");
            fld.Orientation = XlPivotFieldOrientation.xlColumnField;

            fld = pvt.PivotFields("Total USD");
            fld.Orientation = XlPivotFieldOrientation.xlDataField;
            fld.Position = 2;
            fld.NumberFormat = "#,##0.00";

            pvt.DataPivotField.Orientation = XlPivotFieldOrientation.xlColumnField;

            Range collor1 = GetCellsToSelect("B4:C4");
            collor1.Interior.Color = System.Drawing.Color.Beige;

            Range collor2 = GetCellsToSelect("D4:E10");
            collor2.Interior.Color = System.Drawing.Color.LightGoldenrodYellow;

            Range collor3 = GetCellsToSelect("F4:G4");
            collor3.Interior.Color = System.Drawing.Color.Bisque;

            Range collor4 = GetCellsToSelect("H4:I4");
            collor4.Interior.Color = System.Drawing.Color.LightYellow;

            Range collor5 = newSheet.Range["J4:K10"];
            collor5.Interior.Color = System.Drawing.Color.LightSalmon;

            //laterais
            Range collor6 = newSheet.Range["A5:A10"];
            collor6.Interior.Color = System.Drawing.Color.LightSteelBlue;
            collor6.Cells.Font.Bold = true;

            Range collor7 = newSheet.Range["A10:K10"];
            collor7.Interior.Color = System.Drawing.Color.LightSteelBlue;
            collor7.Cells.Font.Bold = true;

            Range collor8 = newSheet.Range["B3:K3"];
            collor8.Interior.Color = System.Drawing.Color.LightSteelBlue;
            collor8.Cells.Font.Bold = true;


            newSheet.Columns.AutoFit();

        }


        public static void NoDisponible_()
        {
            Worksheet noDisponible = Globals.ThisAddIn.getActiveWorkbook().Sheets["No Disponible"];
            noDisponible.Activate();
            Range init = noDisponible.Range["A4"];

            init.PasteSpecial(XlPasteType.xlPasteAll);

            Range columnRange = GetCellsToSelect("B4").NumberFormat = "0";

            if (init.Value != null)
            {
                Clipboard.Clear();
            }
        }


        public static void NoDispFormulas()
        {
            Worksheet noDisponible = Globals.ThisAddIn.getActiveWorkbook().Sheets["No Disponible"];
            noDisponible.Activate();
            Range range = GetCellsToSelect("B4");
            int rows = range.Count + 3;

            //Custo Init
            noDisponible.Range["J4:J" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:I,9,0)";

            //Custo Total
            noDisponible.Range["K4:K" + rows].Formula = @"=J4*E4";

            //Segment
            noDisponible.Range["L4:L" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:C,3,0)";

            //Classification
            noDisponible.Range["M4:M" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:K,11,0)";

            //Amount USD
            noDisponible.Range["P4:P" + rows].Formula = @"=K4/$J$1";

            //get USD
            noDisponible.Range["J1"].Formula = @"='M7'!I1";

            //subtotal
            noDisponible.Range["K2"].Formula = @"=SUBTOTAL(9,K4:K" + rows + ")";



            Range m4 = GetCellsToSelect("M4");

            if(m4.AutoFilter(13, "#N/D"))
            {
                Range all = GetCellsToSelect("A4:S4");
                all.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();
        }



    }
}

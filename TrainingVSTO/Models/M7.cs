using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TrainingVSTO.Models
{
    public class M7
    {
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


        public static void M7Formulas()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];

            Range range = Utils.GetCellsToSelect("B4");
            int rows = range.SpecialCells(XlCellType.xlCellTypeVisible).Count + 3;

            Range f1 = currentSheet.Range["K4:K" + rows];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";

            FilterDataToM7();

            Range range2 = Utils.GetCellsToSelect("K4");
            int rowsCount = range2.Count + 3;

            Range f2 = currentSheet.Range["M4:M" + rowsCount];
            f2.Formula = @"=J4/$I$1";

            Range f3 = currentSheet.Range["O4:O" + rowsCount];
            f3.Formula = @"=J4/5.0758";


            currentSheet.Range["J2"].Formula = @"=SUBTOTAL(9,J4:J" + rowsCount + ")";
            currentSheet.Range["M2"].Formula = @"=SUBTOTAL(9,M4:M" + rowsCount + ")";
            currentSheet.Range["O2"].Formula = @"=SUBTOTAL(9,O4:O" + rowsCount + ")";

            FilterDataToClient();

            //Variação.


            //Concatenar colunas e adicionar novas.
            Range AnewC = currentSheet.Columns[1];
            AnewC.Insert();
            currentSheet.Range["A3:A" + rowsCount].Formula = "=CONCAT(B3,C3)";

            currentSheet.Range["S2"].Formula = @"=SUBTOTAL(9,S4:S" + rowsCount + ")";

            currentSheet.Range["T2"].Formula = @"=SUBTOTAL(9,T4:S" + rowsCount + ")";

            currentSheet.Range["U2"].Formula = @"=SUBTOTAL(9,U4:S" + rowsCount + ")";

            currentSheet.Range["V2"].Formula = @"=SUBTOTAL(9,V4:S" + rowsCount + ")";

            //ranges
            Range H4 = currentSheet.Range["H4:H" + rows];
            Range N4 = currentSheet.Range["N4:N" + rows];
            Range Q4 = currentSheet.Range["Q4:Q" + rows]; Q4.Style = "Percent";
            Range R4 = currentSheet.Range["R4:R" + rows];
            Range S4 = currentSheet.Range["S4:S" + rows];
            Range T4 = currentSheet.Range["T4:T" + rows];
            Range U4 = currentSheet.Range["U4:U" + rows];
            Range V4 = currentSheet.Range["V4:V" + rows];

            //Inserindo forumlas
            Utils.VlookUp("M7", -1, Q4, @"=(H4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0))/VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0)");

            Utils.VlookUp("M7", -1, R4, @"=H4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0)");

            Utils.VlookUp("M7", -1, S4, @"=N4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$N,14,0)");

            Utils.VlookUp("M7", -7, T4, @"=N4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$N,14,0)");

            Utils.VlookUp("M7", -15, U4, @"=N4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$N,14,0)");

            Utils.VlookUp("M7", -30, V4, @"=N4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$N,14,0)");

            //filtrando nullos
            Q4.AutoFilter(17, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            Q4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();
            Utils.refreshFilter();

            R4.AutoFilter(18, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            R4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            Utils.refreshFilter();

            S4.AutoFilter(19, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            S4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            Utils.refreshFilter();

            T4.AutoFilter(20, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            T4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            Utils.refreshFilter();

            U4.AutoFilter(21, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            U4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            Utils.refreshFilter();

            V4.AutoFilter(22, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            V4.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            Utils.refreshFilter();

            //number format
            S4.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* " + "-" + "??_ ;_-@_ ";
            T4.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* " + "-" + "??_ ;_-@_ ";
            U4.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* " + "-" + "??_ ;_-@_ ";
            V4.NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* " + "-" + "??_ ;_-@_ ";

        }


        public static void FilterDataToM7()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range k3 = Utils.GetCellsToSelect("K3");

            #region Lists to Filter
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
            #endregion

            //filtragem por classificação e descrição.
            if (k3.AutoFilter(11, "#N/D"))
            {
                Range d3 =  Utils.GetCellsToSelect("D3");
                d3.AutoFilter(4, filterCriteria1, XlAutoFilterOperator.xlFilterValues);

                Range data = Utils.GetCellsToSelect("A4:K4");
                data.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

                d3.AutoFilter(4, filterCriteria2, XlAutoFilterOperator.xlFilterValues);

                Utils.GetCellsToSelect("K4").Value = "Raw Material";
            }
            //deletando linhas vazias em classificação e descrição
            if (k3.AutoFilter(11, "="))
            {
                Range d3 = Utils.GetCellsToSelect("D3");
                d3.AutoFilter(4, "=");
                d3.SpecialCells(XlCellType.xlCellTypeBlanks).EntireRow.Delete();

            }

            Utils.refreshFilter();

            //filtragem por subconta
            Range range = Utils.GetCellsToSelect("B4");
            int alt = range.Count + 3;
            Range c4 = Utils.GetCellsToSelect("C4:C" + alt);

            c4.AutoFilter(3, filterCriteria3, XlAutoFilterOperator.xlFilterValues);
            c4.Value = "SW";

            c4.AutoFilter(3, "TRM");
            c4.Value = "ISS";

            Utils.refreshFilter();

        }


        public static void FilterDataToClient()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range f4 = Utils.GetCellsToSelect("F4");
            int i = f4.Count + 3;

            Range l4 = currentSheet.Range["L4:L" + i];
            Range n4 = currentSheet.Range["N4:N" + i];


            #region Filters to Client

            if (f4.AutoFilter(6, "*MAN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "MAN";
            }
            if (f4.AutoFilter(6, "*COMAN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "";
            }
            if (f4.AutoFilter(6, "*FIA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FIAT";
            }
            if (f4.AutoFilter(6, "*GM*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "GM";
            }
            if (f4.AutoFilter(6, "*REN*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "RENAULT";
            }
            if (f4.AutoFilter(6, "*HON*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }
            if (f4.AutoFilter(6, "*PSA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "PSA";
            }
            if (f4.AutoFilter(6, "*TOY*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Toyota";
            }
            if (f4.AutoFilter(6, "*FCA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FCA";
            }
            if (f4.AutoFilter(6, "*FI AT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FIAT";
            }
            if (f4.AutoFilter(6, "*FIA T*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "FIAT";
            }
            if (f4.AutoFilter(6, "*VW*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "VW";
            }
            if (f4.AutoFilter(6, "*V W*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
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

            if (f4.AutoFilter(6, "*HO NDA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }

            if (f4.AutoFilter(6, "*HONDA*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }
            if (f4.AutoFilter(6, "*HYUND*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "Hyundai";
            }
            if (f4.AutoFilter(6, "*MITSUB*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HPE";
            }
            if (f4.AutoFilter(6, "*RENAUL*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
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
            }
            if (f4.AutoFilter(6, "*CIVI*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "HONDA";
            }
            if (f4.AutoFilter(6, "*PEUGEOT*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "PEUGEOT";
            }
            if (f4.AutoFilter(6, "*PEUG*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "PEUGEOT";
            }
            if (f4.AutoFilter(6, "*COROL*", XlAutoFilterOperator.xlAnd, Type.Missing, true))
            {
                l4.Value = "TOYOTA";
            }

            Utils.refreshFilter();
            #endregion

            //procv nas planilhas para Clients
            if (l4.AutoFilter(12, "="))
            {

                Utils.PreviousDayProcv("M7", l4, @"=VLOOKUP(A4,'[M7 - STK 01.07.2023 -.xlsx]M7'!$A:$L,12,0)");

            }
            l4.AutoFilter(12, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            l4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Others";
            Utils.refreshFilter();


            //procv nas planilhas para CS
            Utils.PreviousDayProcv("M7", n4, @"=VLOOKUP(A4,'[M7 - STK 01.07.2023 -.xlsx]M7'!$B:$O,14,0)");
        }


        public static void DynimicTable()
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range all = currentSheet.Range[Utils.GetCellsToSelect("A3"), Utils.GetCellsToSelect("A3").End[XlDirection.xlToRight]];

            Workbook workbook = Globals.ThisAddIn.getActiveWorkbook();
            Worksheet newSheet = workbook.Sheets.Add();
            newSheet.Name = "M7 summary";

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

            Range collor1 = Utils.GetCellsToSelect("B4:C4");
            collor1.Interior.Color = System.Drawing.Color.Beige;

            Range collor2 = Utils.GetCellsToSelect("D4:E10");
            collor2.Interior.Color = System.Drawing.Color.LightGoldenrodYellow;

            Range collor3 = Utils.GetCellsToSelect("F4:G4");
            collor3.Interior.Color = System.Drawing.Color.Bisque;

            Range collor4 = Utils.GetCellsToSelect("H4:I4");
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


    }
}

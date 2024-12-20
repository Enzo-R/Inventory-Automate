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
using Microsoft.Office.Tools.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Net.NetworkInformation;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TrainingVSTO.Models
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar elementos dentro do Excel

        static string[] filterCriteriaNull = new string[]
        {
                "#N/D",
                "0",
                "="
        };

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

        public static void M7()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];

            Range range = GetCellsToSelect("B4");
            int rows = range.SpecialCells(XlCellType.xlCellTypeVisible).Count + 3;

            Range f1 = currentSheet.Range["K4:K" + rows];
            f1.Formula = @"=VLOOKUP(B4,'Base Referencias'!A:C,3,0)";

            FilterDataToM7();

            Range range2 = GetCellsToSelect("K4");
            int rowsCount = range2.Count + 3;

            Range m4 = currentSheet.Range["M4:M" + rowsCount];
            m4.Formula = @"=J4/$I$1";

            Range o4 = currentSheet.Range["O4:O" + rowsCount];
            o4.Formula = @"=J4/5.0758";


            currentSheet.Range["J2"].Formula = @"=SUBTOTAL(9,J4:J" + rowsCount + ")";
            currentSheet.Range["M2"].Formula = @"=SUBTOTAL(9,M4:M" + rowsCount + ")";
            currentSheet.Range["O2"].Formula = @"=SUBTOTAL(9,O4:O" + rowsCount + ")";

            FilterDataToClient();

            //Variação.
            Variation(currentSheet, rowsCount);

            currentSheet.Activate();

        }

        public static void FilterDataToM7()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range k3 = GetCellsToSelect("K3");

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
                Range d3 = GetCellsToSelect("D3");
                d3.AutoFilter(4, filterCriteria1, XlAutoFilterOperator.xlFilterValues);

                Range data = GetCellsToSelect("A4:K4");
                data.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();

                d3.AutoFilter(4, filterCriteria2, XlAutoFilterOperator.xlFilterValues);

                GetCellsToSelect("K4").Value = "Raw Material";
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

        public static void FilterDataToClient()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range f4 = GetCellsToSelect("F4");
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

            refreshFilter();
            #endregion

            //procv nas planilhas para Clients
            if (l4.AutoFilter(12, "="))
            {
                PreviousDayProcv("M7", l4, @"=VLOOKUP(A4,'[M7 - STK 01.08.2023 -.xlsx]M7'!$A:$L,12,0)");
            }
            refreshFilter();

            l4.AutoFilter(12, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            l4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Others";
            refreshFilter();


            //procv nas planilhas para CS
            PreviousDayProcv("M7", n4, @"=VLOOKUP(A4,'[M7 - STK 01.08.2023 -.xlsx]M7'!$B:$O,14,0)");

            if (n4.AutoFilter(14, "#N/D"))
            {
                Range visible = n4.SpecialCells(XlCellType.xlCellTypeVisible);
                Range firstCell = visible.Cells[1];
                string c = firstCell.Row.ToString();
                visible.Formula = "=VLOOKUP(L" + c + ",'Base Referencias'!E:F,2,0)";
            }
            refreshFilter();
        }

        public static void Variation(Worksheet currentSheet, int rowsCount)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            //Concatenar colunas e adicionar novas.
            Range AnewC = currentSheet.Columns[1];
            AnewC.Insert();
            currentSheet.Range["A3:A" + rowsCount].Formula = "=CONCAT(B3,C3)";

            currentSheet.Range["S2"].Formula = @"=SUBTOTAL(9,S4:S" + rowsCount + ")";

            //ranges
            Range Q4 = currentSheet.Range["Q4:Q" + rowsCount]; Q4.Style = "Percent";
            Range R4 = currentSheet.Range["R4:R" + rowsCount];
            Range S4 = currentSheet.Range["S4:S" + rowsCount]; S4.NumberFormat = @"#,##0.00_ ;[Red]-#,##0.00";

            //Inserindo forumlas
            VlookUp("M7", -1, Q4, @"=(H4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0))/VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0)");

            VlookUp("M7", -1, R4, @"=H4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$H,8,0)");

            VlookUp("M7", -1, S4, @"=N4-VLOOKUP(A4,'[M7 - STK 30.06.2023 -.xlsx]M7'!$A:$N,14,0)");

            //filtrando nullos
            Q4.AutoFilter(17, "#N/D");
            Q4.SpecialCells(XlCellType.xlCellTypeVisible).Value = 0;
            refreshFilter();
            Q4.AutoFilter(17, "#DIV/0!");
            Q4.SpecialCells(XlCellType.xlCellTypeVisible).Value = 0;
            refreshFilter();

            //aplicando segundo filtro para valores nullos
            R4.AutoFilter(18, "#N/D");
            
            Range visibleCells = R4.SpecialCells(XlCellType.xlCellTypeVisible);
            Range firstCell = visibleCells.Cells[1];
            string c = firstCell.Row.ToString();

            VlookUp("M7", -1, visibleCells, @"=H" + c + "-VLOOKUP(B" + c + ",'[M7 - STK 30.06.2023 -.xlsx]M7'!$B:$H,7,0)");
            R4.AutoFilter(18, "#N/D");
            visibleCells.Value = 0;
            refreshFilter();

            //1
            S4.AutoFilter(19, "#N/D");
            Range visibleCells1 = S4.SpecialCells(XlCellType.xlCellTypeVisible);
            Range firstCell1 = visibleCells1.Cells[1];
            string c1 = firstCell1.Row.ToString();

            VlookUp("M7", -1, visibleCells1, @"=N" + c1 + "-VLOOKUP(B" + c1 + ",'[M7 - STK 30.06.2023 -.xlsx]M7'!$B:$N,13,0)");
            S4.AutoFilter(19, "#N/D");
            visibleCells1.Value = 0;
            refreshFilter();

            //Segunda parte
            Range All1 = GetCellsToSelect("B4:S4");

            //iniciando objeto
            Workbook M7Pbix = Globals.ThisAddIn.getActiveApp().Workbooks.Open(Excel.PathToPbix, UpdateLinks: false);
            Worksheet M7 = M7Pbix.Sheets["M7 Dayli"];
            Worksheet M7V = M7Pbix.Sheets["M7 Variation"];

            //Passando os dados para m7
            M7.Activate();
            Range All2 = GetCellsToSelect("A2:R2");
            All2.EntireRow.Delete();
            All1.Copy();
            M7.Range["A2"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone);
            Clipboard.Clear();

            //passando para a temp
            Range All3 = GetCellsToSelect("A1:S1");
            All3.Copy();

            //tranformando os dados
            M7Pbix.Sheets.Add();
            Worksheet temp = M7Pbix.Sheets["Planilha1"];
            temp.Range["A1"].PasteSpecial(XlPasteType.xlPasteValues);
            Clipboard.Clear();
            temp.Range["B:B"].EntireColumn.Delete();
            temp.Range["C:C"].EntireColumn.Delete();
            temp.Range["C:C"].EntireColumn.Delete();
            temp.Range["E:E"].EntireColumn.Delete();
            temp.Range["E:E"].EntireColumn.Delete();
            temp.Range["E:E"].EntireColumn.Delete();
            temp.Range["H:H"].EntireColumn.Delete();
            temp.Range["H:H"].EntireColumn.Delete();
            temp.Range["H:H"].EntireColumn.Delete();

            //colocando a data
            Range a1 = GetCellsToSelect("A1");
            int tmpAllCount = a1.Count;
            Range j1 = temp.Range["J2:J" + tmpAllCount];
            j1.Formula = "=TODAY()";

            //aplicando o filtro
            _ = temp.EnableAutoFilter;
            temp.Rows["1:1"].AutoFilter();

            //apagando os sem valores
            Range i2 = GetCellsToSelect("I2");
            temp.Range["I:I"].NumberFormat = @"#,##0.00_ ;[Red]-#,##0.00 ";
            i2.AutoFilter(9, "0,00");
            i2.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();

            refreshFilter();

            i2.Sort(i2.Columns[1], XlSortOrder.xlDescending, Type.Missing, Type.Missing,
                    XlSortOrder.xlAscending, Type.Missing, XlSortOrder.xlAscending,
                    XlYesNoGuess.xlGuess, Type.Missing, Type.Missing,
                    XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin,
                    XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal,
                    XlSortDataOption.xlSortNormal);

            Range TempAll = temp.Range["A2:J" + tmpAllCount];

            //Passando os dados para m7 de variação
            M7V.Activate();
            int rowsLength = GetCellsToSelect("A1").Count + 1;
            Range lastRow = M7V.Range["A" + rowsLength];
            TempAll.Copy();
            lastRow.PasteSpecial(XlPasteType.xlPasteValues);

            M7V.Range["J:J"].NumberFormat = "dd/mm/yyyy";
            M7V.Range["I:I"].NumberFormat = @"#,##0.00_ ;[Red]-#,##0.00 ";


            temp.Delete();

            M7Pbix.Save();
            M7Pbix.Close(false);
        }

        public static void DynimicTable()
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range all = currentSheet.Range[GetCellsToSelect("A3"), GetCellsToSelect("A3").End[XlDirection.xlToRight]];

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

        public static void NoDispProcess()
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            //trocar o formato numerico.
            GetCellsToSelect("B4").NumberFormat = "0";
            GetCellsToSelect("D4").NumberFormat = "0";

            Worksheet noDisponible = Globals.ThisAddIn.getActiveWorkbook().Sheets["No Disponible"];
            Range range = GetCellsToSelect("B4");
            int rows = range.Count + 3;

            ////Formulas - PASSO 4
            //Custo Init
            noDisponible.Range["J4:J" + rows].Formula = @"=VLOOKUP(B4,'M7'!B:J,9,0)";

            //Custo Total
            noDisponible.Range["K4:K" + rows].Formula = @"=J4*E4";

            //Segment
            noDisponible.Range["L4:L" + rows].Formula = @"=VLOOKUP(B4,'M7'!B:D,3,0)";

            //Classification
            noDisponible.Range["M4:M" + rows].Formula = @"=VLOOKUP(B4,'M7'!B:L,11,0)";

            //Disponível
            noDisponible.Range["N4:N" + rows].Value = "não";

            //Disponível(MRP)
            noDisponible.Range["O4:O" + rows].Value = "não";

            //Amount USD
            noDisponible.Range["P4:P" + rows].Formula = @"=K4/$J$1";

            //get USD
            noDisponible.Range["J1"].Formula = @"='M7'!J1";

            //subtotal
            noDisponible.Range["K2"].Formula = @"=SUBTOTAL(9,K4:K" + rows + ")";

            //Valor para comparação de %
            noDisponible.Range["K1"].Copy();
            noDisponible.Range["M1"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //filtros limpar dados N/D - PASSO 5
            if (GetCellsToSelect("M4").AutoFilter(13, "#N/D"))
            {
                Range all = GetCellsToSelect("A4:S4");
                all.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();


            //Segunda parte do processo
            Range r = GetCellsToSelect("A4");
            int rowsCount = r.Count + 3;

            Range D4 = noDisponible.Range["D4: D" + rowsCount];
            Range I4 = noDisponible.Range["I4: I" + rowsCount];
            Range L4 = noDisponible.Range["L4: L" + rowsCount];
            Range Q4 = noDisponible.Range["Q4: Q" + rowsCount];
            Range R4 = noDisponible.Range["R4: R" + rowsCount];
            Range S4 = noDisponible.Range["S4: S" + rowsCount];


            ////Procv no dia anterior - PASSO 6
            //Gestores
            PreviousDayProcv("No Disponible", Q4, @"=VLOOKUP(D4,'[M7 - STK 01.08.2023 -.xlsx]No Disponible'!$D:$Q,14,0)");

            //Resp.Inventário
            PreviousDayProcv("No Disponible", R4, @"=VLOOKUP(D4,'[M7 - STK 01.08.2023 -.xlsx]No Disponible'!$D:$R,15,0)");

            //Descrição Lugar
            PreviousDayProcv("No Disponible", S4, @"=VLOOKUP(Q4,'[M7 - STK 01.08.2023 -.xlsx]No Disponible'!$Q:$S,3,0)");

            //filtrar por lugar - PASSO 7
            if (D4.AutoFilter(4, "9ACERTO"))
            {
                Q4.AutoFilter(17, "SCM/Logistica [Pedro Iak]");
                R4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "William Baisi";
            }
            refreshFilter();
            
            //filtrar por lugar - PASSO 8
            if (I4.AutoFilter(9, "ENGENH"))
            {
                Q4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Engenharia Produtos [Luiz Facioli]";
                R4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Marcelo Perobelli";
            }
            refreshFilter();

            //Deletando as sucatas - PASSO 9
            if (Q4.AutoFilter(17, "#N/D"))
            {
                I4.AutoFilter(9, "SUCATA");
                I4.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();

            //Deletando MEMO - PASSO 10
            if (Q4.AutoFilter(17, "#N/D"))
            {
                D4.AutoFilter(4, "MEMO");
                D4.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();

            //Atribuindo aos N/D - PASSO 10
            Range found = Q4.Find("#N/D");
            if(found != null)
            {
                if (Q4.AutoFilter(17, "#N/D"))
                {
                    Range visible = Q4.SpecialCells(XlCellType.xlCellTypeVisible);
                    Range firstCell = visible.Cells[1];
                    string l = firstCell.Row.ToString();
                    visible.Formula = @"=VLOOKUP(D" + l + ",'Base Referencias'!H:J,3,0)";
                }
                refreshFilter();
            }


            //Filtar gestores para - PASSO 11
            Q4.AutoFilter(17, "Producao [Sergio Castro]");
            if (L4.AutoFilter(12, "SB", XlAutoFilterOperator.xlOr, "SW", XlAutoFilterOperator.xlFilterValues))
            {
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Rodrigo Mendonça]";
                R4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Rodrigo Mendonça";
            }
            refreshFilter();

            Q4.AutoFilter(17, "Producao [Rodrigo Mendonça]");
            if (L4.AutoFilter(12, "AB", XlAutoFilterOperator.xlOr, "ISS", XlAutoFilterOperator.xlFilterValues))
            {
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Sergio Castro]";
                R4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Sergio Castro";
            }
            refreshFilter();

            PBIX(noDisponible, "B4:S4", 3);
        }

        public static void FG_expedicao()
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            //Selecionar a planilha expedição
            Worksheet expeSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["FG_Expediçao"];
            expeSheet.Activate();

            //Pegar o tamanho das linhas
            Range range = GetCellsToSelect("A2");
            int rows = range.Count + 1;

            //Selecionar as colunas e executar procv - PASSO 2
            //Client
            Range p3 = expeSheet.Range["P3:P" + rows];
            PreviousDayProcv("FG_Expediçao", p3, @"=VLOOKUP(B3,'[M7 - STK 01.08.2023 -.xlsx]FG_Expediçao'!$B:$P,15,0)");
            p3.AutoFilter(16, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            p3.SpecialCells(XlCellType.xlCellTypeVisible).Formula = @"=VLOOKUP(B3,'M7'!B:M,12,0)";

            expeSheet.ShowAllData();

            //CS
            expeSheet.Range["Q3: Q" + rows].Formula = @"=VLOOKUP(B3,'M7'!B:O,14,0)";

            //Custo unit
            expeSheet.Range["R3: R" + rows].Formula = @"=VLOOKUP(B3,'M7'!B:J,9,0)";

            //Total BRL
            expeSheet.Range["S3:S" + rows].Formula = @"=R3*H3";

            //Total USD
            expeSheet.Range["T3:T" + rows].Formula = @"=S3/'M7'!$J$1";

            //Subtotal BRL
            expeSheet.Range["S1"].Formula = @"=SUBTOTAL(9,S3:S"+rows+")";

            //Subtotal USD
            expeSheet.Range["T1"].Formula = @"=SUBTOTAL(9,T3:T"+rows+")";

            expeSheet.Columns["O:O"].Delete();

            //Apagando valores nulls - PASSO 3
            Range t3 = expeSheet.Range["T3: T" + rows];

            Range found = t3.Find("#N/D");
            if (found != null)
            {
                if (t3.AutoFilter(20, "#N/D"))
                {
                    t3.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
                }
                refreshFilter("2:2");
            }
            PBIX(expeSheet, "B3:S3", 4);
        }

        public static void PBIX(Worksheet currentSheet, string range, int sheetNum)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            currentSheet.Activate();

            //Segunda parte
            Range All1 = GetCellsToSelect(range);

            //iniciando objeto
            Workbook M7Pbix = Globals.ThisAddIn.getActiveApp().Workbooks.Open(Excel.PathToPbix, UpdateLinks: false);
            Worksheet sheets = M7Pbix.Sheets[sheetNum];

            //Passando os dados para o destino
            sheets.Activate();
            Range All2 = GetCellsToSelect("A2:R2");
            All2.EntireRow.Delete();
            All1.Copy();
            sheets.Range["A2"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone);
            Clipboard.Clear();

            M7Pbix.Save();
            M7Pbix.Close(false);
        }

        public static void VlookUp(string sheet, int days, Range cells, string formula)
        {
            Globals.ThisAddIn.getActiveWorksheet();

            int mes = DateTime.Now.Month;
            int ano = DateTime.Now.Year;
            int yy = ano % 100;
            string dir = mes + "-" + yy;

            //Obtenha o nome do arquivo competo
            DateTime previousDay = DateTime.Today.AddDays(days);
            string dateValidate = previousDay.ToString("d").Replace("/", ".");
            string previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\"+ano+@"\M7 - STK "+dir+@"\M7 - STK " + dateValidate + " -.xlsx";

            string defaultData = "30.06.2023";

            if (!File.Exists(previousFile))
            {
                for (int d = -1; d > -10; d--)
                {
                    previousDay = DateTime.Today.AddDays(days+d);
                    dateValidate = previousDay.ToString("d").Replace("/", ".");
                    previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\" + ano.ToString() + @"\M7 - STK " + dir + @"\M7 - STK " + dateValidate + " -.xlsx";
                    if (File.Exists(previousFile))
                    {
                        //Selecione o arquivo para o procv
                        Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                        Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                        worksheetTemp.Activate();

                        string realV = formula.Replace(defaultData, dateValidate);

                        cells.Formula = realV;

                        workbookTemp.Close(false);
                        break;
                    }
                }
                if (!File.Exists(previousFile) && previousFile.Contains(dir))
                {
                    for (int d = -0; d > -10; d--)
                    {
                        DateTime dataAtual = DateTime.Today;
                        DateTime primeiroDiaDoMesAtual = new DateTime(dataAtual.Year, dataAtual.Month, 1);
                        DateTime ultimoDiaDoMesPassado = primeiroDiaDoMesAtual.AddDays(-1 + d);

                        string valid = ultimoDiaDoMesPassado.ToString("d").Replace("/", ".");
                        string month = ultimoDiaDoMesPassado.ToString("MM/yy").Replace("/", "-");

                        previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\" + ano + @"\M7 - STK " + dir + @"\M7 - STK " + valid + " -.xlsx";

                        string newPath = previousFile.Replace(dir, month);

                        if (File.Exists(newPath))
                        {
                            //Selecione o arquivo para o procv
                            Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(newPath, UpdateLinks: false);
                            Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                            worksheetTemp.Activate();

                            string realV = formula.Replace(defaultData, valid);

                            cells.Formula = realV;

                            workbookTemp.Close(false);
                            break;
                        }
                    }
                }
            }
            else
            {
                //Selecione o arquivo para o procv
                Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                worksheetTemp.Activate();
                //
                string realV = formula.Replace(defaultData, dateValidate);
                cells.Formula = realV;
                //close temp
                workbookTemp.Close(false);
            }
        }
    

        public static void PreviousDayProcv(string sheet, Range cell, string procv)
        {
            Globals.ThisAddIn.getActiveWorksheet();

            int mes = DateTime.Now.Month;
            int ano = DateTime.Now.Year;
            int yy = ano % 100;
            string dir = mes + "-" + yy;
            int lastDay = 31;
            
            //Obtenha o nome do arquivo competo
            DateTime previousDay = DateTime.Today.AddDays(-1);
            string dateValidate = previousDay.ToString("dd/MM/yyyy").Replace("/", ".");
            string previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\"+ ano + @"\M7 - STK "+ dir + @"\M7 - STK " +dateValidate+" -.xlsx";
            string defaultData = "01.08.2023";

            if (!File.Exists(previousFile))
            {
                for (int d = -1; d > -10; d--)
                {
                    previousDay = DateTime.Today.AddDays(-1 + d);
                    dateValidate = previousDay.ToString("d").Replace("/", ".");
                    previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\" + ano + @"\M7 - STK " + dir + @"\M7 - STK " + dateValidate + " -.xlsx";
                    
                    if (File.Exists(previousFile))
                    {
                        //Selecione o arquivo para o procv
                        Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                        Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                        worksheetTemp.Activate();

                        string realV = procv.Replace(defaultData, dateValidate);
                        cell.Formula = realV;

                        workbookTemp.Close(false);
                        break;
                    }
                }
                if (!File.Exists(previousFile) && previousFile.Contains(dir))
                {
                    for (int d = -0; d > -10; d--)
                    {
                        DateTime dataAtual = DateTime.Today;
                        DateTime primeiroDiaDoMesAtual = new DateTime(dataAtual.Year, dataAtual.Month, 1);
                        DateTime ultimoDiaDoMesPassado = primeiroDiaDoMesAtual.AddDays(-1+d);

                        string valid = ultimoDiaDoMesPassado.ToString("d").Replace("/", ".");
                        string month = ultimoDiaDoMesPassado.ToString("MM/yy").Replace("/", "-");

                        previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\" + ano + @"\M7 - STK " + dir + @"\M7 - STK " + valid + " -.xlsx";

                        string newPath = previousFile.Replace(dir, month);

                        if (File.Exists(newPath))
                        {
                            //Selecione o arquivo para o procv
                            Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(newPath, UpdateLinks: false);
                            Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                            worksheetTemp.Activate();

                            string realV = procv.Replace(defaultData, dateValidate);
                            cell.Formula = realV;

                            workbookTemp.Close(false);
                            break;
                        }
                    }
                }
            }
            else
            {

                //Selecione o arquivo para o procv
                Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                //verificar isso
                Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                worksheetTemp.Activate();

                //troque o dia colocado para o dia anterior
                string realV = procv.Replace(defaultData, dateValidate);

                cell.Formula = realV;

                //ver forma de fechar o arquivo ao final do processo
                workbookTemp.Close(false);
            }
        }


        public static void refreshFilter(string range = "3:3")
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            currentSheet.AutoFilterMode = false;

            try
            {
                _ = currentSheet.EnableAutoFilter;
                currentSheet.Rows[range].AutoFilter();
            }
            catch (Exception)
            {
                MessageBox.Show("Filtro não foi ativado");
            }
        }


        public static Range GetCellsToSelect(String cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range cellSelect = currentSheet.Range[cell];
            Range sl = currentSheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
            return sl;
        }


        public static void GetData(string tmpSheet, string range)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[tmpSheet];
            Range cells = currentSheet.Range[range];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];
            select.ClearFormats();
            select.Copy();
        }


        public static void SetData(string tmpSheet, string range , string cell, string sheet, Workbook wb)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            excelApp.DisplayAlerts = false;

            GetData(tmpSheet, range);

            Worksheet Wsheet = wb.Sheets[sheet];
            Wsheet.Activate();
            Range init = Wsheet.Range[cell];

            init.PasteSpecial(XlPasteType.xlPasteAll);

            if (init.Value != null)
            {
                Clipboard.Clear();
            }
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
    }
}

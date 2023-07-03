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


        public static void M7Formulas()
        {
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
            currentSheet.Range["M2"].Formula = @"=SUBTOTAL(9,M4:M" + rowsCount + ")";
            currentSheet.Range["O2"].Formula = @"=SUBTOTAL(9,O4:O" + rowsCount + ")";

            FilterDataToClient();

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
                data.SpecialCells(XlCellType.xlCellTypeVisible).Clear();

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
            //n4.SpecialCells(XlCellType.xlCellTypeVisible).Formula = @"=VLOOKUP(L4,'Clientes'!A:B,2,0)";


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

                PreviousDayProcv("M7", l4, @"=VLOOKUP(A4,'[M7 - STK 01.07.2023 -.xlsx]M7'!$A:$L,12,0)");
                
            }
            l4.AutoFilter(12, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            l4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Others";
            refreshFilter();


            //procv nas planilhas para CS
            PreviousDayProcv("M7", n4, @"=VLOOKUP(A4,'[M7 - STK 01.07.2023 -.xlsx]M7'!$B:$O,14,0)");

            n4.AutoFilter(14, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            n4.SpecialCells(XlCellType.xlCellTypeVisible).Formula = @"=VLOOKUP(L4,'Clientes'!A:B,2,0)";
            n4.AutoFilter(14, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            n4.SpecialCells(XlCellType.xlCellTypeVisible).Cells.Clear();
            refreshFilter();

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
            //trocar o formato numerico.
            GetCellsToSelect("B4").NumberFormat = "0";
            GetCellsToSelect("D4").NumberFormat = "0";

            Worksheet noDisponible = Globals.ThisAddIn.getActiveWorkbook().Sheets["No Disponible"];
            Range range = GetCellsToSelect("B4");
            int rows = range.Count + 3;

            ////Formulas - PASSO 4
            //Custo Init
            noDisponible.Range["J4:J" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:I,9,0)";

            //Custo Total
            noDisponible.Range["K4:K" + rows].Formula = @"=J4*E4";

            //Segment
            noDisponible.Range["L4:L" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:C,3,0)";

            //Classification
            noDisponible.Range["M4:M" + rows].Formula = @"=VLOOKUP(B4,'M7'!A:K,11,0)";

            //Disponível
            noDisponible.Range["N4:N" + rows].Value = "não";

            //Disponível(MRP)
            noDisponible.Range["O4:O" + rows].Value = "não";

            //Amount USD
            noDisponible.Range["P4:P" + rows].Formula = @"=K4/$J$1";

            //get USD
            noDisponible.Range["J1"].Formula = @"='M7'!I1";

            //subtotal
            noDisponible.Range["K2"].Formula = @"=SUBTOTAL(9,K4:K" + rows + ")";

            //Valor para comparação de %
            noDisponible.Range["M1"].Value = noDisponible.Range["K1"].Value;


            //Segunda parte do processo
            Range D4 = noDisponible.Range["D4:D" + rows];
            Range I4 = noDisponible.Range["I4: I" + rows];
            Range L4 = noDisponible.Range["L4: L" + rows];
            Range Q4 = noDisponible.Range["Q4: Q" + rows];
            Range R4 = noDisponible.Range["R4: R" + rows];
            Range S4 = noDisponible.Range["S4: S" + rows];

            //filtros limpar dados N/D - PASSO 5
            string[] listCriteria = new string[]
            {
                "ANALISE",
                "BLOQUEAD",
                "VENCIDO"
            };

            GetCellsToSelect("I4").AutoFilter(9, listCriteria, XlAutoFilterOperator.xlFilterValues);

            Q4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Qualidade [Marcelo Santos]";
            R4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "Julio Moura";
            S4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "QUALIDADE PRODUÇAO";

            refreshFilter();

            //filtros limpar dados N/D - PASSO 6
            if (GetCellsToSelect("M4").AutoFilter(13, "#N/D"))
            {
                Range all = GetCellsToSelect("A4:S4");
                all.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();


            ////Procv no dia anterior - PASSO 7
            //Gestores
            PreviousDayProcv("No Disponible", Q4, @"=VLOOKUP(D4,'[M7 - STK 01.06.2023 -.xlsx]No Disponible'!$D:$Q,14,0)");

            //Resp.Inventário
            PreviousDayProcv("No Disponible", R4, @"=VLOOKUP(Q4,'[M7 - STK 01.06.2023 -.xlsx]No Disponible'!$Q:$R,2,0)");

            //Descrição Lugar
            PreviousDayProcv("No Disponible", S4, @"=VLOOKUP(Q4,'[M7 - STK 01.06.2023 -.xlsx]No Disponible'!$Q:$S,3,0)");


            //Filtar gestores para- PASSO 8
            if (Q4.AutoFilter(17, "Producao [Douglas Vale]"))
            {
                string[] filterCriteria = new string[] { "SW", "SB" };
                L4.AutoFilter(12, filterCriteria, XlAutoFilterOperator.xlFilterValues);
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Rodrigo Mendonça]";

            }
            if (Q4.AutoFilter(17, "Producao [Rodrigo Mendonça]"))
            {
                string[] filterCriteria = new string[] { "AB", "ISS" };
                L4.AutoFilter(12, filterCriteria, XlAutoFilterOperator.xlFilterValues);
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Douglas Vale]";
            }
            refreshFilter();

            //filtrar por lugar - PASSO 9
            if (D4.AutoFilter(4, "9ACERTO"))
            {
                Q4.AutoFilter(17, "SCM/Logistica [Pedro Iak]");
                R4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "William Baisi";
            }
            refreshFilter();

            //Deletando as sucatas - PASSO 10
            if (Q4.AutoFilter(17, "#N/D"))
            {
                I4.AutoFilter(9, "SUCATA");
                I4.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            refreshFilter();

        }


        public static void FG_expedicao()
        {
            //Selecionar a planilha expedição
            Worksheet expeSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["FG_Expediçao"];
            expeSheet.Activate();

            //Pegar o tamanho das linhas
            Range range = GetCellsToSelect("A2");
            int rows = range.Count + 1;

            //Selecionar as colunas e executar procv - PASSO 2
            //Client
            Range p3 = expeSheet.Range["P3:P" + rows];
            PreviousDayProcv("FG_Expediçao", p3, @"=VLOOKUP(B3,'[M7 - STK 01.07.2023 -.xlsx]FG_Expediçao'!$B:$P,15,0)");
            p3.AutoFilter(16, filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
            p3.SpecialCells(XlCellType.xlCellTypeVisible).Formula = @"=VLOOKUP(B3,'M7'!A:L,12,0)";

            expeSheet.AutoFilterMode = false;

            //CS
            expeSheet.Range["Q3: Q" + rows].Formula = @"=VLOOKUP(B3,'M7'!A:N,14,0)";

            //Custo unit
            expeSheet.Range["R3: R" + rows].Formula = @"=VLOOKUP(B3,'M7'!A:I,9,0)";

            //Total BRL
            expeSheet.Range["S3:S" + rows].Formula = @"=R3*H3";

            //Total USD
            expeSheet.Range["T3:T" + rows].Formula = @"=S3/'M7'!$I$1";

            //Subtotal BRL
            expeSheet.Range["S1"].Formula = @"=SUBTOTAL(9,S3:S"+rows+")";

            //Subtotal USD
            expeSheet.Range["T1"].Formula = @"=SUBTOTAL(9,T3:T"+rows+")";

            //Atualizar tabela dinamica - PASSO 3
        }


        public static void PreviousDayProcv(string sheet, Range cell, string procv)
        {
            Globals.ThisAddIn.getActiveWorksheet();

            //Obtenha o nome do arquivo competo
            DateTime previousDay = DateTime.Today.AddDays(-1);
            string dateValidate = previousDay.ToString("dd/MM/yyyy").Replace("/", ".");
            string previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK "+dateValidate+" -.xlsx";

            if (File.Exists(previousFile))
            {
                //Selecione o arquivo para o procv
                Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                //verificar isso
                Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                worksheetTemp.Activate();

                //troque o dia colocado para o dia anterior
                string realV = procv.Replace("01.07.2023", dateValidate);

                cell.Formula = realV;

                //ver forma de fechar o arquivo ao final do processo
                workbookTemp.Close(false);
            }
            else
            {
                previousDay = previousDay.AddDays(-1);
                dateValidate = previousDay.ToString("d").Replace("/", ".");
                previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

                if (File.Exists(previousFile))
                {
                    //Selecione o arquivo para o procv
                    Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                    //verificar isso
                    Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                    worksheetTemp.Activate();

                    //troque o dia colocado para o dia anterior
                    string realV = procv.Replace("01.07.2023", dateValidate);

                    cell.Formula = realV;

                    //ver forma de fechar o arquivo ao final do processo
                    workbookTemp.Close(false);
                }
                else
                {
                    previousDay = previousDay.AddDays(-1);
                    dateValidate = previousDay.ToString("d").Replace("/", ".");
                    previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

                    if (File.Exists(previousFile))
                    {
                        //Selecione o arquivo para o procv
                        Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                        //verificar isso
                        Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                        worksheetTemp.Activate();

                        //troque o dia colocado para o dia anterior
                        string realV = procv.Replace("01.07.2023", dateValidate);

                        cell.Formula = realV;

                        //ver forma de fechar o arquivo ao final do processo
                        workbookTemp.Close(false);
                    }
                    else
                    {
                        previousDay = previousDay.AddDays(-1);
                        dateValidate = previousDay.ToString("d").Replace("/", ".");
                        previousFile = @"C:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

                        if (File.Exists(previousFile))
                        {
                            //Selecione o arquivo para o procv
                            Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                            //verificar isso
                            Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                            worksheetTemp.Activate();

                            //troque o dia colocado para o dia anterior
                            string realV = procv.Replace("01.07.2023", dateValidate);

                            cell.Formula = realV;

                            //ver forma de fechar o arquivo ao final do processo
                            workbookTemp.Close(false);
                        }
                    }
                }
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
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[tmpSheet];
            Range cells = currentSheet.Range[range];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];
            select.ClearFormats();
            select.Copy();
        }


        public static void SetData(string tmpSheet, string range , string cell, string sheet, Workbook wb)
        {
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

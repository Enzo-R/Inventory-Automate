using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainingVSTO.Models
{
    public class NoDisp
    {
        public static void NoDispProcess()
        {
            //trocar o formato numerico.
            Utils.GetCellsToSelect("B4").NumberFormat = "0";
            Utils.GetCellsToSelect("D4").NumberFormat = "0";

            Worksheet noDisponible = Globals.ThisAddIn.getActiveWorkbook().Sheets["No Disponible"];
            Range range = Utils.GetCellsToSelect("B4");
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
            if (Utils.GetCellsToSelect("M4").AutoFilter(13, "#N/D"))
            {
                Range all = Utils.GetCellsToSelect("A4:S4");
                all.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            Utils.refreshFilter();


            //Segunda parte do processo
            Range D4 = noDisponible.Range["D4: D" + rows];
            Range I4 = noDisponible.Range["I4: I" + rows];
            Range L4 = noDisponible.Range["L4: L" + rows];
            Range Q4 = noDisponible.Range["Q4: Q" + rows];
            Range R4 = noDisponible.Range["R4: R" + rows];
            Range S4 = noDisponible.Range["S4: S" + rows];


            ////Procv no dia anterior - PASSO 6
            //Gestores
            Utils.PreviousDayProcv("No Disponible", Q4, @"=VLOOKUP(D4,'[M7 - STK 01.07.2023 -.xlsx]No Disponible'!$D:$Q,14,0)");

            //Resp.Inventário
            Utils.PreviousDayProcv("No Disponible", R4, @"=VLOOKUP(Q4,'[M7 - STK 01.07.2023 -.xlsx]No Disponible'!$Q:$R,2,0)");

            //Descrição Lugar
            Utils.PreviousDayProcv("No Disponible", S4, @"=VLOOKUP(Q4,'[M7 - STK 01.07.2023 -.xlsx]No Disponible'!$Q:$S,3,0)");



            //filtrar por lugar - PASSO 7
            if (D4.AutoFilter(4, "9ACERTO"))
            {
                Q4.AutoFilter(17, "SCM/Logistica [Pedro Iak]");
                R4.SpecialCells(XlCellType.xlCellTypeVisible).Value = "William Baisi";
            }
            Utils.refreshFilter();

            //Deletando as sucatas - PASSO 8
            if (Q4.AutoFilter(17, "#N/D"))
            {
                I4.AutoFilter(9, "SUCATA");
                I4.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            Utils.refreshFilter();

            //Deletando MEMO - PASSO 9
            if (Q4.AutoFilter(17, "#N/D"))
            {
                D4.AutoFilter(4, "MEMO");
                D4.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
            }
            Utils.refreshFilter();

            //Filtar gestores para - PASSO 10
            Q4.AutoFilter(17, "Producao [Rodrigo Mendonça]");
            if (L4.AutoFilter(12, "AB", XlAutoFilterOperator.xlOr, "ISS", XlAutoFilterOperator.xlFilterValues))
            {
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Douglas Vale]";
            }
            Utils.refreshFilter();

            Q4.AutoFilter(17, "Producao [Douglas Vale]");
            if (L4.AutoFilter(12, "SB", XlAutoFilterOperator.xlOr, "SW", XlAutoFilterOperator.xlFilterValues))
            {
                Q4.SpecialCells(XlCellType.xlCellTypeVisible)
                    .Value = "Producao [Rodrigo Mendonça]";
            }
            Utils.refreshFilter();
        }

    }
}

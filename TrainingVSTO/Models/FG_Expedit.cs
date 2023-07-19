using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainingVSTO.Models
{
    public class FG_Expedit
    {
        public static void FG_expedicao()
        {
            //Selecionar a planilha expedição
            Worksheet expeSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["FG_Expediçao"];
            expeSheet.Activate();

            //Pegar o tamanho das linhas
            Range range = Utils.GetCellsToSelect("A2");
            int rows = range.Count + 1;

            //Selecionar as colunas e executar procv - PASSO 2
            //Client
            Range p3 = expeSheet.Range["P3:P" + rows];
            Utils.PreviousDayProcv("FG_Expediçao", p3, @"=VLOOKUP(B3,'[M7 - STK 01.07.2023 -.xlsx]FG_Expediçao'!$B:$P,15,0)");
            p3.AutoFilter(16, Utils.filterCriteriaNull, XlAutoFilterOperator.xlFilterValues);
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
            expeSheet.Range["S1"].Formula = @"=SUBTOTAL(9,S3:S" + rows + ")";

            //Subtotal USD
            expeSheet.Range["T1"].Formula = @"=SUBTOTAL(9,T3:T" + rows + ")";


            //Apagando valores nulls - PASSO 3
            Range R3 = expeSheet.Range["R3: R" + rows];
            if (R3.AutoFilter(18, "#N/D", XlAutoFilterOperator.xlFilterValues))
            {
                R3.SpecialCells(XlCellType.xlCellTypeBlanks).EntireRow.Delete();
            }

        }

    }
}

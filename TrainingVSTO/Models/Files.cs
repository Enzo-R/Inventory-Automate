﻿using System;
            string PathToServer = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2022\Teste"
                                                                        + @"\M7 - STK " + dateValidate + " -.xlsx";
                    .SaveAs(@"C:\Users\EROLIVEIRA\OneDrive - Joyson Group\Área de Trabalho\Joyson\M7 - STK " + dateValidate + " -.xlsx");

        public static void OpenNoDispSTK()
        {
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(Excel.PathToM7DOpen);
            Worksheet Sheet = workbook.Sheets["M7"];
        }
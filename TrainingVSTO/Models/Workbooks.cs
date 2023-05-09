﻿using System;
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

            string Text = File.ReadAllText(path);
            string[] rows = Text.Split('\n');

            // Cria um loop para iterar pelas linhas e copiar cada uma para a planilha do Excel
            int linhaAtual = 1;
            foreach (string row in rows)
            {
                // Seleciona a célula atual na planilha do Excel
                Range range = currentSheet.Cells[linhaAtual, 1];

                // Cola a linha atual na célula selecionada
                range.Value2 = row;

                // Incrementa o número da linha atual
                linhaAtual++;
            }

            ReleaseObject(currentSheet);

            //range.PasteSpecial(XlPasteType.xlPasteAll);
            //como copiar o conteudo de um arquivo de texto
            //string conteudo = File.ReadAllText(path);
            //Range position = currentSheet.Columns[1];
            //position.Value = conteudo;

            //string[] linhas = File.ReadAllLines(path);

            ////Escrever os dados do arquivo de texto na planilha Excel
            //for (int i = 0; i < linhas.Length; i++)
            //{
            //    string[] colunas = linhas[i].Split('\t');// Separador de colunas no arquivo de texto, neste caso é o TAB
            //    for (int j = 0; j < colunas.Length; j++)
            //    {
            //        currentSheet.Cells[i + 1, j + 1].Value = colunas[j];                    
            //    }
            //}

        }
        public static object Data(string sheet)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[sheet];
            object dados = currentSheet.Range["B5 : K20000"].Value;
            return dados;
        }
        public static void JoinClass()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture =
                            new System.Globalization.CultureInfo("en-US");

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range f1;

            f1 = currentSheet.Range["K4:K20000"];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";
            //f1.Cells.AutoFill(f1);

            //Range col = currentSheet.Columns.Count;

            //retira os valores nullos
            //if (f1.Cells.Value == "#N/D")
            //{
            //    f1["A4 : J20000"].AutoFilter(1, "Valor 1", XlAutoFilterOperator.xlOr, 2, "Valor 2");

            //    currentSheet.Range["K"].AutoFilter(1, Criteria1: "#N/D");
            //    currentSheet.Range["K4:K20000"]
            //        .SpecialCells(XlCellType.xlCellTypeVisible)
            //        .Delete();
            //}

            Models.Workbooks.ReleaseObject(currentSheet);
        }

    }
}

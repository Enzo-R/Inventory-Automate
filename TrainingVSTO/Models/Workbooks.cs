using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TrainingVSTO.Models
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar elementos dentro do Excel
        public static void SheetSelect(string sheet, string path)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook workbook = excelApp.Workbooks.Open(path);
            Worksheet originalSheet = excelApp.ActiveWorkbook.Sheets[sheet];
            originalSheet.Activate();
        }
        public static void releaseObject(object obj)
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

            releaseObject(currentSheet);

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
        public static void GetData(string sheet)
        {
            SheetSelect(sheet, Models.Excel.PathToM7D);
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            object dados = currentSheet.Range["B5 : K20000"].Value;
            Excel.Data = dados;
        }

    }
}

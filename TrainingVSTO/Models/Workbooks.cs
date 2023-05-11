using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;

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
        public static void VLookUp()
        {
            //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];
            Range f1;
            //forma de obter o tamanho das colunas.
            f1 = currentSheet.Range["K4:K14598"];
            f1.Formula = @"=VLOOKUP(B4,'Base Contas'!A:C,3,0)";

            FilterData();


        }
        public static void FilterData()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets["M7"];

            string[] filterCriteria = new string[] {
                "BENS CAPITAL EM PROCESSO",
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
                "MERCADORIAS EM TRANSITO",
                "MERCADORIAS PARA REVENDA",
                "MOVEIS E UTENSILIOS",
                "PRODUTOS ACABADOS",
                "PRODUTOS SEMIACABADOS",
                "SUBPRODUTO",
                "USO CONSUMO MAQ.EQUIP."
            };
            Range dataRange = currentSheet.Range["$A$3:$P$14619"];

            if (currentSheet.Range["K:K"].AutoFilter(11, "#N/D"))
            {
                dataRange.AutoFilter(4, filterCriteria, XlAutoFilterOperator.xlFilterValues);
                currentSheet.Range["$A$4:$P$14619"].SpecialCells(XlCellType.xlCellTypeVisible).Clear();

            }
        }
    }
}

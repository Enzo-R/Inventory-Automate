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

namespace TrainingVSTO
{
    public class Workbooks
    {
        //classe responsavel por manipular e criar elementos dentro do Excel
        public static void clearWorksheet()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            var text = currentSheet.Columns[1].Value;

            if (text != null)
            {
                currentSheet.Columns.Rows.Clear();
                text = null;
            }
        }
        public static void ReadAndWriteArq(string path)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            //string[] linhas = File.ReadAllLines(path);

            ////Escrever os dados do arquivo de texto na planilha Excel
            //for (int i = 0; i < linhas.Length; i++)
            //{
            //    string[] colunas = linhas[i].Split('\t');// Separador de colunas no arquivo de texto, neste caso é o TAB
            //    for (int j = 0; j < colunas.Length; j++)
            //    {
            //        var arq = currentSheet.Cells[i + 1, j + 1].Value = colunas[j];
            //    }
            //}

            void releaseObject(object obj)
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
            releaseObject(currentSheet);

            //range.PasteSpecial(XlPasteType.xlPasteAll);
            //como copiar o conteudo de um arquivo de texto
            string conteudo = File.ReadAllText(path);
            Range position = currentSheet.Columns[1];
            position.Value = conteudo;

        }
    }
}

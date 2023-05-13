using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using TrainingVSTO;

namespace TrainingVSTO
{
    public partial class Ribbon1
    {
        private void AbreModeloClick(object sender, RibbonControlEventArgs e)
        {
            string day = editBox1.Text;
            Models.Workbooks.Data("M7 EF");
            if (day != "")
            {
                Models.Files.CreateM7D(day);
            }
            else
            {
                MessageBox.Show("Insira uma data para o inventário!");
            }
        }
        private void OpenFile_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = Globals.ThisAddIn.getActiveWorkbook();
            Worksheet Sheet = workbook.Sheets["Original"];
            Sheet.Activate();
            try
            {
                Sheet.Cells.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally 
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "text (*.txt)|*.txt";
                openFile.Title = "Open the file";

                // Exibir o diálogo e verificar se o usuário clicou em "OK"
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    string path = openFile.FileName;
                    Models.Workbooks.ReadAndWriteArq(path);
                }
            }
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Models.Workbooks.VLookUp();
            //string name = editBox1.Text;
            //string path = @"C:\Users\Enzo\Desktop\clientes\"+name;

            //if (name != "")
            //{
            //    Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            //    currentSheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path, XlFixedFormatQuality.xlQualityMinimum);
            //}
            //else
            //{
            //    MessageBox.Show("Insira uma data para o inventário!");
            //}

        }
    }
}

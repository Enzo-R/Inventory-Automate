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

            if (day != "")
            {
                Files.CreateM7D(day);
            }
            else
            {
                MessageBox.Show("Insira uma data para o inventário!");
            }

        }
        private void OpenFile_Click(object sender, RibbonControlEventArgs e)
        {
            string sh = "Original";
            Workbooks.SheetSelect(sh);

            try
            {
                Workbooks.ClearWorksheet();
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
                    Workbooks.ReadAndWriteArq(path);
                }

            }
        }
    }
}

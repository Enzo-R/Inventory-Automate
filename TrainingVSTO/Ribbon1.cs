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
            Models.Workbooks.Data("M7 EF", "B5 : K5");
            Models.Files.CreateM7D();

        }
        private void OpenFile_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = Globals.ThisAddIn.getActiveWorkbook();
            string wbName = workbook.Name;

            if (wbName.Contains("AbreModelo7"))
            {
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
            else
            {

                Models.Workbooks.Data("Ddos", "A2 : I2");

                try
                {
                    //Sheet.Cells.Clear();
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

        }
        private void InventoryNoDisponible_(object sender, RibbonControlEventArgs e)
        {

        }

    }
}

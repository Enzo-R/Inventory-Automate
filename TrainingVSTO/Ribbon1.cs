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
using TrainingVSTO.Models;

namespace TrainingVSTO
{
    public partial class Ribbon1
    {
        private void AbreModeloClick(object sender, RibbonControlEventArgs e)
        {
            Models.Workbooks.GetData("M7 EF", "B5:K5");
            Models.Files.CreateM7D();
            //Models.Files.OpenM7Model().Close(true);

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
        private void InventoryNoDisponible_(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel (*.xlsx)|*.xlsx";
                openFile.Title = "Open No disponible";

                // Exibir o diálogo e verificar se o usuário clicou em "OK"
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    string path = openFile.FileName;
                    Files.OpenNoDispSTK(path, "Ddos");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        private void OpenFG(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel (*.xlsx)|*.xlsx";
                openFile.Title = "Open FG_export";

                // Exibir o diálogo e verificar se o usuário clicou em "OK"
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    string path = openFile.FileName;
                    Files.OpenFG(path, "Ddos");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

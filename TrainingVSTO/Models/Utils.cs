using System;
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
using System.Net.Http;
using Microsoft.Office.Tools.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Net.NetworkInformation;

namespace TrainingVSTO.Models
{
    public class Utils
    {
        //classe responsavel por manipular e criar elementos dentro do Excel

        public static string[] filterCriteriaNull = new string[]
        {
                "#N/D",
                "0",
                "="
        };


        public static void VlookUp(string sheet, int days, Range cells, string formula)
        {
            Globals.ThisAddIn.getActiveWorksheet();

            //Obtenha o nome do arquivo competo
            DateTime previousDay = DateTime.Today.AddDays(days);
            string dateValidate = previousDay.ToString("d").Replace("/", ".");
            string previousFile = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

            string defaultData = "30.06.2023";

            if (!File.Exists(previousFile))
            {
                for (int d = -1; d > -10; d--)
                {
                    previousDay = DateTime.Today.AddDays(days+d);
                    dateValidate = previousDay.ToString("d").Replace("/", ".");
                    previousFile = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

                    if (File.Exists(previousFile))
                    {
                        //Selecione o arquivo para o procv
                        Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                        Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                        worksheetTemp.Activate();

                        string realV = formula.Replace(defaultData, dateValidate);

                        cells.Formula = realV;

                        workbookTemp.Close(false);
                        break;
                    }
                }
                if (!File.Exists(previousFile) && previousFile.Contains("07-23"))
                {
                    for (int d = -0; d > -10; d--)
                    {
                        DateTime imim = DateTime.Today.AddMonths(-1);
                        previousDay = DateTime.Today.AddDays(days + d);
                        dateValidate = previousDay.ToString("d").Replace("/", ".");
                        previousFile = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";
                        string month = imim.ToString("MM/yy").Replace("/", "-");
                        string newPath = previousFile.Replace("07-23", month);

                        if (File.Exists(newPath))
                        {
                            //Selecione o arquivo para o procv
                            Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(newPath, UpdateLinks: false);
                            Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                            worksheetTemp.Activate();

                            string realV = formula.Replace(defaultData, dateValidate);

                            cells.Formula = realV;

                            workbookTemp.Close(false);
                            break;
                        }
                    }
                }
            }
            else
            {
                //Selecione o arquivo para o procv
                Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                worksheetTemp.Activate();
                //
                string realV = formula.Replace(defaultData, dateValidate);
                cells.Formula = realV;
                //close temp
                workbookTemp.Close(false);
            }
        }
    

        public static void PreviousDayProcv(string sheet, Range cell, string procv)
        {
            Globals.ThisAddIn.getActiveWorksheet();

            //Obtenha o nome do arquivo competo
            DateTime previousDay = DateTime.Today.AddDays(-1);
            string dateValidate = previousDay.ToString("dd/MM/yyyy").Replace("/", ".");
            string previousFile = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK "+dateValidate+" -.xlsx";
            string defaultData = "01.07.2023";

            if (!File.Exists(previousFile))
            {
                for (int d = -1; d > -10; d--)
                {
                    previousDay = DateTime.Today.AddDays(-1 + d);
                    dateValidate = previousDay.ToString("d").Replace("/", ".");
                    previousFile = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 07-23\M7 - STK " + dateValidate + " -.xlsx";

                    if (File.Exists(previousFile))
                    {
                        //Selecione o arquivo para o procv
                        Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);
                        Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                        worksheetTemp.Activate();

                        string realV = procv.Replace(defaultData, dateValidate);
                        cell.Formula = realV;

                        workbookTemp.Close(false);
                        break;
                    }
                }
            }
            else
            {
                //Selecione o arquivo para o procv
                Workbook workbookTemp = Globals.ThisAddIn.getActiveApp().Workbooks.Open(previousFile, UpdateLinks: false);

                //verificar isso
                Worksheet worksheetTemp = workbookTemp.Worksheets[sheet];
                worksheetTemp.Activate();

                //troque o dia colocado para o dia anterior
                string realV = procv.Replace(defaultData, dateValidate);

                cell.Formula = realV;

                //ver forma de fechar o arquivo ao final do processo
                workbookTemp.Close(false);
            }
        }


        public static void refreshFilter(string range = "3:3")
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            currentSheet.AutoFilterMode = false;

            try
            {
                _ = currentSheet.EnableAutoFilter;
                currentSheet.Rows[range].AutoFilter();
            }
            catch (Exception)
            {
                MessageBox.Show("Filtro não foi ativado");
            }
        }


        public static Range GetCellsToSelect(String cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            Range cellSelect = currentSheet.Range[cell];
            Range sl = currentSheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
            return sl;
        }


        public static void GetData(string tmpSheet, string range)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorkbook().Sheets[tmpSheet];
            Range cells = currentSheet.Range[range];
            Range select = currentSheet.Range[cells, cells.End[XlDirection.xlDown]];
            select.ClearFormats();
            select.Copy();
        }


        public static void SetData(string tmpSheet, string range , string cell, string sheet, Workbook wb)
        {
            GetData(tmpSheet, range);

            Worksheet Wsheet = wb.Sheets[sheet];
            Wsheet.Activate();
            Range init = Wsheet.Range[cell];

            init.PasteSpecial(XlPasteType.xlPasteAll);

            if (init.Value != null)
            {
                Clipboard.Clear();
            }
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
    }
}

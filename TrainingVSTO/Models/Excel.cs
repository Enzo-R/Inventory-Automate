using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TrainingVSTO.Models
{
    public static class Excel
    {
        public static string PathToM7DOpen
        {
            get { return @"S:\Log_Planej_Adm\PERSONAL\Enzo Rodrigues\Default Files\M7 - ex -.xlsx"; }
        }
        public static string PathToM7DModel
        {
            get { return @"S:\Log_Planej_Adm\PERSONAL\Enzo Rodrigues\Default Files\AbreModelo7 - Rev1.xlsm"; }
        }

        public static DateTime date = DateTime.Today;
        public static string date1 = date.ToString("d");
        public static string dateValidate = date1.Replace("/", ".");
        public static string PathToServer = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 06-23\M7 - STK " + dateValidate + " -.xlsx";
        
    }
}

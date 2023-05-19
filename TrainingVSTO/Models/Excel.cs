using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TrainingVSTO.Models
{
    public static class Excel
    {
        public static string PathToM7DOpen
        {
            get { return "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\M7 - STK 28.04.2023 - novo.xlsx"; }
        }
        public static string PathToM7DModel
        {
            get { return "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\AbreModelo7 - Rev1.xlsm"; }
        }

        public static DateTime date = DateTime.Today;
    }
}

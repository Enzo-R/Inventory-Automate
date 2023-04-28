using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrainingVSTO.Models
{
    public static class Excel
    {
        public static object Data { get; set; }
        public static string PathM7C
        {
            get { return "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\Model.xlsx"; }
        }
        public static string PathToM7D
        {
            get { return "C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\AbreModelo7 - Rev1 - Copia.xlsm"; }
        }
    }
}

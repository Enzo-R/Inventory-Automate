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
        public static DateTime dateToday = DateTime.Today;
        public static string dateValidate = dateToday.ToString("d").Replace("/", ".");
        
        //@"S:\Log_Planej_Adm\PERSONAL\Enzo Rodrigues\Default Files\M7 - ex -.xlsx"
        public static string PathToM7DOpen
        {
            get { return @"S:\Log_Planej_Adm\PERSONAL\Enzo Rodrigues\Default Files\M7 - STK ex.xlsx"; }
        }
  
        public static string PathToM7DModel
        {
            get { return @"C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\AbreModelo7 - Rev1.xlsm"; }
        }


        public static string PathToServer = @"S:\Log_Planej_Adm\CY Inventory Tracking\Relatório Estoque Geral\2023\M7 - STK 08-23\M7 - STK " + dateValidate + " -.xlsx";
        
        //api to return the convertion of dolar value
        public static async void getDollar()
        {
            using (HttpClient api = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
                string url = "https:/api.bcb.gov.br/dados/serie/bcdata.sgs.10813/dados/ultimos/1?formato=json";
                HttpResponseMessage response = await api.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    using (FileStream fileStream = File.Create("C:\\Users\\Enzo\\OneDrive\\Área de Trabalho\\Joyson\\local.html"))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                }
                else
                {
                    MessageBox.Show("Erro ao obter o valor do dólar");
                }
            }
        }
    }
}

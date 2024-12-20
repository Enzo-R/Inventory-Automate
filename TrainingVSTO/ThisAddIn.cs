﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using TrainingVSTO;
using Microsoft.Office.Interop.Excel;

namespace TrainingVSTO
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Models.Files.OpenM7Model();
            //Models.Excel.getDollar();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Models.Workbooks.ReleaseObject(getActiveApp());
            Models.Workbooks.ReleaseObject(getActiveWorkbook());
            Models.Workbooks.ReleaseObject(getActiveWorksheet());

        }

        public Excel.Worksheet getActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        public Excel.Workbook getActiveWorkbook()
        {
            return (Excel.Workbook)Application.ActiveWorkbook;
        }

        public Excel.Application getActiveApp()
        {
            return (Excel.Application)Application.Application;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
        }

        #endregion
    }
}

namespace TrainingVSTO
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.AddIns = this.Factory.CreateRibbonTab();
            this.Actions = this.Factory.CreateRibbonGroup();
            this.OpenFile = this.Factory.CreateRibbonButton();
            this.BtnAbre = this.Factory.CreateRibbonButton();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.AddIns.SuspendLayout();
            this.Actions.SuspendLayout();
            this.SuspendLayout();
            // 
            // AddIns
            // 
            this.AddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.AddIns.Groups.Add(this.Actions);
            this.AddIns.Label = "AddIns";
            this.AddIns.Name = "AddIns";
            // 
            // Actions
            // 
            this.Actions.Items.Add(this.OpenFile);
            this.Actions.Items.Add(this.BtnAbre);
            this.Actions.Items.Add(this.editBox1);
            this.Actions.Label = "Actions";
            this.Actions.Name = "Actions";
            // 
            // OpenFile
            // 
            this.OpenFile.Label = "Search";
            this.OpenFile.Name = "OpenFile";
            this.OpenFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenFile_Click);
            // 
            // BtnAbre
            // 
            this.BtnAbre.Label = "Open Model 7";
            this.BtnAbre.Name = "BtnAbre";
            this.BtnAbre.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AbreModeloClick);
            // 
            // editBox1
            // 
            this.editBox1.Label = "Day:";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.AddIns);
            this.AddIns.ResumeLayout(false);
            this.AddIns.PerformLayout();
            this.Actions.ResumeLayout(false);
            this.Actions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AddIns;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Actions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAbre;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

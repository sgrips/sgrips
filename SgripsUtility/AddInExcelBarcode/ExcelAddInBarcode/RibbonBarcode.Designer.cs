namespace ExcelAddInBarcode
{
    partial class BarcodeForExcel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public BarcodeForExcel()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BarcodeForExcel));
            this.Barcode = this.Factory.CreateRibbonTab();
            this.TabToolSgrips = this.Factory.CreateRibbonTab();
            this.Control = this.Factory.CreateRibbonGroup();
            this.btnBarcode = this.Factory.CreateRibbonButton();
            this.Barcode.SuspendLayout();
            this.TabToolSgrips.SuspendLayout();
            this.Control.SuspendLayout();
            this.SuspendLayout();
            // 
            // Barcode
            // 
            this.Barcode.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Barcode.ControlId.OfficeId = "Barcode";
            this.Barcode.Label = "Barcode";
            this.Barcode.Name = "Barcode";
            // 
            // TabToolSgrips
            // 
            this.TabToolSgrips.Groups.Add(this.Control);
            this.TabToolSgrips.Label = "Barcode";
            this.TabToolSgrips.Name = "TabToolSgrips";
            // 
            // Control
            // 
            this.Control.Items.Add(this.btnBarcode);
            this.Control.Label = "Tools";
            this.Control.Name = "Control";
            // 
            // btnBarcode
            // 
            this.btnBarcode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBarcode.Image = ((System.Drawing.Image)(resources.GetObject("btnBarcode.Image")));
            this.btnBarcode.Label = "Barcode";
            this.btnBarcode.Name = "btnBarcode";
            this.btnBarcode.ShowImage = true;
            this.btnBarcode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBarcode_Click);
            // 
            // BarcodeForExcel
            // 
            this.Name = "BarcodeForExcel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Barcode);
            this.Tabs.Add(this.TabToolSgrips);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonBarcode_Load);
            this.Barcode.ResumeLayout(false);
            this.Barcode.PerformLayout();
            this.TabToolSgrips.ResumeLayout(false);
            this.TabToolSgrips.PerformLayout();
            this.Control.ResumeLayout(false);
            this.Control.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Barcode;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabToolSgrips;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Control;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBarcode;
    }

    partial class ThisRibbonCollection
    {
        internal BarcodeForExcel RibbonBarcode
        {
            get { return this.GetRibbon<BarcodeForExcel>(); }
        }
    }
}

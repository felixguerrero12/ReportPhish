namespace PhishingReportAddin
{
    partial class PhishingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PhishingRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            System.Diagnostics.Debug.WriteLine("PhishingRibbon constructor called");
            InitializeComponent();
            this.tab1.Visible = true;  // Ensure tab visibility
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PhishingRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnReportPhishing = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Phishing Report";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.btnReportPhishing);
            this.group1.Label = "Actions";
            this.group1.Name = "group1";
            // 
            // btnReportPhishing
            // 
            this.btnReportPhishing.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReportPhishing.Image = ((System.Drawing.Image)(resources.GetObject("btnReportPhishing.Image")));
            this.btnReportPhishing.Label = "Report Phishing";
            this.btnReportPhishing.Name = "btnReportPhishing";
            this.btnReportPhishing.ScreenTip = "Report Phishing Email";
            this.btnReportPhishing.ShowImage = true;
            this.btnReportPhishing.SuperTip = "Forward this email to the security team as a potential phishing attempt";
            this.btnReportPhishing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReportPhishing_Click);
            // 
            // PhishingRibbon
            // 
            this.Name = "PhishingRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";

            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PhishingRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportPhishing;
    }

    partial class ThisRibbonCollection
    {
        internal PhishingRibbon PhishingRibbon
        {
            get { return this.GetRibbon<PhishingRibbon>(); }
        }
    }
}

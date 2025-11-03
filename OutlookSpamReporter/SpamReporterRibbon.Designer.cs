namespace OutlookSpamReporter
{
    partial class SpamReporterRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SpamReporterRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SpamReporterRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnReportSpam = this.Factory.CreateRibbonButton();
            this.tabRead = this.Factory.CreateRibbonTab();
            this.groupRead = this.Factory.CreateRibbonGroup();
            this.btnReportSpamRead = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.tabRead.SuspendLayout();
            this.groupRead.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnReportSpam);
            this.group1.Label = "Report phishing";
            this.group1.Name = "group1";
            // 
            // btnReportSpam
            // 
            this.btnReportSpam.Image = ((System.Drawing.Image)(resources.GetObject("btnReportSpam.Image")));
            this.btnReportSpam.Label = "Report phishing";
            this.btnReportSpam.Name = "btnReportSpam";
            this.btnReportSpam.ShowImage = true;
            this.btnReportSpam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReportSpam_Click);
            // 
            // tabRead
            // 
            this.tabRead.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabRead.ControlId.OfficeId = "TabReadMessage";
            this.tabRead.Groups.Add(this.groupRead);
            this.tabRead.Label = "TabReadMessage";
            this.tabRead.Name = "tabRead";
            // 
            // groupRead
            // 
            this.groupRead.Items.Add(this.btnReportSpamRead);
            this.groupRead.Label = "Spam Reporter";
            this.groupRead.Name = "groupRead";
            // 
            // btnReportSpamRead
            // 
            this.btnReportSpamRead.Image = ((System.Drawing.Image)(resources.GetObject("btnReportSpamRead.Image")));
            this.btnReportSpamRead.Label = "Report Spam";
            this.btnReportSpamRead.Name = "btnReportSpamRead";
            this.btnReportSpamRead.ShowImage = true;
            this.btnReportSpamRead.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReportSpamRead_Click);
            // 
            // SpamReporterRibbon
            // 
            this.Name = "SpamReporterRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabRead);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SpamReporterRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.tabRead.ResumeLayout(false);
            this.tabRead.PerformLayout();
            this.groupRead.ResumeLayout(false);
            this.groupRead.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportSpam;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabRead;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRead;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportSpamRead;
    }

    partial class ThisRibbonCollection
    {
        internal SpamReporterRibbon SpamReporterRibbon
        {
            get { return this.GetRibbon<SpamReporterRibbon>(); }
        }
    }
}

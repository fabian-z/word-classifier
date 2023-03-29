namespace word_classifier
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ClassifierGroup = this.Factory.CreateRibbonGroup();
            this.tlpHelp = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.tlpWhite = this.Factory.CreateRibbonButton();
            this.tlpGreen = this.Factory.CreateRibbonButton();
            this.tlpAmber = this.Factory.CreateRibbonButton();
            this.tlpRed = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ClassifierGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.ClassifierGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // ClassifierGroup
            // 
            this.ClassifierGroup.Items.Add(this.tlpHelp);
            this.ClassifierGroup.Items.Add(this.separator1);
            this.ClassifierGroup.Items.Add(this.tlpWhite);
            this.ClassifierGroup.Items.Add(this.tlpGreen);
            this.ClassifierGroup.Items.Add(this.tlpAmber);
            this.ClassifierGroup.Items.Add(this.tlpRed);
            this.ClassifierGroup.Label = "Classifier";
            this.ClassifierGroup.Name = "ClassifierGroup";
            // 
            // tlpHelp
            // 
            this.tlpHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tlpHelp.Image = global::word_classifier.Properties.Resources.icon_64;
            this.tlpHelp.Label = "TLP";
            this.tlpHelp.Name = "tlpHelp";
            this.tlpHelp.ShowImage = true;
            this.tlpHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tlpHelp_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // tlpWhite
            // 
            this.tlpWhite.Image = global::word_classifier.Properties.Resources.icon_32;
            this.tlpWhite.Label = "Mark white";
            this.tlpWhite.Name = "tlpWhite";
            this.tlpWhite.ShowImage = true;
            this.tlpWhite.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tlpWhite_Click);
            // 
            // tlpGreen
            // 
            this.tlpGreen.Image = global::word_classifier.Properties.Resources.icon_green_32;
            this.tlpGreen.Label = "Mark green";
            this.tlpGreen.Name = "tlpGreen";
            this.tlpGreen.ShowImage = true;
            this.tlpGreen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tlpGreen_Click);
            // 
            // tlpAmber
            // 
            this.tlpAmber.Image = global::word_classifier.Properties.Resources.icon_orange_32;
            this.tlpAmber.Label = "Mark amber";
            this.tlpAmber.Name = "tlpAmber";
            this.tlpAmber.ShowImage = true;
            this.tlpAmber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tlpAmber_Click);
            // 
            // tlpRed
            // 
            this.tlpRed.Image = global::word_classifier.Properties.Resources.icon_red_64;
            this.tlpRed.Label = "Mark red";
            this.tlpRed.Name = "tlpRed";
            this.tlpRed.ShowImage = true;
            this.tlpRed.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tlpRed_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ClassifierGroup.ResumeLayout(false);
            this.ClassifierGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ClassifierGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tlpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tlpWhite;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tlpGreen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tlpAmber;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tlpRed;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}

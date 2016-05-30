namespace PPTProgressMaker
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnHorizTOC = this.Factory.CreateRibbonButton();
            this.btnVertToc = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.cbGradient = this.Factory.CreateRibbonCheckBox();
            this.cbRTL = this.Factory.CreateRibbonCheckBox();
            this.box1 = this.Factory.CreateRibbonBox();
            this.glActive = this.Factory.CreateRibbonGallery();
            this.glNormal = this.Factory.CreateRibbonGallery();
            this.cbFirstSlide = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Ariel";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnHorizTOC);
            this.group1.Items.Add(this.btnVertToc);
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.box1);
            this.group1.Label = "Contents";
            this.group1.Name = "group1";
            // 
            // btnHorizTOC
            // 
            this.btnHorizTOC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHorizTOC.Image = global::PPTProgressMaker.Properties.Resources.Bottom;
            this.btnHorizTOC.Label = "Create Horizontal";
            this.btnHorizTOC.Name = "btnHorizTOC";
            this.btnHorizTOC.ShowImage = true;
            this.btnHorizTOC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHorizTOC_Click);
            // 
            // btnVertToc
            // 
            this.btnVertToc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnVertToc.Image = global::PPTProgressMaker.Properties.Resources.Side;
            this.btnVertToc.Label = "Create Vertical";
            this.btnVertToc.Name = "btnVertToc";
            this.btnVertToc.ShowImage = true;
            this.btnVertToc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVertToc_Click);
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.cbGradient);
            this.box2.Items.Add(this.cbRTL);
            this.box2.Items.Add(this.cbFirstSlide);
            this.box2.Name = "box2";
            // 
            // cbGradient
            // 
            this.cbGradient.Checked = true;
            this.cbGradient.Label = "Gradient";
            this.cbGradient.Name = "cbGradient";
            // 
            // cbRTL
            // 
            this.cbRTL.Label = "Right To Left";
            this.cbRTL.Name = "cbRTL";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.glActive);
            this.box1.Items.Add(this.glNormal);
            this.box1.Name = "box1";
            // 
            // glActive
            // 
            this.glActive.Label = "Active Color";
            this.glActive.Name = "glActive";
            this.glActive.ShowImage = true;
            this.glActive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.glActive_Click);
            // 
            // glNormal
            // 
            this.glNormal.Label = "Normal Color";
            this.glNormal.Name = "glNormal";
            this.glNormal.ShowImage = true;
            this.glNormal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.glNormal_Click);
            // 
            // cbFirstSlide
            // 
            this.cbFirstSlide.Label = "On first slide";
            this.cbFirstSlide.Name = "cbFirstSlide";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHorizTOC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVertToc;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbRTL;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glNormal;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glActive;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbGradient;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbFirstSlide;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}

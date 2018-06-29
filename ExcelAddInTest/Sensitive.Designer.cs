namespace ExcelAddInTest
{
    partial class Sensitive : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Sensitive()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Sensitive));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.menuClass = this.Factory.CreateRibbonMenu();
            this.toggleButtonSecret = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonConfidential = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonInternal = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonPublic = this.Factory.CreateRibbonToggleButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.splitButtonMark = this.Factory.CreateRibbonSplitButton();
            this.toggleButtonMarkYes = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonMarkNo = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.menuClass);
            this.group1.Label = "密级";
            this.group1.Name = "group1";
            // 
            // menuClass
            // 
            this.menuClass.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuClass.Description = "分类";
            this.menuClass.Image = ((System.Drawing.Image)(resources.GetObject("menuClass.Image")));
            this.menuClass.ImageName = "Classification";
            this.menuClass.Items.Add(this.toggleButtonSecret);
            this.menuClass.Items.Add(this.toggleButtonConfidential);
            this.menuClass.Items.Add(this.toggleButtonInternal);
            this.menuClass.Items.Add(this.toggleButtonPublic);
            this.menuClass.Items.Add(this.separator1);
            this.menuClass.Items.Add(this.splitButtonMark);
            this.menuClass.Label = "分类";
            this.menuClass.Name = "menuClass";
            this.menuClass.ShowImage = true;
            // 
            // toggleButtonSecret
            // 
            this.toggleButtonSecret.Label = "核心商密";
            this.toggleButtonSecret.Name = "toggleButtonSecret";
            this.toggleButtonSecret.ShowImage = true;
            this.toggleButtonSecret.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonSecret_Click);
            // 
            // toggleButtonConfidential
            // 
            this.toggleButtonConfidential.Label = "普通商迷";
            this.toggleButtonConfidential.Name = "toggleButtonConfidential";
            this.toggleButtonConfidential.ShowImage = true;
            this.toggleButtonConfidential.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonConfidential_Click);
            // 
            // toggleButtonInternal
            // 
            this.toggleButtonInternal.Label = "内部文件";
            this.toggleButtonInternal.Name = "toggleButtonInternal";
            this.toggleButtonInternal.ShowImage = true;
            this.toggleButtonInternal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonInternal_Click);
            // 
            // toggleButtonPublic
            // 
            this.toggleButtonPublic.Label = "公开文件";
            this.toggleButtonPublic.Name = "toggleButtonPublic";
            this.toggleButtonPublic.ShowImage = true;
            this.toggleButtonPublic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPublic_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // splitButtonMark
            // 
            this.splitButtonMark.Image = ((System.Drawing.Image)(resources.GetObject("splitButtonMark.Image")));
            this.splitButtonMark.Items.Add(this.toggleButtonMarkYes);
            this.splitButtonMark.Items.Add(this.toggleButtonMarkNo);
            this.splitButtonMark.Label = "首页标记";
            this.splitButtonMark.Name = "splitButtonMark";
            // 
            // toggleButtonMarkYes
            // 
            this.toggleButtonMarkYes.Label = "YES 标记";
            this.toggleButtonMarkYes.Name = "toggleButtonMarkYes";
            this.toggleButtonMarkYes.ShowImage = true;
            // 
            // toggleButtonMarkNo
            // 
            this.toggleButtonMarkNo.Label = "NO 标记";
            this.toggleButtonMarkNo.Name = "toggleButtonMarkNo";
            this.toggleButtonMarkNo.ShowImage = true;
            // 
            // Sensitive
            // 
            this.Name = "Sensitive";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Sensitive_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuClass;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonSecret;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonConfidential;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonInternal;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPublic;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonMark;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMarkYes;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMarkNo;
    }

    partial class ThisRibbonCollection
    {
        internal Sensitive Sensitive
        {
            get { return this.GetRibbon<Sensitive>(); }
        }
    }
}

﻿namespace PowerPointAddInConfidential
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.toggleButtonSecret = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonConfidential = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonInternal = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonPublic = this.Factory.CreateRibbonToggleButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.buttonClearTags = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.menu1);
            this.group1.Label = "密级";
            this.group1.Name = "group1";
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Items.Add(this.toggleButtonSecret);
            this.menu1.Items.Add(this.toggleButtonConfidential);
            this.menu1.Items.Add(this.toggleButtonInternal);
            this.menu1.Items.Add(this.toggleButtonPublic);
            this.menu1.Items.Add(this.separator1);
            this.menu1.Items.Add(this.buttonClearTags);
            this.menu1.Items.Add(this.splitButtonMark);
            this.menu1.Label = "分类";
            this.menu1.Name = "menu1";
            this.menu1.OfficeImageId = "FileMarkAsFinal";
            this.menu1.ShowImage = true;
            this.menu1.SuperTip = "对此Workbook进行密级标定，密级等级参考公司信息分类分级目录。";
            // 
            // toggleButtonSecret
            // 
            this.toggleButtonSecret.Label = "核心商密";
            this.toggleButtonSecret.Name = "toggleButtonSecret";
            this.toggleButtonSecret.ShowImage = true;
            this.toggleButtonSecret.SuperTip = "公司核心商业信息，参考公司分类分级目录。";
            this.toggleButtonSecret.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonSecret_Click);
            // 
            // toggleButtonConfidential
            // 
            this.toggleButtonConfidential.Label = "普通商密";
            this.toggleButtonConfidential.Name = "toggleButtonConfidential";
            this.toggleButtonConfidential.ShowImage = true;
            this.toggleButtonConfidential.SuperTip = "公司普通商业秘密信息，参考公司信息分类分级目录。";
            this.toggleButtonConfidential.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonConfidential_Click);
            // 
            // toggleButtonInternal
            // 
            this.toggleButtonInternal.Label = "内部文件";
            this.toggleButtonInternal.Name = "toggleButtonInternal";
            this.toggleButtonInternal.ShowImage = true;
            this.toggleButtonInternal.SuperTip = "公司内部信息，参考公司信息分类分级目录。";
            this.toggleButtonInternal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonInternal_Click);
            // 
            // toggleButtonPublic
            // 
            this.toggleButtonPublic.Label = "公开文件";
            this.toggleButtonPublic.Name = "toggleButtonPublic";
            this.toggleButtonPublic.ShowImage = true;
            this.toggleButtonPublic.SuperTip = "公司可公开披露信息，参考公司信息分类分级目录。";
            this.toggleButtonPublic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPublic_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // buttonClearTags
            // 
            this.buttonClearTags.Label = "清除标记";
            this.buttonClearTags.Name = "buttonClearTags";
            this.buttonClearTags.OfficeImageId = "MasterDocumentUnlinkSubdocument";
            this.buttonClearTags.ShowImage = true;
            this.buttonClearTags.SuperTip = "清除页眉标记图片";
            this.buttonClearTags.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonClearTags_Click);
            // 
            // splitButtonMark
            // 
            this.splitButtonMark.Items.Add(this.toggleButtonMarkYes);
            this.splitButtonMark.Items.Add(this.toggleButtonMarkNo);
            this.splitButtonMark.Label = "可视标记";
            this.splitButtonMark.Name = "splitButtonMark";
            this.splitButtonMark.OfficeImageId = "PostReplyToFolder";
            this.splitButtonMark.SuperTip = "是否在此Sheet页页眉插入图片标记。";
            // 
            // toggleButtonMarkYes
            // 
            this.toggleButtonMarkYes.Label = "YES 母版标记";
            this.toggleButtonMarkYes.Name = "toggleButtonMarkYes";
            this.toggleButtonMarkYes.ShowImage = true;
            this.toggleButtonMarkYes.SuperTip = "插入母版页标记图片。";
            this.toggleButtonMarkYes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonMarkYes_Click);
            // 
            // toggleButtonMarkNo
            // 
            this.toggleButtonMarkNo.Label = "NO 母版标记";
            this.toggleButtonMarkNo.Name = "toggleButtonMarkNo";
            this.toggleButtonMarkNo.ShowImage = true;
            this.toggleButtonMarkNo.SuperTip = "不插入母版页标记图片。";
            this.toggleButtonMarkNo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonMarkNo_Click);
            // 
            // Sensitive
            // 
            this.Name = "Sensitive";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonSecret;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonConfidential;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonInternal;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPublic;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonMark;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMarkYes;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMarkNo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClearTags;
    }

    partial class ThisRibbonCollection
    {
        internal Sensitive Sensitive
        {
            get { return this.GetRibbon<Sensitive>(); }
        }
    }
}

namespace Word_AddIns
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabAddinTool = this.Factory.CreateRibbonTab();
            this.grpTemplate = this.Factory.CreateRibbonGroup();
            this.btnNewDocument = this.Factory.CreateRibbonButton();
            this.btnSetDefault = this.Factory.CreateRibbonButton();
            this.grpImport = this.Factory.CreateRibbonGroup();
            this.btnImportMarkdown = this.Factory.CreateRibbonButton();
            this.btnImportText = this.Factory.CreateRibbonButton();
            this.grpDesign = this.Factory.CreateRibbonGroup();
            this.btnUniformFont = this.Factory.CreateRibbonButton();
            this.btnUniformTableStyle = this.Factory.CreateRibbonButton();
            this.btnRemoveStyles = this.Factory.CreateRibbonButton();
            this.btnApplyTemplate = this.Factory.CreateRibbonButton();
            this.mnuWatermark = this.Factory.CreateRibbonMenu();
            this.btnAddWatermark = this.Factory.CreateRibbonButton();
            this.btnModifyWatermark = this.Factory.CreateRibbonButton();
            this.btnRemoveWatermark = this.Factory.CreateRibbonButton();
            this.btnUpdateFields = this.Factory.CreateRibbonButton();
            this.grpExport = this.Factory.CreateRibbonGroup();
            this.btnExportPDF = this.Factory.CreateRibbonButton();
            this.grpMore = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnFeedback = this.Factory.CreateRibbonButton();
            this.tabAddinTool.SuspendLayout();
            this.grpTemplate.SuspendLayout();
            this.grpImport.SuspendLayout();
            this.grpDesign.SuspendLayout();
            this.grpExport.SuspendLayout();
            this.grpMore.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabAddinTool
            // 
            this.tabAddinTool.Groups.Add(this.grpTemplate);
            this.tabAddinTool.Groups.Add(this.grpImport);
            this.tabAddinTool.Groups.Add(this.grpDesign);
            this.tabAddinTool.Groups.Add(this.grpExport);
            this.tabAddinTool.Groups.Add(this.grpMore);
            this.tabAddinTool.Label = global::Word_AddIns.Properties.Resources.RibbonTab_Label;
            this.tabAddinTool.Name = "tabAddinTool";
            // 
            // grpTemplate
            // 
            this.grpTemplate.Items.Add(this.btnNewDocument);
            this.grpTemplate.Items.Add(this.btnSetDefault);
            this.grpTemplate.Label = global::Word_AddIns.Properties.Resources.TemplateGroup_Label;
            this.grpTemplate.Name = "grpTemplate";
            // 
            // btnNewDocument
            // 
            this.btnNewDocument.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewDocument.Image = global::Word_AddIns.Properties.Resources.btnNewDocument;
            this.btnNewDocument.Label = global::Word_AddIns.Properties.Resources.btnNewDocument_Label;
            this.btnNewDocument.Name = "btnNewDocument";
            this.btnNewDocument.ScreenTip = global::Word_AddIns.Properties.Resources.btnNewDocument_ScreenTip;
            this.btnNewDocument.ShowImage = true;
            this.btnNewDocument.SuperTip = global::Word_AddIns.Properties.Resources.btnNewDocument_SuperTip;
            this.btnNewDocument.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnNewDocument);
            // 
            // btnSetDefault
            // 
            this.btnSetDefault.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetDefault.Image = global::Word_AddIns.Properties.Resources.btnSetDefault;
            this.btnSetDefault.Label = global::Word_AddIns.Properties.Resources.btnSetDefault_Label;
            this.btnSetDefault.Name = "btnSetDefault";
            this.btnSetDefault.ScreenTip = global::Word_AddIns.Properties.Resources.btnSetDefault_ScreenTip;
            this.btnSetDefault.ShowImage = true;
            this.btnSetDefault.SuperTip = global::Word_AddIns.Properties.Resources.btnSetDefault_SuperTip;
            this.btnSetDefault.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnSetDefault);
            // 
            // grpImport
            // 
            this.grpImport.Items.Add(this.btnImportMarkdown);
            this.grpImport.Items.Add(this.btnImportText);
            this.grpImport.Label = global::Word_AddIns.Properties.Resources.ImportGroup_Label;
            this.grpImport.Name = "grpImport";
            // 
            // btnImportMarkdown
            // 
            this.btnImportMarkdown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportMarkdown.Image = global::Word_AddIns.Properties.Resources.btnImportMarkdown;
            this.btnImportMarkdown.Label = global::Word_AddIns.Properties.Resources.btnImportMarkdown_Label;
            this.btnImportMarkdown.Name = "btnImportMarkdown";
            this.btnImportMarkdown.ScreenTip = global::Word_AddIns.Properties.Resources.btnImportMarkdown_ScreenTip;
            this.btnImportMarkdown.ShowImage = true;
            this.btnImportMarkdown.SuperTip = global::Word_AddIns.Properties.Resources.btnImportMarkdown_SuperTip;
            this.btnImportMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportMarkdown);
            // 
            // btnImportText
            // 
            this.btnImportText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportText.Image = global::Word_AddIns.Properties.Resources.btnImportText;
            this.btnImportText.Label = global::Word_AddIns.Properties.Resources.btnImportText_Label;
            this.btnImportText.Name = "btnImportText";
            this.btnImportText.ScreenTip = global::Word_AddIns.Properties.Resources.btnImportText_ScreenTip;
            this.btnImportText.ShowImage = true;
            this.btnImportText.SuperTip = global::Word_AddIns.Properties.Resources.btnImportText_SuperTip;
            this.btnImportText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportText);
            // 
            // grpDesign
            // 
            this.grpDesign.Items.Add(this.btnUniformFont);
            this.grpDesign.Items.Add(this.btnUniformTableStyle);
            this.grpDesign.Items.Add(this.btnRemoveStyles);
            this.grpDesign.Items.Add(this.btnApplyTemplate);
            this.grpDesign.Items.Add(this.mnuWatermark);
            this.grpDesign.Items.Add(this.btnUpdateFields);
            this.grpDesign.Label = global::Word_AddIns.Properties.Resources.DesignGroup_Label;
            this.grpDesign.Name = "grpDesign";
            // 
            // btnUniformFont
            // 
            this.btnUniformFont.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUniformFont.Image = global::Word_AddIns.Properties.Resources.btnUniformFont;
            this.btnUniformFont.Label = global::Word_AddIns.Properties.Resources.btnUniformFont_Label;
            this.btnUniformFont.Name = "btnUniformFont";
            this.btnUniformFont.ScreenTip = global::Word_AddIns.Properties.Resources.btnUniformFont_ScreenTip;
            this.btnUniformFont.ShowImage = true;
            this.btnUniformFont.SuperTip = global::Word_AddIns.Properties.Resources.btnUniformFont_SuperTip;
            this.btnUniformFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnUniformFont);
            // 
            // btnUniformTableStyle
            // 
            this.btnUniformTableStyle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUniformTableStyle.Image = global::Word_AddIns.Properties.Resources.btnUniformTableStyle;
            this.btnUniformTableStyle.Label = global::Word_AddIns.Properties.Resources.btnUniformTableStyle_Label;
            this.btnUniformTableStyle.Name = "btnUniformTableStyle";
            this.btnUniformTableStyle.ScreenTip = global::Word_AddIns.Properties.Resources.btnUniformTableStyle_ScreenTip;
            this.btnUniformTableStyle.ShowImage = true;
            this.btnUniformTableStyle.SuperTip = global::Word_AddIns.Properties.Resources.btnUniformTableStyle_SuperTip;
            this.btnUniformTableStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnUniformTableStyle);
            // 
            // btnRemoveStyles
            // 
            this.btnRemoveStyles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRemoveStyles.Image = global::Word_AddIns.Properties.Resources.btnRemoveStyles;
            this.btnRemoveStyles.Label = global::Word_AddIns.Properties.Resources.btnRemoveStyles_Label;
            this.btnRemoveStyles.Name = "btnRemoveStyles";
            this.btnRemoveStyles.ScreenTip = global::Word_AddIns.Properties.Resources.btnRemoveStyles_ScreenTip;
            this.btnRemoveStyles.ShowImage = true;
            this.btnRemoveStyles.SuperTip = global::Word_AddIns.Properties.Resources.btnRemoveStyles_SuperTip;
            this.btnRemoveStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnRemoveStyles);
            // 
            // btnApplyTemplate
            // 
            this.btnApplyTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApplyTemplate.Image = global::Word_AddIns.Properties.Resources.btnSetDefault;
            this.btnApplyTemplate.Label = global::Word_AddIns.Properties.Resources.btnApplyTemplate_Label;
            this.btnApplyTemplate.Name = "btnApplyTemplate";
            this.btnApplyTemplate.ShowImage = true;
            this.btnApplyTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnApplyTemplate);
            // 
            // mnuWatermark
            // 
            this.mnuWatermark.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mnuWatermark.Image = global::Word_AddIns.Properties.Resources.mnuWatermark;
            this.mnuWatermark.Items.Add(this.btnAddWatermark);
            this.mnuWatermark.Items.Add(this.btnModifyWatermark);
            this.mnuWatermark.Items.Add(this.btnRemoveWatermark);
            this.mnuWatermark.Label = global::Word_AddIns.Properties.Resources.mnuWatermark_Label;
            this.mnuWatermark.Name = "mnuWatermark";
            this.mnuWatermark.ScreenTip = global::Word_AddIns.Properties.Resources.mnuWatermark_ScreenTip;
            this.mnuWatermark.ShowImage = true;
            this.mnuWatermark.SuperTip = global::Word_AddIns.Properties.Resources.mnuWatermark_SuperTip;
            // 
            // btnAddWatermark
            // 
            this.btnAddWatermark.Image = global::Word_AddIns.Properties.Resources.btnAddWatermark;
            this.btnAddWatermark.Label = global::Word_AddIns.Properties.Resources.btnAddWatermark_Label;
            this.btnAddWatermark.Name = "btnAddWatermark";
            this.btnAddWatermark.ShowImage = true;
            this.btnAddWatermark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnAddWatermark);
            // 
            // btnModifyWatermark
            // 
            this.btnModifyWatermark.Image = global::Word_AddIns.Properties.Resources.btnModifyWatermark;
            this.btnModifyWatermark.Label = global::Word_AddIns.Properties.Resources.btnModifyWatermark_Label;
            this.btnModifyWatermark.Name = "btnModifyWatermark";
            this.btnModifyWatermark.ShowImage = true;
            this.btnModifyWatermark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnModifyWatermark);
            // 
            // btnRemoveWatermark
            // 
            this.btnRemoveWatermark.Image = global::Word_AddIns.Properties.Resources.btnRemoveWatermark;
            this.btnRemoveWatermark.Label = global::Word_AddIns.Properties.Resources.btnRemoveWatermark_Label;
            this.btnRemoveWatermark.Name = "btnRemoveWatermark";
            this.btnRemoveWatermark.ShowImage = true;
            this.btnRemoveWatermark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnRemoveWatermark);
            // 
            // btnUpdateFields
            // 
            this.btnUpdateFields.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateFields.Image = global::Word_AddIns.Properties.Resources.btnUpdateFields;
            this.btnUpdateFields.Label = global::Word_AddIns.Properties.Resources.btnUpdateFields_Label;
            this.btnUpdateFields.Name = "btnUpdateFields";
            this.btnUpdateFields.ScreenTip = global::Word_AddIns.Properties.Resources.btnUpdateFields_ScreenTip;
            this.btnUpdateFields.ShowImage = true;
            this.btnUpdateFields.SuperTip = global::Word_AddIns.Properties.Resources.btnUpdateFields_SuperTip;
            this.btnUpdateFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnUpdateFields);
            // 
            // grpExport
            // 
            this.grpExport.Items.Add(this.btnExportPDF);
            this.grpExport.Label = global::Word_AddIns.Properties.Resources.ExportGroup_Label;
            this.grpExport.Name = "grpExport";
            // 
            // btnExportPDF
            // 
            this.btnExportPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportPDF.Image = global::Word_AddIns.Properties.Resources.btnExportPDF;
            this.btnExportPDF.Label = global::Word_AddIns.Properties.Resources.btnExportPDF_Label;
            this.btnExportPDF.Name = "btnExportPDF";
            this.btnExportPDF.ScreenTip = global::Word_AddIns.Properties.Resources.btnExportPDF_ScreenTip;
            this.btnExportPDF.ShowImage = true;
            this.btnExportPDF.SuperTip = global::Word_AddIns.Properties.Resources.btnExportPDF_SuperTip;
            this.btnExportPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnExportPDF);
            // 
            // grpMore
            // 
            this.grpMore.Items.Add(this.btnAbout);
            this.grpMore.Items.Add(this.btnFeedback);
            this.grpMore.Label = global::Word_AddIns.Properties.Resources.MoreGroup_Label;
            this.grpMore.Name = "grpMore";
            // 
            // btnAbout
            // 
            this.btnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAbout.Image = global::Word_AddIns.Properties.Resources.btnAbout;
            this.btnAbout.Label = global::Word_AddIns.Properties.Resources.btnAbout_Label;
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ScreenTip = global::Word_AddIns.Properties.Resources.btnAbout_ScreenTip;
            this.btnAbout.ShowImage = true;
            this.btnAbout.SuperTip = global::Word_AddIns.Properties.Resources.btnAbout_SuperTip;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnAbout);
            // 
            // btnFeedback
            // 
            this.btnFeedback.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFeedback.Image = global::Word_AddIns.Properties.Resources.btnFeedback;
            this.btnFeedback.Label = global::Word_AddIns.Properties.Resources.btnFeedback_Label;
            this.btnFeedback.Name = "btnFeedback";
            this.btnFeedback.ScreenTip = global::Word_AddIns.Properties.Resources.btnFeedback_ScreenTip;
            this.btnFeedback.ShowImage = true;
            this.btnFeedback.SuperTip = global::Word_AddIns.Properties.Resources.btnFeedback_SuperTip;
            this.btnFeedback.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnFeedback);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabAddinTool);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabAddinTool.ResumeLayout(false);
            this.tabAddinTool.PerformLayout();
            this.grpTemplate.ResumeLayout(false);
            this.grpTemplate.PerformLayout();
            this.grpImport.ResumeLayout(false);
            this.grpImport.PerformLayout();
            this.grpDesign.ResumeLayout(false);
            this.grpDesign.PerformLayout();
            this.grpExport.ResumeLayout(false);
            this.grpExport.PerformLayout();
            this.grpMore.ResumeLayout(false);
            this.grpMore.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAddinTool;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDesign;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUniformFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUniformTableStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuWatermark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddWatermark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifyWatermark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveWatermark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewDocument;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetDefault;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMore;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFeedback;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

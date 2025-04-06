namespace YaGPTforExcel
{
    partial class YaGPT4ExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup groupGPT;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerate;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnTogglePanel;


        private void InitializeComponent()
        {
            this.tabMain = this.Factory.CreateRibbonTab();
            this.groupGPT = this.Factory.CreateRibbonGroup();
            this.btnGenerate = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.btnTogglePanel = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.groupGPT.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMain.Groups.Add(this.groupGPT);
            this.tabMain.Label = "YaGPT";
            this.tabMain.Name = "tabMain";
            // 
            // groupGPT
            // 
            this.groupGPT.Items.Add(this.btnGenerate);
            this.groupGPT.Items.Add(this.btnSettings);
            this.groupGPT.Label = "Yandex GPT";
            this.groupGPT.Name = "groupGPT";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Label = "Сгенерировать текст";
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerate_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.Label = "Настройки API";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            ///
            this.btnTogglePanel = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();

            // 
            // btnTogglePanel
            // 
            this.btnTogglePanel.Label = "Переключить панель";
            this.btnTogglePanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePanel_Click);

            // 
            // add to tabMain
            // 
            this.tabMain.Groups.Add(this.group1);           

            // 
            // group1
            // 
            this.group1.Items.Add(this.btnTogglePanel);
            this.group1.Label = "Управление";

            // 
            // YaGPT4ExcelRibbon
            // 
            this.Name = "YaGPT4ExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.YaGPT4ExcelRibbon_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.groupGPT.ResumeLayout(false);
            this.groupGPT.PerformLayout();
            this.ResumeLayout(false);

        }

        #region Component Designer generated code

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private System.ComponentModel.IContainer components = null;

        #endregion
    }
}

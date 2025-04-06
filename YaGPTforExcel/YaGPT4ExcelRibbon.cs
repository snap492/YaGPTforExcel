// YaGPT4ExcelRibbon.cs — VSTO надстройка для Excel

using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using YaGPTforExcel.Services;

namespace YaGPTforExcel
{
    public partial class YaGPT4ExcelRibbon : RibbonBase
    {
        public YaGPT4ExcelRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }
        private void YaGPT4ExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            btnGenerate.Enabled = true;
            btnGenerate.Label = "Сгенерировать текст";
        }
        private async void btnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string token = Properties.Settings.Default.Token;
                string folderId = Properties.Settings.Default.FolderId;
                var service = new Yagpt4excelService(token, folderId);
                string prompt = GetSelectedRangeAsPrompt();
                if (string.IsNullOrEmpty(prompt))
                    return;
                string result = await service.GenerateText(prompt);

                Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
                activeCell.Value2 = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            var settingsForm = new SettingsForm();
            settingsForm.ShowDialog();
        }

        private void btnTogglePanel_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane.Visible)
            {
                Globals.ThisAddIn.TaskPane.Visible = false; // скрываем панель
            }
            else
            {
                Globals.ThisAddIn.TaskPane.Visible = true; // показываем панель
            }
        }


        /// <summary>
        /// Преобразует выделенный диапазон Excel в текстовую строку для prompt
        /// </summary>
        /// <returns>Строка с содержимым выделенной таблицы</returns>
        private string GetSelectedRangeAsPrompt()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;

            if (selection == null || selection.Cells.Count == 0)
            {
                MessageBox.Show("Пожалуйста, выделите диапазон ячеек в Excel.");
                return null;
            }

            string selectedText = "";
            foreach (Excel.Range row in selection.Rows)
            {
                string rowText = "";
                foreach (Excel.Range cell in row.Columns)
                {
                    rowText += $"{cell.Text}\t"; // табуляция между ячейками
                }
                selectedText += rowText.TrimEnd('\t') + "\n";
            }

            return $"Проанализируй следующие данные:\n{selectedText}\nСделай выводы или рекомендации.";
        }

    }

    partial class ThisRibbonCollection
    {
        internal YaGPT4ExcelRibbon YaGPT4ExcelRibbon
        {
            get { return this.GetRibbon<YaGPT4ExcelRibbon>(); }
        }
    }
}
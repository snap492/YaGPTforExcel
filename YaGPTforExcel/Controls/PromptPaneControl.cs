using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using YaGPTforExcel.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Controls
{
    public partial class PromptPaneControl : UserControl
    {
        public PromptPaneControl()
        {
            InitializeComponent();
        }

        private async void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                string userPrompt = txtPrompt.Text.Trim();
                if (string.IsNullOrEmpty(userPrompt))
                {
                    MessageBox.Show("Пожалуйста, введите текст запроса.");
                    return;
                }

                userPrompt = PromptBuilderService.BuildPrompt(userPrompt);

                var service = new Yagpt4excelService(Properties.Settings.Default.Token, Properties.Settings.Default.FolderId);
                var result = await service.GenerateText(userPrompt);

                txtResult.Invoke((Action)(() =>
                {
                    if (txtResult.Text == "здесь будет отображаться ответ от Yandex GPT")
                        txtResult.Text = "";

                    txtResult.Text = result;
                    txtResult.SelectionStart = txtResult.Text.Length;
                }));

                if (chkInsert.Checked)
                {
                    if (ExcelInsertService.IsSelectedRange())
                    {
                        ExcelInsertService.InsertIntoSelectedRange(result);
                    }
                    else
                    {
                        ExcelInsertService.InsertIntoActiveSheet(result);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using YaGPTforExcel.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel
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

                // Получаем данные из формы
                string userPrompt = txtPrompt.Text.Trim();
                if (string.IsNullOrEmpty(userPrompt))
                {
                    MessageBox.Show("Пожалуйста, введите текст запроса.");
                    return;
                }
                //Проверяем есть ли выделенная таблица и добавляем ее в promt
                userPrompt = TryAddMarkdownIfRelevant(userPrompt);

                // Создаем сервис
                var service = new Yagpt4excelService(Properties.Settings.Default.Token, Properties.Settings.Default.FolderId);

                // Генерируем результат
                var result = await service.GenerateText(userPrompt);

                // Выводим результат в TextBox
                txtResult.Invoke((Action)(() =>
                {
                    if (txtResult.Text == "здесь будет отображаться ответ от Yandex GPT")
                        txtResult.Text = "";

                    txtResult.Text = result;
                    txtResult.SelectionStart = txtResult.Text.Length;
                }));

                // Вставка в Excel
                if (chkInsert.Checked)
                {
                    if (IsSelectedRange())
                    {
                        // Если выделен диапазон, вставляем в него
                        InsertIntoSelectedRange(result);
                    }
                    else
                    {
                        // Если не выделен диапазон, вставляем в активный лист
                     InsertIntoActiveSheat(result);
                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        // Преобразование ответа в таблицу
        private List<List<string>> ParseResponseToTable(string response)
        {
            var tableData = new List<List<string>>();

            if (string.IsNullOrWhiteSpace(response))
                return tableData;

            var lines = response.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            bool isTableStarted = false;

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                // Начало таблицы: строка должна начинаться с '|'
                if (!isTableStarted)
                {
                    if (trimmedLine.StartsWith("|"))
                    {
                        isTableStarted = true;
                    }
                    else
                    {
                        continue; // пока таблица не началась — пропускаем строки
                    }
                }

                // Пропускаем строки типа "|---|---|"
                if (trimmedLine.StartsWith("|") && trimmedLine.Replace("|", "").Trim().StartsWith("-"))
                {
                    continue;
                }

                // Парсим строку таблицы
                if (trimmedLine.StartsWith("|"))
                {
                    var cells = trimmedLine
                        .Trim('|')                      // убираем крайние |
                        .Split('|')                     // разделяем по |
                        .Select(cell => cell.Trim())    // убираем пробелы
                        .ToList();

                    tableData.Add(cells);
                }
            }

            return tableData;
        }

        // Вставка данных таблицы в Excel
        private void InsertTableIntoExcel(List<List<string>> tableData, Excel.Range startCell)
        {
            // Индексация ячеек начинается с 1, так что используем смещения от первой ячейки
            int row = 0;
            foreach (var tableRow in tableData)
            {
                if (tableRow == null) continue; // Пропускаем пустые строки
                int col = 0;
                foreach (var cellValue in tableRow)
                {
                    // Убедитесь, что cellValue не null, чтобы избежать ошибок при вставке
                    if (cellValue == null) MessageBox.Show(cellValue);

                    // Вставляем данные в ячейку с использованием Offset
                    startCell.Offset[row, col].Value = cellValue;
                    col++;
                }
                row++;
            }
        }
        private void InsertIntoActiveSheat(string result)
        {
            // Получаем ссылку на активное приложение Excel в контексте надстройки
            var excelApp = Globals.ThisAddIn.Application;
            var activeWorkbook = excelApp.ActiveWorkbook;
            var activeSheet = (Excel.Worksheet)activeWorkbook.ActiveSheet;

            Excel.Range startCell;

            // Проверяем, пустой ли лист (используем UsedRange
            bool isSheetEmpty = activeSheet.UsedRange == null || activeSheet.UsedRange.Cells.Count == 1 && string.IsNullOrEmpty(Convert.ToString(activeSheet.UsedRange.Value2));

            if (isSheetEmpty)
            {
                startCell = activeSheet.Cells[1, 1];
            }
            else
            {
                startCell = excelApp.ActiveCell;
            }

            // Преобразуем ответ в таблицу
            var tableData = ParseResponseToTable(result);

            // Вставляем таблицу в Excel
            InsertTableIntoExcel(tableData, startCell);

            // Применяем автоширину ко всем столбцам в диапазоне
            var usedRange = activeSheet.UsedRange;
            usedRange.Columns.AutoFit();
        }
        private void InsertIntoSelectedRange(string result)
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection as Excel.Range;
            var activeSheet = app.ActiveSheet as Excel.Worksheet;

            if (activeSheet == null)
            {
                MessageBox.Show("Не удалось получить активный лист.");
                return;
            }

            var tableData = ParseResponseToTable(result);
            if (tableData == null || tableData.Count == 0)
            {
                MessageBox.Show("Ответ не содержит данных в виде таблицы.");
                return;
            }

            Excel.Range targetStartCell;

            if (selection != null && selection.Rows.Count > 0 && selection.Columns.Count > 0)
            {
                // Ищем первую пустую строку в выделении
                int insertStartRow = 1;
                for (int i = 1; i <= selection.Rows.Count; i++)
                {
                    var cell = (Excel.Range)selection.Cells[i, 1];
                    if (cell.Value2 == null || string.IsNullOrWhiteSpace(cell.Text.ToString()))
                    {
                        insertStartRow = i;
                        break;
                    }

                    if (i == selection.Rows.Count)
                    {
                        insertStartRow = selection.Rows.Count + 1;
                    }
                }

                targetStartCell = selection.Cells[insertStartRow, 1];
            }
            else
            {
                // Если нет выделения — ищем первую пустую строку на листе
                int lastRow = activeSheet.UsedRange.Rows.Count;
                targetStartCell = activeSheet.Cells[lastRow + 1, 1];
            }

            // Вставляем данные
            int rowOffset = 0;
            foreach (var row in tableData)
            {
                int colOffset = 0;
                foreach (var cell in row)
                {
                    targetStartCell.Offset[rowOffset, colOffset].Value2 = cell;
                    colOffset++;
                }
                rowOffset++;
            }

            // Применяем автоширину ко всем столбцам в диапазоне
            var usedRange = activeSheet.UsedRange;
            usedRange.Columns.AutoFit();
        }
        private bool IsSelectedRange()
        {
            // Получаем выделенный диапазон в Excel
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            if (selection != null)
            {
                // Проверяем, что выделение не пустое
                if (selection.Cells.Count > 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Не удалось получить выделение в Excel.");
                return false;
            }
        }
        private string GetSelectedRangeTextAsMarkdown()
        {
            try
            {
                var excelApp = Globals.ThisAddIn.Application;
                var selection = excelApp.Selection as Excel.Range;

                if (selection != null && selection.Cells.Count > 0)
                {
                    int rows = selection.Rows.Count;
                    int cols = selection.Columns.Count;
                    var markdown = "";

                    // Считываем таблицу в массив
                    string[,] table = new string[rows, cols];

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            var cell = selection.Cells[r, c] as Excel.Range;
                            table[r - 1, c - 1] = cell?.Text?.ToString() ?? "";
                        }
                    }

                    // Заголовок
                    for (int c = 0; c < cols; c++)
                        markdown += $"| {table[0, c]} ";
                    markdown += "|\n";

                    // Разделитель
                    for (int c = 0; c < cols; c++)
                        markdown += "|---";
                    markdown += "|\n";

                    // Остальные строки
                    for (int r = 1; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                            markdown += $"| {table[r, c]} ";
                        markdown += "|\n";
                    }

                    return markdown.Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Markdown: {ex.Message}");
            }

            return "";
        }
        private string TryAddMarkdownIfRelevant(string userPrompt)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var selection = excelApp.Selection as Excel.Range;

            // Проверка ключевых слов в prompt
            var lowered = userPrompt.ToLower();
            bool mentionsTable = lowered.Contains("таблица") ||
                                 lowered.Contains("анализируй") ||
                                 lowered.Contains("выделение") ||
                                 lowered.Contains("данные") ||
                                 lowered.Contains("найди в таблице") ||
                                 lowered.Contains("по таблице");
                                 lowered.Contains("добавь");
                                 lowered.Contains("добави");

            // Если есть выделение и запрос к таблице
            if (selection != null && selection.Cells.Count > 1 && mentionsTable)
            {
                string markdown = GetSelectedRangeTextAsMarkdown();
                if (!string.IsNullOrEmpty(markdown))
                {
                    return $"Вот выделенная таблица в формате Markdown:\n\n{markdown}\n\n{userPrompt}";
                }
            }

            return userPrompt;
        }

    }
}

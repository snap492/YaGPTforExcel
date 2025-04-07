using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Services
{
    public static class PromptBuilderService
    {
        public static string TryAddMarkdownIfRelevant(string userPrompt)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var selection = excelApp.Selection as Excel.Range;

            var lowered = userPrompt.ToLower();
            bool mentionsTable = lowered.Contains("таблица") ||
                                 lowered.Contains("анализируй") ||
                                 lowered.Contains("выделение") ||
                                 lowered.Contains("данные") ||
                                 lowered.Contains("найди в таблице") ||
                                 lowered.Contains("найди кто") ||
                                 lowered.Contains("по таблице") ||
                                 lowered.Contains("определи") ||
                                 lowered.Contains("добавь") ||
                                 lowered.Contains("добави");

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

        private static string GetSelectedRangeTextAsMarkdown()
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

                    string[,] table = new string[rows, cols];

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            var cell = selection.Cells[r, c] as Excel.Range;
                            table[r - 1, c - 1] = cell?.Text?.ToString() ?? "";
                        }
                    }

                    markdown += GenerateMarkdownHeader(table, cols);
                    markdown += GenerateMarkdownSeparator(cols);
                    markdown += GenerateMarkdownRows(table, rows, cols);

                    return markdown.Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании Markdown: {ex.Message}");
            }

            return "";
        }

        private static string GenerateMarkdownHeader(string[,] table, int cols)
        {
            var header = "";
            for (int c = 0; c < cols; c++)
                header += $"| {table[0, c]} ";
            header += "|\n";
            return header;
        }

        private static string GenerateMarkdownSeparator(int cols)
        {
            var separator = "";
            for (int c = 0; c < cols; c++)
                separator += "|---";
            separator += "|\n";
            return separator;
        }

        private static string GenerateMarkdownRows(string[,] table, int rows, int cols)
        {
            var rowsMarkdown = "";
            for (int r = 1; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                    rowsMarkdown += $"| {table[r, c]} ";
                rowsMarkdown += "|\n";
            }
            return rowsMarkdown;
        }
    }
}

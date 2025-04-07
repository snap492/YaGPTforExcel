using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Utils
{
    public static class MarkdownGenerator
    {
        /// <summary>
        /// ����������� ���������� �������� Excel � Markdown-�������.
        /// </summary>
        /// <returns>������ � ������� Markdown</returns>
        public static string GetSelectedRangeTextAsMarkdown(Excel.Range selection)
        {
            try
            {
                if (selection != null && selection.Cells.Count > 0)
                {
                    int rows = selection.Rows.Count;
                    int cols = selection.Columns.Count;
                    var markdown = "";

                    string[,] table = new string[rows, cols];

                    // ��������� �������� ����� � ��������� ������
                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            var cell = selection.Cells[r, c] as Excel.Range;
                            table[r - 1, c - 1] = cell?.Text?.ToString() ?? "";
                        }
                    }

                    // ���������� ������� � ������� Markdown
                    markdown += GenerateMarkdownHeader(table, cols);
                    markdown += GenerateMarkdownSeparator(cols);
                    markdown += GenerateMarkdownRows(table, rows, cols);

                    return markdown.Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� ������������ Markdown: {ex.Message}");
            }

            return "";
        }

        /// <summary>
        /// ��������� ��������� ������� Markdown �� ������ ������ ���������.
        /// </summary>
        private static string GenerateMarkdownHeader(string[,] table, int cols)
        {
            var header = "";
            for (int c = 0; c < cols; c++)
                header += $"| {table[0, c]} ";
            header += "|\n";
            return header;
        }

        /// <summary>
        /// ���������� ������-����������� ����� ���������� � ����� ������� Markdown.
        /// </summary>
        private static string GenerateMarkdownSeparator(int cols)
        {
            var separator = "";
            for (int c = 0; c < cols; c++)
                separator += "|---";
            separator += "|\n";
            return separator;
        }

        /// <summary>
        /// ��������� ������ ������� Markdown, ������� �� ������ ������ (������).
        /// </summary>
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

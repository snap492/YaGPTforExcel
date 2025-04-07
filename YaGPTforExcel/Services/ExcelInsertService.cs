using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Services
{
    public static class ExcelInsertService
    {
        /// <summary>
        /// Вставляет таблицу из текста в активный лист Excel.
        /// Если лист пустой — вставка начинается с ячейки A1, иначе — с активной ячейки.
        /// </summary>
        /// <param name="result">Результат в формате Markdown-таблицы</param>
        public static void InsertIntoActiveSheet(string result)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var activeWorkbook = excelApp.ActiveWorkbook;
            var activeSheet = (Excel.Worksheet)activeWorkbook.ActiveSheet;

            Excel.Range startCell = GetStartCell(activeSheet);

            var tableData = ParseResponseToTable(result);
            InsertTableIntoExcel(tableData, startCell);

            var usedRange = activeSheet.UsedRange;
            usedRange.Columns.AutoFit();
        }

        /// <summary>
        /// Вставляет таблицу в выделенный диапазон Excel.
        /// Если в диапазоне уже есть данные, вставка начинается с первой пустой строки.
        /// </summary>
        /// <param name="result">Результат в формате Markdown-таблицы</param>
        public static void InsertIntoSelectedRange(string result)
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

            Excel.Range targetStartCell = GetTargetStartCell(selection, activeSheet);
            InsertTableData(tableData, targetStartCell);

            var usedRange = activeSheet.UsedRange;
            usedRange.Columns.AutoFit();
        }

        /// <summary>
        /// Проверяет, выбрал ли пользователь диапазон из нескольких ячеек.
        /// </summary>
        public static bool IsSelectedRange()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            return selection != null && selection.Cells.Count > 1;
        }

        /// <summary>
        /// Определяет начальную ячейку для вставки в зависимости от того, пустой лист или нет.
        /// </summary>
        private static Excel.Range GetStartCell(Excel.Worksheet activeSheet)
        {
            var excelApp = Globals.ThisAddIn.Application;
            bool isSheetEmpty = activeSheet.UsedRange == null
                                || activeSheet.UsedRange.Cells.Count == 1
                                && string.IsNullOrEmpty(Convert.ToString(activeSheet.UsedRange.Value2));

            return isSheetEmpty ? activeSheet.Cells[1, 1] : excelApp.ActiveCell;
        }

        /// <summary>
        /// Вычисляет начальную ячейку для вставки в выделенном диапазоне.
        /// Если есть свободные строки — выбирает первую пустую строку.
        /// Иначе вставляет после последней строки диапазона.
        /// </summary>
        private static Excel.Range GetTargetStartCell(Excel.Range selection, Excel.Worksheet activeSheet)
        {
            if (selection != null && selection.Rows.Count > 0 && selection.Columns.Count > 0)
            {
                int insertStartRow = GetInsertStartRow(selection);
                return selection.Cells[insertStartRow, 1];
            }
            else
            {
                int lastRow = activeSheet.UsedRange.Rows.Count;
                return activeSheet.Cells[lastRow + 1, 1];
            }
        }

        /// <summary>
        /// Определяет, с какой строки начинать вставку в пределах выделенного диапазона.
        /// Возвращает индекс первой пустой строки, либо следующей за последней заполненной.
        /// </summary>
        private static int GetInsertStartRow(Excel.Range selection)
        {
            for (int i = 1; i <= selection.Rows.Count; i++)
            {
                var cell = (Excel.Range)selection.Cells[i, 1];
                if (cell.Value2 == null || string.IsNullOrWhiteSpace(cell.Text.ToString()))
                {
                    return i;
                }

                if (i == selection.Rows.Count)
                {
                    return selection.Rows.Count + 1;
                }
            }
            return 1;
        }

        /// <summary>
        /// Вставляет данные таблицы в Excel, начиная с заданной ячейки.
        /// </summary>
        private static void InsertTableData(List<List<string>> tableData, Excel.Range targetStartCell)
        {
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
        }

        /// <summary>
        /// Преобразует текст в формате Markdown-таблицы в список строк-ячееек.
        /// Игнорирует строки-разделители "---" и текст до первой строки таблицы.
        /// </summary>
        private static List<List<string>> ParseResponseToTable(string response)
        {
            var tableData = new List<List<string>>();

            if (string.IsNullOrWhiteSpace(response))
                return tableData;

            var lines = response.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            bool isTableStarted = false;

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                if (!isTableStarted)
                {
                    if (trimmedLine.StartsWith("|"))
                    {
                        isTableStarted = true;
                    }
                    else
                    {
                        continue;
                    }
                }

                if (trimmedLine.StartsWith("|") && trimmedLine.Replace("|", "").Trim().StartsWith("-"))
                {
                    continue;
                }

                if (trimmedLine.StartsWith("|"))
                {
                    var cells = trimmedLine
                        .Trim('|')
                        .Split('|')
                        .Select(cell => cell.Trim())
                        .ToList();

                    tableData.Add(cells);
                }
            }

            return tableData;
        }

        /// <summary>
        /// Вставляет таблицу в Excel, начиная с указанной ячейки.
        /// Используется для вставки с активного листа.
        /// </summary>
        private static void InsertTableIntoExcel(List<List<string>> tableData, Excel.Range startCell)
        {
            int row = 0;
            foreach (var tableRow in tableData)
            {
                if (tableRow == null) continue;
                int col = 0;
                foreach (var cellValue in tableRow)
                {
                    if (cellValue == null) MessageBox.Show(cellValue);
                    startCell.Offset[row, col].Value = cellValue;
                    col++;
                }
                row++;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Services
{
    public static class ExcelInsertService

    {
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

        public static void InsertIntoSelectedRange(string result)
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection as Excel.Range;
            var activeSheet = app.ActiveSheet as Excel.Worksheet;

            if (activeSheet == null)
            {
                MessageBox.Show("Ќе удалось получить активный лист.");
                return;
            }

            var tableData = ParseResponseToTable(result);
            if (tableData == null || tableData.Count == 0)
            {
                MessageBox.Show("ќтвет не содержит данных в виде таблицы.");
                return;
            }

            Excel.Range targetStartCell = GetTargetStartCell(selection, activeSheet);
            InsertTableData(tableData, targetStartCell);

            var usedRange = activeSheet.UsedRange;
            usedRange.Columns.AutoFit();
        }

        public static bool IsSelectedRange()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            return selection != null && selection.Cells.Count > 1;
        }

        private static Excel.Range GetStartCell(Excel.Worksheet activeSheet)
        {
            var excelApp = Globals.ThisAddIn.Application;
            bool isSheetEmpty = activeSheet.UsedRange == null || activeSheet.UsedRange.Cells.Count == 1 && string.IsNullOrEmpty(Convert.ToString(activeSheet.UsedRange.Value2));

            return isSheetEmpty ? activeSheet.Cells[1, 1] : excelApp.ActiveCell;
        }

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

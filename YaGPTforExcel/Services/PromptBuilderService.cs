using System;
using System.Windows.Forms;
using YaGPTforExcel.Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Services
{
    public static class PromptBuilderService
    {
        /// <summary>
        /// ѕровер€ет, содержит ли пользовательский prompt запрос, св€занный с таблицей,
        /// и если выделен диапазон в Excel, автоматически добавл€ет его содержимое в формате Markdown.
        /// </summary>
        /// <param name="userPrompt">»сходный запрос пользовател€</param>
        /// <returns>«апрос с добавленной таблицей (если применимо)</returns>
        public static string BuildPrompt(string userPrompt)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var selection = excelApp.Selection as Excel.Range;

            // ѕриводим запрос к нижнему регистру дл€ поиска ключевых слов
            var lowered = userPrompt.ToLower();

            // ѕроверка на наличие "табличных" фраз
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

            // ≈сли выделена таблица и запрос св€зан с таблицами
            if (selection != null && selection.Cells.Count > 1 && mentionsTable)
            {
                string markdown = MarkdownGenerator.GetSelectedRangeTextAsMarkdown(selection);
                if (!string.IsNullOrEmpty(markdown))
                {
                    return $"¬от выделенна€ таблица в формате Markdown:\n\n{markdown}\n\n{userPrompt}";
                }
            }

            return userPrompt;
        }
    }
}

using System;
using System.Windows.Forms;
using YaGPTforExcel.Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace YaGPTforExcel.Services
{
    public static class PromptBuilderService
    {
        /// <summary>
        /// ���������, �������� �� ���������������� prompt ������, ��������� � ��������,
        /// � ���� ������� �������� � Excel, ������������� ��������� ��� ���������� � ������� Markdown.
        /// </summary>
        /// <param name="userPrompt">�������� ������ ������������</param>
        /// <returns>������ � ����������� �������� (���� ���������)</returns>
        public static string BuildPrompt(string userPrompt)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var selection = excelApp.Selection as Excel.Range;

            // �������� ������ � ������� �������� ��� ������ �������� ����
            var lowered = userPrompt.ToLower();

            // �������� �� ������� "���������" ����
            bool mentionsTable = lowered.Contains("�������") ||
                                 lowered.Contains("����������") ||
                                 lowered.Contains("���������") ||
                                 lowered.Contains("������") ||
                                 lowered.Contains("����� � �������") ||
                                 lowered.Contains("����� ���") ||
                                 lowered.Contains("�� �������") ||
                                 lowered.Contains("��������") ||
                                 lowered.Contains("������") ||
                                 lowered.Contains("������");

            // ���� �������� ������� � ������ ������ � ���������
            if (selection != null && selection.Cells.Count > 1 && mentionsTable)
            {
                string markdown = MarkdownGenerator.GetSelectedRangeTextAsMarkdown(selection);
                if (!string.IsNullOrEmpty(markdown))
                {
                    return $"��� ���������� ������� � ������� Markdown:\n\n{markdown}\n\n{userPrompt}";
                }
            }

            return userPrompt;
        }
    }
}

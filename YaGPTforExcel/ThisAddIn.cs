using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using YaGPTforExcel.Controls;

namespace YaGPTforExcel
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        // Свойство для доступа к taskPane
        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return this.taskPane; }
            private set { }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Создаем экземпляр панели
            var promptPaneControl = new PromptPaneControl();

            // Создаем TaskPane и привязываем его к приложению Excel
            taskPane = this.CustomTaskPanes.Add(promptPaneControl, "Prompt Panel");

            // Закрепляем панель сбоку
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

            // Задаем начальные размеры панели
            taskPane.Width = 300; // Ширина панели
      
            taskPane.Visible = true; // Панель будет видимой
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Освобождаем ресурсы
            if (taskPane != null)
            {
                try
                {
                    taskPane.Visible = false;
                    taskPane.Dispose();
                }
                catch { }
            }
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

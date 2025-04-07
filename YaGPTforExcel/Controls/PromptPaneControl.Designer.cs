using System;
using System.Windows.Forms;

namespace YaGPTforExcel.Controls
{
    public partial class PromptPaneControl : UserControl
    {
        private TextBox txtPrompt;
        private TextBox txtResult;
        private Button btnSend;
        private CheckBox chkInsert;

        public event EventHandler<string> PromptSubmitted;
        public bool InsertIntoExcel => chkInsert.Checked;

       

        private void InitializeComponent()
        {
            this.txtPrompt = new TextBox();
            this.txtResult = new TextBox();
            this.btnSend = new Button();
            this.chkInsert = new CheckBox();
            this.SuspendLayout();

            // txtResult (вверху)
            this.txtResult.Multiline = true;
            this.txtResult.ScrollBars = ScrollBars.Vertical;
            this.txtResult.Location = new System.Drawing.Point(10, 10);
            this.txtResult.Size = new System.Drawing.Size(250, 600);
            this.txtResult.ReadOnly = true;
            this.txtResult.Text = "здесь будет отображаться ответ от Yandex GPT";
            this.txtResult.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // txtPrompt (под результатом)
            this.txtPrompt.Multiline = true;
            this.txtPrompt.ScrollBars = ScrollBars.Vertical;
            this.txtPrompt.Location = new System.Drawing.Point(10, 620);
            this.txtPrompt.Size = new System.Drawing.Size(250, 80);
            this.txtPrompt.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // btnSend (ниже prompt, слева)
            this.btnSend.Text = "Отправить";
            this.btnSend.Location = new System.Drawing.Point(10, 710);
            this.btnSend.Size = new System.Drawing.Size(100, 30);
            this.btnSend.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            this.btnSend.Click += new EventHandler(this.btnSend_Click);

            // chkInsert (рядом с кнопкой, справа)
            this.chkInsert.AutoSize = true;
            this.chkInsert.Text = "Вставить в Excel";
            this.chkInsert.Location = new System.Drawing.Point(120, 715);
            this.chkInsert.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // PromptPaneControl
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.txtPrompt);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.chkInsert);
            this.Size = new System.Drawing.Size(270, 260);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
       
    }
}

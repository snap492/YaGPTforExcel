using System;
using System.Windows.Forms;

namespace YaGPTforExcel.Controls
{
    public partial class SettingsForm : Form
    {
        public string Token { get; private set; }
        public string FolderId { get; private set; }

        public SettingsForm()
        {
            InitializeComponent();
            // Предзаполним поля, если настройки уже сохранены
            txtToken.Text = Properties.Settings.Default.Token;
            txtFolderId.Text = Properties.Settings.Default.FolderId;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Сохраняем введенные значения в настройки приложения
            Properties.Settings.Default.Token = txtToken.Text;
            Properties.Settings.Default.FolderId = txtFolderId.Text;
            Properties.Settings.Default.Save();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {

        }

       
    }
}

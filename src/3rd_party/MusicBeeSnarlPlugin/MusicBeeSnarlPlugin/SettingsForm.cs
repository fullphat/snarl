using System;
using System.Windows.Forms;

namespace MusicBeeSnarlPlugin
{
    public partial class SettingsForm : Form
    {
        private Settings settings;

        public SettingsForm(Settings settings)
        {
            InitializeComponent();
            this.settings = settings;
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            SetFormContent();
        }

        private void BrowseDefaultIconButton_Click(object sender, EventArgs e)
        {
            SelectDefaultIconFileDialog.ShowDialog();
            if (!String.IsNullOrWhiteSpace(SelectDefaultIconFileDialog.FileName))
            {
                settings.DefaultIconPath = SelectDefaultIconFileDialog.FileName;
            }
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            settings.Reset();
            SetFormContent();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SetSettingsContent();
        }

        private void SetFormContent()
        {
            HeadlineTextBox.Text = settings.Headline;
            TextTextBox.Text = settings.Text;
            DefaultIconPathTextBox.Text = settings.DefaultIconPath;
        }

        private void SetSettingsContent()
        {
            settings.Headline = HeadlineTextBox.Text;
            settings.Text = TextTextBox.Text;
            settings.DefaultIconPath = DefaultIconPathTextBox.Text;
        }
    }
}

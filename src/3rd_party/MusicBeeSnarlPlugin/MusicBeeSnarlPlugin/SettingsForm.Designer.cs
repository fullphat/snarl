namespace MusicBeeSnarlPlugin
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.SaveButton = new System.Windows.Forms.Button();
            this.MyCancelButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.HeadlineTextBox = new System.Windows.Forms.TextBox();
            this.TextTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.DefaultIconPathTextBox = new System.Windows.Forms.TextBox();
            this.SelectDefaultIconFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.BrowseDefaultIconButton = new System.Windows.Forms.Button();
            this.ResetButton = new System.Windows.Forms.Button();
            this.VariablesToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // SaveButton
            // 
            this.SaveButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.SaveButton.Location = new System.Drawing.Point(36, 141);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(105, 23);
            this.SaveButton.TabIndex = 7;
            this.SaveButton.Text = "&Save";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // MyCancelButton
            // 
            this.MyCancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.MyCancelButton.Location = new System.Drawing.Point(147, 141);
            this.MyCancelButton.Name = "MyCancelButton";
            this.MyCancelButton.Size = new System.Drawing.Size(105, 23);
            this.MyCancelButton.TabIndex = 8;
            this.MyCancelButton.Text = "&Cancel";
            this.MyCancelButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Default icon:";
            // 
            // HeadlineTextBox
            // 
            this.HeadlineTextBox.Location = new System.Drawing.Point(89, 12);
            this.HeadlineTextBox.Name = "HeadlineTextBox";
            this.HeadlineTextBox.Size = new System.Drawing.Size(297, 20);
            this.HeadlineTextBox.TabIndex = 0;
            this.HeadlineTextBox.Text = "Headline text";
            this.VariablesToolTip.SetToolTip(this.HeadlineTextBox, "Artist: %artist%\r\nTitle: %track_title%\r\nAlbum: %album%\r\nTrack number: %track_no%\r" +
        "\nTrack year: %track_year%\r\n\r\nNew line: %newline%");
            // 
            // TextTextBox
            // 
            this.TextTextBox.Location = new System.Drawing.Point(89, 38);
            this.TextTextBox.Name = "TextTextBox";
            this.TextTextBox.Size = new System.Drawing.Size(297, 20);
            this.TextTextBox.TabIndex = 2;
            this.TextTextBox.Text = "%track_title%%newline%%album%";
            this.VariablesToolTip.SetToolTip(this.TextTextBox, "Artist: %artist%\r\nTitle: %track_title%\r\nAlbum: %album%\r\nTrack number: %track_no%\r" +
        "\nTrack year: %track_year%\r\n\r\nNew line: %newline%");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Headline:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(52, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Text:";
            // 
            // DefaultIconPathTextBox
            // 
            this.DefaultIconPathTextBox.Location = new System.Drawing.Point(89, 64);
            this.DefaultIconPathTextBox.Name = "DefaultIconPathTextBox";
            this.DefaultIconPathTextBox.Size = new System.Drawing.Size(297, 20);
            this.DefaultIconPathTextBox.TabIndex = 4;
            this.DefaultIconPathTextBox.Text = "%musicbee_path%\\Plugins\\MusicBee_Logo.png";
            this.VariablesToolTip.SetToolTip(this.DefaultIconPathTextBox, "MusicBee path: %musicbee_path%");
            // 
            // SelectDefaultIconFileDialog
            // 
            this.SelectDefaultIconFileDialog.DefaultExt = "png";
            this.SelectDefaultIconFileDialog.Filter = "Images|*.jpg,*.png|All files|*.*";
            this.SelectDefaultIconFileDialog.Title = "Select default icon";
            // 
            // BrowseDefaultIconButton
            // 
            this.BrowseDefaultIconButton.Location = new System.Drawing.Point(89, 90);
            this.BrowseDefaultIconButton.Name = "BrowseDefaultIconButton";
            this.BrowseDefaultIconButton.Size = new System.Drawing.Size(105, 23);
            this.BrowseDefaultIconButton.TabIndex = 6;
            this.BrowseDefaultIconButton.Text = "&Browse...";
            this.BrowseDefaultIconButton.UseVisualStyleBackColor = true;
            this.BrowseDefaultIconButton.Click += new System.EventHandler(this.BrowseDefaultIconButton_Click);
            // 
            // ResetButton
            // 
            this.ResetButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ResetButton.Location = new System.Drawing.Point(258, 141);
            this.ResetButton.Name = "ResetButton";
            this.ResetButton.Size = new System.Drawing.Size(105, 23);
            this.ResetButton.TabIndex = 9;
            this.ResetButton.Text = "&Reset";
            this.ResetButton.UseVisualStyleBackColor = true;
            this.ResetButton.Click += new System.EventHandler(this.ResetButton_Click);
            // 
            // VariablesToolTip
            // 
            this.VariablesToolTip.AutoPopDelay = 30000;
            this.VariablesToolTip.InitialDelay = 500;
            this.VariablesToolTip.ReshowDelay = 100;
            this.VariablesToolTip.ToolTipTitle = "Variables";
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 175);
            this.Controls.Add(this.ResetButton);
            this.Controls.Add(this.BrowseDefaultIconButton);
            this.Controls.Add(this.DefaultIconPathTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TextTextBox);
            this.Controls.Add(this.HeadlineTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.MyCancelButton);
            this.Controls.Add(this.SaveButton);
            this.Name = "SettingsForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Snarl settings";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.Button MyCancelButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox HeadlineTextBox;
        private System.Windows.Forms.TextBox TextTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox DefaultIconPathTextBox;
        private System.Windows.Forms.OpenFileDialog SelectDefaultIconFileDialog;
        private System.Windows.Forms.Button BrowseDefaultIconButton;
        private System.Windows.Forms.Button ResetButton;
        private System.Windows.Forms.ToolTip VariablesToolTip;
    }
}
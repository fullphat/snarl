namespace SnarlInterfaceExample1
{
	partial class Form1
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
			this.label1 = new System.Windows.Forms.Label();
			this.SnarlStatusLabel = new System.Windows.Forms.Label();
			this.SendNormalButton = new System.Windows.Forms.Button();
			this.SendCriticalButton = new System.Windows.Forms.Button();
			this.LogTextBox = new System.Windows.Forms.TextBox();
			this.SendLowButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(13, 13);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(65, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Snarl status:";
			// 
			// SnarlStatusLabel
			// 
			this.SnarlStatusLabel.Location = new System.Drawing.Point(84, 13);
			this.SnarlStatusLabel.Name = "SnarlStatusLabel";
			this.SnarlStatusLabel.Size = new System.Drawing.Size(93, 13);
			this.SnarlStatusLabel.TabIndex = 1;
			this.SnarlStatusLabel.Text = "Running";
			// 
			// SendNormalButton
			// 
			this.SendNormalButton.Location = new System.Drawing.Point(12, 42);
			this.SendNormalButton.Name = "SendNormalButton";
			this.SendNormalButton.Size = new System.Drawing.Size(165, 23);
			this.SendNormalButton.TabIndex = 2;
			this.SendNormalButton.Text = "Normal message";
			this.SendNormalButton.UseVisualStyleBackColor = true;
			this.SendNormalButton.Click += new System.EventHandler(this.SendNormalButton_Click);
			// 
			// SendCriticalButton
			// 
			this.SendCriticalButton.Location = new System.Drawing.Point(12, 71);
			this.SendCriticalButton.Name = "SendCriticalButton";
			this.SendCriticalButton.Size = new System.Drawing.Size(165, 23);
			this.SendCriticalButton.TabIndex = 3;
			this.SendCriticalButton.Text = "Critical message";
			this.SendCriticalButton.UseVisualStyleBackColor = true;
			this.SendCriticalButton.Click += new System.EventHandler(this.SendCriticalButton_Click);
			// 
			// LogTextBox
			// 
			this.LogTextBox.Location = new System.Drawing.Point(183, 6);
			this.LogTextBox.Multiline = true;
			this.LogTextBox.Name = "LogTextBox";
			this.LogTextBox.Size = new System.Drawing.Size(347, 117);
			this.LogTextBox.TabIndex = 4;
			// 
			// SendLowButton
			// 
			this.SendLowButton.Location = new System.Drawing.Point(12, 100);
			this.SendLowButton.Name = "SendLowButton";
			this.SendLowButton.Size = new System.Drawing.Size(165, 23);
			this.SendLowButton.TabIndex = 5;
			this.SendLowButton.Text = "Low priority message";
			this.SendLowButton.UseVisualStyleBackColor = true;
			this.SendLowButton.Click += new System.EventHandler(this.SendLowButton_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(542, 133);
			this.Controls.Add(this.SendLowButton);
			this.Controls.Add(this.LogTextBox);
			this.Controls.Add(this.SendCriticalButton);
			this.Controls.Add(this.SendNormalButton);
			this.Controls.Add(this.SnarlStatusLabel);
			this.Controls.Add(this.label1);
			this.Name = "Form1";
			this.Text = "SnarlInterface example1";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label SnarlStatusLabel;
		private System.Windows.Forms.Button SendNormalButton;
		private System.Windows.Forms.Button SendCriticalButton;
		private System.Windows.Forms.TextBox LogTextBox;
		private System.Windows.Forms.Button SendLowButton;
	}
}


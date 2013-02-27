using System;
using System.Windows.Forms;

using Snarl.V41;


namespace QuickStartGuide_CSharp
{
	public partial class Form1 : Form
	{
		private SnarlConnector snarl = new SnarlConnector();

		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			snarl.RegisterApp("QuickStartGuide_CSharp", "QuickStartGuide_CSharp", "", IntPtr.Zero, 0, 0);
		}

		private void button2_Click(object sender, EventArgs e)
		{
			snarl.EZNotify("MsgId", "Title", "Text message", 10, "", 0, "", "");
		}

		private void Form1_FormClosed(object sender, FormClosedEventArgs e)
		{
			snarl.UnregisterApp();
		}
	}
}

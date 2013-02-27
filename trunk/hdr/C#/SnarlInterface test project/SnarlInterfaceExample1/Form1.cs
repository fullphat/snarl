using System;
using System.Windows.Forms;
using Snarl.V42;
using System.Text;
using System.Security.Principal;


namespace SnarlInterfaceExample1
{
	/// <summary>
	/// SnarlInterface Example application<br />
	/// <br />
	/// This is a simple example of what I, as maintainer of SnarlInterface, consider best practise
	/// when it comes to using SnarlInterface. It is not meant as a display of all Snarl features.
	/// Feel free to use the code in you own application.
	/// </summary>
	public partial class Form1 : Form
	{
		#region Snarl
		SnarlInterface snarlInterface = new SnarlInterface();

		// Snarl message classes
		const String SnarlClassNormal = "Normal";
		const String SnarlClassCritical = "Critical";

		private String snarlPassword = CreateSnarlPassword(8);

		// Action callback values
		enum SnarlActions
		{
			DoSomething = 1,
			DoSomethingElse
		}
		#endregion
		

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			InitializeSnarl();
		}
		
		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
		{
			// Clean up Snarl - There should be no need to unregister the event handlers at this point
			snarlInterface.Unregister();
		}

		private void InitializeSnarl()
		{
			var vers = SnarlInterface.GetVersion();

			// ReRegisterSnarl() is called when first starting, and when a launch of Snarl is detected after this program is started.
			ReRegisterSnarl();

			// After registering, setup event handlers.
			// Not needed to do more than once, unless you call UnregisterCallbackWindow()
			snarlInterface.CallbackEvent += CallbackEventHandler;

			// Using lambda expression
			snarlInterface.GlobalSnarlEvent += (snarlInstance, args) =>
			{
				Log("Received global event: " + args.GlobalEvent);

				if (args.GlobalEvent == SnarlInterface.GlobalEvent.SnarlLaunched)
					ReRegisterSnarl();
				else if (args.GlobalEvent == SnarlInterface.GlobalEvent.SnarlQuit)
					SnarlStatusLabel.Text = "Not running";
			};

			// Update UI
			SnarlStatusLabel.Text = SnarlInterface.GetSnarlWindow() == IntPtr.Zero ? "Not running" : "Running";
		}

		private void ReRegisterSnarl()
		{
			int result = 0;
			String snarlIcon = SnarlInterface.GetIconsPath() + "presence.png";

			result = snarlInterface.RegisterWithEvents("application/Noer_IT.Example1", "SnarlInterface example1", snarlIcon, snarlPassword, this.Handle, null, SnarlInterface.AppFlags.None);

			// result = snarlInterface.RegisterWithEvents("application/Noer_IT.Example1", "SnarlInterface example1", snarlIcon, snarlPassword, this.Handle, null);
			// result = snarlInterface.RegisterWithEvents("application/Noer_IT.Example1", "SnarlInterface example1", snarlIcon, snarlPassword);

			if (result < (int)SnarlInterface.SnarlStatus.Success)
				Log("Failed to register with Snarl. Error=" + ((SnarlInterface.SnarlStatus)(Math.Abs(result))).ToString());

			snarlInterface.AddClass(SnarlClassNormal, "Normal messages");
			snarlInterface.AddClass(SnarlClassCritical, "Critical messages");

			SnarlStatusLabel.Text = SnarlInterface.GetSnarlWindow() == IntPtr.Zero ? "Not running" : "Running";
		}

		void CallbackEventHandler(SnarlInterface sender, SnarlInterface.CallbackEventArgs e)
		{
			switch (e.SnarlEvent)
			{
				case SnarlInterface.SnarlStatus.NotifyAction:
					HandleActionCallback(e.Parameter, e.MessageToken);
					break;

				case SnarlInterface.SnarlStatus.CallbackInvoked:
					Log("Left button clicked on {0}.", e.MessageToken);
					break;

				case SnarlInterface.SnarlStatus.CallbackTimedOut:
					Log("Message with token={0} timed out.", e.MessageToken);
					break;

				default:
					Log("Received callback event: " + e.SnarlEvent);
					break;
			}
		}

		private void HandleActionCallback(UInt16 actionData, int msgToken)
		{
			switch ((SnarlActions)actionData)
			{
				case SnarlActions.DoSomething:
					Log("DoSomething action callback (msgToken={0})", msgToken);
					break;
				case SnarlActions.DoSomethingElse:
					Log("DoSomethingElse action callback (msgToken={0})", msgToken);
					break;
			}
		}

		private void SendNormalButton_Click(object sender, EventArgs e)
		{
			Int32 msgToken = snarlInterface.Notify(SnarlClassNormal, "Normal message", "Some text", null, null, null,  SnarlInterface.MessagePriority.Normal);

			snarlInterface.AddAction(msgToken, "Do something", "@" + (int)SnarlActions.DoSomething);
			snarlInterface.AddAction(msgToken, "Do something else", "@" + (int)SnarlActions.DoSomethingElse);
		}

		private void SendCriticalButton_Click(object sender, EventArgs e)
		{
			Int32 msgToken = snarlInterface.Notify(SnarlClassNormal, "Critical message", "Some text\nNo need to manually escape & = # etc. btw.", null, null, null, SnarlInterface.MessagePriority.High);

			snarlInterface.AddAction(msgToken, "Do something", "@" + (int)SnarlActions.DoSomething);
			snarlInterface.AddAction(msgToken, "Do something else", "@" + (int)SnarlActions.DoSomethingElse);
		}
		
		private void SendLowButton_Click(object sender, EventArgs e)
		{
			Int32 msgToken = snarlInterface.Notify(SnarlClassNormal, "Low priority message", "Some text", 10, null, null, SnarlInterface.MessagePriority.Low);
		}

		private void Log(String msg, params object[] args)
		{
			String formattedMsg = "";
			msg = msg + Environment.NewLine;
			if (args.Length > 0)
				formattedMsg = String.Format(msg, args);
			else
				formattedMsg = msg;

			if (LogTextBox.InvokeRequired)
			{
				LogTextBox.Invoke((Action)(() => Log(formattedMsg)));
			}
			else
			{
				LogTextBox.AppendText(formattedMsg);
			}
		}

		private static string CreateSnarlPassword(int length)
		{
			// It is a bad idea to create a random generated password, as this will make register fail,
			// if the application is quit without proper unregister call. (Since passwords won't match between the two application instances.)

			// Generate "static" password
			String pass = WindowsIdentity.GetCurrent().Name.ToString() + "Snarl";
			return pass;

			// Generate random password
			/*Random random = new Random();
			StringBuilder sb = new StringBuilder(length);

			for (int i = 0; i < length; ++i)
			{
				sb.Append(Convert.ToChar(random.Next(65, 65 + 25)));
			}
			return sb.ToString();*/
			
		}
	}
}

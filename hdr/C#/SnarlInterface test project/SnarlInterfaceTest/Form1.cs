using System;
using System.Windows.Forms;
using V39 = Snarl.V39;
using V41 = Snarl.V41;
using V42 = Snarl.V42;

namespace SnarlConnectorTest
{
	public partial class Form1: Form
	{
		private const string SNARL_APP_NAME = "SnarlTestApp";
		private const string TEST_ALERT = "TestAlert";

		private const string SNARL_CLASS_1 = "Class1";
		private const string SNARL_CLASS_2 = "Class2";
		
		private const string SPECIAL_STRING = "Special characters: 完了しました != 完乾Eました and おはよう != おEよう";
		
		private string _iconPath = "";
		private string _soundPath = "";
		private Int32 _messageHandle = Int32.MinValue;
		private Int32 _globalMsgV39 = 0;

		private const Int32 ReplyMsgV39 = 0x400 + 100;  // The uReplyMsg used in SendMessage / SendMessageEx - 100 is a random number in WM_USER range
		private const Int32 ReplyMsgV41 = 0x400 + 101;
		private const Int32 ReplyMsgV42 = 0x400 + 102;


		public Form1()
		{
			InitializeComponent();

			// Hardcode path values
			_iconPath = Environment.GetEnvironmentVariable("programfiles") + "\\full phat\\Snarl\\etc\\icons\\info.png";
			_soundPath = Environment.GetEnvironmentVariable("windir") + "\\Media\\tada.wav";

			// Print Snarl information
			Log(String.Format("Snarl handle: {0}", V39.SnarlConnector.GetSnarlWindow()));

			UInt16 major, minor;
			bool snarlsnGetVersion = V39.SnarlConnector.GetVersion( out major, out minor );
			Log( String.Format( "GetVersion: {0}, Version: {1}.{2}", snarlsnGetVersion, major, minor ) );

			int nVersionEx = V39.SnarlConnector.GetVersionEx();
			if (nVersionEx < 0)
				Log(String.Format("GetVersionEx: {0}", V39.SnarlConnector.ConvertToMResult(nVersionEx).ToString()));
			else
				Log(String.Format("GetVersionEx: v{0}", nVersionEx));

			int nRevision = V39.SnarlConnector.GetRevision();
			if (nRevision <= 0)
				Log(String.Format("GetRevision: {0}{1}", V39.SnarlConnector.ConvertToMResult(nRevision).ToString(), Environment.NewLine));
			else
				Log(String.Format("GetRevision: {0}{1}", nRevision, Environment.NewLine));


			Log(String.Format("Snarl path: {0}", V39.SnarlConnector.GetAppPath()));
			Log(String.Format("Icons path: {0}{1}", V39.SnarlConnector.GetIconsPath(), Environment.NewLine));

			_globalMsgV39 = V39.SnarlConnector.GetGlobalMsg();
			Log(String.Format("Global message: {0}{1}", _globalMsgV39, Environment.NewLine));
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
		{
			V39.SnarlConnector.UnregisterApp();
			V39.SnarlConnector.RevokeConfig(this.Handle);

			v41SnarlConnector.UnregisterApp();
			v42Snarl.Unregister();
		}

		/// <summary>
		/// Override the message pump to receive the global Snarl messages.
		/// Use to automatically register with Snarl and to receive mouse click information
		/// </summary>
		protected override void WndProc(ref Message m)
		{
			WndProc_V39(ref m);
			WndProc_V41(ref m);
			WndProc_V42(ref m);

			base.WndProc(ref m);
		}

		private void WndProc_V39(ref Message m)
		{
			if (m.Msg == _globalMsgV39)
			{
				if (m.WParam == (IntPtr)V39.SnarlConnector.SNARL_LAUNCHED)
				{
					Log("Received a global message that Snarl has been started.");
					V39.SnarlConnector.GetSnarlWindow(true); // Force an update of the window handle
				}
				else if (m.WParam == (IntPtr)V39.SnarlConnector.SNARL_QUIT)
					Log("Received a global message that Snarl has quit.");
			}
			else if (m.Msg == ReplyMsgV39)
			{
				// An event has happened for one of our messages

				if (m.WParam.ToInt32() == V39.SnarlConnector.SNARL_NOTIFICATION_ACK)
				{
					Log("Received a left mouse click");

					// System.Diagnostics.Process.Start("http://www.fullphat.net/");
				}
				else if (m.WParam.ToInt32() == V39.SnarlConnector.SNARL_NOTIFICATION_CANCELLED)
				{
					Log("Received a right mouse click");
				}
				else if (m.WParam.ToInt32() == V39.SnarlConnector.SNARL_NOTIFICATION_MIDDLE_BUTTON)
				{
					Log("Received a middle mouse click");
				}
				else if (m.WParam.ToInt32() == V39.SnarlConnector.SNARL_NOTIFICATION_TIMED_OUT)
				{
					Log("Received message timed out");
				}
			}
		}

		private void Log( string message )
		{
			this.textBox1.AppendText( message + Environment.NewLine );
		}

		
	#region V39 API

		private void button1_Click( object sender, EventArgs e )
		{
			_messageHandle = V39.SnarlConnector.ShowMessage("Snarl!", "Message Text Here...\n" + SPECIAL_STRING, 0, _iconPath, this.Handle, (V39.WindowsMessage)ReplyMsgV39);
			if (V39.SnarlConnector.IsValidMessageId(_messageHandle))
				Log(String.Format("Showing message: {0}", _messageHandle ) );
			else
				Log(String.Format("ShowMessage returned error: {0}", (V39.M_RESULT)(uint)_messageHandle));
		}

		private void button2_Click( object sender, EventArgs e )
		{
			if ( _messageHandle == Int32.MinValue )
				return;

			Log(String.Format("Hiding message: {0}", V39.SnarlConnector.HideMessage(_messageHandle)));
		}

		private void button3_Click( object sender, EventArgs e )
		{
			if ( _messageHandle == Int32.MinValue )
				return;

			Log(String.Format("Snarl message visible: {0}", V39.SnarlConnector.IsMessageVisible(_messageHandle)));
		}

		private void button4_Click( object sender, EventArgs e )
		{
			V39.M_RESULT result = V39.SnarlConnector.RevokeConfig(this.Handle);
			Log(String.Format( "Disconnected: {0}", result ) );
		}

		private void button5_Click( object sender, EventArgs e )
		{
			Log(String.Format("Config registered [2]: {0}", V39.SnarlConnector.RegisterConfig(this.Handle, SNARL_APP_NAME, V39.WindowsMessage.WM_MDIMAXIMIZE, _iconPath)));
		}

		private void button6_Click( object sender, EventArgs e )
		{
			Log(String.Format("{0} registered: {1}", TEST_ALERT, V39.SnarlConnector.RegisterAlert(SNARL_APP_NAME, TEST_ALERT)));
		}

		private void button7_Click( object sender, EventArgs e )
		{
			if ( _messageHandle == Int32.MinValue )
				return;

			Log(String.Format("Timeout set to 5 seconds: {0}", V39.SnarlConnector.SetTimeout(_messageHandle, 5)));
		}

		private void button8_Click( object sender, EventArgs e )
		{
			_messageHandle = V39.SnarlConnector.ShowMessageEx(TEST_ALERT, "Test Alert!", "Blah blah blah...\n\n" + SPECIAL_STRING, 10, _iconPath, this.Handle, (V39.WindowsMessage)ReplyMsgV39, _soundPath);
			if (V39.SnarlConnector.IsValidMessageId(_messageHandle))
				Log(String.Format("Showing extended message: {0}", _messageHandle ) );
			else
				Log(String.Format("ShowMessageEx returned error: {0}", (V39.M_RESULT)(uint)_messageHandle));
		}

		private void button9_Click( object sender, EventArgs e )
		{
			if ( _messageHandle == Int32.MinValue )
				return;

			Log("Updating message: " + V39.SnarlConnector.UpdateMessage(_messageHandle, "Updated Title", "Updated text...\n\n" + SPECIAL_STRING, @"C:\Program Files\full phat\Snarl\etc\icons\about.png").ToString());
		}

		private void button10_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}
		
		private void button11_Click(object sender, EventArgs e)
		{
			Log("Requesting Snarl Test Message");
			V39.SnarlConnector.TestMessage();
		}

		private void button12_Click(object sender, EventArgs e)
		{
			Log("Setting this app as a Snarl application");
			V39.SnarlConnector.SetAsSnarlApp(this.Handle, V39.SnarlConnector.SNARL_APP_FLAGS.SNARL_APP_HAS_ABOUT);
		}

		private void button13_Click(object sender, EventArgs e)
		{
			uint n = V39.SnarlConnector.GetAppMsg();
			Log("Application message ID: " + n);
		}

		private void button14_Click(object sender, EventArgs e)
		{
			Int32 result = V39.SnarlConnector.RegisterApp(SNARL_APP_NAME, "", "", this.Handle, ReplyMsgV39 + 1);
			if (result <= 0)
				Log("RegisterApp response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("RegisterApp response: " + result);
		}

		private void button15_Click(object sender, EventArgs e)
		{
			Int32 result = V39.SnarlConnector.UnregisterApp();
			if (result <= 0)
				Log("UnregisterApp response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("UnregisterApp response: " + result);
		}

		private void button16_Click(object sender, EventArgs e)
		{
			Int32 result = V39.SnarlConnector.ShowNotification(SNARL_CLASS_1, "Test title", "ShowNotification test\n10 sec timeout\n" + SPECIAL_STRING, 10, "", this.Handle, ReplyMsgV39, "");
			if (result < 0)
				Log("ShowNotification response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else {
				Log("ShowNotification response: " + result);
				_messageHandle = result;
			}
		}

		private void button20_Click(object sender, EventArgs e)
		{
			Int32 result = V39.SnarlConnector.GetRevision();
			if (result < 0)
				Log("GetRevision response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("GetRevision response: " + result);
		}

		private void button19_Click(object sender, EventArgs e)
		{
			Int32 result = V39.SnarlConnector.AddClass(SNARL_CLASS_1, "Class 1", V39.SnarlConnector.SNARL_CLASS_FLAGS.SNARL_CLASS_ENABLED, "Default title", "", 5);
			if (result <= 0)
				Log("AddClass response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("AddClass response: " + result);
		}

		private void button17_Click(object sender, EventArgs e)
		{
			//Int32 result = SnarlConnector.ChangeAttribute(_messageHandle, SnarlConnector.SNARL_ATTRIBUTES.SNARL_ATTRIBUTE_TEXT, "New text");
			Int32 result = V39.SnarlConnector.ChangeAttribute(_messageHandle, V39.SnarlConnector.SNARL_ATTRIBUTES.SNARL_ATTRIBUTE_TITLE, "New title");
			if (result <= 0)
				Log("ChangeAttribute response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("ChangeAttribute response: " + result);
		}

		private void button18_Click(object sender, EventArgs e)
		{
			Log("Changing timeout for Class1 to 15 sec.");

			Int32 result = V39.SnarlConnector.SetClassDefault(SNARL_CLASS_1, V39.SnarlConnector.SNARL_ATTRIBUTES.SNARL_ATTRIBUTE_TIMEOUT, "15");
			if (result <= 0)
				Log("SetClassDefault response: " + V39.SnarlConnector.ConvertToMResult(result).ToString());
			else
				Log("SetClassDefault response: " + result);
			
		}

	#endregion


	#region V41 API

		private const String AppId = "CSharpTestApp";
		private const String Class1 = "class1";

		private Int32 classToken1 = 0;
		private V41.SnarlConnector v41SnarlConnector = new V41.SnarlConnector();
		private uint _globalMsgV41 = 0;

		private void WndProc_V41(ref Message m)
		{
			if (m.Msg == _globalMsgV41)
			{
				if (m.WParam == (IntPtr)V41.SnarlConnector.GlobalEvent.SnarlLaunched)
					Log("v41SnarlConnector: Received a global message that Snarl has been started.");
				else if (m.WParam == (IntPtr)V41.SnarlConnector.GlobalEvent.SnarlQuit)
					Log("v41SnarlConnector: Received a global message that Snarl has quit.");
			}
			else if (m.Msg == ReplyMsgV41)
			{
				// An event has happened for one of our messages

				if (m.WParam == (IntPtr)V41.SnarlConnector.MessageEvent.NotificationAck)
				{
					Log("v41SnarlConnector: Received a left mouse click");

					// System.Diagnostics.Process.Start("http://www.fullphat.net/");
				}
				else if (m.WParam == (IntPtr)V41.SnarlConnector.MessageEvent.NotificationCancelled)
				{
					Log("v41SnarlConnector: Received a right mouse click");
				}
				else if (m.WParam == (IntPtr)V41.SnarlConnector.MessageEvent.NotificationMiddleButton)
				{
					Log("v41SnarlConnector: Received a middle mouse click");
				}
				else if (m.WParam == (IntPtr)V41.SnarlConnector.MessageEvent.NotificationTimedOut)
				{
					Log("v41SnarlConnector: Received message timed out");
				}
			}
		}

		private void button29_Click(object sender, EventArgs e)
		{
			String iconPath = v41SnarlConnector.GetIconsPath() + "debug.png";

			V41.SnarlConnector.AppFlags flags = V41.SnarlConnector.AppFlags.AppHasPrefs | V41.SnarlConnector.AppFlags.AppIsWindowless;

			Log("v41SnarlConnector.RegisterApp: " + v41SnarlConnector.RegisterApp(AppId, "CSharp Test App", iconPath, this.Handle, ReplyMsgV41, flags));

			_globalMsgV41 = Snarl.V41.SnarlConnector.Broadcast();
		}

		private void button28_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.UnregisterApp: " + v41SnarlConnector.UnregisterApp());
		}

		private void button27_Click(object sender, EventArgs e)
		{
			classToken1 = v41SnarlConnector.AddClass(Class1, "Test class 1", true);
			Log("v41SnarlConnector.AddClass: " + classToken1);
		}

		private void button26_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.RemoveClass: " + v41SnarlConnector.RemoveClass(Class1, false));
		}

		private void button25_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.RemoveAllClasses: " + v41SnarlConnector.RemoveAllClasses(true));
		}

		private void button24_Click(object sender, EventArgs e)
		{
			String packetData = "title::Message title#?text::" + SPECIAL_STRING + "#?icon::" + v41SnarlConnector.GetIconsPath() + "snarl.png";

			Log("v41SnarlConnector.Notify: " + v41SnarlConnector.Notify(Class1, packetData));
		}

		private void button23_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.EZNotify: " + v41SnarlConnector.EZNotify("id", "title", SPECIAL_STRING, 10, "", Snarl.V41.SnarlConnector.MessagePriority.Normal, "ack", "value"));
		}

		private void button22_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.EZUpdate: " + v41SnarlConnector.EZUpdate(v41SnarlConnector.GetLastMsgToken(), "Updated title", "Updated text", 0, ""));
		}

		private void button21_Click(object sender, EventArgs e)
		{
			String packetData = "title::Updated message#?text::Updated message text#?icon::" + v41SnarlConnector.GetIconsPath() + "snarl-update.png";

			Log("v41SnarlConnector.Update: " + v41SnarlConnector.Update(v41SnarlConnector.GetLastMsgToken(), packetData));
		}

		private void button30_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.Hide: " + v41SnarlConnector.Hide(v41SnarlConnector.GetLastMsgToken()));
		}

		private void button32_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.GetVersion: " + v41SnarlConnector.GetVersion());
		}

		private void button31_Click(object sender, EventArgs e)
		{
			Log("v41SnarlConnector.Lasterror: " + ((V41.SnarlConnector.SnarlStatus)v41SnarlConnector.GetLastError()).ToString());
		}

		#endregion

	#region V42 API

		private const String v42AppId = "NoerIT/CSharpTestApp";
		private const String v42Class1 = "class1";

		private V42.SnarlInterface v42Snarl = new V42.SnarlInterface();
		private String v42Password = "CSharpTestApp/password";
		private uint _globalMsgV42 = 0;

		private void WndProc_V42(ref Message m)
		{
			/*
			if (m.Msg == _globalMsgV42)
			{
				if (m.WParam == (IntPtr)V42.SnarlInterface.GlobalEvent.SnarlLaunched)
					Log("v42SnarlConnector: Received a global message that Snarl has been started.");
				else if (m.WParam == (IntPtr)V42.SnarlInterface.GlobalEvent.SnarlQuit)
					Log("v42SnarlConnector: Received a global message that Snarl has quit.");
			}
			else if (m.Msg == ReplyMsgV42)
			{
				// An event has happened for one of our messages
				Int16 eventCode = (Int16)(m.WParam.ToInt32() & 0xffff);
				Int16 data = (Int16)(m.WParam.ToInt32() >> 16);

				if (eventCode == (Int16)V42.SnarlInterface.SnarlStatus.CallbackInvoked)
				{
					Log("v42SnarlConnector: Received a left mouse click");

					// System.Diagnostics.Process.Start("http://www.fullphat.net/");
				}
				else if (eventCode == (Int16)V42.SnarlInterface.SnarlStatus.CallbackRightClick)
				{
					Log("v42SnarlConnector: Received a right mouse click");
				}
				else if (eventCode == (Int16)V42.SnarlInterface.SnarlStatus.CallbackMiddleClick)
				{
					Log("v42SnarlConnector: Received a middle mouse click");
				}
				else if (eventCode == (Int16)V42.SnarlInterface.SnarlStatus.CallbackTimedOut)
				{
					Log("v42SnarlConnector: Received message timed out");
				}
				else if (eventCode == (Int16)V42.SnarlInterface.SnarlStatus.NotifyAction)
				{
					Log("v42SnarlConnector: Received action callback. lowowrd=" + eventCode + " hiword=" + data);
				}
			}*/
		}

		private void button44_Click(object sender, EventArgs e)
		{
			String iconPath = V42.SnarlInterface.GetIconsPath() + "debug.png";

			Log("v42SnarlConnector.RegisterApp: " + v42Snarl.Register(v42AppId, "CSharp Test App", iconPath, v42Password, this.Handle, ReplyMsgV42, V42.SnarlInterface.AppFlags.None));

			_globalMsgV42 = V42.SnarlInterface.Broadcast();
		}

		private void button43_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.UnregisterApp: " + v42Snarl.Unregister());
		}

		private void button42_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.AddClass: " + v42Snarl.AddClass(v42Class1, "Class1"));
		}

		private void button41_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.RemoveClass: " + v42Snarl.RemoveClass(v42Class1));
		}

		private void button40_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.ClearClasses: " + v42Snarl.ClearClasses());
		}

		private void button33_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.GetVersion: " + V42.SnarlInterface.GetVersion());
		}

		private void button38_Click(object sender, EventArgs e)
		{
			// Needs escaping
			const string SPECIAL_STRING2 = "Special characters: 完了しました 完乾Eました & some escaped chars ? == && = おはよう ! おEよう";

			//Log("v42SnarlConnector.Notify: " + v42Snarl.Notify(v42Class1, "C# test title", "Test message", 0));
			Log("v42SnarlConnector.Notify: " + v42Snarl.Notify(v42Class1, "C# test title", SPECIAL_STRING2, 0));
			Log("v42SnarlConnector.Notify: " + v42Snarl.Notify(v42Class1, "C# test - low priority 1", "Test message low priority", 0, null, null, V42.SnarlInterface.MessagePriority.Low));
			Log("v42SnarlConnector.Notify: " + v42Snarl.Notify(v42Class1, "C# test - low priority 2", "Test message low priority", 0, null, null, V42.SnarlInterface.MessagePriority.Low));
			// Log("v42SnarlConnector.Notify: " + v42Snarl.Notify(v42Class1, "C# test - low priority 2", SPECIAL_STRING2, 0, null, null, V42.SnarlInterface.MessagePriority.Low));
		}

		private void button37_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.Update - Not implemented");
		}

		private void button36_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Open C:\\", "C:\\"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback 1", "@1"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback 2", "@2"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback -1", "@-1"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback -2", "@-2"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback 32767", "@32767"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback 32768", "@32768"));
			Log("v42SnarlConnector.AddAction: " + v42Snarl.AddAction(v42Snarl.GetLastMsgToken(), "Dynamic callback 65535", "@65535"));
		}

		private void button35_Click(object sender, EventArgs e)
		{
			Log("v42SnarlConnector.ClearActions: " + v42Snarl.ClearActions(v42Snarl.GetLastMsgToken()));
		}

	#endregion

		private void button34_Click(object sender, EventArgs e)
		{
			/*v42Snarl.AttachCallbackWindow(this.Handle);
			v42Snarl.CallbackEvent += CallbackEventHandler;
			v42Snarl.GlobalSnarlEvent += GlobalSnarlEventHandler;*/
		}

		void GlobalSnarlEventHandler(Snarl.V42.SnarlInterface sender, Snarl.V42.SnarlInterface.GlobalEventArgs e)
		{
			if (e.GlobalEvent == V42.SnarlInterface.GlobalEvent.SnarlLaunched)
				Log("SnarlActionEventHandler: Snarl launched");
			else if (e.GlobalEvent == V42.SnarlInterface.GlobalEvent.SnarlQuit)
				Log("SnarlActionEventHandler: Snarl quit");
		}

		void CallbackEventHandler(Snarl.V42.SnarlInterface sender, Snarl.V42.SnarlInterface.CallbackEventArgs e)
		{
			switch (e.SnarlEvent)
			{
				case V42.SnarlInterface.SnarlStatus.NotifyAction:
					Log(String.Format("Action callback - @data={0}, MsgToken={1}", e.Parameter, e.MessageToken));
					break;
				case V42.SnarlInterface.SnarlStatus.CallbackClosed:
					Log(String.Format("CallbackClosed - MsgToken={0}", e.MessageToken));
					break;
				case V42.SnarlInterface.SnarlStatus.CallbackTimedOut:
					Log(String.Format("CallbackTimedOut - MsgToken={0}", e.MessageToken));
					break;
			}
		}

		private void button39_Click(object sender, EventArgs e)
		{
			v42Snarl.CallbackEvent -= CallbackEventHandler;
			v42Snarl.GlobalSnarlEvent -= GlobalSnarlEventHandler;
			v42Snarl.Unregister();
		}
	}
}

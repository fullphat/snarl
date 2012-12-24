using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;


namespace Snarl.V41
{
	/// <summary>
	/// SnarlConnector
	/// Implementation of the V41 API.
	/// 
	/// Please note the following changes compared to the VB6 (official API) dokumentation:        
	///  - Naming of constants and variables are generally changed to follow Microsoft C# standard.
	///    - Grouped variables like SNARL_LAUNCHED, SNARL_QUIT is in enum GlobalEvent.
	///    - Message events like SNARL_NOTIFICATION_CLICKED, is found in enum MessageEvent.
	///  - Some functions in the VB API takes an appToken as first parameter. This token is a
	///    member variable in C# version, so it is omitted from the functions.
	///    (Always call RegisterApp as first function!)
	///  - Functions manipulating messages (Update, Hide etc.) still takes a message token as
	///    parameter, but you can get the last message token calling GetLastMsgToken();
	///    Example: snarl.Hide(snarl.GetLastMsgToken());
	///
	/// Note on optional parameters:
	///   Since C# 3.0 and prior doesn't have optional parameters, the C# SnarlConnector does not
	///   use this feature!
	///   Documentation for the C# functions includes the Optional keyword and a value expected
	///   passed for default behavior.
	///   Example:
	///     <param name="timeout">Optional (-1)</param>
	///     <param name="icon">Optional ("" or null)</param>
	///     Means that timeout parameter should be set to -1 and icon to either an empty string or
	///     null to get the default behavior from Snarl.
	/// 
	/// Funtions special to C# V41 API compared to VB version:
	///   GetLastMsgToken()
	///   GetAppPath()
	///   GetIconsPath()
	/// </summary>
	/// 
	/// <VersionHistory>
	/// 2011-01-15 : Added MessagePriority enum
	/// 2010-09-12 : Made AppMsg, Broadcast, IsSnarlRunning, GetSnarlWindow static (to be the same as C++ version)
	/// 2010-08-25 : Added overloaded versions of RegisterApp, EZNotify and EZUpdate.
	/// 2010-08-20 : Fixed not sending correct PackedData string in UpdateApp.
	/// 2010-08-14 : Clean-up, more error checking and documentation.
	///            : Converted global events and message events to enums.
	/// 2010-08-13 : Updated to last changes before 2.3 release.
	/// 2010-07-31 : Initial release of V41 API (for 2.3RC1)
	/// </VersionHistory>

	public class SnarlConnector
	{
		#region Public constants and enums

		/// <summary>
		/// Global event identifiers.
		/// Identifiers marked with a '*' are sent by Snarl in two ways:
		///   1. As a broadcast message (uMsg = 'SNARL_GLOBAL_MSG')
		///   2. To the window registered in snRegisterConfig() or snRegisterConfig2()
		///      (uMsg = reply message specified at the time of registering)
		/// In both cases these values appear in wParam.
		///   
		/// Identifiers not marked are not broadcast; they are simply sent to the application's registered window.
		/// </summary>
		public enum GlobalEvent
		{
			SnarlLaunched = 1,      // Snarl has just started running*
			SnarlQuit = 2,          // Snarl is about to stop running*
			SnarlAskAppletVer = 3,  // (R1.5) Reserved for future use
			SnarlShowAppUi = 4      // (R1.6) Application should show its UI
		}

		/// <summary>
		/// Message event identifiers.
		/// These are sent by Snarl to the window specified in RegisterApp() when the
		/// Snarl Notification raised times out or the user clicks on it.
		/// </summary>
		public enum MessageEvent
		{
			NotificationClicked = 32,      // Notification was right-clicked by user
			NotificationCancelled = 32,    // Added in V37 (R1.6) -- same value, just improved the meaning of it
			NotificationTimedOut = 33,     // 
			NotificationAck = 34,          // Notification was left-clicked by user
			NotificationMenu = 35,         // Menu item selected (V39)
			NotificationMiddleButton = 36, // Notification middle-clicked by user (V39)
			NotificationClosed = 37        // User clicked the close gadget (V39)
		}

		/// <summary>
		/// Error values returned by calls to GetLastError().
		/// </summary>
		public enum SnarlStatus : short
		{
			Success = 0,

			ErrorFailed = 101,        // miscellaneous failure
			ErrorUnknownCommand,      // specified command not recognised
			ErrorTimedOut,            // Snarl took too long to respond

			ErrorArgMissing = 109,    // required argument missing
			ErrorSystem,              // internal system error

			ErrorNotRunning = 201,    // Snarl handling window not found
			ErrorNotRegistered,       // 
			ErrorAlreadyRegistered,   // not used yet; RegisterApp() returns existing token
			ErrorClassAlreadyExists,  // not used yet; AddClass() returns existing token
			ErrorClassBlocked,
			ErrorClassNotFound,
			ErrorNotificationNotFound
		}

		/// <summary>
		/// The priority of messages.
		/// See <cref>http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#notify</cref>
		/// </summary>
		public enum MessagePriority
		{
			Low = -1,
			Normal = 0,
			High = 1
		}

		/// <summary>
		/// Application flags - features this app supports.
		/// </summary>
		[Flags]
		public enum AppFlags
		{
			AppHasPrefs = 1,
			AppHasAbout = 2,
			AppIsWindowless = 0x8000
		}

		public const uint WM_SNARLTEST = (uint)WindowsMessage.WM_USER + 237;

		#endregion

		#region Internal constants and enums

		protected const string SnarlWindowClass = "w>Snarl";
		protected const string SnarlWindowTitle = "Snarl";

		protected const string SnarlGlobalMsg = "SnarlGlobalEvent";
		protected const string SnarlAppMsg = "SnarlAppMessage";

		protected const int SnarlPacketDataSize = 4096;

		protected enum SnarlCommand : short
		{
			RegisterApp = 1,           // for this command, SNARLMSG->Token is actually the sending app's PID
			UnregisterApp,
			UpdateApp,
			SetCallback,
			AddClass,
			RemoveClass,
			Notify,
			UpdateNotification,
			HideNotification,
			IsNotificationVisible,
			LastError                  // deprecated but retained for backwards compatability
		}

		[StructLayout(LayoutKind.Sequential, Pack = 4)]
		protected struct SnarlMessage
		{
			public SnarlCommand Command;
			public Int32 Token;

			[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlPacketDataSize)]
			public byte[] PacketData;
		}

		#endregion

		#region Member variables

		private Int32 appToken = 0;
		private Int32 lastMsgToken = 0;
		private SnarlStatus localError = 0;

		#endregion


		// -------------------------------------------------------------------

		/// <summary>
		/// Register application with Snarl.
		/// </summary>
		/// <param name="signatur"></param>
		/// <param name="title"></param>
		/// <param name="icon"></param>
		/// <param name="hWndReply">Optional (IntPtr.Zero)</param>
		/// <param name="msgReply">Optional (0)</param>
		/// <param name="flags">Optional (0)</param>
		/// <returns></returns>
		public Int32 RegisterApp(String signature, String title, String icon, IntPtr hWndReply, Int32 msgReply, AppFlags flags)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.RegisterApp;
			msg.Token = 0;
			msg.PacketData = StringToUtf8(
				"id::" + signature + 
				"#?title::" + title +
				"#?icon::"  + icon +
				"#?hwnd::"  + hWndReply.ToString() + 
				"#?umsg::"  + msgReply.ToString() + 
				"#?flags::" + ((int)flags).ToString() );

			appToken = Send(msg);
			lastMsgToken = 0;

			return appToken;
		}

		public Int32 RegisterApp(String signature, String title, String icon, IntPtr hWndReply, Int32 msgReply)
		{
			return RegisterApp(signature, title, icon, hWndReply, msgReply, 0);
		}

		public Int32 RegisterApp(String signature, String title, String icon)
		{
			return RegisterApp(signature, title, icon, IntPtr.Zero, 0, 0);
		}


		/// <summary>
		/// Unregister application.
		/// </summary>
		/// <returns></returns>
		public Int32 UnregisterApp()
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.UnregisterApp;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8("");

			appToken = 0;
			lastMsgToken = 0;

			return Send(msg);
		}

		/// <summary>
		/// UpdateApp
		/// </summary>
		/// <param name="title">Optional (null)</param>
		/// <param name="icon">Optional (null)</param>
		/// <returns></returns>
		public Int32 UpdateApp(String title, String icon)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.UpdateApp;
			msg.Token = appToken;

			String str = "";
			if (title != null && title.Trim().Length > 0)
				str += "title::" + title.Trim();
			
			if (icon != null && icon.Trim().Length > 0)
			{
				str += str.Length > 0 ? "#?icon::" : "icon::";
				str += icon.Trim();
			}
			msg.PacketData = StringToUtf8(str);

			return Send(msg);
		}

		/// <summary>
		/// AddClass
		/// </summary>
		/// <param name="name"></param>
		/// <param name="description"></param>
		/// <param name="enabled">Optional (true)</param>
		/// <returns></returns>
		public Int32 AddClass(String className, String description, bool enabled)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.AddClass;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8(
				"id::" + className +
				"#?name::" + description +
				"#?enabled::" + (enabled ? "1" : "0") );

			return Send(msg);
		}

		public Int32 AddClass(String className, String description)
		{
			return AddClass(className, description, true);
		}

		/// <summary>
		/// RemoveClass
		/// </summary>
		/// <param name="className"></param>
		/// <param name="forgetSettings">Optional (false)</param>
		/// <returns></returns>
		public Int32 RemoveClass(String className, bool forgetSettings)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.RemoveClass;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8(
				"id::" + className +
				"#?forget::" + (forgetSettings ? "1" : "0") );

			return Send(msg);
		}

		public Int32 RemoveClass(String className)
		{
			return RemoveClass(className, false);
		}
		
		/// <summary>
		/// RemoveAllClasses
		/// </summary>
		/// <param name="forgetSettings">Optional (false)</param>
		public Int32 RemoveAllClasses(bool forgetSettings)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.RemoveClass;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8(
				"all::1" +
				"#?forget::" + (forgetSettings ? "1" : "0") );

			return Send(msg);
		}

		public Int32 RemoveAllClasses()
		{
			return RemoveAllClasses(false);
		}

		/// <summary>
		/// EZNotify
		/// </summary>
		/// <param name="className"></param>
		/// <param name="title"></param>
		/// <param name="text"></param>
		/// <param name="timeout">Optional (Default -1)</param>
		/// <param name="icon">Optional  ("" or null)</param>
		/// <param name="priority">Optional (Default Normal)</param>
		/// <param name="acknowledge">Optional ("" or null)</param>
		/// <param name="value">Optional ("" or null)</param>
		/// <returns></returns>
		public Int32 EZNotify(String className, String title, String text, Int32 timeout, String icon, MessagePriority priority, String acknowledge, String value)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.Notify;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8(
				"id::" + className +
				"#?title::" + title +
				"#?text::" + text +
				"#?timeout::" + timeout.ToString() +
				"#?icon::" + ((icon != null) ? icon : "") +
				"#?priority::" + (int)priority +
				"#?ack::" + ((acknowledge != null) ? acknowledge : "") +
				"#?value::" + ((value != null) ? value : ""));

			lastMsgToken = Send(msg);
			return lastMsgToken;
		}

		public Int32 EZNotify(String className, String title, String text, Int32 timeout, String icon)
		{
			return EZNotify(className, title, text, timeout, icon, 0, "", "");
		}

		public Int32 EZNotify(String className, String title, String text, Int32 timeout)
		{
			return EZNotify(className, title, text, timeout, "", 0, "", "");
		}

		public Int32 EZNotify(String className, String title, String text)
		{
			return EZNotify(className, title, text, -1, "", 0, "", "");
		}

		/// <summary>
		/// Notify
		/// </summary>
		/// <param name="className"></param>
		/// <param name="packetData"></param>
		/// <returns></returns>
		public Int32 Notify(String className, String packetData)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.Notify;
			msg.Token = appToken;
			msg.PacketData = StringToUtf8(
				"id::" + className +
				"#?" + packetData );

			lastMsgToken = Send(msg);
			return lastMsgToken;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="msgToken"></param>
		/// <param name="title">Optional ("" or null)</param>
		/// <param name="text">Optional ("" or null)</param>
		/// <param name="timeout">Optional (-1)</param>
		/// <param name="icon">Optional ("" or null)</param>
		/// <returns></returns>
		public Int32 EZUpdate(Int32 msgToken, String title, String text, Int32 timeout, String icon)
		{
			StringBuilder sb = new StringBuilder();

			// All paramaters are optional - build PacketData
			if (title != null && title.Length > 0)
				sb.Append("title::" + title);
			
			if (text != null && text.Length > 0)
				sb.Append( ((sb.Length > 0) ? "#?" : "") + "text::" + text);

			if (icon != null && icon.Length > 0)
				sb.Append( ((sb.Length > 0) ? "#?" : "") + "icon::" + icon);

			if (timeout != -1)
				sb.Append( ((sb.Length > 0) ? "#?" : "") + "timeout::" + timeout);

			SnarlMessage msg;
			msg.Command = SnarlCommand.UpdateNotification;
			msg.Token = msgToken;
			msg.PacketData = StringToUtf8(sb.ToString());
			
			return Send(msg);
		}

		public Int32 EZUpdate(Int32 msgToken, String title, String text, Int32 timeout)
		{
			return EZUpdate(msgToken, title, text, timeout, null);
		}

		public Int32 EZUpdate(Int32 msgToken, String title, String text)
		{
			return EZUpdate(msgToken, title, text, -1, null);
		}

		/// <summary>
		/// Update
		/// </summary>
		/// <param name="msgToken"></param>
		/// <param name="packetData"></param>
		/// <returns></returns>
		public Int32 Update(Int32 msgToken, String packetData)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.UpdateNotification;
			msg.Token = msgToken;
			msg.PacketData = StringToUtf8(packetData);

			return Send(msg);
		}

		/// <summary>
		/// Hide
		/// </summary>
		/// <param name="msgToken"></param>
		/// <returns></returns>
		public Int32 Hide(Int32 msgToken)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.HideNotification;
			msg.Token = msgToken;
			msg.PacketData = StringToUtf8("");

			return Send(msg);
		}

		/// <summary>
		/// IsVisible
		/// </summary>
		/// <param name="msgToken"></param>
		/// <returns>
		///		Returns -1 if message is visible. 0 if not visible or if an error occured.
		///	</returns>
		public Int32 IsVisible(Int32 msgToken)
		{
			SnarlMessage msg;
			msg.Command = SnarlCommand.IsNotificationVisible;
			msg.Token = msgToken;
			msg.PacketData = StringToUtf8("");
			
			return Send(msg);
		}

		/// <summary>
		/// GetLastError
		/// </summary>
		/// <returns></returns>
		public SnarlStatus GetLastError()
		{
			return localError;
		}

		/// <summary>
		/// Returns true if Snarl system was found running.
		/// </summary>
		/// <returns></returns>
		static public bool IsSnarlRunning()
		{
			return IsWindow(GetSnarlWindow());
		}

		/// <summary>
		/// Returns the version of Snarl running.
		/// </summary>
		/// <returns></returns>
		public Int32 GetVersion()
		{
			localError = 0;

			IntPtr hWnd = GetSnarlWindow();
			if (!IsWindow(hWnd))
			{
				localError = SnarlStatus.ErrorNotRunning;
				return 0;
			}

			return GetProp(hWnd, "_version").ToInt32();
		}

		/// <summary>
		/// Returns the value of Snarl's global registered message.
		/// Notes:
		///   Snarl registers SNARL_GLOBAL_MSG during startup which it then uses to communicate
		///   with all running applications through a Windows broadcast message. This function can
		///   only fail if for some reason the Windows RegisterWindowMessage() function fails
		///   - given this, this function *cannnot* be used to test for the presence of Snarl.
		/// </summary>
		/// <returns>A 16-bit value (translated to 32-bit) which is the registered Windows message for Snarl.</returns>
		static public uint Broadcast()
		{
			return RegisterWindowMessage(SnarlGlobalMsg);
		}

		/// <summary>
		/// Returns the global Snarl Application message  (V39)
		/// </summary>
		/// <returns>Snarl Application registered message.</returns>
		static public uint AppMsg()
		{
			return RegisterWindowMessage(SnarlAppMsg);
		}

		/// <summary>
		/// Returns a handle to the Snarl Dispatcher window  (V37)
		/// Notes:
		///   This is now the preferred way to test if Snarl is actually running.
		/// </summary>
		/// <returns>Returns handle to Snarl Dispatcher window, or zero if it's not found</returns>
		static public IntPtr GetSnarlWindow()
		{
			IntPtr hwnd = FindWindow(SnarlWindowClass, SnarlWindowTitle);

			return hwnd;
		}

		/// <summary>
		/// Returns a fully qualified path to Snarl's installation folder.
		/// This is a V39 API method.
		/// </summary>
		/// <returns>Path to Snarl's installation folder. Empty string on failure.</returns>
		public string GetAppPath()
		{
			StringBuilder sb = new StringBuilder(512);

			IntPtr hwnd = GetSnarlWindow();
			if (hwnd != IntPtr.Zero)
			{
				IntPtr hWndPath = FindWindowEx(hwnd, IntPtr.Zero, "static", null);
				if (hWndPath != IntPtr.Zero)
				{
					GetWindowText(hWndPath, sb, 512);
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Returns a fully qualified path to Snarl's default icon location.
		/// This is a V39 API method.
		/// </summary>
		/// <returns>Path to Snarl's default icon location. Empty string on failure.</returns>
		public string GetIconsPath()
		{
			return Path.Combine(GetAppPath(), "etc\\icons\\");
		}

		/// <summary>
		/// GetLastMsgToken() returns token of the last message sent to Snarl.
		/// This function is not in the official API!
		/// </summary>
		/// <returns></returns>
		public Int32 GetLastMsgToken()
		{
			return lastMsgToken;
		}

		#region Private functions

		/// <summary>
		/// Send message to Snarl.
		/// </summary>
		/// <param name="msg"></param>
		/// <returns>Return zero on failure.</returns>
		private Int32 Send(SnarlMessage msg)
		{
			Int32 nReturn = 0; // Failure
			localError = 0;

			IntPtr hWnd = GetSnarlWindow();
			if (!IsWindow(hWnd))
			{
				localError = SnarlStatus.ErrorNotRunning;
				return 0;
			}

			IntPtr nSendMessageResult = IntPtr.Zero;
			IntPtr ptrToSnarlMessage = IntPtr.Zero;
			IntPtr ptrToCds = IntPtr.Zero;

			try
			{
				COPYDATASTRUCT cds = new COPYDATASTRUCT();
				cds.dwData = (IntPtr)0x534E4C02; // "SNL",2
				cds.cbData = Marshal.SizeOf(typeof(SnarlMessage));

				ptrToSnarlMessage = Marshal.AllocHGlobal(cds.cbData);
				Marshal.StructureToPtr(msg, ptrToSnarlMessage, false);
				cds.lpData = ptrToSnarlMessage;

				ptrToCds = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(COPYDATASTRUCT)));
				Marshal.StructureToPtr(cds, ptrToCds, false);

				if (SendMessageTimeout(hWnd,
						  (uint)WindowsMessage.WM_COPYDATA,
						  (IntPtr)GetCurrentProcessId(),
						  ptrToCds,
						  SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | SendMessageTimeoutFlags.SMTO_NOTIMEOUTIFNOTHUNG,
						  500,
						  out nSendMessageResult) == IntPtr.Zero)
				{
					// return zero on failure
					localError = SnarlStatus.ErrorTimedOut;
					return 0;
				}
				
				// return result and cache LastError
				nReturn = unchecked((Int32)nSendMessageResult.ToInt64()); // Avoid aritmetic overflow error
				localError = (SnarlStatus)GetProp(hWnd, "last_error");
			}
			finally
			{
				Marshal.FreeHGlobal(ptrToCds);
				Marshal.FreeHGlobal(ptrToSnarlMessage);
			}

			return nReturn;
		}

		/// <summary>
		/// Use this function to convert a string into an UTF8 encoded byte[]
		/// </summary>
		/// <param name="strToConvert">The managed string object to convert.</param>
		/// <returns><c>byte[]</c> with the converted string.</returns>
		private static byte[] StringToUtf8(string strToConvert)
		{
			byte[] returnString = new byte[SnarlPacketDataSize];

			UTF8Encoding utf8 = new UTF8Encoding();
			utf8.GetBytes(strToConvert, 0, strToConvert.Length, returnString, 0);

			return returnString;
		}

		#endregion

		#region Interop imports and structures

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, Int32 nMaxCount);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		internal static extern uint RegisterWindowMessage(string lpString);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		internal static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

		[DllImport("user32.dll")]
		internal static extern bool IsWindow(IntPtr hWnd);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		internal static extern IntPtr GetProp(IntPtr hWnd, string lpString);

		[DllImport("kernel32.dll")]
		internal static extern uint GetCurrentProcessId();

		[Flags]
		internal enum SendMessageTimeoutFlags : uint
		{
			SMTO_NORMAL = 0x0000,
			SMTO_BLOCK = 0x0001,
			SMTO_ABORTIFHUNG = 0x0002,
			SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
		}

		[StructLayout(LayoutKind.Sequential)]
		internal struct COPYDATASTRUCT
		{
			public IntPtr dwData;   // DWORD
			public Int32  cbData;   // DWORD
			public IntPtr lpData;   // PVOID
		}

		#endregion

		#region WindowsMessage enum

		public enum WindowsMessage : uint
		{
			WM_ACTIVATE = 0x6,
			WM_ACTIVATEAPP = 0x1C,
			WM_AFXFIRST = 0x360,
			WM_AFXLAST = 0x37F,
			WM_APP = 0x8000,
			WM_ASKCBFORMATNAME = 0x30C,
			WM_CANCELJOURNAL = 0x4B,
			WM_CANCELMODE = 0x1F,
			WM_CAPTURECHANGED = 0x215,
			WM_CHANGECBCHAIN = 0x30D,
			WM_CHAR = 0x102,
			WM_CHARTOITEM = 0x2F,
			WM_CHILDACTIVATE = 0x22,
			WM_CLEAR = 0x303,
			WM_CLOSE = 0x10,
			WM_COMMAND = 0x111,
			WM_COMPACTING = 0x41,
			WM_COMPAREITEM = 0x39,
			WM_CONTEXTMENU = 0x7B,
			WM_COPY = 0x301,
			WM_COPYDATA = 0x4A,
			WM_CREATE = 0x1,
			WM_CTLCOLORBTN = 0x135,
			WM_CTLCOLORDLG = 0x136,
			WM_CTLCOLOREDIT = 0x133,
			WM_CTLCOLORLISTBOX = 0x134,
			WM_CTLCOLORMSGBOX = 0x132,
			WM_CTLCOLORSCROLLBAR = 0x137,
			WM_CTLCOLORSTATIC = 0x138,
			WM_CUT = 0x300,
			WM_DEADCHAR = 0x103,
			WM_DELETEITEM = 0x2D,
			WM_DESTROY = 0x2,
			WM_DESTROYCLIPBOARD = 0x307,
			WM_DEVICECHANGE = 0x219,
			WM_DEVMODECHANGE = 0x1B,
			WM_DISPLAYCHANGE = 0x7E,
			WM_DRAWCLIPBOARD = 0x308,
			WM_DRAWITEM = 0x2B,
			WM_DROPFILES = 0x233,
			WM_ENABLE = 0xA,
			WM_ENDSESSION = 0x16,
			WM_ENTERIDLE = 0x121,
			WM_ENTERMENULOOP = 0x211,
			WM_ENTERSIZEMOVE = 0x231,
			WM_ERASEBKGND = 0x14,
			WM_EXITMENULOOP = 0x212,
			WM_EXITSIZEMOVE = 0x232,
			WM_FONTCHANGE = 0x1D,
			WM_GETDLGCODE = 0x87,
			WM_GETFONT = 0x31,
			WM_GETHOTKEY = 0x33,
			WM_GETICON = 0x7F,
			WM_GETMINMAXINFO = 0x24,
			WM_GETOBJECT = 0x3D,
			WM_GETSYSMENU = 0x313,
			WM_GETTEXT = 0xD,
			WM_GETTEXTLENGTH = 0xE,
			WM_HANDHELDFIRST = 0x358,
			WM_HANDHELDLAST = 0x35F,
			WM_HELP = 0x53,
			WM_HOTKEY = 0x312,
			WM_HSCROLL = 0x114,
			WM_HSCROLLCLIPBOARD = 0x30E,
			WM_ICONERASEBKGND = 0x27,
			WM_IME_CHAR = 0x286,
			WM_IME_COMPOSITION = 0x10F,
			WM_IME_COMPOSITIONFULL = 0x284,
			WM_IME_CONTROL = 0x283,
			WM_IME_ENDCOMPOSITION = 0x10E,
			WM_IME_KEYDOWN = 0x290,
			WM_IME_KEYLAST = 0x10F,
			WM_IME_KEYUP = 0x291,
			WM_IME_NOTIFY = 0x282,
			WM_IME_REQUEST = 0x288,
			WM_IME_SELECT = 0x285,
			WM_IME_SETCONTEXT = 0x281,
			WM_IME_STARTCOMPOSITION = 0x10D,
			WM_INITDIALOG = 0x110,
			WM_INITMENU = 0x116,
			WM_INITMENUPOPUP = 0x117,
			WM_INPUTLANGCHANGE = 0x51,
			WM_INPUTLANGCHANGEREQUEST = 0x50,
			WM_KEYDOWN = 0x100,
			WM_KEYFIRST = 0x100,
			WM_KEYLAST = 0x108,
			WM_KEYUP = 0x101,
			WM_KILLFOCUS = 0x8,
			WM_LBUTTONDBLCLK = 0x203,
			WM_LBUTTONDOWN = 0x201,
			WM_LBUTTONUP = 0x202,
			WM_MBUTTONDBLCLK = 0x209,
			WM_MBUTTONDOWN = 0x207,
			WM_MBUTTONUP = 0x208,
			WM_MDIACTIVATE = 0x222,
			WM_MDICASCADE = 0x227,
			WM_MDICREATE = 0x220,
			WM_MDIDESTROY = 0x221,
			WM_MDIGETACTIVE = 0x229,
			WM_MDIICONARRANGE = 0x228,
			WM_MDIMAXIMIZE = 0x225,
			WM_MDINEXT = 0x224,
			WM_MDIREFRESHMENU = 0x234,
			WM_MDIRESTORE = 0x223,
			WM_MDISETMENU = 0x230,
			WM_MDITILE = 0x226,
			WM_MEASUREITEM = 0x2C,
			WM_MENUCHAR = 0x120,
			WM_MENUCOMMAND = 0x126,
			WM_MENUDRAG = 0x123,
			WM_MENUGETOBJECT = 0x124,
			WM_MENURBUTTONUP = 0x122,
			WM_MENUSELECT = 0x11F,
			WM_MOUSEACTIVATE = 0x21,
			WM_MOUSEFIRST = 0x200,
			WM_MOUSEHOVER = 0x2A1,
			WM_MOUSELAST = 0x20A,
			WM_MOUSELEAVE = 0x2A3,
			WM_MOUSEMOVE = 0x200,
			WM_MOUSEWHEEL = 0x20A,
			WM_MOVE = 0x3,
			WM_MOVING = 0x216,
			WM_NCACTIVATE = 0x86,
			WM_NCCALCSIZE = 0x83,
			WM_NCCREATE = 0x81,
			WM_NCDESTROY = 0x82,
			WM_NCHITTEST = 0x84,
			WM_NCLBUTTONDBLCLK = 0xA3,
			WM_NCLBUTTONDOWN = 0xA1,
			WM_NCLBUTTONUP = 0xA2,
			WM_NCMBUTTONDBLCLK = 0xA9,
			WM_NCMBUTTONDOWN = 0xA7,
			WM_NCMBUTTONUP = 0xA8,
			WM_NCMOUSEHOVER = 0x2A0,
			WM_NCMOUSELEAVE = 0x2A2,
			WM_NCMOUSEMOVE = 0xA0,
			WM_NCPAINT = 0x85,
			WM_NCRBUTTONDBLCLK = 0xA6,
			WM_NCRBUTTONDOWN = 0xA4,
			WM_NCRBUTTONUP = 0xA5,
			WM_NEXTDLGCTL = 0x28,
			WM_NEXTMENU = 0x213,
			WM_NOTIFY = 0x4E,
			WM_NOTIFYFORMAT = 0x55,
			WM_NULL = 0x0,
			WM_PAINT = 0xF,
			WM_PAINTCLIPBOARD = 0x309,
			WM_PAINTICON = 0x26,
			WM_PALETTECHANGED = 0x311,
			WM_PALETTEISCHANGING = 0x310,
			WM_PARENTNOTIFY = 0x210,
			WM_PASTE = 0x302,
			WM_PENWINFIRST = 0x380,
			WM_PENWINLAST = 0x38F,
			WM_POWER = 0x48,
			WM_PRINT = 0x317,
			WM_PRINTCLIENT = 0x318,
			WM_QUERYDRAGICON = 0x37,
			WM_QUERYENDSESSION = 0x11,
			WM_QUERYNEWPALETTE = 0x30F,
			WM_QUERYOPEN = 0x13,
			WM_QUERYUISTATE = 0x129,
			WM_QUEUESYNC = 0x23,
			WM_QUIT = 0x12,
			WM_RBUTTONDBLCLK = 0x206,
			WM_RBUTTONDOWN = 0x204,
			WM_RBUTTONUP = 0x205,
			WM_RENDERALLFORMATS = 0x306,
			WM_RENDERFORMAT = 0x305,
			WM_SETCURSOR = 0x20,
			WM_SETFOCUS = 0x7,
			WM_SETFONT = 0x30,
			WM_SETHOTKEY = 0x32,
			WM_SETICON = 0x80,
			WM_SETREDRAW = 0xB,
			WM_SETTEXT = 0xC,
			WM_SETTINGCHANGE = 0x1A,
			WM_SHOWWINDOW = 0x18,
			WM_SIZE = 0x5,
			WM_SIZECLIPBOARD = 0x30B,
			WM_SIZING = 0x214,
			WM_SPOOLERSTATUS = 0x2A,
			WM_STYLECHANGED = 0x7D,
			WM_STYLECHANGING = 0x7C,
			WM_SYNCPAINT = 0x88,
			WM_SYSCHAR = 0x106,
			WM_SYSCOLORCHANGE = 0x15,
			WM_SYSCOMMAND = 0x112,
			WM_SYSDEADCHAR = 0x107,
			WM_SYSKEYDOWN = 0x104,
			WM_SYSKEYUP = 0x105,
			WM_SYSTIMER = 0x118,  // undocumented, see http://support.microsoft.com/?id=108938
			WM_TCARD = 0x52,
			WM_TIMECHANGE = 0x1E,
			WM_TIMER = 0x113,
			WM_UNDO = 0x304,
			WM_UNINITMENUPOPUP = 0x125,
			WM_USER = 0x400,
			WM_USERCHANGED = 0x54,
			WM_VKEYTOITEM = 0x2E,
			WM_VSCROLL = 0x115,
			WM_VSCROLLCLIPBOARD = 0x30A,
			WM_WINDOWPOSCHANGED = 0x47,
			WM_WINDOWPOSCHANGING = 0x46,
			WM_WININICHANGE = 0x1A,
			WM_XBUTTONDBLCLK = 0x20D,
			WM_XBUTTONDOWN = 0x20B,
			WM_XBUTTONUP = 0x20C
		}

		#endregion
	}
}

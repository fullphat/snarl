using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace Snarl.V42
{
	/// <summary>
	/// Snarl C# interface implementation
	/// API version 42
	///
	/// http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Windows_API
	/// https://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API
	///
	/// Written and maintained by Toke Noer Nøttrup (toke@noer.it)
	/// 
	/// Please note the following changes compared to the general (official API) dokumentation:
	///  - Naming of constants and variables are generally changed to follow Microsoft C# standard.
	///    - Grouped variables like SNARL_LAUNCHED, SNARL_QUIT is in enum GlobalEvent.
	///    
	///  - Some functions in the general API takes an appToken as first parameter. This token
	///    is a member variable in C# version, so it is omitted from the functions.
	///    (Always call RegisterApp as first function!)
	///    
	///  - Functions manipulating messages (Update, Hide etc.) still takes a message token as
	///    parameter, but you can get the last message token calling GetLastMsgToken();
	///    Example: snarl.Hide(snarl.GetLastMsgToken());
	///    
	///  - Implements RegisterWithEvents() and event handlers. See SnarlInterfaceExample1 project.
	///
	/// Note on optional parameters:
	///   Since C# 3.0 and prior doesn't have optional parameters, the C# SnarlInterface does not
	///   use this feature! Some functions have overloaded versions though. If no suitable
	///   overload exists, pass null for optional parameters.
	/// 
	/// </summary>
	/// 
	/// <VersionHistory>
	/// 2011-07-25 : Updated register with AppFlags - General update to match SVN rev. 232.
	/// 2011-04-30 : Added GetErrorText().
	/// 2011-04-22 : Implementing events and RegisterWithEvents() + Various small fixes.
	/// 2011-03-13 : Implemented Update()
	///            : Initial event implementation. Needs cleanup.
	/// 2011-02-05 : Removed MessageEvent enum. Only SnarlStatus enum now (same as in VB code)
	/// 2011-02-04 : Updated per rev. 3 of General API documentation.
	///            : Implemented Escape function. Updated some documentation and fixes.
	///            : Should be usable by now!
	/// 2011-02-03 : Initial release of V42 API
	/// </VersionHistory>
	///
	/// <Todo>
	///  - Update documentation all around
	/// </Todo>

	public class SnarlInterface
	{
		#region Public constants and enums

		/// <summary>
		/// Global event identifiers - sent as Windows broadcast messages.
		/// These values appear in wParam of the message.
		/// </summary>
		public enum GlobalEvent
		{
			SnarlLaunched = 1,   // Snarl has just started running
			SnarlQuit = 2,       // Snarl is about to stop running
			SnarlStopped = 3,    // Sent when stopped by user - Also sent to registered window
			SnarlStarted = 4,    // Sent when started by user - Also sent to registered window
		}

		/// <summary>
		/// Snarl status codes.
		/// Containes error codes for function calls, as well as callback values sent by Snarl
		/// to the window specified in Register() when fx. a Snarl notification times out
		/// or the user clicks on it.
		/// </summary>
		public enum SnarlStatus : short
		{
			Success = 0,

			// Snarl-Stopped/Started/UserAway/UserBack is defined in the GlobalEvent struct in VB6 code,
			// but are sent directly to a registered window, so in C# they are defined here instead.
			// Implemented as of Snarl R2.4 Beta3
			SnarlStopped = 3,              // Sent when stopped by user - Also sent as broadcast message
			SnarlStarted,                  // Sent when started by user - Also sent as broadcast message
			SnarlUserAway,                 // Away mode was enabled
			SnarlUserBack,                 // Away mode was disabled

			// Win32 callbacks (renamed under V42)
			CallbackRightClick = 32,       // Deprecated as of V42, ex. SNARL_NOTIFICATION_CLICKED/SNARL_NOTIFICATION_CANCELLED
			CallbackTimedOut,
			CallbackInvoked,               // left clicked and no default callback assigned
			CallbackMenuSelected,          // HIWORD(wParam) contains 1-based menu item index
			CallbackMiddleClick,           // Deprecated as of V42
			CallbackClosed,

			// critical errors
			ErrorFailed = 101,             // miscellaneous failure
			ErrorUnknownCommand,           // specified command not recognised
			ErrorTimedOut,                 // Snarl took too long to respond
			//104 gen critical #4
			//105 gen critical #5
			ErrorBadSocket = 106,          // invalid socket (or some other socket-related error)
			ErrorBadPacket = 107,          // badly formed request
			ErrorInvalidArg = 108,         // arg supplied was invalid (Added in v42.56)
			ErrorArgMissing = 109,         // required argument missing
			ErrorSystem,                   // internal system error
			//120 libsnarl critical block
			ErrorAccessDenied = 121,       // libsnarl only
			//130 SNP/3.0-specific
			ErrorUnsupportedVersion = 131, // requested SNP version is not supported
			ErrorNoActionsProvided,        // empty request
			ErrorUnsupportedEncryption,    // requested encryption type is not supported
			ErrorUnsupportedHashing,       // requested message hashing type is not supported

			// warnings
			ErrorNotRunning = 201,         // Snarl handling window not found
			ErrorNotRegistered,
			ErrorAlreadyRegistered,        // not used yet; sn41RegisterApp() returns existing token
			ErrorClassAlreadyExists,       // not used yet
			ErrorClassBlocked,
			ErrorClassNotFound,
			ErrorNotificationNotFound,
			ErrorFlooding,                 // notification generated by same class within quantum
			ErrorDoNotDisturb,             // DnD mode is in effect was not logged as missed
			ErrorCouldNotDisplay,          // not enough space on-screen to display notification
			ErrorAuthFailure,              // password mismatch
			// Release 2.4.2
			ErrorDiscarded,                // discarded for some reason, e.g. foreground app match
			ErrorNotSubscribed,            // subscriber not found

			// informational
			// code 250 reserved for future use
			WasMerged = 251,               // notification was merged, returned token is the one we merged with

			// callbacks
			// code 300 reserved for future use
			NotifyGone = 301,              // reserved for future use

			// The following are currently specific to SNP 2.0 and are effectively the
			// Win32 SNARL_CALLBACK_nnn constants with 270 added to them

			// SNARL_NOTIFY_CLICK = 302    // indicates notification was right-clicked (deprecated as of V42)
			NotifyExpired = 303,
			NotifyInvoked = 304,           // note this was "ACK" in a previous life
			NotifyMenu,                    // indicates an item was selected from user-defined menu (deprecated as of V42)
			// SNARL_NOTIFY_EX_CLICK       // user clicked the middle mouse button (deprecated as of V42)
			NotifyClosed = 307,            // // user clicked the notification's close gadget (GNTP only)

			// the following is generic to SNP and the Win32 API
			NotifyAction = 308              // user picked an action from the list, the data value will indicate which one
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
			None = 0,
			AppHasPrefs = 1,
			AppHasAbout = 2,
			AppIsWindowless = 0x8000
		}

		/// <summary>Application requests - these values appear in wParam.<para>Application should launch its settings UI</para></summary>
		public static readonly IntPtr AppDoPrefs = new IntPtr(1);
		/// <summary>Application requests - these values appear in wParam.<para>Application should show its About... dialog</para></summary>
		public static readonly IntPtr AppDoAbout = new IntPtr(2);

		public const uint WM_SNARLTEST = (uint)WindowsMessage.WM_USER + 237;

		#endregion

		#region Public events

		public delegate void CallbackEventHandler(SnarlInterface sender, CallbackEventArgs args);
		public delegate void GlobalEventHandler(SnarlInterface sender, GlobalEventArgs args);
		public event CallbackEventHandler CallbackEvent;
		public event GlobalEventHandler GlobalSnarlEvent;

		public class GlobalEventArgs : EventArgs
		{
			public GlobalEvent GlobalEvent { get; set; }

			public GlobalEventArgs(GlobalEvent globalEvent)
			{
				GlobalEvent = globalEvent;
			}
		}

		public class CallbackEventArgs : EventArgs
		{
			public Int32 MessageToken { get; set; }
			public SnarlStatus SnarlEvent { get; set; }

			/// <summary>
			/// The data member if an action callback. Menu index if popup menu.
			/// </summary>
			public UInt16 Parameter { get; set; }

			public CallbackEventArgs(SnarlStatus snarlEvent, UInt16 parameter, int msgToken)
			{
				SnarlEvent = snarlEvent;
				MessageToken = msgToken;
				Parameter = parameter;
			}
		}

		#endregion

		#region Internal constants and enums

		protected const string SnarlWindowClass = "w>Snarl";
		protected const string SnarlWindowTitle = "Snarl";

		protected const string SnarlGlobalMsg = "SnarlGlobalEvent";
		protected const string SnarlAppMsg = "SnarlAppMessage";

		// protected const Int32 WM_DEFAULT_APPMSG = (Int32)WindowsMessage.WM_APP + 0x2fff; // AFFF
		protected const Int32 WM_DEFAULT_APPMSG = (Int32)WindowsMessage.WM_USER + 0x1fff;  // 23FF

		#endregion

		#region Member variables

		private Int32 appToken = 0;
		private Int32 lastMsgToken = 0;
		private String appSignature = "";
		private String password = null;

		// Used for RegisterWithEvents functionality
		private Object instanceLock = new Object();
		private SnarlCallbackNativeWindow callbackWindow = null;
		private Int32 msgReply = 0; // message number Snarl will send to registered window.

		#endregion

		#region Static Snarl functions

		public struct Requests
		{
			public static String AddAction    { get { return "addaction"; } }
			public static String AddClass     { get { return "addclass"; } }
			public static String ClearActions { get { return "clearactions"; } }
			public static String ClearClasses { get { return "clearclasses"; } }
			public static String Hello        { get { return "hello"; } }
			public static String Hide         { get { return "hide"; } }
			public static String IsVisible    { get { return "isvisible"; } }
			public static String Notify       { get { return "notify"; } }
			public static String Register     { get { return "register"; } }
			public static String RemoveClass  { get { return "remclass"; } }
			public static String Unregister   { get { return "unregister"; } }
			public static String UpdateApp    { get { return "updateapp"; } }
			public static String Update       { get { return "update"; } }
			public static String Version      { get { return "version"; } }
		}

		// ------------------------------------------------------------------------------------

		/// <summary>
		/// Send message to Snarl.
		/// Will UTF8 encode the message before sending.
		/// </summary>
		/// <param name="request"></param>
		/// <param name="replyTimeout">(Optional - default = 1000)</param>
		/// <returns>Return zero or positive on succes. Negative on error.</returns>
		static public Int32 DoRequest(String request, UInt32 replyTimeout)
		{
			Int32 nReturn = -1;
			IntPtr nSendMessageResult = IntPtr.Zero;
			IntPtr ptrToUtf8Request = IntPtr.Zero;
			IntPtr ptrToCds = IntPtr.Zero;
			byte[] utf8Request = null;

			// Test if Snarl is running
			IntPtr hWnd = GetSnarlWindow();
			if (!IsWindow(hWnd))
				return -(Int32)SnarlStatus.ErrorNotRunning;

			try
			{
				// Convert to UTF8
				UTF8Encoding utf8 = new UTF8Encoding();
				utf8Request = new byte[utf8.GetMaxByteCount(request.Length)];
				int convertCount = utf8.GetBytes(request, 0, request.Length, utf8Request, 0);

				// Create interop struct
				COPYDATASTRUCT cds = new COPYDATASTRUCT();
				cds.dwData = (IntPtr)0x534E4C03; // "SNL",3
				cds.cbData = convertCount;

				// Create unmanaged byte[] and copy utf8Request into it
				ptrToUtf8Request = Marshal.AllocHGlobal(convertCount);
				Marshal.Copy(utf8Request, 0, ptrToUtf8Request, convertCount);
				cds.lpData = ptrToUtf8Request;

				// Create unmanaged pointer to COPYDATASTRUCT
				ptrToCds = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(COPYDATASTRUCT)));
				Marshal.StructureToPtr(cds, ptrToCds, false);

				if (SendMessageTimeout(hWnd,
						  (uint)WindowsMessage.WM_COPYDATA,
						  (IntPtr)GetCurrentProcessId(),
						  ptrToCds,
						  SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | SendMessageTimeoutFlags.SMTO_NOTIMEOUTIFNOTHUNG,
						  replyTimeout,
						  out nSendMessageResult) == IntPtr.Zero)
				{
					// Error
					int nError = Marshal.GetLastWin32Error();
					if (nError == ERROR_TIMEOUT)
						nReturn = -(Int32)SnarlStatus.ErrorTimedOut;
					else
						nReturn = -(Int32)SnarlStatus.ErrorFailed;
				}
				else
					nReturn = unchecked((Int32)nSendMessageResult.ToInt64()); // Avoid aritmetic overflow error
			}
			finally
			{
				utf8Request = null;
				Marshal.FreeHGlobal(ptrToCds);
				Marshal.FreeHGlobal(ptrToUtf8Request);
			}

			return nReturn;
		}

		static public Int32 DoRequest(String request)
		{
			return DoRequest(request, 1000);
		}

		static public String Escape(String str)
		{
			// Optimize for nothing to escape (all parameters are run through this function!)
			int pos = str.IndexOfAny(new char[] { '=', '&' });

			if (pos != -1)
			{
				// Escape the string
				int len = str.Length;
				char[] charArray = str.ToCharArray();
				StringBuilder sb = new StringBuilder(str.Length + 10); // A random guess at max number of escaped chars

				for (int i = 0; i < str.Length; ++i)
				{
					if (charArray[i] == '=')
						sb.Append("=");
					else if (charArray[i] == '&')
						sb.Append("&");
					sb.Append(charArray[i]);
				}
				return sb.ToString();
			}
			else
				return str;
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
		/// Convenience function for converting a Snarl function result to a string with the error.
		/// </summary>
		/// <param name="callResult"></param>
		/// <returns></returns>
		static public String GetErrorText(int callResult)
		{
			return ((SnarlInterface.SnarlStatus)(Math.Abs(callResult))).ToString();
		}

		/// <summary>
		/// Returns the version of Snarl running.
		/// </summary>
		/// <returns></returns>
		static public Int32 GetVersion()
		{
			return DoRequest(Requests.Version);
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
			return FindWindow(SnarlWindowClass, SnarlWindowTitle);
		}

		/// <summary>
		/// Returns a fully qualified path to Snarl's installation folder.
		/// This is a V39 API method.
		/// </summary>
		/// <returns>Path to Snarl's installation folder. Empty string on failure.</returns>
		static public string GetAppPath()
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
		static public string GetIconsPath()
		{
			return Path.Combine(GetAppPath(), "etc\\icons\\");
		}

		#endregion

		#region SnarlInterface member functions

		/// <summary>
		/// AddAction
		/// </summary>
		/// <param name="notificationToken"></param>
		/// <param name="label"></param>
		/// <param name="command">If using dynamic callback command (@data), use numbers in the range [0, 32767]</param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#addaction"/></remarks>
		public Int32 AddAction(Int32 msgToken, String label, String command)
		{
			// addaction?[token=<notification token>|app-sig=<signature>&uid=<uid>][&password=<password>]&label=<label>&cmd=<command>

			SnarlParameterList spl = new SnarlParameterList(4);
			spl.Add("token", msgToken);
			spl.Add("label", label);
			spl.Add("cmd", command);
			spl.Add("password", password);

			return DoRequest(Requests.AddAction, spl);
		}

		/// <summary>
		/// AddClass
		/// </summary>
		/// <param name="classId"></param>
		/// <param name="name"></param>
		/// <param name="title"></param>
		/// <param name="text"></param>
		/// <param name="icon"></param>
		/// <param name="sound"></param>
		/// <param name="duration"></param>
		/// <param name="callback"></param>
		/// <param name="enabled"></param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#addclass"/></remarks>
		public Int32 AddClass(String classId, String name, String title, String text, String icon, String sound, Int32? duration, String callback, bool enabled)
		{
			// addclass?[app-sig=<signature>|token=<application token>][&password=<password>]&id=<class identifier>&name=<class name>[&enabled=<0|1>][&callback=<callback>]
			//          [&title=<title>][&text=<text>][&icon=<icon>][&sound=<sound>][&duration=<duration>]

			SnarlParameterList spl = new SnarlParameterList(9);
			spl.Add("token", appToken);
			spl.Add("id", classId);
			spl.Add("name", name);
			spl.Add("enabled", enabled);
			spl.Add("callback", callback);
			spl.Add("title", title);
			spl.Add("text", text);
			spl.Add("icon", icon);
			spl.Add("sound", sound);
			spl.Add("duration", duration);
			spl.Add("password", password);

			return DoRequest(Requests.AddClass, spl);
		}

		public Int32 AddClass(String classId, String name)
		{
			return AddClass(classId, name, null, null, null, null, null, null, true);
		}

		/// <summary>
		/// ClearActions
		/// </summary>
		/// <param name="notificationToken"></param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#clearactions"/></remarks>
		public Int32 ClearActions(Int32 msgToken)
		{
			// clearactions?[token=<notification token>|app-sig=<app-sig>&uid=<uid>][&password=<password>]

			SnarlParameterList spl = new SnarlParameterList(2);
			spl.Add("token", msgToken);
			spl.Add("password", password);

			return DoRequest(Requests.ClearActions, spl);
		}

		/// <summary>
		/// ClearClasses
		/// </summary>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#clearclasses"/></remarks>
		public Int32 ClearClasses()
		{
			// clearclasses?[token=app-sig=<signature>|token=<application token>][&password=<password>]

			SnarlParameterList spl = new SnarlParameterList(2);
			spl.Add("token", appToken);
			spl.Add("password", password);

			return DoRequest(Requests.ClearClasses, spl);
		}

		/// <summary>
		/// Hide
		/// </summary>
		/// <param name="msgToken"></param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#hide"/></remarks>
		public Int32 Hide(Int32 msgToken)
		{
			// hide?[token=<notification token>|app-sig=<app-sig>&uid=<uid>][&password=<password>]

			SnarlParameterList spl = new SnarlParameterList(2);
			spl.Add("token", msgToken);
			spl.Add("password", password);

			return DoRequest(Requests.Hide, spl);
		}

		/// <summary>
		/// IsVisible
		/// </summary>
		/// <param name="msgToken"></param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#isvisible"/></remarks>
		public Int32 IsVisible(Int32 msgToken)
		{
			// isvisible?[token=<notification token>|app-sig=<app-sig>&uid=<uid>][&password=<password>]

			SnarlParameterList spl = new SnarlParameterList(2);
			spl.Add("token", msgToken);
			spl.Add("password", password);

			return DoRequest(Requests.IsVisible, spl);
		}

		/// <summary>
		/// Notify
		/// </summary>
		/// <param name="classId">Optional</param>
		/// <param name="title">Optional</param>
		/// <param name="text">Optional</param>
		/// <param name="timeout">Optional</param>
		/// <param name="iconPath">Optional</param>
		/// <param name="iconBase64">Optional</param>
		/// <param name="callback">Optional</param>
		/// <param name="priority">Optional</param>
		/// <param name="uid">Optional</param>
		/// <param name="value">Optional</param>
		/// <returns></returns>
		/// <remarks>
		/// All parameters are optional. Pass null to use class default values.
		/// <para><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#notify"/></para>
		/// </remarks>
		public Int32 Notify(String classId, String title, String text, Int32? timeout, String iconPath, String iconBase64, MessagePriority? priority, String uid, String callback, String value)
		{
			// notify?[app-sig=<signature>|token=<application token>][&password=<password>][&id=<class identifier>]
			//        [&title=<title>][&text=<text>][&timeout=<timeout>][&icon=<icon path>][&icon-base64=<MIME data>][&callback=<default callback>]
			//        [&priority=<priority>][&uid=<notification uid>][&value=<value>]

			SnarlParameterList spl = new SnarlParameterList(12);
			spl.Add("token", appToken);
			spl.Add("password", password);

			spl.Add("id", classId);
			spl.Add("title", title);
			spl.Add("text", text);
			spl.Add("timeout", timeout);
			spl.Add("icon", iconPath);
			spl.Add("icon-base64", iconBase64);
			spl.Add("callback", callback);
			spl.Add("priority", (Int32?)priority);
			spl.Add("uid", uid);
			spl.Add("value", value);

			lastMsgToken = DoRequest(Requests.Notify, spl);

			return lastMsgToken;
		}

		public Int32 Notify(String classId, String title, String text, Int32? timeout, String iconPath, String iconBase64, MessagePriority? priority)
		{
			return Notify(classId, title, text, timeout, iconPath, iconBase64, priority, null, null, null);
		}

		public Int32 Notify(String classId, String title, String text, Int32? timeout, String iconPath, String iconBase64)
		{
			return Notify(classId, title, text, timeout, iconPath, iconBase64, null, null, null, null);
		}

		public Int32 Notify(String classId, String title, String text, Int32? timeout)
		{
			return Notify(classId, title, text, timeout, null, null, null, null, null, null);
		}

		public Int32 Notify(String classId, String title, String text)
		{
			return Notify(classId, title, text, null, null, null, null, null, null, null);
		}

		/// <summary>
		/// Register application with Snarl.
		/// </summary>
		/// <param name="signatur"></param>
		/// <param name="title"></param>
		/// <param name="icon">Optional (null)</param>
		/// <param name="password">Optional (null)</param>
		/// <param name="hWndReply">Optional (IntPtr.Zero)</param>
		/// <param name="msgReply">Optional (0)</param>
		/// <param name="flags">Optional (AppFlags.None)</param>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#register"/></remarks>
		public Int32 Register(String signature, String title, String icon, String password, IntPtr hWndReplyTo, Int32 msgReply, AppFlags flags)
		{
			// register?app-sig=<signature>&title=<title>[&icon=<icon>][&password=<password>][&reply-to=<reply window>][&reply=<reply message>]

			SnarlParameterList spl = new SnarlParameterList(7);
			spl.Add("app-sig", signature);
			spl.Add("title", title);
			spl.Add("icon", icon);
			spl.Add("password", password);
			spl.Add("reply-to", hWndReplyTo);
			spl.Add("reply", msgReply);
			spl.Add("flags", (int)flags);

			// If password was given, save and use in all other functions requiring password
			if (!String.IsNullOrEmpty(password))
				this.password = password;

			Int32 request = DoRequest(Requests.Register, spl);
			if (request > 0)
			{
				this.appToken = request;
				this.msgReply = msgReply;
				this.appSignature = signature;
			}

			return request;
		}

		public Int32 Register(String signature, String title, String icon, String password)
		{
			return Register(signature, title, icon, password, IntPtr.Zero, 0, AppFlags.None);
		}

		public Int32 Register(String signature, String title, String icon)
		{
			return Register(signature, title, icon, null, IntPtr.Zero, 0, AppFlags.None);
		}

		/// <summary>
		/// Register with Snarl and hook the window supplied in hWndReplyTo, enabling SnarlInterface
		/// to send the CallbackEvent and GlobalSnarlEvent.
		/// </summary>
		/// <param name="signature"></param>
		/// <param name="title"></param>
		/// <param name="icon"></param>
		/// <param name="password"></param>
		/// <param name="hWndReplyTo">The HWND of the window which should be hooked. If IntPtr.Zero a new listening Window will be created.</param>
		/// <param name="msgReply">The message Snarl should send back to the window on callbacks. If null and internal default value will be used.</param>
		/// <returns></returns>
		/// <remarks>This is not part of the official API.</remarks>
		public Int32 RegisterWithEvents(String signature, String title, String icon, String password, IntPtr hWndReplyTo, Int32? msgReply, AppFlags flags)
		{
			if (msgReply == null || msgReply.Value == 0)
				msgReply = WM_DEFAULT_APPMSG;

			lock (instanceLock)
			{
				UnregisterCallbackWindow();
				callbackWindow = new SnarlCallbackNativeWindow(this, hWndReplyTo);
			}

			return Register(signature, title, icon, password, callbackWindow.Handle, msgReply.Value, flags);
		}

		/// <summary>
		/// Registers with Snarl and creates a new Window listening for messages from Snarl.
		/// This enables the use of the CallbackEvent and GlobalSnarlEvent events.
		/// </summary>
		public Int32 RegisterWithEvents(String signature, String title, String icon, String password)
		{
			return RegisterWithEvents(signature, title, icon, password, IntPtr.Zero, null, AppFlags.None);
		}

		/// <summary>
		/// RemoveClass
		/// </summary>
		/// <param name="classID"></param>
		/// <returns></returns>
		/// <seealso cref="ClearClasses"/>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#remclass"/></remarks>
		public Int32 RemoveClass(String classID)
		{
			// remclass?[app-sig=<signature>|token=<application token>][&password=<password>][&id=<class identifier>|&all=<0|1>]

			SnarlParameterList spl = new SnarlParameterList(3);
			spl.Add("token", appToken);
			spl.Add("id", classID);
			spl.Add("password", password);
			// spl.Add("all", password); // Use ClearClasses

			return DoRequest(Requests.RemoveClass, spl);
		}

		/// <summary>
		/// Unregister application.
		/// </summary>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#unregister"/></remarks>
		public Int32 Unregister()
		{
			// unregister?[app-sig=<signature>|token=<application token>][&password=<password>]

			SnarlParameterList spl = new SnarlParameterList(2);
			spl.Add("app-sig", appSignature);
			spl.Add("password", password);

			appToken = 0;
			lastMsgToken = 0;
			appSignature = "";
			password = null;
			msgReply = 0;

			UnregisterCallbackWindow(); // Will only do work if RegisterWithEvents has been called

			return DoRequest(Requests.Unregister, spl);
		}

		/// <summary>
        /// Updates a Snarl notification already visible.
		/// </summary>
		/// <returns></returns>
		/// <remarks><see cref="http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#update"/></remarks>
		public Int32 Update(Int32 msgToken, String classId, String title, String text, Int32? timeout, String iconPath, String iconBase64, MessagePriority? priority, String callback, String value)
		{
			// Made from best guess - no documentation available yet
			// Following parameters left out: "reply-to", "reply", "uid"
			SnarlParameterList spl = new SnarlParameterList(11);
			spl.Add("token", msgToken);
			spl.Add("password", password);

			spl.Add("id", classId);
			spl.Add("title", title);
			spl.Add("text", text);
			spl.Add("icon", iconPath);
			spl.Add("icon-base64", iconBase64);
			spl.Add("callback", callback);
			spl.Add("value", value);
			spl.Add("timeout", timeout);
			spl.Add("priority", (Int32?)priority);

			return DoRequest(Requests.Update, spl);
		}

		public Int32 Update(Int32 msgToken, String classId, String title, String text, Int32? timeout, String iconPath, String iconBase64, MessagePriority? priority)
		{
			return Update(msgToken, classId, title, text, timeout, iconPath, iconBase64, priority, null, null);
		}

		public Int32 Update(Int32 msgToken, String classId, String title, String text, Int32? timeout)
		{
			return Update(msgToken, classId, title, text, timeout, null, null, null, null, null);
		}

		public Int32 Update(Int32 msgToken, String classId, String title, String text)
		{
			return Update(msgToken, classId, title, text, null, null, null, null, null, null);
		}

		/// <summary>
		/// GetLastMsgToken() returns token of the last message sent to Snarl.
		/// </summary>
		/// <returns>Token returned from Snarl on the last call to <see cref="Notify"/>.</returns>
		public Int32 GetLastMsgToken()
		{
			return lastMsgToken;
		}

		#endregion

		#region Private functions

		/// <summary>
		/// Internal helper function for constructing the Snarl messages
		/// </summary>
		/// <param name="request"></param>
		/// <param name="spl"></param>
		/// <param name="replyTimeout"></param>
		/// <returns></returns>
		static private Int32 DoRequest(String request, SnarlParameterList spl, UInt32 replyTimeout)
		{
			// <action>[?<data>=<value>[&<data>=<value>]]

			if (spl.ParamList.Count > 0)
			{
				StringBuilder sb = new StringBuilder(request);
				sb.Append("?");

				foreach (KeyValuePair<String, String> kvp in spl.ParamList)
				{
					if (kvp.Value != null && kvp.Value.Length > 0)
					{
						sb.AppendFormat("{0}={1}&", kvp.Key, Escape(kvp.Value));
					}
				}

				// Delete last &
				sb.Remove(sb.Length - 1, 1);

				return DoRequest(sb.ToString(), replyTimeout);
			}
			else
				return DoRequest(request, replyTimeout);
		}

		static private Int32 DoRequest(String request, SnarlParameterList spl)
		{
			return DoRequest(request, spl, 1000);
		}

		/// <summary>
		/// Releases the callback window.
		/// </summary>
		private void UnregisterCallbackWindow()
		{
			if (callbackWindow != null)
			{
				lock (instanceLock)
				{
					if (callbackWindow != null)
					{
						callbackWindow.Detach();
						callbackWindow = null;
					}
				}
			}
		}

		/// <summary>
		/// Creates a new window or hook a window and receives messages.
		/// </summary>
		protected class SnarlCallbackNativeWindow : System.Windows.Forms.NativeWindow
		{
			private readonly bool hasCreatedCallbackWindow;
			private readonly UInt32 snarlGlobalMessage;
			private readonly SnarlInterface parent;


			public SnarlCallbackNativeWindow(SnarlInterface parent, IntPtr windowHandle)
			{
				this.parent = parent;
				snarlGlobalMessage = SnarlInterface.Broadcast();

				if (windowHandle == IntPtr.Zero)
				{
					hasCreatedCallbackWindow = true;

					CreateParams cp = new CreateParams();
					cp.Caption = GetType().FullName;
					CreateHandle(cp);
				}
				else
				{
					hasCreatedCallbackWindow = false;
					AssignHandle(windowHandle);
				}
			}

			public void Detach()
			{
				if (hasCreatedCallbackWindow)
					DestroyHandle();
				else
					ReleaseHandle();
			}

			protected override void WndProc(ref System.Windows.Forms.Message m)
			{
				if (m.Msg == snarlGlobalMessage)
				{
					if (parent.GlobalSnarlEvent == null)
						return;

					GlobalEventArgs eventArgs = new GlobalEventArgs((SnarlInterface.GlobalEvent)(m.WParam.ToInt64() & 0xffffffff));
					parent.GlobalSnarlEvent(parent, eventArgs);
				}
				else if (m.Msg == parent.msgReply)
				{
					if (parent.CallbackEvent == null)
						return;

					// Parse out parameters
					UInt16 loword, hiword;
					ConvertToUInt16(m.WParam, out loword, out hiword);
					Int32 msgToken = m.LParam.ToInt32();

					CallbackEventArgs eventArgs = new CallbackEventArgs((SnarlStatus)loword, hiword, msgToken);

					parent.CallbackEvent(parent, eventArgs);
				}

				base.WndProc(ref m);
			}

			protected void ConvertToUInt16(IntPtr input, out UInt16 outLow, out UInt16 outHigh)
			{
				// Assumes only 32 bit
				UInt32 tmp = (UInt32)input.ToInt64();

				outLow = (UInt16)(tmp & 0x0000ffff);
				outHigh = (UInt16)(tmp >> 16);
			}
		}

		/// <summary>
		/// Helper class used with internal DoRequest()
		/// </summary>
		protected class SnarlParameterList
		{
			protected List<KeyValuePair<String, String>> list = new List<KeyValuePair<String, String>>();
			public IList<KeyValuePair<String, String>> ParamList
			{
				get { return list; }
			}

			public SnarlParameterList()
			{
			}

			public SnarlParameterList(int initialCapacity)
			{
				list.Capacity = initialCapacity;
			}

			public void Add(String key, String value)
			{
				if (value != null)
					list.Add(new KeyValuePair<String, String>(key, value));
			}

			public void Add(String key, bool value)
			{
				list.Add(new KeyValuePair<String, String>(key, value ? "1" : "0"));
			}

			public void Add(String key, Int32? value)
			{
				if (value != null && value.HasValue)
					list.Add(new KeyValuePair<String, String>(key, value.Value.ToString()));
			}

			public void Add(String key, IntPtr value)
			{
				if (value != null)
					list.Add(new KeyValuePair<String, String>(key, value.ToString()));
			}
		}

		#endregion

		#region Interop imports and structures

		[DllImport("user32.dll", SetLastError = false)]
		internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", SetLastError = false)]
		internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, Int32 nMaxCount);

		[DllImport("user32.dll", SetLastError = false, CharSet = CharSet.Auto)]
		internal static extern uint RegisterWindowMessage(string lpString);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		internal static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

		[DllImport("user32.dll")]
		internal static extern bool IsWindow(IntPtr hWnd);

		// [DllImport("user32.dll", SetLastError = false, CharSet = CharSet.Auto)]
		// internal static extern IntPtr GetProp(IntPtr hWnd, string lpString);

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
			public Int32 cbData;   // DWORD
			public IntPtr lpData;   // PVOID
		}

		private const int ERROR_TIMEOUT = 1460;

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

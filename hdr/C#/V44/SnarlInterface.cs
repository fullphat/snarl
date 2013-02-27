using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Snarl.V44
{
    /// <summary>
    /// Snarl C# interface implementation
    /// API version 44 (Snarl 3.0)
    ///
    /// https://sites.google.com/site/snarlapp/developers/api-reference
    /// https://sites.google.com/site/snarlapp/win32-api
    ///
    /// Written and maintained by Toke Noer Nøttrup (toke@noer.it)
    /// 
    /// Please note the following changes compared to the general (official API) documentation:
    ///  - Naming of constants and variables are generally changed to follow Microsoft C# standard.
    ///    - Grouped variables like SNARL_LAUNCHED, SNARL_QUIT is in enum SnarlGlobalEvent.
    ///    
    ///  - Some functions in the general API takes an appToken/app-sig as first parameter. This token
    ///    is a member variable in C# version, so it is omitted from the functions.
    ///    (Always call Register/RegisterWithEvents as first function!)
    ///    
    ///  - Implements RegisterWithEvents() and event handlers. See SnarlInterfaceWinFormExample_V44 project.
    /// </summary>
    /// 
    /// Changes from V42 to V44:
    ///  - Removed the old "token" style stuff - now uses app-sig and unique message id's instead.
    ///    This also means IsVisible() and Update() is removed - use Notify and specify a uId parameter instead.
    ///  - Removed AddAction (deprecated by Snarl) - Notify() takes an array of SnarlActions instead.
    ///  - Removed various overloads - replaced with C# 4.0 optional arguments instead (requires .NET 4.0)
    ///  - Renamed some functions
    ///    (Broadcast -> GetBroadcastMessage, AppMsg -> GetAppMsg, GetAppPath -> GetSnarlPath,
    ///  - Moved most of the inner classes out of SnarlInterface - this does give a bit more work for users when upgrading, sorry.
    ///    (Upside is less typing - ie. SnarlInterface.GlobalEvent.SnarlStopped becomes SnarlGlobalEvent.SnarlStopped)
    /// <VersionHistory>
    /// 2012-08-14 : Initial release of V44 API
    /// </VersionHistory>
    public class SnarlInterface
    {
        #region Public events

        public delegate void CallbackEventHandler(object sender, SnarlCallbackEventArgs e);
        public delegate void GlobalEventHandler(object sender, SnarlGlobalEventArgs e);
        public event CallbackEventHandler CallbackEvent;
        public event GlobalEventHandler GlobalSnarlEvent;

        #endregion

        public const uint WM_SNARLTEST = (uint)WindowsMessage.WM_USER + 237;

        #region Internal constants and enums

        protected const string SnarlWindowClass = "w>Snarl";
        protected const string SnarlWindowTitle = "Snarl";

        protected const string SnarlGlobalMsg = "SnarlGlobalEvent";
        protected const string SnarlAppMsg = "SnarlAppMessage";

        protected const Int32 WM_DEFAULT_APPMSG = (Int32)WindowsMessage.WM_USER + 0x1fff;  // 23FF

        #endregion

        #region Member variables

        private String appSignature = null;
        private String password = null;

        // Used for RegisterWithEvents functionality
        private Object instanceLock = new Object();
        private SnarlCallbackNativeWindow callbackWindow = null;
        private Int32 msgReply = 0; // message number Snarl will send to registered window.

        #endregion

        #region Properties

        public String AppSignature
        {
            get { return appSignature; }
            set { appSignature = value; }
        }

        public String Password
        {
            get { return password; }
            set { password = value; }
        }

        #endregion

        #region Static Snarl functions

        // ------------------------------------------------------------------------------------

        /// <summary>
        /// Send message to Snarl.
        /// Will UTF8 encode the message before sending.
        /// </summary>
        /// <param name="request">The Snarl request key/value pair strings.</param>
        /// <param name="replyTimeout">Timeout before returning an error if Snarl is too slow to respond.</param>
        /// <returns>Return zero or positive on success. Negative on error.</returns>
        public static Int32 DoRequest(String request, UInt32 replyTimeout = 1000)
        {
            Int32 nReturn = -1;
            IntPtr nSendMessageResult = IntPtr.Zero;
            IntPtr ptrToUtf8Request = IntPtr.Zero;
            IntPtr ptrToCds = IntPtr.Zero;
            byte[] utf8Request = null;

            // Test if Snarl is running
            IntPtr hWnd = GetSnarlWindow();
            if (!NativeMethods.IsWindow(hWnd))
                return -(int)SnarlStatus.ErrorNotRunning;

            try
            {
                // Convert to UTF8
                UTF8Encoding utf8 = new UTF8Encoding();
                utf8Request = new byte[utf8.GetMaxByteCount(request.Length)];
                int convertCount = utf8.GetBytes(request, 0, request.Length, utf8Request, 0);

                // Create interop struct
                var cds = new NativeMethods.COPYDATASTRUCT();
                cds.dwData = (IntPtr)0x534E4C03; // "SNL",3
                cds.cbData = convertCount;

                // Create unmanaged byte[] and copy utf8Request into it
                ptrToUtf8Request = Marshal.AllocHGlobal(convertCount);
                Marshal.Copy(utf8Request, 0, ptrToUtf8Request, convertCount);
                cds.lpData = ptrToUtf8Request;

                // Create unmanaged pointer to COPYDATASTRUCT
                ptrToCds = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NativeMethods.COPYDATASTRUCT)));
                Marshal.StructureToPtr(cds, ptrToCds, false);

                if (NativeMethods.SendMessageTimeout(hWnd,
                          (uint)WindowsMessage.WM_COPYDATA,
                          (IntPtr)NativeMethods.GetCurrentProcessId(),
                          ptrToCds,
                          NativeMethods.SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | NativeMethods.SendMessageTimeoutFlags.SMTO_NOTIMEOUTIFNOTHUNG,
                          replyTimeout,
                          out nSendMessageResult) == IntPtr.Zero)
                {
                    // Error
                    int nError = Marshal.GetLastWin32Error();
                    if (nError == NativeMethods.ERROR_TIMEOUT)
                        nReturn = -(Int32)SnarlStatus.ErrorTimedOut;
                    else
                        nReturn = -(Int32)SnarlStatus.ErrorFailed;
                }
                else
                {
                    // SendMessage success
                    nReturn = unchecked((Int32)nSendMessageResult.ToInt64()); // Avoid arithmetic overflow error
                }
            }
            finally
            {
                utf8Request = null;
                Marshal.FreeHGlobal(ptrToCds);
                Marshal.FreeHGlobal(ptrToUtf8Request);
            }

            return nReturn;
        }

        /// <summary>
        /// Constructs a request string from the request parameter and the given parameter list
        /// and sends the action request to Snarl.
        /// </summary>
        /// <param name="action">The Snarl action to perform - ie. notify. <seealso cref="SnarlRequests"/>.</param>
        /// <param name="spl">List of key/value pairs with parameters for the request.</param>
        /// <param name="replyTimeout">(Optional - default = 1000 milliseconds)</param>
        /// <returns>Return zero or positive on success. Negative on error.</returns>
        public static Int32 DoRequest(String action, SnarlParameterList spl, UInt32 replyTimeout = 1000)
        {
            // Creates string of format: <action>[?<data>=<value>[&<data>=<value>]]
            // and sends the String through other DoRequest() overload.

            if (spl.ParamList.Count > 0)
            {
                StringBuilder sb = new StringBuilder(action);
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
            {
                return DoRequest(action, replyTimeout);
            }
        }

        public static String Escape(String str)
        {
            // Optimize for nothing to escape (all parameters are run through this function!)
            int pos = str.IndexOfAny(new char[] { '=', '&' });
            if (pos == -1)
                return str;

            // str contains symbols that needs to be escaped
            int len = str.Length;
            char[] charArray = str.ToCharArray();
            StringBuilder sb = new StringBuilder(len + 10); // A "guess" at max number of escaped chars

            for (int i = 0; i < len; ++i)
            {
                if (charArray[i] == '=')
                    sb.Append("=");
                else if (charArray[i] == '&')
                    sb.Append("&");
                sb.Append(charArray[i]);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Returns true if Snarl system was found running.
        /// </summary>
        public static bool IsSnarlRunning()
        {
            return NativeMethods.IsWindow(GetSnarlWindow());
        }

        /// <summary>
        /// Convenience function for converting a Snarl function result to a string with the error.
        /// </summary>
        /// <param name="callResult">The result returned by <see cref="DoRequest"/>.</param>
        /// <exception cref="ArgumentException">Throws if error code is not found in <see cref="StatusCode"/> enum.</exception>
        public static String GetErrorText(int callResult)
        {
            String name = Enum.GetName(typeof(SnarlStatus), Math.Abs(callResult));
            if (name == null)
                throw new ArgumentException("Error text could not be found for callResult.");
            
            return name;
        }

        /// <summary>
        /// Returns the system version number of Snarl.
        /// </summary>
        public static Int32 GetVersion()
        {
            return DoRequest(SnarlRequests.Version);
        }

        /// <summary>
        /// Returns the value of Snarl's global registered message.
        /// Notes:
        ///   Snarl registers SNARL_GLOBAL_MSG during startup which it then uses to communicate
        ///   with all running applications through a Windows broadcast message. This function can
        ///   only fail if for some reason the Windows RegisterWindowMessage() function fails
        ///   - given this, this function *cannot* be used to test for the presence of Snarl.
        /// </summary>
        /// <returns>A 16-bit value (translated to 32-bit) which is the registered Windows message for Snarl.</returns>
        public static uint GetBroadcastMessage()
        {
            return NativeMethods.RegisterWindowMessage(SnarlGlobalMsg);
        }

        /// <summary>
        /// Returns the global Snarl Application message.
        /// </summary>
        /// <returns>Snarl Application registered message.</returns>
        public static uint GetAppMsg()
        {
            return NativeMethods.RegisterWindowMessage(SnarlAppMsg);
        }

        /// <summary>
        /// Returns a handle to the Snarl Dispatcher window.
        /// Notes:
        ///   This is now the preferred way to test if Snarl is actually running.
        /// </summary>
        /// <returns>Returns handle to Snarl Dispatcher window, or zero if it's not found</returns>
        public static IntPtr GetSnarlWindow()
        {
            return NativeMethods.FindWindow(SnarlWindowClass, SnarlWindowTitle);
        }

        /// <summary>
        /// Returns a fully qualified path to Snarl's installation folder.
        /// </summary>
        /// <returns>Path to Snarl's installation folder. Empty string on failure.</returns>
        public static string GetSnarlPath()
        {
            StringBuilder sb = new StringBuilder(512);

            IntPtr hwnd = GetSnarlWindow();
            if (hwnd != IntPtr.Zero)
            {
                IntPtr hWndPath = NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "static", null);
                if (hWndPath != IntPtr.Zero)
                {
                    int result = NativeMethods.GetWindowText(hWndPath, sb, 512);
                    if (result == 0)
                        throw new SnarlException("GetWindowText native function returned 0.");
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Returns a fully qualified path to Snarl's default icon location.
        /// </summary>
        /// <returns>Path to Snarl's default icon location. Empty string on failure.</returns>
        public static string GetIconsPath()
        {
            return Path.Combine(GetSnarlPath(), "etc\\icons\\");
        }

        #endregion

        #region SnarlInterface member functions

        /// <summary>
        /// Adds a notification class to a previously registered application.
        /// </summary>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-addclass"/></remarks>
        public Int32 AddClass(String classId, String className = null,
                              String title = null, String text = null, String icon = null,
                              String sound = null, Int32? duration = null, String callback = null, bool enabled = true)
        {
            SnarlParameterList spl = new SnarlParameterList(11);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);

            spl.Add("class-id", classId);
            spl.Add("class-name", className);
            spl.Add("enabled", enabled);
            spl.Add("callback", callback);
            spl.Add("title", title);
            spl.Add("text", text);
            spl.Add("icon", icon);
            spl.Add("sound", sound);
            spl.Add("duration", duration);

            return DoRequest(SnarlRequests.AddClass, spl);
        }

        /// <summary>
        /// Removes all actions associated with the specified notification.
        /// </summary>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-clearactions"/></remarks>
        public Int32 ClearActions(String uId)
        {
            SnarlParameterList spl = new SnarlParameterList(3);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);
            spl.Add("uid ", uId);

            return DoRequest(SnarlRequests.ClearActions, spl);
        }

        /// <summary>
        /// Removes all classes associated with a particular application.
        /// </summary>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-clearclasses"/></remarks>
        public Int32 ClearClasses()
        {
            SnarlParameterList spl = new SnarlParameterList(2);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);

            return DoRequest(SnarlRequests.ClearClasses, spl);
        }

        /// <summary>
        /// Removes the specified notification from the screen or missed list.
        /// </summary>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-hide"/></remarks>
        public Int32 Hide(String uId)
        {
            SnarlParameterList spl = new SnarlParameterList(2);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);
            spl.Add("uid", uId);

            return DoRequest(SnarlRequests.Hide, spl);
        }

        /// <summary>
        /// Displays a notification.
        /// </summary>
        /// <param name="uid">(Optional)A unique identifier for the notification so the sending application can track events related to it and update it if required. </param>
        /// <param name="classId">(Optional)The identifier of the class to use.</param>
        /// <param name="title">(Optional)Notification title.</param>
        /// <param name="text">(Optional) Notification body text.</param>
        /// <param name="timeout">(Optional)The amount of time, in seconds, the notification should remain on screen.</param>
        /// <param name="icon">(Optional) The icon to use, see the notes below for details of supported formats.</param>
        /// <param name="iconBase64">(Optional) Base64-encoded bytes to be used as the icon.</param>
        /// <param name="sound">(Optional) The path to a sound file to play.</param>
        /// <param name="callback">(Optional) Callback to be invoked when the user clicks on the main body area of the notification.</param>
        /// <param name="priority">(Optional) The urgency of the notification, <see cref="SnarlMessagePriority"/> and Priorities in the Developer Guide for more information on how Snarl deals with different priorities. </param>
        /// <param name="sensitivity">(Optional) The sensitivity of the notification.  See the Sensitivity section in the Developer Guide for more information.</param>
        /// <param name="value-percent">(Optional) A decimal percent value to be included with the notification.  Certain styles can display this value as a meter or other visual representation.  See Custom Values in the Developer Guide for more information.</param>
        /// <param name="action">Optional) Actions to add to the notification - <see cref="ActionCollection"/> and see Action section in the Developer Guide for more information.</param>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-notify"/></remarks>
        public Int32 Notify(String uId = null, String classId = null, String title = null, String text = null,
                            Int32? timeout = null, String icon = null, String iconBase64 = null,
                            String sound = null, String callback = null, String callbackLabel = null, SnarlMessagePriority? priority = null,
                            String sensitivity = null, decimal? valuePercent = null, SnarlAction[] actions = null)
        {
            if (appSignature == null)
                throw new InvalidOperationException("AppSignature is null - the applications is properly not registered correctly with Snarl.");

            // Build parameter list
            int paramListSize = 15 + (actions != null ? actions.Length : 0);
            SnarlParameterList spl = new SnarlParameterList(paramListSize);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);

            spl.Add("uid", uId);
            spl.Add("id", classId);
            spl.Add("title", title);
            spl.Add("text", text);
            spl.Add("timeout", timeout);
            spl.Add("icon", icon);
            spl.Add("icon-base64", iconBase64);
            spl.Add("sound", sound);
            spl.Add("callback", callback);
            spl.Add("callback-label", callbackLabel);
            spl.Add("priority", (Int32?)priority);
            spl.Add("sensitivity", sensitivity);
            spl.Add("value-percent", sensitivity);
            spl.Add(actions);

            return DoRequest(SnarlRequests.Notify, spl);
        }

        /// <summary>
        /// Registers an application with Snarl.
        /// </summary>
        /// <param name="appSignatur">The application's signature.</param>
        /// <param name="title">The application's name.</param>
        /// <param name="password">Password used during registration.</param>
        /// <param name="icon">(Optional) The icon to use.</param>
        /// <param name="hWndReply">(Optional) The handle to a window (HWND) which Snarl should post events to. <seealso cref="RegisterWithEvents"/></param>
        /// <param name="msgReply">(Optional) The message Snarl should use to post to the hWndReply Window.</param>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-register"/></remarks>
        public Int32 Register(String appSignature, String title, String password,
                              String icon = null, IntPtr hWndReplyTo = default(IntPtr), Int32 msgReply = 0)
        {
            if (String.IsNullOrEmpty(appSignature))
                throw new ArgumentNullException("appSignature");
            if (String.IsNullOrEmpty(title))
                throw new ArgumentNullException("title");
            if (String.IsNullOrEmpty(password))
                throw new ArgumentNullException("password");

            SnarlParameterList spl = new SnarlParameterList(7);
            spl.Add("app-sig", appSignature);
            spl.Add("title", title);
            spl.Add("password", password);

            spl.Add("icon", icon);
            spl.Add("reply-to", hWndReplyTo);
            spl.Add("reply", msgReply);
            spl.Add("flags", 0);


            Int32 request = DoRequest(SnarlRequests.Register, spl);
            if (request > 0)
            {
                this.appSignature = appSignature;
                this.password = password;
                this.msgReply = msgReply;
            }

            return request;
        }

        /// <summary>
        /// Register with Snarl and hook the window supplied in hWndReplyTo, enabling SnarlInterface
        /// to send the CallbackEvent and GlobalSnarlEvent.
        /// </summary>
        /// <param name="appSignatur">The application's signature.</param>
        /// <param name="title">The application's name.</param>
        /// <param name="password">Password used during registration.</param>
        /// <param name="icon">(Optional) The icon to use.</param>
        /// <param name="hWndReply">(Optional) The handle to a window (HWND) which Snarl should post events to.
        ///     If parameter is eqaul to IntPtr.Zero or omitted a new listening Window will be created.</param>
        /// <param name="msgReply">(Optional) The message Snarl should use to post to the hWndReply Window. If 0 or omitted a default value will be used.</param>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-register"/></remarks>
        /// <remarks>This is not part of the official API, but enables the SnarlInterface class to send events.</remarks>
        public Int32 RegisterWithEvents(String appSignatur, String title, String password,
                                        String icon = null, IntPtr hWndReplyTo = default(IntPtr), Int32 msgReply = 0)
        {
            if (msgReply == 0)
                msgReply = WM_DEFAULT_APPMSG;

            lock (instanceLock)
            {
                UnregisterCallbackWindow();
                callbackWindow = new SnarlCallbackNativeWindow(this, hWndReplyTo);
            }

            return Register(appSignatur, title, password, icon, callbackWindow.Handle, msgReply);
        }

        /// <summary>
        /// Removes a particular notification class from a registered application.
        /// </summary>
        /// <seealso cref="ClearClasses"/>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-remclass"/></remarks>
        public Int32 RemoveClass(String classId)
        {
            SnarlParameterList spl = new SnarlParameterList(3);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);
            spl.Add("id", classId);
            
            return DoRequest(SnarlRequests.RemoveClass, spl);
        }

        /// <summary>
        /// Unregisters the application.
        /// </summary>
        /// <remarks><see cref="https://sites.google.com/site/snarlapp/developers/api-reference#TOC-unregister"/></remarks>
        public Int32 Unregister()
        {
            SnarlParameterList spl = new SnarlParameterList(2);
            spl.Add("app-sig", appSignature);
            spl.Add("password", password);

            appSignature = null;
            password = null;
            msgReply = 0;

            UnregisterCallbackWindow(); // Will only do work if RegisterWithEvents has been called

            return DoRequest(SnarlRequests.Unregister, spl);
        }

        #endregion

        #region Private functions

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
        /// Creates a new window or hook a window to receive Windows messages.
        /// </summary>
        private sealed class SnarlCallbackNativeWindow : System.Windows.Forms.NativeWindow
        {
            private readonly bool hasCreatedCallbackWindow;
            private readonly UInt32 snarlGlobalMessage;
            private readonly SnarlInterface parent;


            public SnarlCallbackNativeWindow(SnarlInterface parent, IntPtr windowHandle)
            {
                this.parent = parent;
                snarlGlobalMessage = SnarlInterface.GetBroadcastMessage();

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

                    SnarlGlobalEventArgs eventArgs = new SnarlGlobalEventArgs((SnarlGlobalEvent)(m.WParam.ToInt64() & 0xffffffff));
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

                    SnarlCallbackEventArgs eventArgs = new SnarlCallbackEventArgs((SnarlStatus)loword, hiword, msgToken);

                    parent.CallbackEvent(parent, eventArgs);
                }

                base.WndProc(ref m);
            }

            private void ConvertToUInt16(IntPtr input, out UInt16 outLow, out UInt16 outHigh)
            {
                // Assumes only 32 bit
                UInt32 tmp = (UInt32)input.ToInt64();

                outLow = (UInt16)(tmp & 0x0000ffff);
                outHigh = (UInt16)(tmp >> 16);
            }
        }

        #endregion

        #region Interop imports and structures

        protected static class NativeMethods
        {
            public const int ERROR_TIMEOUT = 1460;

            [DllImport("user32.dll", SetLastError = false, CharSet = CharSet.Unicode)]
            internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", SetLastError = false, CharSet = CharSet.Unicode)]
            internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, Int32 nMaxCount);

            [DllImport("user32.dll", SetLastError = false, CharSet = CharSet.Unicode)]
            internal static extern uint RegisterWindowMessage(string lpString);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            internal static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
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

    /// <summary>
    /// Requests supported by Snarl - used for custom building commands.
    /// </summary>
    public struct SnarlRequests
    {
        public static String AddAction     { get { return "addaction"; } }
        public static String AddClass      { get { return "addclass"; } }
        public static String ClearActions  { get { return "clearactions"; } }
        public static String ClearClasses  { get { return "clearclasses"; } }
        public static String Hello         { get { return "hello"; } }
        public static String Hide          { get { return "hide"; } }
        public static String IsVisible     { get { return "isvisible"; } }
        public static String Notify        { get { return "notify"; } }
        public static String Register      { get { return "register"; } }
        public static String RemoveClass   { get { return "remclass"; } }
        // public static String Subscribe  { get { return "subscribe"; } } Not available for Win32 API
        public static String Test          { get { return "test"; } }
        public static String Unregister    { get { return "unregister"; } }
        public static String UpdateApp     { get { return "updateapp"; } }
        public static String Update        { get { return "update"; } }
        public static String Version       { get { return "version"; } }
    }

    /// <summary>
    /// Snarl status codes.
    /// Contains error codes for function calls, as well as callback values sent by Snarl
    /// to the window specified in Register() when fx. a Snarl notification times out
    /// or the user clicks on it.
    /// </summary>
    public enum SnarlStatus : short
    {
        Success = 0,

        // Snarl-Stopped/Started/UserAway/UserBack is defined in the GlobalEvent struct in VB6 code,
        // but are sent directly to a registered window, so in C# they are defined here as well.
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
        ErrorUnknownCommand,           // specified command not recognized
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
        // ErrorAlreadyRegistered,     // (Deprecated it seems - reregistering is not an error any more!)
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

        // Events - https://sites.google.com/site/snarlapp/developers/api-reference#TOC-Events
        NotifyGone = 301,              // reserved for future use
        // NotifyClick = 302,          // Deprecated: notification was right-clicked
        NotifyExpired = 303,           // Notification timed out
        NotifyInvoked = 304,           // Notification was clicked by the user
        // NotifyMenu,                 // Deprecated: item was selected from the notification's menu
        // NotifyExClick               // Deprecated: user clicked the middle mouse button on the notification 
        NotifyClosed = 307,            // User clicked the notification's Close gadget  
        NotifyAction = 308,            // User selected an action from the notification's Actions menu 

        // other events

        // reserved app event 320
        NotifyAppDoAbout = 321,
        NotifyAppDoPrefs,
        NotifyAppActivated,
        NotifyAppQuit,
    }

    /// <summary>
    /// Global event identifiers - sent as Windows broadcast messages.
    /// These values appear in wParam of the message.
    /// </summary>
    public enum SnarlGlobalEvent
    {
        Undefined = 0,
        SnarlLaunched = 1,   // Snarl has just started running
        SnarlQuit = 2,       // Snarl is about to stop running
        SnarlStopped = 3,    // Sent when stopped by user - Also sent to registered window
        SnarlStarted = 4,    // Sent when started by user - Also sent to registered window
        UserAway,            // Away mode was enabled
        UserBack,            // Away mode was disabled
    }

    /// <summary>
    /// The priority of messages.
    /// See <cref>http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=Generic_API#notify</cref>
    /// </summary>
    public enum SnarlMessagePriority
    {
        Low = -1,
        Normal = 0,
        High = 1
    }

    /// <summary>
    /// Application flags - features this app supports.
    /// </summary>
    [Flags]
    public enum SnarlAppFlags
    {
        None = 0,
        AppHasPrefs = 1,
        AppHasAbout = 2,
        AppIsWindowless = 0x8000
    }

    /// <summary>
    /// The SnarlParameterList is used to build the key/value pairs used when 
    /// making Snarl requests. The class can also be used by users in cases where the 
    /// normal functions does not provide the required functionality or lacks behind the
    /// native Snarl API. (<see cref="DoRequest"/>.)
    /// </summary>
    public sealed class SnarlParameterList
    {
        private List<KeyValuePair<String, String>> list = new List<KeyValuePair<String, String>>();

        public SnarlParameterList()
        {
        }

        public SnarlParameterList(int initialCapacity)
        {
            list.Capacity = initialCapacity;
        }

        public IList<KeyValuePair<String, String>> ParamList
        {
            get { return list; }
        }

        public SnarlParameterList Add(String key, String value)
        {
            DebugCheckCapacity();

            if (value != null)
                list.Add(new KeyValuePair<String, String>(key, value));

            return this;
        }

        public SnarlParameterList Add(String key, bool value)
        {
            DebugCheckCapacity();

            list.Add(new KeyValuePair<String, String>(key, value ? "1" : "0"));

            return this;
        }

        public SnarlParameterList Add(String key, Int32? value)
        {
            DebugCheckCapacity();

            if (value != null && value.HasValue)
                list.Add(new KeyValuePair<String, String>(key, value.Value.ToString()));

            return this;
        }

        public SnarlParameterList Add(String key, IntPtr value)
        {
            DebugCheckCapacity();

            if (value != null)
                list.Add(new KeyValuePair<String, String>(key, value.ToString()));

            return this;
        }

        public SnarlParameterList Add(SnarlAction[] actions)
        {
            if (actions != null)
            {
                foreach (var a in actions)
                {
                    String action = a.Label + "," + a.Callback;
                    list.Add(new KeyValuePair<String, String>("action", action));
                }
            }

            return this;
        }

        /// <summary>
        /// Ensures that initial capacity is large enough. (Only in debug builds.)
        /// </summary>
        private void DebugCheckCapacity()
        {
            Debug.Assert(list.Capacity >= list.Count + 1);
        }
    }

    public class SnarlGlobalEventArgs : EventArgs
    {
        public SnarlGlobalEvent GlobalEvent { get; set; }

        public SnarlGlobalEventArgs(SnarlGlobalEvent globalEvent)
        {
            GlobalEvent = globalEvent;
        }
    }

    public class SnarlCallbackEventArgs : EventArgs
    {
        public Int32 MessageToken { get; set; }
        public SnarlStatus SnarlEvent { get; set; }

        /// <summary>
        /// The data member if an action callback. Menu index if popup menu.
        /// </summary>
        public UInt16 Parameter { get; set; }

        public SnarlCallbackEventArgs(SnarlStatus snarlEvent, UInt16 parameter, int msgToken)
        {
            SnarlEvent = snarlEvent;
            MessageToken = msgToken;
            Parameter = parameter;
        }
    }

    public struct SnarlAction
    {
        public String Label { get; set; }
        public String Callback { get; set; }
    }

    [Serializable]
    public class SnarlException : Exception
    {
        public SnarlException() { }
        public SnarlException(string message) : base(message) { }
        public SnarlException(string message, Exception inner) : base(message, inner) { }
        protected SnarlException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}

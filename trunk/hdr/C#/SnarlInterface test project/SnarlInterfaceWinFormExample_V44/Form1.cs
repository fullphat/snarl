using System;
using System.Security.Principal;
using System.Windows.Forms;
using Snarl.V44;

namespace SnarlInterfaceWinFormExample_V44
{
    /// <summary>
    /// SnarlInterface example application.<br />
    /// <br />
    /// This is a simple example of what I, as maintainer of SnarlInterface, consider best practise
    /// when it comes to using SnarlInterface. It is not meant as a display of all Snarl features.
    /// Feel free to use the code in you own application.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Snarl

        SnarlInterface snarl = new SnarlInterface();

        // Snarl message classes
        const String SnarlClassNormal = "Normal";
        const String SnarlClassCritical = "Critical";
        const String SnarlClassLow = "Low";

        const int NormalMsgCallbackValue = 1;

        String snarlPassword = CreateSnarlPassword(8);

        // Action callback values
        enum SnarlActions
        {
            DoSomething = 1,
            DoSomethingElse,
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
            snarl.Unregister();
        }

        private void InitializeSnarl()
        {
            var vers = SnarlInterface.GetVersion();

            // ReRegisterSnarl() is called when first starting, and when a launch of Snarl is detected after this program is started.
            ReRegisterSnarl();

            // After registering, setup event handlers.
            // Not needed to do more than once, unless you call UnregisterCallbackWindow()
            snarl.CallbackEvent += CallbackEventHandler;

            // Using lambda expression
            snarl.GlobalSnarlEvent += (snarlInstance, args) =>
            {
                Log("Received global event: " + (int)args.GlobalEvent);
                if (args.GlobalEvent == SnarlGlobalEvent.SnarlLaunched)
                    ReRegisterSnarl();
                else if (args.GlobalEvent == SnarlGlobalEvent.SnarlQuit)
                    SnarlStatusLabel.Text = "Not running";
                else
                    Log("Received global event: " + args.GlobalEvent);
            };

            // Update UI
            SnarlStatusLabel.Text = SnarlInterface.GetSnarlWindow() == IntPtr.Zero ? "Not running" : "Running";
        }

        private void ReRegisterSnarl()
        {
            int result = 0;
            String snarlIcon = SnarlInterface.GetIconsPath() + "presence.png";

            // Use the form window to receive messages from Snarl
            // result = snarlInterface.RegisterWithEvents("application/Noer_IT.Example1", "SnarlInterface example1", snarlIcon, snarlPassword, this.Handle, null, SnarlAppFlags.None);

            // Alternative: Use a new auto created window
            result = snarl.RegisterWithEvents("application/Noer_IT.V44WinFormExample", "SnarlInterface WinForm example", snarlIcon, snarlPassword);

            if (result < (int)SnarlStatus.Success)
                Log("Failed to register with Snarl. Error=" + ((SnarlStatus)Math.Abs(result)).ToString());

            snarl.AddClass(SnarlClassNormal, "Normal messages");
            snarl.AddClass(SnarlClassCritical, "Critical messages");
            snarl.AddClass(SnarlClassLow, "Low priority messages");

            SnarlStatusLabel.Text = SnarlInterface.GetSnarlWindow() == IntPtr.Zero ? "Not running" : "Running";
        }

        private void CallbackEventHandler(object sender, SnarlCallbackEventArgs e)
        {
            switch (e.SnarlEvent)
            {
                case SnarlStatus.NotifyAction:
                    HandleActionCallback(e.Parameter, e.MessageToken);
                    break;

                case SnarlStatus.NotifyInvoked:
                case SnarlStatus.CallbackInvoked:
                    if (e.Parameter == NormalMsgCallbackValue)
                        Log("NotifyInvoked event - Normal message callback");
                    else
                        Log("{0} event - MessageToken={1} | Parameter={2}", e.SnarlEvent, e.MessageToken, e.Parameter);
                    break;

                case SnarlStatus.NotifyExpired:
                case SnarlStatus.CallbackTimedOut:
                    Log("{0} event - message with token={1} timed out.", e.SnarlEvent, e.MessageToken);
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
            String uid = "NormalMessage"; // Static uid - this means that the message will be update every time it send.
            Int32 msgToken = snarl.Notify(uId: uid,
                                          classId: SnarlClassNormal,
                                          title: "Normal message",
                                          text: "Some text\nTime: " + DateTime.Now.ToLongTimeString(),
                                          priority: SnarlMessagePriority.Normal,
                                          callback: "@" + NormalMsgCallbackValue,
                                          actions: GetDefaultActions());
            Log("Created message with uid=[{0}] - token returned={1}", uid, msgToken);
        }

        private void SendCriticalButton_Click(object sender, EventArgs e)
        {
            String uid = Guid.NewGuid().ToString();
            Int32 msgToken = snarl.Notify(uId: uid,
                                          classId: SnarlClassNormal,
                                          title: "Critical message",
                                          text: "Some text",
                                          priority: SnarlMessagePriority.High,
                                          actions: GetDefaultActions());
            Log("Created message with uid=[{0}] - token returned={1}", uid, msgToken);
        }
        
        private void SendLowButton_Click(object sender, EventArgs e)
        {
            String uid = Guid.NewGuid().ToString();
            Int32 msgToken = snarl.Notify(uId: uid, classId: SnarlClassLow, priority: SnarlMessagePriority.Low,
                                          title: "Low priority message",
                                          text: "Only one low priority messages are displayed at a time - subsequent notifications replaces the current.",
                                          actions: new SnarlAction[]
                                          {
                                              new SnarlAction() { Label = "Do something", Callback = "@" + (int)SnarlActions.DoSomething },
                                              new SnarlAction() { Label = "Do something else", Callback = "@" + (int)SnarlActions.DoSomethingElse },
                                          });
            //actions: GetDefaultActions());
            Log("Created message with uid=[{0}] - token returned={1}", uid, msgToken);
        }

        private SnarlAction[] GetDefaultActions()
        {
            return new SnarlAction[]
                {
                    new SnarlAction() { Label = "Do something", Callback = "@" + (int)SnarlActions.DoSomething },
                    new SnarlAction() { Label = "Do something else", Callback = "@" + (int)SnarlActions.DoSomethingElse },
                };
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
            // Though the official Snarl guideline says to generate a random password, this can be quite annoying when testing,
            // as this will make register fail if the application is quit without proper unregister call.
            // (Since passwords won't match between the two application instances.)

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

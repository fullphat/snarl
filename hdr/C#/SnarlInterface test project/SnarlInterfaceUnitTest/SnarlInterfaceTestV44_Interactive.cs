using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Snarl.V44;

namespace SnarlConnectorUnitTest
{
    [TestClass]
    public class SnarlInterfaceTestV44_Interactive
    {
        private readonly String AppSignatur = "application/Noer_IT.SnarlInterface_V44_UnitTest";
        private readonly String Title = "SnarlInterface UnitTest Interactive";
        private readonly String Password = "12345";

        private SnarlCallbackEventArgs savedCallbackArgs = null; // Variable used to check callback 
        private SnarlGlobalEventArgs savedGlobalArgs = null; // Variable used to check callback 


        [TestMethod]
        public void V44_Interactive_GlobalEventTest1()
        {
            ManualResetEvent stepEvent = new ManualResetEvent(false);
            var snarl = new SnarlInterface();

            SnarlInterface.GlobalEventHandler stoppedEventHandler = (sender, args) =>
            {
                Assert.AreSame(snarl, sender);
                Assert.AreNotEqual(null, args);
                savedGlobalArgs = args;
                stepEvent.Set();
            };
            snarl.GlobalSnarlEvent += stoppedEventHandler;

            SetupSnarlTest(snarl, stepEvent, false);

            snarl.Notify(text: "Stop Snarl");
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlGlobalEvent.SnarlStopped, savedGlobalArgs.GlobalEvent);
            stepEvent.Reset();

            MessageBox.Show("Test complete - start Snarl again before proceeding.");
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlGlobalEvent.SnarlStarted, savedGlobalArgs.GlobalEvent);
            stepEvent.Reset();

            snarl.GlobalSnarlEvent -= stoppedEventHandler;
            DoCleanup(stepEvent);
        }

        [TestMethod]
        public void V44_Interactive_CallbackEventTest1()
        {
            ManualResetEvent stepEvent = new ManualResetEvent(false);
            var snarl = new SnarlInterface();
            SetupSnarlTest(snarl, stepEvent);

            snarl.Notify(text: "Close me");
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.CallbackClosed, savedCallbackArgs.SnarlEvent);
            stepEvent.Reset();

            snarl.Notify(text: "Click Test button", callback: "@", callbackLabel: "Test");
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.CallbackInvoked, savedCallbackArgs.SnarlEvent);
            stepEvent.Reset();

            snarl.Notify(text: "Click show button", callback: "@1");
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.CallbackInvoked, savedCallbackArgs.SnarlEvent);
            Assert.AreEqual(1, savedCallbackArgs.Parameter);
            stepEvent.Reset();

            System.Windows.Forms.Application.Exit();
            stepEvent.WaitOne();
        }

        [TestMethod]
        public void V44_Interactive_TimeoutTest()
        {
            ManualResetEvent stepEvent = new ManualResetEvent(false);
            var snarl = new SnarlInterface();
            SetupSnarlTest(snarl, stepEvent);

            snarl.Notify(text: "Wait on timeout (2s)", timeout: 2);
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.CallbackTimedOut, savedCallbackArgs.SnarlEvent);
            stepEvent.Reset();

            DoCleanup(stepEvent);
        }

        [TestMethod]
        public void V44_Interactive_ActionEventTest1()
        {
            ManualResetEvent stepEvent = new ManualResetEvent(false);
            var snarl = new SnarlInterface();
            SetupSnarlTest(snarl, stepEvent);

            var actions = new SnarlAction[]
            {
                new SnarlAction() { Label = "Action 1", Callback = "@1" },
                new SnarlAction() { Label = "Action 2", Callback = "@2" },
                new SnarlAction() { Label = "Action 3", Callback = "d:\\" },
            };

            snarl.Notify(title: "Action test", text: "Click action 1", actions: actions);
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.NotifyAction, savedCallbackArgs.SnarlEvent);
            stepEvent.Reset();

            snarl.Notify(title: "Action test", text: "Click action 2", actions: actions);
            stepEvent.WaitOne();
            Assert.AreEqual(SnarlStatus.NotifyAction, savedCallbackArgs.SnarlEvent);
            stepEvent.Reset();

            DoCleanup(stepEvent);
        }

        private void SetupSnarlTest(SnarlInterface snarl, ManualResetEvent stepEvent, bool addCallbackEvent = true)
        {
            ThreadPool.QueueUserWorkItem((_) =>
            {
                var icon = Path.Combine(SnarlInterface.GetIconsPath(), "good.png");
                int result = 0;

                result = snarl.RegisterWithEvents(AppSignatur, Title, Password, icon);
                Assert.IsTrue(result > 0);

                stepEvent.Set();

                EventHandler exitExp = null;
                exitExp = (_x, _y) =>
                {
                    snarl.Unregister();
                    stepEvent.Set();

                    System.Windows.Forms.Application.ApplicationExit -= exitExp;
                };
                System.Windows.Forms.Application.ApplicationExit += exitExp;

                System.Windows.Forms.Application.Run(); // Start pumping messages
            });
            // Wait on new thread to register with Snarl
            if (!stepEvent.WaitOne(TimeSpan.FromSeconds(30)))
                Assert.Inconclusive("Thread was never started.");
            stepEvent.Reset();

            if (addCallbackEvent)
            {
                snarl.CallbackEvent += (sender, args) =>
                {
                    Assert.AreSame(snarl, sender);
                    Assert.AreNotEqual(null, args);
                    Assert.AreNotEqual(0, args.MessageToken);
                    Assert.AreNotEqual(0, args.SnarlEvent);
                    savedCallbackArgs = args;
                    stepEvent.Set();
                };
            }
        }
        
        private void DoCleanup(ManualResetEvent stepEvent)
        {
            System.Windows.Forms.Application.Exit();
            stepEvent.WaitOne(); // Wait on unregister before completing test
        }
    }
}

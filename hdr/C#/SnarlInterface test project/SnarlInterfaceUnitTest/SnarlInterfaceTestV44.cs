using System;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Snarl.V44;

namespace SnarlConnectorUnitTest
{
    [TestClass]
    public class SnarlInterfaceTestV44
    {
        private readonly String AppSignatur = "application/Noer_IT.SnarlInterface_V44_UnitTest";
        private readonly String Title = "SnarlInterface UnitTest";
        private readonly String Password = "12345";

        #region Static methods

        [TestMethod]
        public void V44_Static_IsSnarlRunningTest()
        {
            Assert.IsTrue(SnarlInterface.IsSnarlRunning());
        }

        [TestMethod]
        public void V44_Static_GetErrorTextTest1()
        {
            String msg = SnarlInterface.GetErrorText(-(int)SnarlStatus.ErrorFailed);
            Assert.AreEqual("ErrorFailed", msg);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void V44_Static_GetErrorTextTest2()
        {
            String msg = SnarlInterface.GetErrorText(Int32.MaxValue);
        }

        [TestMethod]
        public void V44_Static_GetVersionTest()
        {
            int version = SnarlInterface.GetVersion();
            Assert.IsTrue(version >= 43); 
        }

        [TestMethod]
        public void V44_Static_GetBroadcastMessageTest()
        {
            Assert.AreNotEqual(0, SnarlInterface.GetBroadcastMessage());
        }

        [TestMethod]
        public void V44_Static_GetAppMsgTest()
        {
            Assert.AreNotEqual(0, SnarlInterface.GetAppMsg());
        }

        [TestMethod]
        public void V44_Static_GetSnarlWindowTest()
        {
            Assert.AreNotEqual(IntPtr.Zero, SnarlInterface.GetSnarlWindow());
        }

        [TestMethod]
        public void V44_Static_GetSnarlPathTest()
        {
            var path = SnarlInterface.GetSnarlPath();
            Assert.IsTrue(Directory.Exists(path));
        }

        [TestMethod]
        public void V44_Static_GetIconsPathTest()
        {
            var path = SnarlInterface.GetIconsPath();
            Assert.IsTrue(Directory.Exists(path));
        }

        #endregion

        #region Member tests

        [TestMethod]
        public void V44_RegisterUnregisterTest1()
        {
            // Arrange
            int result = 0;
            var snarl = new SnarlInterface();

            // Act
            result = snarl.Register(AppSignatur, Title, Password);
            
            // Assert
            Assert.IsTrue(result > 0);

            // Unregister
            result = snarl.Unregister();
            AssertSuccess(result);
        }

        /// <summary>
        /// Tests register/unregister error codes.
        /// </summary>
        [TestMethod]
        public void V44_RegisterUnregisterTest2()
        {
            int result = 0;
            var snarl = new SnarlInterface();

            result = snarl.Register(AppSignatur, Title, Password);
            Assert.IsTrue(result > 0);

            result = snarl.Register(AppSignatur, Title, Password);
            // Assert.AreEqual(-(int)SnarlStatus.ErrorAlreadyRegistered, result);
            Assert.IsTrue(result > 0); // ErrorAlreadyRegistered seems to be depricated

            // Unregister
            result = snarl.Unregister();
            AssertSuccess(result);

            result = snarl.Unregister();
            Assert.AreEqual(-(int)SnarlStatus.ErrorNotRegistered, result);
        }

        /// <summary>
        /// Tests register with icon.
        /// </summary>
        [TestMethod]
        public void V44_RegisterUnregisterTest3()
        {
            int result = 0;
            var snarl = new SnarlInterface();
            var icon = Path.Combine(SnarlInterface.GetIconsPath(), "good.png");

            result = snarl.Register(AppSignatur, Title, Password, icon);
            Assert.IsTrue(result > 0);

            snarl.Notify(text: "Check that registration has an icon");

            // Unregister
            AssertSuccess(snarl.Unregister());
        }

        [TestMethod]
        public void V44_AddClassTest1()
        {
            SnarlInterface snarl = null;
            try
            {
                int result;
                var icon = Path.Combine(SnarlInterface.GetIconsPath(), "good.png");
                snarl = RegisterSnarl();

                result = snarl.AddClass("Class1", "Class 1", "Default title", "Default text", icon);
                AssertSuccess(result);

                result = snarl.Notify(classId: "Class1");
                Assert.IsTrue(result > 0);

                result = snarl.Notify(classId: "Class1", title: "This should have text = \"Default text\".");
                Assert.IsTrue(result > 0);

                result = snarl.Notify(classId: "Class1", text: "Test - this should have title = \"Default title\".");
                Assert.IsTrue(result > 0);
            }
            finally
            {
                if (snarl != null)
                    UnregisterSnarl(snarl);
            }
        }

        /// <summary>
        /// Tests register with icon.
        /// </summary>
        [TestMethod]
        public void V44_NotifyTest1()
        {
            SnarlInterface snarl = null;
            try
            {
                int result;
                var icon = Path.Combine(SnarlInterface.GetIconsPath(), "good.png");
                snarl = RegisterSnarl();

                result = snarl.Notify();
                Assert.IsTrue(result > 0);

                // Test update
                result = snarl.Notify(uId: "123", title: "V44_NotifyTest1", text: "Test");
                Assert.IsTrue(result > 0);

                for (int i = 1; i <= 5; ++i)
                {
                    Thread.Sleep(500);
                    int result2 = snarl.Notify(uId: "123", title: "V44_NotifyTest1", text: "Updated text - " + i);
                    Assert.AreEqual(result, result2); // Check that same token is returned
                }
            }
            finally
            {
                UnregisterSnarl(snarl);
            }
        }

        private SnarlInterface RegisterSnarl()
        {
            int result = 0;
            var snarl = new SnarlInterface();
            var icon = Path.Combine(SnarlInterface.GetIconsPath(), "good.png");

            result = snarl.Register(AppSignatur, Title, Password, icon);
            Assert.IsTrue(result > 0);
            
            return snarl;
        }

        private void UnregisterSnarl(SnarlInterface snarl)
        {
            snarl.Unregister();
        }

        #endregion

        /// <summary>
        /// Helper function to avoid to cast the SnarlStutus enum all the time.
        /// </summary>
        private void AssertSuccess(int result)
        {
            Assert.AreEqual((short)SnarlStatus.Success, result);
        }
    }
}

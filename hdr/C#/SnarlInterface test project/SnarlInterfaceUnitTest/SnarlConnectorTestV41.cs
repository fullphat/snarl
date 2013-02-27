using Snarl.V41;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;

namespace SnarlConnectorUnitTest
{
	
	
	/// <summary>
	///This is a test class for SnarlConnectorTest and is intended to contain all SnarlConnectorTest Unit Tests.
	///
	/// General API note from Chris:
	/// Every function should:
	/// 1. Return a LONG value
	/// 2. If only a success/fail result is required, function should return -1 (0xFFFFFFFF) to indicate success or 0 to indicate failure
	/// 3. All functions must set LastError to zero ("no error") on success
	/// 4. All functions must set LastError to the most suitable error value (a set are defined) if it fails
	/// 5. Application registrations are managed via a ULONG token - this is returned after a successful call to sn41RegisterApp()
	/// 6. Class registrations are managed via a BSTR identifier - this can be anything but is usually something short and simple (e.g. "class1")
	/// 7. Notifications are managed via a ULONG token - this is returned after a successful call to sn41Notify() or sn41EZNotify()
	///</summary>
	[TestClass()]
	public class SnarlConnectorTestV41
	{
		private const String ClassId1 = "Class1";
		private const String ClassId2 = "Class2";
		private const Int32 DefaultMsgTimeout = 10;

		private SnarlConnector snarl = new SnarlConnector();
		private Int32 snarlToken = 0;

		
		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		private TestContext testContextInstance;
		public TestContext TestContext
		{
			get
			{
				return testContextInstance;
			}
			set
			{
				testContextInstance = value;
			}
		}

		#region Additional test attributes

		// Use ClassInitialize to run code before running the first test in the class
		// [ClassInitialize()]
		// public static void MyClassInitialize(TestContext testContext)
		// {
		// }
		// 
		// // Use ClassCleanup to run code after all tests in a class have run
		// [ClassCleanup()]
		// public static void MyClassCleanup()
		// {
		// }
		
		// Use TestInitialize to run code before running each test
		[TestInitialize()]
		public void MyTestInitialize()
		{
			Int32 value = 0;
			Int32 actual = 0;

			snarlToken = snarl.RegisterApp("CSharpUnitTest", "C# unit test", null, IntPtr.Zero, 0, 0);
			Assert.AreNotEqual(0, snarlToken);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}
		
		// Use TestCleanup to run code after each test has run
		[TestCleanup()]
		public void MyTestCleanup()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = -1;
			actual = snarl.UnregisterApp();
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}
		
		#endregion


		/// <summary>
		///A test for RegisterApp
		///</summary>
		[TestMethod()]
        public void V41_RegisterAppTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			// Subsequent calls should return same token as first call
			value = snarlToken;
			actual = snarl.RegisterApp("CSharpUnitTest", "C# unit test", null, IntPtr.Zero, 0, 0);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			value = -1;
			actual = snarl.UnregisterApp();
			Assert.AreEqual(value, actual);

			// Test overloaded version
			value = 0;
			actual = snarl.RegisterApp("CSharpUnitTest", "C# unit test", null);
			Assert.AreNotEqual(value, actual);

			value = -1;
			actual = snarl.UnregisterApp();
			Assert.AreEqual(value, actual);

			// Test invalid parameters
			value = 0;
			actual = snarl.RegisterApp("", "C# unit test", null, IntPtr.Zero, 0, 0);
			Assert.AreEqual(value, actual);

			// LastError
			value = (Int32)SnarlConnector.SnarlStatus.ErrorArgMissing;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Leave registered with Snarl
			value = 0;
			actual = snarl.RegisterApp("CSharpUnitTest", "C# unit test", null, IntPtr.Zero, 0, 0);
			Assert.AreNotEqual(value, actual);
		}

		/// <summary>
		///A test for UnregisterApp
		///</summary>
		[TestMethod()]
        public void V41_UnregisterAppTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			// Post condition : Leave Snarl registered
			value = -1;
			actual = snarl.UnregisterApp();
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			value = 0;
			actual = snarl.UnregisterApp();
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.ErrorNotRegistered;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			value = 0;
			actual = snarl.RegisterApp("CSharpUnitTest", "C# unit test", null, IntPtr.Zero, 0, 0);
			Assert.AreNotEqual(value, actual);
		}

		///<summary>
		///A test for AddClass.
		///
		///Note: RemoveClassTest and RemoveAllClassesTest calls this test!
		///</summary>
		[TestMethod()]
        public void V41_AddClassTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = -1;
			actual = snarl.AddClass(ClassId1, "Class 1", true);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			value = -1;
			actual = snarl.AddClass(ClassId2, "Class 2", true);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Test error
			value = 0;
			actual = snarl.AddClass(ClassId1, "Class 1", true);
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.ErrorClassAlreadyExists;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for Hide
		///</summary>
		[TestMethod()]
        public void V41_HideTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = 0;
			actual = snarl.EZNotify("id", "title", "text", DefaultMsgTimeout, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);

			Thread.Sleep(1000);

			value = -1;
			actual = snarl.Hide(actual);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
			
			// Error test
			value = 0;
			actual = snarl.Hide(actual);
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.ErrorNotificationNotFound;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		/// A test for IsVisible
		///
		/// sn41IsVisible() will return -1 if it's visible and 0 if it's not.  If zero is return you can call sn41GetLastError() to see if something  went wrong.
		/// If sn41GetLastError() returns 0 then  you know that sn41IsVisible() succeeded and that the notification wasn't visible.
		/// </summary>
		[TestMethod()]
        public void V41_IsVisibleTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = 0;
			actual = snarl.EZNotify("IsVisibleId", "title", "text", DefaultMsgTimeout, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);
			
			Thread.Sleep(1000);

			value = -1;
			actual = snarl.IsVisible(snarl.GetLastMsgToken());
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Error test
			value = 0;
			actual = snarl.EZNotify("IsVisibleId", "title", "text", 1, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);

			Thread.Sleep(3000); // Message should be gone

			value = 0;
			actual = snarl.IsVisible(snarl.GetLastMsgToken());
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for EZNotify
		///</summary>
		[TestMethod()]
        public void V41_EZNotifyTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "EZNotifyTest", "Full version", DefaultMsgTimeout, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Test overloaded functions
			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "EZNotifyTest", "5 parameters version", DefaultMsgTimeout, null);
			Assert.AreNotEqual(value, actual);

			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "EZNotifyTest", "4 parameters version", DefaultMsgTimeout);
			Assert.AreNotEqual(value, actual);

			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "EZNotifyTest", "3 parameters version");
			Assert.AreNotEqual(value, actual);

		}

		/// <summary>
		///A test for Notify
		///</summary>
		[TestMethod()]
        public void V41_NotifyTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			String packetData = 
				"title::NotifyTest" +
				"#?text::Text" +
				"#?timeout::" + DefaultMsgTimeout;

			value = 0;
			actual = snarl.Notify("NotifyId", packetData);
			Assert.AreNotEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for RemoveClass
		///</summary>
		[TestMethod()]
        public void V41_RemoveClassTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

            V41_AddClassTest();

			value = -1;
			actual = snarl.RemoveClass(ClassId1, true);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Error test
			value = 0;
			actual = snarl.RemoveClass(ClassId1, true);
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.ErrorClassNotFound;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			value = -1;
			actual = snarl.RemoveClass(ClassId2, false);
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for RemoveAllClasses
		///</summary>
		[TestMethod()]
        public void V41_RemoveAllClassesTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

            V41_AddClassTest();

			value = -1;
			actual = snarl.RemoveAllClasses(true);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Error test
			value = -1;
			actual = snarl.RemoveAllClasses(true);
			Assert.AreEqual(value, actual);

			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for EZUpdate
		///</summary>
		[TestMethod()]
        public void V41_EZUpdateTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			// Test without message - Returns zero on failure
			value = 0;
			actual = snarl.EZUpdate(snarl.GetLastMsgToken(), "Updated title", "Updated text", DefaultMsgTimeout, "");
			Assert.AreEqual(value, actual);

			// LastError - Should return ErrorFailed
			value = (Int32)SnarlConnector.SnarlStatus.ErrorNotificationNotFound;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Test success case
			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "title", "text", DefaultMsgTimeout, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);

			Thread.Sleep(1000);

			value = -1;
			actual = snarl.EZUpdate(snarl.GetLastMsgToken(), "Updated title", "Updated text", DefaultMsgTimeout, "");
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for Update
		///</summary>
		[TestMethod()]
        public void V41_UpdateTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			String packetData = "title::Updated message#?text::Updated message text#?icon::" + snarl.GetIconsPath() + "snarl-update.png";

			// Test without message - Returns zero on failure
			value = 0;
			actual = snarl.Update(snarl.GetLastMsgToken(), packetData);
			Assert.AreEqual(value, actual);

			// LastError - Should return ErrorFailed
			value = (Int32)SnarlConnector.SnarlStatus.ErrorNotificationNotFound;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);

			// Test success case
			value = 0;
			actual = snarl.EZNotify("EZNotifyId", "title", "text", DefaultMsgTimeout, null, 0, "acknowledge", "value");
			Assert.AreNotEqual(value, actual);

			Thread.Sleep(1000);

			value = -1;
			actual = snarl.Update(snarl.GetLastMsgToken(), packetData);
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}

		/// <summary>
		///A test for UpdateApp
		///</summary>
		[TestMethod()]
        public void V41_UpdateAppTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			value = -1;
			actual = snarl.UpdateApp("C# updated", "");
			Assert.AreEqual(value, actual);

			// LastError - Should return success
			value = (Int32)SnarlConnector.SnarlStatus.Success;
			actual = (Int32)snarl.GetLastError();
			Assert.AreEqual(value, actual);
		}
	}
}

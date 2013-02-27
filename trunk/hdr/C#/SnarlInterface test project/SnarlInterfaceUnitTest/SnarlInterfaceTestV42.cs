using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Snarl.V42;

namespace SnarlConnectorUnitTest
{
	
	
	/// <summary>
	///This is a test class for SnarlInterfaceTest and is intended
	///to contain all SnarlInterfaceTest Unit Tests
	///</summary>
	[TestClass()]
	public class SnarlInterfaceTestV42
	{
		private const String appId = "CSharp/Interfacetest";
		private const String appTitle = "CSharp Unit test";
		private const String ClassId1 = "Class1";
		private const String ClassId2 = "Class2";
		private const Int32 DefaultMsgTimeout = 10;

		private SnarlInterface snarl = new SnarlInterface();
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
		// 
		//You can use the following additional attributes as you write your tests:
		//
		//Use ClassInitialize to run code before running the first test in the class
		//[ClassInitialize()]
		//public static void MyClassInitialize(TestContext testContext)
		//{
		//}
		//
		//Use ClassCleanup to run code after all tests in a class have run
		//[ClassCleanup()]
		//public static void MyClassCleanup()
		//{
		//}
		
		// Use TestInitialize to run code before running each test
		[TestInitialize()]
		public void MyTestInitialize()
		{
			snarlToken = snarl.Register(appId, appTitle, null);
			Assert.IsTrue(snarlToken > 0);
		}
		
		//Use TestCleanup to run code after each test has run
		[TestCleanup()]
		public void MyTestCleanup()
		{
			Assert.AreEqual(snarl.Unregister(), 0);
		}
		
		#endregion

		/// <summary>
		///A test for Notify
		///</summary>
		[TestMethod()]
        public void V42_NotifyTest()
		{
			String result = "", str = "", expected = "";

			str = "Some random string, which should not be escaped";
			expected = str;
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);

			str = "Some random string. & this should be escape once each =";
			expected = "Some random string. && this should be escape once each ==";
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);

			str = "Some random string. && this should be escape once each ==";
			expected = "Some random string. &&&& this should be escape once each ====";
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);

			str = "&";
			expected = "&&";
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);

			str = "&&&&";
			expected = "&&&&&&&&";
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);

			str = "&=";
			expected = "&&==";
			result = SnarlInterface.Escape(str);
			Assert.AreEqual(expected, result);
		}

		/// <summary>
		///A test for RegisterApp
		///</summary>
		[TestMethod()]
		public void V42_RegisterAppTest()
		{
			Int32 expected = 0;
			Int32 actual = 0;

			// Subsequent calls should return same token as first call
			expected = snarlToken;
			actual = snarl.Register(appId, appTitle, null);
			Assert.AreEqual(expected, actual);

			expected = 0;
			actual = snarl.Unregister();
			Assert.AreEqual(expected, actual);

			// Test with password - should not be able to unregister without password
			expected = 0;
			actual = snarl.Register(appId, appTitle, null, "MyPassword");
			Assert.AreNotEqual(expected, actual);

			expected = -(Int32)SnarlInterface.SnarlStatus.ErrorAuthFailure;
			actual = SnarlInterface.DoRequest(SnarlInterface.Requests.Unregister + "?app-sig=" + appId);
			Assert.AreEqual(expected, actual);

			expected = 0;
			actual = snarl.Unregister();
			Assert.AreEqual(expected, actual);

			// Test invalid parameters
			expected = 0; // not used
			actual = snarl.Register("", "C# unit test", null);  // Should be an error
			Assert.IsTrue(actual < 0);

			// Leave registered with Snarl
			MyTestInitialize();
		}

		/// <summary>
		///A test for RemoveClass
		///</summary>
		[TestMethod()]
        public void V42_RemoveClassTest()
		{
			/*Int32 value = 0;
			Int32 actual = 0;

			AddClassTest();

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
			Assert.AreEqual(value, actual);*/
		}

		/// <summary>
		///A test for UnregisterApp
		///</summary>
		[TestMethod()]
        public void V42_UnregisterAppTest()
		{
			Int32 value = 0;
			Int32 actual = 0;

			// Post condition : Leave Snarl registered
			value = 0;
			actual = snarl.Unregister();
			Assert.AreEqual(value, actual);

			value = -(Int32)SnarlInterface.SnarlStatus.ErrorNotRegistered;
			actual = snarl.Unregister();
			Assert.AreEqual(value, actual);

			MyTestInitialize();
		}
	}
}

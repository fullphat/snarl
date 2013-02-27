using System;
using System.Runtime.InteropServices;

namespace Snarl.V39
{
	[StructLayout( LayoutKind.Sequential, Pack = 4 )]
	struct SNARLSTRUCTEX
	{
		public Int16 Cmd;         // what to do...
		public Int32 Id;          // snarl message id (returned by snShowMessage())
		public Int32 Timeout;     // timeout in seconds (0=sticky)
		public Int32 LngData2;    // reserved

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Title;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Text;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Icon;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Class;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Extra;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Extra2;

		public Int32 Reserved1;
		public Int32 Reserved2;
	}
}

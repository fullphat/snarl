using System;
using System.Runtime.InteropServices;

namespace Snarl.V39
{
	[StructLayout(LayoutKind.Sequential, Pack = 4)]
	struct SNARLSTRUCT
	{
		public Int16 Cmd;           // what to do...
		public Int32 Id;            // snarl message id (returned by snShowMessage())
		public Int32 Timeout;       // timeout in seconds (0=sticky)
		public Int32 LngData2;      // reserved

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Title;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Text;

		[MarshalAs(UnmanagedType.ByValArray, SizeConst = SnarlConnector.SNARL_STRING_LENGTH)]
		public byte[] Icon;
	}
}

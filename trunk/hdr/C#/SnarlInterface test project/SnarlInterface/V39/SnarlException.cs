using System;
using System.Collections.Generic;
using System.Text;

namespace Snarl.V39
{
	public class SnarlException: Exception
	{
		public SnarlException( M_RESULT code )
			: base( "Snarl encountered an error while processing the request" )
		{
			_code = code;
		}

		private M_RESULT _code;
		public M_RESULT Code
		{
			get
			{
				return _code;
			}
			set
			{
				_code = value;
			}
		}
	}
}

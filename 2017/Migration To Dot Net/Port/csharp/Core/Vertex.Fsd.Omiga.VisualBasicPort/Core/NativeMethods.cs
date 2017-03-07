using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// Wraps P/Invoke methods.
	/// </summary>
	public static class NativeMethods
	{
		/// <summary>
		/// Gets the context associated with current COM+ object. Wraps the comsvcs.dll GetObjectContext method.
		/// </summary>
		/// <returns>The COM+ object context, or null if there is no current context.</returns>
		public static object GetObjectContext()
		{
			object objectContext = null;
			GetObjectContext(out objectContext);
			return objectContext;
		}

		[DllImport("comsvcs.dll")]
		private static extern int GetObjectContext([MarshalAs(UnmanagedType.IUnknown)] out object ppInstanceContext);
	}
}

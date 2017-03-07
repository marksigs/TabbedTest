/*
--------------------------------------------------------------------------------------------
Workfile:			ContextUtility.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Safe alternative to ContextUtil.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		09/05/2007	First version.
--------------------------------------------------------------------------------------------
*/
using System;
using System.EnterpriseServices;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// The ContextUtility class is an alternative to the 
	/// <see cref="System.EnterpriseServices.ContextUtil"/> class.
	/// </summary>
	public static class ContextUtility
	{
		/// <summary>
		/// Gets the current COM+ context.
		/// </summary>
		/// <returns>The current COM+ context, or null if there is no COM+ context.</returns>
		public static object GetObjectContext()
		{
			return NativeMethods.GetObjectContext();
		}

		/// <summary>
		/// Returns <b>true</b> if there is a COM+ context.
		/// </summary>
		/// <remarks>
		/// As an alternative it may be possible to use the <see cref="ContextUtil.IsInTransaction"/> 
		/// property, which in testing seems to return <b>false</b> if there is no COM+ context. 
		/// However, the MSDN documentation states that <see cref="ContextUtil.IsInTransaction"/> 
		/// will throw a COMException if there is not a COM+ context.
		/// </remarks>
		public static bool IsObjectContext
		{
			get { return GetObjectContext() != null; }
		}

		/// <summary>
		/// Calls <see cref="ContextUtil.SetComplete"/> if there is a COM+ context, 
		/// otherwise does nothing.
		/// </summary>
		/// <remarks>
		/// <see cref="ContextUtil.SetComplete"/> will throw a COMException if there is no COM+ 
		/// context. If a component may or may not be running in COM+ it is safer to use 
		/// ContextUtility.SetComplete.
		/// </remarks>
		public static void SetComplete()
		{
			if (IsObjectContext) ContextUtil.SetComplete();
		}

		/// <summary>
		/// Calls <see cref="ContextUtil.SetAbort"/> if there is a COM+ context, 
		/// otherwise does nothing.
		/// </summary>
		/// <remarks>
		/// <see cref="ContextUtil.SetAbort"/> will throw a COMException if there is no COM+ 
		/// context. If a component may or may not be running in COM+ it is safer to use 
		/// ContextUtility.SetAbort.
		/// </remarks>
		public static void SetAbort()
		{
			if (IsObjectContext) ContextUtil.SetAbort();
		}
	}
}

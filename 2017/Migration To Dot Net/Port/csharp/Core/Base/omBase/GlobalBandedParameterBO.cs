/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalBandedParameterDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Global Banded Parameter Business Object which 'supports transactions' only.
					Methods which do not require transaction support reside in this
					class. Any methods that require transactions will be delegated to
					GlobalBandedParameterTxBO
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		30/06/99	Created
DRC		3/10/01		SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		27/07/2007	First .Net version. Ported from GlobalParameterBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.GlobalBandedParameterBO")]
	[Guid("0923114B-1510-4889-8984-670521A8B47F")]
	[Transaction(TransactionOption.Supported)]
	public class GlobalBandedParameterBO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:  Gets the current values for the parameter and qualifier passed in
		// pass:		 name   Parameter for which current data is required
		// qualifier	   Qualifier to select correct band
		// return:	   GetCurrentParameter xml Response data stream containing results of
		// operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameter(string name, string qualifier) 
		{
			string response = "";

			try 
			{
				string data = "";
				using (GlobalBandedParameterDO globalBandedParameterDO = new GlobalBandedParameterDO())
				{
					data = globalBandedParameterDO.GetCurrentParameter(name, qualifier);
				}

				// if we get here, everything has completed OK
				response = "<RESPONSE TYPE='SUCCESS'>" + data + "</RESPONSE>";
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}

		protected override bool CanBePooled()
		{
			return true;
		}

		#region Not implemented
		/*
		// These methods were defined in the VB6 code as raising a "not implemented" error.
		// Comment out here so it is impossible to call them.
		public string Create(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Update(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Delete(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string DeleteAll(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string GetData(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string FindList(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}
		*/
		#endregion

	}

}

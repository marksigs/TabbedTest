/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterTxBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Transactioned Global Parameter Business Object which requires transactions
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		21/07/99	Created
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		27/07/2007	First .Net version. Ported from GlobalParameterTxBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.EnterpriseServices;
using System.Runtime.InteropServices;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.GlobalParameterTxBO")]
	[Guid("138A30B6-417E-422E-9B10-CE0C4B38D935")]
	[Transaction(TransactionOption.Required)]
	public class GlobalParameterTxBO : ServicedComponent
	{
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
		*/

		#endregion

	}

}

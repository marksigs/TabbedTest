/*
--------------------------------------------------------------------------------------------
Workfile:			FirstTitleNoTxBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		Class to escape an existing transaction
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	14/12/2005	MAR855 Created
PSC		10/01/2006	MAR961 Amend HandleInboundMessage and HandleOutboundMessage to take a 
                    process context for logging
GHun	10/01/2006	MAR972 Changed ProcessInboundMessage
GHun	27/03/2006  MAR1539 Changed ProcessOutboundMessage
--------------------------------------------------------------------------------------------
*/
using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;

namespace Vertex.Fsd.Omiga.omFirstTitle
{
	/// <summary>
	/// This class is used to escape existing transactions
	/// </summary>

	[ProgId("omFirstTitle.FirstTitleNoTxBO")]
	[ComVisible(true)]
	[Guid("F89FA28F-640F-40cd-8434-25ACFBC76C9D")]
	[Transaction(TransactionOption.NotSupported)]
	public class FirstTitleNoTxBO : ServicedComponent
	{

		public FirstTitleNoTxBO()
		{
		}
		
		// PSC 10/01/2006 MAR961
		public void ProcessInboundMessage(string xmlData, string processContext)
		{
			try 
			{
				// PSC 10/01/2006 MAR961
				//MAR972 GHun Call through NTxBO to create a new transaction
				FirstTitleNTxBO ntxBO = new FirstTitleNTxBO();
				ntxBO.ProcessInboundMessage(xmlData, processContext);
				//ContextUtil.SetComplete();
			}
			catch
			{
				//ContextUtil.SetAbort();
			}
		}

		// PSC 10/01/2006 MAR961
		public void ProcessOutboundMessage(string xmlData, string processContext)
		{
			try 
			{
				// PSC 10/01/2006 MAR961
				// MAR1392 Start
				//FirstTitleOutboundBO outBO = new FirstTitleOutboundBO(processContext);
				FirstTitleOutboundBO outBO = new FirstTitleOutboundBO();
				//MAR1539 GHun pass processContext as a parameter
				//outBO.ProcessContext = processContext; 
				// MAR1392 End
				outBO.ProcessOutboundMessage(xmlData, processContext);
				//ContextUtil.SetComplete();
			}
			catch
			{
				//ContextUtil.SetAbort();
			}
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			ESurvNoTxBO.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	15/12/2005	MAR821 Created
PSC		03/01/2006	MAR961 Amend HandleInboundMessage and HandleOutboundMessage to take a 
                    process context for logging
PSC		13/02/2006  Amend HandleInboundMessage to pass back return code and log any errors
PSC		09/05/2006	MAR1742 Add additional logging
--------------------------------------------------------------------------------------------
*/

//GHun	MAR821
using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using MQL = MESSAGEQUEUECOMPONENTVCLib;			// PSC 13/02/2006 MAR1207
using omLogging = Vertex.Fsd.Omiga.omLogging;	// PSC 13/02/2006 MAR1207

namespace Vertex.Fsd.Omiga.omESurv
{
	/// <summary>
	/// This class does not support transactions and is used to escape an existing transaction
	/// </summary>
	
	[ProgId("omESurv.ESurvNoTxBO")]
	[ComVisible(true)]
	[Guid("4E43D2FD-5DF7-41AE-AEDA-E9A8CFEE1C11")]
	[Transaction(TransactionOption.NotSupported)]
	public class ESurvNoTxBO : ServicedComponent
	{
		private omLogging.Logger _logger = null;		// PSC 13/02/2006 MAR1207
		
		public ESurvNoTxBO()
		{
		}

		// PSC 03/01/2006 MAR961
		public int HandleInboundMessage(string data, string processContext)
		{
			string functionName = System.Reflection.MethodBase.GetCurrentMethod().Name;
			try
			{
				_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, processContext);

				// PSC 09/05/2006 MAR1742 - Start
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Starting");
				}
				// PSC 09/05/2006 MAR1742 - End

				// PSC 03/01/2006 MAR961
				InboundBO inBo = new InboundBO(processContext);
				int Response = inBo.HandleInboundMessage(data);
				return Response;
				//ContextUtil.SetComplete();
			}
			catch (Exception exp)
			{
				if(_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": An unexpected error occurred handling inbound message", exp);
				}
				return (int)MQL.MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
			}
			// PSC 09/05/2006 MAR1742 - Start
			finally
			{
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Completed");
				}
			}
			// PSC 09/05/2006 MAR1742 - End
		}
		// PSC 13/02/2006 MAR1207 - End

		// PSC 03/01/2006 MAR961
		public void HandleOutboundMessage(string data, string processContext)
		{
			try
			{
				_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, processContext);
				
				// PSC 03/01/2006 MAR961
				OutboundBO outBo = new OutboundBO(processContext);
				outBo.HandleOutboundMessage(data);
				//ContextUtil.SetComplete();
			}
			catch
			{
				//ContextUtil.SetAbort();
			}
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			FirstTitleNTxBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		Class to create a new transaction
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	14/12/2005	MAR855 Created
PSC		10/01/2006	MAR961 Use omLogging wrapper
GHun	10/01/2006	MAR972 Added ProcessInboundMessage
--------------------------------------------------------------------------------------------
*/
using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using System.Xml;
using omMQ;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 10/01/2006 MAR961

namespace Vertex.Fsd.Omiga.omFirstTitle
{
	/// <summary>
	/// This class is used to create a new transaction
	/// </summary>

	[ProgId("omFirstTitle.FirstTitleNTxBO")]
	[ComVisible(true)]
	[Guid("3ADE9115-DC91-4cf8-9D26-F7865569BDE3")]
	[Transaction(TransactionOption.RequiresNew)]
	public class FirstTitleNTxBO : ServicedComponent
	{
		private  omLogging.Logger _logger = null;

		public FirstTitleNTxBO()
		{
		}
		
		//MAR821 Moved from OutboundBO to here so it will always run in a new transaction
		// PSC 10/01/2006 MAR961
		// RF 09/02/2006 MAR1392 Don't pass logger or you get the wrong component named in the log
		//public bool AddMessageToQueue(string strData, bool bAddToOutbound, bool bAddDelay, string interfaceName, string applicationNumber, string userId, string unitId, string userAuthorityLevel, string channelId, omLogging.Logger logger)
		public bool AddMessageToQueue(string strData, 
			bool bAddToOutbound, bool bAddDelay, string interfaceName, 
			string applicationNumber, string userId, string unitId, 
			string userAuthorityLevel, string channelId, string processContext)
		{
			const string cFunctionName = "AddMessageToQueue";
		
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, 
				processContext);

			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting " + cFunctionName + ": Data = " + strData);
				//_logger.Debug("Transaction ID: " + ContextUtil.TransactionId.ToString()); 
			}

			string Response = string.Empty;
			XmlDocument xdResponse = new XmlDocument();
			omMQBOClass omigaMessageQueue = new omMQBOClass();

			try
			{
				XmlDocument requestDocument = new XmlDocument();

				string targetQueueName;
				string progId;
				if(bAddToOutbound)
				{
					targetQueueName = new GlobalParameter("FirstTitleOutboundQueueName").String;
					progId = "omFirstTitle.FirstTitleOutboundBO";
				}
				else
				{
					targetQueueName = new GlobalParameter("FirstTitleInboundQueueName").String;
					progId = "omFirstTitle.FirstTitleInboundBO";
				}

				// Prepare the request for Queue
				XmlElement rootElement = requestDocument.CreateElement("REQUEST");
				rootElement.SetAttribute("OPERATION", "SendToQueue");
				requestDocument.AppendChild(rootElement);
			
				// Set up message element
				XmlElement messageElement = requestDocument.CreateElement("MESSAGEQUEUE");
				messageElement.SetAttribute("QUEUENAME", targetQueueName);
				messageElement.SetAttribute("PROGID", progId);

				if (bAddDelay)
				{
					//Add an executeafter time of now + 1hour
					string datestr = DateTime.Now.AddHours(1).ToString("dd/MM/yyyy HH:mm:ss");
					messageElement.SetAttribute("EXECUTEAFTERDATE", datestr);
				}
				rootElement.AppendChild(messageElement);

				// Set up data element
				XmlDocument xmlInData = new XmlDocument();
				xmlInData.LoadXml(strData);
				XmlNode xmlDataNode = xmlInData.SelectSingleNode("XML");

				if (xmlDataNode == null)
					xmlDataNode = xmlInData.DocumentElement;

				XmlNode dataElement = requestDocument.CreateElement("XML");
				messageElement.AppendChild(dataElement);

				XmlElement directGatewayRequest = requestDocument.CreateElement("REQUEST");
				directGatewayRequest.SetAttribute("USERAUTHORITYLEVEL", userAuthorityLevel);
				directGatewayRequest.SetAttribute("MACHINEID", System.Environment.MachineName);
				directGatewayRequest.SetAttribute("USERID", userId);
				directGatewayRequest.SetAttribute("UNITID", unitId);
				directGatewayRequest.SetAttribute("CHANNELID", channelId);
				directGatewayRequest.SetAttribute("OPERATION", "OnMessage");
				if (bAddToOutbound)
				{
					directGatewayRequest.SetAttribute("APPLICATIONNUMBER", applicationNumber);
					directGatewayRequest.SetAttribute("INTERFACENAME", interfaceName);
				}

				XmlNodeList NodesToImport= xmlDataNode.ChildNodes;
				for(int i = 0; i < NodesToImport.Count; i++)
				{
					XmlNode importedNode = requestDocument.ImportNode(NodesToImport.Item(i),true);
					directGatewayRequest.AppendChild(importedNode);
				}
				
				dataElement.AppendChild (directGatewayRequest);

				// RF 09/03/2006 MAR1392 
				//omMQBOClass omigaMessageQueue = new omMQBOClass(); 

				xdResponse.LoadXml(omigaMessageQueue.omRequest(requestDocument.OuterXml));

				string responseType = xdResponse.DocumentElement.GetAttribute("TYPE");
				if (responseType != "SUCCESS")
				{
					if(_logger.IsDebugEnabled)
					{
						_logger.Debug(cFunctionName + " failed: " + xdResponse.OuterXml);
					}
					throw new OmigaException(cFunctionName + ": error in MessageQueue call:" + 
						xdResponse.OuterXml);
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				ContextUtil.SetAbort();
				throw new OmigaException(ex);
			}
			// RF 09/03/2006 MAR1392 
			finally
			{
				Marshal.ReleaseComObject(omigaMessageQueue);
			}

			ContextUtil.SetComplete();
			return true;
		}

		//MAR972 GHun
		public void ProcessInboundMessage(string xmlData, string processContext)
		{
			try 
			{
				FirstTitleInboundBO inBO = new FirstTitleInboundBO(processContext);
				bool success = inBO.ProcessFirstTitleInboundResponse(xmlData);
				if (success)
					ContextUtil.SetComplete();
				else
					ContextUtil.SetAbort();
			}
			catch
			{
				ContextUtil.SetAbort();
			}
		}
		//MAR972 End
	}
}

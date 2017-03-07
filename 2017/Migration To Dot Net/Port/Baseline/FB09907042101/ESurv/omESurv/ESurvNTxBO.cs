/*
--------------------------------------------------------------------------------------------
Workfile:			ESurvNTxBO.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Creates a new transaction.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
GHun	15/12/2005	MAR821 Created
PSC		03/01/2006	MAR961 Use omLogging wrapper
PSC		14/02/2006	MAR1207 Correct resending of message
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly

--------------------------------------------------------------------------------------------
*/

//GHun	MAR821
using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using System.Xml;
using omMQ;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 03/01/2006 MAR961

namespace Vertex.Fsd.Omiga.omESurv
{
	/// <summary>
	/// This class is used to create a new transaction
	/// </summary>
	
	[ProgId("omESurv.ESurvNTxBO")]
	[ComVisible(true)]
	[Guid("0A6C411D-66B4-41c9-BD37-038CA322A8DD")]
	[Transaction(TransactionOption.RequiresNew)]
	public class ESurvNTxBO : ServicedComponent
	{
		public ESurvNTxBO()
		{
		}

		//MAR821 Moved from OutboundBO to here so it will always run in a new transaction
		// PSC 03/01/2006 MAR961
		public void AddMessageToQueue(string strData, bool bAddDelay, bool bAddToOutbound, string userId, string unitId, string userAuthorityLevel, omLogging.Logger logger)
		{
			string Response = string.Empty;
			XmlDocument xdResponse = new XmlDocument();
			const string cFunctionName = "AddMessageToQueue";

			if(logger.IsDebugEnabled)
			{
				logger.Debug(cFunctionName + ": strData=[" + strData + "]");
			}

			try
			{
				XmlNode xnNode;
				XmlDocument requestDocument = new XmlDocument();

				//Header
				XmlNode xnHeader = xdResponse.CreateElement("Header");
				xdResponse.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdResponse.CreateElement("TimeStamp");
				xnNode.InnerText  = DateTime.Today.ToString("yyyy/MM/dd");
				xnHeader.AppendChild (xnNode);

				string targetQueueName;
				if(bAddToOutbound)
					targetQueueName = new GlobalParameter("ESurvOutboundQueueName").String;
				else
					targetQueueName = new GlobalParameter("ESurvInboundQueueName").String;

				// Prepare the request for Queue
				XmlElement rootElement = requestDocument.CreateElement("REQUEST");
				rootElement.SetAttribute("OPERATION", "SendToQueue");
				requestDocument.AppendChild(rootElement);
				
				// Set up message element
				XmlElement messageElement = requestDocument.CreateElement("MESSAGEQUEUE");
				messageElement.SetAttribute("QUEUENAME", targetQueueName);
				if (bAddToOutbound)
					messageElement.SetAttribute("PROGID", "omESurv.OutBoundBO");
				else
					messageElement.SetAttribute("PROGID", "omESurv.InboundBO");
				//messageElement.SetAttribute("XML", strData);
				if (bAddDelay)
				{
					//Add an executeafter time of now + 1hour
					string datestr = DateTime.Now.AddHours(1).ToString("dd/MM/yyyy HH:mm:ss");	//MAR744 GHun
					messageElement.SetAttribute("EXECUTEAFTERDATE", datestr);	//MAR744 GHun
				}
				rootElement.AppendChild(messageElement);

				// Set up data element
				XmlDocument xmlInData = new XmlDocument();
				xmlInData.LoadXml(strData);
				XmlNode xmlDataNode = xmlInData.SelectSingleNode("XML");
				XmlNode dataElement = requestDocument.CreateElement("XML");
				messageElement.AppendChild(dataElement);

				// PSC 14/02/2007 MAR1207 - Start
				if (xmlDataNode != null)
				{
					XmlElement directGatewayRequest = requestDocument.CreateElement("REQUEST");
					directGatewayRequest.SetAttribute("USERAUTHORITYLEVEL", userAuthorityLevel);
					directGatewayRequest.SetAttribute("MACHINEID", System.Environment.MachineName);
					directGatewayRequest.SetAttribute("USERID", userId);
					directGatewayRequest.SetAttribute("UNITID", unitId);
					directGatewayRequest.SetAttribute("OPERATION", "OnMessage");
					XmlNodeList NodesToImport= xmlDataNode.ChildNodes;
					for(int i = 0; i < NodesToImport.Count; i++) //MAR597
					{
						XmlNode importedNode = requestDocument.ImportNode(NodesToImport.Item(i),true);
						directGatewayRequest.AppendChild(importedNode);
					}
						
					dataElement.AppendChild (directGatewayRequest);
				}
				else
				{
					dataElement.InnerXml = strData;
				}
				// PSC 14/02/2007 MAR1207 - End
				
				omMQBOClass omigaMessageQueue = new omMQBOClass();
				// MAR1603 M Heys 10/04/2006 start
				try
				{
					xdResponse.LoadXml(omigaMessageQueue.omRequest(requestDocument.OuterXml));
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(omigaMessageQueue);
				}
				// MAR1603 M Heys 10/04/2006 end

				string responseType = (xdResponse.DocumentElement != null) ? xdResponse.DocumentElement.GetAttribute("TYPE") : "";
				if (responseType != "SUCCESS")
				{
					if(logger.IsDebugEnabled)
					{
						logger.Debug(cFunctionName + ": error calling omMQ: " + xdResponse.OuterXml);
					}
					throw new OmigaException(cFunctionName + ": Error in MessageQueue Call: " + xdResponse.OuterXml);
				}
			}
			catch(Exception ex)
			{
				if(logger.IsDebugEnabled)
				{
					logger.Debug(cFunctionName + ": " + ex.Message);
				}
				ContextUtil.SetAbort();
				throw new OmigaException(cFunctionName + ": " + ex.Message);
			}

			ContextUtil.SetComplete();
		}
		//MAR821 End
	}
}

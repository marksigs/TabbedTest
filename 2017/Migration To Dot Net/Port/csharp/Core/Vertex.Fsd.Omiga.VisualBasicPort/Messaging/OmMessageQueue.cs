/*
--------------------------------------------------------------------------------------------
Workfile:			OmMessageQueue.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Omiga support for MSMQ.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		21/06/2000	Created
PSC		26/04/2001	SYS2292 Amend to call correct methodand return a node rather than a
					document
PSC		02/05/2001	Amend to send the correct messgage data to AsyncSend
DM		20/09/2001	SYS2716 Removed if and added case to IomMessageQueue_SendToQueue
------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from OmMessageQueue.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Messaging
{
	/// <summary>
	/// Omiga message queueing.
	/// </summary>
	public static class OmMessageQueue
	{
		/// <summary>
		/// Sends a message to a queue.
		/// </summary>
		/// <param name="xmlRequest">An xml node containing the message to send to the queue and details of how to send the message to the queue.</param>
		/// <returns>
		/// An xml response containing details about whether the message was successfully sent to the queue.
		/// </returns>
		/// <remarks>
		/// <para>
		/// The global parameter <b>MessageQueueType</b> determines which type of queue is used: 
		/// 1 for MSMQ (not implemented), 2 or 3 for database queues.
		/// </para>
		/// <para>
		/// The <paramref name="xmlRequest"/> parameter has the format:
		/// <code>
		/// &lt;REQUEST&gt;
		///	 &lt;MESSAGEQUEUE&gt;
		///		 &lt;QUEUENAME&gt;<i>queueName</i>&lt;/QUEUNAME&gt;
		///		 &lt;PROGID&gt;<i>progId</i>&lt;/PROGID&gt;
		///		 &lt;PRIORITY&gt;<i>priority</i>&lt;/PRIORITY&gt;
		///		 &lt;EXECUTEAFTERDATE&gt;<i>YYYY-MM-DD HH:MM</i>&lt;/EXECUTEAFTERDATE&gt;
		///		 &lt;CONNECTIONSTRING&gt;Provider=<i>provider</i>;Data Source=<i>dataSource</i>;User ID=<i>userId</i>;Password=<i>password</i>;&lt;/CONNECTIONSTRING&gt;
		///		 &lt;XML&gt;<i>xml</i>&lt;/XML&gt;
		///	 &lt;/MESSAGEQUEUE&gt;
		/// &lt;/REQUEST&gt;
		/// </code>
		/// <i>xml</i> is the xml that forms the message that will be placed on the queue <i>queueName</i>. 
		/// The message will be sent asynchronously to the component identified by <i>progId</i>.
		/// PRIORITY, EXECUTEAFTERDATE and CONNECTIONSTRING are optional elements. If CONNECTIONSTRING
		/// is not supplied, then the Omiga 4 connection string is used (i.e. from the registry).
		/// </para>
		/// </remarks>
		public static XmlNode SendToQueue(XmlNode xmlRequest) 
		{
			XmlDocument xmlDocument = new XmlDocument();
			
			try 
			{
				OmigaMessageQueue omigaToMessageQueue = null;
				// Get MessageQueuetype GlobalParameter
				int messageQueueType = GlobalAssist.GetGlobalParamAmount("MessageQueueType");
				// There are three types of Queue available at the moment.
				// Must ensure that the correct Global parameter is set up for the Queue type.
				// SYS2716
				switch (messageQueueType)
				{
					case 1:
						// SQL Server MSMQ
						omigaToMessageQueue = new OmigaMessageQueueMSMQ();
						break;
					case 2:
					case 3:
						// SQL Server OMMQ Oracle OMMQ
						omigaToMessageQueue = new OmigaMessageQueueOMMQ();
						break;
					default:
						// Error Message Queue type not supported.
						throw new InvalidMessageQueueTypeException("Message Queue Type not supported, check global parameter MessageQueueType");
				}

				// Make Request element based
				XmlNode xmlNode = XmlAssist.CreateElementRequestFromNode(xmlRequest, "MESSAGEQUEUE", true);
				// Get the message queue element and remove the actual message text to pass down
				// separately to AsyncSend
				XmlNode xmlMessageQueue = XmlAssist.GetMandatoryNode(xmlNode, ".//MESSAGEQUEUE");
				XmlNode xmlMessage = XmlAssist.GetMandatoryNode(xmlMessageQueue, "XML");
				xmlMessageQueue.RemoveChild(xmlMessage);
				// PSC 26/04/01
				string response = omigaToMessageQueue.AsyncSend(xmlNode.OuterXml, xmlMessage.FirstChild.OuterXml);
				xmlDocument = XmlAssist.Load(response);
				// RF 11/01/02
				ErrAssistException.CheckXmlResponseNode(xmlDocument.DocumentElement, null, true);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				ErrAssistException errAssistException = new ErrAssistException(exception, TransactionAction.SetAbort);
				errAssistException.CheckError();
			}

			return xmlDocument.DocumentElement;
		}
	}
}

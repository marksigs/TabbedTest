/*
--------------------------------------------------------------------------------------------
Workfile:			OmigaMessageQueue.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Omiga support for OMMQ.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
LD		16/10/2000	Created
AW		28/03/2001	Changed order of last 2 parameters of ADO command to match sp signature
LD		11/06/2001	SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
LD		19/06/2001	SYS2386 All projects to use guidassist.bas rather than guidassist.cls
LD		12/10/2001	SYS2705 Support for SQL Server added
LD		12/10/2001	SYS2708 Add optional element CONNECTIONSTRING
DS		04/12/2001	SYS3298 Add LOGICALQUEUNAME processing
LD		02/05/2002	SYS4414 Amendments for BLOB support, and execute after date, and the threshold
					of large message size for sql server
------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from OmigaMessageQueue.cls.
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
	/// Base class for different types of Omiga message queues.
	/// </summary>
	public abstract class OmigaMessageQueue
	{
		/// <summary>
		/// Send a message to a queue.
		/// </summary>
		/// <param name="request">An xml request containing details about how to send the message.</param>
		/// <param name="message">The message to send to the queue.</param>
		/// <returns>An xml response containing details about whether the message was successfully sent to the queue.</returns>
		public abstract string AsyncSend(string request, string message);
	}

	/// <summary>
	/// Omiga message queue using Microsoft Message Queues (MSMQ).
	/// </summary>
	public class OmigaMessageQueueMSMQ : OmigaMessageQueue
	{
		/// <summary>
		/// Send a message to an MSMQ queue.
		/// </summary>
		/// <param name="request">An xml request containing details about how to send the message.</param>
		/// <param name="message">The message to send to the queue.</param>
		/// <returns>An xml response containing details about whether the message was successfully sent to the queue.</returns>
		/// <remarks>
		/// This method is not implemented.
		/// </remarks>
		/// <exception cref="Vertex.Fsd.Omiga.VisualBasicPort.Assists.NotImplementedException">Always thrown.</exception>
		public override string AsyncSend(string request, string message)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.Assists.NotImplementedException();
		}
	}
	
	/// <summary>
	/// Omiga message queue using database queues.
	/// </summary>
	public class OmigaMessageQueueOMMQ : OmigaMessageQueue
	{
		/// <summary>
		/// Sends a message to a database queue.
		/// </summary>
		/// <param name="request">An xml request containing details about how to send the message.</param>
		/// <param name="message">The message to send to the queue.</param>
		/// <returns>An xml response containing details about whether the message was successfully sent to the queue.</returns>
		/// <remarks>
		/// The <paramref name="request"/> parameter has the format:
		/// <code>
		/// &lt;REQUEST&gt;
		///	 &lt;MESSAGEQUEUE&gt;
		///		 &lt;QUEUENAME&gt;<i>queueName</i>&lt;/QUEUNAME&gt;
		///		 &lt;PROGID&gt;<i>progId</i>&lt;/PROGID&gt;
		///		 &lt;PRIORITY&gt;<i>priority</i>&lt;/PRIORITY&gt;
		///		 &lt;EXECUTEAFTERDATE&gt;<i>YYYY-MM-DD HH:MM</i>&lt;/EXECUTEAFTERDATE&gt;
		///		 &lt;CONNECTIONSTRING&gt;Provider=<i>provider</i>;Data Source=<i>dataSource</i>;User ID=<i>userId</i>;Password=<i>password</i>;&lt;/CONNECTIONSTRING&gt;
		///	 &lt;/MESSAGEQUEUE&gt;
		/// &lt;/REQUEST&gt;
		/// </code>
		/// PRIORITY, EXECUTEAFTERDATE and CONNECTIONSTRING are optional elements. If CONNECTIONSTRING
		/// is not supplied, then the Omiga 4 connection string is used (i.e. from the registry).
		/// </remarks>
		public override string AsyncSend(string request, string message)
		{
			string response = "";

			try
			{
				// Extract the request details
				XmlDocument xmlRequest = XmlAssist.Load(request);

				XmlNode xmlNode = XmlAssist.GetMandatoryNode(xmlRequest, "REQUEST/MESSAGEQUEUE");

				string logicalQueueName = XmlAssist.GetNodeText(xmlNode, "LOGICALQUEUENAME");

				string queueName = null;
				if (logicalQueueName != null && logicalQueueName.Length > 0)
				{
					queueName = GlobalAssist.GetGlobalParamString(logicalQueueName);
				}
				else
				{
					queueName = XmlAssist.GetMandatoryNodeText(xmlNode, "QUEUENAME");
				}

				string progIdText = XmlAssist.GetMandatoryNodeText(xmlNode, "PROGID");
				string priorityText = XmlAssist.GetNodeText(xmlNode, "PRIORITY");

				// 3 is the default
				int priority = 0;
				if (priorityText == null || priorityText.Length == 0)
				{
					priority = 3;
				}
				else
				{
					priority = Convert.ToInt32(priorityText);
				}

				string executeAfterDateText = XmlAssist.GetNodeText(xmlNode, "EXECUTEAFTERDATE");
				string connectionString = XmlAssist.GetNodeText(xmlNode, "CONNECTIONSTRING");
				DBPROVIDER databaseProvider = DBPROVIDER.omiga4DBPROVIDERUnknown;
				if (connectionString == null || connectionString.Length == 0)
				{
					connectionString = AdoAssist.GetDbConnectString();
					databaseProvider = AdoAssist.GetDbProvider();
				}
				else if (connectionString.ToUpper().Contains("SQLOLEDB"))
				{
					databaseProvider = DBPROVIDER.omiga4DBPROVIDERSQLServer;
				}
				else
				{
					databaseProvider = DBPROVIDER.omiga4DBPROVIDEROracle;
				}

				bool isLargeMessage = message.Length > 2000;

				string storedProcedureName = null;
				if (databaseProvider == DBPROVIDER.omiga4DBPROVIDEROracle)
				{
					throw new Vertex.Fsd.Omiga.VisualBasicPort.Assists.NotImplementedException("Oracle support is not implemented.");
				}
				else if (databaseProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer)
				{
					storedProcedureName = isLargeMessage ? "USP_MQLOMMQSENDMESSAGENTEXT" : "USP_MQLOMMQSENDMESSAGE";
				}

				using (SqlConnection sqlConnection = new SqlConnection(connectionString))
				{
					using (SqlCommand sqlCommand = new SqlCommand(storedProcedureName, sqlConnection))
					{
						sqlCommand.CommandType = CommandType.StoredProcedure;

						SqlParameter sqlParameter = null;

						string guid = GuidAssist.CreateGUID();
						sqlCommand.Parameters.Add("p_MessageId", SqlDbType.Binary, 16).Value = SqlAssist.GuidStringToByteArray(guid);
						sqlCommand.Parameters.Add("p_QueueName", SqlDbType.NVarChar, queueName.Length).Value = queueName;
						sqlCommand.Parameters.Add("p_ProgId", SqlDbType.NVarChar, progIdText.Length).Value = progIdText;

						// threshold for large message is 2000 characters.
						SqlDbType sqlDbType = isLargeMessage ? SqlDbType.NText : SqlDbType.NVarChar;
						sqlCommand.Parameters.Add("p_Xml", sqlDbType, message.Length).Value = message;

						sqlCommand.Parameters.Add("p_Priority", SqlDbType.Int).Value = priority;

						sqlParameter = sqlCommand.Parameters.Add("p_dtExecuteAfterDate", SqlDbType.DateTime);
						sqlParameter.IsNullable = true;
						if (executeAfterDateText != null && executeAfterDateText.Length > 0)
						{
							sqlParameter.Value = ConvertAssist.CSafeDate(executeAfterDateText);
						}
						else
						{
							sqlParameter.Value = System.DBNull.Value;
						}

						// send the message
						sqlConnection.Open();
						sqlCommand.ExecuteNonQuery();
					}
				}

				response = "<RESPONSE TYPE=\"SUCCESS\"/>";

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception, TransactionAction.SetAbort).CreateErrorResponse();
			}

			return response;
		}
	}
}

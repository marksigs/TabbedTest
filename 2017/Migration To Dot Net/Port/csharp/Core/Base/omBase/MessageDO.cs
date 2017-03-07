/*
--------------------------------------------------------------------------------------------
Workfile:			MessageDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Message data object.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		21/10/99	Created
LD		07/11/00	Explicitly close recordsets
DM		16/10/01	SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		31/07/2007	First .Net version. Ported from MessageDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.MessageDO")]
	[Guid("83261909-0829-4C08-ACC8-97A85A20F450")]
	[Transaction(TransactionOption.Supported)]
	public class MessageDO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:
		// Get the data for a single instance of the persistant data associated with
		// this data object
		// pass:
		// request  xml Request data stream containing data to which identifies
		// the instance of the persistant data to be retrieved
		// return:
		// GetData		 string containing XML data stream representation of
		// data retrieved
		// Raise Errors: if record not found, raise omiga4RecordNotFound
		// ------------------------------------------------------------------------------------------
		public string GetData(string request) 
		{
			string response = "";

			try 
			{
				response = DOAssist.GetData(request, LoadMessageData());
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return response;
		}

		private string LoadMessageData() 
		{
			const string classDefinition = 
				"<TABLENAME>" +
					"MESSAGE" +
					"<PRIMARYKEY>MESSAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" +
					"<OTHERS>MESSAGETEXT<TYPE>dbdtString</TYPE></OTHERS>" +
					"<OTHERS>MESSAGETYPE<TYPE>dbdtString</TYPE></OTHERS>" +
				"</TABLENAME>";
			return classDefinition;
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// XML elements must be created for any derived values as specified.
		// Add any derived values to XML. E.g. data type 'double' fields will
		// need to be formatted as strings to required precision & rounding.
		// pass:
		// data		 base XML data stream
		// as:
		// <tablename>
		// <element1>element1 value</element1>
		// <elementn>elementn value</elementn>
		// return:
		// AddDerivedData	  base XML data stream plus any derived values
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string AddDerivedData(string data) 
		{
			return data;
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Get the data for a single instance of the persistant data associated with
		// this data object
		// pass:
		// request  xml Request data stream containing data to which identifies
		// the instance of the persistant data to be retrieved
		// return:
		// string containing XML data stream representation of
		// data retrieved
		// ------------------------------------------------------------------------------------------
		public string GetMessageDetails(string request) 
		{
			string response = "";
			
			try
			{

				XmlDocument xmlIn = XmlAssist.Load(request);

				string messageNumberText = XmlAssist.GetTagValue(xmlIn.DocumentElement, "MESSAGENUMBER");
				string cmdText = "";
				if (messageNumberText.Length == 0)
				{
					response = CreateErrorDocument(messageNumberText);
				}
				else
				{
					cmdText = "SELECT * FROM MESSAGE WHERE MESSAGENUMBER = @messageNumber";

					using (SqlCommand sqlCommand = new SqlCommand(cmdText))
					{
						sqlCommand.Parameters.Add("messageNumber", SqlDbType.Int).Value = Convert.ToInt32(messageNumberText);
						using (DataSet dataSet = AdoAssist.ExecuteGetRecordSet(sqlCommand))
						{
							DataTable dataTable = DOAssist.GetDataTable(dataSet);
							if (dataTable == null || dataTable.Rows.Count == 0)
							{
								response = CreateErrorDocument(messageNumberText);
							}
							else
							{
								string classDefinition = LoadMessageData();
								XmlDocument xmlOut = new XmlDocument();
								foreach (DataRow dataRow in dataTable.Rows)
								{
									XmlDocument xmlDoc = XmlAssist.Load(DOAssist.GetXMLFromRecordSet(dataRow, classDefinition));
									xmlOut.AppendChild(xmlOut.ImportNode(xmlDoc.DocumentElement, true));
								}
								response = xmlOut.OuterXml;
							}
						}
					}
				}
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}

		private string CreateErrorDocument(string messageNumberText) 
		{
			XmlDocument xmlOut = new XmlDocument();
			string messageText = "Message number " + messageNumberText + " not found";
			XmlElement xmlTableElem = xmlOut.CreateElement("MESSAGE");
			xmlOut.AppendChild(xmlTableElem);
			XmlElement xmlElem = xmlOut.CreateElement("MESSAGENUMBER");
			xmlElem.InnerText = messageNumberText;
			xmlTableElem.AppendChild(xmlElem);
			xmlElem = xmlOut.CreateElement("MESSAGETEXT");
			xmlElem.InnerText = messageText;
			xmlTableElem.AppendChild(xmlElem);
			xmlElem = xmlOut.CreateElement("MESSAGETYPE");
			xmlElem.InnerText = "Error";
			xmlTableElem.AppendChild(xmlElem);
			ErrAssistException exception = new ErrAssistException(OMIGAERROR.InvalidMessageNo, messageText);

			return xmlOut.OuterXml;
		}

		protected override bool CanBePooled()
		{
			return true;
		}
	}
}

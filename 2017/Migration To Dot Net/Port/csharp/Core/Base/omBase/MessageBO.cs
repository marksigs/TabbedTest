/*
--------------------------------------------------------------------------------------------
Workfile:			MessageBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Message Business Object which 'supports transactions' only.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		30/06/99	Created
RF		30/09/99	Applied changes raised by code review of 30/09/99, including:
					removed AnonInterfaceFunction
					removed calls to Validate
					improved error handling
RF		04/10/99	Added profiling
DRC		3/10/01	 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		31/07/2007	First .Net version. Ported from MessageBO.cls.
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
	[ProgId("omBase.MessageBO")]
	[Guid("B1AC8FE9-B23A-4680-944C-7DCC4C0D38B3")]
	[Transaction(TransactionOption.Supported)]
	public class MessageBO : ServicedComponent
	{
		private const string _tableName = "MESSAGE";

		// header ----------------------------------------------------------------------------------
		// description:  Get the data for a single instance of the persistant data associated with
		// this data object
		// pass:		 request  xml Request data stream containing data to which identifies
		// instance of the persistant data to be retrieved
		// return:					   xml Response data stream containing results of operation
		// either: TYPE="SUCCESS" and xml representation of data
		// or: TYPE="SYSERR" and <ERROR> element
		// ------------------------------------------------------------------------------------------
		public string GetData(string request) 
		{
			string response = "";

			try
			{
				XmlDocument xmlIn = XmlAssist.Load(request);
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElem = xmlOut.CreateElement("RESPONSE");
				XmlNode xmlDataNode = xmlOut.AppendChild(xmlResponseElem);
				xmlResponseElem.SetAttribute("TYPE", "SUCCESS");
				XmlNode xmlRequestNode = xmlIn.GetElementsByTagName(_tableName)[0];
				// call Data Object GetData function
				if (xmlRequestNode == null)
				{
					throw new MissingPrimaryTagException(_tableName + " tag not found");
				}

				string data = "";
				using (MessageDO messageDO = new MessageDO())
				{
					data = messageDO.GetData(xmlRequestNode.OuterXml);
				}
				XmlDocument xmlData = XmlAssist.Load(data);
				xmlDataNode.AppendChild(xmlOut.ImportNode(xmlData.DocumentElement, true));
				response = xmlOut.OuterXml;
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
	}
}

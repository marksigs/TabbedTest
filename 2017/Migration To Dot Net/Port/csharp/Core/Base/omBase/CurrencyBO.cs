/*
--------------------------------------------------------------------------------------------
Workfile:			CurrencyDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Collects all Currency data for use with the Currency Calculator
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PF		09/04/2001	Created
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		24/07/2007	First .Net version. Ported from CurrencyBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.CurrencyBO")]
	[Guid("75DAF8CE-9490-41D6-A958-A5EF89317758")]
	[Transaction(TransactionOption.Supported)]
	public class CurrencyBO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:  Get all instances of the persistant data associated with this
		// business object
		// pass:		 vstrXmlRequest  xml Request data stream containing data to be persisted
		// return:					   xml Response data stream containing results of operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// ------------------------------------------------------------------------------------------
		public string FindList() 
		{
			string response = null;

			try
			{
				// Create default response block
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");

				// Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
				XmlNode xmlTempResponseNode = xmlOut.ImportNode(FindListAsXml(), true);
				ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElement, true);
				XmlAssist.AttachResponseData((XmlNode)xmlResponseElement, (XmlElement)xmlTempResponseNode);
				response = xmlResponseElement.OuterXml;
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

		// header ----------------------------------------------------------------------------------
		// description:  Get all instances of the persistant data associated with this
		// business object
		// pass:		 vstrXmlRequest  xml Request data stream containing data to be persisted
		// return:					   xml Response data stream containing results of operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// ------------------------------------------------------------------------------------------
		public XmlNode FindListAsXml() 
		{
			XmlNode xmlResponseNode = null;

			try
			{
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlNode xmlDataNode = null;
				using (CurrencyDO currencyDO = new CurrencyDO())
				{
					xmlDataNode = currencyDO.FindList();
				}
				xmlResponseElement.InnerXml = xmlDataNode.OuterXml;
				xmlResponseNode = xmlResponseElement;
			}
			catch (Exception exception)
			{
				xmlResponseNode = new ErrAssistException(exception).CreateErrorResponseNode();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return xmlResponseNode;
		}
	}
}

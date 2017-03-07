using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Messaging;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class OmigaMessageQueueTest
	{
		[Test]
		public void TestAsyncSend()
		{
			string request =
				"<REQUEST>" +
					"<MESSAGEQUEUE>Test" +
						/*
						"<QUEUENAME>astanley</QUEUENAME>" +
						"<PROGID>omBase.Combo</PROGID>" +
						"<PRIORITY>5</PRIORITY>" +
						"<EXECUTEAFTERDATE>20/06/2007 09:12:34</EXECUTEAFTERDATE>" +
						 */
					"</MESSAGEQUEUE>" +
				"</REQUEST>";
			string message =
				"<TABLENAME>" +
					"CUSTOMERADDRESS" +
					"<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" +
					"<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" +
					"<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" +
					"<OTHERS>ADDRESSTYPE<TYPE>dbdtComboId</TYPE>" +
					"<COMBO>CustomerAddressType</COMBO></OTHERS>" +
					"<OTHERS>DATEMOVEDIN<TYPE>dbdtDate</TYPE></OTHERS>" +
					"<OTHERS>DATEMOVEDOUT<TYPE>dbdtDate</TYPE></OTHERS>" +
					"<OTHERS>NATUREOFOCCUPANCY<TYPE>dbdtComboId</TYPE>" +
					"<COMBO>NatureOfOccupancy</COMBO></OTHERS>" +
					"<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" +
					"<OTHERS>LASTAMENDEDDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" +
				"</TABLENAME>";
			OmigaMessageQueueOMMQ omigaToMessageQueue = new OmigaMessageQueueOMMQ();
			string response = omigaToMessageQueue.AsyncSend(request, message);

			Assert.AreEqual(response, "<RESPONSE TYPE=\"SUCCESS\"/>");
		}

		[Test]
		public void TestSendToQueue()
		{
			string request =
				"<REQUEST NAME='Test'>" +
					"<MESSAGEQUEUE>" +
						"<QUEUENAME>astanley</QUEUENAME>" +
						"<PROGID>omBase.Combo</PROGID>" +
						"<PRIORITY>5</PRIORITY>" +
						"<EXECUTEAFTERDATE>20/06/2007 09:12:34</EXECUTEAFTERDATE>" +
//						"QUEUENAME='astanley' " +
//						"PROGID='omBase.Combo' " +
//						"PRIORITY='5' " +
//						"TITLE_TEXT='Mr' " +
//						"EXECUTEAFTERDATE='20/06/2007 09:12:34'>" +
						"<XML>" +
							"<TABLENAME ENTITY='Entity1' COLUMN='Column1'>" +
								"CUSTOMERADDRESS" +
								"<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" +
								"<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" +
								"<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" +
								"<OTHERS>ADDRESSTYPE<TYPE>dbdtComboId</TYPE>" +
								"<COMBO>CustomerAddressType</COMBO></OTHERS>" +
								"<OTHERS>DATEMOVEDIN<TYPE>dbdtDate</TYPE></OTHERS>" +
								"<OTHERS>DATEMOVEDOUT<TYPE>dbdtDate</TYPE></OTHERS>" +
								"<OTHERS>NATUREOFOCCUPANCY<TYPE>dbdtComboId</TYPE>" +
								"<COMBO>NatureOfOccupancy</COMBO></OTHERS>" +
								"<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" +
								"<OTHERS>LASTAMENDEDDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" +
							"</TABLENAME>" +
						"</XML>" +
					"</MESSAGEQUEUE>" +
				"</REQUEST>";
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlNode xmlResponse = OmMessageQueue.SendToQueue(xmlRequest.DocumentElement);
			string response = xmlResponse != null ? xmlResponse.OuterXml : "";
			Assert.AreEqual(response, "<RESPONSE TYPE=\"SUCCESS\" />");
		}
	}
}

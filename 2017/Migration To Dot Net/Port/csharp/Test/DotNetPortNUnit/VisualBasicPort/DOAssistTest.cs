using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class DOAssistTest
	{
		[Test]
		public void TestGetData()
		{
			string request = 
				"<REQUEST>" + 
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000000</CUSTOMERNUMBER>" + 
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition = 
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
			string response = null;
			try
			{
				response = DOAssist.GetData(request, classDefinition);
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception, TransactionAction.SetOnErrorType).CreateErrorResponse();
			}

			request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>DoesNotExist</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			try
			{
				response = DOAssist.GetData(request, classDefinition);
				Assert.Fail();
			}
			catch (ErrAssistException exception)
			{
				exception.SetOnErrorType();
			}
		}

		[Test]
		public void TestGetDataEx()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000000</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
				"<TABLENAME>" +
					"CUSTOMERADDRESS" +
					"<SQLNOLOCK/>" + 
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlDocument xmlClassDefinition = new XmlDocument();
			xmlClassDefinition.LoadXml(classDefinition);
			XmlNode xmlNode = DOAssist.GetDataEx(xmlRequest.DocumentElement, xmlClassDefinition);
			string response = xmlNode.OuterXml;
		}

		[Test]
		public void TestGetComponentData()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000000</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
					"<CRITERIA>" +
						"<CUSTOMERNUMBER>10000000</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CRITERIA>" + 
					"<OUTPUT ALLFIELDS='0'/>" +
					"<FIELD>ADDRESSTYPE</FIELD>" +
					"<FIELD>DATEMOVEDIN</FIELD>" +
					"<FIELD>DATEMOVEDOUT</FIELD>" +
				"</REQUEST>";
			string classDefinition =
				"<TABLENAME>" +
					"CUSTOMERADDRESS" +
					"<SQLNOLOCK/>" +
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlDocument xmlClassDefinition = new XmlDocument();
			xmlClassDefinition.LoadXml(classDefinition);
			XmlNode xmlNode = DOAssist.GetComponentData(xmlRequest.DocumentElement, xmlClassDefinition);
			string response = xmlNode.OuterXml;
		}

		[Test]
		public void TestDelete()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>DoesNotExist</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			try
			{
				DOAssist.Delete(request, classDefinition);
				Assert.Fail();
			}
			catch (NoRowsAffectedException)
			{
			}
		}

		[Test]
		public void TestFindList()
		{
			string request = 
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			string response = DOAssist.FindList(request, classDefinition);
		}

		[Test]
		public void TestFindListEx()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlDocument xmlClassDefinition = new XmlDocument();
			xmlClassDefinition.LoadXml(classDefinition);
			XmlNode xmlNode = DOAssist.FindListEx(xmlRequest.DocumentElement, xmlClassDefinition);
			string response = xmlNode.OuterXml;
		}

		[Test]
		public void TestFindListMultipleEx()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlDocument xmlClassDefinition = new XmlDocument();
			xmlClassDefinition.LoadXml(classDefinition);
			XmlNode xmlNode = DOAssist.FindListMultipleEx(xmlRequest.DocumentElement, xmlClassDefinition);
			string response = xmlNode.OuterXml;
		}

		[Test]
        public void TestCreate()
        {
            string request =
                "<REQUEST>" +
                    "<CUSTOMERADDRESS>" +
                        "<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
                        "<CUSTOMERVERSIONNUMBER>5</CUSTOMERVERSIONNUMBER>" +
                        "<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
                    "</CUSTOMERADDRESS>" +
                "</REQUEST>";
            string classDefinition =
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
            DOAssist.Create(request, classDefinition);
        }

        [Test]
        public void TestUpdate()
        {
            string request =
                "<REQUEST>" +
                    "<CUSTOMERADDRESS>" +
                        "<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
                        "<CUSTOMERVERSIONNUMBER>5</CUSTOMERVERSIONNUMBER>" +
                        "<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
                        "<ADDRESSTYPE>1</ADDRESSTYPE>" +
                    "</CUSTOMERADDRESS>" +
                "</REQUEST>";
            string classDefinition =
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
            DOAssist.Update(request, classDefinition);
        }

        [Test]
        public void TestGetNextSequenceNumber()
        {
            string request =
                "<REQUEST>" +
                    "<CUSTOMERADDRESS>" +
                        "<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
                        "<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
                        "<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
                    "</CUSTOMERADDRESS>" +
                "</REQUEST>";
            string classDefinition =
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
            int sequenceNumber = DOAssist.GetNextSequenceNumber(request, classDefinition, "CUSTOMERADDRESSSEQUENCENUMBER");
            Assert.IsTrue(sequenceNumber == 2);
        }
	
		[Test]
		public void TestGenerateSequenceNumber()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			DOAssist.GenerateSequenceNumber(xmlRequest, classDefinition, "CUSTOMERADDRESSSEQUENCENUMBER");
		}

		[Test]
		public void TestGenerateSequenceNumberEx()
		{
			string request =
				"<REQUEST>" +
					"<CUSTOMERADDRESS>" +
						"<CUSTOMERNUMBER>10000001</CUSTOMERNUMBER>" +
						"<CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>" +
						"<CUSTOMERADDRESSSEQUENCENUMBER>1</CUSTOMERADDRESSSEQUENCENUMBER>" +
					"</CUSTOMERADDRESS>" +
				"</REQUEST>";
			string classDefinition =
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
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlDocument xmlClassDefinition = new XmlDocument();
			xmlClassDefinition.LoadXml(classDefinition);
			DOAssist.GenerateSequenceNumberEx(xmlRequest.DocumentElement, xmlClassDefinition, "CUSTOMERADDRESSSEQUENCENUMBER");
		}
	}
}

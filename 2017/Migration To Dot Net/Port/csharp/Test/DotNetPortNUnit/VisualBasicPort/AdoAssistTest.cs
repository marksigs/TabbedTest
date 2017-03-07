using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class AdoAssistTest
	{
		[Test]
		public void TestBuildDbConnectionString()
		{
			Assert.IsNotEmpty(AdoAssist.GetDbConnectString());
		}

		[Test]
		public void TestGetRecordSetAsXMLApplication()
		{
			XmlDocument xmlRequestDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();
			XmlDocument xmlResponseDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlRequestDocument.Load(assemblyDirectory + "\\TestGetRecordSetAsXML.application.request.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestGetRecordSetAsXML.application.schema.xml");
			XmlElement xmlElement = xmlResponseDocument.CreateElement("RESPONSE");
			XmlNode xmlResponseNode = xmlResponseDocument.AppendChild(xmlElement);
			AdoAssist.GetRecordSetAsXML(xmlRequestDocument.DocumentElement, xmlSchemaDocument.DocumentElement, xmlResponseNode);
			Assert.IsTrue(xmlResponseNode.OuterXml.Length > 0);
		}

		[Test]
		public void TestGetAsXMLApplication()
		{
			XmlDocument xmlRequestDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();
			XmlDocument xmlResponseDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlRequestDocument.Load(assemblyDirectory + "\\TestGetRecordSetAsXML.application.request.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestGetRecordSetAsXML.application.schema.xml");
			XmlElement xmlElement = xmlResponseDocument.CreateElement("RESPONSE");
			XmlNode xmlResponseNode = xmlResponseDocument.AppendChild(xmlElement);
			AdoAssist adoAssist = new AdoAssist(Assembly.GetExecutingAssembly());
			adoAssist.GetAsXML(xmlRequestDocument.DocumentElement, xmlResponseNode, "APPLICATION");
			Assert.IsTrue(xmlResponseNode.OuterXml.Length > 0);
		}

		[Test]
		public void TestGetGeneratedKeyGuid()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestGetGeneratedKey.guid.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestGetGeneratedKey.guid.schema.xml");
			//AdoAssist.GetGeneratedKey(xmlDataDocument.SelectSingleNode("//RESPONSE/CONTACTDETAILS"), xmlSchemaDocument.SelectSingleNode("//CONTACTDETAILS/CONTACTDETAILSGUID"));
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestGetGeneratedKeyShort()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestGetGeneratedKey.short.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestGetGeneratedKey.short.schema.xml");
			//AdoAssist.GetGeneratedKey(xmlDataDocument.SelectSingleNode("//RESPONSE/ADDITIONALBORROWINGFEE"), xmlSchemaDocument.SelectSingleNode("//ADDITIONALBORROWINGFEE/ADDITIONALBORROWINGFEESET"));
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestCreateAdditionalBorrowingFee()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestCreate.additionalborrowingfee.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestCreate.additionalborrowingfee.schema.xml");
			AdoAssist.Create(xmlDataDocument.SelectSingleNode("//RESPONSE/ADDITIONALBORROWINGFEE"), xmlSchemaDocument.SelectSingleNode("//ADDITIONALBORROWINGFEE"));
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestCreateFromNodeAdditionalBorrowingFee()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestCreate.additionalborrowingfee.data.xml");
			AdoAssist adoAssist = new AdoAssist(Assembly.GetExecutingAssembly());
			adoAssist.CreateFromNode(xmlDataDocument.SelectSingleNode("//RESPONSE/ADDITIONALBORROWINGFEE"), "ADDITIONALBORROWINGFEE");
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestUpdateAdditionalBorrowingFee()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestUpdate.additionalborrowingfee.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestUpdate.additionalborrowingfee.schema.xml");
			AdoAssist.Update(xmlDataDocument.SelectSingleNode("//RESPONSE/ADDITIONALBORROWINGFEE"), xmlSchemaDocument.SelectSingleNode("//ADDITIONALBORROWINGFEE"));
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestDeleteAdditionalBorrowingFee()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestDelete.additionalborrowingfee.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestDelete.additionalborrowingfee.schema.xml");
			AdoAssist.Delete(xmlDataDocument.SelectSingleNode("//RESPONSE/ADDITIONALBORROWINGFEE"), xmlSchemaDocument.SelectSingleNode("//ADDITIONALBORROWINGFEE"));
			Assert.IsTrue(xmlDataDocument.OuterXml.Length > 0);
		}

		[Test]
		public void TestGetStoredProcAsXML()
		{
			XmlDocument xmlDataDocument = new XmlDocument();
			XmlDocument xmlSchemaDocument = new XmlDocument();
			XmlDocument xmlResponseDocument = new XmlDocument();

			string assemblyDirectory = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
			xmlDataDocument.Load(assemblyDirectory + "\\TestGetStoredProcAsXML.data.xml");
			xmlSchemaDocument.Load(assemblyDirectory + "\\TestGetStoredProcAsXML.schema.xml");
			XmlElement xmlElement = xmlResponseDocument.CreateElement("RESPONSE");
			XmlNode xmlResponseNode = xmlResponseDocument.AppendChild(xmlElement);
			AdoAssist.GetStoredProcAsXML(xmlDataDocument.DocumentElement, xmlSchemaDocument.SelectSingleNode("//PT_SCHEMA/DOCUMENTS"), xmlResponseNode);
			Assert.IsTrue(xmlResponseNode.OuterXml.Length > 0);
		}

		[Test]
		public void TestLoadSchema()
		{
			AdoAssist adoAssist1 = new AdoAssist(System.Reflection.Assembly.GetExecutingAssembly());
			Assert.IsTrue(adoAssist1.GetSchema("DOCUMENTS") != null);
			AdoAssist adoAssist2 = new AdoAssist(System.Reflection.Assembly.GetExecutingAssembly());
			Assert.IsTrue(adoAssist2.GetSchema("DOCUMENTS") != null);
		}

		[Test]
		public void TestExecuteSqlCommand()
		{
			int recordsAffected = AdoAssist.ExecuteSQLCommand(
				"insert into account values (0x100ED7EEB43F477B9FEFB837D2668FEB, null, null, null, null, null, null, null)");
			Assert.IsTrue(recordsAffected > 0);
		}

		[Test]
		public void TestCheckSingleRecordExists()
		{
			AdoAssist.CheckSingleRecordExists("application", "applicationnumber = '0500414'");
		}

		[Test]
		public void TestGetNumberOfRecords()
		{
			int numberOfRecords = AdoAssist.GetNumberOfRecords("application", "applicationnumber = '0500414'");
			Assert.IsTrue(numberOfRecords == 1);
		}

		[Test]
		public void TestGetValueFromTable()
		{
			//object value = AdoAssist.GetValueFromTable("account", "accountguid = 0x100ED7EEB43F477B9FEFB837D2668FEA", "accountnumber");
			//object value = AdoAssist.GetValueFromTable("account", "accountguid = 0x11B6D52A9C194082AA6386679088A3BF", "accountnumber");
			object value = AdoAssist.GetValueFromTable("account", "accountguid = 0x11B6D52A9C194082AA6386679088A3BF", "accountguid");
			Assert.IsTrue(value != null);
		}
	}
}

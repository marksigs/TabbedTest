using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class XmlAssistTest
	{
		[Test]
		public void TestLoad()
		{
			XmlDocument xmlDocument = XmlAssist.Load("<REQUEST />");
			Assert.AreEqual(xmlDocument.OuterXml, "<REQUEST />");
			try
			{
				XmlAssist.Load("<REQUEST>");
				Assert.Fail("ErrAssistException not thrown");
			}
			catch (ErrAssistException)
			{
			}
		}

		[Test]
		public void TestSetAttributeValue()
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml("<REQUEST><DATA/></REQUEST>");
			XmlAssist.SetAttributeValue(xmlDocument.DocumentElement, "TYPE", "SUCCESS");
			Assert.AreEqual(xmlDocument.DocumentElement.GetAttribute("TYPE"), "SUCCESS");
		}

		[Test]
		public void TestGetRequestNode()
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml("<MESSAGE><REQUEST><DATA/></REQUEST></MESSAGE>");
			XmlNode xmlNode = XmlAssist.GetRequestNode(xmlDocument.DocumentElement);
			Assert.IsTrue(xmlNode != null);
		}

		[Test]
		public void TestMakeNodeElementBased()
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(
				"<REQUEST " +
					"ATTRIBUTE1='A1' " + 
					"ATTRIBUTE2='A2' " + 
					"ATTRIBUTE3='A3' " + 
					">" +
					"<ELEMENT1 A1='1' A2='2' A3='3'>" +
						"<ELEMENT2>" +
						"</ELEMENT2>" +
					"</ELEMENT1>" +
				"</REQUEST>");
			XmlNode xmlNode = XmlAssist.MakeNodeElementBased(xmlDocument.DocumentElement, true, null, "ATTRIBUTE1", "ATTRIBUTE2");
			Assert.IsTrue(xmlNode != null);
		}
	}
}

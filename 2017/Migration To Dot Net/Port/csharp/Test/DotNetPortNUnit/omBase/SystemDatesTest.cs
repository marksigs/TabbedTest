using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.omBase;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class SystemDatesTest
	{
		[Test]
		public void TestSystemDatesDOCheckNonWorkingOccurence()
		{
			SystemDatesDO systemDatesDO = new SystemDatesDO();
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(
				"<REQUEST>" + 
					"<DATE>25/12/2006</DATE>" +
					"<CHANNELID>1</CHANNELID>" +
				"</REQUEST>");
			XmlNode xmlNode = systemDatesDO.CheckNonWorkingOccurence(xmlDocument.DocumentElement);
			Assert.AreEqual(xmlNode.InnerText, "1");
			xmlDocument.LoadXml(
				"<REQUEST>" +
					"<DATE>01/12/2006</DATE>" +
					"<CHANNELID>1</CHANNELID>" +
				"</REQUEST>");
			xmlNode = systemDatesDO.CheckNonWorkingOccurence(xmlDocument.DocumentElement);
			Assert.AreEqual(xmlNode.InnerText, "0");
		}

		[Test]
		public void TestSystemDatesBOCheckNonWorkingOccurence()
		{
			SystemDatesBO systemDatesBO = new SystemDatesBO();
			XmlDocument xmlDocument = new XmlDocument();
			string request = 
				"<REQUEST>" +
					"<SYSTEMDATE>" +
						"<DATE>25/12/2006</DATE>" +
						"<CHANNELID>1</CHANNELID>" +
					"</SYSTEMDATE>" +
				"</REQUEST>";
			xmlDocument.LoadXml(request);
			XmlNode xmlNode = systemDatesBO.CheckNonWorkingOccurence(xmlDocument.DocumentElement);
			Assert.AreEqual(xmlNode.InnerText, "1");
			string response = systemDatesBO.CheckNonWorkingOccurence(request);

			request =
				"<REQUEST>" +
					"<SYSTEMDATE>" +
						"<DATE>01/12/2006</DATE>" +
						"<CHANNELID>1</CHANNELID>" +
					"</SYSTEMDATE>" +
				"</REQUEST>";
			xmlDocument.LoadXml(request);
			xmlNode = systemDatesBO.CheckNonWorkingOccurence(xmlDocument.DocumentElement);
			Assert.AreEqual(xmlNode.InnerText, "0");
		}

		[Test]
		public void TestSystemDatesBOFindWorkingDay()
		{
			SystemDatesBO systemDatesBO = new SystemDatesBO();
			string response = systemDatesBO.FindWorkingDay(
				"<REQUEST>" +
					"<SYSTEMDATE>" + 
						"<DATE>01/12/2006</DATE>" +
						"<DIRECTION>-</DIRECTION>" +
						"<DAYS>2</DAYS>" +
						"<HOURS>12</HOURS>" +
						"<CHANNELID>1</CHANNELID>" + 
					"</SYSTEMDATE>" +
				"</REQUEST>");
		}
	}
}

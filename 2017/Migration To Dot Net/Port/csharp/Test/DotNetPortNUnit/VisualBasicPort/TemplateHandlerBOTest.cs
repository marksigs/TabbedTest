using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Printing;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class TemplateHandlerBOTest
	{
		private TemplateHandlerBO _templateHandlerBO = new TemplateHandlerBO();

		[Test]
		public void TestGetTemplate()
		{
			string request =
				"<REQUEST>" +
					"<TEMPLATE>10</TEMPLATE>" +
				"</REQUEST>";
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlNode xmlNode = _templateHandlerBO.GetTemplate(xmlRequest.DocumentElement);
			string response = xmlNode.OuterXml;
		}

		[Test]
		public void TestGetTemplateString()
		{
			string request =
				"<REQUEST>" +
					"<TEMPLATE>" +
						"<SECURITYLEVEL>10</SECURITYLEVEL>" +
						"<STAGENUMBER>0</STAGENUMBER>" +
						"<LANGUAGE>1</LANGUAGE>" +
					"</TEMPLATE>" +
				"</REQUEST>";
			string response = _templateHandlerBO.GetTemplate(request);
		}

		[Test]
		public void TestFindAvailableTemplates()
		{
			string request =
				"<REQUEST>" +
					"<TEMPLATE>" + 
						"<SECURITYLEVEL>10</SECURITYLEVEL>" +
						"<STAGENUMBER>0</STAGENUMBER>" +
						"<LANGUAGE>1</LANGUAGE>" +
					"</TEMPLATE>" +
				"</REQUEST>";
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlNode xmlNode = _templateHandlerBO.FindAvailableTemplates(xmlRequest.DocumentElement);
			string response = xmlNode.OuterXml;
		}

		[Test]
		public void TestFindAvailableTemplatesString()
		{
			string request =
				"<REQUEST>" +
					"<TEMPLATE>" +
						"<SECURITYLEVEL>10</SECURITYLEVEL>" +
						"<STAGENUMBER>0</STAGENUMBER>" +
						"<LANGUAGE>1</LANGUAGE>" +
					"</TEMPLATE>" +
				"</REQUEST>";
			string response = _templateHandlerBO.FindAvailableTemplates(request);
		}
	}
}

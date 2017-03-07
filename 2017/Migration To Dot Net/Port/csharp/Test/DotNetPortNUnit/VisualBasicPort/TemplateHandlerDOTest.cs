using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Printing;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class TemplateHandlerDOTest
	{
		private TemplateHandlerDO _templateHandlerDO = new TemplateHandlerDO();

		[Test]
		public void TestFindAvailableTemplates()
		{
			string request =
				"<REQUEST>" +
					"<SECURITYLEVEL>10</SECURITYLEVEL>" +
					"<STAGENUMBER>0</STAGENUMBER>" +
					"<LANGUAGE>1</LANGUAGE>" +
				"</REQUEST>";
			XmlDocument xmlRequest = new XmlDocument();
			xmlRequest.LoadXml(request);
			XmlNode xmlNode = _templateHandlerDO.FindAvailableTemplates(xmlRequest.DocumentElement);
			string response = xmlNode.OuterXml;
		}
	}
}

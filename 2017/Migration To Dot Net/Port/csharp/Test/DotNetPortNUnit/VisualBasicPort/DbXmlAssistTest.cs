using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class DbXmlAssistTest
	{
		[Test]
		public void Test()
		{
			XmlNode xmlNode1 = DbXmlAssist.GetCurrentParameterXML("DocumentMangementSystemType");
			Assert.IsNotNull(xmlNode1);
			XmlNode xmlNode2 = DbXmlAssist.GetTemplateXML("1");
			Assert.IsNotNull(xmlNode2);
		}

	}
}

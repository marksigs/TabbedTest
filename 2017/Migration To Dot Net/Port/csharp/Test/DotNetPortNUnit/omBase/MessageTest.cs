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
	public class MessageTest
	{
		[Test]
		public void TestMessageBOGetData()
		{
			MessageBO messageBO = new MessageBO();
			string response = messageBO.GetData(
				"<REQUEST>" +
					"<MESSAGE>" + 
						"<MESSAGENUMBER>500</MESSAGENUMBER>" +
					"</MESSAGE>" +
				"</REQUEST>");
		}

		[Test]
		public void TestMessageDOGetMessageDetails()
		{
			MessageDO messageDO = new MessageDO();
			string response = messageDO.GetMessageDetails(
				"<REQUEST>" + 
					"<MESSAGENUMBER>500</MESSAGENUMBER>" +
				"</REQUEST>");
		}
	}
}

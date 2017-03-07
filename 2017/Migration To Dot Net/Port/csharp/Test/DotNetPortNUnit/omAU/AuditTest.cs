using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.omAU;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class AuditTest
	{
		[Test]
		public void TestAuditDOCreateAccessAudit()
		{
			using (AuditDO auditDO = new AuditDO())
			{
				XmlDocument xmlDocument = new XmlDocument();
				string request =
					"<ACCESSAUDIT>" +
						"<USERID>astanley</USERID>" +
						"<AUDITRECORDTYPE>L</AUDITRECORDTYPE>" +
						"<MACHINEID>CH019505</MACHINEID>" +
						"<SUCCESSINDICATOR>1</SUCCESSINDICATOR>" +
						"<ONBEHALFOFUSERID></ONBEHALFOFUSERID>" +
						"<APPLICATIONNUMBER></APPLICATIONNUMBER>" +
						"<PASSWORDCREATIONDATE>optional</PASSWORDCREATIONDATE>" +
					"</ACCESSAUDIT>";
				xmlDocument.LoadXml(request);
				((IAuditDO)auditDO).CreateAccessAudit(xmlDocument.DocumentElement);
			}
		}
	}
}

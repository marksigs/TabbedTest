using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class EncryptAssistTest
	{
		public EncryptAssistTest()
		{
		}

		[Test]
		public void TestEncrypt()
		{
			Assert.AreEqual(EncryptAssist.Encrypt("password"), "~Ym~R~K~f~P~[}");
		}
	}
}

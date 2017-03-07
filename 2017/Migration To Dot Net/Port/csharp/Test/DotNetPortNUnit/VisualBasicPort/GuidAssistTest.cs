using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class GuidAssistTest
	{
		public GuidAssistTest()
		{
		}

		[Test]
		public void TestCreateGUID()
		{
			string guidText = GuidAssist.CreateGUID();
			Assert.AreEqual(guidText.Length, 32);
		}
	}
}

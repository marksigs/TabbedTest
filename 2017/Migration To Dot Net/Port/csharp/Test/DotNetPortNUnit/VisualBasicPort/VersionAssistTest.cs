using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class VersionAssistTest
	{
		[Test]
		public void TestGetVersionList()
		{
			string response = VersionAssist.GetVersionList();
		}
	}
}

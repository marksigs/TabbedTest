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
	public class CurrencyTest
	{
		[Test]
		public void TestCurrencyBOFindList()
		{
			using (CurrencyBO currencyBO = new CurrencyBO())
			{
				string response = currencyBO.FindList();
			}
		}

		[Test]
		public void TestCurrencyDOFindList()
		{
			using (CurrencyDO currencyDO = new CurrencyDO())
			{
				XmlNode xmlNode = currencyDO.FindList();
			}
		}
	}
}

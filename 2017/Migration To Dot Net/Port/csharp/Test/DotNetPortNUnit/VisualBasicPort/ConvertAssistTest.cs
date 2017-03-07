using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class ConvertAssistTest
	{
		public ConvertAssistTest()
		{
		}

		[Test]
		public void TestByteArrays()
		{
			byte[] byteArray = new byte[6];
			const string inputText = "ABCDE";
			ConvertAssist.StringToByteArray(ref byteArray, inputText);
			string outputText = ConvertAssist.ByteArrayToString(byteArray);
			Assert.AreEqual(inputText, outputText);
		}

		[Test]
		public void TestCSafeDate()
		{
			DateTime safe = ConvertAssist.CSafeDate("30/12/2007 13:14:15");
			Assert.AreEqual(safe.ToString(), new DateTime(2007, 12, 30, 13, 14, 15).ToString());
		}

		[Test]
		public void TestCSafeLng()
		{
			int safe = ConvertAssist.CSafeLng("1234");
			Assert.AreEqual(safe, 1234);
		}

		[Test]
		public void TestCSafeDbl()
		{
			double safe = ConvertAssist.CSafeDbl("1234.5678");
			Assert.AreEqual(safe, 1234.5678);
		}

		[Test]
		public void TestCSafeByte()
		{
			byte safe = ConvertAssist.CSafeByte("A");
			Assert.AreEqual(safe, 65);
		}

		[Test]
		public void TestCSafeInt()
		{
			int safe = ConvertAssist.CSafeInt("1234");
			Assert.AreEqual(safe, 1234);
		}

		[Test]
		public void TestCSafeBool()
		{
			bool safe = ConvertAssist.CSafeBool("true");
			Assert.AreEqual(safe, true);
		}
	}
}

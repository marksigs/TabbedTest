using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Printing;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class GeminiTest
	{
		public GeminiTest()
		{
		}

		[Test]
		public void TestIsFileVersionLocked()
		{
			GeminiDocument document = new GeminiDocument();
			document.FileContentsGuid = "1120C206113D48479DBA8FF170CD3460";
			document.DocumentVersion = "V1";
			string userId = null;
			//Gemini.IsFileVersionLocked(document, out userId);
			Assert.IsTrue(userId.Length > 0);
		}

		[Test]
		public void TestSendToFulfillment()
		{
			GeminiPack pack = new GeminiPack();
			pack.ApplicationNumber = "1003319";
			pack.PackFulfillmentGuid = "12690963310043278C6FC276786FDD73";
			GeminiDocument document = new GeminiDocument();
			document.FileContentsGuid = "B7CB4E8C6B614028991088C03BDA484A";
			document.DocumentDetails.FileContentsType = "FILECONTENTSTYPE";
			document.DocumentDetails.CompressionMethod = "COMPRESSIONMETHOD";
			document.DocumentDetails.DeliveryType = 10;
			document.DocumentDetails.FileContents = "FILECONTENTS";
			pack.Documents.Add(document);
			List<GeminiPack> packs = Gemini.GeminiSendToFulfillment(pack);
		}
	}
}

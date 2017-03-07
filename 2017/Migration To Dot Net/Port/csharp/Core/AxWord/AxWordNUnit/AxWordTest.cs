using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.AxWord;

namespace Vertex.Fsd.Omiga.AxWord.Tests
{
	[TestFixture]
	public class AxWordTest
	{
		[Test]
		public void TestDisplayRtfZlib()
		{
			TestDisplay("axword.kfi.rtf.zlib.xml", ".rtf", "ZLIB", 30);
		}

		[Test]
		public void TestDisplayPdfZlib()
		{
			TestDisplay("axword.kfi.pdf.zlib.xml", ".pdf", "ZLIB", 20);
		}

		[Test]
		public void TestDisplayDocZlib()
		{
			TestDisplay("axword.kfi.doc.zlib.xml", ".doc", "ZLIB", 10);
		}

		[Test]
		public void TestDisplayTifZlib()
		{
			TestDisplay("axword.kfi.tif.zlib.xml", ".tiff", "ZLIB", 50);
		}

		[Test]
		public void TestDisplayAfpTif()
		{
			TestDisplay("axword.kfi.afp.xml", ".afp", "", 0);
		}

		private void TestDisplay(string fileName, string fileExtension, string compressionMethod, int deliveryType)
		{
			fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar + fileName;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load(fileName);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("RESPONSE/PRINTDOCUMENTDETAILS/@PRINTDOCUMENT");
			string request =
				"<REQUEST OPERATION='' FILEEXTENSION='" + fileExtension + "'>" +
					"<DOCUMENTCONTENTS FILECONTENTS='" + xmlNode.Value + "'>" +
					"</DOCUMENTCONTENTS>" +
					"<CONTROLDATA " +
						"COMPRESSIONMETHOD='" + compressionMethod + "' " +
						"DOCUMENTID='100' " +
						"DOCUMENTTITLE='KFI' " +
						"PRINTER='' " +
						"FIRSTPAGEPRINTERTRAY='' " +
						"OTHERPAGESPRINTERTRAY='' " +
						"COPIES='' " +
						"DELIVERYTYPE='" + deliveryType.ToString() + "'>" +
					"</CONTROLDATA>" +
				"</REQUEST>";
			AxWordClass axWordClass = new AxWordClass();
			axWordClass.EnableLog = true;
			axWordClass.ShowProgressBar = false;
			string response = axWordClass.DisplayDocumentNative(request);
		}


		[Test]
		public void TestPrintRtfZlib()
		{
			TestPrint("axword.kfi.rtf.zlib.xml", ".rtf", "ZLIB", 20);
		}

		[Test]
		public void TestPrintPdfZlib()
		{
			TestPrint("axword.kfi.pdf.zlib.xml", ".pdf", "ZLIB", 20);
		}

		private void TestPrint(string fileName, string fileExtension, string compressionMethod, int deliveryType)
		{
			fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar + fileName;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load(fileName);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("RESPONSE/PRINTDOCUMENTDETAILS/@PRINTDOCUMENT");
			string request =
				"<REQUEST OPERATION='' FILEEXTENSION='" + fileExtension + "'>" +
					"<DOCUMENTCONTENTS FILECONTENTS='" + xmlNode.Value + "'>" +
					"</DOCUMENTCONTENTS>" +
					"<CONTROLDATA " +
						"COMPRESSIONMETHOD='" + compressionMethod + "' " +
						"DOCUMENTID='100' " +
						"DOCUMENTTITLE='KFI' " +
						"PRINTER='\\\\printserv1\\JH_G_HPLJ4000_3' " +
						"FIRSTPAGEPRINTERTRAY='' " +
						"OTHERPAGESPRINTERTRAY='' " +
						"COPIES='' " +
						"DELIVERYTYPE='" + deliveryType.ToString() + "'>" +
					"</CONTROLDATA>" +
				"</REQUEST>";
			AxWordClass axWordClass = new AxWordClass();
			axWordClass.EnableLog = false;
			axWordClass.ShowPrintDialog = true;
			axWordClass.ShowProgressBar = true;
			axWordClass.DisablePrintOut = true;
			axWordClass.PersistState = true;
			bool printed = axWordClass.PrintDocument(request);

			//Assert.IsTrue(printed);
		}

		[Test]
		public void TestGetPrinters()
		{
			PrinterCollection printerCollection = new PrinterCollection();
		}

		[Test]
		public void TestGetPrintersAsXml()
		{
			AxWordClass axWordClass = new AxWordClass();
			string printers = axWordClass.GetPrintersAsXML();
		}

		[Test]
		public void TestConvertFile()
		{
			AxWordClass axWordClass = new AxWordClass();
			string binBase64 = axWordClass.ConvertFileToBase64(@"C:\DOCS\TAD.DOC", "ZLIB", false);
			axWordClass.ConvertBase64ToFile(binBase64, @"C:\DOCS\TAD.NEW.DOC", "ZLIB");
		}
	}
}

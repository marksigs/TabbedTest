using System;
using System.Collections.Generic;
using System.Text;
using Vertex.Fsd.Omiga.AxWord.Tests;

namespace AxWordTestHarness
{
	class Program
	{
		static void Main(string[] args)
		{
			AxWordTest axWordTest = new AxWordTest();
			//axWordTest.TestDisplayRtfZlib();
			//axWordTest.TestDisplayPdfZlib();
			//axWordTest.TestDisplayDocZlib();
			//axWordTest.TestDisplayTifZlib();
			//axWordTest.TestDisplayAfpTif();

			axWordTest.TestPrintRtfZlib();
			//axWordTest.TestPrintPdfZlib();

			//axWordTest.TestGetPrinters();
			//axWordTest.TestGetPrintersAsXml();
			//axWordTest.TestConvertFile();
		}
	}
}

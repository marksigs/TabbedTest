using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class DocumentPdf : DocumentBinary
	{
		public DocumentPdf()
		{
		}

		public DocumentPdf(string fileName) : base(fileName)
		{
		}

		public override bool Print(Form frmParent, bool disablePrintOut, PrintJob printJob)
		{
			bool success = false;

			if (string.IsNullOrEmpty(FileName))
			{
				throw new InvalidOperationException("Undefined FileName property.");
			}
			if (printJob == null)
			{
				throw new ArgumentNullException("printJob");
			}

			Logger.WriteLine("->DocumentPdf.Print('" + FileName + "', " + disablePrintOut.ToString() + ")");

			int pagesPrinted = PrintPdfFile(FileName, printJob, disablePrintOut);
			success = pagesPrinted > 0;

			Logger.WriteLine("<-DocumentPdf.Print() = " + success.ToString());

			return success;
		}

		private static int PrintPdfFile(string fileName, PrintJob printJob, bool disablePrintOut)
		{
			Logger.WriteLine("->DocumentPdf.PrintPdfFile('" + fileName + "')");

			if (string.IsNullOrEmpty(fileName))
			{
				throw new ArgumentNullException(fileName);
			}

			int pagesPrinted = 0;
			const string licenceName = "Marlborough-Stirling";
			const string licenceKey = "GAjAd9yU9dk9GA9kyRe";
			const int licenceCode = 5384401;
			const string optionsSeparator = "\r\n";

			string options = "";
			string printerName = "";
			int copies = 0;
			if (!string.IsNullOrEmpty(printJob.Printer.Name))
			{
				// AS 01/02/06 CORE234 pdfPrint treats spaces (or commas) in the printer name as delimiters,
				// unless the option is double quoted.
				printerName = printJob.Printer.Name;
				options = "'PRINTER=" + printerName + "'";
			}

			if (options.Length > 0)
			{
				options += optionsSeparator;
			}
			copies = printJob.Copies;
			options += "COPIES=" + Convert.ToString(copies);

			if (printJob.FirstPagePrinterBin.Number != MicrosoftWordConstants.wdPrinterDefaultBin)
			{
				if (options.Length > 0)
				{
					options += optionsSeparator;
				}
				options += "TRAY=" + Convert.ToString(printJob.FirstPagePrinterBin.Number);
			}

			if (printJob.OtherPagesPrinterBin.Number != MicrosoftWordConstants.wdPrinterDefaultBin)
			{
				if (options.Length > 0)
				{
					options += optionsSeparator;
				}
				options += "TRAY2=" + Convert.ToString(printJob.OtherPagesPrinterBin.Number);
			}

			if (options.Length > 0)
			{
				options += optionsSeparator;
			}
			options += "QUIET=0";
			if (options.Length > 0)
			{
				options += optionsSeparator;
			}
			options += "LISTPRINTER=1";
			if (options.Length > 0)
			{
				options += optionsSeparator;
			}
			options += "LISTTRAY=1";

			if (disablePrintOut)
			{
				pagesPrinted = 1;
			}
			else
			{
				string password = null;
				System.Diagnostics.Debug.WriteLine("->pdfPrint(" + fileName + ", " + printerName + ", Copies: " + copies.ToString() + ", " + options + ")");
				Logger.WriteLine("->pdfPrint(" + fileName + ", " + printerName + ", Copies: " + copies.ToString() + ", " + options + ")");
				pagesPrinted = NativeMethods.PdfPrint(fileName, password, licenceName, licenceKey, licenceCode, options);
				System.Diagnostics.Debug.WriteLine("<-pdfPrint(" + fileName + ", " + printerName + ", Copies: " + copies.ToString() + ", " + options + ") = " + pagesPrinted.ToString());
				Logger.WriteLine("<-pdfPrint(" + fileName + ", " + printerName + ", Copies: " + copies.ToString() + ", " + options + ") = " + pagesPrinted.ToString());
				if (pagesPrinted == 0)
				{
					throw new InvalidOperationException("Error printing " + fileName + " - pdfPrint returned 0");
				}
			}

			Logger.WriteLine("<-DocumentPdf.PrintPdfFile() = " + pagesPrinted.ToString());

			return pagesPrinted;
		}
	}
}

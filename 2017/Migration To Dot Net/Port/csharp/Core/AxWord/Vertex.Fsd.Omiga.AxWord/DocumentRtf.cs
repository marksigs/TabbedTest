using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class DocumentRtf : DocumentText
	{
		protected override Encoding GetEncoding(byte[] bytes)
		{
			Encoding encoding = null;

			const string rtfTag = @"{\rtf";
			if (bytes.Length > rtfTag.Length && string.Compare(Encoding.UTF8.GetString(bytes, 0, rtfTag.Length), rtfTag, true) == 0)
			{
				encoding = Encoding.UTF8;
			}
			else if (bytes.Length > rtfTag.Length * 2 && string.Compare(Encoding.Unicode.GetString(bytes, 0, rtfTag.Length * 2), rtfTag, true) == 0)
			{
				encoding = Encoding.Unicode;
			}
			else
			{
				throw new InvalidOperationException("Invalid RTF format '" + Encoding.UTF8.GetString(bytes, 0, rtfTag.Length * 2) + "'.");
			}

			return encoding;
		}

		public override bool Print(Form frmParent, bool disablePrintOut, PrintJob printJob)
		{
			DocumentPdf documentPdf = ToPdf();
			return documentPdf.Print(frmParent, disablePrintOut, printJob);
		}

		private DocumentPdf ToPdf()
		{
			DocumentPdf documentPdf = new DocumentPdf(FileName + ".pdf");

			Logger.WriteLine("->ToPdf('" + FileName + "', '" + documentPdf.FileName + "')");

			wPDF_X01.PDFControl pdfControl = null;
			try
			{
				pdfControl = new wPDF_X01.PDFControl();
				if (pdfControl != null)
				{
					const string password = "";

					pdfControl.PDF_Compression = wPDF_X01.TxPDFCompression.deflate;
					pdfControl.PDF_FontMode = wPDF_X01.TxPDFFontMode.UseTrueTypeFonts;
					pdfControl.SECURITY_Permission = 0;
					pdfControl.SECURITY_OwnerPassword = "R3NrstsEs4EdeC7rtff4";
					pdfControl.SECURITY_UserPassword = password;
					pdfControl.INFO_Author = "Marlborough-Stirling";
					pdfControl.INFO_Date = DateTime.Now.ToString("dd MMMM yyyy at HH:mm:ss");

					string wRtf2Pdf02Path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar + "wRTF2PDF02.dll";
					if (pdfControl.StartEngine(wRtf2Pdf02Path, "Marlborough Stirling plc", "R3NrstsEs4EdeC7rtff4", 16575595))
					{
						int result = pdfControl.BeginDoc(documentPdf.FileName, 0);
						if (result > 0)
						{
							pdfControl.ExecIntCommand(1024, 1);	// Use printer to determine sizes							
							pdfControl.ExecIntCommand(1000, 0);
							pdfControl.ExecStrCommand(1002, FileName);
							pdfControl.ExecIntCommand(1100, 0);
							pdfControl.EndDoc();
						}
						else
						{
							throw new InvalidOperationException("Error converting RTF to PDF - pdfControl.BeginDoc returned " + Convert.ToString(result));
						}

						pdfControl.StopEngine();
					}
					else
					{
						throw new InvalidOperationException("Error printing RTF - unable to load " + wRtf2Pdf02Path);
					}
				}
				else
				{
					throw new InvalidOperationException("Error printing RTF - unable to create wPDF_X01.PDFControl");
				}
			}
			finally
			{
				if (pdfControl != null) Marshal.ReleaseComObject(pdfControl);
			}

			Logger.WriteLine("<-ToPdf('" + FileName + "', '" + documentPdf.FileName + "')");

			return documentPdf;
		}

	}
}

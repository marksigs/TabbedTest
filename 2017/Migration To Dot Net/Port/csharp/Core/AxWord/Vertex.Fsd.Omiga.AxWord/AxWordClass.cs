using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Vertex.Fsd.Omiga.AxWord
{
	[ComVisible(true)]
	[ProgId("axword.axwordclass")]
	[Guid("68575265-E41C-4C2E-808A-CBA63B53D0EF")]
	public class AxWordClass
	{
		private AxWordManager _axWordManager = new AxWordManager();

		public bool ReadOnly
		{
			get { return _axWordManager.ReadOnly; }
			set { _axWordManager.ReadOnly = value; }
		}

		public bool PersistState
		{
			get { return _axWordManager.PersistState; }
			set { _axWordManager.PersistState = value; }
		}

		public bool ShowPrintDialog
		{
			get { return _axWordManager.ShowPrintDialog; }
			set { _axWordManager.ShowPrintDialog = value; }
		}

		public bool ShowProgressBar
		{
			get { return _axWordManager.ShowProgressBar; }
			set { _axWordManager.ShowProgressBar = value; }
		}

		public bool DisablePrintOut
		{
			get { return _axWordManager.DisablePrintOut; }
			set { _axWordManager.DisablePrintOut = value; }
		}

		public bool EnableLog
		{
			get { return _axWordManager.LogEnabled; }
			set { _axWordManager.LogEnabled = value; }
		}

		public string LogFileName
		{
			get { return _axWordManager.LogFileName; }
			set { _axWordManager.LogFileName = value; }
		}

		public bool DocumentEdited
		{
			get { return _axWordManager.DocumentEdited; }
		}

		public bool DocumentPrinted
		{
			get { return _axWordManager.DocumentPrinted; }
		}

		public bool Modeless
		{
			get { return _axWordManager.Modeless; }
			set { _axWordManager.Modeless = value; }
		}

		public string DisplayDocumentNative(string xmlTextIn)
		{
			return _axWordManager.DisplayDocument(new AxWordRequest(xmlTextIn, _axWordManager.PersistState));
		}

		public bool PrintDocument(string xmlTextIn)
		{
			return _axWordManager.PrintDocument(new AxWordRequest(xmlTextIn, _axWordManager.PersistState));
		}

		public string GetPrintersAsXML()
		{
			PrinterCollection printerCollection = new PrinterCollection();
			return printerCollection.ToXmlString();
		}

		public string ConvertFileToBase64(string fileName, string compressionMethod, bool deleteFile)
		{
			return _axWordManager.ConvertFileToBase64(fileName, compressionMethod, deleteFile);
		}

		public bool ConvertBase64ToFile(string binBase64, string fileName, string compressionMethod)
		{
			return _axWordManager.ConvertBase64ToFile(binBase64, fileName, compressionMethod);
		}

	}
}

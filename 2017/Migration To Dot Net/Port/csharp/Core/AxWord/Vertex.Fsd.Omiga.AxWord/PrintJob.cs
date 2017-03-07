using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Microsoft.Win32;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class PrintJob
	{
		#region Fields and properties
		private const string _companyName = "Vertex Financial Services";
		private const string _applicationName = "AxWord";
		private const string _sectionName = "PrintJob";

		private bool _persistState;
		public bool PersistState
		{
			get { return _persistState; }
			set { _persistState = value; }
		}

		private Document _document;
		public Document Document
		{
			get { return _document; }
			set { _document = value; }
		}

		private Printer _printer = new Printer();
		public Printer Printer
		{
			get { return _printer; }
			set { _printer = value; }
		}

		private PrinterBin _firstPagePrinterBin = new PrinterBin();
		public PrinterBin FirstPagePrinterBin
		{
			get { return _firstPagePrinterBin; }
			set { _firstPagePrinterBin = value; }
		}

		private PrinterBin _otherPagesPrinterBin = new PrinterBin();
		public PrinterBin OtherPagesPrinterBin
		{
			get { return _otherPagesPrinterBin; }
			set { _otherPagesPrinterBin = value; }
		}

		private int _copies = 1;
		public int Copies
		{
			get { return _copies; }
			set { _copies = value; }
		}

		private bool _useDifferentBinForOtherPages;
		public bool UseDifferentBinForOtherPages
		{
			get { return _useDifferentBinForOtherPages; }
			set { _useDifferentBinForOtherPages = value; }
		}
		#endregion

		protected PrintJob()
		{
		}

		public PrintJob(XmlDocument xmlDocument, Document document, bool persistState)
		{
			_document = document;
			_persistState = persistState;

			XmlNode xmlControlDataNode = xmlDocument.SelectSingleNode("REQUEST/CONTROLDATA");
			if (xmlControlDataNode != null)
			{
				XmlNode xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("PRINTER");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					_printer = new Printer(xmlAttribute.InnerText);
				}

				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("FIRSTPAGEPRINTERTRAY");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					_firstPagePrinterBin = new PrinterBin(Convert.ToInt32(xmlAttribute.InnerText));
				}

				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("OTHERPAGESPRINTERTRAY");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					_otherPagesPrinterBin = new PrinterBin(Convert.ToInt32(xmlAttribute.InnerText));
				}

				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("COPIES");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					_copies = Convert.ToInt32(xmlAttribute.InnerText);
				}
			}

			int copies = _copies;

			if (_persistState)
			{
				// Load request defaults from registry for this document. This will override any
				// equivalent settings in the input request.
				LoadState();
			}

			// AS 12/05/2005 CORE124 Do not override COPIES in request.
			if (copies > -1)
			{
				_copies = copies;
			}
		}

		public void LoadState()
		{
			if (!string.IsNullOrEmpty(_document.DocumentId))
			{
				// Load default settings for this document.
				LoadStateSection(_sectionName + "\\" + _document.DocumentId);
				if (_printer == null || string.IsNullOrEmpty(_printer.Name))
				{
					// No entries for this document, so load default settings for all documents.
					LoadStateSection(_sectionName);
				}
			}
			else
			{
				LoadStateSection(_sectionName);
			}
		}

		private void LoadStateSection(string sectionName)
		{
			using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("Software\\" + _companyName + "\\" + _applicationName + "\\" + sectionName))
			{
				if (registryKey != null)
				{
					_printer.Name = (string)registryKey.GetValue("DeviceName", _printer.Name);
					_copies = (int)registryKey.GetValue("Copies", _copies);
					_firstPagePrinterBin.Number = (int)registryKey.GetValue("FirstPageBin", _firstPagePrinterBin.Number);
					_otherPagesPrinterBin.Number = (int)registryKey.GetValue("OtherPagesBin", _otherPagesPrinterBin.Number);
					_useDifferentBinForOtherPages = (int)registryKey.GetValue("UseDifferentBinForOtherPages", _useDifferentBinForOtherPages) == 1;
				}
			}
		}

		public void SaveState()
		{
			string sectionName = _sectionName;

			if (!string.IsNullOrEmpty(_document.DocumentId))
			{
				// Save default settings for this document.
				sectionName += "\\" + _document.DocumentId;
			}

			SaveStateSection(sectionName);
		}

		private void SaveStateSection(string sectionName)
		{
			using (RegistryKey registryKey = Registry.LocalMachine.CreateSubKey("Software\\" + _companyName + "\\" + _applicationName + "\\" + sectionName))
			{
				if (!string.IsNullOrEmpty(_printer.Name)) registryKey.SetValue("DeviceName", _printer.Name, RegistryValueKind.String);
				registryKey.SetValue("Copies", _copies, RegistryValueKind.DWord);
				registryKey.SetValue("FirstPageBin", _firstPagePrinterBin.Number, RegistryValueKind.DWord);
				registryKey.SetValue("OtherPagesBin", _otherPagesPrinterBin.Number, RegistryValueKind.DWord);
				registryKey.SetValue("UseDifferentBinForOtherPages", _useDifferentBinForOtherPages ? 1 : 0, RegistryValueKind.DWord);
			}
		}
	}
}

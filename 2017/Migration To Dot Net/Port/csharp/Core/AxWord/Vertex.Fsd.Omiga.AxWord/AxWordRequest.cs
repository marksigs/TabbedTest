using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class AxWordRequest
	{
		private bool _persistState;

		private Document _document;
		public Document Document
		{
			get { return _document; }
			set { _document = value; }
		}

		private PrintJob _printJob;
		public PrintJob PrintJob
		{
			get { return _printJob; }
			set { _printJob = value; }
		}

		private string _operation;
		public string Operation
		{
			get { return _operation; }
			set { _operation = value; }
		}

		public AxWordRequest(string xmlText, bool persistState)
		{
			_persistState = persistState;

			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(xmlText);

			XmlNode xmlAttribute = xmlDocument.DocumentElement.Attributes.GetNamedItem("OPERATION");
			if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
			{
				_operation = xmlAttribute.InnerText.ToUpper();
			}

			_document = Document.CreateDocument(xmlDocument);
			_printJob = new PrintJob(xmlDocument, _document, _persistState);

			XmlNode xmlControlDataNode = xmlDocument.SelectSingleNode("REQUEST/CONTROLDATA");
		}
	}
}

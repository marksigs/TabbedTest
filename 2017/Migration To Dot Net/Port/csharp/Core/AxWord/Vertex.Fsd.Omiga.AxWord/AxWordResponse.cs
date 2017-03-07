using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class AxWordResponse
	{
		private bool _persistState;

		private Document _document;
		public Document Document
		{
			get { return _document; }
			set { _document = value; }
		}

		private Exception _exception;
		public Exception Exception
		{
			get { return _exception; }
			set { _exception = value; }
		}

		public AxWordResponse(Document document, Exception exception, bool persistState)
		{
			_document = document;
			_exception = exception;
			_persistState = persistState;
		}

		public override string ToString()
		{
			string xmlString = null;

			using (StringWriter stringWriter = new StringWriter())
			{
				XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
				xmlWriterSettings.OmitXmlDeclaration = true;
				using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, xmlWriterSettings))
				{
					xmlWriter.WriteStartDocument();
					xmlWriter.WriteStartElement("RESPONSE");
					xmlWriter.WriteAttributeString("TYPE", _exception == null ? "SUCCESS" : "APPERR");
					if (_exception == null)
					{
						xmlWriter.WriteStartElement("DOCUMENTCONTENTS");
						xmlWriter.WriteAttributeString("FILECONTENTS", _document.FileContentsBinBase64);
						xmlWriter.WriteAttributeString("COMPRESSIONMETHOD", Document.ToString(_document.DocumentCompressionMethod));
						xmlWriter.WriteEndElement();
					}
					else
					{
						string id = Guid.NewGuid().ToString("N");
						xmlWriter.WriteStartElement("ERROR");
						xmlWriter.WriteElementString("EXCEPTIONTYPE", _exception.GetType().ToString());
						xmlWriter.WriteElementString("ID", id);
						xmlWriter.WriteElementString("DESCRIPTION", _exception.Message);
						xmlWriter.WriteElementString("VERSION", Assembly.GetExecutingAssembly().FullName);
						xmlWriter.WriteElementString("SOURCE", _exception.StackTrace);
						xmlWriter.WriteEndElement();
					}
					xmlWriter.WriteEndElement();
					xmlWriter.WriteEndDocument();
					xmlWriter.Flush();
				}

				xmlString = stringWriter.ToString();
			}

			return xmlString;
		}

	}
}

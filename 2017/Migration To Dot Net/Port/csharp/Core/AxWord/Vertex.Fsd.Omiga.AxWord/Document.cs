using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public abstract class Document : IDisposable
	{
		#region Fields and properties

		private static Logger _logger = Logger.GetLogger();
		protected static Logger Logger
		{
			get { return _logger; }
			set { _logger = value; }
		}

		private bool _disposed;
#if DEBUG
		private bool _deleteFilesOnDisposal = true;
#else
		private bool _deleteFilesOnDisposal = true;
#endif
		public bool DeleteFilesOnDisposal
		{
			get { return _deleteFilesOnDisposal; }
			set { _deleteFilesOnDisposal = value; }
		}

		private StringCollection _disposalFiles = new StringCollection();
		public StringCollection DisposalFiles
		{
			get { return _disposalFiles; }
		}

		private string _fileName;
		public string FileName
		{
			get { return _fileName; }
			set { _fileName = value; }
		}

		private string _fileExtension;
		public string FileExtension
		{
			get { return _fileExtension; }
			set { _fileExtension = value; }
		}

		private string _documentId;
		public string DocumentId
		{
			get { return _documentId; }
			set { _documentId = value; }
		}

		private string _documentTitle;
		public string DocumentTitle
		{
			get { return _documentTitle; }
			set { _documentTitle = value; }
		}

		private DocumentDeliveryType _deliveryType = DocumentDeliveryType.Word;
		public DocumentDeliveryType DeliveryType
		{
			get { return _deliveryType; }
			set { _deliveryType = value; }
		}

		private string _fileContentsBinBase64;
		public string FileContentsBinBase64
		{
			get { return _fileContentsBinBase64; }
			set { _fileContentsBinBase64 = value; }
		}

		private DocumentCompressionMethod _compressionMethod = DocumentCompressionMethod.Unknown;
		public DocumentCompressionMethod DocumentCompressionMethod
		{
			get { return _compressionMethod; }
			set { _compressionMethod = value; }
		}

		private bool _compressed;
		public bool Compressed
		{
			get { return _compressed; }
			set { _compressed = value; }
		}
		#endregion

		#region Constructors
		protected Document()
		{
		}

		protected Document(string fileName)
		{
			_fileName = fileName;
			_disposalFiles.Add(fileName);
		}

		~Document()
		{
			Dispose(false);
		}

		public static Document CreateDocument(string fileName, DocumentCompressionMethod compressionMethod)
		{
			return CreateDocument(fileName, compressionMethod, null);
		}

		public static Document CreateDocument(string fileName, DocumentCompressionMethod compressionMethod, string binBase64)
		{
			if (string.IsNullOrEmpty(fileName))
			{
				throw new ArgumentNullException("fileName");
			}

			string fileExtension = Path.GetExtension(fileName).ToLower();
			DocumentDeliveryType deliveryType = ToDocumentDeliveryType(fileExtension);
			Document document = CreateDocument(deliveryType);
			document._fileName = fileName;
			document._fileExtension = fileExtension;
			document._compressionMethod = compressionMethod;
			document._fileContentsBinBase64 = binBase64;

			return document;
		}

		public static DocumentCompressionMethod ToDocumentCompressionMethod(string compressionMethodText)
		{
			DocumentCompressionMethod compressionMethod = DocumentCompressionMethod.Unknown;

			if (compressionMethodText == null)
			{
				compressionMethod = DocumentCompressionMethod.Unknown;
			}
			else if (compressionMethodText.Length == 0)
			{
				compressionMethod = DocumentCompressionMethod.None;
			}
			else if (string.Compare(compressionMethodText, "zlib", true) == 0)
			{
				compressionMethod = DocumentCompressionMethod.Zlib;
			}
			else if (string.Compare(compressionMethodText, "compapi", true) == 0)
			{
				compressionMethod = DocumentCompressionMethod.Compapi;
			}

			return compressionMethod;
		}

		public static DocumentDeliveryType ToDocumentDeliveryType(string fileExtension)
		{
			DocumentDeliveryType deliveryType = DocumentDeliveryType.Unknown;

			if (fileExtension == null)
			{
				deliveryType = DocumentDeliveryType.Unknown;
			}
			else if (string.Compare(fileExtension, ".doc", true) == 0)
			{
				deliveryType = DocumentDeliveryType.Word;
			}
			else if (string.Compare(fileExtension, ".pdf", true) == 0)
			{
				deliveryType = DocumentDeliveryType.Pdf;
			}
			else if (string.Compare(fileExtension, ".rtf", true) == 0)
			{
				deliveryType = DocumentDeliveryType.Rtf;
			}
			else if (string.Compare(fileExtension, ".tiff", true) == 0)
			{
				deliveryType = DocumentDeliveryType.Tif;
			}
			else
			{
				deliveryType = DocumentDeliveryType.Unknown;
			}

			return deliveryType;
		}

		private static Document CreateDocument(DocumentDeliveryType deliveryType)
		{
			Document document = null;

			switch (deliveryType)
			{
				case DocumentDeliveryType.Word:
					document = new DocumentWord();
					break;
				case DocumentDeliveryType.Pdf:
					document = new DocumentPdf();
					break;
				case DocumentDeliveryType.Rtf:
					document = new DocumentRtf();
					break;
				case DocumentDeliveryType.Tif:
					document = new DocumentTif();
					break;
				case DocumentDeliveryType.Unknown:
					document = new DocumentBinary();
					break;
				default:
					throw new InvalidOperationException("Invalid delivery type: " + deliveryType.ToString() + ".");
			}

			document._deliveryType = deliveryType;

			return document;
		}

		public static Document CreateDocument(XmlDocument xmlDocument)
		{
			if (xmlDocument == null)
			{
				throw new ArgumentNullException("xmlDocument");
			}

			Document document = null;
			DocumentDeliveryType deliveryType = DocumentDeliveryType.Unknown;
			string fileExtension = null;

			// Check XML itself - the structure of the document.
			XmlNode xmlDocContentsNode = xmlDocument.SelectSingleNode("REQUEST/DOCUMENTCONTENTS");
			if (xmlDocContentsNode == null)
			{
				xmlDocContentsNode = xmlDocument.SelectSingleNode("REQUEST/PRINTDOCUMENTDATA");
				if (xmlDocContentsNode == null)
				{
					throw new ArgumentException("Invalid XML request - missing mandatory node: REQUEST/DOCUMENTCONTENTS");
				}
			}

			XmlNode xmlAttribute = null;
			XmlNode xmlControlDataNode = xmlDocument.SelectSingleNode("REQUEST/CONTROLDATA");
			if (xmlControlDataNode != null)
			{
				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("DELIVERYTYPE");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					deliveryType = ParseDeliveryType(xmlAttribute.InnerText);
				}
			}

			xmlAttribute = xmlDocument.DocumentElement.Attributes.GetNamedItem("FILEEXTENSION");
			if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
			{
				fileExtension = xmlAttribute.InnerText;
				if (string.IsNullOrEmpty(fileExtension))
				{
					switch (deliveryType)
					{
						case DocumentDeliveryType.Word:
							fileExtension = ".doc";
							break;
						case DocumentDeliveryType.Pdf:
							fileExtension = ".pdf";
							break;
						case DocumentDeliveryType.Rtf:
							fileExtension = ".rtf";
							break;
						case DocumentDeliveryType.Tif:
							fileExtension = ".tiff";
							break;
						default:
							fileExtension = ".doc";
							break;
					}
				}
				else
				{
					deliveryType = ToDocumentDeliveryType(fileExtension);
				}
			}

			document = CreateDocument(deliveryType);
			document._fileExtension = fileExtension;

			// Check XML itself - attribute that holds the file contents.
			xmlAttribute = xmlDocContentsNode.Attributes.GetNamedItem("FILECONTENTS");
			if (xmlAttribute == null)
			{
				throw new ArgumentException("Invalid XML request - missing mandatory attribute: REQUEST/DOCUMENTCONTENTS/@FILECONTENTS");
			}
			// Now get the XML data.
			document._fileContentsBinBase64 = xmlAttribute.InnerText;

			if (xmlControlDataNode != null)
			{
				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("COMPRESSIONMETHOD");
				if (xmlAttribute != null)
				{
					document._compressionMethod = ParseCompressionMethod(xmlAttribute.InnerText);
					document._compressed = document._compressionMethod != DocumentCompressionMethod.Unknown && document._compressionMethod != DocumentCompressionMethod.None;
				}

				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("DOCUMENTID");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					document._documentId = xmlAttribute.InnerText;
				}

				xmlAttribute = xmlControlDataNode.Attributes.GetNamedItem("DOCUMENTTITLE");
				if (xmlAttribute != null && xmlAttribute.InnerText.Length > 0)
				{
					document._documentTitle = xmlAttribute.InnerText;
				}
			}

			return document;
		}
		#endregion

		protected abstract void WriteToFile(FileStream fileStream, byte[] bytes);
		protected abstract byte[] ReadFromFile(FileStream fileStream);

		public string ToFile()
		{
			if (string.IsNullOrEmpty(_fileContentsBinBase64))
			{
				throw new InvalidOperationException("Undefined FileContentsBinBase64 property.");
			}

			if (string.IsNullOrEmpty(_fileName))
			{
				_fileName = Path.GetTempPath() + GenerateFilename(_documentTitle) + _fileExtension;
			}

			byte[] bytes = Convert.FromBase64String(_fileContentsBinBase64);

			_compressed = Decompress(ref bytes, _compressionMethod);

			using (FileStream fileStream = new FileStream(_fileName, FileMode.Create, FileAccess.Write, FileShare.None))
			{
				WriteToFile(fileStream, bytes);
			}

			// AS 03/10/2006 EP1194 Basic check to ensure file is valid.
			FileInfo fileInfo = new FileInfo(_fileName);
			if (fileInfo.Length == 0)
			{
				throw new InvalidOperationException("Invalid file length: " + _fileName);
			}

			_disposalFiles.Add(_fileName);

			return _fileName;
		}

		public string FromFile()
		{
			if (string.IsNullOrEmpty(_fileName))
			{
				throw new InvalidOperationException("Undefined FileName property.");
			}

			byte[] bytes = null;
			using (FileStream fileStream = new FileStream(_fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				bytes = ReadFromFile(fileStream);
			}

			_compressed = Compress(ref bytes, _compressionMethod);

			_fileContentsBinBase64 = Convert.ToBase64String(bytes);

			return _fileContentsBinBase64;
		}

		private static string GenerateFilename(string prefix)
		{
			string filename = null;

			Guid guid = Guid.NewGuid();

			if (!string.IsNullOrEmpty(prefix))
			{
				filename = prefix + "." + guid.ToString("N");
			}
			else
			{
				filename = guid.ToString("N");
			}

			return filename;
		}

		public virtual bool Print(Form frmParent, bool disablePrintOut, PrintJob printJob)
		{
			throw new NotImplementedException("Printing of " + _fileName + " is not supported.");
		}

		#region Compression
		public static bool Compress(ref byte[] bytes, DocumentCompressionMethod compressionMethod)
		{
			bool compressed = false;

			if (compressionMethod != DocumentCompressionMethod.Unknown && compressionMethod != DocumentCompressionMethod.None)
			{
				dmsCompression.dmsCompression1Class dmsCompression = null;
				try
				{
					dmsCompression = new dmsCompression.dmsCompression1Class();

					dmsCompression.CompressionMethod = ToString(compressionMethod);
					bytes = (byte[])dmsCompression.SafeArrayCompressToSafeArray(bytes);
					compressed = true;
				}
				catch (Exception exception)
				{
					throw new InvalidOperationException("Error compressing data.", exception);
				}
				finally
				{
					if (dmsCompression != null) Marshal.ReleaseComObject(dmsCompression);
				}
			}

			return compressed;
		}

		public static bool Decompress(ref byte[] bytes, DocumentCompressionMethod compressionMethod)
		{
			bool compressed = false;

			if (compressionMethod != DocumentCompressionMethod.Unknown && compressionMethod != DocumentCompressionMethod.None)
			{
				compressed = true;

				dmsCompression.dmsCompression1Class dmsCompression = null;
				try
				{
					dmsCompression = new dmsCompression.dmsCompression1Class();

					dmsCompression.CompressionMethod = ToString(compressionMethod);
					bytes = (byte[])dmsCompression.SafeArrayDecompressToSafeArray(bytes);
					compressed = false;
				}
				catch (Exception exception)
				{
					throw new InvalidOperationException("Error decompressing data.", exception);
				}
				finally
				{
					if (dmsCompression != null) Marshal.ReleaseComObject(dmsCompression);
				}
			}

			return compressed;
		}
		#endregion

		#region Parsing

		internal static DocumentDeliveryType ParseDeliveryType(int deliveryType)
		{
			return ParseDeliveryType(deliveryType.ToString());
		}

		internal static DocumentDeliveryType ParseDeliveryType(string deliveryTypeText)
		{
			DocumentDeliveryType deliveryType;

			if (deliveryTypeText == null)
			{
				throw new ArgumentNullException("deliveryTypeText");
			}

			string deliveryTypeTextLower = deliveryTypeText.ToLower();
			switch (deliveryTypeTextLower)
			{
				case "0":
					deliveryType = DocumentDeliveryType.Unknown;
					break;
				case "10":
				case "doc":
					deliveryType = DocumentDeliveryType.Word;
					break;
				case "20":
				case "pdf":
					deliveryType = DocumentDeliveryType.Pdf;
					break;
				case "30":
				case "rtf":
					deliveryType = DocumentDeliveryType.Rtf;
					break;
				case "40":
				case "xml":
					deliveryType = DocumentDeliveryType.Xml;
					break;
				case "50":
				case "tif":
				case "tiff":
					deliveryType = DocumentDeliveryType.Tif;
					break;
				default:
					throw new ArgumentException("Invalid DocumentDeliveryType: " + deliveryTypeText + ".", "deliveryTypeText");
			}

			return deliveryType;
		}

		internal static DocumentCompressionMethod ParseCompressionMethod(string compressionMethodText)
		{
			return ParseCompressionMethod(compressionMethodText, DocumentCompressionMethod.None);
		}

		internal static DocumentCompressionMethod ParseCompressionMethod(string compressionMethodText, DocumentCompressionMethod defaultCompressionMethod)
		{
			DocumentCompressionMethod compressionMethod = defaultCompressionMethod;

			if (compressionMethodText != null && compressionMethodText.Length > 0)
			{
				switch (compressionMethodText.ToUpper())
				{
					case "COMPAPI":
						compressionMethod = DocumentCompressionMethod.Compapi;
						break;
					case "ZLIB":
						compressionMethod = DocumentCompressionMethod.Zlib;
						break;
					default:
						throw new ArgumentException("Invalid CompressionMethod: " + compressionMethodText + ".", "compressionMethodText");
				}
			}
			else
			{
				// Default compression method.
				compressionMethod = defaultCompressionMethod;
			}

			return compressionMethod;
		}
		#endregion

		#region Conversion
		public static int ToInt32(DocumentDeliveryType deliveryType)
		{
			int deliveryTypeInt = 10;

			switch (deliveryType)
			{
				case DocumentDeliveryType.Unknown:
					deliveryTypeInt = 0;
					break;
				case DocumentDeliveryType.Word:
					deliveryTypeInt = 10;
					break;
				case DocumentDeliveryType.Pdf:
					deliveryTypeInt = 20;
					break;
				case DocumentDeliveryType.Rtf:
					deliveryTypeInt = 30;
					break;
				case DocumentDeliveryType.Tif:
					deliveryTypeInt = 50;
					break;
			}

			return deliveryTypeInt;
		}

		public static string ToString(DocumentCompressionMethod compressionMethod)
		{
			string compressionMethodText = null;

			switch (compressionMethod)
			{
				case DocumentCompressionMethod.Compapi:
					compressionMethodText = "COMPAPI";
					break;
				case DocumentCompressionMethod.Zlib:
					compressionMethodText = "ZLIB";
					break;
			}

			return compressionMethodText;
		}

		public static string ToString(DocumentDeliveryType deliveryType)
		{
			string deliveryTypeText = null;

			switch (deliveryType)
			{
				case DocumentDeliveryType.Unknown:
					deliveryTypeText = "";
					break;
				case DocumentDeliveryType.Word:
					deliveryTypeText = "doc";
					break;
				case DocumentDeliveryType.Pdf:
					deliveryTypeText = "pdf";
					break;
				case DocumentDeliveryType.Rtf:
					deliveryTypeText = "rtf";
					break;
				case DocumentDeliveryType.Tif:
					deliveryTypeText = "tiff";
					break;
			}

			return deliveryTypeText;
		}
		#endregion

		#region IDisposable
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					// Dispose managed resources.
				}

				// Dispose unmanaged resources.
				// Delete any temporary files associated with this document.
				if (_deleteFilesOnDisposal)
				{
					foreach (string fileName in _disposalFiles)
					{
						try
						{
							if (File.Exists(fileName))
							{
								if (_logger.Enabled)
								{
									// For debug purposes, do not delete temporary files, but rename instead.
									string deletedFileName = Path.GetDirectoryName(fileName) + "\\deleted." + Path.GetFileName(fileName);
									if (File.Exists(deletedFileName))
									{
										_logger.WriteLine("Deleting '" + deletedFileName + "'");
										File.SetAttributes(deletedFileName, FileAttributes.Normal);
										File.Delete(deletedFileName);
									}
									_logger.WriteLine("Moving '" + fileName + "' to '" + deletedFileName + "'");
									File.Move(fileName, deletedFileName);
								}
								else
								{
									File.SetAttributes(fileName, FileAttributes.Normal);
									File.Delete(fileName);
								}
							}
						}
						catch (Exception exception)
						{
							// Do not throw exception.
							if (_logger.Enabled) _logger.WriteLine(exception.ToString());
						}
					}
				}

				_disposalFiles.Clear();
			}

			_disposed = true;
		}
		#endregion IDisposable
	}
}

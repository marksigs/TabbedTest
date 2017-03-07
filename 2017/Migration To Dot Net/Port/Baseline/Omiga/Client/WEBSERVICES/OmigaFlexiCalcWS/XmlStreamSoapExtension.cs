//
//--------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/OmigaFlexiCalcWS/XmlStreamSoapExtension.cs $
//Workfile:			$Workfile: XmlStreamSoapExtension.cs $
//Current Version	$Revision: 1 $
//Last Modified		$Modtime: 16/03/07 13:37 $
//Modified By		$Author: Mheys $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//IK	28/02/2007	created
//--------------------------------------------------------------------------------------------
//
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Web.Services.Description;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.IO;
using System.Text;
using System.Reflection;
using Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.Web.Services.OmigaFlexiCalcWS
{
    // XmlStreamSoapExtension exposes raw SOAP messages to
    // an ASP.NET Web service
    public class XmlStreamSoapExtension : SoapExtension
    {
		private static readonly Logger _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		bool output = false;		// flag indicating input or output
        Stream httpOutputStream;    // HTTP output stream to send
                                    // real output to
        Stream chainedOutputStream; // output stream for ASP.NET
                                    // plumbing to write to
        Stream appOutputStream;     // output stream for method
                                    // to write to
		private Stream httpInputStream;		// replacement HTTP input stream 
		// to avoid de-serialisation
		private Stream chainedInputStream;	// copy of input stream

		StringReader soapPayload;

		private string GetXMLPath()
		{
			string thePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

			if (thePath.ToUpper().IndexOf("DLLDOTNET") >= 0)
			{
				thePath = thePath.Replace("\\DLLDOTNET", "\\XML");
			}
			else if (thePath.ToUpper().IndexOf("OMIGAWEBSERVICES") >= 0)
			{
				int position = thePath.ToUpper().IndexOf("OMIGAWEBSERVICES");
				thePath = thePath.Substring(0, position) + @"XML";
			}
			return(thePath);
		}

        // ChainStream replaces original stream with
        // extension’s stream
        public override Stream ChainStream(Stream stream)
        {
			const string soapPreamble = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><soap:Body>";
			const string soapPostamble = "</soap:Body></soap:Envelope>";
			const string emptyOmigaFlexiCalcRequest = soapPreamble + "<OmigaFlexiCalcRequest xmlns='http://Request.FlexiCalc.Omiga.vertex.co.uk'></OmigaFlexiCalcRequest>" + soapPostamble;
			
			Stream result = stream;
            // only replace output stream with memory stream
            if (output)
            {
                httpOutputStream = stream;
                result = chainedOutputStream = new MemoryStream();
            }
            else 
            {
				//	EP2_159
				httpInputStream = stream;

				chainedInputStream = new MemoryStream();
				StreamWriter sw = new StreamWriter(chainedInputStream);
				sw.Write(emptyOmigaFlexiCalcRequest);
				sw.Flush();

				chainedInputStream.Position = 0;
				result = chainedInputStream;
				// EP2_159_ends
				
				output = true;
            }
            return result;
        }
        
        public override object GetInitializer(System.Type serviceType)
        {
            return null;
        }

        public override object GetInitializer(System.Web.Services.Protocols.LogicalMethodInfo methodInfo, System.Web.Services.Protocols.SoapExtensionAttribute attribute)
        {
            return null;
        }

        public override void Initialize(object initializer)
        {
        
        }

        // ProcessMessage is called to process SOAP messages
        // after inbound messages are deserialized to input
        // parameters and output parameters are serialized to
        // outbound messages
        public override void ProcessMessage(SoapMessage message)
        {
            switch (message.Stage)
            {
				case SoapMessageStage.AfterDeserialize :
                {
					// rewind HTTP input stream to start and store
                    // reference in current HTTP context
                    HttpContext.Current.Request.InputStream.Position = 0;
                    HttpContext.Current.Items["SoapInputStream"] =
                        HttpContext.Current.Request.InputStream;

					HttpContext.Current.Request.InputStream.Position = 0;
					XmlReader reader = new XmlTextReader(HttpContext.Current.Request.InputStream);
					reader.ReadStartElement("Envelope",
						"http://schemas.xmlsoap.org/soap/envelope/");
					reader.MoveToContent();
					if (reader.LocalName == "Header") reader.Skip();
					reader.ReadStartElement("Body",
						"http://schemas.xmlsoap.org/soap/envelope/");
					reader.MoveToContent();

					soapPayload = new StringReader(reader.ReadOuterXml());
					
					// create new memory stream for method to write
                    // output message to and store reference in
                    // current HTTP context
                    appOutputStream = new MemoryStream();
                    HttpContext.Current.Items["SoapOutputStream"] =
                        appOutputStream;

                    break;
                }
                case SoapMessageStage.AfterSerialize :
                {
					// scan chained output stream looking for SOAP fault
                    chainedOutputStream.Position = 0;
                    XmlReader reader = new XmlTextReader(chainedOutputStream);
                    reader.ReadStartElement("Envelope",
                        "http://schemas.xmlsoap.org/soap/envelope/");
                    reader.MoveToContent();
                    if (reader.LocalName == "Header") reader.Skip();
                    reader.ReadStartElement("Body",
                        "http://schemas.xmlsoap.org/soap/envelope/");
                    reader.MoveToContent();
                    if (reader.LocalName == "Fault" && reader.NamespaceURI ==
                        "http://schemas.xmlsoap.org/soap/envelope/")
                    {
                        // fault was found, so rewind chained output stream
                        // and copy it to HTTP output stream
                        chainedOutputStream.Position = 0;
                        CopyStream(chainedOutputStream, httpOutputStream);
                    }
                    else
                    {
                        // No fault was found, so flush memory stream that
                        // method wrote its output message to, rewind it
                        // and copy it to the HTTP output stream

						chainedOutputStream.Position = 0;
						reader = new XmlTextReader(chainedOutputStream);
						reader.MoveToContent();
						XmlDocument xd = new XmlDocument();
						xd.LoadXml(reader.ReadOuterXml());
						Debug.WriteLine(xd.OuterXml);

						om4Interface(xd);

                        appOutputStream.Flush();
                        appOutputStream.Position = 0;
                        CopyStream(appOutputStream, httpOutputStream);
					}
                    appOutputStream.Close();
                    break;
                }
            }
        }    
        
        // CopyStream copies the contents of a source stream
        // to a destination stream
        private void CopyStream(Stream src, Stream dest)
        {
            StreamReader reader = new StreamReader(src);
            StreamWriter writer = new StreamWriter(dest);
            writer.Write(reader.ReadToEnd());
            writer.Flush();
        }

		private void om4Interface(XmlDocument soapOutXDoc)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			string sx = soapPayload.ReadToEnd();
			string sPath = GetXMLPath() + "\\";

			XslTransform xslt = new XslTransform();
			xslt.Load(sPath + "transformOmigaFlexiCalcRequest.xslt");

			XPathDocument sourceDocument = new XPathDocument(new StringReader(sx));

			StringWriter targetStringWriter = new StringWriter();
			XmlTextWriter targetXmlWriter = new XmlTextWriter(targetStringWriter);
			targetXmlWriter.Formatting = Formatting.None;

			xslt.Transform(sourceDocument, null, targetXmlWriter, null);
			sx = targetStringWriter.ToString();

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: alpha request: {1}", functionName,sx);
			}

			//	Debug.WriteLine(sx);
			
			XmlDocument xd = new XmlDocument();

			AlphaCOMPlus.Alpha ace = new AlphaCOMPlus.AlphaClass();
			string aceResponse = ace.aceRequest(sx);

			targetStringWriter = new StringWriter();
			targetXmlWriter = new XmlTextWriter(targetStringWriter);
			targetXmlWriter.Formatting = Formatting.None;

			xslt.Load(sPath + "transformOmigaFlexiCalcResponse.xslt");
			sourceDocument = new XPathDocument(new StringReader(aceResponse));
			xslt.Transform(sourceDocument, null, targetXmlWriter, null);
			aceResponse = targetStringWriter.ToString();

			xd.LoadXml(aceResponse);

			//	Debug.WriteLine(xd.OuterXml);

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: raw alpha response: {1}", functionName,xd.OuterXml);
			}

			XmlNode omigaResponseNode = xd.DocumentElement;

			XmlElement soapResponseNode = (XmlElement)soapOutXDoc.DocumentElement.FirstChild.FirstChild;

			soapResponseNode.InnerXml = omigaResponseNode.OuterXml;
			// IK 19/09/2006 EP2_1 ends

			//	Debug.WriteLine(soapOutXDoc.DocumentElement.OuterXml);

			StreamWriter sw = new StreamWriter(appOutputStream);
			sw.Write(soapOutXDoc.DocumentElement.OuterXml);
			sw.Flush();
		}
    }


    public class SoapStreams
    {
		public static Stream InputMessage
        {
            get 
            {
                return (Stream) HttpContext.Current.Items["SoapInputStream"];
            }
        }
        public static Stream OutputMessage
        {
            get
            {
                return (Stream) HttpContext.Current.Items["SoapOutputStream"];
            }
        }
    }

    [AttributeUsage(AttributeTargets.Method)]
    public class XmlStreamSoapExtensionAttribute : SoapExtensionAttribute
    {
        private int priority = 1;
        public override Type ExtensionType 
        {
            get { return typeof(XmlStreamSoapExtension); }
        }

        public override int Priority 
        {
            get { return priority; }
            set { priority = 1; }
        }
    }
}

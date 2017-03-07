//
//--------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/SubmitStopAndSaveFMAWS/XmlStreamSoapExtension.cs $
//Workfile:			$Workfile: XmlStreamSoapExtension.cs $
//Current Version	$Revision: 2 $
//Last Modified		$Modtime: 28/11/06 10:07 $
//Modified By		$Author: Dbarraclough $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//IK	19/09/2006	EP2_1 remove xmlns="" etc.
//SR    23/10/2006  EP2_1 : Modified to check in
//IK    21/11/2006  EP2_134 : prevent REQUEST de-serialization
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

namespace Vertex.Fsd.Omiga.Web.Services.SubmitStopAndSaveFMAWS
{
    // XmlStreamSoapExtension exposes raw SOAP messages to
    // an ASP.NET Web service
    public class XmlStreamSoapExtension : SoapExtension
    {
        bool output = false;		// flag indicating input or output
        Stream httpOutputStream;    // HTTP output stream to send
                                    // real output to
        Stream chainedOutputStream; // output stream for ASP.NET
                                    // plumbing to write to
        Stream appOutputStream;     // output stream for method
                                    // to write to
		// EP2_134
		private Stream httpInputStream;		// replacement HTTP input stream 
		// to avoid de-seria;isation
		private Stream chainedInputStream;	// copy of input stream

		StringReader soapPayload;

        // ChainStream replaces original stream with
        // extension’s stream
        public override Stream ChainStream(Stream stream)
        {
			// EP2_134
			const string soapPreamble = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><soap:Body>";
			const string soapPostamble = "</soap:Body></soap:Envelope>";
			const string emptyStopAndSaveFMARequest = soapPreamble + "<REQUEST xmlns='http://Request.SubmitStopAndSaveFMA.Omiga.vertex.co.uk'></REQUEST>" + soapPostamble;
			// EP2_134_ends
			
			Stream result = stream;
            if (output)
            {
                httpOutputStream = stream;
                result = chainedOutputStream = new MemoryStream();
            }
            else 
            {
				//	EP2_134
				httpInputStream = stream;

				chainedInputStream = new MemoryStream();
				StreamWriter sw = new StreamWriter(chainedInputStream);
				sw.Write(emptyStopAndSaveFMARequest);
				sw.Flush();

				chainedInputStream.Position = 0;
				result = chainedInputStream;
				// EP2_134_ends

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

						//	ik_wip
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
			string sx = soapPayload.ReadToEnd();
			Debug.WriteLine(sx);

			Type lbt = Type.GetTypeFromProgID("omWSInterface.omWSInterfaceBO");
			object lbo = Activator.CreateInstance(lbt);
			object[] lbop = new object[] {"StopAndSaveFMA",sx};
			XmlDocument xd = new XmlDocument();
			xd.LoadXml((string)lbt.InvokeMember("omRequest",BindingFlags.InvokeMethod,null,lbo,lbop));
			lbo = null;
			Debug.WriteLine(xd.OuterXml);

			XmlNode omigaResponseNode = xd.DocumentElement;
			XmlElement soapResponseNode = (XmlElement)soapOutXDoc.DocumentElement.FirstChild.FirstChild;

			// IK 19/09/2006 EP2_1
			foreach(XmlAttribute xa in omigaResponseNode.Attributes)
			{
				soapResponseNode.SetAttribute(xa.Name,xa.Value);
			}

			soapResponseNode.InnerXml = omigaResponseNode.InnerXml;
			// IK 19/09/2006 EP2_1 ends

			Debug.WriteLine(soapOutXDoc.DocumentElement.OuterXml);

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

//
//------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/IntroducerPipelineWS/XmlStreamSoapExtension.cs $
//Workfile:			$Workfile: XmlStreamSoapExtension.cs $
//Current Version	$Revision: 1 $
//Last Modified		$Modtime: 5/12/06 12:09 $
//Modified By		$Author: Dbarraclough $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//PSC	21/11/2006	EP2_138 Created
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
using Microsoft.Win32;

using omLogging = Vertex.Fsd.Omiga.omLogging;
using Vertex.Fsd.Omiga.omIntroducer;
using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.Web.Services.IntroducerPipelineWS
{
    // XmlStreamSoapExtension exposes raw SOAP messages to
    // an ASP.NET Web service
    public class XmlStreamSoapExtension : SoapExtension
    {		
		// Validation Error Count
		private static int errorsCount = 0;
		// Validation Error Message
		private static string errorMessage = "";

		private bool output = false;		// flag indicating input or output

// prevent (input) deserialisation
 
		private Stream httpInputStream;		// replacement HTTP input stream 
											// to avoid de-seria;isation
		private Stream chainedInputStream;	// copy of input stream

//	prevent (input) deserialisation ends

		private Stream httpOutputStream;    // HTTP output stream to send
						                    // real output to
        private Stream chainedOutputStream; // output stream for ASP.NET
						                    // plumbing to write to

		private Stream appOutputStream;     // output stream for method
						                    // to write to

		private StringReader soapPayload;

		private Stream soapPayloadStream;

		// schema cache
		private XmlSchemaCollection schemaCache;	
	
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private static void ValidationHandler(object sender,ValidationEventArgs args)
		{
			errorMessage = errorMessage + args.Message + " | ";
			errorsCount ++;
		}

        // ChainStream replaces original stream with
        // extension’s stream
        public override Stream ChainStream(Stream stream)
        {
//			const string soapPreamble = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><soap:Body>";
//			const string soapPostamble = "</soap:Body></soap:Envelope>";
//
//			const string getIntroducerPipelineRequest = soapPreamble + "<REQUEST xmlns='http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk'></REQUEST>" + soapPostamble;
            
			Stream result = stream;

			if (output)
            {
                httpOutputStream = stream;
                result = chainedOutputStream = new MemoryStream();
            }
            else 
            {

/* prevent (input) deserialisation

				httpInputStream = stream;

// ik_debug
//				StreamReader sr = new StreamReader(httpInputStream);
//				string s = sr.ReadToEnd();
//				Debug.Write(s);

				string soapAction = HttpContext.Current.Request.Headers.Get("SOAPAction");

				chainedInputStream = new MemoryStream();
				StreamWriter sw = new StreamWriter(chainedInputStream);

				// 'REQUEST' node depends on soap action
				if(soapAction.IndexOf("GetIntroducerPipeline") > 0)
				{
					sw.Write(getIntroducerPipelineRequest);
				}
				
				sw.Flush();
				
// ik_debug
//				chainedInputStream.Position = 0;
//				sr = new StreamReader(chainedInputStream);
//				s = sr.ReadToEnd();
//				Debug.Write(s);

				chainedInputStream.Position = 0;
				result = chainedInputStream;

	prevent (input) deserialisation ends */
				
				output = true;
			}
            return result;
        }
        
        public override object GetInitializer(System.Type serviceType)
        {
            return InitializerAssist();
        }

        public override object GetInitializer(System.Web.Services.Protocols.LogicalMethodInfo methodInfo, System.Web.Services.Protocols.SoapExtensionAttribute attribute)
        {
            return InitializerAssist();
        }

		private object InitializerAssist()
		{
			XmlSchemaCollection sc = new XmlSchemaCollection();

			//	EP2_30
			if(ValidationRequired())
			{
				sc.Add(LoadSchema("GetIntroducerPipelineRequest.xsd"));
				sc.Add(LoadSchema("GetIntroducerPipelineResponse.xsd"));
			}
			return(sc);
		}

		private XmlSchema LoadSchema(string schemaName)
		{
			XmlTextReader r = new XmlTextReader(GetXsdPath() + schemaName);
			XmlSchema s = XmlSchema.Read(r,null);
			return(s);
		}

        public override void Initialize(object initializer)
        {
			schemaCache = (XmlSchemaCollection)initializer;        
        }

        // ProcessMessage is called to process SOAP messages
        // after inbound messages are deserialized to input
        // parameters and output parameters are serialized to
        // outbound messages
        public override void ProcessMessage(SoapMessage message)
        {
			string functionName = System.Reflection.MethodInfo.GetCurrentMethod().Name; 

			// Clear out logging thread context 
			if (message.Stage == SoapMessageStage.BeforeDeserialize)
			{
				omLogging.ThreadContext.Properties["UserId"] = "Unknown";
				omLogging.ThreadContext.Properties["IntroducerId"] = "Unknown";
			}

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Processing message. Message stage = {1}.", functionName, message.Stage.ToString());
			}

			try
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

						soapPayloadStream = new MemoryStream();
						StreamWriter sw = new StreamWriter(soapPayloadStream);
						sw.Write(soapPayload.ReadToEnd());
						sw.Flush();

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

							Om4Interface(xd);

							appOutputStream.Flush();
							appOutputStream.Position = 0;
							CopyStream(appOutputStream, httpOutputStream);
						}
						appOutputStream.Close();
						break;
					}
				}
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Processed message. Message stage = {1}.", functionName, message.Stage.ToString());
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

		private void Om4Interface(XmlDocument soapOutXDoc)
		{
			string functionName = System.Reflection.MethodInfo.GetCurrentMethod().Name;
 
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				Debug.WriteLine(HttpContext.Current.Request.Headers.Get("SOAPAction"));

				string resp = "";
				string methodName = string.Empty;
				string soapAction = HttpContext.Current.Request.Headers.Get("SOAPAction");

				soapPayloadStream.Position = 0;
				StreamReader sr = new StreamReader(soapPayloadStream);
				string sx = sr.ReadToEnd();
				Debug.WriteLine(sx);
	            
				XmlDocument xReq = new XmlDocument();
				xReq.LoadXml(sx);

				if(soapAction.IndexOf("GetIntroducerPipeline") > 0)
				{
					resp = GetIntroducerPipeline(xReq);
				}

				XmlDocument xResp = new XmlDocument();
				xResp.LoadXml(resp);
				XmlNode omigaResponseNode = xResp.DocumentElement;

				if(xResp.SelectSingleNode("RESPONSE[@TYPE != 'SUCCESS']") != null)
				{
					AddErrorNamespaces(xResp);
				}

				XmlElement soapResponseNode = (XmlElement)soapOutXDoc.DocumentElement.FirstChild.FirstChild;

				// copy attributes from omiga response
				foreach(XmlAttribute xa in omigaResponseNode.Attributes)
				{
					soapResponseNode.SetAttribute(xa.Name,xa.Value);
				}

				soapResponseNode.InnerXml = omigaResponseNode.InnerXml;

				if(ValidationRequired())
				{
					if(soapAction.IndexOf("GetIntroducerPipeline") > 0)
					{
						string schemaName = "GetIntroducerPipelineResponse.xsd";

						try
						{
							if (_logger.IsDebugEnabled)
							{
								_logger.DebugFormat("{0}: Validating response against schema {1}.", functionName, schemaName);
							}

							ValidateMessage(soapResponseNode.OuterXml, schemaName);

							if (_logger.IsDebugEnabled)
							{
								_logger.DebugFormat("{0}: Response validated against schema {1}.", functionName, schemaName);
							}

						}
						catch (XmlException exp)
						{
							if (_logger.IsErrorEnabled)
							{
								_logger.Error(functionName + ": Failed to validate the response against the schema " + schemaName, exp);
							}

							throw new XmlException("The response did not validate against the schema " + schemaName + " Details: " + exp.Message);
						}
					}
				}

				StreamWriter sw = new StreamWriter(appOutputStream);
				sw.Write(soapOutXDoc.DocumentElement.OuterXml); 
				sw.Flush();
			
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		private string GetIntroducerPipeline(XmlDocument xReq)
		{
			string functionName = System.Reflection.MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				if(ValidationRequired())
				{
					string schemaName = "GetIntroducerPipelineRequest.xsd";
					
					try
					{

						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Validating request against schema {1}.", functionName, schemaName);
						}

						ValidateMessage(xReq.OuterXml, schemaName);
					
						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Request validated against schema {1}.", functionName, schemaName);
						}
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + ": Failed to validate the request against the schema " + schemaName, exp);
						}

						throw new XmlException("The request did not validate against the schema " + schemaName + " Details: " + exp.Message);
					}
				}

				xReq.LoadXml(StripNamespaces(xReq.OuterXml));
				Debug.WriteLine(xReq.OuterXml);		

				xReq.DocumentElement.SetAttribute("OPERATION", functionName);

				omIntroducer.IntroducerBO introducerBO = new omIntroducer.IntroducerBO();

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling IntroducerBO with request: {1}", functionName, xReq.OuterXml);
				}

				string resp = introducerBO.omRequest(xReq.OuterXml);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Called IntroducerBO with response: {1}", functionName, resp);
				}

				return(resp);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		private bool ValidationRequired()
		{
			try
			{
				RegistryKey hklm = Registry.LocalMachine;
				RegistryKey hkSoftware = hklm.OpenSubKey("Software");
				RegistryKey hkOmiga4 = hkSoftware.OpenSubKey("Omiga4");
				RegistryKey hkValidateWS = hkOmiga4.OpenSubKey("wsValidate");
				if(hkValidateWS.GetValue(null).ToString() == "1") return(true);
				else return(false);
			}
			catch
			{
				return(false);
			}
		}

		private void ValidateMessage(string xmlIn, string schemaName)
		{
			errorMessage = string.Empty;
			errorsCount = 0;

			XmlTextReader xtr = new XmlTextReader(GetXsdPath() + schemaName);
				
			XmlValidatingReader xvr = new XmlValidatingReader(xmlIn,XmlNodeType.Document,null);
			xvr.Schemas.Add(schemaCache);

			xvr.ValidationType = ValidationType.Schema;
			xvr.ValidationEventHandler +=
				new ValidationEventHandler(ValidationHandler);
				
			while(xvr.Read());
			xvr.Close();

			// Raise exception, if XML validation fails
			if (errorsCount > 0)
			{
				throw new XmlException(errorMessage);
			}
		}

		private string StripNamespaces(string xml)
		{
			XmlDocument xmlIn = new XmlDocument();
			xmlIn.LoadXml(xml);

			XmlDocument xmlOut = new XmlDocument();

			StripNode(xmlIn.DocumentElement,xmlOut,xmlOut);
			return(xmlOut.OuterXml);
		}

		private void StripNode(XmlNode sourceNode, XmlNode targetNode, XmlDocument targetDoc)
		{
			XmlElement xe = targetDoc.CreateElement(sourceNode.LocalName);
			foreach(XmlAttribute xa in sourceNode.Attributes)
			{
				if(xa.Name != "xmlns")
				{
					xe.SetAttribute(xa.Name,xa.Value);
				}
			}
			XmlNode xn = targetNode.AppendChild(xe);
			if(sourceNode.HasChildNodes)
			{
				for (int i=0; i < sourceNode.ChildNodes.Count; i++)
				{
					StripNode(sourceNode.ChildNodes[i],xn,targetDoc);
				}
			}
		}

		private void AddErrorNamespaces(XmlDocument xmlResponse)
		{
			XmlAttribute xa = xmlResponse.CreateAttribute("xmlns");
			xa.Value = "http://msgtypes.Omiga.vertex.co.uk";

			XmlNode xn = xmlResponse.DocumentElement.FirstChild;
			if(xn != null)
			{
				if(xn.LocalName == "MESSAGE" || xn.LocalName == "ERROR")
				{
					for(int i =0; i < xn.ChildNodes.Count; i++)
					{
						xn.ChildNodes[i].Attributes.SetNamedItem(xa);
					}
				}
			}
		}

		private string GetXsdPath()
		{
			try
			{
				RegistryKey hklm = Registry.LocalMachine;
				RegistryKey hkSoftware = hklm.OpenSubKey("Software");
				RegistryKey hkOmiga4 = hkSoftware.OpenSubKey("Omiga4");
				RegistryKey hkValidateWS = hkOmiga4.OpenSubKey("wsValidate");
				if(hkValidateWS.GetValue("xsdPath").ToString().Length == 0)
				{
					throw(new Exception("no .xsd path specified for xml validation"));
				}
				string xsdPath = hkValidateWS.GetValue("xsdPath").ToString();
				if(xsdPath.EndsWith(@"\") == false) xsdPath += @"\";
				return(xsdPath);
			}
			catch
			{
				throw(new Exception("no .xsd path specified for xml validation"));
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

//
//------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/IntroducerLogonAndValidationWS/XmlStreamSoapExtension.cs $
//Workfile:			$Workfile: XmlStreamSoapExtension.cs $
//Current Version	$Revision: 3 $
//Last Modified		$Modtime: 10/11/06 8:57 $
//Modified By		$Author: Dbarraclough $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//IK	01/11/2006	created for EP2_22
//IK	02/11/2006	EP2_28 - add namespace directives to error responses
//IK	02/11/2006	EP2_28 - add schema cache
//IK	02/11/2006	EP2_30 - fix to schema cache (coding standards)
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
using Vertex.Fsd.Omiga.omIntroducer;
using Microsoft.Win32;

namespace Vertex.Fsd.Omiga.Web.Services.IntroducerLogonAndValidationWS
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
		
		private static void ValidationHandler(object sender,ValidationEventArgs args)
		{
			errorMessage = errorMessage + args.Message + " | ";
			errorsCount ++;
		}

        // ChainStream replaces original stream with
        // extension’s stream
        public override Stream ChainStream(Stream stream)
        {
			const string soapPreamble = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><soap:Body>";
			const string soapPostamble = "</soap:Body></soap:Envelope>";

			const string getIntroducerUserFirmsRequest = soapPreamble + "<REQUEST xmlns='http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk'></REQUEST>" + soapPostamble;
			const string authenticateFirmRequest = soapPreamble + "<REQUEST xmlns='http://Request.AuthenticateFirm.Omiga.vertex.co.uk'></REQUEST>" + soapPostamble;
			const string validateBrokerRequest = soapPreamble + "<REQUEST xmlns='http://Request.ValidateBroker.Omiga.vertex.co.uk'></REQUEST>" + soapPostamble;

            Stream result = stream;

			if (output)
            {
                httpOutputStream = stream;
                result = chainedOutputStream = new MemoryStream();
            }
            else 
            {

// /* prevent (input) deserialisation

				httpInputStream = stream;

// ik_debug
//				StreamReader sr = new StreamReader(httpInputStream);
//				string s = sr.ReadToEnd();
//				Debug.Write(s);

				string soapAction = HttpContext.Current.Request.Headers.Get("SOAPAction");

				chainedInputStream = new MemoryStream();
				StreamWriter sw = new StreamWriter(chainedInputStream);

				// 'REQUEST' node depends on soap action
				if(soapAction.IndexOf("GetIntroducerUserFirms") > 0)
				{
					sw.Write(getIntroducerUserFirmsRequest);
				}
				else if (soapAction.IndexOf("AuthenticateFirm") > 0)
				{
					sw.Write(authenticateFirmRequest);
				}
				else if (soapAction.IndexOf("ValidateBroker") > 0)
				{
					sw.Write(validateBrokerRequest);
				}
				
				sw.Flush();
				
// ik_debug
//				chainedInputStream.Position = 0;
//				sr = new StreamReader(chainedInputStream);
//				s = sr.ReadToEnd();
//				Debug.Write(s);

				chainedInputStream.Position = 0;
				result = chainedInputStream;

//	prevent (input) deserialisation ends */
				
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
				sc.Add(LoadSchema("AuthenticateFirmRequest.xsd"));
				sc.Add(LoadSchema("AuthenticateFirmResponse.xsd"));
				sc.Add(LoadSchema("GetIntroducerUserFirmsRequest.xsd"));
				sc.Add(LoadSchema("GetIntroducerUserFirmsResponse.xsd"));
				sc.Add(LoadSchema("ValidateBrokerRequest.xsd"));
				sc.Add(LoadSchema("ValidateBrokerResponse.xsd"));
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
						Debug.WriteLine(xd.OuterXml);

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

				if(soapAction.IndexOf("GetIntroducerUserFirms") > 0)
				{
					resp = GetIntroducerUserFirms(xReq);
				}
				else if (soapAction.IndexOf("AuthenticateFirm") > 0)
				{
					resp = AuthenticateFirm(xReq);
				}
				else if (soapAction.IndexOf("ValidateBroker") > 0)
				{
					resp = ValidateBroker(xReq);
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
					if(soapAction.IndexOf("GetIntroducerUserFirms") > 0)
					{
						ValidateMessage(soapResponseNode.OuterXml,"GetIntroducerUserFirmsResponse.xsd");
					}
					else if (soapAction.IndexOf("AuthenticateFirm") > 0)
					{
						ValidateMessage(soapResponseNode.OuterXml,"AuthenticateFirmResponse.xsd");
					}
					else if (soapAction.IndexOf("ValidateBroker") > 0)
					{
						ValidateMessage(soapResponseNode.OuterXml,"ValidateBrokerResponse.xsd");
					}
				}

				StreamWriter sw = new StreamWriter(appOutputStream);
				sw.Write(soapOutXDoc.DocumentElement.OuterXml); 
				sw.Flush();
			
			}
			catch (Vertex.Fsd.Omiga.Core.OmigaException oe)
			{
				Debug.WriteLine(oe.FullMessage);
			}
		}

		private string GetIntroducerUserFirms(XmlDocument xReq)
		{
			string methodName = "GetIntroducerUserFirms";

			if(ValidationRequired())
				ValidateMessage(xReq.OuterXml,"GetIntroducerUserFirmsRequest.xsd");

			xReq.LoadXml(StripNamespaces(xReq.OuterXml));
			Debug.WriteLine(xReq.OuterXml);		

			xReq.DocumentElement.SetAttribute("OPERATION",methodName);

			XmlElement xe = (XmlElement)xReq.SelectSingleNode("REQUEST/INTRODUCERCREDENTIALS");
			if(xe != null)
			{
				XmlAttribute xa = (XmlAttribute)xe.Attributes.GetNamedItem("INTRODUCERUSERPASSWORD");
				if(xa != null)
				{
					xe.SetAttribute("INTRODUCERUSERPASSWORD",Encrypt(xa.Value));
				}
			}
			Debug.WriteLine(xReq.OuterXml);

			omIntroducer.IntroducerBO introducerBO = new omIntroducer.IntroducerBO();

			string resp = introducerBO.omRequest(xReq.OuterXml);
			return(resp);
		}

		private string AuthenticateFirm(XmlDocument xReq)
		{
			string methodName = "AuthenticateFirm";

			if(ValidationRequired())
				ValidateMessage(xReq.OuterXml,"AuthenticateFirmRequest.xsd");

			xReq.LoadXml(StripNamespaces(xReq.OuterXml));
			Debug.WriteLine(xReq.OuterXml);		

			xReq.DocumentElement.SetAttribute("OPERATION",methodName);

			omIntroducer.IntroducerBO introducerBO = new omIntroducer.IntroducerBO();

			string resp = introducerBO.omRequest(xReq.OuterXml);
			return(resp);
		}

		private string ValidateBroker(XmlDocument xReq)
		{
			string methodName = "ValidateBroker";
			string resp = null;

			if(ValidationRequired())
				ValidateMessage(xReq.OuterXml,"ValidateBrokerRequest.xsd");

			xReq.LoadXml(StripNamespaces(xReq.OuterXml));
			Debug.WriteLine(xReq.OuterXml);		

			xReq.DocumentElement.SetAttribute("OPERATION",methodName);

			omIntroducer.IntroducerBO introducerBO = new omIntroducer.IntroducerBO();
			resp = introducerBO.omRequest(xReq.OuterXml);
			Debug.WriteLine(resp);		

			XmlDocument respXml = new XmlDocument();
			respXml.LoadXml(resp);
			if(respXml.SelectNodes("RESPONSE/ARFIRM").Count > 1)
			{
				foreach(XmlNode arNode in respXml.SelectNodes("RESPONSE/ARFIRM"))
				{
					while(arNode.HasChildNodes)
					{
						arNode.RemoveChild(arNode.FirstChild);
					}
				}
			}

			resp = respXml.DocumentElement.OuterXml;
			Debug.WriteLine(resp);		
			return(resp);
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
			try
			{
				errorMessage = "";
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
					throw new Exception(errorMessage);
				}
			}
			catch(Exception e)
			{
				throw(e);
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

		private string Encrypt(string plainText)
		{
			int inputLength = plainText.Length;
			int charPosition = 0;
			string cypherText = string.Empty;
			byte increment;
			byte char1;
			byte char2;
			const byte asciiTilde = 126;

			byte[] random = new byte[30] {23, 12, 33, 40, 17, 31, 23, 25, 12, 43, 44, 43, 48, 43, 13, 35, 17, 23, 17, 37, 21, 34, 10, 38, 11, 12, 41, 21, 43, 23};
			byte randomIndex = 0;

			while (charPosition < inputLength)
			{
				increment = random[randomIndex++];
				char1 = (byte) plainText.ToCharArray(charPosition++, 1)[0];
				char2 = (byte) (char1 + increment);

				if (char2 >= asciiTilde)
				{
					cypherText += '~';
					char2 = (byte) (char1 - increment);
				}

				cypherText += (char) char2;
			}
			return cypherText;
		}
	}

/*
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
*/

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

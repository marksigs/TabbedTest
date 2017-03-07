/*--------------------------------------------------------------------------------------------
Workfile:		AdminInterfaceBO.cs
Copyright:		Copyright © 2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog		Date		Description
RFairlie	08/11/2005	MAR451 Ensure completions interface works as part of batch job. 
						This is a continuation of MAR225. Added debugging.
RFairlie	10/11/2005	MAR499 Need to implement handling of response from createmortgageaccount
PCarter		11/11/2005	MAR352 Added methods to call Customer or Prospect Web Services
PCarter		14/11/2005	MAR520 Additional completions work
PCarter		14/11/2005	MAR352 Additional changes
PCarter		16/11/2005	MAR591 Amend UpdateCRSCustomer to get OMIGACUSTMERNUMBER correctly
PCarter		17/11/2005	MAR618 Amend UpdateCRSCustomer to get CIF correctly
HAldred     17/11/2005  MAR641 Amend FindCustomer to format DOB correctly
GHun		18/11/2005	MAR651 Amend CreateSOAPRequestForCompletions to set xsi:nil="true" on Operator
PCarter		22/11/2005	MAR671 Amend FindCustomer to backout MAR641 and correct other parameters
PCarter		24/11/2005	MAR673 Amend OnMessage so that it tries to return exception as a
                               completion reponse if possible. Amended the way error responses 
							   are built
PCarter		30/11/2005	MAR733 Amend to create correct error response message
PCarter		05/12/2005	MAR804 Amend SendSOAPRequestForCompletions to cater for null Proxy
GHun		13/12/2005	MAR852 set Operator in generic request					   
PCarter		13/12/2005	MAR457 Use omLogging wrapper
PCarter		20/12/2005  MAR926 Fix the Time Out Issue
MHeys		10/04/2006	MAR1603	Terminate COM+ objects cleanly
MHeys		12/04/2006	MAR1603	Build errors rectified
RFairlie	19/06/2006  MAR1911 Configurably stub out calls to DG 
--------------------------------------------------------------------------------------------*/

using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Text;
using MESSAGEQUEUECOMPONENTVCLib;
using ExternalXML = Vertex.Fsd.Omiga.ExternalXML;
using WebRef = Vertex.Fsd.Omiga.omAdminCS.WebRefCreateMortgageAccount;
using System.Net;
using Vertex.Fsd.Omiga.Core;

// PSC 03/11/2005 MAR352 - Start
using System.Xml.Xsl;
using System.Reflection;
using CustomerSearch = Vertex.Fsd.Omiga.omAdminCS.WebRefSearchForCustomerOrProspect;
using UpdateCustomer = Vertex.Fsd.Omiga.omAdminCS.WebRefUpdateCustomerOrProspect;
using GetCustomer = Vertex.Fsd.Omiga.omAdminCS.WebRefGetCustomerOrProspectDetails;
using omDG = Vertex.Fsd.Omiga.omDg;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
// PSC 03/11/2005 MAR352 - End

// PSC 14/11/2005 MAR520 - Start
using omMQ;
// PSC 14/11/2005 MAR520 - End


namespace Vertex.Fsd.Omiga.omAdminCS
{
	/// <summary>
	/// </summary>
	[ProgId("omAdminCS.AdminInterfaceBO")]
	[Guid("82BAF7D3-B048-40d5-8A55-1CB7CCD217D1")]
	[ComVisible(true)]
	public class AdminInterfaceBO : IMessageQueueComponentVC2 
	{
		// PSC 13/12/2005 MAR457
		private  omLogging.Logger _logger = null;

		public AdminInterfaceBO()
		{
			// PSC 13/12/2005 MAR457
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.FullName + ".Default");
		}

		// PSC 14/11/2005 MAR520
		private XmlDocument CreateSOAPRequestForCompletions(string in_xmlData, out AdminInterfaceBO_OnMessage_Input convertedInput)
		{
			// in_xmlData is the message from the queue

			#region Validation

			//			// validate
			//			XmlValidatingReader xmlValidatingReader = null;
			//			try
			//			{
			//				XmlReader xmlReader = new XmlTextReader(new StringReader(in_xmlData));
			//				xmlValidatingReader = new XmlValidatingReader(xmlReader);
			//
			//				XmlSchemaCollection schemas = new XmlSchemaCollection();
			//				ValidationEventHandler validationHandlerSchema = new ValidationEventHandler(ValidationHandlerSchema);
			//
			//				// add schemas
			//				System.IO.Stream stream;
			//				XmlSchema schema;
			//				// ... GenericMessages
			//				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.GenericMessages.xsd");
			//				schema = XmlSchema.Read(stream, validationHandlerSchema);
			//				schemas.Add(schema);
			//				// ... CommonTypes
			//				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.CommonTypes.xsd");
			//				schema = XmlSchema.Read(stream, validationHandlerSchema);
			//				schemas.Add(schema);
			//				// ... CreateMortgageAccount
			//				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.CreateMortgageAccount.xsd");
			//				schema = XmlSchema.Read(stream, validationHandlerSchema);
			//				schemas.Add(schema);
			//				// ... AdminInterfaceBO_OnMessage_Input
			//				stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.omAdminCS.AdminInterfaceBO_OnMessage_Input.xsd");
			//				schema = XmlSchema.Read(stream, validationHandlerSchema);
			//				schemas.Add(schema);
			//				xmlValidatingReader.Schemas.Add(schemas);
			//				xmlValidatingReader.ValidationType = ValidationType.Schema;
			//				xmlValidatingReader.ValidationEventHandler += new ValidationEventHandler(ValidationHandlerSchemaInstance);
			//				while (xmlValidatingReader.Read()) ;
			//
			//				xmlValidatingReader.Close();
			//				xmlReader.Close();
			//			}
			//			catch (Exception ex)
			//			{
			//				throw;
			//			}
			//			finally
			//			{
			//				if (xmlValidatingReader != null)
			//				{
			//					xmlValidatingReader.Close();
			//				}
			//			}

			#endregion

			// Theres a bug in .Net framework 1.1 which prevents CreateMortgageAccount from being deserialized
			// (fixed in .Net Framework 2.0, but we cannot use it yet)

			// ... AdminInterfaceBO_OnMessage_Input modified to just get access to the string contents
			XmlReader xmlReader = new XmlTextReader(new StringReader(in_xmlData));
			XmlSerializer serializer = new XmlSerializer(typeof(AdminInterfaceBO_OnMessage_Input));
			AdminInterfaceBO_OnMessage_Input input = (AdminInterfaceBO_OnMessage_Input)serializer.Deserialize(xmlReader);
			string createMortgageAccountRequestString = ((XmlElement)input.OtherElements[0]).OuterXml;
			string userId = input.USERID;
			string password = input.PASSWORD;

			// now working in strings rather than classes

			// ... create an empty soap envelope which can be populated into which the string can be added
			XmlDocument doc = new XmlDocument();
			doc.LoadXml("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"><soap:Body><request xmlns=\"http://interfaces\"/></soap:Body></soap:Envelope>");
			// ... add the body
			XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
			nsmgr.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/");
			nsmgr.AddNamespace("CreateMortgageAccount", "http://createmortgageaccount.services.dg.ingdirect.com/0.0.1");
			nsmgr.AddNamespace("interfaces", "http://interfaces");
			XmlNode interfacesRequest = doc.SelectSingleNode("soap:Envelope/soap:Body/interfaces:request", nsmgr);
			interfacesRequest.InnerXml = createMortgageAccountRequestString;

			// ... need to modify the createMortgageAccountRequest in the soap body before sending
			//const string emptyNamespace = "";
			Vertex.Fsd.Omiga.omDg.DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();				
			XmlElement requestElement = (XmlElement)interfacesRequest.SelectSingleNode("CreateMortgageAccount:CreateMortgageAccountRequest", nsmgr);
				
			// ... ... CustomerNumber
			string customerNumber = input.OTHERSYSTEMCUSTOMERNUMBER;
			if (customerNumber == null || customerNumber.Length == 0)
			{
				customerNumber = "NA";
			}
			//XmlInsertElementBeforeFirstChild(doc, requestElement, "CustomerNumber", emptyNamespace, customerNumber);
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//CustomerNumber");
				temp.InnerXml = customerNumber;
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set CustomerNumber: " + ex.Message);
			}

			//XmlInsertElementBeforeFirstChild(doc, requestElement, "ServiceName", emptyNamespace, dg.GetServiceName());
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//ServiceName");
				temp.InnerXml = dg.GetServiceName();
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set ServiceName: " + ex.Message);
			}

			//XmlInsertNilElementBeforeFirstChild(doc, requestElement, "CommunicationDirection", emptyNamespace);

			//XmlInsertElementBeforeFirstChild(doc, requestElement, "CommunicationChannel", emptyNamespace, dg.GetCommunicationChannel(userId != null && userId.Length > 0));
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//CommunicationChannel");
				temp.InnerXml = dg.GetCommunicationChannel(userId != null && userId.Length > 0);
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set CommunicationChannel: " + ex.Message);
			}

			//XmlInsertNilElementBeforeFirstChild(doc, requestElement, "SessionID", emptyNamespace);
				
			//XmlInsertElementBeforeFirstChild(doc, requestElement, "ProductType", emptyNamespace, dg.GetProductType());
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//ProductType");
				temp.InnerXml = dg.GetProductType();
			}
			catch(OmigaException ex)
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(ex.ToOmiga4Response(), ex);
				}
				throw new Exception("Failed to set ProductType: " + ex.Message);
			}
			catch(Exception ex)
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Failed to set ProductType: ", ex);
				}
				throw new Exception("Failed to set ProductType: " + ex.Message);
			}

			//XmlInsertNilElementBeforeFirstChild(doc, requestElement, "Operator", emptyNamespace);

			//XmlInsertElementBeforeFirstChild(doc, requestElement, "ProxyPwd", emptyNamespace, dg.GetProxyPwd());
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//ProxyPwd");
				temp.InnerXml = dg.GetProxyPwd();
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set ProxyPwd: " + ex.Message);
			}
			//XmlInsertElementBeforeFirstChild(doc, requestElement, "ProxyID", emptyNamespace, dg.GetProxyId());
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//ProxyID");
				temp.InnerXml = dg.GetProxyId();
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set ProxyID: " + ex.Message);
			}

			// ... ... Id and Pwd
			if (userId != null && password != null &&
				userId.Length > 0 && password.Length > 0)
			{
				//XmlInsertElementBeforeFirstChild(doc, requestElement, "TellerPwd", emptyNamespace, password);
				try
				{
					XmlNode temp = requestElement.SelectSingleNode(".//TellerPwd");
					temp.InnerXml = password;
				}
				catch(Exception ex)
				{
					throw new Exception("Failed to set TellerPwd: " + ex.Message);
				}
				//XmlInsertElementBeforeFirstChild(doc, requestElement, "TellerID", emptyNamespace, userId);
				try
				{
					XmlNode temp = requestElement.SelectSingleNode(".//TellerID");
					temp.InnerXml = userId;
				}
				catch(Exception ex)
				{
					throw new Exception("Failed to set TellerID: " + ex.Message);
				}
			}
			//				else
			//				{
			//					//XmlInsertNilElementBeforeFirstChild(doc, requestElement, "TellerPwd", emptyNamespace);
			//					//XmlInsertNilElementBeforeFirstChild(doc, requestElement, "TellerID", emptyNamespace);
			//				}
			//XmlInsertElementBeforeFirstChild(doc, requestElement, "ClientDevice", emptyNamespace, dg.GetClientDevice());
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//ClientDevice");
				temp.InnerXml = dg.GetClientDevice();
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set ClientDevice: " + ex.Message);
			}

			//MAR852 GHun
			try
			{
				XmlNode temp = requestElement.SelectSingleNode(".//Operator");
				temp.InnerXml = dg.GetOperator();
			}
			catch(Exception ex)
			{
				throw new Exception("Failed to set Operator: " + ex.Message);
			}

			////MAR651 GHun
			//XmlElement xOperator = (XmlElement) requestElement.SelectSingleNode(".//Operator");
			//if (xOperator.InnerText.Length == 0)
			//	xOperator.SetAttribute("nil", "http://www.w3.org/2001/XMLSchema-instance", "true");
			////MAR651 End
			//MAR852 End

			// PSC 14/11/2005 MAR520
			convertedInput = input;

			return doc;
		}


		private string SendSOAPRequestForCompletions(XmlDocument doc)
		{
			// create a memory stream which will hold the soap request, UTF-8 encoded.
			MemoryStream ms = new MemoryStream();
			TextWriter sw = new StreamWriter(ms, Encoding.UTF8);
			XmlTextWriter xw = new XmlTextWriter(sw);
			doc.Save(xw);

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Request to WS: " + doc.InnerXml);
			}

			// ... create the web request
			//WebRequest webRequest = WebRequest.Create("http://localhost:8080");				
				
			Vertex.Fsd.Omiga.omDg.DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();				

			// RF Use HttpWebRequest 
			HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(dg.GetServiceUrl());
				
			//IWebProxy proxy = new WebProxy("LOCALHOST:8080", true);
			//proxy.Credentials = CredentialCache.DefaultCredentials;
			//webRequest.Proxy = proxy;

			// PSC 05/12/2005 - MAR804
			WebProxy currentProxy = dg.GetProxy();

			if (currentProxy != null)
			{
				webRequest.Proxy = currentProxy;
			}
			else
			{
				webRequest.Proxy = GlobalProxySelection.GetEmptyWebProxy();
			}
			// PSC 05/12/2005 - MAR804

			webRequest.Method = "POST";
			webRequest.ContentType = "text/xml";
			webRequest.Headers.Add("SOAPAction", "\"\"");

			System.Text.Encoding encoding = System.Text.Encoding.UTF8;
			byte[] bytes = ms.GetBuffer();
			int length = (int)ms.Length; //bytes.GetLength(0);
			webRequest.ContentLength = length;

			// PSC 20/12/2005 MAR926 - Start
			Stream thisStream = webRequest.GetRequestStream();
			thisStream.Write(bytes, 0, length);
			thisStream.Close();
			// PSC 20/12/2005 MAR926 - End
			
			HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse();

			Stream respStream = response.GetResponseStream();

			// use a StreamReader to read the entire response into a string 
			// (n.b. response.ContentLength is not supported)
			StreamReader reader = new StreamReader(respStream, Encoding.UTF8);
			String respHTML = reader.ReadToEnd();

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Response from WS: " + respHTML);
			}

			// Close the response stream.
			respStream.Close();

			return respHTML;
		}


		// PSC 14/11/2005 MAR520 - Start
		private void HandleSOAPResponseForCompletions(AdminInterfaceBO_OnMessage_Input convertedInput, string gatewayResponse)
		{
			// PSC 24/11/2005 MAR673 - Start
			// Set up request for the queue
			XmlDocumentEx messageDocument = new XmlDocumentEx();
			XmlDocumentEx gatewayResponseDocument = new XmlDocumentEx();

			gatewayResponseDocument.LoadXml(gatewayResponse);

			XmlElementEx requestElement = (XmlElementEx) messageDocument.CreateElement("REQUEST");
			requestElement.SetAttribute("OPERATION", "CompleteInterfacing");
			XmlElementEx headerElement = BuildResponseHeader(convertedInput);
			headerElement = (XmlElementEx)messageDocument.ImportNode(headerElement, true);
			requestElement.AppendChild(headerElement);
			// PSC 24/11/2005 MAR673 - End
			
			XmlElementEx responseElement = (XmlElementEx) messageDocument.CreateElement("RESPONSE");
			requestElement.AppendChild(responseElement);

			string errorCode = gatewayResponseDocument.GetMandatoryNodeText(".//ErrorCode");

			if (errorCode != "0")
			{
				responseElement.SetAttribute("TYPE", "APPERR");
				XmlElementEx errorElement = (XmlElementEx) messageDocument.CreateElement("ERROR");
				responseElement.AppendChild(errorElement);
				XmlElementEx tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORNUMBER");	
				tempElement.InnerText =  errorCode;
				errorElement.AppendChild(tempElement);
				tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORDESCRIPTION");
				tempElement.InnerText = gatewayResponseDocument.GetNodeText(".//ErrorMessage");
				errorElement.AppendChild(tempElement);
				tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORSOURCE");
				tempElement.InnerText = "Administration System";
				errorElement.AppendChild(tempElement);
			}
			else
			{
				responseElement.SetAttribute("TYPE", "SUCCESS");
				headerElement.SetAttribute("ACCOUNTNUMBER", gatewayResponseDocument.GetMandatoryNodeText(".//MortgageAccountNumber"));
				XmlElementEx tempElement = (XmlElementEx) messageDocument.CreateElement("DISBURSEMENTPAYMENT");	
				responseElement.AppendChild(tempElement);

				// Convert date to Omiga Format
				string firstPaymentDate = gatewayResponseDocument.GetMandatoryNodeText(".//DateOfFirstPayment");
				firstPaymentDate = firstPaymentDate.Substring(8,2) + "/" +
								   firstPaymentDate.Substring(5,2) + "/" +
								   firstPaymentDate.Substring(0,4);

				tempElement.SetAttribute("FIRSTPAYMENTDATE", firstPaymentDate);
				tempElement.SetAttribute("FIRSTPAYMENT", gatewayResponseDocument.GetMandatoryNodeText(".//AmountOfFirstPayment"));
				tempElement.SetAttribute("REGULARPAYMENT", gatewayResponseDocument.GetMandatoryNodeText(".//AmountOfSubsequentPayment"));
			}

			// PSC 24/11/2005 MAR673
			SendCompletionsResponse(requestElement);
		}
		// PSC 14/11/2005 MAR520 - End


		int IMessageQueueComponentVC2.OnMessage(string in_xmlConfig, string in_xmlData)
		{
			MESSQ_RESP result = MESSQ_RESP.MESSQ_RESP_SUCCESS;

			try
			{
				// PSC 13/12/2005 MAR457
				_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.FullName + ".MQL");

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Message received by omAdminCS: " + in_xmlData);
				}

				// PSC 14/11/2005 MAR520 - Start
				AdminInterfaceBO_OnMessage_Input convertedInput;
				XmlDocument requestDoc = CreateSOAPRequestForCompletions(in_xmlData, out convertedInput);
				// PSC 14/11/2005 MAR520 - End 

				string resp = SendSOAPRequestForCompletions(requestDoc);

				// PSC 14/11/2005 MAR520
				HandleSOAPResponseForCompletions(convertedInput, resp);
			}
			// PSC 24/11/2005 MAR673 - Start
			catch(OmigaException ex)
			{
				try
				{
					XmlElementEx requestElement = BuildExceptionResponse(in_xmlData, ex);
					SendCompletionsResponse(requestElement);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("Exception caught. Returning MESSQ_RESP_RETRY_MOVE_MESSAGE. " + 
							exp.ToOmiga4Response(), exp);
					}
					result = MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
				}
				catch (Exception exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("Exception caught. Returning MESSQ_RESP_RETRY_MOVE_MESSAGE", exp);
					}
					result = MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
				}
			}
			catch (Exception ex)
			{
				try
				{
					// PSC 30/11/2005 MAR733
					OmigaException thisException = new OmigaException("An error occured processing the message. " + ex.Message, ex);
					XmlElementEx requestElement = BuildExceptionResponse(in_xmlData, thisException);
					SendCompletionsResponse(requestElement);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("Exception caught. Returning MESSQ_RESP_RETRY_MOVE_MESSAGE. " + 
							exp.ToOmiga4Response(), exp);
					}
					result = MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
				}
				catch (Exception exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("Exception caught. Returning MESSQ_RESP_RETRY_MOVE_MESSAGE", exp);
					}
					result = MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
				}
			}
			// PSC 24/11/2005 MAR673 - End

			return (int)result;
		}
	
		private static void ValidationHandlerSchema(object sender, ValidationEventArgs args)
		{
			System.Diagnostics.Debugger.Break();
		}
		private static void ValidationHandlerSchemaInstance(object sender, ValidationEventArgs args)
		{
			System.Diagnostics.Debugger.Break();
		}

		private void Validate(string xml)
		{
			// validate (requires .Net Framework 1.1 with SERVICE PACK 1 applied)
			XmlValidatingReader xmlValidatingReader = null;
			try
			{
				XmlReader xmlReader = new XmlTextReader(new StringReader(xml));
				xmlValidatingReader = new XmlValidatingReader(xmlReader);

				XmlSchemaCollection schemas = new XmlSchemaCollection();
				ValidationEventHandler validationHandlerSchema = new ValidationEventHandler(ValidationHandlerSchema);

				// add schemas
				System.IO.Stream stream;
				XmlSchema schema;
				// ... GenericMessages
				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.GenericMessages.xsd");
				schema = XmlSchema.Read(stream, validationHandlerSchema);
				schemas.Add(schema);
				// ... CommonTypes
				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.CommonTypes.xsd");
				schema = XmlSchema.Read(stream, validationHandlerSchema);
				schemas.Add(schema);
				// ... CreateMortgageAccount
				stream = ExternalXML.Assembly.GetAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.ExternalXML.IDUK.DirectGateway.wsdl.com.ingdirect.dg.business.services.CreateMortgageAccount.xsd");
				schema = XmlSchema.Read(stream, validationHandlerSchema);
				schemas.Add(schema);
				// ... AdminInterfaceBO_OnMessage_Input
				stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Vertex.Fsd.Omiga.omAdminCS.AdminInterfaceBO_OnMessage_Input.xsd");
				schema = XmlSchema.Read(stream, validationHandlerSchema);
				schemas.Add(schema);
				xmlValidatingReader.Schemas.Add(schemas);
				xmlValidatingReader.ValidationType = ValidationType.Schema;
				xmlValidatingReader.ValidationEventHandler += new ValidationEventHandler(ValidationHandlerSchemaInstance);
				while (xmlValidatingReader.Read()) ;

				xmlValidatingReader.Close();
				xmlReader.Close();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			finally
			{
				if (xmlValidatingReader != null)
				{
					xmlValidatingReader.Close();
				}
			}
		}	

		// PSC 03/11/2005 MAR352 - Start
		public string ExecuteService(string serviceRequest)
		{
			const string functionName = "ExecuteService";

			XmlDocumentEx requestDoc = new XmlDocumentEx();
			XmlElementEx requestElement = null;
			string operation = string.Empty;
			string response = string.Empty;

			try
			{
				try
				{
					// Parse the input
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " Request: " + serviceRequest); 
					}

					requestDoc.LoadXml(serviceRequest);

					// Determine the operation
					requestElement = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
					operation = requestElement.GetMandatoryAttribute("OPERATION");

				}
				catch (XmlException exp)
				{				
					throw new OmigaException("The request did not contain valid xml", exp);
				}

				// Forward request to appropriate method
				switch (operation.ToUpper())
				{
					case "FINDCUSTOMER":
					{
						response = FindCustomer(requestElement);
						break;
					}
					case "GETCUSTOMERDETAILS":
					{
						response = GetCustomerDetails(requestElement);
						break;
					}
					case "UPDATECRSCUSTOMER":
					{
						response = UpdateCRSCustomer(requestElement);
						break;
					}
					default:
					{
						throw new OmigaException("Operation " + operation + " is not supported.");
					}
				}

				// PSC 23/11/2005 MAR673 - Start
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " Response: " + response); 
				}
				// PSC 23/11/2005 MAR673 - End
				
				return response;
			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " " + exp.Message, exp); 
				}
				return exp.ToOmiga4Response();
			}
			catch (Exception exp)
			{
				string errorMessage = "An unexpected error occurred";
				
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " " + errorMessage, exp); 
				}
				OmigaException thisException = new OmigaException(errorMessage, exp);

				return thisException.ToOmiga4Response();
			}
		}

		private string FindCustomer(XmlElementEx requestElement)
		{
			const string functionName = "FindCustomer";

			CustomerSearch.SearchForCustomerOrProspectRequestType gatewayRequest = new CustomerSearch.SearchForCustomerOrProspectRequestType();
			CustomerSearch.SearchForCustomerOrProspectResponseType gatewayResponse = new CustomerSearch.SearchForCustomerOrProspectResponseType();
			CustomerSearch.DirectGatewaySoapService soapService = new CustomerSearch.DirectGatewaySoapService(); 
			omDG.DirectGatewayBO directGatewayHelper = new omDG.DirectGatewayBO();
			
			try
			{
				// Create the request for the direct gateway
				soapService.Url = directGatewayHelper.GetServiceUrl();

//				IWebProxy proxy = new WebProxy("LOCALHOST:8080", true);
//				proxy.Credentials = CredentialCache.DefaultCredentials;
//				soapService.Proxy = proxy;
				soapService.Proxy = directGatewayHelper.GetProxy();

				XmlElementEx customerElement = (XmlElementEx)requestElement.SelectMandatorySingleNode("CUSTOMER");
		
				// PSC 22/11/2005 MAR671 - Start
				string lastName = customerElement.GetAttribute("SURNAME");

				if (lastName.Length > 0)
				{
					gatewayRequest.LastName = lastName.Replace("*", "%");
				}

				string firstName = customerElement.GetAttribute("FIRSTFORENAME");

				if (firstName.Length > 0)
				{	
					gatewayRequest.FirstName = firstName.Replace("*", "%");
				}
				
				string dateOfBirth = customerElement.GetAttribute("DATEOFBIRTH");

				//MAR641  DoB is stored yyyy-mm-dd
				if (dateOfBirth.Length > 0)
				{
					gatewayRequest.DOB = dateOfBirth.Substring(6,4) + "/" +
										 dateOfBirth.Substring(3,2) + "/" +
										 dateOfBirth.Substring(0,2);
				}

				string postCode = customerElement.GetAttribute("POSTCODE");
				
				if (postCode.Length > 0)
				{
					gatewayRequest.PostCode = postCode.Replace("*", "%");
				}
				// PSC 22/11/2005 MAR671 - End

				gatewayRequest.ClientDevice = directGatewayHelper.GetClientDevice();
				gatewayRequest.ServiceName = directGatewayHelper.GetServiceName();
	
				string userId = (requestElement != null) ? requestElement.GetAttribute("USERID"): String.Empty;
				string password = (requestElement != null) ? requestElement.GetAttribute("PASSWORD") : String.Empty;

				if (userId.Length > 0 && password.Length > 0)
				{
					gatewayRequest.TellerID = userId;
					gatewayRequest.TellerPwd = password;
				}
		
				gatewayRequest.ProxyID = directGatewayHelper.GetProxyId();
				gatewayRequest.ProxyPwd = directGatewayHelper.GetProxyPwd();
				
				gatewayRequest.Operator = directGatewayHelper.GetOperator();	//MAR852 GHun

				gatewayRequest.ProductType = directGatewayHelper.GetProductType();
				gatewayRequest.CommunicationChannel = directGatewayHelper.GetCommunicationChannel((userId.Length > 0));
				gatewayRequest.CustomerNumber = "NA";

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = directGatewayHelper.GetXmlFromGatewayObject(typeof(CustomerSearch.SearchForCustomerOrProspectRequestType), gatewayRequest);
					_logger.Debug(functionName + " Calling Direct Gateway SearchForCustomerOrProspect. Request: "   + requestString);
				}
			}
			catch (Exception exp)
			{
				throw new OmigaException("An error occurred trying to create the Direct Gateway request", exp);
			}

			try
			{
				// RFairlie	19/06/2006 MAR1911 Start - Configurably stub out calls to DG 

				bool useStub = UseStub();

				if (useStub == true)
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(functionName + ": Using Direct Gateway stub.");
					}

					// create dummy response
					//gatewayResponse.ErrorCode = "0";
					//gatewayResponse.ErrorMessage = "No Error";

					gatewayResponse = 
						(CustomerSearch.SearchForCustomerOrProspectResponseType)LoadGatewayObjectFromFile(
							typeof(CustomerSearch.SearchForCustomerOrProspectResponseType),
							"omAdminCSFindCustomerResponse.xml");

					// Delay for a given time (read from the registry)
					BlockThread();
				}
				else
				{
					// Call the gateway
					gatewayResponse = soapService.execute(gatewayRequest);
				}

				// RFairlie	19/06/2006 MAR1911 End
			}
			catch (Exception exp)
			{
				throw new OmigaException("An error occurred calling the Direct Gateway", exp);
			}

			string responseString = directGatewayHelper.GetXmlFromGatewayObject(typeof(CustomerSearch.SearchForCustomerOrProspectResponseType), gatewayResponse);
		
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(functionName + " Direct Gateway SearchForCustomerOrProspect Response : " + responseString );
			}

			XmlDocumentEx responseDoc = new XmlDocumentEx();
		
			try
			{
				// Parse the response
				responseDoc.LoadXml(responseString);
			}
			catch (XmlException exp)
			{
				throw new OmigaException("The Direct Gateway returned invalid xml", exp);
			}

			try
			{
				return TransformNode(responseDoc, "SearchForCustomerOrProspect.xslt", "AdminToOmigaTranslation.xslt");
			}
			catch (Exception exp)
			{
				throw new OmigaException("An error occurred processing the Direct Gateway response", exp);
			}
		}

		private string GetCustomerDetails(XmlElementEx requestElement)
		{
			const string functionName = "GetCustomerDetails";

			GetCustomer.GetCustomerOrProspectDetailsRequestType gatewayRequest = new GetCustomer.GetCustomerOrProspectDetailsRequestType();
			GetCustomer.GetCustomerOrProspectDetailsResponseType gatewayResponse = new GetCustomer.GetCustomerOrProspectDetailsResponseType();
			GetCustomer.DirectGatewaySoapService soapService = new GetCustomer.DirectGatewaySoapService(); 
			omDG.DirectGatewayBO directGatewayHelper = new omDG.DirectGatewayBO();
			
			try
			{
				// Create the request for the direct gateway
				soapService.Url = directGatewayHelper.GetServiceUrl();
				soapService.Proxy = directGatewayHelper.GetProxy();

				gatewayRequest.ClientDevice = directGatewayHelper.GetClientDevice();
				gatewayRequest.ServiceName = directGatewayHelper.GetServiceName();

				string userId = (requestElement != null) ? requestElement.GetAttribute("USERID"): String.Empty;
				string password = (requestElement != null) ? requestElement.GetAttribute("PASSWORD") : String.Empty;

				if (userId.Length > 0 && password.Length > 0)
				{
					gatewayRequest.TellerID = userId;
					gatewayRequest.TellerPwd = password;
				}
	
				gatewayRequest.ProxyID = directGatewayHelper.GetProxyId();
				gatewayRequest.ProxyPwd = directGatewayHelper.GetProxyPwd();

				gatewayRequest.Operator = directGatewayHelper.GetOperator();	//MAR852 GHun

				gatewayRequest.ProductType = directGatewayHelper.GetProductType();
				gatewayRequest.CommunicationChannel = directGatewayHelper.GetCommunicationChannel((userId.Length > 0));
			}
			catch (Exception exp)
			{
				throw new OmigaException("An error occurred setting up the common gateway attributes", exp);
			}

			XmlDocumentEx tempDoc = new XmlDocumentEx();
			XmlDocumentEx responseDoc = new XmlDocumentEx();
			XmlElementEx responseElement = (XmlElementEx)tempDoc.CreateElement("RESPONSE");
			tempDoc.AppendChild (responseElement);
			XmlNodeList customersIn = null;

			try
			{
				customersIn = requestElement.SelectMandatoryNodes("CUSTOMER");
			}
			catch (XmlException exp)
			{
				throw new OmigaException("An error occurred processing the request", exp);
			}

			foreach (XmlElementEx customerElement in customersIn)
			{
				// RFairlie	19/06/2006 MAR1911 Start - Configurably stub out calls to DG
 
				//gatewayRequest.CustomerNumber = customerElement.GetMandatoryAttribute("CUSTOMERNUMBER");
				string customerNumber = customerElement.GetMandatoryAttribute("CUSTOMERNUMBER");
				gatewayRequest.CustomerNumber = customerNumber;
				
				// RFairlie	19/06/2006 MAR1911 End

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = directGatewayHelper.GetXmlFromGatewayObject(typeof(GetCustomer.GetCustomerOrProspectDetailsRequestType), gatewayRequest);
					_logger.Debug(functionName + " Calling Direct Gateway GetCustomerOrProspectDetails. Request: "   + requestString);
				}
				
				try
				{
					// RFairlie	19/06/2006 MAR1911 Start - Configurably stub out calls to DG 

					bool useStub = UseStub();

					if (useStub == true)
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(functionName + ": Using Direct Gateway stub." );
						}

						// create dummy response
						gatewayResponse = 
							(GetCustomer.GetCustomerOrProspectDetailsResponseType)LoadGatewayObjectFromFile(
								typeof(GetCustomer.GetCustomerOrProspectDetailsResponseType),
								"omAdminCSGetCustomerDetailsResponse.xml");

						// use customer number from request in response
						gatewayResponse.CustomerNumber = customerNumber;
						gatewayResponse.ClientDetails.Number = customerNumber;

						// Delay for a given time (read from the registry)
						BlockThread();
					}
					else
					{
						// Call the gateway
						gatewayResponse = soapService.execute(gatewayRequest);
					}
				}
				catch (Exception exp)
				{
					throw new OmigaException("An error occurred calling the Direct Gateway", exp);
				}

				// RFairlie	19/06/2006 MAR1911 End

				string responseString = directGatewayHelper.GetXmlFromGatewayObject(typeof(GetCustomer.GetCustomerOrProspectDetailsResponseType), gatewayResponse);
	
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + " Direct Gateway GetCustomerOrProspectDetails Response : " + responseString );
				}

				try
				{
					if (gatewayResponse.ErrorCode.CompareTo("0") != 0)
					{
						throw new OmigaException(gatewayResponse.ErrorMessage);
					}

					responseDoc.LoadXml(responseString);
					responseElement.AppendChild(tempDoc.ImportNode(responseDoc.DocumentElement, true));
				}
				catch (XmlException exp)
				{
					throw new OmigaException("The Direct Gateway returned invalid xml", exp);
				}
			}

			try
			{
				return TransformNode(responseElement, "GetCustomerOrProspectDetails.xslt", "AdminToOmigaTranslation.xslt");
			}
			catch (Exception exp)
			{
				throw new OmigaException("An error occurred processing the Direct Gateway response", exp);
			}
		}

		private string UpdateCRSCustomer(XmlElementEx requestElement)
		{
			const string functionName = "UpdateCRSCustomer";

			UpdateCustomer.UpdateCustomerOrProspectRequestType gatewayRequest = new UpdateCustomer.UpdateCustomerOrProspectRequestType();
			UpdateCustomer.UpdateCustomerOrProspectResponseType gatewayResponse = new UpdateCustomer.UpdateCustomerOrProspectResponseType();
			UpdateCustomer.DirectGatewaySoapService soapService = new UpdateCustomer.DirectGatewaySoapService(); 
			omDG.DirectGatewayBO directGatewayHelper = new omDG.DirectGatewayBO();
			
			// Create the request for the direct gateway
			soapService.Url = directGatewayHelper.GetServiceUrl();

			//IWebProxy proxy = new WebProxy("LOCALHOST:8080", true);
			//proxy.Credentials = CredentialCache.DefaultCredentials;
			//soapService.Proxy = proxy;
			soapService.Proxy = directGatewayHelper.GetProxy();

			string transformedRequest = string.Empty;
			XmlDocumentEx transformedDoc = new XmlDocumentEx();
			XmlDocumentEx responseDoc = new XmlDocumentEx();
			XmlDocumentEx tempDoc = new XmlDocumentEx();
			XmlElementEx responseElement = null;
			XmlElementEx customerList = null;
			XmlNodeList transformedCustomers = null;

			try
			{
				transformedRequest = TransformNode(requestElement, "UpdateCRSCustomerRequest.xslt", "UpdateCRSCustomerTranslation.xslt");
				transformedDoc.LoadXml(transformedRequest);
				responseElement = (XmlElementEx)tempDoc.CreateElement("RESPONSE");
				responseElement.SetAttribute("TYPE", "SUCCESS");
				tempDoc.AppendChild (responseElement);
				customerList = (XmlElementEx)tempDoc.CreateElement("CUSTOMERLIST");
				responseElement.AppendChild(customerList);
				transformedCustomers = transformedDoc.SelectMandatoryNodes("REQUEST/UpdateCustomerOrProspectRequestType");
			}
			catch (XmlException exp)
			{
				throw new OmigaException("An error occurred processing the request", exp);
			}

			foreach (XmlElementEx customerElement in transformedCustomers)
			{
				XmlNodeReader thisReader = new XmlNodeReader(customerElement);
				XmlSerializer thisSerialiser = new XmlSerializer(typeof(UpdateCustomer.UpdateCustomerOrProspectRequestType));
				gatewayRequest = (UpdateCustomer.UpdateCustomerOrProspectRequestType)thisSerialiser.Deserialize(thisReader);
				gatewayRequest.ClientDevice = directGatewayHelper.GetClientDevice();
				gatewayRequest.ServiceName = directGatewayHelper.GetServiceName();
	
				string userId = (requestElement != null) ? requestElement.GetAttribute("USERID"): String.Empty;
				string password = (requestElement != null) ? requestElement.GetAttribute("PASSWORD") : String.Empty;

				if (userId.Length > 0 && password.Length > 0)
				{
					gatewayRequest.TellerID = userId;
					gatewayRequest.TellerPwd = password;
				}
		
				gatewayRequest.ProxyID = directGatewayHelper.GetProxyId();
				gatewayRequest.ProxyPwd = directGatewayHelper.GetProxyPwd();

				gatewayRequest.Operator = directGatewayHelper.GetOperator();	//MAR852 GHun

				gatewayRequest.ProductType = directGatewayHelper.GetProductType();
				gatewayRequest.CommunicationChannel = directGatewayHelper.GetCommunicationChannel((userId.Length > 0));

				if(_logger.IsDebugEnabled) 
				{
					// PSC 14/11/2005 MAR352
					string requestString  = directGatewayHelper.GetXmlFromGatewayObject(typeof(UpdateCustomer.UpdateCustomerOrProspectRequestType), gatewayRequest);
					_logger.Debug(functionName + " Calling Direct Gateway UpdateCustomerOrProspect. Request: "   + requestString);
				}
				
				try
				{
					// RFairlie	19/06/2006 MAR1911 Start - Configurably stub out calls to DG 

					bool useStub = UseStub();

					if (useStub == true)
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(functionName + ": Using Direct Gateway stub." );
						}

						// create dummy response
						gatewayResponse = 
							(UpdateCustomer.UpdateCustomerOrProspectResponseType)LoadGatewayObjectFromFile(
							typeof(UpdateCustomer.UpdateCustomerOrProspectResponseType),
							"omAdminCSUpdateCRSCustomerResponse.xml");

						// Delay for a given time (read from the registry)
						BlockThread();
					}
					else
					{
						// Call the gateway
						gatewayResponse = soapService.execute(gatewayRequest);
					}
				}
				catch (Exception exp)
				{
					throw new OmigaException("An error occurred calling the Direct Gateway", exp);
				}

				// RFairlie	19/06/2006 MAR1911 End

				// PSC 14/11/2005 MAR352
				string responseString = directGatewayHelper.GetXmlFromGatewayObject(typeof(UpdateCustomer.UpdateCustomerOrProspectResponseType), gatewayResponse);
	
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + " Direct Gateway UpdateCustomerOrProspect Response : " + responseString );
				}

				try
				{
					if (gatewayResponse.ErrorCode.CompareTo("0") != 0)
					{
						throw new OmigaException(gatewayResponse.ErrorMessage);
					}					
				}
				catch (XmlException exp)
				{
					throw new OmigaException("The Direct Gateway returned invalid xml", exp);
				}

				try
				{
					// PSC 17/11/2005 MAR618
					string CifNumber = gatewayResponse.CIF;
					// PSC 16/11/2005 MAR591
					string omigaNumber = customerElement.GetNodeText("ClientDetail/OMIGACUSTOMERNUMBER");
					XmlElementEx returnCustomer = (XmlElementEx) tempDoc.CreateElement("CUSTOMER");
					returnCustomer.SetAttribute("OTHERSYSTEMCUSTOMERNUMBER", CifNumber);
					returnCustomer.SetAttribute("OMIGACUSTOMERNUMBER", omigaNumber);
					customerList.AppendChild(returnCustomer);
				}
				catch (Exception exp)
				{
					throw new OmigaException("An error occurred processing the Direct Gateway response", exp);
				}
			}

			return tempDoc.OuterXml;
		}


		private string GetXsltPath()
		{
			string XsltPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			return XsltPath.ToUpper().Replace("\\DLLDOTNET", "\\XML");
		}
	

		private string TransformNode(XmlNode sourceNode, string transformFileName)
		{
			return TransformNode(sourceNode, transformFileName, string.Empty);  
		}
		
		private string TransformNode(XmlNode sourceNode, string firstTransformFileName, string secondTransformFileName)
		{
			XslTransform thisTransform = new XslTransform();
			StringWriter thisStringWriter = new StringWriter();
			thisTransform.Load(GetXsltPath() + "\\" + firstTransformFileName);
			thisTransform.Transform(sourceNode, null, thisStringWriter, null);

			if (secondTransformFileName != null && secondTransformFileName.Length > 0)
			{
				XmlDocumentEx transformedResponse = new XmlDocumentEx();
				transformedResponse.LoadXml(thisStringWriter.ToString());

				// Convert the document to use the correct combo values
				thisStringWriter = new StringWriter();
				thisTransform.Load(GetXsltPath() + "\\" + secondTransformFileName);
				thisTransform.Transform(transformedResponse, null, thisStringWriter, null);
			}
			return thisStringWriter.ToString();
		}
		// PSC 03/11/2005 MAR352 - End
		// PSC 24/11/2005 MAR673 - Start
		private void SendCompletionsResponse(XmlElement requestElement)
		{
			string functionName = "SendCompletionsResponse";

			GlobalParameter adminQueue = new GlobalParameter("CompletionsInboundQueueName");

			XmlDocumentEx messageDocument = new XmlDocumentEx();
			XmlElementEx rootElement = (XmlElementEx) messageDocument.CreateElement("REQUEST");
			rootElement.SetAttribute("OPERATION", "SendToQueue");
			messageDocument.AppendChild(rootElement);

			XmlElementEx messageElement = (XmlElementEx) messageDocument.CreateElement("MESSAGEQUEUE");
			messageElement.SetAttribute("QUEUENAME", adminQueue.String);
			messageElement.SetAttribute("PROGID", "omPayProc.PaymentProcessingBO");
			rootElement.AppendChild(messageElement);

			// Set up data element
			XmlElementEx dataElement = (XmlElementEx) messageDocument.CreateElement("XML");
			messageElement.AppendChild(dataElement);

			dataElement.AppendChild(messageDocument.ImportNode(requestElement, true));

			// Send response back to the queue for processing by payment processing
			omMQBOClass omigaMessageQueue = new omMQBOClass();
			
			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(functionName + " SendToQueue Request: " + messageDocument.OuterXml);
			}

			// MAR1603 M Heys 10/04/2006 start			
			Omiga4Response sendResponse;
			try
			{
				//Omiga4Response sendResponse = new Omiga4Response(omigaMessageQueue.omRequest(messageDocument.OuterXml));
				sendResponse = new Omiga4Response(omigaMessageQueue.omRequest(messageDocument.OuterXml));
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(omigaMessageQueue);
			}
			// MAR1603 M Heys 10/04/2006 end

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(functionName + " SendToQueue Response: " + sendResponse.Xml);
			}

			if (!sendResponse.isSuccess())
			{
				throw new OmigaException("Error sending completions response to message queue. " + sendResponse.Xml);
			}
		}

		private XmlElementEx BuildResponseHeader (AdminInterfaceBO_OnMessage_Input convertedInput)
		{
			XmlDocumentEx tempDoc = new XmlDocumentEx();
			XmlElementEx headerElement = (XmlElementEx) tempDoc.CreateElement("HEADER");
			headerElement.SetAttribute("BATCHRUNNUMBER", convertedInput.BATCHAUDIT.BATCHRUNNUMBER);
			headerElement.SetAttribute("BATCHNUMBER", convertedInput.BATCHAUDIT.BATCHNUMBER);
			headerElement.SetAttribute("BATCHAUDITGUID", convertedInput.BATCHAUDIT.BATCHAUDITGUID);
			headerElement.SetAttribute("USERID", convertedInput.USERID);
			headerElement.SetAttribute("UNITID", convertedInput.UNITID);
			headerElement.SetAttribute("APPLICATIONNUMBER", convertedInput.PAYMENTRECORD.APPLICATIONNUMBER);
			headerElement.SetAttribute("PAYMENTSEQUENCENUMBER", convertedInput.PAYMENTRECORD.PAYMENTSEQUENCENUMBER);
		
			return headerElement;
		}

		private XmlElementEx BuildResponseHeader (string requestXml)
		{
			XmlDocumentEx tempDoc = new XmlDocumentEx();
			tempDoc.LoadXml(requestXml);

			XmlElementEx request = (XmlElementEx)tempDoc.SelectMandatorySingleNode("REQUEST");
			XmlElementEx batchAudit = (XmlElementEx)request.SelectMandatorySingleNode("BATCHAUDIT");
			XmlElementEx paymentRecord = (XmlElementEx)request.SelectMandatorySingleNode("PAYMENTRECORD");

			XmlElementEx headerElement = (XmlElementEx) tempDoc.CreateElement("HEADER");
			headerElement.SetAttribute("BATCHRUNNUMBER", batchAudit.GetMandatoryAttribute("BATCHRUNNUMBER"));
			headerElement.SetAttribute("BATCHNUMBER", batchAudit.GetMandatoryAttribute("BATCHNUMBER"));
			headerElement.SetAttribute("BATCHAUDITGUID", batchAudit.GetMandatoryAttribute("BATCHAUDITGUID"));
			headerElement.SetAttribute("USERID", request.GetMandatoryAttribute("USERID"));
			headerElement.SetAttribute("UNITID", request.GetMandatoryAttribute("UNITID"));
			headerElement.SetAttribute("APPLICATIONNUMBER", paymentRecord.GetMandatoryAttribute("APPLICATIONNUMBER"));
			headerElement.SetAttribute("PAYMENTSEQUENCENUMBER", paymentRecord.GetMandatoryAttribute("PAYMENTSEQUENCENUMBER"));
		
			return headerElement;
		}

		private XmlElementEx BuildExceptionResponse(string requestXml, OmigaException exp) 
		{
			// Set up request for the queue
			XmlDocumentEx messageDocument = new XmlDocumentEx();

			XmlElementEx requestElement = (XmlElementEx) messageDocument.CreateElement("REQUEST");
			requestElement.SetAttribute("OPERATION", "CompleteInterfacing");
			XmlElementEx headerElement = BuildResponseHeader(requestXml);
			headerElement = (XmlElementEx)messageDocument.ImportNode(headerElement, true);
			requestElement.AppendChild(headerElement);

			// PSC 30/11/2005 MAR733 - Start
			XmlElementEx responseElement = (XmlElementEx) messageDocument.CreateElement("RESPONSE");
			requestElement.AppendChild(responseElement);
			responseElement.SetAttribute("TYPE", "APPERR");
			XmlElementEx errorElement = (XmlElementEx) messageDocument.CreateElement("ERROR");
			responseElement.AppendChild(errorElement);
			XmlElementEx tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORNUMBER");	
			tempElement.InnerText =  "1000";
			errorElement.AppendChild(tempElement);
			tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORDESCRIPTION");
			tempElement.InnerText = exp.Message;
			errorElement.AppendChild(tempElement);
			tempElement = (XmlElementEx) messageDocument.CreateElement("ERRORSOURCE");
			tempElement.InnerText = exp.ErrorSource;
			errorElement.AppendChild(tempElement);
			return requestElement;
			// PSC 30/11/2005 MAR733 - End
		}
		// PSC 24/11/2005 MAR673 - End


		#region Xml Helper Functions

//		void XmlInsertElementBeforeFirstChild(XmlDocument doc, XmlElement parent, string qualifiedName, string namespaceUri, string value)
//		{
//			XmlElement element = doc.CreateElement(qualifiedName, namespaceUri);
//			element.InnerText = value;
//			parent.InsertBefore(element, parent.FirstChild);
//		}

//		void XmlInsertNilElementBeforeFirstChild(XmlDocument doc, XmlElement parent, string qualifiedName, string namespaceUri)
//		{
//			XmlElement element = doc.CreateElement(qualifiedName, namespaceUri);
//			element.SetAttribute("nil", "http://www.w3.org/2001/XMLSchema-instance", null);
//			parent.InsertBefore(element, parent.FirstChild);
//		}

		#endregion

		#region Stub Helper Functions

		/// <summary>
		/// Determine whether to stub out call to DG
		/// </summary>
		/// <remarks>
		/// Never raises exceptions.
		/// Added for MAR1911. 
		/// </remarks>
		private bool UseStub()
		{
			string functionName = "UseStub";

			bool useStub = false;

			string globalparam = "UseDummyAdminSystem"; 
			try
			{
				useStub = new GlobalParameter(globalparam).Boolean;
			}
			catch(Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + 
						": Failed to read global parameter " + globalparam + 
						", defaulting value to false", ex);
				}
				useStub = false;
			}
			return useStub;
		}

		/// <summary>
		/// Block the current thread for a configurable number of milliseconds
		/// </summary>
		/// <remarks>
		/// Never raises exceptions.
		/// Added for MAR1911. 
		/// </remarks>
		private void BlockThread()
		{
			string functionName = "BlockThread";

			try
			{
				// get configured delay
				string sleepTime = GetStubDelay();

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + ": Sleeping for " + sleepTime + "ms");
				}

				System.Threading.Thread.Sleep(Convert.ToInt32(sleepTime));
			}
			catch(Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("Error in " + functionName, ex);
				}
			}
		}

		/// <summary>
		/// Read delay in milliseconds from registry. Default value = 1000
		/// </summary>
		/// <remarks>
		/// Never raises exceptions.
		/// Added for MAR1911. 
		/// </remarks>
		private string GetStubDelay()
		{	
			string functionName = "GetStubDelay";

			string delay = "1000";

			try
			{
				Microsoft.Win32.RegistryKey key =
					Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
					"Software\\Omiga4\\System Configuration\\Stubs\\omAdminCS", 
					false);
					
				if (key != null)
				{
					delay = (string)key.GetValue("Delay");
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("Error in " + functionName, ex);
				}
			}

			return delay; 
		}

		/// <summary>
		/// Create DG object by reading xml file located in the runtime directory.
		/// </summary>
		/// <remarks>
		/// Added for MAR1911. 
		/// This method could be moved to omDG, but it was added for a patch.
		/// </remarks>
		private object LoadGatewayObjectFromFile(Type objectType, string fileName)
		{
			object dgObj = null;
			XmlSerializer ser = null;
			XmlReader reader = null;

			try
			{
				ser = new XmlSerializer(objectType);
				string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
				reader = new XmlTextReader(appPath + "//" + fileName);
				dgObj = ser.Deserialize(reader);
			}
			catch (Exception ex)
			{
				throw new OmigaException("Unable to load Gateway object from file.", ex);
			}
			finally
			{
				reader.Close();
			}
			
			return dgObj;
		}

		#endregion
	}
}









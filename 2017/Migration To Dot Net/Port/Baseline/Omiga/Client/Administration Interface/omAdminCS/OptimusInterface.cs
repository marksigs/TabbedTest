/*--------------------------------------------------------------------------------------------
Workfile:		OptimusInterface.cs
Copyright:		Copyright © 2007 Marlborough Stirling

Description:	Implementation of administration interface for Optimus		
--------------------------------------------------------------------------------------------
History:

Prog		Date		Description
PSC			11/01/2007	EP2_741 Created
PSC			16/01/2006	EP2_741 Don't inherit from BaseAdministrationInterface
						Change GetCustomerDetails to cater for multiple customers
						Change GetAccountDetails to cater for multiple customers
PSC			25/01/2007	EP2_928 remove duplicate accounts
PSC			26/01/2007	EP2_1028 Correct GetXsltPath
PSC			01/02/2007	EP2_1175 Check response from ODI call
--------------------------------------------------------------------------------------------*/
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.IO;
using System.Text;
using System.Net;
using System.Collections;
using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omLogging;
using Vertex.Fsd.Omiga.omDg;
using MESSAGEQUEUECOMPONENTVCLib;

namespace Vertex.Fsd.Omiga.omAdminCS
{
	/// <summary>
	/// Interface to enable communication with Optimus
	/// </summary>
	[ProgId("omAdminCS.OptimusInterface")]
	[Guid("55287D85-6B06-4528-A7AA-85333371689F")]
	[ComVisible(true)]
	public class OptimusInterface : IMessageQueueComponentVC2
	{
		private  omLogging.Logger _logger = null;
		Hashtable XsltCache = Hashtable.Synchronized(new Hashtable());

		#region Constructors and Destructors

		public OptimusInterface()
		{
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, "Default");
		}

		public OptimusInterface(string processContext)
		{
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, processContext);
		}

		#endregion
        		
		#region Public Methods

		/// <summary>
		/// Executes a call to an external administration system. The call that is made is based
		/// on the operation specified in the request
		/// </summary>
		/// <param name="serviceRequest">Xml string containing the request for the service</param>
		/// <returns>The response from the service</returns>
		public virtual string ExecuteService(string serviceRequest)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			
			string response = null;
			bool createdThreadContext = false;
			XmlTextReader reader = null;

			// Create RequestId logging context if not present already
			if (ThreadContext.Properties["RequestId"] == null || 
				ThreadContext.Properties["RequestId"].ToString().Length == 0)
			{
				ThreadContext.Properties["RequestId"] = Guid.NewGuid().ToString("N");
				createdThreadContext = true;
			}

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// Find the OPERATION and TYPE attributes - these determine how the message is processed.
				reader = new XmlTextReader(new StringReader(serviceRequest));

				if (reader.Read() && reader.NodeType == XmlNodeType.Element && reader.Name == "REQUEST")
				{
					string operation = null;

					if (reader.MoveToAttribute("OPERATION")) 
					{
						operation = reader.Value;
					}
					else 
					{
						throw new OmigaException("Invalid request: missing OPERATION attribute.");
					}

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Operation: {1}. Request: {2}", functionName, operation, serviceRequest);
					}

					switch (operation.ToUpper())
					{
						case "FINDCUSTOMER":
							response = FindCustomer(serviceRequest);
							break;
						case "GETCUSTOMERDETAILS":
							response = GetCustomerDetails(serviceRequest);
							break;
						case "FINDBUSINESSFORCUSTOMER":
							response = FindBusinessForCustomer(serviceRequest);
							break;
						case "GETACCOUNTCUSTOMERS":
							response = GetAccountCustomers(serviceRequest);
							break;
						case "GETACCOUNTDETAILS":
							response = GetAccountDetails(serviceRequest);
							break;
						case "GETACCOUNTSUMMARY":
							response = GetAccountSummary(serviceRequest);
							break;
						case "GETNEWNUMBERS":
							response = GetNewNumbers(serviceRequest);
							break;
						default:
							OmigaException omEx = new OmigaException("Invalid OptimusInterface.ExecuteService operation: " + operation + ".");
					
							if (_logger.IsErrorEnabled)
							{
								_logger.Error(functionName + ": " + omEx.Message, omEx);
							}
							throw omEx;
					}
					
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Response: {1}.", functionName, response);
					}
					
					return response;
				}
				else
				{
					throw new OmigaException("Invalid request: missing REQUEST root element.");
				}
			}
			catch (OmigaException omEx)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": Error processing request.", omEx);
				}

				return omEx.ToOmiga4Response();
			}
			catch (Exception ex)
			{
				OmigaException omEx = new OmigaException("Error processing request.", ex);
				
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(omEx.Message, omEx);
				}

				return omEx.ToOmiga4Response();
			}
			finally
			{
				if (reader != null)
				{
					reader.Close();
				}
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}

				// Remove RequestId logging context if we created it
				if (createdThreadContext)
				{
					ThreadContext.Properties.Remove("RequestId");
				}
			}
		}

		#endregion

		#region IMessageQueueComponentVC2 Interface Implementation

		/// <summary>
		/// Processes an MQL message
		/// </summary>
		/// <param name="in_xmlConfig">MQL configuration data</param>
		/// <param name="in_xmlData">MQL message data</param>
		/// <returns>An MQL response to denote successful processing or otherwise</returns>
		int IMessageQueueComponentVC2.OnMessage(string in_xmlConfig, string in_xmlData)
		{
			// Call the derived class's OnMessage method
			return ProcessMQLMessage(in_xmlConfig, in_xmlData);
		}

		#endregion

		#region Protected Methods

		/// <summary>
		/// Finds customers on the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer search criteria</param>
		/// <returns>Xml string containing a list of matching customers</returns>
		protected virtual string FindCustomer(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				//return ProcessServiceRequest(request, "FindCustomerRequest.xslt", "FindCustomerResponse.xslt");
				return ProcessServiceRequestViaODI(request);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets details of the supplied customer from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifier</param>
		/// <returns>Xml string containing the details for the supplied customer</returns>
		protected virtual string GetCustomerDetails(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// Load request
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				// Create response
				XmlDocumentEx responseDoc = new XmlDocumentEx();
				XmlElementEx responseElement = (XmlElementEx)responseDoc.CreateElement("RESPONSE");
				responseElement.SetAttribute("TYPE", "SUCCESS");
				responseDoc.AppendChild(responseElement);

				// Create temporary request and response documents for Optimus call
				XmlDocumentEx tempRequestDoc = new XmlDocumentEx();
				XmlDocumentEx tempResponseDoc = new XmlDocumentEx();

				// Create template request for Optimus
				XmlElementEx requestElement = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx tempRequestElement = (XmlElementEx)tempRequestDoc.ImportNode(requestElement, false);
				tempRequestDoc.AppendChild(tempRequestElement);
				XmlElementEx customerElement = (XmlElementEx)tempRequestDoc.CreateElement("CUSTOMER");
				tempRequestElement.AppendChild(customerElement);

				// Get a list of customers from  the request passed in
				XmlNodeList customersIn = requestElement.SelectMandatoryNodes("CUSTOMER");

				// For each customer call Optimus to get their details
				foreach (XmlElementEx customerIn in customersIn)
				{
					// Set the correct customer number for this customer
					customerElement.SetAttribute("CUSTOMERNUMBER", customerIn.GetMandatoryAttribute("CUSTOMERNUMBER"));
					
					// Call Optimus
					string response = ProcessServiceRequestViaODI(tempRequestDoc.OuterXml);				
					
					//Load the response. If an error is returned ProcessServiceRequestViaODI will
					// have thrown an exception so can assume call was successful
					tempResponseDoc.LoadXml(response);

					// Get the returned customer details and attach them to the response 
					XmlElementEx customerDetails = (XmlElementEx)tempResponseDoc.SelectMandatorySingleNode("RESPONSE/CUSTOMER");
					customerDetails = (XmlElementEx)responseDoc.ImportNode(customerDetails, true);
					responseElement.AppendChild(customerDetails);
				}

				// Return the details for all requested customers
				return responseDoc.OuterXml;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Finds existing business for the supplied customer from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifier</param>
		/// <returns>Xml string containing a list of existing business for the supplied customer</returns>
		protected virtual string FindBusinessForCustomer(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				//return ProcessServiceRequest(request, "FindBusinessForCustomerRequest.xslt", "FindBusinessForCustomerResponse.xslt");
				return ProcessServiceRequestViaODI(request);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Finds customers linked to the supplied account from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the account identifier</param>
		/// <returns>Xml string containing a list of customers linked to the supplied account</returns>
		protected virtual string GetAccountCustomers(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				//return ProcessServiceRequest(request, "GetAccountCustomersRequest.xslt", "GetAccountCustomersResponse.xslt");
				return ProcessServiceRequestViaODI(request);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}
		
		/// <summary>
		/// Gets the account details for the supplied customers from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifiers</param>
		/// <returns>Xml string containing the account details for the supplied customers</returns>
		protected virtual string GetAccountDetails(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// Load request
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				// Create response
				XmlDocumentEx responseDoc = new XmlDocumentEx();
				XmlElementEx responseElement = (XmlElementEx)responseDoc.CreateElement("RESPONSE");
				responseElement.SetAttribute("TYPE", "SUCCESS");
				responseDoc.AppendChild(responseElement);
				XmlElement responseAccountList = responseDoc.CreateElement("MORTGAGEACCOUNTLIST");
				responseElement.AppendChild(responseAccountList);

				// Create temporary request and response documents for Optimus call
				XmlDocumentEx tempRequestDoc = new XmlDocumentEx();
				XmlDocumentEx tempResponseDoc = new XmlDocumentEx();

				// Create template request for Optimus
				XmlElementEx requestElement = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx tempRequestElement = (XmlElementEx)tempRequestDoc.ImportNode(requestElement,false);
				tempRequestDoc.AppendChild(tempRequestElement);
				XmlElementEx customerElement = (XmlElementEx)tempRequestDoc.CreateElement("CUSTOMER");
				tempRequestElement.AppendChild(customerElement);

				// Get a list of customers from  the request passed in
				XmlNodeList customersIn = requestElement.SelectMandatoryNodes("IMPORTACCOUNTSINTOAPPLICATION/CUSTOMERLIST/CUSTOMER");

				// For each customer call Optimus to get their account details
				foreach (XmlElementEx customerIn in customersIn)
				{
					// Set the correct customer number for this customer
					customerElement.SetAttribute("CUSTOMERNUMBER", customerIn.GetMandatoryAttribute("CUSTOMERNUMBER"));
					
					// Call Optimus
					string response = ProcessServiceRequestViaODI(tempRequestDoc.OuterXml);				
					
					//Load the response. If an error is returned ProcessServiceRequestViaODI will
					// have thrown an exception so can assume call was successful
					tempResponseDoc.LoadXml(response);

					// Get the returned account details and attach them to the response 
					XmlNodeList returnedAccounts = tempResponseDoc.SelectNodes("RESPONSE/MORTGAGEACCOUNTLIST/MORTGAGEACCOUNT");

					foreach (XmlElementEx account in returnedAccounts)
					{
						XmlElementEx accountDetails = (XmlElementEx)responseDoc.ImportNode(account, true);
						responseAccountList.AppendChild(accountDetails);
					}
				}

				// Remove customers that aren't in the request and add the Omiga customer number to
				// those that are
				XmlNodeList accountCustomers = responseDoc.SelectNodes(".//CUSTOMER");

				foreach (XmlElementEx accountCustomer in accountCustomers)
				{
					XmlElementEx requestCustomer = (XmlElementEx)requestElement.SelectSingleNode("IMPORTACCOUNTSINTOAPPLICATION/CUSTOMERLIST/CUSTOMER[@CUSTOMERNUMBER='" + accountCustomer.GetMandatoryAttribute("CUSTOMERNUMBER") + "']");

					if (requestCustomer != null)
					{
						accountCustomer.SetAttribute("OMIGACUSTOMERNUMBER", requestCustomer.GetMandatoryAttribute("OMIGACUSTOMERNUMBER"));
					}
					else
					{
						accountCustomer.ParentNode.RemoveChild(accountCustomer);
					}
				}

				// Return the details for all requested customers
				// PSC 25/01/2007 EP2_928
				return Transform(responseDoc.OuterXml, "GetAccountDetailsResponse.xslt");
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets the summary details for the supplied account from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the account identifier</param>
		/// <returns>Xml string containing the details for the supplied account</returns>
		protected virtual string GetAccountSummary(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				//return ProcessServiceRequest(request, "GetAccountSummaryRequest.xslt", "GetAccountSummaryResponse.xslt");
				return ProcessServiceRequestViaODI(request);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets numbers from the external administration system to allocate to applications and customers
		/// to be used at the completion stage
		/// </summary>
		/// <param name="request">Xml string specifiying which type of numbers are required</param>
		/// <returns>Xml string containing the requested numbers</returns>
		protected virtual string GetNewNumbers(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// Create temporary request for Optimus call
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				XmlDocumentEx responseDoc = new XmlDocumentEx();
				XmlDocumentEx tempRequestDoc = new XmlDocumentEx();
				XmlDocumentEx tempResponseDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				XmlElementEx responseElement = (XmlElementEx)responseDoc.CreateElement("RESPONSE");
				responseElement.SetAttribute("TYPE", "SUCCESS");
				responseDoc.AppendChild(responseElement);
								
				XmlElementEx numberRequest = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST/NUMBERREQUEST");
				string numberRequired = numberRequest.GetAttribute("NUMBERREQUIRED");
				
				XmlElementEx requestElement = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx tempRequestElement = (XmlElementEx)tempRequestDoc.ImportNode(requestElement,false);
				tempRequestDoc.AppendChild(tempRequestElement);
				
				if(numberRequired.Length > 0)
				{
					string typeRequired = numberRequest.GetMandatoryAttribute("NUMBERTYPEREQUIRED");
					
					if (typeRequired == "A")
					{
						tempRequestElement.SetAttribute("OPERATION", "GetNewAccountNumber");
					}
					else if(typeRequired == "C")
					{
						tempRequestElement.SetAttribute("OPERATION", "GetNewCustomerNumber");
					}
					string response = ProcessServiceRequestViaODI(tempRequestDoc.OuterXml);	
					tempResponseDoc.LoadXml(response);
				
					XmlElementEx otherSystemNumberElem = (XmlElementEx)responseDoc.CreateElement("OTHERSYSTEMNUMBER");
					responseElement.AppendChild(otherSystemNumberElem);
					
					string otherSystemNumber = tempResponseDoc.SelectMandatorySingleNode("RESPONSE/NUMBERRESPONSE").Attributes.GetNamedItem("OTHERSYSTEMNUMBER").Value;					
					otherSystemNumberElem.SetAttribute ("OTHERSYSTEMNUMBER", otherSystemNumber);
					otherSystemNumberElem.SetAttribute ("NUMBERREQUIRED", numberRequired);
					otherSystemNumberElem.SetAttribute ("NUMBERTYPE", typeRequired);
				}
				return responseDoc.OuterXml;				
				//return ProcessServiceRequest(request, "GetNewNumbersRequest.xslt", "GetNewNumbersResponse.xslt");				
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets a limited set of account details for the supplied account from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the account identifiers</param>
		/// <returns>Xml string containing the account details for the supplied account</returns>
		protected virtual string GetAccountRefresh(string request)
		{
			throw new OmigaException("Method GetAccountRefresh has not been implemented.");	
		}

		/// <summary>
		/// Gets contact log information for the supplied customer from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifier</param>
		/// <returns>Xml string containing the contact log details for the supplied customer</returns>
		protected virtual string GetCRSContactData(string request)
		{
			throw new OmigaException("Method GetCRSContactData has not been implemented.");;
		}

		/// <summary>
		/// Updates the contact log for the supplied customer from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifier and contact log details</param>
		/// <returns>Xml string containing the outcome of the update</returns>
		protected virtual string UpdateCRSContactLog(string request)
		{
			throw new OmigaException("Method UpdateCRSContactLog has not been implemented.");;
		}

		/// <summary>
		/// Updates the details for the supplied customer from the external administration system
		/// </summary>
		/// <param name="request">Xml string containing the customer identifier and their details</param>
		/// <returns>Xml string containing the outcome of the update</returns>
		protected virtual string UpdateCRSCustomer(string request)
		{
			throw new OmigaException("Method UpdateCRSCustomer has not been implemented.");;
		}

		/// <summary>
		/// Process a completion by sending the details to the external system
		/// </summary>
		/// <param name="request">Xml string containing the completion request</param>
		/// <returns>Xml string containing the denoting if the processing was successful</returns>
		protected virtual string ProcessCompletionsInterface(string request)
		{
			throw new OmigaException("Method ProcessCompletionsInterface has not been implemented.");;
		}

		/// <summary>
		/// Does the call to the external system
		/// </summary>
		/// <param name="request">Xml string containing the request in tbe format expected by the external system</param>
		/// <returns>The response from the external system</returns>
		protected virtual string CallExternalInterface(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				MemoryStream memoryStream = new MemoryStream();
				StreamWriter streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);
				XmlTextWriter xmlWriter = new XmlTextWriter(streamWriter);
				xmlWriter.WriteRaw(request);

				DirectGatewayBO dg = new DirectGatewayBO();

				HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(dg.GetServiceUrl());
				WebProxy currentProxy = dg.GetProxy();

				if (currentProxy != null)
				{
					webRequest.Proxy = currentProxy;
				}
				else
				{
					webRequest.Proxy = GlobalProxySelection.GetEmptyWebProxy();
				}

				webRequest.Method = "POST";
				webRequest.ContentType = "text/xml";

				System.Text.Encoding encoding = System.Text.Encoding.UTF8;
				byte[] bytes = memoryStream.GetBuffer();
				int length = (int)memoryStream.Length;
				webRequest.ContentLength = length;

				Stream requestStream = null;

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: external interface request: {1}.", functionName, request);
				}

				try
				{
					requestStream = webRequest.GetRequestStream();
					requestStream.Write(bytes, 0, length);
				}
				finally
				{
					if (requestStream != null)
					{
						requestStream.Close();
					}
				}

				HttpWebResponse OsgResponse = (HttpWebResponse)webRequest.GetResponse();

				Stream responseStream = null;
				StreamReader streamReader = null;
				string response = string.Empty;

				try
				{
					responseStream = OsgResponse.GetResponseStream();
					streamReader = new StreamReader(responseStream, Encoding.UTF8);
					response = streamReader.ReadToEnd();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: external interface response: {1}.", functionName, response);
					}
				}
				finally
				{
					if (responseStream != null)
					{
						responseStream.Close();
					}

					if (streamReader != null)
					{
						streamReader.Close();
					}
				}

				return response;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Processes an MQL message
		/// </summary>
		/// <param name="in_xmlConfig">MQL configuration data</param>
		/// <param name="in_xmlData">MQL message data</param>
		/// <returns>An MQL response to denote successful processing or otherwise</returns>
		protected virtual int ProcessMQLMessage(string in_xmlConfig, string in_xmlData)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			bool createdThreadContext = false;

			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, "MQL");

			// Create RequestId logging context if not present already
			if (ThreadContext.Properties["RequestId"] == null || 
				ThreadContext.Properties["RequestId"].ToString().Length == 0)
			{
				ThreadContext.Properties["RequestId"] = Guid.NewGuid().ToString("N");
				createdThreadContext = true;
			}
			
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				return (int)MESSAGEQUEUECOMPONENTVCLib.MESSQ_RESP.MESSQ_RESP_SUCCESS;
			}
			finally
			{

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}

				// Remove RequestId logging context if we created it
				if (createdThreadContext)
				{
					ThreadContext.Properties.Remove("RequestId");
				}
			}
		}

		/// <summary>
		/// Processes a request by transforming the request from the Omiga format to the
		/// format expected by the external system, calling the external system and then 
		/// transforming the response from the external system format to the Omiga format
		/// </summary>
		/// <param name="request">Xml string containing the request in Omiga format</param>
		/// <param name="preCallTransform">The transform to apply prior to the call to the external system</param>
		/// <param name="postCallTransform">The transform to apply after the call to the external system</param>
		/// <returns>The response from the external system in Omiga format</returns>
		protected string ProcessServiceRequest(string request, string preCallTransform, string postCallTransform)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				string transformedXml = string.Empty;

				if (preCallTransform != null && preCallTransform.Length != 0)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Transforming request.", functionName);
					}

					transformedXml = Transform(request, preCallTransform);
				}
				else
				{
					transformedXml = request;
				}
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling external interface.", functionName);
				}

				string response = CallExternalInterface(transformedXml);
				
				if (postCallTransform != null && postCallTransform.Length != 0)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Transforming response.", functionName);
					}

					transformedXml = Transform(response, postCallTransform);
				}
				else
				{
					transformedXml = response;
				}

				return transformedXml;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Processes the service request using ODI rather than using OSG2
		/// </summary>
		/// <param name="request"></param>
		/// <returns></returns>
		protected string ProcessServiceRequestViaODI(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			string adminSystemState = string.Empty;
			object transformerBO = null;
			Omiga4Response OdiResponse = null;
			Type OdiType = null;

			try
			{
				// Get Optimus login parameters
				GlobalParameter userEnvironment = new GlobalParameter("ODIAutoLogonUserEnv");
				GlobalParameter host = new GlobalParameter("ODIAutoLogonHost");
				GlobalParameter userName = new GlobalParameter("ODIAutoLogonUserName");
				GlobalParameter password = new GlobalParameter("ODIAutoLogonPassword");
				GlobalParameter OdiEnvironment = new GlobalParameter("ODIAutoLogonODIEnv");

				// Create login request
				XmlDocument logonRequestDoc = new XmlDocument();
				XmlElement logonRequest = logonRequestDoc.CreateElement("REQUEST");
				logonRequest.SetAttribute("OPERATION", "ValidateUserLogon");
				logonRequest.SetAttribute("CHANNELID", "A1");
				logonRequestDoc.AppendChild(logonRequest);
				XmlElement userElement = logonRequestDoc.CreateElement("USER");
				userElement.SetAttribute("ENVIRONMENT", userEnvironment.String);
				userElement.SetAttribute("HOST", host.String);
				userElement.SetAttribute("USERNAME", userName.String);
				userElement.SetAttribute("PASSWORDVALUE", password.String);
				logonRequest.AppendChild(userElement);
				XmlElement odiInitialisation = logonRequestDoc.CreateElement("ODIINITIALISATION");
				odiInitialisation.SetAttribute("ODIENVIRONMENT", OdiEnvironment.String);
				logonRequest.AppendChild(odiInitialisation);

				// Late bind to ODI so we don't need a reference in the project
				try
				{
					OdiType = Type.GetTypeFromProgID("ODITransformer.ODITransformerBO", true);
				}
				catch (COMException exp)
				{
					throw new OmigaException("Unable to get type from ProgID ODITransformer.ODITransformerBO.", exp);
				}
			
				try
				{
					transformerBO = Activator.CreateInstance(OdiType);
				}
				catch (Exception exp)
				{
					throw new OmigaException("Unable create an instance of ODITransformer.ODITransformerBO.", exp);
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Logging in to Optimus.", functionName);
				}

				string logonResponse = (string)OdiType.InvokeMember("Request", BindingFlags.InvokeMethod, null, transformerBO, new object[] {logonRequestDoc.OuterXml});

				// Determine if there is a successful response. This will throw an exception 
				// if the response is an APPERR or SYSERR
				OdiResponse = new Omiga4Response(logonResponse);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Optimus login successful.", functionName);
				}

				// Get the admin system state from the response
				XmlDocument logonResponseDoc = new XmlDocumentEx();
				logonResponseDoc.LoadXml(logonResponse);
				XmlElement systemStateElement = (XmlElement)logonResponseDoc.SelectSingleNode("RESPONSE/ADMINSYSTEMSTATE");
				
				if (systemStateElement != null)
				{
					adminSystemState = systemStateElement.OuterXml;
				}

				// Load the request and get the operation. Certain operations in ODI do not match
				// the operation names from Omiga. These need to be changed where required.
				// In addition add in the Optimus Admin System State 
				XmlDocumentEx requestDocument = new XmlDocumentEx();
				requestDocument.LoadXml(request);
				XmlElementEx requestElement = (XmlElementEx)requestDocument.SelectMandatorySingleNode("REQUEST");
				string operation = requestElement.GetMandatoryAttribute("OPERATION");

				switch (operation.ToUpper())
				{
					case "FINDCUSTOMER":
						operation = "FindCustomerList";
						break;
					case "GETACCOUNTDETAILS":
						operation = "FindAccountDetails";
						break;
				}

				requestElement.SetAttribute("OPERATION", operation);
				requestElement.SetAttribute("ADMINSYSTEMSTATE", adminSystemState); 
			
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: ODI request: {1}.", functionName, requestDocument.OuterXml);
				}

				string response = (string)OdiType.InvokeMember("Request", BindingFlags.InvokeMethod, null, transformerBO, new object[] {requestDocument.OuterXml});
				OdiResponse = new Omiga4Response(response);	// PSC 01/02/2007 EP2_1175

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: ODI response: {1}.", functionName, response);
				}

				return response;
			}
			// PSC 01/02/2007 EP2_1175 - Start
			catch (OmigaException exp)
			{
				string message = "An error occurred calling Optimus.";

				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": " + message, exp);
				}

				throw new OmigaException(message, exp);	
			}
			// PSC 01/02/2007 EP2_1175 - End
			finally
			{
				// Logoff from Optimus if we have logged on. If this fails write an error to the
				// log and continue
				if (transformerBO != null && adminSystemState.Length > 0)
				{
					try
					{
						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Logging out of Optimus.", functionName);
						}

						XmlDocument logoffRequestDoc = new XmlDocument();
						XmlElement logoffRequest = logoffRequestDoc.CreateElement("REQUEST");
						logoffRequest.SetAttribute("OPERATION", "LogOffUser");
						logoffRequest.SetAttribute("ADMINSYSTEMSTATE", adminSystemState);
						logoffRequestDoc.AppendChild(logoffRequest);
					
						string logoffResponse = (string)OdiType.InvokeMember("Request", BindingFlags.InvokeMethod, null, transformerBO, new object[] {logoffRequestDoc.OuterXml});
						OdiResponse = new Omiga4Response(logoffResponse);

						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Optimus logout successful.", functionName);
						}
					}
					catch (Exception exp)
					{
						if (_logger.IsErrorEnabled)
						{
							// PSC 01/02/2007 EP2_1175	
							_logger.Error(functionName + ": Failed to log out of Optimus.", exp);
						}
					}
				}
		
				// Release the ODI Transformer COM object
				if (transformerBO != null)
				{
					Marshal.ReleaseComObject(transformerBO);
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Transforms the source xml using the Xslt file supplied
		/// </summary>
		/// <param name="sourceXml">Xml string containing the source to be transformed</param>
		/// <param name="transformFilename">The filename of the Xslt to use for the transformation</param>
		/// <returns>Xml string containing the transformed Xml</returns>
		protected string Transform(string sourceXml, string transformFilename)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Source Xml: {1}.", functionName, sourceXml);
				}

				string transformedXml = sourceXml;
				XPathDocument sourceDocument = new XPathDocument(new StringReader(sourceXml));
			
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Getting Xsl transform {1}.", functionName, transformFilename);
				}

				XslTransform xslt = GetTransform(transformFilename);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Got Xsl transform {1}.", functionName, transformFilename);
				}

				StringWriter targetStringWriter = null;
				XmlTextWriter targetXmlWriter = null;

				try
				{
					targetStringWriter = new StringWriter();
					targetXmlWriter = new XmlTextWriter(targetStringWriter);
					targetXmlWriter.Formatting = Formatting.None;

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Transforming Source Xml using {1}.", functionName, transformFilename);
					}

					xslt.Transform(sourceDocument, null, targetXmlWriter, null);
					transformedXml = targetStringWriter.ToString();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Transformed Source Xml. Transformed Xml: {1}", functionName, transformedXml);
					}
					return transformedXml;
				
				}
				finally
				{
					if (targetXmlWriter != null)
					{
						targetXmlWriter.Close();
					}

					if (targetStringWriter != null)
					{
						targetStringWriter.Close();
					}
				}
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets the requested transform. The cache is searched first and if the transform
		/// is not found it is read from the file, compiled and stored in the cache
		/// </summary>
		/// <param name="transformFilename">The filename of the Xslt to use for the transformation</param>
		/// <returns>The requested transform</returns>
		protected XslTransform GetTransform(string transformFilename)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				XslTransform xslt = null;

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Checking cache for Xsl transform {1}.", functionName, transformFilename);
				}

				if (XsltCache.Contains(transformFilename))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Xsl transform {1} found in cache.", functionName, transformFilename);
					}

					xslt = (XslTransform)XsltCache[transformFilename];
				}
				else
				{
					string XsltFilePath = GetXsltPath() + "\\" + transformFilename;
					xslt = new XslTransform();

					try
					{
						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Compiling Xsl transform {1}.", functionName, transformFilename);
						}

						xslt.Load(XsltFilePath);

						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Adding Xsl transform {1} to cache.", functionName, transformFilename);
						}

						XsltCache.Add(transformFilename, xslt);
					}
					catch (System.IO.FileNotFoundException exp)
					{
						throw new OmigaException("Unable to apply transform. Xslt file " + XsltFilePath + " not found." , exp);
					}
				}

				return xslt;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}
		
		/// <summary>
		/// Determines the path where the Xslt files are stored
		/// </summary>
		/// <returns>The xslt file path</returns>
		protected string GetXsltPath()
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// PSC 26/01/2007 EP2_1028 - Start
				string XsltPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
				int position = -1;
				
				if (XsltPath.ToUpper().IndexOf("DLLDOTNET") >= 0)
				{
					position = XsltPath.ToUpper().IndexOf("DLLDOTNET");
				}
				else if (XsltPath.ToUpper().IndexOf("OMIGAWEBSERVICES") >= 0)
				{
					position = XsltPath.ToUpper().IndexOf("OMIGAWEBSERVICES");
				}

				if (position >= 0)
				{
					XsltPath = XsltPath.Substring(0, position) + @"XML";
				}
				// PSC 26/01/2007 EP2_1028 - End
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Xslt file location is {1}.", functionName, XsltPath);
				}

				return(XsltPath);
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		#endregion

	}
}

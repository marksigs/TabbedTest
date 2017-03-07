/*--------------------------------------------------------------------------------------------
Workfile:		FulfilmentDataManager.cs
Copyright:		Copyright © 2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog		Date		Description
PSC			25/11/2005	MAR588: Correct gateway parameters for CancelPack
GHun		07/12/2005	MAR795 Get ObjectStore from a GlobalParameter
GHun		13/12/2005	MAR852 set Operator in generic request
PSC			14/11/2005	MAR802 Amend SendToSMSFulfilment and SendToEMAILFulfilment to use
							   customer node rather than address node for the phone details
							   and contact email address
PSC			04/01/2006 MAR993  Add ClientDevice	and ServiceName to SendToEmailFulfilment and SendToFulfilment						   
PSC			06/01/2006 MAR922  Amend CheckFolderExists to look within parent folder
HMA         11/01/2006 MAR978  Amend PackTypeID and ImageReference in request to Fulfilment
PSC			11/01/2006 MAR994  Amend SaveFile to not save minor versions 
							   Rewrite getting and creation of folders so that the root path
							   is a global parameter and colders are created recursively down
							   th path
DRC         15/01/2006  MAR978 Revisited - sorted out sope missing nodes on the response.							   
DRC         18/01/2006  MAR978 Revisited - adde specialneed to response.							   
PSC			19/01/2006	MAR1087 Changed PackTypeID for SMS and EMail
RF			03/02/2006  MAR1191 Add document name to fulfilment request
PSC			06/02/2006	MAR1197	Add content type to fulfilment request and return content type 
                                from filenet when getting record
RF			21/02/2006	MAR1196 Follow C# logging standards.
								Doc id not always returned as a property.		
DRC         23/02/2006  MAR1193 Country for Customer address and check for Scottish Property												
HMA         02/03/2006  MAR1302 Add Person details to SendToEMAILFulfilment and SendToSMSFulfilment
GHun		03/03/2006	MAR1332	Changed FulfilmentHandler to make eventkeys consistent
RF			23/03/2006  MAR1299 For Fulfilment set CommunicationChannel to PHONE when there is a logged on user and otherwise INTERNAL; 
								set CommunicationDirection to “OUT”; set SessionID to to the Pack Id 
								Follow C# logging standards for Fulfilment helper.
								Check responses from fulfilment.
PE			13/06/2006	MAR1405	Change of document class for storage of document in filenet.
PSC			19/04/2006	MAR1634	Improve logging
HMA         27/04/2006  MAR1638 Change SessionID to numeric representation of PackID.
HMA         03/05/2006	MAR1638 Check that Pack ID is not empty.
PSC			05/05/2006  MAR1593 Amend SendToSMSFulfilment and SendToEmailFulfilment to get PACKID from request
HMA         10/05/2006  MAR1638 Use SeqNextNumber table to generate a number.
HMA                             Error logging.
HMA                             Send Customer Number for Cancel and Resend.
HMA         06/07/2006  MAR1896 Read global parameters and perform dummy calls to FileNet or Fulfilment if necessary.
--------------------------------------------------------------------------------------------*/
using System;
using Microsoft.Web.Services2;
using Microsoft.Web.Services2.Security.Tokens;
using System.Web;
using System.Web.Services;
using Vertex.Fsd.Omiga.omFDM.WebRefFileNet;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Net; // for proxy object
using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omDg;
using omLogging = Vertex.Fsd.Omiga.omLogging;  
using System.Reflection; 
using System.Data;
using System.Data.SqlClient;
using omigaException = Vertex.Fsd.Omiga.Core.OmigaException;


namespace Vertex.Fsd.Omiga.omFDM
{
	[ProgId("omFDM.FileNetInterfaceBO")]
	[ComVisible(true)]
	[Guid("05CA40F0-ACE6-4573-89F2-95F89D71B1FC")]

	public class FileNetInterfaceBO
	{
		public bool traceOn = false;
		public string traceLocation = "";
		public string fileNetParentFolder = "";
		public string fileNetFolderName = "";
		public string fileNetParentFolderId = "";		// PSC 11/01/2006 MAR994
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		[STAThread]
		static void Main() 
		{
			FileNetInterfaceBO app = new FileNetInterfaceBO();
		}

		public FileNetInterfaceBO()
		{
		}

		public string getFileNetRecord(string userName, string password, 
			string applicationNumber, string fileNetGUID)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				FileNetHandler FNHandler = new FileNetHandler(applicationNumber, "", userName, password);
				
				return  FNHandler.GetFileNetRecord(traceOn, traceLocation, fileNetGUID);
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;	
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}
	
		// PSC 11/01/2006 MAR994
		public string createPrimaryFolder(string userName, string password)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				FileNetHandler FNHandler = new FileNetHandler("", "", userName, password);
				
				// PSC 11/01/2006 MAR994
				return  FNHandler.CreateFolder(traceOn, traceLocation, "{0F1E2D3C-4B5A-6978-8796-A5B4C3D2E1F0}", fileNetFolderName);
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		public string getFileNetFolder(string userName, string password)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				FileNetHandler FNHandler = new FileNetHandler("", "", userName, password);
				
				return  FNHandler.GetFileNetFolder(traceOn, traceLocation, fileNetParentFolder, fileNetFolderName);
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// PSC 11/01/2006 MAR994
		public string createFolder(string userName, string password)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				FileNetHandler FNHandler = new FileNetHandler("", "", userName, password);
				
				// PSC 11/01/2006 MAR994
				return  FNHandler.CreateFolder(traceOn, traceLocation, fileNetParentFolderId, fileNetFolderName);
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		public string checkFolderExists(string userName, string password)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				FileNetHandler FNHandler = new FileNetHandler("", "", userName, password);
				
				// PSC 11/01/2006 MAR994
				return FNHandler.CheckFolderExists(traceOn, traceLocation, fileNetParentFolder, fileNetFolderName);
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		public string getFileNetGUID(
			string applicationNumber, 
			string documentData, 
			string userName, 
			string password, 
			string documentName, // MAR1191
			string contentType,	// PSC 06/02/2006 MAR1197
			string documentclass) // PE 13/04/2006 MAR1405
		{
			// RF MAR1196
			omLogging.ThreadContext.Properties["ApplicationNumber"] = applicationNumber;

			if(_logger.IsDebugEnabled) 
			{
				// PSC 06/02/2006 MAR1197
				_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
					": Starting for application " + applicationNumber +
					" and document " + documentName +
					" with content type of " + contentType +
					" and document class " + documentclass);
			}

			try
			{
				//MAR1896 Return dummy guid if UseFileNetStub global parameter is true.
				bool bUseStub;
                string docId;

				try
				{
					bUseStub = new GlobalParameter("UseFileNetStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					// Delay for a given time (read from the registry)
					BlockThread();

					// Return dummy File Net Guid
					Guid myGuid = new Guid();
					myGuid = Guid.NewGuid();
					docId = "{" + myGuid.ToString() + "}";
				}
				else
				{
					// PSC 06/02/2006 MAR1197
					FileNetHandler FNHandler = new FileNetHandler(
						applicationNumber, documentData, userName, password);
				
					// MAR1191 Start
					//return  FNHandler.SaveFileToFolder(traceOn, traceLocation);
					// PSC 06/02/2006 MAR1197
					docId = FNHandler.SaveFileToFolder(documentName, contentType, traceOn, traceLocation, documentclass);
				}

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Doc id from SaveFileToFolder = " + docId);
				}

				return docId;
				// MAR1191 End
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// initial request to fulfilment, i.e. SendToFulfilment
		public string initialRequest(string userName, string password, string inputXML)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}
			try
			{
				//MAR1896 Return success if UseFulfilmentStub global parameter is TRUE
				bool bUseStub;
				string resp;

				try
				{
					bUseStub = new GlobalParameter("UseFulfilmentStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					//Delay for a given time (read from the registry)
					BlockThread();

					resp = "Done";		
				}
				else
				{
					FulfilmentHandler FMHandler = new FulfilmentHandler(userName, password, inputXML, "");
				
					// RF 27/02/2006 MAR1299 Start
					//return FMHandler.SendToFulfilment(traceOn, traceLocation);
					resp = FMHandler.SendToFulfilment();
				}

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name +
						": Response from SendToFulfilment = " + resp); 
				}

				return resp;
				// RF 27/02/2006 MAR1299 End
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		public string sendToSMS(string userName, string password, string inputXML)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				//MAR1896 Return success if UseFulfilmentStub global parameter is TRUE
				bool bUseStub;
				string resp;

				try
				{
					bUseStub = new GlobalParameter("UseFulfilmentStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					//Delay for a given time (read from the registry)
					BlockThread();

					resp = "Done";		
				}
				else
				{
					FulfilmentHandler FMHandler = new FulfilmentHandler(userName, password, inputXML, "");
				
					//return FMHandler.SendToSMSFulfilment(traceOn, traceLocation);
					resp = FMHandler.SendToSMSFulfilment(traceOn, traceLocation);
				}
				
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name +
						": Response from SendToSMSFulfilment = " + resp); 
				}
				return resp;
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		public string sendToEmail(string userName, string password, string inputXML)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				//MAR1896 Return success if UseFulfilmentStub global parameter is TRUE
				bool bUseStub;
				string resp;

				try
				{
					bUseStub = new GlobalParameter("UseFulfilmentStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					//Delay for a given time (read from the registry)
					BlockThread();

					resp = "Done";		
				}
				else
				{
					FulfilmentHandler FMHandler = new FulfilmentHandler(userName, password, inputXML, "");
				
					//return FMHandler.SendToEMAILFulfilment(traceOn, traceLocation);

					resp = FMHandler.SendToEMAILFulfilment(traceOn, traceLocation);
				}
				
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name +
						": Response from SendToEMAILFulfilment = " + resp); 
				}
				return resp;
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// MAR1638 Send customer number in request.
		public string reSendPack(string packID, string userName, string password, string customernumber)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting: Pack id = " + packID);  // MAR1638
			}

			try
			{
				//MAR1896 Return success if UseFulfilmentStub global parameter is TRUE
				bool bUseStub;
				string resp;

				try
				{
					bUseStub = new GlobalParameter("UseFulfilmentStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					//Delay for a given time (read from the registry)
					BlockThread();

					resp = "Done";		
				}
				else
				{
					FulfilmentHandler FMHandler = new FulfilmentHandler(userName, password, "", packID);
				
					// RF 27/02/2006 MAR1299 Start
					//return FMHandler.ReSendPack(traceOn, traceLocation);
					resp = FMHandler.ReSendPack(customernumber);   // MAR1638
				}
				
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name +
						": Response from ReSendPack = " + resp); 
				}
				return resp;
				// RF 27/02/2006 MAR1299 End
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// MAR1638  Pass customer number in the request.
		public string cancelPack(string packID, string userName, string password, string customernumber)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting: Pack id = " + packID);  // MAR1638

			}

			try
			{
				//MAR1896 Return success if UseFulfilmentStub global parameter is TRUE
				bool bUseStub;
				string resp;

				try
				{
					bUseStub = new GlobalParameter("UseFulfilmentStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Direct Gateway stub.");
					}

					//Delay for a given time (read from the registry)
					BlockThread();

					resp = "Done";		
				}
				else
				{
					FulfilmentHandler FMHandler = new FulfilmentHandler(
						userName, password, "", packID);
				
					// RF 27/02/2006 MAR1299 Start
					//return FMHandler.CancelPack(traceOn, traceLocation);
					resp = FMHandler.CancelPack(customernumber); 
				}

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name +
						": Response from CancelPack = " + resp); 
				}
				return resp;
				// RF 27/02/2006 MAR1299 End
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": An exception occurred", ex);
				}
				throw;				
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}	

		private void BlockThread()
		{
			//Read delay in milliseconds from registry. Default value = 1000ms
			string sDelay;
			int nWaitTime;

			Microsoft.Win32.RegistryKey omFDMDelayKey;
		
			sDelay = "1000";

			try
			{
				omFDMDelayKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Omiga4\\System Configuration\\Stubs\\omFDM", false);
					
				if (omFDMDelayKey != null)
				{
					sDelay = (string)omFDMDelayKey.GetValue("Delay");
				}
			}
			catch
			{
			}

			//Delay for the given time
			nWaitTime = Convert.ToInt16(sDelay);
			TimeSpan waitTime = new TimeSpan(0,0,0,0,nWaitTime);

			System.Threading.Thread.Sleep(waitTime);		
		}
	}

	internal class FulfilmentHandler
	{
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
		private XmlDocumentEx packXML = new XmlDocumentEx();	// PSC 05/05/2006 MAR1593
		private string m_packID;
		private string m_sessionID = "";                 // MAR1638
		private static string traceFolder = "";
		private string m_objectStore = new GlobalParameter("FileNetObjectStore").String; //MAR795 GHun
		private omDg.DirectGatewayBO dg = new omDg.DirectGatewayBO();
		private string _communicationDirection = "OUT"; // RF 27/02/2006 MAR1299 

		// PSC 25/11/2005 MAR588 - Start
		private string _userName = string.Empty;
		private string _password = string.Empty;
		// PSC 25/11/2005 MAR588 - End

		/* MAR1332 GHun Not used
		enum eventkey
		{
			Created = 0,
			Edited = 1,
			Viewed = 2,
			Reprinted = 3,
			FulfilmentSendSuccess = 4,
			FulfilmentResendSuccess = 5,
			FulfilmentCancelSuccess = 6,
			SMSSuccess = 7,
			EmailSuccess = 8,
			Recategorisation = 9
		}
		*/

		public FulfilmentHandler(
			string userName, string password, 
			string inputXML, string packID)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting: Pack id = " + packID);  // MAR1638

			}

			try
			{
				// PSC 25/11/2005 MAR588 - Start
				_userName = userName;
				_password = password;
				// PSC 25/11/2005 MAR588 - End

				m_packID = packID;

				//MAR1638 Retrieve a sequence number from SeqNextNumber to use as Session ID
				m_sessionID = GetNextSequenceNumber("FDMSessionID");
				
				if (inputXML != "")
				{
					packXML.LoadXml(inputXML);
				}
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed: Session id = " + m_sessionID);  // MAR1638
				}
			}
		}

		// RF 27/02/2006 MAR1299 Use C# logging standards
		public string CancelPack(string customernumber)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Create a wse-enabled web service object to provide access to SOAP header
			try
			{
				CancelFulfilmentRequest.DirectGatewaySoapServiceWse wseService = 
				new CancelFulfilmentRequest.DirectGatewaySoapServiceWse();
				wseService.Url = dg.GetServiceUrl();
				wseService.Proxy = dg.GetProxy();

				CancelFulfilmentRequest.ProcessCancelFulfilmentResponseType webServicesResponse = 
					new CancelFulfilmentRequest.ProcessCancelFulfilmentResponseType();

				SoapContext soapContext = wseService.RequestSoapContext;

				// Add security token to SOAP header with username and password
				UsernameToken token = new UsernameToken(dg.GetProxyId(),  dg.GetProxyPwd(), 
					PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);         
				
				CancelFulfilmentRequest.ProcessCancelFulfilmentRequestType CFRequest = 
					new CancelFulfilmentRequest.ProcessCancelFulfilmentRequestType();
				CFRequest.PackID = m_packID;
				CFRequest.ClientDevice = dg.GetClientDevice();
				CFRequest.ServiceName = dg.GetServiceName();

				// PSC 25/11/2005 MAR588 - Start
				if (_userName.Length > 0 && _password.Length > 0)
				{
					CFRequest.TellerID = _userName;
					CFRequest.TellerPwd = _password;
				}

				CFRequest.CommunicationChannel = dg.GetCommunicationChannel((_userName.Length > 0));

				// RF 27/02/2006 MAR1299 Start 
				CFRequest.CommunicationDirection = _communicationDirection;
				CFRequest.SessionID = m_sessionID;    // MAR1638
				// RF 27/02/2006 MAR1299 End

				CFRequest.ProxyID = dg.GetProxyId();
				CFRequest.ProxyPwd = dg.GetProxyPwd();
				CFRequest.ProductType = dg.GetProductType();
				CFRequest.CustomerNumber = customernumber;      // MAR1638

				// PSC 25/11/2005 MAR588 - End

				CFRequest.Operator = dg.GetOperator();	//MAR852 GHun

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(CancelFulfilmentRequest.ProcessCancelFulfilmentRequestType), CFRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
				}

				try
				{
					webServicesResponse = wseService.execute(CFRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": A WebException occurred", ex);
					}
					throw new Exception("An exception occurred while sending a Cancel Pack request to Fulfilment: [" 
						+ ex.Message + "] " + webServicesResponse.ToString());
				}         

				if(_logger.IsDebugEnabled) 
				{
					string resp = dg.GetXmlFromGatewayObject(
						typeof(CancelFulfilmentRequest.ProcessCancelFulfilmentResponseType), 
						webServicesResponse);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
				}

				//MAR1638 Check the error code returned
				string strResponse = dg.GetXmlFromGatewayObject(typeof(CancelFulfilmentRequest.ProcessCancelFulfilmentResponseType), webServicesResponse);
				string strErrorType = webServicesResponse.ErrorCode;

				if (strErrorType != "0")
				{
					throw new omigaException(MethodInfo.GetCurrentMethod().Name + ": " + webServicesResponse.ErrorMessage);
				}

				return "Done";     // MAR1638
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// RF 27/02/2006 MAR1299 Use C# logging standards
		//MAR1638 Send customer number in request
		public string ReSendPack(string customernumber)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Create a wse-enabled web service object to provide access to SOAP header
			try
			{
				ResendFulfilmentRequest.DirectGatewaySoapServiceWse wseService = 
				new ResendFulfilmentRequest.DirectGatewaySoapServiceWse();
				wseService.Url = dg.GetServiceUrl();
				wseService.Proxy = dg.GetProxy();

				ResendFulfilmentRequest.ProcessResendFulfilmentResponseType webServicesResponse = 
					new ResendFulfilmentRequest.ProcessResendFulfilmentResponseType();

				SoapContext soapContext = wseService.RequestSoapContext;

				// Add security token to SOAP header with username and password
				UsernameToken token = new UsernameToken(dg.GetProxyId(),  dg.GetProxyPwd(), 
					PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);         
				
				ResendFulfilmentRequest.ProcessResendFulfilmentRequestType RSRequest = 
					new ResendFulfilmentRequest.ProcessResendFulfilmentRequestType();
				RSRequest.PackID = m_packID;
				RSRequest.ClientDevice = dg.GetClientDevice();
				RSRequest.ServiceName = dg.GetServiceName();
				
				//MAR852 GHun
				if (_userName.Length > 0 && _password.Length > 0)
				{
					RSRequest.TellerID = _userName;
					RSRequest.TellerPwd = _password;
				}
				RSRequest.ProxyID = dg.GetProxyId();
				RSRequest.ProxyPwd = dg.GetProxyPwd();
				RSRequest.CommunicationChannel = dg.GetCommunicationChannel((_userName.Length > 0));

				// RF 27/02/2006 MAR1299 Start
				RSRequest.CommunicationDirection = _communicationDirection;
				RSRequest.SessionID = m_sessionID;    //MAR1638
				// RF 27/02/2006 MAR1299 End

				RSRequest.ProductType = dg.GetProductType();
				RSRequest.CustomerNumber = customernumber;    // MAR1638
				RSRequest.Operator = dg.GetOperator();
				//MAR852 End
				
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(ResendFulfilmentRequest.ProcessResendFulfilmentRequestType), RSRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
				}

				try
				{
					webServicesResponse = wseService.execute(RSRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": A WebException occurred", ex);
					}
					throw new Exception("An exception occurred while re-sending a Pack to Fulfilment: [" 
						+ ex.Message + "] " + webServicesResponse.ToString());
				}         
				
				if(_logger.IsDebugEnabled) 
				{
					string resp  = dg.GetXmlFromGatewayObject(
						typeof(ResendFulfilmentRequest.ProcessResendFulfilmentResponseType), webServicesResponse);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
				}

				//MAR1638 Check the error code returned
				string strResponse = dg.GetXmlFromGatewayObject(typeof(ResendFulfilmentRequest.ProcessResendFulfilmentResponseType), webServicesResponse);
				string strErrorType = webServicesResponse.ErrorCode;

				if (strErrorType != "0")
				{
					throw new omigaException(MethodInfo.GetCurrentMethod().Name + ": " + webServicesResponse.ErrorMessage);
				}

				return "Done";     
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// Send document to SMS
		public string SendToSMSFulfilment(bool bLogRequestAndResponse, string traceFolderLocation)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Create a wse-enabled web service object to provide access to SOAP header
			try
			{
				SMSFulfilmentRequest.DirectGatewaySoapServiceWse wseService = 
				new SMSFulfilmentRequest.DirectGatewaySoapServiceWse();
				wseService.Url = dg.GetServiceUrl();
				wseService.Proxy = dg.GetProxy();

				// MAR1638
				SMSFulfilmentRequest.ProcessSMSFulfilmentResponseType webServicesResponse = 
					new SMSFulfilmentRequest.ProcessSMSFulfilmentResponseType();

				SoapContext soapContext = wseService.RequestSoapContext;

				// Add security token to SOAP header with username and password
				UsernameToken token = new UsernameToken(dg.GetProxyId(),  dg.GetProxyPwd(), 
					PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);         
		
				traceFolder = traceFolderLocation;

				// PSC 05/05/2006 MAR1593
				XmlElementEx packControlNode = (XmlElementEx)packXML.SelectMandatorySingleNode("//PACKCONTROL");
				XmlNode addressNode = packControlNode.SelectSingleNode("//ADDRESS");
				XmlNode customerNode = addressNode.SelectSingleNode("//CUSTOMER");
				XmlNodeList packMemberNodeList = packControlNode.SelectNodes("//PACKMEMBER");

				SMSFulfilmentRequest.ProcessSMSFulfilmentRequestType SMSRequest = new SMSFulfilmentRequest.ProcessSMSFulfilmentRequestType();
				SMSFulfilmentRequest.PersonDetailsType SMSPersonDetails = new SMSFulfilmentRequest.PersonDetailsType();
				SMSFulfilmentRequest.StructuredAddressDetailsType SMSAddressDetails = new SMSFulfilmentRequest.StructuredAddressDetailsType();
				SMSFulfilmentRequest.PackDetailsType SMSPackDetails = new SMSFulfilmentRequest.PackDetailsType();

				SMSRequest.ClientDevice = dg.GetClientDevice();
				SMSRequest.ServiceName = dg.GetServiceName();
				SMSRequest.PackID = GetAttribute(packControlNode, "PACKGUID");
				SMSRequest.CustomerNumber = GetAttribute(customerNode, "OTHERSYSTEMCUSTOMERNUMBER");
				// PSC 14/12/2005 MAR802
				SMSRequest.ContactNumber = GetAttribute(customerNode, "AREACODE") + GetAttribute(customerNode, "TELEPHONENUMBER");
				
				//MAR852 GHun
				if (_userName.Length > 0 && _password.Length > 0)
				{
					SMSRequest.TellerID = _userName;
					SMSRequest.TellerPwd = _password;
				}
				SMSRequest.ProxyID = dg.GetProxyId();
				SMSRequest.ProxyPwd = dg.GetProxyPwd();
				SMSRequest.ProductType = dg.GetProductType();
				SMSRequest.Operator = dg.GetOperator();
				//MAR852 End

				SMSRequest.CommunicationChannel = "SMS";

				// RF 27/02/2006 MAR1299 Start
				SMSRequest.CommunicationDirection = _communicationDirection;
				SMSRequest.SessionID = m_sessionID;    // MAR1638
				// RF 27/02/2006 MAR1299 End

				// PSC 19/01/2006 MAR1087
				// PSC 05/05/2006 MAR1593
				SMSRequest.PackTypeID = packControlNode.GetMandatoryAttribute("PACKTYPEID");

				SMSPersonDetails.Salutation = GetAttribute(customerNode, "TITLE");
				SMSPersonDetails.FirstName = GetAttribute(customerNode, "FIRSTFORENAME");
				SMSPersonDetails.MiddleName = GetAttribute(customerNode, "SECONDFORENAME");
				SMSPersonDetails.LastName = GetAttribute(customerNode, "SURNAME");
				
				if (GetAttribute(customerNode, "SPECIALNEEDS") != "")
				{
					switch (GetAttribute(customerNode, "SPECIALNEEDS").Substring(0, 1).ToUpper())
					{
						case "A":
							SMSPersonDetails.SpecialNeed = SMSFulfilmentRequest.SpecialNeedType.A;
							break;
						case "L":
							SMSPersonDetails.SpecialNeed = SMSFulfilmentRequest.SpecialNeedType.L;
							break;
						case "B":
							SMSPersonDetails.SpecialNeed = SMSFulfilmentRequest.SpecialNeedType.B;
							break;
						default:
							SMSPersonDetails.SpecialNeed = SMSFulfilmentRequest.SpecialNeedType.Item0;
							break;
					}
				}

				SMSAddressDetails.HouseOrBuildingName = GetAttribute(addressNode,"BUILDINGORHOUSENAME");
				SMSAddressDetails.HouseOrBuildingNumber = GetAttribute(addressNode,"BUILDINGORHOUSENUMBER");
				SMSAddressDetails.FlatNameOrNumber = GetAttribute(addressNode,"FLATNUMBER");
				SMSAddressDetails.Street = GetAttribute(addressNode,"STREET");
				SMSAddressDetails.District = GetAttribute(addressNode,"DISTRICT");
				SMSAddressDetails.TownOrCity = GetAttribute(addressNode,"TOWN");
				SMSAddressDetails.County = GetAttribute(addressNode,"COUNTY");
				SMSAddressDetails.PostCode = GetAttribute(addressNode,"POSTCODE");
				SMSAddressDetails.CountryCode = GetAttribute(addressNode,"COUNTRYCODE");

				if (GetAttribute(customerNode, "CUSTOMERORDER") == "1")
				{
					SMSPackDetails.MainApplicantIndicator = Vertex.Fsd.Omiga.omFDM.SMSFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@true;
				}
				else
				{
					SMSPackDetails.MainApplicantIndicator = Vertex.Fsd.Omiga.omFDM.SMSFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@false;
				}

				SMSPackDetails.ProductAccountNumber = ""; //??????????
				SMSPackDetails.ProductApplicationNumber = GetAttribute(addressNode, "APPLICATIONNUMBER");
				SMSPackDetails.Address = SMSAddressDetails;

				if (SMSAddressDetails.CountryCode == "3")
				{
					SMSPackDetails.PropertyLocation = Vertex.Fsd.Omiga.omFDM.SMSFulfilmentRequest.PackDetailsTypePropertyLocation.Item1; //Scottish
				}
				else
				{
					SMSPackDetails.PropertyLocation = Vertex.Fsd.Omiga.omFDM.SMSFulfilmentRequest.PackDetailsTypePropertyLocation.Item0; //Scottish
				}
	
				//MAR1302  Add Person details to request
				SMSPackDetails.Person = SMSPersonDetails;

				SMSRequest.PackDetails = SMSPackDetails;
	
				if (bLogRequestAndResponse)
				{
					try
					{
						string fileOut = traceFolder + "SMS_Request" + SMSRequest.PackID + ".xml";
						XmlSerializer serOut = new XmlSerializer(typeof(SMSFulfilmentRequest.ProcessSMSFulfilmentRequestType));
						XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
						serOut.Serialize(writer, SMSRequest);
						writer.Close();
					}
					catch(Exception ex)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(MethodInfo.GetCurrentMethod().Name + 
								": An exception occurred", ex);
						}
						throw new Exception("An exception occurred while writing to the logging file: [" + ex.Message + "]");
					}
				}

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(SMSFulfilmentRequest.ProcessSMSFulfilmentRequestType), SMSRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
				}

				try
				{
					webServicesResponse = wseService.execute(SMSRequest);  // MAR1638
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": A WebException occurred", ex);
					}

					throw new Exception("An exception occurred while sending a document to SMS Fulfilment: [" 
						+ ex.Message + "] " + webServicesResponse.ToString());
				}
				
				//MAR1638
				if(_logger.IsDebugEnabled) 
				{
					string resp  = dg.GetXmlFromGatewayObject(
						typeof(SMSFulfilmentRequest.ProcessSMSFulfilmentResponseType), webServicesResponse);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
				}

				//MAR1638 Check the error code returned
				string strResponse = dg.GetXmlFromGatewayObject(typeof(SMSFulfilmentRequest.ProcessSMSFulfilmentResponseType), webServicesResponse);
				string strErrorType = webServicesResponse.ErrorCode;

				if (strErrorType != "0")
				{
					throw new omigaException(MethodInfo.GetCurrentMethod().Name + ": " + webServicesResponse.ErrorMessage);
				}

				return "Done";     
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// Send document to Email
		public string SendToEMAILFulfilment(bool bLogRequestAndResponse, string traceFolderLocation)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Create a wse-enabled web service object to provide access to SOAP header
			try
			{
				EmailFulfilmentRequest.DirectGatewaySoapServiceWse wseService = new EmailFulfilmentRequest.DirectGatewaySoapServiceWse();
				wseService.Url = dg.GetServiceUrl();
				wseService.Proxy = dg.GetProxy();

				// MAR1638
				EmailFulfilmentRequest.ProcessEmailFulfilmentResponseType webServicesResponse = 
					new EmailFulfilmentRequest.ProcessEmailFulfilmentResponseType();

				SoapContext soapContext = wseService.RequestSoapContext;

				// Add security token to SOAP header with username and password
				UsernameToken token = new UsernameToken(dg.GetProxyId(),  dg.GetProxyPwd(), 
					PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);         
					
				traceFolder = traceFolderLocation;
				// PSC 05/05/2006 MAR1593
				XmlElementEx packControlNode = (XmlElementEx)packXML.SelectMandatorySingleNode("//PACKCONTROL");
				XmlNode addressNode = packControlNode.SelectSingleNode("//ADDRESS");
				XmlNode customerNode = addressNode.SelectSingleNode("//CUSTOMER");
				XmlNodeList packMemberNodeList = packControlNode.SelectNodes("//PACKMEMBER");

				EmailFulfilmentRequest.ProcessEmailFulfilmentRequestType EMailRequest = new EmailFulfilmentRequest.ProcessEmailFulfilmentRequestType();
				EmailFulfilmentRequest.PersonDetailsType EMailPersonalDetails = new EmailFulfilmentRequest.PersonDetailsType();
				EmailFulfilmentRequest.StructuredAddressDetailsType EMailAddressDetails = new EmailFulfilmentRequest.StructuredAddressDetailsType();
				EmailFulfilmentRequest.PackDetailsType EMailPackDetails = new EmailFulfilmentRequest.PackDetailsType();
				
				EMailRequest.PackID = GetAttribute(packControlNode, "PACKGUID");
				EMailRequest.CustomerNumber = GetAttribute(customerNode, "OTHERSYSTEMCUSTOMERNUMBER");
				
				// PSC 01/04/2006 MAR993 - Start
				EMailRequest.ClientDevice = dg.GetClientDevice();
				EMailRequest.ServiceName = dg.GetServiceName();
				// PSC 01/04/2006 MAR993 - End

				//MAR852 GHun
				if (_userName.Length > 0 && _password.Length > 0)
				{
					EMailRequest.TellerID = _userName;
					EMailRequest.TellerPwd = _password;
				}
				EMailRequest.CommunicationChannel = "EMAIL";

				// RF 27/02/2006 MAR1299 Start
				EMailRequest.CommunicationDirection = _communicationDirection;
				EMailRequest.SessionID = m_sessionID;    // MAR1638
				// RF 27/02/2006 MAR1299 End

				EMailRequest.Operator = dg.GetOperator();
				EMailRequest.ProductType = dg.GetProductType();
				EMailRequest.ProxyID = dg.GetProxyId();
				EMailRequest.ProxyPwd = dg.GetProxyPwd();
				//MAR852 End
				// PSC 19/01/2006 MAR1087
				// PSC 05/05/2006 MAR1593
				EMailRequest.PackTypeID = packControlNode.GetMandatoryAttribute("PACKTYPEID");

				EMailPersonalDetails.Salutation = GetAttribute(customerNode, "TITLE");
				EMailPersonalDetails.FirstName = GetAttribute(customerNode, "FIRSTFORENAME");
				EMailPersonalDetails.MiddleName = GetAttribute(customerNode, "SECONDFORENAME");
				EMailPersonalDetails.LastName = GetAttribute(customerNode, "SURNAME");
				
				if (GetAttribute(customerNode, "SPECIALNEEDS") != "")
				{
					switch (GetAttribute(customerNode, "SPECIALNEEDS").Substring(0, 1).ToUpper())
					{
						case "A":
							EMailPersonalDetails.SpecialNeed = EmailFulfilmentRequest.SpecialNeedType.A;
							break;
						case "L":
							EMailPersonalDetails.SpecialNeed = EmailFulfilmentRequest.SpecialNeedType.L;
							break;
						case "B":
							EMailPersonalDetails.SpecialNeed = EmailFulfilmentRequest.SpecialNeedType.B;
							break;
						default:
							EMailPersonalDetails.SpecialNeed = EmailFulfilmentRequest.SpecialNeedType.Item0;
							break;
					}
				}

				EMailAddressDetails.HouseOrBuildingName = GetAttribute(addressNode,"BUILDINGORHOUSENAME");
				EMailAddressDetails.HouseOrBuildingNumber = GetAttribute(addressNode,"BUILDINGORHOUSENUMBER");
				EMailAddressDetails.FlatNameOrNumber = GetAttribute(addressNode,"FLATNUMBER");
				EMailAddressDetails.Street = GetAttribute(addressNode,"STREET");
				EMailAddressDetails.District = GetAttribute(addressNode,"DISTRICT");
				EMailAddressDetails.TownOrCity = GetAttribute(addressNode,"TOWN");
				EMailAddressDetails.County = GetAttribute(addressNode,"COUNTY");
				EMailAddressDetails.PostCode = GetAttribute(addressNode,"POSTCODE");
				EMailAddressDetails.CountryCode = GetAttribute(addressNode,"COUNTRYCODE");

				if (GetAttribute(customerNode, "CUSTOMERORDER") == "1")
				{
					EMailPackDetails.MainApplicantIndicator = Vertex.Fsd.Omiga.omFDM.EmailFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@true;
				}
				else
				{
					EMailPackDetails.MainApplicantIndicator = Vertex.Fsd.Omiga.omFDM.EmailFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@false;
				}

				EMailPackDetails.ProductAccountNumber = ""; //??????????
				EMailPackDetails.ProductApplicationNumber = GetAttribute(addressNode, "APPLICATIONNUMBER");
				EMailPackDetails.Address = EMailAddressDetails;

				if (EMailAddressDetails.CountryCode == "3")
				{
					EMailPackDetails.PropertyLocation = Vertex.Fsd.Omiga.omFDM.EmailFulfilmentRequest.PackDetailsTypePropertyLocation.Item1; //Scottish
				}
				else
				{
					EMailPackDetails.PropertyLocation = Vertex.Fsd.Omiga.omFDM.EmailFulfilmentRequest.PackDetailsTypePropertyLocation.Item0; //Scottish
				}

				//MAR1302  Add Person details to request.
				EMailPackDetails.Person = EMailPersonalDetails;

				EMailRequest.PackDetails = EMailPackDetails;
				// PSC 14/12/2005 MAR802
				EMailRequest.EmailAddress = GetAttribute(customerNode, "CONTACTEMAILADDRESS");
	
				if (bLogRequestAndResponse)
				{
					try
					{
						string fileOut = traceFolder + "EMail_Request" + EMailRequest.PackID + ".xml";
						XmlSerializer serOut = new XmlSerializer(typeof(EmailFulfilmentRequest.ProcessEmailFulfilmentRequestType));
						XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
						serOut.Serialize(writer, EMailRequest);
						writer.Close();
					}
					catch(Exception ex)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(MethodInfo.GetCurrentMethod().Name + 
								": An exception occurred", ex);
						}
						throw new Exception("An exception occurred while writing to the logging file: [" + ex.Message + "]");
					}
				}

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(EmailFulfilmentRequest.ProcessEmailFulfilmentRequestType), EMailRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
				}

				try
				{
					webServicesResponse = wseService.execute(EMailRequest);   // MAR1638
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("An exception occurred while sending an EMail document to Fulfilment: [" 
						+ ex.Message + "] " + webServicesResponse.ToString());
				}
				
				//MAR1638
				if(_logger.IsDebugEnabled) 
				{
					string resp  = dg.GetXmlFromGatewayObject(
						typeof(EmailFulfilmentRequest.ProcessEmailFulfilmentResponseType), webServicesResponse);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
				}

				//MAR1638 Check the error code returned
				string strResponse = dg.GetXmlFromGatewayObject(typeof(EmailFulfilmentRequest.ProcessEmailFulfilmentResponseType), webServicesResponse);
				string strErrorType = webServicesResponse.ErrorCode;

				if (strErrorType != "0")
				{
					throw new omigaException(MethodInfo.GetCurrentMethod().Name + ": " + webServicesResponse.ErrorMessage);
				}

				return "Done";     // MAR1638			
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		// Send document to Fulfilment
		// RF 27/02/2006 MAR1299 Use C# logging standards
		public string SendToFulfilment()
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Create a wse-enabled web service object to provide access to SOAP header
			try
			{
				PrintFulfilmentRequest.DirectGatewaySoapServiceWse wseService = 
				new PrintFulfilmentRequest.DirectGatewaySoapServiceWse();
				wseService.Url = dg.GetServiceUrl();
				wseService.Proxy = dg.GetProxy();

				PrintFulfilmentRequest.ProcessPrintFulfilmentResponseType webServicesResponse = 
					new PrintFulfilmentRequest.ProcessPrintFulfilmentResponseType();

				SoapContext soapContext = wseService.RequestSoapContext;

				// Add security token to SOAP header with username and password
				UsernameToken token = new UsernameToken(dg.GetProxyId(), dg.GetProxyPwd(), 
					PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);         
				
				XmlNode packControlNode = packXML.SelectSingleNode("//PACKCONTROL");
				XmlNode addressNode = packControlNode.SelectSingleNode("//ADDRESS");
				XmlNode customerNode = addressNode.SelectSingleNode("//CUSTOMER");
				XmlNodeList packMemberNodeList = packControlNode.SelectNodes("//PACKMEMBER");

				PrintFulfilmentRequest.ProcessPrintFulfilmentRequestType PFRequest = new PrintFulfilmentRequest.ProcessPrintFulfilmentRequestType();
				PrintFulfilmentRequest.PersonDetailsType PFPersonDetails = new PrintFulfilmentRequest.PersonDetailsType();
				PrintFulfilmentRequest.StructuredAddressDetailsType PFAddressDetails = new PrintFulfilmentRequest.StructuredAddressDetailsType();
				PrintFulfilmentRequest.PackDetailsType PFPackDetails = new PrintFulfilmentRequest.PackDetailsType();
				PFRequest.PackID = GetAttribute(packControlNode, "PACKGUID");
				PFRequest.CustomerNumber = GetAttribute(customerNode, "OTHERSYSTEMCUSTOMERNUMBER");
				
				// PSC 04/01/2006 MAR993 - Start
				PFRequest.ClientDevice = dg.GetClientDevice();
				PFRequest.ServiceName = dg.GetServiceName();
				// PSC 04/01/2006 MAR993 - End

				//MAR852 GHun
				if (_userName.Length > 0 && _password.Length > 0)
				{
					PFRequest.TellerID = _userName;
					PFRequest.TellerPwd = _password;
				}
				PFRequest.ProxyID = dg.GetProxyId();
				PFRequest.ProxyPwd = dg.GetProxyPwd();
				PFRequest.CommunicationChannel = "MAIL";

				// RF 27/02/2006 MAR1299 
				PFRequest.CommunicationDirection = _communicationDirection;
				PFRequest.SessionID = m_sessionID;    // MAR1638
				// RF 27/02/2006 MAR1299 End

				PFRequest.ProductType = dg.GetProductType();
				PFRequest.Operator = dg.GetOperator();
				//MAR852 End

				PFRequest.PackTypeID = GetAttribute(packControlNode, "PACKCONTROLNUMBER");  // MAR978

				PFPersonDetails.Salutation = GetAttribute(customerNode, "TITLE");
				PFPersonDetails.FirstName = GetAttribute(customerNode, "FIRSTFORENAME");
				PFPersonDetails.MiddleName = GetAttribute(customerNode, "SECONDFORENAME");
				PFPersonDetails.LastName = GetAttribute(customerNode, "SURNAME");
				PFPersonDetails.SpecialNeed = PrintFulfilmentRequest.SpecialNeedType.Item0; //DRC 978 revisited
				if (GetAttribute(customerNode, "SPECIALNEEDS") != "")
				{
					switch (GetAttribute(customerNode, "SPECIALNEEDS").Substring(0, 1).ToUpper())
					{
						case "A":
							PFPersonDetails.SpecialNeed = PrintFulfilmentRequest.SpecialNeedType.A;
							break;
						case "L":
							PFPersonDetails.SpecialNeed = PrintFulfilmentRequest.SpecialNeedType.L;
							break;
						case "B":
							PFPersonDetails.SpecialNeed = PrintFulfilmentRequest.SpecialNeedType.B;
							break;
						default:
							PFPersonDetails.SpecialNeed = PrintFulfilmentRequest.SpecialNeedType.Item0;
							break;
					}
				}
                PFPersonDetails.SpecialNeedSpecified = true; //DRC MAR978 revisited
				PFAddressDetails.HouseOrBuildingName = GetAttribute(addressNode,"BUILDINGORHOUSENAME");
				PFAddressDetails.HouseOrBuildingNumber = GetAttribute(addressNode,"BUILDINGORHOUSENUMBER");
				PFAddressDetails.FlatNameOrNumber = GetAttribute(addressNode,"FLATNUMBER");
				PFAddressDetails.Street = GetAttribute(addressNode,"STREET");
				PFAddressDetails.District = GetAttribute(addressNode,"DISTRICT");
				PFAddressDetails.TownOrCity = GetAttribute(addressNode,"TOWN");
				PFAddressDetails.County = GetAttribute(addressNode,"COUNTY");
				PFAddressDetails.PostCode = GetAttribute(addressNode,"POSTCODE");
				PFAddressDetails.CountryCode = GetAttribute(addressNode,"COUNTRY");

				if (GetAttribute(customerNode, "CUSTOMERORDER") == "1")
				{
					PFPackDetails.MainApplicantIndicator = PrintFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@true;		
				}
				else
				{
					PFPackDetails.MainApplicantIndicator = PrintFulfilmentRequest.PackDetailsTypeMainApplicantIndicator.@false;
				}
				PFPackDetails.MainApplicantIndicatorSpecified = true; //DRC MAR978
				PFPackDetails.ProductAccountNumber = ""; //??????????
				PFPackDetails.ProductApplicationNumber = GetAttribute(addressNode, "APPLICATIONNUMBER");
				//DRC MAR978
				PFPackDetails.DateTimeCreated = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
				PFPackDetails.Address = PFAddressDetails;
				PFPackDetails.Person = PFPersonDetails; //DRC MAR978

				if (GetAttribute(packControlNode, "INSCOTLAND")  == "1") //DRC MAR1193
				{
					PFPackDetails.PropertyLocation = PrintFulfilmentRequest.PackDetailsTypePropertyLocation.Item1; //Scottish
				}
				else
				{
					PFPackDetails.PropertyLocation = PrintFulfilmentRequest.PackDetailsTypePropertyLocation.Item0; //Elsewhere
				}
				PFPackDetails.PropertyLocationSpecified = true; //DRC MAR978

				PFRequest.ImageReference = new string[packMemberNodeList.Count];
				int x = 0;
				foreach (XmlNode packMemberNode in packMemberNodeList)
				{
					PFRequest.ImageReference[x] = m_objectStore + "|" + GetAttribute(packMemberNode,"FILENETIMAGEREF");
					if (GetAttribute(packMemberNode, "PRIMARY") == "true")
					{
						PFRequest.PrimaryImageReference = m_objectStore + "|" + GetAttribute(packMemberNode, "FILENETIMAGEREF");
					}
					x++;
				}
				PFRequest.PackDetails = PFPackDetails;
	
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(PrintFulfilmentRequest.ProcessPrintFulfilmentRequestType), PFRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
				}

				try
				{
					webServicesResponse = wseService.execute(PFRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					//MAR1638
					throw new Exception("An exception occurred while sending a document to Fulfilment: [" 
						+ ex.Message + "] " + webServicesResponse.ToString());
				}         

				if(_logger.IsDebugEnabled) 
				{
					string resp = dg.GetXmlFromGatewayObject(
						typeof(PrintFulfilmentRequest.ProcessPrintFulfilmentResponseType), 
						webServicesResponse);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
				}      

				//MAR1638 Check the error code returned
				string strResponse = dg.GetXmlFromGatewayObject(typeof(PrintFulfilmentRequest.ProcessPrintFulfilmentResponseType), webServicesResponse);
				string strErrorType = webServicesResponse.ErrorCode;

				if (strErrorType != "0")
				{
					throw new omigaException(MethodInfo.GetCurrentMethod().Name + ": " + webServicesResponse.ErrorMessage);
				}

				return "Done";     
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}
		}

		//MAR1638  Add code to get next sequence number
		private string GetNextSequenceNumber(string sequenceName)
		{
			string connString = Global.DatabaseConnectionString + "Enlist=false;";	
			string nextNumber;

			using (SqlConnection conn = new SqlConnection(connString))
			{
				SqlCommand cmd = new SqlCommand("USP_GetNextSequenceNumber", conn);
				cmd.CommandType = CommandType.StoredProcedure;
				
				SqlParameter paramSName = cmd.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, sequenceName.Length);
				paramSName.Value = sequenceName;
				SqlParameter paramNextNumber = cmd.Parameters.Add("@p_NextNumber", SqlDbType.NVarChar, 12);
				paramNextNumber.Direction = ParameterDirection.Output;

				conn.Open();
				cmd.ExecuteNonQuery();

				if (paramNextNumber.Value != DBNull.Value)	
					nextNumber = (string) paramNextNumber.Value;
				else
					nextNumber = string.Empty;
			}

			return nextNumber;
		}

		private string GetAttribute(XmlNode theNode, string attributeName)
		{
			XmlNode tempAttribute = theNode.Attributes.GetNamedItem(attributeName);
			if (tempAttribute == null)
			{
				return "";			
			}
			else
			{
				return tempAttribute.InnerText;
			}
		}

		private string[] ConvertStringToStringArray(string theString)
		{
			string[] temp = new string[theString.Length];
			for (int i = 0; i < theString.Length; i++)
			{
				temp[i] = theString.Substring(i, 1);
			}
			return temp;
		}
	}

	internal class FileNetHandler
	{
		private string m_stringData;
		private string m_applicationNumber;
		private string m_url;
		private string m_objectStore = new GlobalParameter("FileNetObjectStore").String; //MAR795 GHun
		private WebRefFileNet.FNCEWS10ServiceWse wseService;
		private static string traceFolder = "";
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		// PSC 06/01/2006 MAR922
		private const string _fileNetRootFolderId = "{0F1E2D3C-4B5A-6978-8796-A5B4C3D2E1F0}";

		public FileNetHandler(string applicationNumber, string documentData, string userName, string password)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				m_applicationNumber = applicationNumber;
				m_stringData = documentData;
			
				//			m_url = "http://195.245.252.218:80/FNCEWS10SOAP/";
				Core.GlobalParameter GlobalParam = new GlobalParameter("FileNetURL");
				m_url = GlobalParam.String;

				// Create a wse-enabled web service object to provide access to SOAP header
				wseService = new WebRefFileNet.FNCEWS10ServiceWse(); 

				omDg.DirectGatewayBO dg = new omDg.DirectGatewayBO();
				Core.GlobalParameter ProxyServer = new GlobalParameter("ProxyServer");
				if (ProxyServer.String.Length > 0)
				{
					IWebProxy proxy = new WebProxy(ProxyServer.String, true);
					proxy.Credentials = CredentialCache.DefaultCredentials;
					wseService.Proxy = proxy;
				}

				wseService.Url = m_url;

				SoapContext soapContext = wseService.RequestSoapContext;         
				
				// Add security token to SOAP header with your username and password
				UsernameToken token = new UsernameToken(  
					userName,  password, PasswordOption.SendPlainText);
				soapContext.Security.Tokens.Add(token);
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		public string GetFileNetRecord(bool bLogRequestAndResponse, string traceFolderLocation, string fileNetGUID)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				traceFolder = traceFolderLocation;
				ObjectReference objRootFolder = new ObjectReference();
				objRootFolder.classId = "Document";

				string documentPath = "/Mortgages/" + m_applicationNumber + "/" + fileNetGUID; // A pre-existing document must be here.

				// Create an object reference to the root folder

				ObjectReference objRootDocument = new ObjectReference();

				//Specify its class, object ID, and object store

				objRootDocument.classId = "Document";
				objRootDocument.objectId = fileNetGUID;
				objRootDocument.objectStore = m_objectStore;
			
				ObjectRequestType elemObjectRequestType = new ObjectRequestType();

				elemObjectRequestType.SourceSpecification = objRootDocument;

				ObjectRequestType[] elemObjectRequestTypeArray = new ObjectRequestType[1];

				// Filters
				FilterElementType elemFilterElementType = new FilterElementType();

				elemFilterElementType.maxRecursion = 1;
				elemFilterElementType.maxRecursionSpecified = true;
				elemFilterElementType.Value = "ContentData";  

				FilterElementType[] elemFilterElementTypeArray = new FilterElementType[1];

				elemFilterElementTypeArray[0] = elemFilterElementType;

				PropertyFilterType elemPropertyFilterType = new PropertyFilterType();

				elemPropertyFilterType.IncludeTypes = elemFilterElementTypeArray;
				elemObjectRequestType.PropertyFilter = elemPropertyFilterType;

				// PSC 06/02/2006 MAR1197 - Start
				elemPropertyFilterType.IncludeProperties = new FilterElementType[1];
				elemPropertyFilterType.IncludeProperties[0] = new FilterElementType();
				elemPropertyFilterType.IncludeProperties[0].Value = "MimeType";
				// PSC 06/02/2006 MAR1197 - End

				elemObjectRequestTypeArray[0] = elemObjectRequestType;

				ObjectResponseType[] elemObjectResponseTypeArray = null;

				if (bLogRequestAndResponse)
				{
					try
					{
						string fileOut = traceFolder + "/" + m_applicationNumber + "_GetRecord_Request.xml";
						XmlSerializer serOut = new XmlSerializer(typeof(ObjectRequestType[]));
						XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
						serOut.Serialize(writer, elemObjectRequestTypeArray);
						writer.Close();
					}
					catch(Exception ex)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(MethodInfo.GetCurrentMethod().Name + 
								": An exception occurred", ex);
						}
						throw new Exception("An exception occurred writing to the log file." + ex.Message + "]");
					}
				}

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string requestString  = dg.GetXmlFromGatewayObject(
						typeof(ObjectRequestType[]), elemObjectRequestTypeArray);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString);
					dg = null;
				}

				try
				{
					elemObjectResponseTypeArray = wseService.GetObjects(elemObjectRequestTypeArray);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("No document returned. " + ex.Message);
				}

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string resp = dg.GetXmlFromGatewayObject(
						typeof(ObjectResponseType[]), 
						elemObjectResponseTypeArray);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + resp);
					dg = null;
				}      

				if (bLogRequestAndResponse)
				{
					try
					{
						string fileOut = traceFolder + "/" + m_applicationNumber + "_GetRecord_Response.xml";
						XmlSerializer serOut = new XmlSerializer(typeof(ObjectResponseType[]));
						XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
						serOut.Serialize(writer, elemObjectResponseTypeArray);
						writer.Close();
					}
					catch(Exception ex)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(MethodInfo.GetCurrentMethod().Name + 
								": An exception occurred", ex);
						}
						throw new Exception("An exception occurred writing to the log file." + ex.Message + "]");
					}
				}

				XmlSerializer elmRequestXmlSerializer = new XmlSerializer(typeof(ObjectResponseType[]));
				StringWriter elmRequestStringWriter = new StringWriter();
				XmlTextWriter elmRequestTextWriter = new XmlTextWriter(elmRequestStringWriter);
				elmRequestXmlSerializer.Serialize(elmRequestTextWriter, elemObjectResponseTypeArray);
				elmRequestTextWriter.Flush();
				elmRequestTextWriter.Close();

				string documentData = elmRequestStringWriter.ToString();

				if (documentData == "")
				{
					throw new Exception("Failed to convert document to string");
				}	
				try
				{
					// PSC 06/02/2006 MAR1197 - Start
					// Get content type 
					string mimeType = string.Empty;

					ObjectValue returnedObject = (ObjectValue)elemObjectResponseTypeArray[0].Item;
			
					foreach (PropertyType returnedProperty in returnedObject.Property)
					{
						if (returnedProperty.propertyId == "MimeType")
						{
							mimeType = ((SingletonString)returnedProperty).Value;
							break;
						}
					}

					string contentType = string.Empty;

					switch (mimeType.ToUpper())
					{
						case "APPLICATION/PDF":
						{
							contentType = "pdf";
							break;
						}
						case "APPLICATION/RTF":
						{
							contentType = "rtf";
							break;
						}
						case "APPLICATION/MSWORD":
						{
							contentType = "doc";
							break;
						}
						case "IMAGE/TIFF":
						{
							contentType = "tif";
							break;
						}
						default:
						{
							throw new Exception ("Unsupported MIME type " + mimeType);
						}
					}
				
					int x = documentData.LastIndexOf("<Binary>");
					int y = documentData.LastIndexOf("</Binary>");

					XmlDocumentEx documentContents = new XmlDocumentEx();
					XmlElementEx filenetRecord = (XmlElementEx)documentContents.CreateElement("FILENETRECORD");
					documentContents.AppendChild(filenetRecord);
					filenetRecord.SetAttribute("FILECONTENTS", documentData.Substring(x + 8, y - x - 8));
					filenetRecord.SetAttribute("CONTENTTYPE", contentType);

					return documentContents.OuterXml;
					// PSC 06/02/2006 MAR1197 - End
				}
				catch (Exception ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("No data content for this document");
				}
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		// PSC 11/01/2006 MAR994 - Start
		public string CheckFolderExists(bool bLogRequestAndResponse, string traceFolderLocation, string folderPath, string folderName)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				traceFolder = traceFolderLocation;

				string revisedPath = folderPath.Replace("\\", "/");
				char [] folderSeperator = new char[] {'/'};
				revisedPath = revisedPath.TrimEnd(folderSeperator);
				revisedPath = revisedPath.TrimStart(folderSeperator);

				string parentFolderId = string.Empty;
				string folderId = GetFolderId(bLogRequestAndResponse, traceFolderLocation, revisedPath, folderName);

				if (folderId.Length == 0)
				{
					if (revisedPath.Length == 0)
					{
						parentFolderId = _fileNetRootFolderId;
					}
					else
					{
						string parentFolder = string.Empty;
						string parentFolderPath  = string.Empty;

						int lastSlash = revisedPath.LastIndexOf('/');

						if (lastSlash == -1)
						{
							parentFolder = revisedPath;
						}
						else
						{
							parentFolder = revisedPath.Substring(lastSlash + 1);
							parentFolderPath = revisedPath.Substring(0, lastSlash);
						}


						parentFolderId = CheckFolderExists(bLogRequestAndResponse, traceFolderLocation, parentFolderPath, parentFolder);
					}
		
					folderId = CreateFolder(bLogRequestAndResponse, traceFolderLocation, parentFolderId, folderName);
				}

				return folderId;
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		public string GetFolderId(bool bLogRequestAndResponse, string traceFolderLocation, string folderPath, string folderName)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				traceFolder = traceFolderLocation;

				ObjectReference objRootFolder = new ObjectReference();
				objRootFolder.classId = "Folder";

				// Specify the scope of the search
				ObjectStoreScope elemObjectStoreScope = new ObjectStoreScope();
				elemObjectStoreScope.objectStore = m_objectStore;         // Create the search for unfiled doc
				RepositorySearch elemRepositorySearch = new RepositorySearch();
				elemRepositorySearch.repositorySearchMode = RepositorySearchModeType.Rows;
				elemRepositorySearch.repositorySearchModeSpecified = true;
				elemRepositorySearch.SearchScope = elemObjectStoreScope;
				// PSC 06/01/2006 MAR922
				elemRepositorySearch.SearchSQL = "SELECT [Id] FROM [Folder] WHERE ([This] INFOLDER '/" + folderPath + "' And [FolderName] = '" + folderName + "')"; 

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string requestString = dg.GetXmlFromGatewayObject(typeof(RepositorySearch), elemRepositorySearch);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service request: " + requestString);
					dg = null;
				}      

				// Invoke the ExecuteSearch operation
				ObjectSetType objObjectSet = wseService.ExecuteSearch(elemRepositorySearch);

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string responseString = dg.GetXmlFromGatewayObject(typeof(ObjectSetType), objObjectSet);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service response: " + responseString);
					dg = null;
				}      
				
				SingletonId propId;
				try
				{
					propId = (SingletonId)objObjectSet.Object[0].Property[0];
				}
				catch (Exception)
				{
					return "";				
				}
				objRootFolder.objectId = propId.Value;
				return objRootFolder.objectId;
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		private bool FolderExists(bool bLogRequestAndResponse, string traceFolderLocation, string parentFolderId, string folderName)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			// Specify the scope of the search
			try
			{
				ObjectStoreScope elemObjectStoreScope = new ObjectStoreScope();
				elemObjectStoreScope.objectStore = m_objectStore;         // Create the search for unfiled doc
				RepositorySearch elemRepositorySearch = new RepositorySearch();
				elemRepositorySearch.repositorySearchMode = RepositorySearchModeType.Rows;
				elemRepositorySearch.repositorySearchModeSpecified = true;
				elemRepositorySearch.SearchScope = elemObjectStoreScope;
				elemRepositorySearch.SearchSQL = "SELECT [Id] FROM [Folder] WHERE ([This] INFOLDER '" + parentFolderId + "' And [FolderName] = '" + folderName + "')"; 

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string requestString = dg.GetXmlFromGatewayObject(typeof(RepositorySearch), elemRepositorySearch);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service request: " + requestString);
					dg = null;
				}      

				// Invoke the ExecuteSearch operation
				ObjectSetType objObjectSet = wseService.ExecuteSearch(elemRepositorySearch);
				
				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string responseString = dg.GetXmlFromGatewayObject(typeof(ObjectSetType), objObjectSet);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service response: " + responseString);
					dg = null;
				}      

				SingletonId propId;
				try
				{
					propId = (SingletonId)objObjectSet.Object[0].Property[0];
				}
				catch (Exception)
				{
					return false;				
				}
				return true;
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}
		// PSC 11/01/2006 MAR994 - End

		// PSC 11/01/2006 MAR994 - Start
		public string GetFileNetFolder(bool bLogRequestAndResponse, string traceFolderLocation, string folderPath, string folderName)
		{		
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				traceFolder = traceFolderLocation;

				string revisedPath = folderPath.Replace("\\", "/");
				char [] folderSeperator = new char[] {'/'};
				revisedPath = revisedPath.TrimEnd(folderSeperator);
				revisedPath = revisedPath.TrimStart(folderSeperator);

				string folderId = string.Empty;

				folderId = CheckFolderExists(bLogRequestAndResponse, traceFolderLocation, revisedPath, folderName);
			
				return "/" + revisedPath + "/" + folderName;
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}
		// PSC 11/01/2006 MAR994 - End

		// PSC 11/01/2006 MAR994
		public string CreateFolder(bool bLogRequestAndResponse, string traceFolderLocation, string parentFolderId, string folderName)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				traceFolder = traceFolderLocation;

				string folderId = string.Empty;
				System.DateTime dateCreated = new System.DateTime();
				ObjectReference parentFolder = new ObjectReference();	// PSC 11/01/2006 MAR994		
				parentFolder.classId = "Folder";						// PSC 11/01/2006 MAR994
			
				parentFolder.objectId = parentFolderId;					// PSC 11/01/2006 MAR994
				// PSC 06/01/2006 MAR922 - End

				// Now check that the folderName does not exist
				// PSC 06/01/2006 MAR922
				// PSC 11/01/2006 MAR994
				if (FolderExists(bLogRequestAndResponse, traceFolderLocation, parentFolderId, folderName))
				{
					throw new Exception("The folder '" + folderName + "' already exists");
				}

				// Now create the folder
				CreateAction verbCreate = new CreateAction();
				verbCreate.classId = "Folder";          // Assign the action to the ChangeRequestType element
				ChangeRequestType elemChangeRequestType = new ChangeRequestType();
				elemChangeRequestType.Action = new ActionType[1];
				elemChangeRequestType.Action[0] = (ActionType)verbCreate; // Assign Create action
         
				// Specify the target object (an object store) for the actions
				elemChangeRequestType.TargetSpecification = new ObjectReference(); 
				elemChangeRequestType.TargetSpecification.classId = "ObjectStore";
				elemChangeRequestType.TargetSpecification.objectId = m_objectStore; 
				elemChangeRequestType.id = "1";         // Build a list of properties to set in the new doc
				ModifiablePropertyType[] elemInputProps = new ModifiablePropertyType[2];         // Specify and set a string-valued property for the FolderName property
				SingletonString propFolderName = new SingletonString();
				propFolderName.propertyId = "FolderName";
				propFolderName.Value = folderName;
				elemInputProps[0] = propFolderName; // Add to property list         // Create an object reference to the root folder
			
				// PSC 11/01/2006 MAR994
				parentFolder.objectStore = m_objectStore;         // Specify and set an object-valued property for the Parent property
				SingletonObject propParent = new SingletonObject();
				propParent.propertyId = "Parent"; 
				// PSC 11/01/2006 MAR994
				propParent.Value = (ObjectEntryType)parentFolder; // Set its value to the RootFolder object
				elemInputProps[1] = propParent; // Add to property list
         
				// Assign list of folder properties to set in ChangeRequestType element
				elemChangeRequestType.ActionProperties = elemInputProps;         // Build a list of properties to exclude on the new folder object that will be returned
				string[] excludeProps = new string[2];
				excludeProps[0] = "Owner";
				excludeProps[1] = "DateLastModified";         // Assign the list of excluded properties to the ChangeRequestType element
				elemChangeRequestType.RefreshFilter = new PropertyFilterType();
				elemChangeRequestType.RefreshFilter.ExcludeProperties = excludeProps;         // Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
				ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1];
				elemChangeRequestTypeArray[0] = elemChangeRequestType;         // Create ChangeResponseType element array 
				ChangeResponseType[] elemChangeResponseTypeArray = null;
         
				// Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
				ExecuteChangesRequest elemExecuteChangesRequest = new ExecuteChangesRequest(); 
				elemExecuteChangesRequest.ChangeRequest = elemChangeRequestTypeArray;
				elemExecuteChangesRequest.refresh = true; // return a refreshed object
				elemExecuteChangesRequest.refreshSpecified = true;
				
				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string requestString = dg.GetXmlFromGatewayObject(typeof(ExecuteChangesRequest), elemExecuteChangesRequest);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service request: " + requestString);
					dg = null;
				}      

				try
				{
					// Call ExecuteChanges operation to implement the folder creation
					elemChangeResponseTypeArray = wseService.ExecuteChanges(elemExecuteChangesRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An WebException occurred", ex);
					}
					throw new Exception("An exception occurred while creating a folder: [" + ex.Message + "]");
				}         // The new folder object should be returned, unless there is an error
				if (elemChangeResponseTypeArray==null || elemChangeResponseTypeArray.Length < 1)
				{
					throw new Exception("A valid object was not returned from the ExecuteChanges operation");

				}
				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string responseString = dg.GetXmlFromGatewayObject(typeof(ChangeResponseType[]), elemChangeResponseTypeArray);
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Web service response: " + responseString);
					dg = null;
				}      

				// Capture value of the FolderName property in the returned doc object
				foreach (PropertyType propProperty in elemChangeResponseTypeArray[0].Property)
				{
					// If property found, store its value
					if (propProperty.propertyId == "FolderName")
					{
						folderName = ((SingletonString)propProperty).Value;
						break;
					}
				}         // Capture value of the DateCreated property in the returned doc object
				foreach (PropertyType propProperty in elemChangeResponseTypeArray[0].Property)
				{
					// If property found, store its value
					if (propProperty.propertyId == "DateCreated")
					{
						dateCreated = ((SingletonDateTime)propProperty).Value;
						break;
					}
				}
				// PSC 11/01/2006 MAR994 - Start
				foreach (PropertyType propProperty in elemChangeResponseTypeArray[0].Property)
				{
					// If property found, store its value
					if (propProperty.propertyId == "Id")
					{
						folderId = ((SingletonId)propProperty).Value;
						break;
					}
				}

				return folderId;
				// PSC 11/01/2006 MAR994 - End
		
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		// Create and checkin in a document with inline content.
		// PSC 06/02/2006 MAR1197
		public string SaveFileToFolder(string documentName, string contentType, 
			bool bLogRequestAndResponse, string traceFolderLocation, string documentclass)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}
			
			traceFolder = traceFolderLocation;

			try
			{
				// PSC 06/02/2006 MAR1197
				string docId = SaveFile(documentName, contentType, bLogRequestAndResponse, documentclass);
				LinkFileToFolder(docId, bLogRequestAndResponse, documentclass);
				return docId;
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		// PSC 06/02/2006 MAR1197
		public string SaveFile(string documentName, string contentType, bool bLogRequestAndResponse, string documentclass)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				System.DateTime dateCreated = new System.DateTime();         
				
				// Build the Create action for a document
				CreateAction verbCreate = new CreateAction();
				verbCreate.classId = documentclass;          
				
				// Build the Checkin action
				CheckinAction verbCheckin = new CheckinAction();
				
				// PSC 11/01/2006 MAR994 - Start
				verbCheckin.checkinMinorVersion = false;
				verbCheckin.checkinMinorVersionSpecified = false;
				// PSC 11/01/2006 MAR994 - End
			
				// Assign the actions to the ChangeRequestType element
				ChangeRequestType elemChangeRequestType = new ChangeRequestType();
				elemChangeRequestType.Action = new ActionType[2];
				elemChangeRequestType.Action[0] = (ActionType)verbCreate; // Assign Create action
				elemChangeRequestType.Action[1] = (ActionType)verbCheckin; // Assign Checkin action
   
				// Specify the target object (an object store) for the actions
				elemChangeRequestType.TargetSpecification = new ObjectReference(); 
				elemChangeRequestType.TargetSpecification.classId = "ObjectStore";
				elemChangeRequestType.TargetSpecification.objectId = m_objectStore; 
				elemChangeRequestType.id = "1";         
				
				// Build a list of properties to set in the new doc
				// MAR1191 Start
				//ModifiablePropertyType[] elemInputProps = new ModifiablePropertyType[2];         
				ModifiablePropertyType[] elemInputProps = new ModifiablePropertyType[3];         
				// MAR1191 End
				
				// Specify and set a string-valued property for the DocumentTitle property
				SingletonString propDocumentTitle = new SingletonString();
				propDocumentTitle.Value = m_applicationNumber;
				propDocumentTitle.propertyId = "DocumentTitle"; 
				elemInputProps[0] = propDocumentTitle; // Add to property list         
				
				// Create an object reference to dependently persistable ContentTransfer object
				DependentObjectType objContentTransfer = new DependentObjectType();
				objContentTransfer.classId = "ContentTransfer";
				objContentTransfer.dependentAction = DependentObjectTypeDependentAction.Insert;
				objContentTransfer.dependentActionSpecified = true;
				objContentTransfer.Property = new PropertyType[2];         
				
				// Create reference to the object set of ContentTransfer objects 
				// returned by the Document.ContentElements property
				ListOfObject propContentElement = new ListOfObject();
				propContentElement.propertyId = "ContentElements";
				propContentElement.Value = new DependentObjectType[1];
				propContentElement.Value[0] = objContentTransfer;          

				InlineContent elemInlineContent = GetDocumentContent(m_stringData);
				
				// Create reference to Content pseudo-property
				ContentData propContent = new ContentData();
				propContent.Value = (ContentType)elemInlineContent;
				propContent.propertyId = "Content";         
				
				// Assign Content property to ContentTransfer object 
				objContentTransfer.Property[0] = propContent;         				
				
				// MAR1191 Start
				SingletonString propOmigaUniqueFileName = new SingletonString();
				propOmigaUniqueFileName.Value = documentName;
				propOmigaUniqueFileName.propertyId = "OmigaUniqueFileName"; 
				elemInputProps[1] = propOmigaUniqueFileName; // Add to property list         
				// MAR1191 End

				// Create and assign ContentType string-valued property to ContentTransfer object
				SingletonString propContentType = new SingletonString();
				propContentType.propertyId = "ContentType";
				// Set MIME-type depending on content type
				// PSC 06/02/2006 MAR1197 - Start
				switch (contentType.ToUpper())
				{
					case "RTF":
					{
						propContentType.Value = "application/rtf";
						break;
					}
					case "DOC":
					{
						propContentType.Value = "application/msword";
						break;
					}
					default:
					{
						propContentType.Value = "application/pdf";
						break;
					}
				}
				// PSC 06/02/2006 MAR1197 - End

				objContentTransfer.Property[1] = propContentType;
				// MAR1191 Start
				//elemInputProps[1] = propContentElement;         
				elemInputProps[2] = propContentElement;         
				// MAR1191 End

				// Assign list of document properties to set in ChangeRequestType element
				elemChangeRequestType.ActionProperties = elemInputProps;         
				
				// Build a list of properties to exclude on the new doc object that will be returned
				string[] excludeProps = new string[2];
				excludeProps[0] = "Owner";
				excludeProps[1] = "DateLastModified";         
				
				// Assign the list of excluded properties to the ChangeRequestType element
				elemChangeRequestType.RefreshFilter = new PropertyFilterType();
				elemChangeRequestType.RefreshFilter.ExcludeProperties = excludeProps;         
				
				// Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
				ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1];
				elemChangeRequestTypeArray[0] = elemChangeRequestType;         
				
				// Create ChangeResponseType element array 
				ChangeResponseType[] elemChangeResponseTypeArray = null;
   
				// Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
				ExecuteChangesRequest elemExecuteChangesRequest = new ExecuteChangesRequest(); 
				elemExecuteChangesRequest.ChangeRequest = elemChangeRequestTypeArray;
				elemExecuteChangesRequest.refresh = true; // return a refreshed object
				elemExecuteChangesRequest.refreshSpecified = true;

				// RF 21/02/2006 MAR1196 Start - Follow C# logging standards							
//				if (bLogRequestAndResponse)
//				{
//					Logger lr = new Logger(System.Reflection.MethodInfo.GetCurrentMethod().Name);
//					lr.LogRequest(elemExecuteChangesRequest);
//				}
   
				if(_logger.IsDebugEnabled) 
				{
					string message = MethodInfo.GetCurrentMethod().Name + 
							": Web service request: ";
					try
					{
						DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
						string request = dg.GetXmlFromGatewayObject(typeof(ExecuteChangesRequest), elemExecuteChangesRequest);
						dg = null;

						// don't log binary data 
						int x = request.IndexOf("<Binary>");
						int y = request.LastIndexOf("</Binary>");
						if (x > 0 && y > 0 )
						{
							System.Text.StringBuilder sb = new System.Text.StringBuilder(
								request.Substring(0, x + 8), request.Length);
							sb.Append("*** Binary data removed ***");
							sb.Append(request.Substring(y, request.Length - y));
							message += sb.ToString();
						}
						else
						{
							message += request;
						}
					}
					catch(Exception ex)
					{
						message += "Failed to get request as string: " +ex.Message;
					}
					_logger.Debug(message);
				}
				// RF 21/02/2006 MAR1196 End

				try
				{
					// Call ExecuteChanges operation to implement the doc creation and checkin
					elemChangeResponseTypeArray = wseService.ExecuteChanges(elemExecuteChangesRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("An exception occurred while creating a document: [" + ex.Message + "]");
				}

				// RF 21/02/2006 MAR1196 Start - Follow C# logging standards							
//				if (bLogRequestAndResponse)
//				{
//					Logger lr = new Logger(System.Reflection.MethodInfo.GetCurrentMethod().Name);
//					lr.LogResponse(elemChangeResponseTypeArray);
//				}
				
				// RF 21/02/2006 MAR1196 End

				// The new document object should be returned, unless there is an error
				if (elemChangeResponseTypeArray==null || elemChangeResponseTypeArray.Length < 1)
				{
					throw new Exception("A valid object was not returned from the ExecuteChanges operation");
				}
				
				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string responseString  = dg.GetXmlFromGatewayObject(typeof(ChangeResponseType[]), elemChangeResponseTypeArray);
					dg = null;
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + responseString );
				}

        
				// get doc id
				string newDocId=""; // id of newly created doc

				try
				{				
					// RF 21/02/2006 MAR1196 Start - Doc id not always returned as a property.

					// Capture value of the DocumentTitle property in the returned doc object
//					foreach (PropertyType propProperty in elemChangeResponseTypeArray[0].Property)
//					{
//						// If property found, store its value
//						if (propProperty.propertyId == "Id")
//						{
//							newDocId = ((SingletonId)propProperty).Value;
//							break;
//						}
//					}         

					newDocId = elemChangeResponseTypeArray[0].objectId;

					// RF 21/02/2006 MAR1196 End
				}
				catch(Exception ex)
				{
					throw new Exception("Object id not found: " + ex.Message);
				}

				if (newDocId == "")
				{
					throw new Exception("A valid doc id was not returned from the ExecuteChanges operation");
				}		

				return newDocId; 
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}

		public void LinkFileToFolder(string docId, bool bLogRequestAndResponse, string documentclass)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Starting"); 
			}

			try
			{
				string containmentName = "";
				System.DateTime dateCreated = new System.DateTime();
   
				// Build the Create action for a DynamicReferentialContainmentRelationship object
				CreateAction verbCreate = new CreateAction();
				verbCreate.classId = "DynamicReferentialContainmentRelationship";          
				
				// Assign the action to the ChangeRequestType element
				ChangeRequestType elemChangeRequestType = new ChangeRequestType();
				elemChangeRequestType.Action = new ActionType[1];
				elemChangeRequestType.Action[0] = (ActionType)verbCreate; // Assign Create action
   
				// Specify the target object (an object store) for the actions
				elemChangeRequestType.TargetSpecification = new ObjectReference(); 
				elemChangeRequestType.TargetSpecification.classId = "ObjectStore";
				elemChangeRequestType.TargetSpecification.objectId = m_objectStore; 
				elemChangeRequestType.id = "1";         
				
				// Build a list of properties to set in the new doc
				ModifiablePropertyType[] elemInputProps = new ModifiablePropertyType[3]; 
        
				// Specify and set a string-valued property for the ContainmentName property
				SingletonString propContainmentName = new SingletonString(); 
				propContainmentName.propertyId = "ContainmentName";	
				propContainmentName.Value = docId;
				elemInputProps[0] = propContainmentName; // Add to property list         // Specify the scope of the search
				
				// Create an object reference to the document to file
				ObjectReference objDocument = new ObjectReference();
				objDocument.classId = documentclass;
				objDocument.objectStore = m_objectStore;
				objDocument.objectId = docId;         
				
				// Create an object reference to the folder in which to file the document
				ObjectSpecification objFolder = new ObjectSpecification();
				objFolder.classId = "Folder";
				objFolder.objectStore = m_objectStore;

				// PSC 11/01/2006 MAR994 - Start
				// Get folder path from global parameters
				GlobalParameter folderLocation = new GlobalParameter("FileNetFolder");
				objFolder.path = GetFileNetFolder(bLogRequestAndResponse, traceFolder, folderLocation.String, m_applicationNumber);
				// PSC 11/01/2006 MAR994 - End

				// Specify and set an object-valued property for the Head property
				SingletonObject propHead = new SingletonObject();
				propHead.propertyId = "Head"; 
				propHead.Value = (ObjectEntryType)objDocument; 
				
				// Set its value to the Document object
				elemInputProps[1] = propHead; // Add to property list         
				
				// Specify and set an object-valued property for the Tail property
				SingletonObject propTail = new SingletonObject();
				propTail.propertyId = "Tail"; 
				propTail.Value = (ObjectEntryType)objFolder; // Set its value to the folder object
				elemInputProps[2] = propTail; // Add to property list
   
				// Assign list of DRCR properties to set in ChangeRequestType element
				elemChangeRequestType.ActionProperties = elemInputProps;         
				
				// Build a list of properties to exclude on the new DRCR object that will be returned
				string[] excludeProps = new string[2];
				excludeProps[0] = "Owner";
				excludeProps[1] = "DateLastModified";         
				
				// Assign the list of excluded properties to the ChangeRequestType element
				elemChangeRequestType.RefreshFilter = new PropertyFilterType();
				elemChangeRequestType.RefreshFilter.ExcludeProperties = excludeProps;         
				
				// Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
				ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1];
				elemChangeRequestTypeArray[0] = elemChangeRequestType;         
				
				// Create ChangeResponseType element array 
				ChangeResponseType[] elemChangeResponseTypeArray = null;
   
				// Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
				ExecuteChangesRequest elemExecuteChangesRequest = new ExecuteChangesRequest(); 
				elemExecuteChangesRequest.ChangeRequest = elemChangeRequestTypeArray;
				elemExecuteChangesRequest.refresh = true; // return a refreshed object
				elemExecuteChangesRequest.refreshSpecified = true;         
				
				if (bLogRequestAndResponse)
				{
					Logger lr = new Logger(System.Reflection.MethodInfo.GetCurrentMethod().Name);
					lr.LogRequest(elemExecuteChangesRequest);
				}

				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string requestString  = dg.GetXmlFromGatewayObject(typeof(ExecuteChangesRequest), elemExecuteChangesRequest);
					dg = null;
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service request: " + requestString );
				}

				try
				{
					// Call ExecuteChanges operation to implement the DRCR creation
					elemChangeResponseTypeArray = wseService.ExecuteChanges(elemExecuteChangesRequest);
				}
				catch(System.Net.WebException ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("An exception occurred while creating a folder: [" + ex.Message + "]");
				}         
				
				// The new DRCR object should be returned, unless there is an error
				if (elemChangeResponseTypeArray==null || elemChangeResponseTypeArray.Length < 1)
				{
					throw new Exception("A valid object was not returned from the ExecuteChanges operation");
				}
				
				if(_logger.IsDebugEnabled) 
				{
					DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
					string responseString  = dg.GetXmlFromGatewayObject(typeof(ChangeResponseType[]), elemChangeResponseTypeArray);
					dg = null;
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + 
						": Web service response: " + responseString );
				}
				
				if (bLogRequestAndResponse)
				{
					Logger lr = new Logger(System.Reflection.MethodInfo.GetCurrentMethod().Name);
					lr.LogResponse(elemChangeResponseTypeArray);
				}

	
				// Capture value of the ContainmentName property in the returned doc object
				foreach (PropertyType propProperty in elemChangeResponseTypeArray[0].Property)
				{
					// If property found, store its value
					if (propProperty.propertyId == "ContainmentName")
					{
						containmentName = ((SingletonString)propProperty).Value;
						break;
					}
				}         
			}
			finally
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Completed"); 
				}
			}		
		}		

		public InlineContent GetDocumentContent(string stringData) 
		{
			byte[] binaryData = BinaryHelper.ConvertToByteArray(stringData);

			// Read data stream from file containing the document content
			InlineContent ic = new InlineContent();

			ic.Binary = (binaryData);

			return ic;
		}

		public class BinaryHelper
		{
			// remove any empty spaces and "\r\n"
			private static string FixBase64(string data)
			{
				System.Text.StringBuilder sbText = new System.Text.StringBuilder(data,data.Length);
				sbText.Replace("\r\n", String.Empty);
				sbText.Replace(" ", String.Empty);
				return sbText.ToString();
			}

			public static Byte[] ConvertToByteArray(string sData)
			{
				Byte[] bytArrayData = new Byte[sData.Length];

				if (sData.Length > 0)
				{
					sData = FixBase64(sData);
					bytArrayData = Convert.FromBase64String(sData);
				}
				return bytArrayData; 
			}
		}

		public class Logger
		{
			public Logger(string operation)
			{
				m_operation = operation;
			}

			private string m_operation;

			public void LogRequest(ExecuteChangesRequest request)
			{
				try
				{
					string fileOut = traceFolder + m_operation + "Request.xml";
					XmlSerializer serOut = new XmlSerializer(typeof(ExecuteChangesRequest ));
					XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
					serOut.Serialize(writer, request);
					writer.Close();
				}
				catch(Exception ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("An exception occurred writing to the log file." + ex.Message + "]");
				}
			}

			public void LogResponse(ChangeResponseType[] response)
			{
				try
				{
					string fileOut = traceFolder + m_operation + "Response.xml";
					XmlSerializer serOut = new XmlSerializer(typeof(ChangeResponseType[]));
					XmlWriter writer = new XmlTextWriter(fileOut, System.Text.Encoding.UTF8);          
					serOut.Serialize(writer, response);
					writer.Close();
				}
				catch(Exception ex)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(MethodInfo.GetCurrentMethod().Name + 
							": An exception occurred", ex);
					}
					throw new Exception("An exception occurred writing to the log file." + ex.Message + "]");
				}
			}
		}
	}		
}  


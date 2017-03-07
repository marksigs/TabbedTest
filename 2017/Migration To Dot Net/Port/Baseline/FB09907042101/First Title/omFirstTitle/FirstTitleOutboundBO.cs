/*
--------------------------------------------------------------------------------------------
Workfile:			FirstTitleOutboundBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		17/10/2005	MAR182	Code Amendments
GHun	09/11/2005	MAR182	Use omCore to get db connection string, fixed hardcoded XSLT file path, error messages
GHun	10/11/2005	MAR508	Changed stored procedure parameter code to use data length
HM		12/11/2005	MAR530	Change generated classes to web references, correct namespaces. 
HMA     16/11/2005  MAR530  Further changes. Renamed RunFTContinueOfferInterface to RunFTReIssueOfferInterface.
GHun	30/11/2005	MAR609	Fixed numerous errors throughout
GHun	09/12/2005	MAR835	Changed onMessage to set request attributes if InterfaceAvailability fails
GHun	12/12/2005	MAR852	set Operator in generic request
JD		14/12/2005	MAR783  Check for no rows returned in GetApplicationFirstTitle
GHun	15/12/2005	MAR855	Changed onMessage and added ProcessOutboundMessage
PSC		10/01/2006	MAR961	Use omLogging wrapper
GHun	16/01/2006	MAR1036	Changed timestamps to use DateTime.Now instead of DateTime.Today
HMA     23/01/2006  MAR1048 Save outbound messages.
HMA     01/02/2006  MAR1159 Correct RunFTCancelAppInterface
RF		09/03/2006  MAR1392 Writing MOF message to ApplicationFirstTitle table is timing out 
PSC		24/03/2006	MAR1521	Move instantiation of _logger to before first logging statement
GHun	27/03/2006  MAR1539 Changed ProcessOutboundMessage to instantiate _logger
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
GHun	26/04/2006	MAR1646 Changed ProcessFirstTitleOutboundResponse to pass ChannelId to HandleInterfaceResponse
HMA     12/06/2006  MAR527  Add code to check Interface Availability.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Runtime.InteropServices;
using System.Net;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Xml.Xsl;
using System.ComponentModel;
using System.EnterpriseServices; // MAR1392
//using NUnit.Framework;

//COM Ref
using omMQ;
using omApp;
using omTM;
using MsgTm;
using MQL = MESSAGEQUEUECOMPONENTVCLib;

using omigaException = Vertex.Fsd.Omiga.Core.OmigaException;
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 10/01/2006 MAR961

//Web Ref  MAR530
using MakeOffer = Vertex.Fsd.Omiga.omFirstTitle.MakeOffer;
using RequestInsurance = Vertex.Fsd.Omiga.omFirstTitle.RequestInsurance;
using MiscInstruction = Vertex.Fsd.Omiga.omFirstTitle.MiscellaneousInstruction;

//Class Ref
using GeneralManager = Vertex.Fsd.Omiga.omFirstTitle.General.GeneralManager;
using ConfigDataManager = Vertex.Fsd.Omiga.omFirstTitle.ConfigData.ConfigDataManager;
using XMLManager = Vertex.Fsd.Omiga.omFirstTitle.XML.XMLManager;

namespace Vertex.Fsd.Omiga.omFirstTitle
{
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[ProgId("omFirstTitle.FirstTitleOutboundBO")]
	[ComVisible(true)]
	[Guid("DA086385-CF78-46bd-AE9C-9688BBA8F441")]
	[Transaction(TransactionOption.Supported)] // MAR1392 
	public class FirstTitleOutboundBO : 
		ServicedComponent, // MAR1392   
		MQL.IMessageQueueComponentVC2 
	{		
		string _UserID;
		string _UnitID;
		string _UserAuthorityLevel;
		string _ChannelID;
		string _Password;
		const string _InterfaceType = "First Title";
		//const string _ClassName = "omFirstTitle.FirstTitleOutboundBO";

		XMLManager _xmlMgr = new XMLManager();
		ConfigDataManager _cdMgr = new ConfigDataManager();
		GeneralManager _genMgr = new GeneralManager();

		// PSC 10/01/2006 MAR961 - Start
		private omLogging.Logger _logger = null;
		private string _processContext = "Default";
			
		public FirstTitleOutboundBO()
		{
			// MAR1392 Moved into Execute  
			//_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}
	
		// MAR1392 Start - constructor not supported for serviced component
//		public FirstTitleOutboundBO(string processContext)
//		{
//			if (processContext == null || processContext.Length == 0)
//			{
//				throw new ArgumentNullException("Process context has not been supplied");
//			}
//			_processContext = processContext;
//			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
//		}
		// MAR1392 End
		// PSC 10/01/2006 MAR961 - End
			
		public string UserID
		{
			get { return _UserID; }
			set { _UserID = value; }
		}

		public string UnitID
		{
			get { return _UnitID; }
			set { _UnitID = value; }
		}

		public string UserAuthorityLevel
		{
			get { return _UserAuthorityLevel; }
			set { _UserAuthorityLevel = value; }
		}
		
		public string ChannelID
		{
			get { return _ChannelID; }
			set { _ChannelID = value; }
		}
		
		public string Password
		{
			get { return _Password; }
			set { _Password = value; }
		}

		//MAR1539 GHun
		//// MAR1392 
		//public string ProcessContext
		//{
		//	get { return _processContext; }
		//	set { _processContext = value; }
		//}

		public string Execute(string strRequest)
		{
			const string cFunctionName = "Execute";

			// MAR1392  
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName + ": request = " + strRequest);
			}

			XmlDocument requestXML = new XmlDocument();

			XmlNode requestNode;
			XmlDocument responseXML = new XmlDocument();
			
			XmlNode tempNode;
			XmlNode responseNode;
			XmlNodeList operationNodeList;
			
			XmlAttribute typeAttrib;

			// Create the response node ...
			tempNode = responseXML.CreateElement("RESPONSE");
			responseNode = responseXML.AppendChild(tempNode);	
			typeAttrib = responseXML.CreateAttribute("TYPE");
			responseNode.Attributes.SetNamedItem(typeAttrib);

			string response = "";
			try
			{
				requestXML.LoadXml(strRequest);
				//MAR535 Debug
				//requestXML.Save ("c:\\RUNFTINITIALREQUESTINTERFACE.XML");
				
				requestNode = _xmlMgr.xmlGetMandatoryNode(requestXML, "/REQUEST");
				
				_UserID				= _xmlMgr.xmlGetAttributeText(requestNode, "USERID");
				_UnitID				= _xmlMgr.xmlGetAttributeText(requestNode, "UNITID"); 
				_UserAuthorityLevel = _xmlMgr.xmlGetAttributeText(requestNode, "USERAUTHORITYLEVEL"); 
				_ChannelID			= _xmlMgr.xmlGetAttributeText(requestNode, "CHANNELID");
				_Password			= _xmlMgr.xmlGetAttributeText(requestNode, "PASSWORD");

				if (requestNode.Attributes.GetNamedItem("OPERATION") == null)
				{
					operationNodeList = requestNode.SelectNodes("OPERATION");
					foreach (XmlNode xndOperation in operationNodeList)
					{
						foreach (XmlAttribute xmlAttrib in requestNode.Attributes)
						{
							xndOperation.Attributes.SetNamedItem(xmlAttrib.CloneNode(true));
						}
						ForwardRequest(xndOperation, responseNode);
					}
				}
				else
				{
					ForwardRequest(requestNode, responseNode);
				}
							
				typeAttrib.InnerText = "SUCCESS";
				response = responseXML.OuterXml;

			}
			catch(Exception ex)
			{	
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}

				// Create the response node ...
				//MAR182 GHun
				response = new omigaException(ex).ToOmiga4Response();
			}
			finally
			{
				requestXML = null;
				requestNode = null;
				responseXML = null;
				tempNode = null;
				typeAttrib = null;
				responseNode = null;	
				operationNodeList = null;
			}

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Exiting " + cFunctionName + ": response = " + response);
			}

			return response;
		}
		
		private void ForwardRequest(XmlNode requestNode, XmlNode responseNode)
		{
			const string cFunctionName = "ForwardRequest";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try
			{
				string strOperation = (requestNode.Name == "REQUEST") ? requestNode.Attributes.GetNamedItem("OPERATION").InnerText :
					requestNode.Attributes.GetNamedItem("NAME").InnerText;

				strOperation = strOperation.ToUpper();
				switch (strOperation)
				{
					case "RUNFTINITIALREQUESTINTERFACE":
						RunFTInitialRequestInterface(requestNode, responseNode);
						break;
					case "RUNFTOFFERREQUESTINTERFACE":
						RunFTOfferRequestInterface(requestNode, responseNode);
						break;
					case "RUNFTREISSUEOFFERINTERFACE":
						RunFTReissueOfferInterface(requestNode, responseNode);
						break;
					case "RUNFTCANCELAPPINTERFACE":
						RunFTCancelAppInterface(requestNode, responseNode);
						break;
					case "RUNFTFUNDSRELEASEFAILINTERFACE":
						RunFTFundsReleaseFailInterface(requestNode, responseNode);
						break;
					case "RUNFTMISCUPDATEINTERFACE":
						RunFTMiscUpdateInterface(requestNode, responseNode);
						break;
					case "RUNFTCOMPDATEUPDATEINTERFACE":	//MAR690 GHun interfaceName has to be 30 chars or less
						RunFTCompletionDateUpdateInterface(requestNode, responseNode);
						break;
					case "RUNFTOFFEREXPIRYINTERFACE":
						RunFTOfferExpiryInterface(requestNode, responseNode);
						break;
					case "RUNFTREINSTATEAPPINTERFACE":
						RunFTReInstateAppInterface(requestNode, responseNode);
						break;
					case "ONMESSAGE":
						responseNode = requestNode;
						//OnMessage(requestNode.OuterXml, responseNode.OuterXml);
						break;
					case "GETAPPLICATIONFIRSTTITLE":
						GetApplicationFirstTitle(requestNode, responseNode, "");
						break;
					default:
						break;
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException("Unexpected error occurred", ex);
			}
		}

		private void RunFTInitialRequestInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTInitialRequestInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlDocument xdDoc = new XmlDocument();

			try
			{
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");

				GetRunFTInitialRequestInterfaceData(xnRequest, xnResponse);
				
				//MAR1048
				XmlNode xnAppNode = _xmlMgr.xmlGetMandatoryNode(xnRequest, "//APPLICATION");
				string ApplicationNumber = _xmlMgr.xmlGetAttributeText(xnAppNode, "APPLICATIONNUMBER");

				//Prepare and send to Queue
				if (PrepareFTInitialRequestToMessageQueue(xnResponse))
				{
					// If success then save message and update the casetask with status "COMPLETED"
					
					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "RFI";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}
				
				// Create the response node
				XmlAttribute typeAttrib;
				typeAttrib = xdDoc.CreateAttribute("TYPE");
				typeAttrib.InnerText = "SUCCESS";
				xnResponse.Attributes.SetNamedItem(typeAttrib);
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}
	
		private void GetRunFTInitialRequestInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetRunFTInitialRequestInterfaceData";
			
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlDocument xdResponseDoc = new XmlDocument();
			StringWriter sw = new StringWriter();
			XmlDocument xmlDoc = new XmlDocument();

			try 
			{
				XmlNode xnAppNode = _xmlMgr.xmlGetMandatoryNode(xnRequest, "//APPLICATION");

				string strApplicationNumber = _xmlMgr.xmlGetAttributeText(xnAppNode, "APPLICATIONNUMBER");
				string strApplicationFactFindNumber = _xmlMgr.xmlGetAttributeText(xnAppNode, "APPLICATIONFACTFINDNUMBER");
			
				string strConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(strConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTINITIALREQUESTDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, strApplicationNumber.Length);
					oParam.Value = strApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(strApplicationFactFindNumber, 10);
				
					oConn.Open();
					
					SqlDataReader oReader = oComm.ExecuteReader();
					StringBuilder sb = new StringBuilder("<RESPONSE>");

					do 
					{
						while (oReader.Read())
						{
							sb.Append (oReader.GetString(0));
						}
					} while (oReader.NextResult());

					sb.Append ("</RESPONSE>");
					
					xmlDoc.LoadXml(sb.ToString());
					//MAR535 debug
					//xmlDoc.Save("c:\\FTInitialRequestData.xml");
					XslTransform xsltFile = new XslTransform();
					//MAR182 GHun
					//xsltFile.Load("C:\\Projects\\INGUK\\1 DEV Code\\First Title\\omFirstTitle\\bin\\Debug\\FTInitialRequestData.xslt");
					xsltFile.Load(GetXmlPath() + "\\" + "FTInitialRequestData.xslt");

					xsltFile.Transform(xmlDoc.FirstChild, null, sw, null);
					xdResponseDoc.LoadXml(sw.ToString());
					
					XmlNode ReponseNode = xnResponse.OwnerDocument.ImportNode(xdResponseDoc.SelectSingleNode("//RequestInsuranceRequestType"), true);
					
					XmlAttribute typeAttrib;
					typeAttrib = xmlDoc.CreateAttribute("APPLICATIONNUMBER");
					typeAttrib.InnerText = strApplicationNumber;
					xnResponse.Attributes.SetNamedItem(typeAttrib);

					xnResponse.AppendChild(ReponseNode);
					
					sw.Flush();

					oReader.Close();
				}	
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}

		private void GetApplicationFirstTitle(XmlNode xnRequest, XmlNode xnResponse, string messageType)
		{
			const string cFunctionName = "GetApplicationFirstTitle";
		
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				//Select Mandatory Node
				XmlNode xnAppNode = _xmlMgr.xmlGetMandatoryNode(xnRequest, "//REQUEST/APPLICATION");
				
				//Check the mandatory Attributes
				string ApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONNUMBER");
				
				string ApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONFACTFINDNUMBER");
				
				if (messageType.Length == 0)
					messageType = _xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "MESSAGETYPE");

				if (ApplicationNumber.Length == 0 || messageType.Length == 0 || ApplicationFactFindNumber.Length == 0)
				{
					throw new ArgumentNullException();
				}

				string strConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(strConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETLATESTAPPLICATIONFIRSTTITLE", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(ApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, messageType.Length);
					oParam.Value = messageType;
			
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					StringBuilder sb = new StringBuilder("<RESPONSE>");
					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument xdDoc = new XmlDocument();
					xdDoc.LoadXml(sb.ToString());
				
					if(xdDoc.SelectSingleNode("//APPLICATIONFIRSTTITLE") != null) //Check we have some rows returned. JD MAR783
					{
						XmlNode ReponseNode = xnResponse.OwnerDocument.ImportNode(xdDoc.SelectSingleNode("//APPLICATIONFIRSTTITLE"), true);
						xnResponse.AppendChild(ReponseNode);
					}
					oReader.Close();
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
		}

		private void RunFTOfferRequestInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTOfferRequestInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
				//_logger.Debug("Transaction ID: " + ContextUtil.TransactionId.ToString()); 
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			string Response = string.Empty;
			XmlDocument xdDoc = new XmlDocument();
			XmlDocument xdReqDoc = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Load into a DomDocumnet
				GetRunFTOfferRequestInterfaceData(xnRequest, xnResponse);
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType = _xmlMgr.xmlGetAttributeText(appFirstTitle, "MESSAGETYPE");
				string ApplicationNumber = _xmlMgr.xmlGetAttributeText(appFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");

				string MortgageAdvance = xnResponse.SelectSingleNode("//RESPONSE/MORTGAGESUBQUOTE/@AMOUNTREQUESTED").Value;
				MortgageAdvance  = Convert.ToString (Convert.ToInt32(MortgageAdvance) * 100);

				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("MakeOfferRequestType");
				xdReqDoc.AppendChild(xnRootNode);

				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MOF";
				xnHeader.AppendChild(xnNode);
				
				// Seq No
				xnNode = xdReqDoc.CreateElement("SequenceNo");
				xnNode.InnerText = "";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Mortgage Advance
				xnNode = xdReqDoc.CreateElement("MortgageAdvance");
				xnNode.InnerText  = MortgageAdvance;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, 
					true, false, "RUNFTOFFERREQUESTINTERFACE", ApplicationNumber, 
					_UserID, _UnitID, _UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					// If success then save the message and update the casetask with status "COMPLETED"

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MOF";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}

				// Create the response node
				XmlAttribute typeAttrib;
				typeAttrib = xdDoc.CreateAttribute("TYPE");
				typeAttrib.InnerText = "SUCCESS";
				xnResponse.Attributes.SetNamedItem(typeAttrib);
			}
			catch (Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
		}

		private void GetRunFTOfferRequestInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetRunFTOfferRequestInterfaceData";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlDocument xmlIndoc = new XmlDocument();

			try
			{
				//Select Ma
				XmlNode xmlAppNode = xnRequest.SelectSingleNode("//APPLICATION");
				
				//Check the mandatory Attributes
				string ApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONNUMBER");
				
				string ApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONFACTFINDNUMBER");
				
				string messageType = "ACA";
				
				if (ApplicationNumber.Length == 0 || messageType.Length == 0 || ApplicationFactFindNumber.Length == 0)
				{
					throw new ArgumentNullException();
				}

				string ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTOFFERREQUESTINTERFACEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(ApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, messageType.Length);
					oParam.Value = messageType;
			
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();

					StringBuilder sb = new StringBuilder("<RESPONSE>");
					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument ResponseDoc = new XmlDocument();
					ResponseDoc.LoadXml(sb.ToString());
				
					XmlNode xmlTemp = xnResponse.OwnerDocument.ImportNode(ResponseDoc.FirstChild, true);
					xnResponse.AppendChild(xmlTemp);

					oReader.Close();
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
			finally
			{
				xmlIndoc = null;
			}
		}

		private void RunFTReissueOfferInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTReissueOfferInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			string Response = string.Empty;
			XmlDocument xdDoc = new XmlDocument();
			XmlDocument xdReqDoc = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Load into a DomDocumnet
				GetRunFTOfferRequestInterfaceData(xnRequest, xnResponse);
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType = _xmlMgr.xmlGetAttributeText(appFirstTitle, "MESSAGETYPE");
				string ApplicationNumber = _xmlMgr.xmlGetAttributeText(appFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");

				string MortgageAdvance = xnResponse.SelectSingleNode("//RESPONSE/MORTGAGESUBQUOTE/@AMOUNTREQUESTED").Value;
				MortgageAdvance = Convert.ToString (Convert.ToInt32(MortgageAdvance) * 100);

				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("MakeOfferRequestType");
				xdReqDoc.AppendChild(xnRootNode);

				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//Timestamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MOR";
				xnHeader.AppendChild(xnNode);
				
				// SequenceNo
				xnNode = xdReqDoc.CreateElement("SequenceNo");
				xnNode.InnerText = "";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Mortgage Advance
				xnNode = xdReqDoc.CreateElement("MortgageAdvance");
				xnNode.InnerText  = MortgageAdvance;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTREISSUEOFFERINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					// If success then update the casetask with status "COMPLETED"
					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}
				
				// Create the response node
				XmlAttribute typeAttrib;
				typeAttrib = xdDoc.CreateAttribute("TYPE");
				typeAttrib.InnerText = "SUCCESS";
				xnResponse.Attributes.SetNamedItem(typeAttrib);
			}
			catch (Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
		}

		private void RunFTCancelAppInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTCancelAppInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			XmlDocument xdDoc = new XmlDocument();
			string ExceptionReasonCodeText = string.Empty;
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data and load into a DomDocumnet
				GetRunFTCancelAppInterfaceData(xnRequest, xnResponse);
				
				XmlNode xnAppFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType =  _xmlMgr.xmlGetAttributeText(xnAppFirstTitle, "MESSAGETYPE");
				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(xnAppFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(xnAppFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(xnAppFirstTitle, "CONVEYANCERREFERENCE");
				
				XmlNode xnGlobalParam = xnResponse.SelectSingleNode("//RESPONSE/GLOBALPARAMETER");
				string TMCancelStateID = _xmlMgr.xmlGetAttributeText(xnGlobalParam, "STRING");
				string CurrentStageID = _xmlMgr.xmlGetAttributeText(xnCaseTask, "STAGEID");
				//MAR1159
				if (CurrentStageID == TMCancelStateID) 
				{
					XmlNode xnCombo = xnResponse.SelectSingleNode("//RESPONSE/CASESTAGE/COMBOVALUE");
					if (xnCombo != null)
						ExceptionReasonCodeText = _xmlMgr.xmlGetAttributeText(xnCombo, "VALUENAME");
					else
						ExceptionReasonCodeText = "";
				}

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);

				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//Timestamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "CAN";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = ExceptionReasonCodeText;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTCANCELAPPINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					// If success then save the message and update the casetask with status "COMPLETED"

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MSI";
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGESUBTYPE");
					xnTempNode.InnerText = "CAN";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}

				// Create the response node
				XmlAttribute typeAttrib;
				typeAttrib = xdDoc.CreateAttribute("TYPE");
				typeAttrib.InnerText = "SUCCESS";
				xnResponse.Attributes.SetNamedItem(typeAttrib);

			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
			finally
			{
				xnNode = null;
				requestDocument = null;
				xdDoc  = null;
				xdResponse  = null;
			}
		}

		private void GetRunFTCancelAppInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetRunFTCancelAppInterfaceData";
			
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				//Select Ma
				XmlNode xmlAppNode = xnRequest.SelectSingleNode("//APPLICATION");
							
				//Check the mandatory Attributes
				string ApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONNUMBER");
				
				string ApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONFACTFINDNUMBER");
				
				string messageType = "ACA";
							
				if (ApplicationNumber.Length == 0 || messageType.Length == 0 || ApplicationFactFindNumber.Length == 0)
				{
					throw new ArgumentNullException();
				}

				string ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTCANCELAPPINTERFACEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					//MAR609 GHun fix parameter lengths
					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(ApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, messageType.Length);
					oParam.Value = messageType;
					//MAR609 End
						
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					StringBuilder sb = new StringBuilder("<RESPONSE>");

					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument ResponseDoc = new XmlDocument();
					ResponseDoc.LoadXml(sb.ToString());
				
					XmlNode xmlTemp = xnResponse.OwnerDocument.ImportNode(ResponseDoc.FirstChild, true);
					xnResponse.AppendChild(xmlTemp);

					oReader.Close();
				}	
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ":" + ex.Message);
			}
		}

		private void RunFTFundsReleaseFailInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTFundsReleaseFailInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data and load into a DomDocumnet
				GetFTFundsReleaseFailInterfaceData(xnRequest, xnResponse);
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType = _cdMgr.GetComboValueIDForValidationType("FirstTitleMessageType", "ACA"); //MAR609 GHun
				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");

				XmlNode taskNote = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/TASKNOTE");
				string noteEntry = _xmlMgr.xmlGetAttributeText(taskNote, "NOTEENTRY");

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);
				
				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "NOR";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = noteEntry;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTFUNDSRELEASEFAILINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}

			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
		}

		private void GetFTFundsReleaseFailInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetFTFundsReleaseFailInterfaceData";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				//Select Mandatorynodes
				XmlNode xnAppNode = xnRequest.SelectSingleNode("//APPLICATION");
				
				XmlNode xnCaseTaskNode = xnRequest.SelectSingleNode("//CASETASK");
				
				//Check the mandatory Attributes
				string ApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONNUMBER");
				
				string ApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONFACTFINDNUMBER");
				
				string messageType = "ACA";	//MAR609 GHun
				
				string CaseActivityGUID =
					_xmlMgr.xmlGetMandatoryAttributeText(xnCaseTaskNode, "CASEACTIVITYGUID");

				string StageID = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnCaseTaskNode, "STAGEID");

				string CaseStageSeqNo = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnCaseTaskNode, "CASESTAGESEQUENCENO");

				string TaskID = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnCaseTaskNode, "TASKID");
				
				string TaskInstance =
					_xmlMgr.xmlGetMandatoryAttributeText(xnCaseTaskNode, "TASKINSTANCE");

				string ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTFUNDSRELEASEFAILINTERFACEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
					
					//MAR609 GHun fix parameter lengths
					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(ApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, messageType.Length);
					oParam.Value = messageType;
				
					oParam = oComm.Parameters.Add("@p_CastActivityGUID", SqlDbType.VarBinary, 16);
					oParam.Value = _genMgr.GuidStringToByteArray(CaseActivityGUID);

					oParam = oComm.Parameters.Add("@p_StageID", SqlDbType.NVarChar, StageID.Length);
					oParam.Value = StageID;

					oParam = oComm.Parameters.Add("@p_CaseStageSeqNo", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(CaseStageSeqNo, 10);

					oParam = oComm.Parameters.Add("@p_TaskID", SqlDbType.NVarChar, TaskID.Length);
					oParam.Value = TaskID;
				
					oParam = oComm.Parameters.Add("@p_TaskInstance", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(TaskInstance, 10);
					//MAR609 End

					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					StringBuilder sb = new StringBuilder("<RESPONSE>");

					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument ResponseDoc = new XmlDocument();
					ResponseDoc.LoadXml(sb.ToString());
				
					XmlNode xmlTemp = xnResponse.OwnerDocument.ImportNode(ResponseDoc.DocumentElement, true);
					xnResponse.AppendChild(xmlTemp);

					oReader.Close();
				}	
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
		}

		private void RunFTMiscUpdateInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTMiscUpdateInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data and load into a DomDocument
				GetFTFundsReleaseFailInterfaceData(xnRequest, xnResponse);
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "MESSAGETYPE");
				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(appFirstTitle,  "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");
				
				XmlNode taskNote = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/TASKNOTE");
				string noteEntry = _xmlMgr.xmlGetAttributeText(taskNote, "NOTEENTRY");

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);

				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "MIS";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = noteEntry;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTMISCUPDATEINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					//Save message and update CaseTask as completed

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MSI";
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGESUBTYPE");
					xnTempNode.InnerText = "MIS";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}

			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}

		private void RunFTCompletionDateUpdateInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTCompletionDateUpdateInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data 
				GetFTCompletionDateUpdateInterfaceData(xnRequest, xnResponse);
				
				XmlNode AppFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");	
				string MessageType =  _xmlMgr.xmlGetAttributeText(AppFirstTitle, "MESSAGETYPE");
				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(AppFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(AppFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(AppFirstTitle, "CONVEYANCERREFERENCE");
				
				XmlNode AppFactFind = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFACTFIND");	
				string EarliestCompletionDate = _xmlMgr.xmlGetAttributeText(AppFactFind, "EARLIESTCOMPLETIONDATE");

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);
		
				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "ECD";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = EarliestCompletionDate;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTCOMPLETIONDATEUPDATEINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					// If success then save the message and update the casetask with status "COMPLETED"

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MSI";
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGESUBTYPE");
					xnTempNode.InnerText = "ECD";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}

		private void GetFTCompletionDateUpdateInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetFTCompletionDateUpdateInterfaceData";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				//Select Mandatorynodes
				XmlNode xnAppNode = xnRequest.SelectSingleNode("//APPLICATION");
				
				//Check the mandatory Attributes
				string strApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONNUMBER");
				
				string strApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONFACTFINDNUMBER");
				
				string strMessageType = "ACA";
				
				if (strApplicationNumber.Length == 0 || strMessageType.Length == 0 || strApplicationFactFindNumber.Length == 0)
				{
					throw new ArgumentNullException();
				}

				string ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTCOMPLETIONDATEUPDATEINTERFACEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, strApplicationNumber.Length);
					oParam.Value = strApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(strApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, strMessageType.Length);
					oParam.Value = strMessageType;
			
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					StringBuilder sb = new StringBuilder("<RESPONSE>");

					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument ResponseDoc = new XmlDocument();
					ResponseDoc.LoadXml(sb.ToString());
				
					XmlNode xmlTemp = xnResponse.OwnerDocument.ImportNode(ResponseDoc.FirstChild, true);
					xnResponse.AppendChild(xmlTemp);

					oReader.Close();
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex.Message + ":" + cFunctionName);
			}
		}

		private void RunFTReInstateAppInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTReInstateAppInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument xdResponse = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data and load into a DomDocumnet
				GetApplicationFirstTitle(xnRequest, xnResponse, "ACA");
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");
				string MessageType =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "MESSAGETYPE");
				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);

				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "REI";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = xdReqDoc.CreateElement("XML");
				dataElement.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTREINSTATEAPPINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

				if (success)
				{
					// If success then save the message and update the casetask with status "COMPLETED"

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MSI";
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGESUBTYPE");
					xnTempNode.InnerText = "REI";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(cFunctionName + ": " + ex.Message);
			}
			finally
			{
				xnNode = null;
				xdResponse = null;
			}
		}

		private void RunFTOfferExpiryInterface(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "RunFTOfferExpiryInterface";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlNode xnNode;
			XmlDocument requestDocument = new XmlDocument();

			try
			{
				//Variable Declaration
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");
				
				//Get the Initial Request Data and load into a DomDocumnet
				GetRunFTOfferExpiryInterfaceData(xnRequest, xnResponse);
				
				XmlNode appFirstTitle = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONFIRSTTITLE");

				string ApplicationNumber =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "APPLICATIONNUMBER");
				string HubRef = _xmlMgr.xmlGetAttributeText(appFirstTitle, "HUBREFERENCE");
				string ConveyancerRef =  _xmlMgr.xmlGetAttributeText(appFirstTitle, "CONVEYANCERREFERENCE");
					
				XmlNode appOffer = _xmlMgr.xmlGetMandatoryNode(xnResponse, "//RESPONSE/APPLICATIONOFFER");
				
				string OfferExpiryDate =  _xmlMgr.xmlGetAttributeText(appOffer, "OFFEREXPIRYDATE");

				XmlDocument xdReqDoc = new XmlDocument();
				
				//Create Root Element
				XmlNode xnRootNode = xdReqDoc.CreateElement("SendMiscInstructionToConveyancerRequestType");
				xdReqDoc.AppendChild(xnRootNode);
				
				//Header Node 
				XmlNode xnHeader = xdReqDoc.CreateElement("Header");
				xdReqDoc.DocumentElement.AppendChild(xnHeader);

				//TimeStamp
				xnNode = xdReqDoc.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// Message Type
				xnNode = xdReqDoc.CreateElement("MessageType");
				xnNode.InnerText = "MSI";
				xnHeader.AppendChild(xnNode);
				
				// Message Sub Type
				xnNode = xdReqDoc.CreateElement("MessageSubType");
				xnNode.InnerText = "MOE";
				xnHeader.AppendChild(xnNode);

				// HubRef
				xnNode = xdReqDoc.CreateElement("HubRef");
				xnNode.InnerText = HubRef;
				xnHeader.AppendChild(xnNode);
				
				// Mortgage application No
				xnNode = xdReqDoc.CreateElement("MortgageApplicationNo");
				xnNode.InnerText  = ApplicationNumber;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				// conveyancers Ref
				xnNode = xdReqDoc.CreateElement("ConveyancerRef");
				xnNode.InnerText  = ConveyancerRef;
				xdReqDoc.DocumentElement.AppendChild(xnNode);
				
				// Text
				xnNode = xdReqDoc.CreateElement("Text");
				xnNode.InnerText  = OfferExpiryDate;
				xdReqDoc.DocumentElement.AppendChild(xnNode);

				XmlNode xn = xdReqDoc.CreateElement("XML");
				xn.AppendChild(xdReqDoc.DocumentElement);

				// RF 09/02/2006 MAR1392 Pass process context not logger
				bool success = _genMgr.AddMessageToQueue(xn.OuterXml, true, false, "RUNFTOFFEREXPIRYINTERFACE", 
					ApplicationNumber, _UserID, _UnitID, _UserAuthorityLevel, _ChannelID, _processContext);
			
				// If success then update the casetask with status "COMPLETED"
				if (success)
				{
					// If success then save the message and update the casetask with status "COMPLETED"

					//MAR1048
					//Save the message to the ApplicationFirstTitle table
					XmlDocument xdSaveMessage = new XmlDocument();
					XmlNode xnSaveMessage = xdSaveMessage.CreateElement("APPLICATIONFIRSTTITLE");
					xdSaveMessage.AppendChild(xnSaveMessage);

					XmlNode xnTempNode = xdSaveMessage.CreateElement ("APPLICATIONNUMBER");
					xnTempNode.InnerText = ApplicationNumber;
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGETYPE");
					xnTempNode.InnerText = "MSI";
					xnSaveMessage.AppendChild (xnTempNode);

					xnTempNode = xdSaveMessage.CreateElement ("MESSAGESUBTYPE");
					xnTempNode.InnerText = "MOE";
					xnSaveMessage.AppendChild (xnTempNode);

					SaveAppFirstTitleOutbound(xnSaveMessage.OwnerDocument.DocumentElement);	
					//MAR1048 End

					//Update CaseTask as completed
					UpdateCaseTask(xnRequest, xnCaseTask, "40");
				}

			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}

		private void GetRunFTOfferExpiryInterfaceData(XmlNode xnRequest, XmlNode xnResponse)
		{
			const string cFunctionName = "GetRunFTOfferExpiryInterfaceData";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				//Select Mandatory nodes
				XmlNode xnAppNode = xnRequest.SelectSingleNode("//APPLICATION");
					
				//Check the mandatory Attributes
				string strApplicationNumber =
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONNUMBER");
				
				string strApplicationFactFindNumber = 
					_xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "APPLICATIONFACTFINDNUMBER");
				
				//string strMessageType = _xmlMgr.xmlGetMandatoryAttributeText(xnAppNode, "MESSAGETYPE");
				string strMessageType = "ACA";	//MAR609 GHun
				
				string ConnectionString = Global.DatabaseConnectionString;

				using (SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_GETRUNFTOFFEREXPIRYINTERFACEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, strApplicationNumber.Length);
					oParam.Value = strApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(strApplicationFactFindNumber, 10);

					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, strMessageType.Length);
					oParam.Value = strMessageType;
			
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
				
					StringBuilder sb = new StringBuilder("<RESPONSE>");

					do 
					{
						while (oReader.Read())
						{
							sb.Append(oReader.GetString(0));
						}
					} while (oReader.NextResult());
				
					sb.Append ("</RESPONSE>");
				
					XmlDocument ResponseDoc = new XmlDocument();
					ResponseDoc.LoadXml(sb.ToString());
				
					XmlNode xmlTemp = xnResponse.OwnerDocument.ImportNode(ResponseDoc.FirstChild, true);
					xnResponse.AppendChild(xmlTemp);

					oReader.Close();
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}
		
		private bool PrepareFTInitialRequestToMessageQueue(XmlNode xnReqNode)
		{
			const string cFunctionName = "PrepareFTInitialRequestToMessageQueue";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			XmlDocument xdResponse = new XmlDocument();
			bool success = false;

			string ApplicationNumber = xnReqNode.Attributes.GetNamedItem("APPLICATIONNUMBER").Value;

			try
			{
				XmlNode xnNode;
				XmlDocument requestDocument = new XmlDocument();

				//Header
				XmlNode xmlRootReq = xnReqNode.SelectSingleNode("//RequestInsuranceRequestType");
				XmlNode xnHeader = xmlRootReq.OwnerDocument.CreateElement("Header");
				xmlRootReq.AppendChild(xnHeader);
				
				//TimeStamp
				xnNode = xnHeader.OwnerDocument.CreateElement("Timestamp");
				xnNode.InnerText  = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");	//MAR1036 GHun
				xnHeader.AppendChild (xnNode);
				
				// HubRef
				xnNode = xnHeader.OwnerDocument.CreateElement("HubRef");
				xnNode.InnerText = "";
				xnHeader.AppendChild(xnNode);

				// Set up data element
				XmlElement dataElement = requestDocument.CreateElement("XML");
				dataElement.AppendChild(requestDocument.ImportNode(xmlRootReq, true));

				// RF 09/02/2006 MAR1392 Pass process context not logger
				success = _genMgr.AddMessageToQueue(dataElement.OuterXml, true, false, 
					"RUNFTINITIALREQUESTINTERFACE", ApplicationNumber, _UserID, _UnitID, 
					_UserAuthorityLevel, _ChannelID, _processContext);

			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
			
			return success;
		}
		
		private bool InterfaceAvailability()
		{
			bool reply = false;

			//MAR609 GHun Still needs to be completed
			GlobalParameter gp = new GlobalParameter("FTOutboundIfaceInUseForBACS");
			reply = (!gp.Boolean);

			if (reply == true)
			{
				//MAR527 Check the InterfaceAvailabilityHours table for availability.
				string interfaceValidationType = "FT";
				string conn = Global.DatabaseConnectionString;
			
				try
				{
					using (SqlConnection oConn = new SqlConnection(conn))
					{
						SqlCommand oComm = new SqlCommand("USP_GetInterfaceAvailability", oConn);
						oComm.CommandType = CommandType.StoredProcedure;

						SqlParameter oParam = oComm.Parameters.Add("@p_InterfaceValType", SqlDbType.NVarChar, interfaceValidationType.Length);
						oParam.Value = interfaceValidationType;

						oConn.Open();
				
						SqlDataReader oReader = oComm.ExecuteReader(CommandBehavior.SingleResult);
						if (oReader.HasRows)
						{
							oReader.Read();
							reply = oReader.GetBoolean(0);
						}
					}
				}
				catch 
				{	
					if (_logger.IsErrorEnabled)	
						_logger.Error ("Unable to retrieve interface availability"); 
				}
			}

			return reply;
		}

		int MQL.IMessageQueueComponentVC2.OnMessage(string config, string xmlData)
		{
			const string cFunctionName = "OnMessage";

			// PSC 24/03/2006 MAR1521 Moved to before first logging statement
			// PSC 10/01/2006 MAR961 - Start
			_processContext = "MQL";
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
			// PSC 10/01/2006 MAR961 - End

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			// Set result = success so that message is removed from queue
			MQL.MESSQ_RESP result = MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
			
			try
			{
				FirstTitleNoTxBO noTxBo = new FirstTitleNoTxBO();	//MAR855 GHun
				// PSC 10/01/2006 MAR961
				noTxBo.ProcessOutboundMessage(xmlData, _processContext);	//MAR609 GHun
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				return (int) result;		//MAR855 GHun always return success
			}

			return (int) result;
		}
		
		//MAR1539 GHun add processContext parameter
		internal void ProcessOutboundMessage(string xmlData, string processContext)
		{
			const string cFunctionName = "ProcessOutboundMessage";
	
			//MAR1539 GHun
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, processContext);

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try
			{
				string errCode = string.Empty;
				bool success = false;

				XmlDocument xdDoc = new XmlDocument();
				xdDoc.LoadXml(xmlData);
				XmlNode xnReq = xdDoc.SelectSingleNode("//REQUEST").Clone();
				string interfaceName = _xmlMgr.xmlGetAttributeText(xnReq, "INTERFACENAME");
				string applicationNumber = _xmlMgr.xmlGetAttributeText(xdDoc.DocumentElement, "APPLICATIONNUMBER");

				//MAR835 Get request details before InterfaceAvailability check
				_UserID = _xmlMgr.xmlGetAttributeText(xnReq, "USERID");
				_UnitID = _xmlMgr.xmlGetAttributeText(xnReq, "UNITID");
				_UserAuthorityLevel = _xmlMgr.xmlGetAttributeText(xnReq, "USERAUTHORITYLEVEL");
				_ChannelID = _xmlMgr.xmlGetAttributeText(xnReq, "CHANNELID");
				_Password = _xmlMgr.xmlGetAttributeText(xnReq, "PASSWORD");
				//MAR835 End

				if (InterfaceAvailability())
				{
					string response = string.Empty;
					interfaceName = interfaceName.ToUpper();

					//MAR609 GHun Call Individual Methods
					switch (interfaceName)
					{
						case "RUNFTINITIALREQUESTINTERFACE":
							response = RunRequestInsuranceInterfaceWebService(xdDoc, ref errCode);
							break;
						case "RUNFTOFFERREQUESTINTERFACE":
						case "RUNFTREISSUEOFFERINTERFACE":
							response  = RunMakeOfferInterfaceWebService(xdDoc, ref errCode);
							break;
						case "RUNFTCANCELAPPINTERFACE":
						case "RUNFTFUNDSRELEASEFAILINTERFACE":
						case "RUNFTMISCUPDATEINTERFACE":
						case "RUNFTCOMPLETIONDATEUPDATEINTERFACE":
						case "RUNFTOFFEREXPIRYINTERFACE":
						case "RUNFTREINSTATEAPPINTERFACE":
							response = RunMiscInstructionInterfaceWebService(xdDoc, ref errCode);
							break;
					}
					
					success = ProcessFirstTitleOutboundResponse(interfaceName, applicationNumber, errCode);
					if (!success)
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(cFunctionName + " processing message failed. Adding back onto the queue to retry");
						}

						// RF 09/02/2006 MAR1392 Pass process context not logger
						success = _genMgr.AddMessageToQueue(xmlData, true, true, interfaceName, applicationNumber, 
							_UserID, _UnitID, _UserAuthorityLevel, _ChannelID, _processContext);
					}
				}
				else
				{	
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(cFunctionName + ": The interface is unavailable, so will retry in an hour");
					}
					//MAR609 GHun Add the message back on the queue to try again later
					// RF 09/02/2006 MAR1392 Pass process context not logger
					success = _genMgr.AddMessageToQueue(xmlData, true, true, interfaceName, applicationNumber, _UserID, 
						_UnitID, _UserAuthorityLevel, _ChannelID, _processContext);
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new omigaException(ex);
			}
		}

		private string RunRequestInsuranceInterfaceWebService(XmlDocument xdDoc, ref string errCode)
		{
			string responseString = string.Empty;
			string functionName = "RunRequestInsuranceInterfaceWebService";
			
			XmlNode xnRequestNode = xdDoc.SelectSingleNode("//RequestInsuranceRequestType");
			
			RequestInsurance.RequestInsuranceRequestType request = new RequestInsurance.RequestInsuranceRequestType();
			//MAR530
			RequestInsurance.RequestInsuranceResponseType response = new RequestInsurance.RequestInsuranceResponseType();
			RequestInsurance.DirectGatewaySoapService service = new RequestInsurance.DirectGatewaySoapService();
			DirectGatewayBO dg = new DirectGatewayBO();

			try
			{
				XmlNodeReader nodereader = new XmlNodeReader(xdDoc.ImportNode(xnRequestNode, true));
				XmlSerializer serOut = new XmlSerializer(typeof(RequestInsurance.RequestInsuranceRequestType));
				request = (RequestInsurance.RequestInsuranceRequestType)(serOut.Deserialize(nodereader));
				
				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();

				request.ClientDevice = dg.GetClientDevice();
				request.ServiceName = dg.GetServiceName();

				if (UserID.Length > 0 && Password.Length > 0)
				{
					request.TellerID = UserID;
					request.TellerPwd = Password;
				}
				request.ProxyID = dg.GetProxyId();
				request.ProxyPwd = dg.GetProxyPwd();
				
				request.Operator = dg.GetOperator();	//MAR852 GHun
				
				request.ProductType = dg.GetProductType();
				request.CommunicationChannel = dg.GetCommunicationChannel((UserID.Length > 0));

				XmlNode custNode = xdDoc.SelectSingleNode("OTHERSYSTEMCUSTOMERNUMBER");
				string CIFNumber = (custNode != null) ? custNode.Value: string.Empty;
				if (CIFNumber.Length > 0)
					request.CustomerNumber = CIFNumber;
				else
					request.CustomerNumber = "NA";

				request.ProductType = dg.GetProductType();

				request.Header.SequenceNo = dg.GetNextMessageSequenceNumber("FirstTitleOutboundMSN");
				
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(RequestInsurance.RequestInsuranceRequestType), request);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				response = service.execute(request);
			}
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "SOAP exception occurred: " + ex.Message;
			}
			catch(Exception ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Exception occurred: " + ex.Message;
			}
			catch
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Unknown exception occurred";
			} 
			finally
			{
				request = null;
			}
			
			if (response.ErrorCode != "0")
			{
				errCode = "ERRRFI";
			}

			//MAR530
			responseString = dg.GetXmlFromGatewayObject(typeof(RequestInsurance.RequestInsuranceResponseType), response);

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(functionName + ": Service Response: " + responseString);
			}

			return responseString;
		}

		private string RunMakeOfferInterfaceWebService(XmlDocument xdDoc, ref string errCode)
		{
			const string cFunctionName = "RunMakeOfferInterfaceWebService";
			
			XmlNode xnRequestNode = _xmlMgr.xmlGetMandatoryNode(xdDoc, "//MakeOfferRequestType");
			XmlNode xnHeader = _xmlMgr.xmlGetMandatoryNode(xnRequestNode, "//MakeOfferRequestType/Header");

			string messageType = _xmlMgr.xmlGetNodeText(xnHeader, "MessageType");

			MakeOffer.MakeOfferRequestType request = new MakeOffer.MakeOfferRequestType();
			//MAR530
			MakeOffer.MakeOfferResponseType response = new MakeOffer.MakeOfferResponseType();
			MakeOffer.DirectGatewaySoapService service = new MakeOffer.DirectGatewaySoapService();
			
			DirectGatewayBO dg = new DirectGatewayBO();

			try
			{
				XmlNodeReader nodereader = new XmlNodeReader(xdDoc.ImportNode(xnRequestNode, true));
				XmlSerializer serOut = new XmlSerializer(typeof(MakeOffer.MakeOfferRequestType));
				request = (MakeOffer.MakeOfferRequestType)(serOut.Deserialize(nodereader));
							
				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();

				request.ClientDevice = dg.GetClientDevice();
				request.ServiceName = dg.GetServiceName();

				if (UserID.Length > 0 && Password.Length > 0)
				{
					request.TellerID = UserID;
					request.TellerPwd = Password;
				}
				request.ProxyID = dg.GetProxyId();
				request.ProxyPwd = dg.GetProxyPwd();
				
				request.Operator = dg.GetOperator();	//MAR852 GHun

				request.ProductType = dg.GetProductType();
				request.CommunicationChannel = dg.GetCommunicationChannel((UserID.Length > 0));

				XmlNode custNode = xdDoc.SelectSingleNode("OTHERSYSTEMCUSTOMERNUMBER");
				string CIFNumber = (custNode != null) ? custNode.Value: string.Empty;
				if (CIFNumber.Length > 0)
					request.CustomerNumber = CIFNumber;
				else
					request.CustomerNumber = "NA";

				request.ProductType = dg.GetProductType();

				request.Header.SequenceNo = dg.GetNextMessageSequenceNumber("FirstTitleOutboundMSN");
				
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(MakeOffer.MakeOfferRequestType), request);
					_logger.Debug(cFunctionName + ": Service Request: " + requestString);
				}

				response = service.execute(request);
			}
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "SOAP exception occurred: " + ex.Message;
			}
			catch(Exception ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Exception occurred: " + ex.Message;
			}
			catch
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Unknown exception occurred";
			} 
			finally
			{
				request = null;
			}

			//MAR530
			string responseString = dg.GetXmlFromGatewayObject(typeof(MakeOffer.MakeOfferResponseType), response);
				
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(cFunctionName + ": Service Response: " + responseString);
			}

			if (response.ErrorCode != "0")
			{
				errCode = "ERR" + messageType;
			}

			return responseString;
		}

		private string RunMiscInstructionInterfaceWebService(XmlDocument xdDoc, ref string errCode)
		{
			string functionName = "RunMiscInstructionInterfaceWebService";

			XmlNode xnRequestNode = _xmlMgr.xmlGetMandatoryNode(xdDoc, "//SendMiscInstructionToConveyancerRequestType"); //MAR609 GHun
			
			//MAR530
			MiscInstruction.SendMiscInstructionToConveyancerRequestType request = new MiscInstruction.SendMiscInstructionToConveyancerRequestType();
			MiscInstruction.SendMiscInstructionToConveyancerResponseType response = new MiscInstruction.SendMiscInstructionToConveyancerResponseType();
			MiscInstruction.DirectGatewaySoapService service = new MiscInstruction.DirectGatewaySoapService();
			DirectGatewayBO dg = new DirectGatewayBO();

			try
			{
				XmlNodeReader nodereader = new XmlNodeReader(xdDoc.ImportNode(xnRequestNode, true));
				//MAR530
				XmlSerializer serOut = new XmlSerializer(typeof(MiscInstruction.SendMiscInstructionToConveyancerRequestType));
				request = (MiscInstruction.SendMiscInstructionToConveyancerRequestType)(serOut.Deserialize(nodereader));
								
				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();

				request.ClientDevice = dg.GetClientDevice();
				request.ServiceName = dg.GetServiceName();

				if (UserID.Length > 0 && Password.Length > 0)
				{
					request.TellerID = UserID;
					request.TellerPwd = Password;
				}
				request.ProxyID = dg.GetProxyId();
				request.ProxyPwd = dg.GetProxyPwd();

				request.Operator = dg.GetOperator();	//MAR852 GHun

				request.ProductType = dg.GetProductType();
				request.CommunicationChannel = dg.GetCommunicationChannel((UserID.Length > 0));

				XmlNode custNode = xdDoc.SelectSingleNode("OTHERSYSTEMCUSTOMERNUMBER");
				string CIFNumber = (custNode != null) ? custNode.Value: string.Empty;
				if (CIFNumber.Length > 0)
					request.CustomerNumber = CIFNumber;
				else
					request.CustomerNumber = "NA";

				request.ProductType = dg.GetProductType();

				request.Header.SequenceNo = dg.GetNextMessageSequenceNumber("FirstTitleOutboundMSN");
				
				if(_logger.IsDebugEnabled) 
				{
					//MAR530
					string requestString  = dg.GetXmlFromGatewayObject(typeof(MiscInstruction.SendMiscInstructionToConveyancerRequestType), request);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				response = service.execute(request);
			}
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "SOAP exception occurred: " + ex.Message;
			}
			catch(Exception ex)
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Exception occurred: " + ex.Message;
			}
			catch
			{
				response.ErrorCode = "SYSERR";
				response.ErrorMessage = "Unknown exception occurred";
			} 
			finally
			{
				request = null;
			}

			if (response.ErrorCode != "0")
			{
				XmlNode xnHeader = _xmlMgr.xmlGetMandatoryNode(xnRequestNode, "Header");
				string messageSubType = _xmlMgr.xmlGetNodeText(xnHeader, "MessageSubType");
				errCode = "ERRMSI" + messageSubType;				
			}

			// MAR530
			string responseString = dg.GetXmlFromGatewayObject(typeof(MiscInstruction.SendMiscInstructionToConveyancerResponseType), response);
				
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(functionName + ": Service Response: " + responseString);
			}

			return responseString;
		}

		private bool ProcessFirstTitleOutboundResponse(string interfaceName, 
			string applicationNumber, string errorCode) 
		{
			
			const string cFunctionName = "ProcessFirstTitleOutboundResponse";
			bool success = true;

			//If there is an error then 
			if ((errorCode.Length > 0) && (errorCode.Substring(0, 3) == "ERR"))
			{
				XmlNode xnRootNode;
				XmlDocument xdResponse = new XmlDocument();

				//Step 1:
				string messageType = _cdMgr.GetComboValueIDForValidationType("FirstTitleMessageType", errorCode); //MAR609 GHun
			
				//MAR609 GHun ApplicationLock no longer required
				//// Create Application Lock
				//success = _genMgr.CreateApplicationLock(applicationNumber, UserID, UnitID, UserAuthorityLevel, ChannelID, _logger);
				//if (!success)
				//{
				//	return false;
				//}
				
				//Step 2:
				
				//Call omTMBO.HANDLEINTERFACERESPONSE to create adhoc tasks for the above error

				//Build Request for HandleInterfaceResponse Call
				XmlDocument xdRequestDoc = new XmlDocument();
					
				xnRootNode = xdRequestDoc.CreateElement("REQUEST");
				xdRequestDoc.AppendChild(xnRootNode);
			
				_xmlMgr.xmlSetAttributeValue(xnRootNode, "USERID", UserID);
				_xmlMgr.xmlSetAttributeValue(xnRootNode, "UNIID", UnitID);
				_xmlMgr.xmlSetAttributeValue(xnRootNode, "USERAUTHORITYLEVEL", UserAuthorityLevel);
				_xmlMgr.xmlSetAttributeValue(xnRootNode, "CHANNELID", _ChannelID);	//MAR1646 GHun
				_xmlMgr.xmlSetAttributeValue(xnRootNode, "OPERATION", "HANDLEINTERFACERESPONSE");
				
				XmlNode xnInterfaceDetailsNode = xdRequestDoc.CreateElement("APPLICATION");
				
				_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "APPLICATIONNUMBER", applicationNumber);
				_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "APPLICATIONFACTFINDNUMBER", "1");
				xnRootNode.AppendChild(xnInterfaceDetailsNode);

				xnInterfaceDetailsNode = xdRequestDoc.CreateElement("INTERFACE");
				
				string InterfaceTypeValueID = _cdMgr.GetComboValueIDForValueName("InterfaceType", _InterfaceType);

				_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "INTERFACETYPE", InterfaceTypeValueID);
				_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "MESSAGETYPE", messageType);	//MAR609 GHun
				_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "CREATETASKFLAG", "1");
				
				xnRootNode.AppendChild(xnInterfaceDetailsNode);

				OmTmBOClass omTMBO = new OmTmBOClass();
				// MAR1603 M Heys 10/04/2006 start
				try
				{
					xdResponse.LoadXml(omTMBO.OmTmRequest(xdRequestDoc.OuterXml));
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(omTMBO);
				}
				// MAR1603 M Heys 10/04/2006 end
			
				string responseType = _xmlMgr.xmlGetAttributeText(xdResponse.DocumentElement, "TYPE");
				if (responseType != "SUCCESS")
				{
					success = false;
					if(_logger.IsErrorEnabled) 
					{
						_logger.Error(cFunctionName + ": Error in HandleInterfaceTask method: " + 
							xdResponse.OuterXml);
					}
				}
				
				//Step3

				//// Unlock Application
				//bool unlockSuccess  = _genMgr.UnlockApplication(applicationNumber, _logger);
			}

			return success;
		}

		//MAR182 GHun
		private string GetXmlPath()
		{
			string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			return(appPath.ToUpper().Replace("\\DLLDOTNET", "\\XML"));
		}

		//MAR609 GHun
		// RF 09/03/2006 MAR1392 Ensure error is handled
		/// <summary>
		/// Updates CaseTask as completed
		/// <summary>
		private void UpdateCaseTask(XmlNode xnRequest, XmlNode xnCaseTask, string status)
		{
			const string cFunctionName = "UpdateCaseTask"; 

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try 
			{
				XmlNode xnCaseTaskRootNode;
				XmlDocument xdCaseTask = new XmlDocument();

				xnCaseTaskRootNode = _xmlMgr.xmlGetRequestNode(xnRequest);
				xnCaseTaskRootNode = xdCaseTask.ImportNode(xnCaseTaskRootNode, true);
				_xmlMgr.xmlSetAttributeValue(xnCaseTaskRootNode, "OPERATION", "UPDATECASETASK"); 
						
				xdCaseTask.AppendChild (xnCaseTaskRootNode);
							
				_xmlMgr.xmlSetAttributeValue(xnCaseTask, "TASKSTATUS", status); // Complete
							
				xnCaseTask = xdCaseTask.ImportNode(xnCaseTask, true);
				xnCaseTaskRootNode.AppendChild(xnCaseTask);
				
				XmlDocument xdResponse = new XmlDocument();
				
				MsgTmBOClass MsgTmBO = new MsgTmBOClass();

				string req = xdCaseTask.OuterXml;
				
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Calling MsgTmBO.TmRequest: request = " + req);
				}

				//xdResponse.LoadXml(MsgTmBO.TmRequest(xdCaseTask.OuterXml));

				string resp = MsgTmBO.TmRequest(req);

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Called MsgTmBO.TmRequest: response = " + resp);
				}

				//string responseType = _xmlMgr.xmlGetAttributeText(xdResponse.DocumentElement, "TYPE");
				//			if (responseType != "SUCCESS")
				//			{
				//				if(_logger.IsErrorEnabled) 
				//				{
				//					_logger.Error("Error in Updating casetask: " + xdResponse.OuterXml);
				//				}
				//			}
				Omiga4Response omigaResp = new Omiga4Response(resp);
			}
			catch(OmigaException ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw;
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw;
			}
		}

		// MAR1048 Add routine to save outbound messages to ApplicationFirstTitle table
		// MAR1392 Ensure data is saved within txn 
		private void SaveAppFirstTitleOutbound(XmlNode xnAppFirstTitle)
		{
			const string cFunctionName = "SaveAppFirstTitleOutbound";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + cFunctionName);
			}

			try
			{
				FirstTitleTxBO txBo = new FirstTitleTxBO();	
				txBo.SaveAppFirstTitleOutbound(xnAppFirstTitle, _processContext);
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				throw new OmigaException(ex);
			}
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			FirstTitleBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		02/11/2005	MAR182	Created
GHun	09/11/2005	MAR182	Use omCore to get db connection string, error messages
GHun	30/11/2005	MAR609 Fixed numerous errors throughout
GHun	09/12/2005	MAR535 Handle errors from SaveApplicationFirstTitle correctly
GHun	14/12/2005	MAR855 Changed onMessage
GHun	14/12/2005	MAR885 Change SaveAppFirstTitleResponse to use correct Misc MessageSubType
PSC		10/01/2006	MAR961 Use omLogging wrapper
GHun	10/01/2006	MAR972 Changed ProcessFirstTitleInboundResponse to return a bool
RF		09/03/2006  MAR1392 Writing MOF message to ApplicationFirstTitle table is timing out 
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
GHun	26/04/2006	MAR1646 Changed ProcessFirstTitleInboundResponse to pass ChannelId to HandleInterfaceResponse
HMA     05/05/2006  MAR1713 Add extra data to UnlockApplication request.
GHun	11/05/2006	MAR1715 Changed ProcessFirstTitleInboundResponse and SaveAppFirstTitleResponse
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 10/01/2006 MAR961

using XMLManager = Vertex.Fsd.Omiga.omFirstTitle.XML.XMLManager;
using ConfigDataManager = Vertex.Fsd.Omiga.omFirstTitle.ConfigData.ConfigDataManager;
using omigaException = Vertex.Fsd.Omiga.Core.OmigaException;
using Vertex.Fsd.Omiga.omFirstTitle.General;

using omApp;
using omTM;
using MQL = MESSAGEQUEUECOMPONENTVCLib;

namespace Vertex.Fsd.Omiga.omFirstTitle
{
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[ProgId("omFirstTitle.FirstTitleInboundBO")]
	[ComVisible(true)]
	[Guid("5E0C1B0C-AB8C-4929-85E7-925000374A7A")]
	public class FirstTitleInboundBO: MQL.IMessageQueueComponentVC2
	{
		string _UserID;
		string _UnitID;
		string _UserAuthorityLevel;
		string _ChannelID;
		string _MachineID;               // MAR1713
		//string _Password;
		
		const string _InterfaceType = "First Title";
		const string _ClassName = "omFirstTitle.FirstTitleInboundBO";

		XMLManager _xmlMgr = new XMLManager();
		ConfigDataManager _cdMgr = new ConfigDataManager();
		GeneralManager _genMgr = new GeneralManager();

		// PSC 10/01/2006 MAR961 - Start
		private  omLogging.Logger _logger = null;
		private string _processContext = "Default";
		
		public FirstTitleInboundBO()
		{
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}
	
		public FirstTitleInboundBO(string processContext)
		{
			if (processContext == null || processContext.Length == 0)
			{
				throw new ArgumentNullException("Process context has not been supplied");
			}
			_processContext = processContext;
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}
		// PSC 10/01/2006 MAR961 - End

		int MQL.IMessageQueueComponentVC2.OnMessage(string config, string xmlData)
		{
			const string cFunctionName = "OnMessage";

			MQL.MESSQ_RESP result = MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;

			// PSC 10/01/2006 MAR961 - Start
			_processContext = "MQL";
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
			// PSC 10/01/2006 MAR961 - End

			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting " + cFunctionName + 
					": Inbound message received: " + xmlData);
			}

			try
			{
				FirstTitleNoTxBO noTxBo = new FirstTitleNoTxBO();	//MAR855 GHun
				// PSC 10/01/2006 MAR961
				noTxBo.ProcessInboundMessage(xmlData, _processContext);	//MAR609 GHun
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				return (int) result;	//MAR855 GHun always return success
			}

			// Set result = success so that message is removed from queue
			return (int) result;	//MAR609 GHun
		}

		internal bool ProcessFirstTitleInboundResponse(string DGResponseXml)
		{
			const string cFunctionName = _ClassName + ".ProcessFirstTitleInboundResponse";

			XmlDocument xdResponse = new XmlDocument();
			XmlNode xnRootNode;
			bool success;
			string ApplicationNumber = string.Empty;
			string MessageType = string.Empty;
			string MessageSubType = string.Empty;
			string MessageText = string.Empty;
			string ConveyancerRef = string.Empty;
			string HubRef = string.Empty;
			string TimeStamp = string.Empty;
			string RejectionReason = string.Empty;
			string TitleNo1 = string.Empty;
			string TitleNo2  = string.Empty;
			string TitleNo3 = string.Empty;
			string ExpiryDate  = string.Empty;
			string CompletionDate = string.Empty;
			string MortgageAdvance = string.Empty;
			string TaskNote = string.Empty;
			int	appFirstTitleId;	//MAR1715 GHun

			XmlDocument xdDoc = new XmlDocument();
			xdDoc.LoadXml(DGResponseXml);

			XmlNode requestNode = _xmlMgr.xmlGetMandatoryNode(xdDoc.DocumentElement, "/REQUEST");

			_UserID				= _xmlMgr.xmlGetAttributeText(requestNode, "USERID");
			_UnitID				= _xmlMgr.xmlGetAttributeText(requestNode, "UNITID"); 
			_UserAuthorityLevel = _xmlMgr.xmlGetAttributeText(requestNode, "USERAUTHORITYLEVEL"); 
			_ChannelID			= _xmlMgr.xmlGetAttributeText(requestNode, "CHANNELID");
			_MachineID          = _xmlMgr.xmlGetAttributeText(requestNode, "MACHINEID");                    // MAR1713
			//_Password			= _xmlMgr.xmlGetAttributeText(requestNode, "PASSWORD");
			
			ApplicationNumber = _xmlMgr.xmlGetMandatoryNodeText(requestNode, "//MortgageApplicationNo");	//MAR609 GHun

			//Step 1:
			
			// Create Application Lock
			success = _genMgr.CreateApplicationLock(ApplicationNumber, _UserID, _UnitID, _UserAuthorityLevel, _ChannelID, _logger);
			if (!success)
			{
				//MAR609 GHun Add the message back onto the queue to be processed at a later time
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + ": the application could not be locked, so the message will be added back to the queue to retry later");
				}
				// RF 09/02/2006 MAR1392 Pass process context not logger
				success = _genMgr.AddMessageToQueue(DGResponseXml, false, true, "", 
					ApplicationNumber, _UserID, _UnitID, _UserAuthorityLevel, _ChannelID, 
					_processContext);
				return success; //MAR972 GHun
			}
			//Step 2:
						
			// save ApplicationFirstTitle Data
			XmlNode xnHeader = xdDoc.SelectSingleNode("//Header");
			XmlNode xnStatusMsg = xdDoc.SelectSingleNode("//StatusMessage");
			XmlNode xnMiscMsg = xdDoc.SelectSingleNode("//Miscellaneous");	//MAR609 GHun
			
			XmlDocument xdAppFirstTitle = new XmlDocument();
			XmlNode xnAppFirstTitle = xdAppFirstTitle.CreateElement("APPLICATIONFIRSTTITLE");
			xdAppFirstTitle.AppendChild(xnAppFirstTitle);

			XmlNode xnTempNode = xdAppFirstTitle.CreateElement ("APPLICATIONNUMBER");
			xnTempNode.InnerText = ApplicationNumber;
			xnAppFirstTitle.AppendChild (xnTempNode);

			xnTempNode = xdAppFirstTitle.CreateElement ("APPLICATIONFACTFINDNUMBER");
			xnTempNode.InnerText = "1";
			xnAppFirstTitle.AppendChild (xnTempNode);

			if (xnStatusMsg != null)
			{
				xnTempNode = xdAppFirstTitle.CreateElement ("MESSAGETYPE");
				MessageType = _xmlMgr.xmlGetMandatoryNodeText(xnStatusMsg, "MessageSubType");
				xnTempNode.InnerText = MessageType;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				ConveyancerRef = _xmlMgr.xmlGetNodeText(xnStatusMsg, "ConveyancerRef");
				xnTempNode = xdAppFirstTitle.CreateElement ("CONVEYANCERREF");
				xnTempNode.InnerText = ConveyancerRef;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (ConveyancerRef.Length > 0)
					TaskNote += " ConveyancerRef:" + ConveyancerRef.ToString();

				RejectionReason = _xmlMgr.xmlGetNodeText(xnStatusMsg, "RejectionReason");
				xnTempNode = xdAppFirstTitle.CreateElement ("REJECTIONREASON");
				xnTempNode.InnerText = RejectionReason;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (RejectionReason.Length > 0)
					TaskNote += " RejectionReason:" + RejectionReason.ToString();

				TitleNo1 = _xmlMgr.xmlGetNodeText(xnStatusMsg, "TitleNo[1]");
				xnTempNode = xdAppFirstTitle.CreateElement ("TITLENO1");
				xnTempNode.InnerText = TitleNo1;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (TitleNo1.Length > 0)
					TaskNote += " TitleNo1:" + TitleNo1.ToString();

				TitleNo2 = _xmlMgr.xmlGetNodeText(xnStatusMsg, "TitleNo[2]");
				xnTempNode = xdAppFirstTitle.CreateElement ("TITLENO2");
				xnTempNode.InnerText =  TitleNo2;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (TitleNo2.Length > 0)
					TaskNote += " TitleNo2:" + TitleNo2.ToString();

				TitleNo3 = _xmlMgr.xmlGetNodeText(xnStatusMsg, "TitleNo[3]");
				xnTempNode = xdAppFirstTitle.CreateElement ("TITLENO3");
				xnTempNode.InnerText = TitleNo3;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (TitleNo3.Length > 0)
					TaskNote += " TitleNo3:" + TitleNo3.ToString();

				ExpiryDate = _xmlMgr.xmlGetNodeText(xnStatusMsg, "ExpiryDate");
				xnTempNode = xdAppFirstTitle.CreateElement ("EXPIRYDATE");
				xnTempNode.InnerText = ExpiryDate;
				xnAppFirstTitle.AppendChild (xnTempNode);
				
				if (ExpiryDate.Length > 0)
					TaskNote += " ExpiryDate:" + ExpiryDate.ToString();
				
				CompletionDate = _xmlMgr.xmlGetNodeText(xnStatusMsg, "CompletionDate");
				xnTempNode = xdAppFirstTitle.CreateElement ("COMPLETIONDATE");
				xnTempNode.InnerText = CompletionDate;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (CompletionDate.Length > 0)
					TaskNote += " CompletionDate:" + CompletionDate.ToString();

				MortgageAdvance = _xmlMgr.xmlGetNodeText(xnStatusMsg, "MortgageAdvance");
				xnTempNode = xdAppFirstTitle.CreateElement ("MORTGAGEADVANCE");
				xnTempNode.InnerText = MortgageAdvance;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (MortgageAdvance.Length > 0)
					TaskNote += " MortgageAdvance:" + MortgageAdvance.ToString();
			}

			if(xnMiscMsg != null)
			{
				xnTempNode = xdAppFirstTitle.CreateElement ("MESSAGETYPE");
				MessageType = "MSF";
				xnTempNode.InnerText = MessageType;
				xnAppFirstTitle.AppendChild (xnTempNode);

				xnTempNode = xdAppFirstTitle.CreateElement ("MESSAGESUBTYPE");
				MessageSubType = _xmlMgr.xmlGetMandatoryNodeText(xnMiscMsg, "MessageSubtype");	//MAR609 GHun
				xnTempNode.InnerText = MessageSubType;	//MAR885 GHun
				xnAppFirstTitle.AppendChild (xnTempNode);

				MessageText = _xmlMgr.xmlGetMandatoryNodeText(xnMiscMsg, "Text");	//MAR609 GHun
				xnTempNode = xdAppFirstTitle.CreateElement ("MESSAGETEXT");
				xnTempNode.InnerText = MessageText;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (MessageText.Length > 0)
					TaskNote += " MessageText:" + MessageText.ToString();

				ConveyancerRef = _xmlMgr.xmlGetMandatoryNodeText(xnMiscMsg, "ConveyancerRef");	//MAR609 GHun
				xnTempNode = xdAppFirstTitle.CreateElement ("CONVEYANCERREF");
				xnTempNode.InnerText = ConveyancerRef;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (ConveyancerRef.Length > 0)
					TaskNote += " ConveyancerRef:" + ConveyancerRef.ToString();
			}
			
			if(xnHeader != null)
			{
				HubRef = _xmlMgr.xmlGetNodeText(xnHeader, "HubRef");	//MAR609 GHun
				xnTempNode = xdAppFirstTitle.CreateElement ("HUBREF");
				xnTempNode.InnerText = HubRef;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (HubRef.Length > 0)
					TaskNote += " HubRef:" + HubRef.ToString();

				TimeStamp = _xmlMgr.xmlGetNodeText(xnHeader, "Timestamp");	//MAR609 GHun
				xnTempNode = xdAppFirstTitle.CreateElement ("TIMESTAMP");
				xnTempNode.InnerText = TimeStamp;
				xnAppFirstTitle.AppendChild (xnTempNode);

				if (TimeStamp.Length > 0)
					TaskNote += " DateTime:" + TimeStamp.ToString();
			}
			
			XmlNode xnSaveAppFirstTitleResponse;
			xnSaveAppFirstTitleResponse = xdAppFirstTitle.CreateElement("RESPONSE");
			_xmlMgr.xmlSetAttributeValue(xnSaveAppFirstTitleResponse, "TYPE", "SUCCESS");

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(cFunctionName + ": Calling SaveAppFirstTitleResponse: " + xnAppFirstTitle.OuterXml);
			}

			//MAR1715 GHun Get the ApplicationFirstTitleId back after saving
			appFirstTitleId = SaveAppFirstTitleResponse(xnAppFirstTitle.OwnerDocument.DocumentElement);	//MAR535 GHun
			success = (appFirstTitleId > 0);
			//MAR1715 End

			if (!success)
			{
				//MAR609 GHun Add the message back onto the queue to be processed at a later time
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + ": Saving ApplicationFirstTitle failed, so the message will be added back to the queue to retry later");
				}

				// RF 09/02/2006 MAR1392 Pass process context not logger
				success = _genMgr.AddMessageToQueue(DGResponseXml, false, true, "", 
					ApplicationNumber, _UserID, _UnitID, _UserAuthorityLevel, _ChannelID, 
					_processContext);
				throw new omigaException(cFunctionName + ": Error in saving ApplicationFirstTitle Record: " + xnSaveAppFirstTitleResponse.OuterXml);
			}

			//Step 3
		
			//Build Request for omTMBO.HandleInterfaceResponse Call
			XmlDocument xdRequestDoc = new XmlDocument();
					
			xnRootNode = xdRequestDoc.CreateElement("REQUEST");
			xdRequestDoc.AppendChild(xnRootNode);
			
			_xmlMgr.xmlSetAttributeValue(xnRootNode, "USERID", _UserID);
			_xmlMgr.xmlSetAttributeValue(xnRootNode, "UNITID", _UnitID);
			_xmlMgr.xmlSetAttributeValue(xnRootNode, "USERAUTHORITYLEVEL", _UserAuthorityLevel);

			_xmlMgr.xmlSetAttributeValue(xnRootNode, "CHANNELID", _ChannelID);	//MAR1646 GHun
			_xmlMgr.xmlSetAttributeValue(xnRootNode, "OPERATION", "HANDLEINTERFACERESPONSE");
				
			XmlNode xnInterfaceDetailsNode = xdRequestDoc.CreateElement("APPLICATION");
			
			//MAR609 GHun
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "APPLICATIONNUMBER", ApplicationNumber);
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "APPLICATIONFACTFINDNUMBER", "1");
			xnRootNode.AppendChild(xnInterfaceDetailsNode);

			xnInterfaceDetailsNode = xdRequestDoc.CreateElement("INTERFACE");

			string InterfaceTypeValueID = _cdMgr.GetComboValueIDForValueName("InterfaceType", _InterfaceType);
			
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "INTERFACETYPE", InterfaceTypeValueID);
			MessageType = _cdMgr.GetComboValueIDForValidationType("FirstTitleMessageType", MessageType);	//MAR609 GHun
			_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "MESSAGETYPE", MessageType);
			_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "CREATETASKFLAG", "1");
			_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "TASKNOTES", TaskNote);	//MAR609 GHun
			_xmlMgr.xmlSetAttributeValue(xnInterfaceDetailsNode, "CONTEXT", appFirstTitleId.ToString());	//MAR1715 GHun

			if (MessageSubType.Length > 0)	//MAR609 GHun
			{
				XmlNode xnRootMessageSubType = xdRequestDoc.CreateElement("MESSAGESUBTYPELIST");	//MAR609 GHun
				xnInterfaceDetailsNode.AppendChild(xnRootMessageSubType);

				XmlNode xnMessageSubType = xdRequestDoc.CreateElement("MESSAGESUBTYPE");
				MessageSubType = _cdMgr.GetComboValueIDForValidationType("FirstTitleMessageSubType", MessageSubType);	//MAR609 GHun
				_xmlMgr.xmlSetAttributeValue (xnMessageSubType, "MESSAGESUBTYPE", MessageSubType);
				xnRootMessageSubType.AppendChild(xnMessageSubType);
			}

			xnRootNode.AppendChild(xnInterfaceDetailsNode);

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(cFunctionName + ": Calling HandleInterfaceResponse: " + xdRequestDoc.OuterXml);
			}

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
				//MAR609 GHun HandleInterfaceResponse failed, so add the message back onto the queue to be processed at a later time
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + ": Error in HandleInterfaceResponse: " + xdResponse.OuterXml);
					_logger.Debug(cFunctionName + ": HandleInterfaceResponse failed, so the message will be added back to the queue to retry later");
				}

				// RF 09/02/2006 MAR1392 Pass process context not logger
				success = _genMgr.AddMessageToQueue(DGResponseXml, false, true, "", 
					ApplicationNumber, _UserID, _UnitID, _UserAuthorityLevel, _ChannelID, 
					_processContext);

				throw new omigaException(cFunctionName + ": Error in HandleInterfaceResponse method Call: " + xdResponse.OuterXml);
			}

			//Step 4:
			// Unlock Application
			
			//MAR1713  Add UserID and MachineID to UnlockApplication request.
			success = _genMgr.UnlockApplication(ApplicationNumber, _UserID, _MachineID, _logger);

			if(_logger.IsDebugEnabled && success)
			{
				_logger.Debug(cFunctionName + ": message successfully processed");
			}

			return success; //MAR972 GHun
		}

		//MAR1715 GHun changed type to int
		private int SaveAppFirstTitleResponse(XmlNode xnAppFirstTitle)
		{
			const string cFunctionName = "SaveAppFirstTitleResponse";
			int retVal = -1;	//MAR1715 GHun

			try 
			{
				//Check the mandatory Attributes
				string ApplicationNumber =
					_xmlMgr.xmlGetMandatoryNodeText(xnAppFirstTitle, "APPLICATIONNUMBER");
				
				string ApplicationFactFindNumber = "1";
					
				string MessageType = _xmlMgr.xmlGetMandatoryNodeText(xnAppFirstTitle, "MESSAGETYPE");

				string MessageSubType =	_xmlMgr.xmlGetNodeText(xnAppFirstTitle, "MESSAGESUBTYPE");

				string ConveyancerRef = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "CONVEYANCERREF");
	
				string HubRef = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "HUBREF");

				string TimeStamp = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "TIMESTAMP");

				string ReasonForRejection = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "REJECTIONREASON");

				string TitleNumber1 = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "TITLENO1");

				string TitleNumber2 = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "TITLENO2");

				string TitleNumber3 = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "TITLENO3");

				string RedemptionStatementExpiryDate = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "EXPIRYDATE");

				string CompletionDate = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "COMPLETIONDATE");

				string MessageText = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "MESSAGETEXT");

				string MortgageAdvance = _xmlMgr.xmlGetNodeText(xnAppFirstTitle, "MORTGAGEADVANCE");

				string ConnectionString = Global.DatabaseConnectionString;
				
				using(SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_SAVEAPPFIRSTTITLEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					//MAR1715 GHun Return value
					SqlParameter returnParam = oComm.Parameters.Add("RETURN_VALUE", SqlDbType.Int);
					returnParam.Direction = ParameterDirection.ReturnValue;
					//MAR1715 End

					//ApplicationNumber
					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					//ApplicationFactFindNumber
					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = ApplicationFactFindNumber;

					//MessageType
					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, MessageType.Length);
					if (MessageType.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MessageType;
				
					//MessageSubType
					oParam = oComm.Parameters.Add("@p_MessageSubType", SqlDbType.NVarChar, MessageSubType.Length);
					if (MessageSubType.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MessageSubType;
				
					//ConveyancerRef 
					oParam = oComm.Parameters.Add("@p_ConveyancerRef", SqlDbType.NVarChar, ConveyancerRef.Length);
					if (ConveyancerRef.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = ConveyancerRef;
				
					//HubRef
					oParam = oComm.Parameters.Add("@p_HubRef", SqlDbType.NVarChar, HubRef.Length);
					if (HubRef.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = HubRef;
				
					//DateTime
					oParam = oComm.Parameters.Add("@p_DateTime", SqlDbType.DateTime);
					if (TimeStamp.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = DateTime.Parse(TimeStamp);
				
					//ReasonForRejection 
					oParam = oComm.Parameters.Add("@p_ReasonForRejection", SqlDbType.NVarChar, ReasonForRejection.Length);
					if (ReasonForRejection.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = ReasonForRejection;
				
					//TitleNumber1
					oParam = oComm.Parameters.Add("@p_TitleNumber1", SqlDbType.NVarChar, TitleNumber1.Length);
					if (TitleNumber1.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = TitleNumber1;
				
					//TitleNumber2
					oParam = oComm.Parameters.Add("@p_TitleNumber2", SqlDbType.NVarChar, TitleNumber2.Length);
					if (TitleNumber2.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = TitleNumber2;
				
					//TitleNumber3
					oParam = oComm.Parameters.Add("@p_TitleNumber3", SqlDbType.NVarChar, TitleNumber3.Length);
					if (TitleNumber3.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = TitleNumber3;
				
					//RedemptionStatementExpDate
					oParam = oComm.Parameters.Add("@p_RedempStatementExpiryDate", SqlDbType.DateTime);
					if (RedemptionStatementExpiryDate.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = DateTime.Parse(RedemptionStatementExpiryDate);
				
					//CompletionDate
					oParam = oComm.Parameters.Add("@p_CompletionDate", SqlDbType.DateTime);
					if (CompletionDate.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = DateTime.Parse(CompletionDate);
				
					//MessageText
					oParam = oComm.Parameters.Add("@p_MessageText", SqlDbType.NVarChar, MessageText.Length);
					if (MessageText.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MessageText;
				
					//MortgageAdvance
					oParam = oComm.Parameters.Add("@p_MortgageAdvance", SqlDbType.NVarChar, MortgageAdvance.Length);
					if (MortgageAdvance.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MortgageAdvance;
				
					oConn.Open();
				
					int iRowsAffected = oComm.ExecuteNonQuery();
				
					if (iRowsAffected <= 0)
					{
						//MAR535 GHun log error and return failure
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(cFunctionName + " error: Error in inserting data");
						}
						//MAR535 End
					}
					retVal = (int) returnParam.Value;	//MAR1715 GHun
				}
			}
			catch(Exception ex)
			{
				//MAR535 GHun log error and return failure
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				//MAR535 End
			}

			return retVal;	//MAR1715 GHun
		}
	}
}

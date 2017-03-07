/*
--------------------------------------------------------------------------------------------
Workfile:			InboundBO.cs
Copyright:			Copyright © 2006 Vertex Financial Services

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
JD		28/10/2005	MAR342/MAR366 Continued work. Get config correct
JD		09/11/2005	MAR488 errors in inbound processing
GHun	21/11/2005	MAR538 errors in inbound processing
GHun	23/11/2005	MAR538 changed UpdateCaseTask
GHun	28/11/2005	MAR708
GHun	06/11/2005	MAR818 Fixed call to LTV calc and HandleInitialResponse messageType value
GHun	15/12/2005	MAR821
GHun	22/12/2005	MAR931 Fix messageType for HandleInitialResponse
JD	06/01/2006	MAR997 UpdateCaseTask set the status of the task before importing the node.
PSC	09/01/2006	MAR961 Use omLogging wrapper
JD	08/02/2006	MAR1231 make sure all nodes returned from request broker are sent to the valuationrules.
RF	11/02/2006	MAR1252 Change the Esurv response so that the buildings certification type is  converted
PSC	13/02/2006  	MAR1207 Enhance the error processing and logging
PSC	14/02/2007	MAR1207 Change call to AddMessageToQueue
PSC	16/02/2006	MAR1207	Amend UpdateCaseTask to set task status correctly
PSC	01/03/2006	MAR1343	Amend getValuationReportValuationNode to default saleability and overall condition
JD	16/03/2006	MAR1396 updateCaseTask - save instructionSeqNo as Context.
							ProcessValuationReport - update appointmentdate on valuerinstruction.
BC	24/03/2006	MAR1475	DateReceived (ValnRepSummary) should be populated from DateofInspection not DateofReport
				OverallCondition (ValnRepValuation) should be populated from ExternalAppearance
BC	24/03/2006	MAR1473 Set TYPEOFPROPERTY and PROPERTYDESCRIPTION to correct values
BC	24/03/2006	MAR1472	Amend from NONRESIDENTAILLANDIND to NONRESIDENTIALLANDIND (note the spelling) to update ValnRepPropertyRisks
BC	27/03/2006	MAR1459	NUMBEROFKITCHENS not being set in ValnRepPropertyDetails
BC	27/03/2006	MAR1465	Create attribute STRUCTURE from BuildingConstruction, as well as CONSTRUCTIONTYPEDETAILS
Ghun	04/04/2006	MAR1300 Changed ProcessValuationReport to clear change of property
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
MHeys	12/04/2006	MAR1603	Build errors rectified
MHeys	21/04/2006	MAR1655 Terminate COM+ objects cleanly
JD	23/04/2006	MAR1650 ProcessValuationReport - check input values when reading an external report.
JD	25/04/2006	MAR1670 Corrected error in MAR1650
JD	25/04/2006	MAR1640 If present valuation is zero set new ltv as 999
JD	28/04/2006	MAR1680 <transmission> node does not exist
BC	05/05/2006	MAR17141 Populate OverallCondition from OverallCondition (what a surprise) rather than ExternalAppearance.	
HMA 05/05/2006  MAR1713  Add Machine ID to Unlock Application request
DRC/JD 07/05/2006 MAR1718  Only Update Incomplete Tasks, otherwise create a new one
PSC	09/05/2006	MAR1742	Add additional logging
RF	11/05/2006  MAR1765 Late bind to omValuationRules
JD	11/05/2006	MAR1767 make sure correct 'Overall condition' is picked up based on report type
MH	15/05/2006	MAR1786	More COM+ DLL issues
GHun	18/05/2006	MAR1811 Backed out part of MAR1786 as it released objects too early
--------------------------------------------------------------------------------------------
*/

using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Runtime.InteropServices;
using System.Net;
using System.Xml.Serialization;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.ComponentModel;
using System.EnterpriseServices;

using omMQ;
using omTM;
using MsgTm;
using omApp;
using omAppProc;
using omAQ;
using omCM;
using omRB;
//using omValuationRules; RF 11/05/2006 MAR1765 

using omigaException = Vertex.Fsd.Omiga.Core.OmigaException;
using Vertex.Fsd.Omiga.omDg;

using MQL = MESSAGEQUEUECOMPONENTVCLib;
using WebRef = Vertex.Fsd.Omiga.omESurv.WebRefValuerValuation;
using XMLManager = Vertex.Fsd.Omiga.omESurv.XML.XMLManager;
using Vertex.Fsd.Omiga.Core;
using ConfigManager = Vertex.Fsd.Omiga.omESurv.ConfigData.ConfigDataManager;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 09/01/2006 MAR961

namespace Vertex.Fsd.Omiga.omESurv
{
	/// <summary>
	/// Summary description for InboundBO.
	/// </summary>
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[ProgId("omESurv.InboundBO")]
	[ComVisible(true)]
	[Guid("B644A0C8-6191-430c-8198-83B9A6D35FD9")]
	public class InboundBO : MQL.IMessageQueueComponentVC2 
	{
		XMLManager _xmlMgr = new XMLManager();
		ConfigManager _cdMgr = new ConfigManager();

		const int _noRecordsFound = -2147220492;
		
		// PSC 09/01/2006 MAR961 - Start
		private  omLogging.Logger _logger = null;
		private string _processContext = "Default";
		// PSC 09/01/2006 MAR961 - End
		
		private string _userId;
		private string _unitId;
		private string _userAuthorityLevel;
		private string _channelId;
		private string _machineId;    // MAR1713

		public InboundBO()
		{
			// PSC 09/01/2006 MAR961
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}

		// PSC 09/01/2006 MAR961 - Start
		public InboundBO(string processContext)
		{
			string functionName = System.Reflection.MethodBase.GetCurrentMethod().Name;
		
			// PSC 09/05/2006 MAR1742 - Start
			try
			{
				if (processContext == null || processContext.Length == 0)
				{
					throw new ArgumentNullException("Process context has not been supplied");
				}

				_processContext = processContext;
				_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
				
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Starting");
					_logger.Debug(functionName + ": processContext = " + processContext);
				}
			}
			finally
			{
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Completed");
				}
			}
			// PSC 09/05/2006 MAR1742 - End
		}
		// PSC 09/01/2006 MAR961 - End

		int MQL.IMessageQueueComponentVC2.OnMessage(string config, string xmlData)
		{
			int result = (int) MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
			
			// Reset logging thread context as we don't know the application number at this point
			omLogging.ThreadContext.Properties["ApplicationNumber"] = "Unknown";

			// PSC 09/01/2006 MAR961 - Start
			_processContext = "MQL";

			System.Reflection.MethodBase method = System.Reflection.MethodBase.GetCurrentMethod();
			string functionName = method.Name; 
			_logger = omLogging.LogManager.GetLogger(method.DeclaringType, _processContext);
			// PSC 09/01/2006 MAR961 - End

			try
			{
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Processing message. Config: " + config + " Data: " + xmlData);
				}


				//MAR821 GHun Call HandleInboundMessage through ESurvNoTxBO to escape the transaction
				ESurvNoTxBO noTxBO = new ESurvNoTxBO();
				// PSC 09/01/2006 MAR961
				result = noTxBO.HandleInboundMessage(xmlData, _processContext);	//HandleInboundMessage(xmlData);
				
				return (int)MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;

				//MAR821 End
			}
			catch (Exception exp)
			{
				if(_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": An unexpected error occurred handling inbound message", exp);
				}

				return result;	//MAR821 GHun return success regardless of any exceptions
			}
			// PSC 09/05/2006 MAR1742 - Start
			finally
			{
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + ": Processed message.");
				}
			}
			// PSC 09/05/2006 MAR1742 - End
		}

		public void HandleInitialResponse(string sApplicationNumber, string sApplicationFactFindNumber, string sUserID, string sUnitID, string sUserAuthLevel, string messageTypeValType, string sTaskNotes)
		{
			const string cFunctionName = "HandleInitialResponse";
			XmlDocument xmlDoc = new XmlDocument();
			XmlDocument xmlRespDoc = new XmlDocument();
			XmlNode xmlReqNode;
			xmlReqNode = xmlDoc.CreateElement("REQUEST");
			xmlDoc.AppendChild(xmlReqNode);

			//MAR931 Get the valueId for the message type validation type passed in
			string messageType = _cdMgr.GetComboValueIDForValidationType("ESurvMessageType", messageTypeValType);

			_xmlMgr.xmlSetAttributeValue (xmlReqNode, "USERID", sUserID);
			_xmlMgr.xmlSetAttributeValue (xmlReqNode, "UNITID", sUnitID);
			_xmlMgr.xmlSetAttributeValue (xmlReqNode, "USERAUTHORITYLEVEL", sUserAuthLevel);
			_xmlMgr.xmlSetAttributeValue (xmlReqNode, "OPERATION", "HANDLEINTERFACERESPONSE");
				
			XmlNode xnInterfaceDetailsNode = xmlDoc.CreateElement("APPLICATION");
				
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "APPLICATIONNUMBER", sApplicationNumber);
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
			xmlReqNode.AppendChild(xnInterfaceDetailsNode);

			xnInterfaceDetailsNode = xmlDoc.CreateElement("CASEACTIVITY");
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "CASEID", sApplicationNumber);
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "ACTIVITYID", "10");
			xmlReqNode.AppendChild(xnInterfaceDetailsNode);

			xnInterfaceDetailsNode = xmlDoc.CreateElement("INTERFACE");

			string InterfaceTypeValueID = _cdMgr.GetComboValueIDForValidationType("InterfaceType", "ES");	//MAR708 GHun

			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "INTERFACETYPE", InterfaceTypeValueID);
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "MESSAGETYPE", messageType);	//MAR818 GHun
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "TASKNOTES", sTaskNotes);
			_xmlMgr.xmlSetAttributeValue (xnInterfaceDetailsNode, "CREATETASKFLAG", "1");
			xmlReqNode.AppendChild(xnInterfaceDetailsNode);

			if(_logger.IsDebugEnabled)
			{
				_logger.Debug(cFunctionName + ": calling HandleInterfaceResponse: " + xmlDoc.OuterXml);
			}

			OmTmBO omTMBO = new OmTmBOClass();
			// MAR1603 M Heys 10/04/2006 start
			try 
			{
				xmlRespDoc.LoadXml(omTMBO.OmTmRequest(xmlDoc.OuterXml));
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(omTMBO);
			}
			// MAR1603 M Heys 10/04/2006 end
			string responseType = _xmlMgr.xmlGetAttributeText(xmlRespDoc.DocumentElement, "TYPE");
			
			// PSC 09/01/2006 MAR961 - Start
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(cFunctionName + ": response from HandleInitialResponse: " + xmlRespDoc.OuterXml);
			}
			// PSC 09/01/2006 MAR961 - End
		}

		public int HandleInboundMessage(string sInboundMessage)
		{
			const string cFunctionName = "HandleInboundMessage";
		
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(cFunctionName + ": processing message: " + sInboundMessage);
			}
			
			XmlDocumentEx xmlIn = new XmlDocumentEx();
			xmlIn.LoadXml(sInboundMessage);
			Boolean bValuationReportMessage = false;
			string strApplicationNumber = "";
			string strMessageType = "";
			
			bool postMessageBackToQueue = false;

			try
			{
				XmlElementEx requestElement = (XmlElementEx) xmlIn.SelectMandatorySingleNode("REQUEST");
				_userId = requestElement.GetMandatoryAttribute("USERID");
				_unitId = requestElement.GetMandatoryAttribute("UNITID");
				_userAuthorityLevel = requestElement.GetMandatoryAttribute("USERAUTHORITYLEVEL");
				_channelId = requestElement.GetAttribute("CHANNELID");
				_machineId = requestElement.GetAttribute("MACHINEID");   // MAR1713

			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(cFunctionName + ": Unable to get user details", exp);
				}
				
				return (int)MQL.MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
			}

			//determine what sort of message has been received
			XmlElementEx topElement;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(cFunctionName + ": Checking for StatusDetails");
			}

			topElement = (XmlElementEx) xmlIn.SelectSingleNode("//StatusDetails");
			if(topElement != null)
			{
				//we have some status details. Find the application number
				strApplicationNumber = topElement.GetNodeText("./QuestStatus/StatusData/ApplicationReference");
			}
			else
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(cFunctionName + ": StatusDetails not found checking for MortgageValuationDetails");
				}

				topElement = (XmlElementEx) xmlIn.SelectSingleNode("//MortgageValuationDetails");
				if(topElement != null)
				{
					//We have a valuation report. Find the applicationNumber
					bValuationReportMessage = true;
					strApplicationNumber = topElement.GetNodeText("./MortgageValuation/LenderReference");
					strMessageType = topElement.GetNodeText("./MortgageValuation/MessageType"); //MAR1680
				}
				else
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(cFunctionName + ": MortgageValuationDetails not found checking for ExternalAppraisalDetails");
					}

					topElement = (XmlElementEx) xmlIn.SelectSingleNode("//ExternalAppraisalDetails");
					if(topElement != null)
					{
						//We have a valuation report. Find the applicationNumber
						bValuationReportMessage = true;
						strApplicationNumber = topElement.GetNodeText("./ExternalAppraisal/LenderReference");
						strMessageType = topElement.GetNodeText("./ExternalAppraisal/MessageType");//MAR1680
					}
				}
			}
			if(strApplicationNumber == "")
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(cFunctionName + ": Unable to find application number. Message will be moved to the dead letter queue");
				}
				
				return (int)MQL.MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
			}
			else
			{
				// Set thread context so application number can be picked up in the log
				omLogging.ThreadContext.Properties["ApplicationNumber"] = strApplicationNumber;
				
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + ": Processing application");
					_logger.Debug(cFunctionName + ": About to lock application");
				}

				//ready to lock application and read poller response
				if(CreateApplicationLock(strApplicationNumber))
				{
					//make the xml upper case
					//string sXml = TopNode.OuterXml;
					//sXml = sXml.ToUpper();
					//XmlDocument newDoc = new XmlDocument();
					//newDoc.LoadXml(sXml);
					
					try
					{
						if(bValuationReportMessage == true)
						{
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": About to process valuation report");
							}

							ProcessValuationReport(strApplicationNumber, strMessageType, topElement.CloneNode(true));

							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": Processed valuation report");
							}

						}
						else
						{
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": About to processing Esurv code");
							}

							ProcessESurvCode(strApplicationNumber, topElement.CloneNode(true));
							
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": Processed Esurv code");
							}
						}
					}
					catch (Exception exp)
					{
						if(_logger.IsErrorEnabled) 
						{
							_logger.Error(cFunctionName + ": Failed to process application", exp);
							postMessageBackToQueue = true;
						}
					}

					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(cFunctionName + ": About to unlock application");
					}

					bool successfulUnlock = UnlockApplication(strApplicationNumber);

					if (_logger.IsDebugEnabled)
					{
						if (successfulUnlock)
						{
							_logger.Debug(cFunctionName + ": Unlocked application");
						}
					}
					
					if (_logger.IsWarnEnabled)
					{
						if (!successfulUnlock)
						{
							_logger.Debug(cFunctionName + ": Failed to unlock application");
						}
					}

					if (postMessageBackToQueue)
					{
						try
						{
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": About to put message back onto the queue");
							}

							//add message to queue with a delay to be picked up later
							ESurvNTxBO ntxBO = new ESurvNTxBO();	//MAR821 GHun
							// PSC 14/02/2007 MAR1207
							ntxBO.AddMessageToQueue(sInboundMessage, true, false, _userId, _unitId, _userAuthorityLevel, _logger);
							
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": Put message back onto the queue");
							}
						
							return (int)MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
						}
						catch (Exception exp)
						{
							if (_logger.IsErrorEnabled)
							{
								_logger.Debug(cFunctionName + ": Failed to put message back onto the queue", exp);
							}

							return (int)MQL.MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
						}
					}
					
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(cFunctionName + ": Processed application");
					}

					return (int)MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
				}
				else
				{
					try
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(cFunctionName + ": Failed to lock application. About to put message back onto the queue");
						}

						//add message to queue with a delay to be picked up later
						ESurvNTxBO ntxBO = new ESurvNTxBO();	//MAR821 GHun
						// PSC 14/02/2007 MAR1207
						ntxBO.AddMessageToQueue(sInboundMessage, true, false, _userId, _unitId, _userAuthorityLevel, _logger);	//MAR538 GHun Wrap message in XML tags
					
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(cFunctionName + ": Put message back onto the queue");
						}
						
						return (int)MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
					}
					catch (Exception exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Debug(cFunctionName + ": Failed to put message back onto the queue", exp);
						}

						return (int)MQL.MESSQ_RESP.MESSQ_RESP_RETRY_MOVE_MESSAGE;
					}
				}
			}
		}

		private string GetValuationTypeFromNewProperty(string AppNo, string AppFFNo)
		{
			string sValuationType = "";

			//call NewPropertyBO.GetValuationTypeAndLocation
			omApp.NewPropertyBOClass NewPropBO = new omApp.NewPropertyBOClass();
			XmlDocument xmlNPDoc = new XmlDocument();
			XmlNode xmlReq;
			xmlReq = xmlNPDoc.CreateElement("REQUEST");
			xmlNPDoc.AppendChild(xmlReq);

			XmlNode xmlTempElem;
			XmlNode xmlNewPropNode;
			xmlNewPropNode = xmlNPDoc.CreateElement("NEWPROPERTY");
			xmlReq.AppendChild(xmlNewPropNode);
			xmlTempElem = xmlNPDoc.CreateElement("APPLICATIONNUMBER");
			xmlTempElem.InnerXml = AppNo;
			xmlNewPropNode.AppendChild(xmlTempElem);
			xmlTempElem = xmlNPDoc.CreateElement("APPLICATIONFACTFINDNUMBER");
			xmlTempElem.InnerXml = AppFFNo;
			xmlNewPropNode.AppendChild(xmlTempElem);
			// MAR1603 M Heys 10/04/2006 start
			string strResponse = "";
			try
			{
				strResponse = NewPropBO.GetValuationTypeAndLocation(xmlReq.OuterXml);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(NewPropBO);
			}
			// MAR1603 M Heys 10/04/2006 end

			XmlDocument xmlRespDoc = new XmlDocument();
			XmlNode xmlRespNode;
			xmlRespDoc.LoadXml(strResponse);
			xmlRespNode = xmlRespDoc.SelectSingleNode(".//RESPONSE");

			if(xmlRespNode != null)
			{
				if(_xmlMgr.xmlGetAttributeText(xmlRespNode, "TYPE") == "SUCCESS")
				{
					sValuationType = _xmlMgr.xmlGetNodeText(xmlRespNode, ".//VALUATIONTYPE");
				}
				else
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug("Call to NewPropBO.GetValuationTypeAndLocation failed: " + xmlRespNode.OuterXml);
					}
				}
			}
			return(sValuationType);
		}

		private bool UpdateCaseTask(string sCaseId, string sGlobalParamTaskId, string sInstructionSeqNo)
		{
			//looks for the task on the application. If it is already present then
			// the task is updated to status Completed and Success = true is returned.
			//If the task is not already there, success = false is returned.
			Omiga4Response tempResponse = null;
			const string functionName = "UpdateCaseTask";
			string strResponse = string.Empty;
			string strTaskId = string.Empty;
			bool bSuccess = false;

			//First make sure we've got one. Call FindCaseTaskList
			try
			{
				XmlDocumentEx findCaseTaskDoc = new XmlDocumentEx();
				XmlElementEx findCaseTaskRequest;
				XmlElementEx caseTaskElement;

				findCaseTaskRequest = (XmlElementEx) findCaseTaskDoc.CreateElement("REQUEST");
				findCaseTaskDoc.AppendChild(findCaseTaskRequest);
				findCaseTaskRequest.SetAttribute("OPERATION", "FINDCASETASKLISTLITE");
				caseTaskElement = (XmlElementEx) findCaseTaskDoc.CreateElement("CASETASK");
				findCaseTaskRequest.AppendChild(caseTaskElement);
				caseTaskElement.SetAttribute("CASEID", sCaseId);
				strTaskId = new GlobalParameter(sGlobalParamTaskId).String;
				caseTaskElement.SetAttribute("TASKID", strTaskId);
				//MAR1718 - check for Incomplete tasks only
				caseTaskElement.SetAttribute("TASKSTATUS", "I");

				omTM.OmTmBO omTMBO = new OmTmBOClass();

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling FindCaseTaskListLite. Request: " + findCaseTaskDoc.OuterXml);
				}
				// MAR1603 M Heys 10/04/2006 start
				try
				{
						strResponse = omTMBO.OmTmRequest(findCaseTaskDoc.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(omTMBO);
				}
				// MAR1603 M Heys 10/04/2006 end

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from FindCaseTaskListLite. Response: " + strResponse);
				}

				tempResponse = new Omiga4Response(strResponse);	
			}
			catch (OmigaException exp)
			{
				if (exp.Omiga4Error == null || exp.Omiga4Error.Number != _noRecordsFound)					
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to retrieve current task list", exp);
					}
					throw new OmigaException("Unable to retrieve current task list", exp);
				}
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : An error occurred building the request for retrieving current task list", exp);
				}
				throw new OmigaException("An error occurred building the request for retrieving current task list", exp);
			}
				
			XmlDocumentEx caseTaskRespDoc = new XmlDocumentEx();
			XmlElementEx responseElement;

			try
			{
				caseTaskRespDoc.LoadXml(strResponse);
				responseElement = (XmlElementEx) caseTaskRespDoc.SelectMandatorySingleNode("RESPONSE");
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Invalid response returned from FindCaseTaskListLite", exp);
				}
				throw new OmigaException("Invalid response returned from FindCaseTaskListLite", exp);
			}

			//Update this case task status
			string strPattern = ".//CASETASK[@TASKID='" + strTaskId + "']";
			XmlElementEx thisTaskNode = (XmlElementEx) responseElement.SelectSingleNode(strPattern);
			
			if(thisTaskNode != null)
			{
				try
				{
					XmlDocumentEx updateCaseTaskDoc = new XmlDocumentEx();
					XmlElementEx updateCaseTaskRequest = (XmlElementEx) updateCaseTaskDoc.CreateElement("REQUEST");
					updateCaseTaskDoc.AppendChild(updateCaseTaskRequest); 
					//MAR538 GHun
					updateCaseTaskRequest.SetAttribute("USERID", _userId);
					updateCaseTaskRequest.SetAttribute("UNITID", _unitId);
					updateCaseTaskRequest.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
					//MAR538 End
					updateCaseTaskRequest.SetAttribute("OPERATION", "UPDATECASETASK");
					// PSC 16/02/2006 MAR1207 
					thisTaskNode.SetAttribute("TASKSTATUS", "40"); //Completed   JD MAR997 set new status before importing the node.
					if(sInstructionSeqNo != "")
						thisTaskNode.SetAttribute("CONTEXT", sInstructionSeqNo); //JD MAR1396
					updateCaseTaskRequest.AppendChild(updateCaseTaskDoc.ImportNode(thisTaskNode, true));	//MAR538 GHun
					
					MsgTm.MsgTmBO msgTMObj = new MsgTmBOClass();
					
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling UpdateCaseTask. Request: " + updateCaseTaskDoc.OuterXml);
					}
					// MAR1603 M Heys 10/04/2006 start
					try 
					{
						strResponse = msgTMObj.TmRequest(updateCaseTaskDoc.OuterXml);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(msgTMObj);
					}
					// MAR1603 M Heys 10/04/2006 end
											

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from UpdateCaseTask. Response: " + strResponse);
					}

					tempResponse = new Omiga4Response(strResponse);	
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to update case task", exp);
					}
					throw new OmigaException("Unable to update case task", exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request to update the case task", exp);
					}
					throw new OmigaException("An error occurred building the request to update the case task", exp);
				}
					
				try
				{
					caseTaskRespDoc.LoadXml(strResponse);
					responseElement = (XmlElementEx) caseTaskRespDoc.SelectMandatorySingleNode("RESPONSE");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from UpdateCaseTask", exp);
					}
					throw new OmigaException("Invalid response returned from UpdateCaseTask", exp);
				}
				bSuccess = true;
			}
			return(bSuccess);
		}

		private void ProcessESurvCode(string strApplicationNumber, XmlNode TopNode)
		{
			const string cFunctionName = "ProcessESurvCode";
			//Find the status code that's returned
			string sStatusCode = _xmlMgr.xmlGetNodeText(TopNode, "./QuestStatus/StatusData/status");
			//get any message and add as tasknotes
			string sMessage = _xmlMgr.xmlGetNodeText(TopNode, "./QuestStatus/StatusData/message");
			if(sStatusCode != "")
			{
				//create the associated tasks
				HandleInitialResponse(strApplicationNumber, "1", _userId, _unitId, _userAuthorityLevel, sStatusCode, sMessage);
				
				if(sStatusCode == "B")
				{
					//We also have a dateOfInspection returned. Create the valuerInstruction
					//and set the date
					string sDateOfInspection = _xmlMgr.xmlGetNodeText(TopNode, "./QuestStatus/StatusData/UpdateDate");
					if(sDateOfInspection != "")
					{
						string sValuationType = GetValuationTypeFromNewProperty(strApplicationNumber, "1");

						if(sValuationType != "")
						{
							//call CreateValuerInstructions
							XmlDocument xmlCreateValDoc = new XmlDocument();
							XmlNode xmlCreateValReqNode;
							XmlNode xmlValInstNode;
							xmlCreateValReqNode = xmlCreateValDoc.CreateElement("REQUEST");
							_xmlMgr.xmlSetAttributeValue(xmlCreateValReqNode, "OPERATION", "CreateValuerInstructions");
							xmlValInstNode = xmlCreateValDoc.CreateElement("VALUERINSTRUCTION");
							xmlCreateValReqNode.AppendChild(xmlValInstNode);

							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "APPLICATIONNUMBER", strApplicationNumber);
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "APPLICATIONFACTFINDNUMBER", "1");
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "VALUATIONTYPE", sValuationType);
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "APPOINTMENTDATE", sDateOfInspection);
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "DATEOFINSTRUCTION", sDateOfInspection);
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "VALUERPANELNO", "Default");
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "VALUERTYPE", _cdMgr.GetComboValueIDForValidationType("ValuerType", "E"));
							_xmlMgr.xmlSetAttributeValue(xmlValInstNode, "VALUATIONSTATUS", _cdMgr.GetComboValueIDForValidationType("ValuationStatus", "A"));

							omAppProc.omAppProcBO omAppProcBO = new omAppProcBOClass();

							// MAR1603 M Heys 10/04/2006 start
							string strResponse = "";
							try 
							{
								strResponse = omAppProcBO.OmAppProcRequest(xmlCreateValReqNode.OuterXml);
							}
							finally
							{
								//Free object reference
								System.Runtime.InteropServices.Marshal.ReleaseComObject(omAppProcBO);
							}
							// MAR1603 M Heys 10/04/2006 end
							XmlDocument xmlValInstRespDoc = new XmlDocument();
							xmlValInstRespDoc.LoadXml(strResponse);
							XmlNode xmlRespNode = xmlValInstRespDoc.SelectSingleNode(".//RESPONSE");

							if(xmlRespNode != null)
							{
								if(_xmlMgr.xmlGetAttributeText(xmlRespNode, "TYPE") == "SUCCESS")
								{
									//set task TMEsurvAwaitingBookTaskID as complete
									UpdateCaseTask(strApplicationNumber, "TMEsurvAwaitingBookTaskID", string.Empty);
								}
								else
								{
									if(_logger.IsDebugEnabled) 
									{
										_logger.Debug(cFunctionName + ": call to CreateValuerInstruction failed: " + xmlRespNode.OuterXml);
									}
									throw new OmigaException("CreateValuerInstruction failed");
								}
							}
						}
						else
						{
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug(cFunctionName + ": call to NewPropBO.GetValuationTypeAndLocation failed");
							}
							throw new OmigaException("NewPropertyBO.GetValuationTypeAndLocation failed");
						}
					}
					else
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug(cFunctionName + ": UpdateDate not found in xml: " + TopNode.OuterXml);
						}
						throw new OmigaException("UpdateDate not found in xml");
					}
				}
			}
			else
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + ": StatusCode not found in xml: " + TopNode.OuterXml);
				}
				throw new OmigaException("StatusCode not found in xml");
			}
		}
		
		// RF 11/05/2006 MAR1765 Late bind to omValuationRules
		private string CallValuationRules(string request)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name);
				_logger.Debug("Valuation Rules Request : " + request);
			}

			object vrInstance = null; 
			try 
			{
				// Get the type for omValuationRules on the local machine
				string progID = "omValuationRules.ValuationRulesBO";
				string methodName = "RunValuationRules";
				System.Type vrType; 
				try 
				{
					vrType = System.Type.GetTypeFromProgID(progID, true);
				}
				catch(Exception ex)
				{
					throw new OmigaException("Failed to get ProgID for " + progID, ex);
				}
				
				// Create an instance of omValuationRules on the local machine
				try
				{
					vrInstance = Activator.CreateInstance(vrType);
				}
				catch(Exception ex)
				{
					throw new OmigaException("Failed to create instance of " + progID, ex);
				}

				object[] objArgs = new object[] { request };

				string response = vrType.InvokeMember(
						methodName, 
						BindingFlags.Public | BindingFlags.InvokeMethod, 
						null, 
						vrInstance, 
						objArgs).ToString();	

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("Valuation Rules Response : " + response);
				}

				return response;
			}
			catch (Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name, ex);
				}
				throw new OmigaException("Failed to call valuation rules", ex);
			}
			finally
			{
				if (vrInstance != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(vrInstance);
				}
			}
		}

		private void ProcessValuationReport(string strApplicationNumber, string strMessageType, XmlNode TopNode)
		{
			const string functionName  = "ProcessValuationReport";
			//
			//Step 1 get any existing valuer instructions and create the val report
			//
			Omiga4Response tempResponse = null;
			string strResponse = string.Empty;
			omAppProc.omAppProcBO ValBO = null;
			XmlDocumentEx requestDocument;
			XmlElementEx requestElement;
			// MAR1603 M Heys 10/04/2006 start
			try
			{
			// MAR1603 M Heys 10/04/2006
				try
				{
					requestDocument = new XmlDocumentEx();
					requestElement = (XmlElementEx)requestDocument.CreateElement("REQUEST");
					requestDocument.AppendChild(requestElement);
					requestElement.SetAttribute("OPERATION", "FINDVALUERINSTRUCTIONLIST");
					XmlElementEx  valInstElement = (XmlElementEx)requestDocument.CreateElement("VALUERINSTRUCTION");
					requestElement.AppendChild(valInstElement);
					//make sure we get the last instruction
					valInstElement.SetAttribute("_ORDERBY_", "INSTRUCTIONSEQUENCENO DESC");
					valInstElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
					valInstElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling FindValuerInstructionList. Request: " + requestDocument.OuterXml);
					}
					ValBO = new omAppProcBOClass();
					
					strResponse = ValBO.OmAppProcRequest(requestDocument.OuterXml);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from FindValuerInstructionList. Response: " + strResponse);
					}

					tempResponse = new Omiga4Response(strResponse);	
				}
				catch (OmigaException exp)
				{
					if (exp.Omiga4Error == null || exp.Omiga4Error.Number != _noRecordsFound)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : Unable to retrieve valuer instruction list", exp);
						}
						throw new OmigaException("Unable to retrieve valuer instruction list", exp);
					}
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request for retrieving valuer instruction list", exp);
					}
					throw new OmigaException("An error occurred building the request for retrieving valuer instruction list", exp);
				}

				XmlDocumentEx responseDoc = new XmlDocumentEx();
				XmlElementEx valResponse;

				try
				{
					responseDoc.LoadXml(strResponse);
					valResponse = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from FindValuerInstructionList", exp);
					}
					throw new OmigaException("Invalid response returned from FindValuerInstructionList", exp);
				}

				XmlElementEx valInstruction = (XmlElementEx) valResponse.SelectSingleNode("VALUERINSTRUCTION");
				string strInstructionSequenceNo = "";
			
				if(valInstruction == null)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : No valuer instructions found. Calling CreateValuationReport.");
					}
					//we have no valuer instruction set up. Create a new report and a new instruction
					strInstructionSequenceNo = CreateValuationReport(strApplicationNumber);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Called CreateValuationReport.");
					}
				}
				else
				{				
					try
					{
						//we already have a val instruct so get the latest valuationreport.
						strInstructionSequenceNo =  valInstruction.GetAttribute("INSTRUCTIONSEQUENCENO");
						//This is the latest instructionSeqNo
						requestDocument = new XmlDocumentEx();
						requestElement  = (XmlElementEx)requestDocument.CreateElement("REQUEST");
						requestDocument.AppendChild(requestElement);
						requestElement.SetAttribute("OPERATION", "GETVALUATIONREPORT");
						XmlElementEx tempElement = (XmlElementEx)requestDocument.CreateElement("VALUATION");
						requestElement.AppendChild(tempElement);
						tempElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
						tempElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling GetValuationReport. Request: " + requestDocument.OuterXml);
						}

						strResponse = ValBO.OmAppProcRequest(requestDocument.OuterXml);
					
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Returned from GetValuationReport. Response: " + strResponse);
						}

						tempResponse = new Omiga4Response(strResponse);
					}
					catch (OmigaException exp)
					{
						if (exp.Omiga4Error == null || exp.Omiga4Error.Number != _noRecordsFound)
						{
							if (_logger.IsErrorEnabled)
							{
								_logger.Error(functionName + " : Unable to retrieve valuation report", exp);
							}
							throw new OmigaException("Unable to retrieve valuation report", exp);
						}
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : An error occurred building the request for retrieving the valuation report", exp);
						}
						throw new OmigaException("An error occurred building the request for retrieving the valuation report", exp);
					}
				
					XmlElementEx responseElement;

					try
					{
						responseDoc.LoadXml(strResponse);
						responseElement = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : Invalid response returned from GetValuationReport", exp);
						}
						throw new OmigaException("Invalid response returned from GetValuationReport", exp);
					}

					XmlElementEx valReportElement = (XmlElementEx) responseElement.SelectSingleNode("GETVALUATIONREPORT");

					if (valReportElement != null)
					{
						if (valReportElement.GetAttribute("INSTRUCTIONSEQUENCENO") == strInstructionSequenceNo)
						{
							//The latest val instruct has a report already so create a new report for a new instruction
							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Instruction sequence on report is the same as instruction. Calling CreateValuationReport to create a new report for a new instruction.");
							}

							strInstructionSequenceNo = CreateValuationReport(strApplicationNumber);
						
							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Called CreateValuationReport.");
							}
						}
						else
						{
							//The latest val instruct doesn't have a report so create a new report for this instruction
							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Latest valuer instruction does not have a report. Calling CreateValuationReport to create a new report for instruction " + strInstructionSequenceNo);
							}
							strInstructionSequenceNo = CreateValuationReport(strApplicationNumber, strInstructionSequenceNo);
						
							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Called CreateValuationReport.");
							}
						}
					}
					else
					{
						//We have no valuation report. Create a new report for this instruction
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Valuation report not found. Calling CreateValuationReport to create a new report for instruction " + strInstructionSequenceNo);
						}
						strInstructionSequenceNo = CreateValuationReport(strApplicationNumber, strInstructionSequenceNo);
						
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Called CreateValuationReport.");
						}				
					}	
				}
				if(strInstructionSequenceNo != "")
				{
					try
					{
						// STEP 2
						//we have got a new valuation report ready to fill in
						XmlDocumentEx valReportDocument = new XmlDocumentEx();
						XmlElementEx valReportRequest = (XmlElementEx)valReportDocument.CreateElement("REQUEST");
						valReportDocument.AppendChild(valReportRequest);
						valReportRequest.SetAttribute("OPERATION", "UpdateValuationReport");
						XmlElementEx valuationElement = (XmlElementEx)valReportDocument.CreateElement("VALUATION");
						valReportRequest.AppendChild(valuationElement);
						valuationElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
						valuationElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
						valuationElement.SetAttribute("INSTRUCTIONSEQUENCENO", strInstructionSequenceNo);
						getValuationReportSummaryNode(TopNode, valuationElement);
						getValuationReportValuationNode(strMessageType, TopNode, valuationElement);
						getValuationReportPropertyRisksNode(strMessageType,TopNode, valuationElement);
						getValuationReportPropertyServicesNode(TopNode, valuationElement);
						getValuationReportPropertyDetailsNode(strMessageType, TopNode, valuationElement);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling UpdateValuationReport. Request: " + valReportDocument.OuterXml);
						}
				
						strResponse = ValBO.OmAppProcRequest(valReportDocument.OuterXml);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Returned from UpdateValuationReport. Response: " + strResponse);
						}

						tempResponse = new Omiga4Response(strResponse);
					}
					catch (OmigaException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : Unable to update valuation report", exp);
						}
						throw new OmigaException("Unable to update valuation report", exp);
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : An error occurred building the request for updating the valuation report", exp);
						}
						throw new OmigaException("An error occurred building the request for updating the valuation report", exp);
					}

					//responseDoc.LoadXml(strResponse); // PSC delete
				
					//MAR1300 GHun Update Change of Property if applicable
					try 
					{
						UpdateChangeOfProperty(strApplicationNumber, "1");
					}
					catch (Exception ex)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : An error occurred updating NewProperty", ex);
						}
						throw new OmigaException("An error occurred updating NewProperty", ex);
					}
					//MAR1300 End

					try
					{
						//Valuation report successfully updated. Update ValuerInstruction.
						XmlDocumentEx valInstructionDoc = new XmlDocumentEx();
						XmlElementEx valInstructionRequest = (XmlElementEx)valInstructionDoc.CreateElement("REQUEST");
						valInstructionDoc.AppendChild(valInstructionRequest);
						valInstructionRequest.SetAttribute("OPERATION", "UpdateValuerInstructions");
						XmlElementEx valInstructionElement = (XmlElementEx)valInstructionDoc.CreateElement("VALUERINSTRUCTION");
						valInstructionRequest.AppendChild(valInstructionElement);
						valInstructionElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
						valInstructionElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
						valInstructionElement.SetAttribute("INSTRUCTIONSEQUENCENO", strInstructionSequenceNo);
						if(strMessageType == "External Report")
						{
							//set the valuationType as external appraisal
							valInstructionElement.SetAttribute("VALUATIONTYPE", _cdMgr.GetComboValueIDForValidationType("ValuationType", "DB"));
						}
						else
						{
							//set the valuationtype as the value in NewProperty
							valInstructionElement.SetAttribute("VALUATIONTYPE", GetValuationTypeFromNewProperty(strApplicationNumber, "1"));
						}
						XmlElementEx tempElement;
						tempElement = (XmlElementEx)TopNode.SelectSingleNode(".//Recommendation");
						valInstructionElement.SetAttribute("DATEOFINSPECTION", tempElement.GetNodeText("DateofInspection"));
						//MAR1396 update the appointmentdate also
						valInstructionElement.SetAttribute("APPOINTMENTDATE", tempElement.GetNodeText("DateofInspection"));
						valInstructionElement.SetAttribute("VALUERNAME", tempElement.GetNodeText("ValuerName"));
						valInstructionElement.SetAttribute("VALUATIONSTATUS", _cdMgr.GetComboValueIDForValidationType("ValuationStatus", "C"));
				
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling UpdateValuerInstructions. Request: " + valInstructionDoc.OuterXml);
						}

						strResponse = ValBO.OmAppProcRequest(valInstructionDoc.OuterXml);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Returned from UpdateValuerInstructions. Response: " + strResponse);
						}

						tempResponse = new Omiga4Response(strResponse);
					}
					catch (OmigaException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : Unable to update valuation instruction", exp);
						}
						throw new OmigaException("Unable to update instruction report", exp);
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : An error occurred building the request for updating the valuer instruction", exp);
						}
						throw new OmigaException("An error occurred building the request for updating the valuer instruction", exp);
					}

					//responseDoc.LoadXml(strResponse); // PSC delete
				
					//Valuer Instruction updated successfully.
					//Now calculate the new LTV based on the present valuation

					//MAR818 GHun
					//TempNode = TopNode.SelectSingleNode(".//Recommendation");
					//string strPresentValuation = _xmlMgr.xmlGetNodeText(TempNode, "PresentConditionValuation",false);
					//string strNewLTV = CalcCostModelLTV(strApplicationNumber, strPresentValuation);

					//MAR1640 If present valuation is 0 set LTV as special case 999
					string strNewLTV = "";

					XmlNode TempNode = TopNode.SelectSingleNode(".//Recommendation");
					string strPresentValuation = _xmlMgr.xmlGetNodeText(TempNode, "PresentConditionValuation");
					if(strPresentValuation == "0")
					{
						strNewLTV = "999";
					}
					else
					{
				
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling CalcCostModelLTV");
						}

						strNewLTV = CalcCostModelLTV(strApplicationNumber);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Called CalcCostModelLTV.");
						}			
					}
					//MAR818 End
				
					if(strNewLTV != "")
					{
						//successfully calculated new LTV. Now save it to the mortgagesubquote.
						//Get the current quote details and remember the OLD LTV for the rules processing
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling UpdateMortgageSubquote");
						}

						string sOldLTV = UpdateMortgageSubquote(strApplicationNumber, strNewLTV);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Called UpdateMortgageSubquote");
						}				

						if(sOldLTV != "")
						{
							//we have obtained the old LTV and updated the MSQ successfully
							//
							// STEP 3
							//update task and call valuation rules.
							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Calling UpdateCaseTask");
							}
							bool updatedCaseTask = UpdateCaseTask(strApplicationNumber, "TMEsurvValReportReturned", strInstructionSequenceNo); //JD MAR1396 

							if (_logger.IsDebugEnabled)
							{
								_logger.Debug(functionName + " : Called UpdateCaseTask");
							}				

							if(!updatedCaseTask)
							{
								//need to create a task as complete
								string sTaskId = new GlobalParameter("TMEsurvValReportReturned").String;
								string sTaskStatus = _cdMgr.GetComboValueIDForValidationType("TaskStatus", "CP");

								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : " + sTaskId + " task not updated. Calling CreateAdhocCaseTask for task " + sTaskId);
								}

								CreateAdhocCaseTask(strApplicationNumber, strInstructionSequenceNo, sTaskId, sTaskStatus);

								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : Called CreateAdhocCaseTask for task " + sTaskId);
								}
							}

							XmlDocumentEx requestBrokerDoc;

							try
							{
								//call valuation rules
								//get the requestbroker data to pass in
								requestBrokerDoc = new XmlDocumentEx();
								XmlElementEx rbRequestElement = (XmlElementEx) requestBrokerDoc.CreateElement("REQUEST");
								requestBrokerDoc.AppendChild(rbRequestElement);
								rbRequestElement.SetAttribute( "USERID", _userId);
								rbRequestElement.SetAttribute("UNITID", _unitId);
								rbRequestElement.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
								rbRequestElement.SetAttribute("COMBOLOOKUP", "NO");
								rbRequestElement.SetAttribute("RB_TEMPLATE", "APValnRBTemplate");
								XmlElementEx rbApplicationElement = (XmlElementEx) requestBrokerDoc.CreateElement("APPLICATION");
								rbRequestElement.AppendChild(rbApplicationElement);
								rbApplicationElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
								rbApplicationElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
								rbApplicationElement.SetAttribute("_SCHEMA_", "APPLICATION");

								omRB._OmRequestDO omRBDO = new OmRequestDOClass();
							
								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : Calling Request Broker data retrieval. Request: " + requestBrokerDoc.OuterXml);
								}

								// MAR1603 M Heys 10/04/2006 start
								try
								{
									strResponse = omRBDO.OmDataRequest(requestBrokerDoc.OuterXml);
								}
								finally
								{
									System.Runtime.InteropServices.Marshal.ReleaseComObject(omRBDO);
								}
								// MAR1603 M Heys 10/04/2006 end
								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : Returned from Request Broker data retrieval. Response: " + strResponse);
								}

								tempResponse = new Omiga4Response(strResponse);
							}
							catch (OmigaException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : Unable to retrieve Request Broker data", exp);
								}
								throw new OmigaException("Unable to retrieve Request Broker data", exp);
							}
							catch (XmlException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : An error occurred building the request to retrieve Request Broker data", exp);
								}
								throw new OmigaException("An error occurred building the request to retrieve Request Broker data", exp);
							}
						
							XmlElementEx requestBrokerResponse;

							try
							{
								requestBrokerDoc.LoadXml(strResponse);
								requestBrokerResponse = (XmlElementEx) requestBrokerDoc.SelectMandatorySingleNode(".//RESPONSE");
							}
							catch (XmlException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : Invalid response returned from Request Broker", exp);
								}
								throw new OmigaException("Invalid response returned from Request Broker", exp);
							}

							try
							{					
								//Call the rules
								XmlDocumentEx valRulesDoc = new XmlDocumentEx();
								XmlElementEx  valRulesRequest = (XmlElementEx) valRulesDoc.CreateElement("REQUEST");
								valRulesDoc.AppendChild(valRulesRequest);
								valRulesRequest.SetAttribute("USERID", _userId);
								valRulesRequest.SetAttribute("UNITID", _unitId);
								valRulesRequest.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
								valRulesRequest.SetAttribute("OPERATION", "ProcessValuationReport");
								//MAR1231 Add in all nodes returned from the call to RB
								XmlNodeList RBNodes = null;
								XmlNode RBnode = null;
								XmlNode RepNode = null;
								RBNodes = requestBrokerResponse.ChildNodes;

								foreach (XmlElementEx childElement in requestBrokerResponse.ChildNodes)
								{
									RBnode = childElement.CloneNode(true);
									RepNode = valRulesDoc.ImportNode(RBnode,true);
									valRulesRequest.AppendChild(RepNode);
								}

								XmlElementEx oldLTV = (XmlElementEx) valRulesDoc.CreateElement("OLDLTV");
								oldLTV.InnerText = sOldLTV;
								valRulesRequest.AppendChild(oldLTV);

								// RF 11/05/2006 MAR1765 Late bind to omValuationRules
								strResponse = CallValuationRules(valRulesDoc.OuterXml);

								tempResponse = new Omiga4Response(strResponse);
							}
							catch (OmigaException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : Unable to run valuation rules", exp);
								}
								throw new OmigaException("Unable to run valuation rules", exp);
							}
							catch (XmlException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : An error occurred building the request to valuation rules", exp);
								}
								throw new OmigaException("An error occurred building the request to valuation rules", exp);
							}

							XmlElement responseElement;

							try
							{
								responseDoc.LoadXml(strResponse);
								responseElement = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
							}
							catch (XmlException exp)
							{
								if (_logger.IsErrorEnabled)
								{
									_logger.Error(functionName + " : Invalid response returned from valuation rules", exp);
								}
								throw new OmigaException("Invalid response returned from valuation rules", exp);
							}

							XmlElementEx valuationElement = (XmlElementEx) responseElement.SelectSingleNode("VALUATION[@STATUS]");
							if(valuationElement != null)
							{
								//update the status on the valuation report
								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : Calling UpdateValuationReport");
								}
							
								UpdateValuationReport(strApplicationNumber, strInstructionSequenceNo, valuationElement.GetAttribute("STATUS"));
							
								if (_logger.IsDebugEnabled)
								{
									_logger.Debug(functionName + " : Called UpdateValuationReport");
								}

								//create any tasks returned from the rules component
								XmlNodeList taskList = responseElement.SelectNodes(".//TASKS/TASK");
								string sTaskStatus = _cdMgr.GetComboValueIDForValidationType("TaskStatus", "CP");
							
								foreach(XmlElementEx currentTask in taskList)
								{
									string taskId = currentTask.GetAttribute("TASKID");

									if (_logger.IsDebugEnabled)
									{
										_logger.Debug(functionName + " : Calling CreateAdhocCaseTask for task " + taskId);
									}

									CreateAdhocCaseTask(strApplicationNumber, 
										strInstructionSequenceNo, 
										taskId, 
										sTaskStatus, 
										currentTask.GetAttribute("TASKNOTES"));

									if (_logger.IsDebugEnabled)
									{
										_logger.Debug(functionName + " : Called CreateAdhocCaseTask for task " + taskId);
									}
								}	
							}
						}
					}
				}
			// MAR1603 M Heys 10/04/2006
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ValBO);
			}
			// MAR1603 M Heys 10/04/2006 end
		}

		private void UpdateValuationReport(string strApplicationNumber, string strInstructionSequenceNo, string sStatus)
		{
			const string functionName = "UpdateValuationReport";
			Omiga4Response tempResponse;
			try
			{

				XmlDocument tempDocument = new XmlDocumentEx();
				XmlElementEx  requestElement = (XmlElementEx) tempDocument.CreateElement("REQUEST");
				tempDocument.AppendChild(requestElement);
				requestElement.SetAttribute("OPERATION", "UpdateValuationReport");
				XmlElementEx valuationElement = (XmlElementEx) tempDocument.CreateElement("VALUATION");
				requestElement.AppendChild(valuationElement);
				//XmlNode tempNode = xmlValRepDoc.CreateElement("CREATEVALUATIONREPORTSUMMARY");
				//xmlValRepReqNode.AppendChild(tempNode);
				valuationElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
				valuationElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
				valuationElement.SetAttribute("INSTRUCTIONSEQUENCENO", strInstructionSequenceNo);
				if(sStatus == "SATISFIED")
				{
					valuationElement.SetAttribute("VALUATIONSTATUS", _cdMgr.GetComboValueIDForValidationType("ValuationStatus", "C"));
				}
				else
				{
					//NOTSATISFIED should also have the ValuationStatus set to Complete ("C")
					valuationElement.SetAttribute("VALUATIONSTATUS", _cdMgr.GetComboValueIDForValidationType("ValuationStatus", "C"));
				}

				omAppProc.omAppProcBO omAppProc = new omAppProcBOClass();
				
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling UpdateValuationReport. Request: " + tempDocument.OuterXml);
				}
				// MAR1603 M Heys 10/04/2006 start
				string strResponse = "";
				try
				{
					strResponse = omAppProc.OmAppProcRequest(tempDocument.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(omAppProc);
				}
				// MAR1603 M Heys 10/04/2006 end
				
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from UpdateValuationReport. Response: " + strResponse);
				}

				tempResponse = new Omiga4Response(strResponse);
			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Unable to update valuation report", exp);
				}
				throw new OmigaException("Unable to update valuation report", exp);
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : An error occurred building the request to update valuation report", exp);
				}
				throw new OmigaException("An error occurred building the request to update valuation report", exp);
			}		
		}

		private void CreateAdhocCaseTask(string strApplicationNumber, string strInstructionSequenceNo, string sTaskId, string sTaskStatus)
		{
			CreateAdhocCaseTask(strApplicationNumber, strInstructionSequenceNo, sTaskId, sTaskStatus, string.Empty);
		}

		private void CreateAdhocCaseTask(string strApplicationNumber, string strInstructionSequenceNo, string sTaskId, string sTaskStatus, string sTaskNote)
		{
			const string functionName = "CreateAdhocCaseTask";
			Omiga4Response tempResponse;
			string strResponse = string.Empty;
			string stageId = string.Empty;
			string caseStageSeqNo = string.Empty;
			XmlDocumentEx tempDocument;
			XmlElementEx requestElement;
			XmlElementEx caseActivityElement;
			MsgTm.MsgTmBO msgTmBO;
			// MAR1603 M Heys 10/04/2006 start
			msgTmBO = new MsgTmBOClass();
			try
			{
			// MAR1603 M Heys 10/04/2006
				//msgTmBO = new MsgTmBOClass();
				//first find the activity
				try
				{
					tempDocument = new XmlDocumentEx();
					requestElement = (XmlElementEx) tempDocument.CreateElement("REQUEST");
					requestElement.SetAttribute("OPERATION", "GetCaseActivity");
					tempDocument.AppendChild(requestElement);
					caseActivityElement = (XmlElementEx) tempDocument.CreateElement("CASEACTIVITY");
					requestElement.AppendChild(caseActivityElement);
					caseActivityElement.SetAttribute("CASEID", strApplicationNumber);

					//msgTmBO = new MsgTmBOClass();

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling GetCaseActivity. Request: " + tempDocument.OuterXml);
					}
					
					strResponse = msgTmBO.TmRequest(tempDocument.OuterXml);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from GetCaseActivity. Response: " + strResponse);
					}

					tempResponse = new Omiga4Response(strResponse);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to retrieve case activity", exp);
					}
					throw new OmigaException("Unable to retrieve case activity", exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request to retrieve case activity", exp);
					}
					throw new OmigaException("An error occurred building the request to retrieve case activity", exp);
				}

				XmlDocumentEx responseDoc = new XmlDocumentEx();
				XmlElementEx  responseElement;

				try
				{
					responseDoc.LoadXml(strResponse);
					responseElement = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from GetCaseActivity", exp);
					}
					throw new OmigaException("Invalid response returned from GetCaseActivity", exp);
				}
				
				try
				{
					caseActivityElement = (XmlElementEx) responseElement.SelectSingleNode("CASEACTIVITY");
					
					// now get the current stage
					tempDocument = new XmlDocumentEx();
					requestElement = (XmlElementEx) tempDocument.CreateElement("REQUEST");
					requestElement.SetAttribute("OPERATION", "GetCurrentStage");
					tempDocument.AppendChild(requestElement);
					XmlNode newActivityNode = tempDocument.ImportNode(caseActivityElement.CloneNode(true),true);
					requestElement.AppendChild(newActivityNode);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling GetCurrentStage. Request: " + tempDocument.OuterXml);
					}
					strResponse = msgTmBO.TmRequest(tempDocument.OuterXml);
				
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from GetCurrentStage. Response: " + strResponse);
					}
					tempResponse = new Omiga4Response(strResponse);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to retrieve current stage", exp);
					}
					throw new OmigaException("Unable to retrieve current stage", exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request to retrieve current stage", exp);
					}
					throw new OmigaException("An error occurred building the request to retrieve current stage", exp);
				}
				
				try
				{
					responseDoc.LoadXml(strResponse);
					responseElement = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from GetCurrentStage", exp);
					}
					throw new OmigaException("Invalid response returned from GetCurrentStage", exp);
				}

				try
				{
		
					XmlElementEx stageElement = (XmlElementEx) responseElement.SelectSingleNode("CASESTAGE");
					//now create the task
					tempDocument = new XmlDocumentEx();
					requestElement = (XmlElementEx) tempDocument.CreateElement("REQUEST");
					requestElement.SetAttribute("USERID", _userId);
					requestElement.SetAttribute("UNITID", _unitId);
					requestElement.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
					requestElement.SetAttribute("OPERATION", "CreateAdhocCaseTask");
					tempDocument.AppendChild(requestElement);
					XmlElementEx applicationElement = (XmlElementEx) tempDocument.CreateElement("APPLICATION");
					requestElement.AppendChild(applicationElement);
					applicationElement.SetAttribute("APPLICATIONPRIORITY", "10"); //DON'T KNOW WHAT THIS IS
					XmlElementEx caseTaskElement = (XmlElementEx) tempDocument.CreateElement("CASETASK");
					requestElement.AppendChild(caseTaskElement);
					caseTaskElement.SetAttribute("SOURCEAPPLICATION", "Omiga");
					caseTaskElement.SetAttribute("CASEID", strApplicationNumber);
					caseTaskElement.SetAttribute("ACTIVITYID", _xmlMgr.xmlGetAttributeText(caseActivityElement, "ACTIVITYID"));
					caseTaskElement.SetAttribute("ACTIVITYINSTANCE", _xmlMgr.xmlGetAttributeText(caseActivityElement, "ACTIVITYINSTANCE"));
					stageId = stageElement.GetAttribute("STAGEID");
					caseTaskElement.SetAttribute("STAGEID", stageId);
					caseTaskElement.SetAttribute("TASKID", sTaskId);
					caseTaskElement.SetAttribute("TASKSTATUS", sTaskStatus);
					caseTaskElement.SetAttribute("CONTEXT", strInstructionSequenceNo);
					caseStageSeqNo = stageElement.GetAttribute("CASESTAGESEQUENCENO");	//MAR538 GHun
					caseTaskElement.SetAttribute("CASESTAGESEQUENCENO", caseStageSeqNo);	//MAR538 GHun
					
					omTM.OmTmBO omTMBO = new OmTmBOClass();
					
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling CreateAdhocCaseTask. Request: " + tempDocument.OuterXml);
					}
					//MAR1786 M A Heys 15/05/2006 Interop Marshalling
					try
					{
						strResponse = omTMBO.OmTmRequest(tempDocument.OuterXml);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(omTMBO);
					}
					
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from CreateAdhocCaseTask. Response: " + strResponse);
					}
					
					tempResponse = new Omiga4Response(strResponse);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to create adhoc case task " + sTaskId, exp);
					}
					throw new OmigaException("Unable to create adhoc case task " + sTaskId, exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request to create adhoc case task " + sTaskId, exp);
					}
					throw new OmigaException("An error occurred building the request to create adhoc case task " + sTaskId, exp);
				}

				try
				{
					responseDoc.LoadXml(strResponse);
					responseElement = (XmlElementEx) responseDoc.SelectMandatorySingleNode("RESPONSE");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from CreateAdhocCaseTask", exp);
					}
					throw new OmigaException("Invalid response returned from CreateAdhocCaseTask", exp);
				}
				
				//if we have some tasknotes create those also
				if(sTaskNote != "")
				{
					//MAR538 GHun
					try
					{
						string caseActivityGuid = caseActivityElement.GetAttribute("CASEACTIVITYGUID");
						
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling GetMaxCaseTaskInstance");
						}
						
						int taskInstance = GetMaxCaseTaskInstance(stageId, sTaskId, caseActivityGuid, caseStageSeqNo);

						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Called GetMaxCaseTaskInstance");
						}
						//MAR538 End

						tempDocument  = new XmlDocumentEx();
						requestElement = (XmlElementEx) tempDocument.CreateElement("REQUEST");
						requestElement.SetAttribute("USERID", _userId);
						requestElement.SetAttribute("UNITID", _unitId);
						requestElement.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
						requestElement.SetAttribute("OPERATION", "CreateTaskNote");
						tempDocument.AppendChild(requestElement);
						XmlElementEx caseNoteElement = (XmlElementEx) tempDocument.CreateElement("TASKNOTE");
						requestElement.AppendChild(caseNoteElement);
						//_xmlMgr.xmlSetAttributeValue(xmlCaseNoteNode, "SOURCEAPPLICATION", "Omiga");
						caseNoteElement.SetAttribute("CASEACTIVITYGUID", caseActivityGuid);
						caseNoteElement.SetAttribute("CASESTAGESEQUENCENO", caseStageSeqNo);
						caseNoteElement.SetAttribute("TASKINSTANCE", taskInstance.ToString());
						caseNoteElement.SetAttribute("STAGEID", stageId);
						caseNoteElement.SetAttribute("TASKID", sTaskId);
						caseNoteElement.SetAttribute("NOTEENTRY", sTaskNote);
					
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Calling CreateTaskNote. Request: " + tempDocument.OuterXml);
						}
						strResponse = msgTmBO.TmRequest(tempDocument.OuterXml);
						
						if (_logger.IsDebugEnabled)
						{
							_logger.Debug(functionName + " : Returned from CreateTaskNote. Response: " + strResponse);
						}
				
						tempResponse = new Omiga4Response(strResponse);
					}
					catch (OmigaException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : Unable to create task note", exp);
						}
						throw new OmigaException("Unable to create task note", exp);
					}
					catch (XmlException exp)
					{
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + " : An error occurred building the request to create task note", exp);
						}
						throw new OmigaException("An error occurred building the request to create task note", exp);
					}
				}
			}
			// MAR1603 M Heys 10/04/2006
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(msgTmBO);
			}
			// MAR1603 M Heys 10/04/2006 end
		}

		private string UpdateMortgageSubquote(string strApplicationNumber, string strNewLTV)
		{
			const string functionName = "UpdateMortgageSubquote";
			string sOldLTV = "";
			string strResponse = string.Empty;
			XmlDocumentEx tempDocument;
			XmlElementEx requestElement;
			Omiga4Response tempResponse;

			try
			{
				//MAR818 GHun retrieve AmountRequested
				tempDocument = new XmlDocumentEx();
				XmlElementEx tempElement;
				requestElement = (XmlElementEx)tempDocument.CreateElement("REQUEST");
				tempDocument.AppendChild(requestElement);
				XmlElementEx applicationElement = (XmlElementEx)tempDocument.CreateElement("APPLICATION");
				requestElement.AppendChild(applicationElement);
				tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONNUMBER");
				tempElement.InnerText = strApplicationNumber;
				applicationElement.AppendChild(tempElement);
				tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONFACTFINDNUMBER");
				tempElement.InnerText = "1";
				applicationElement.AppendChild(tempElement);

				omAQ._ApplicationQuoteBO AppQuoteBO = new ApplicationQuoteBOClass();
			
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling GetAcceptedOrActiveQuoteData. Request: " + tempDocument.OuterXml);
				}

				// MAR1603 M Heys 10/04/2006 start
				try
				{
					strResponse = AppQuoteBO.GetAcceptedOrActiveQuoteData(tempDocument.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppQuoteBO);
				}
				// MAR1603 M Heys 10/04/2006 end

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from GetAcceptedOrActiveQuoteData. Response: " + strResponse);
				}

				tempResponse = new Omiga4Response(strResponse);
			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Unable to get accepted or active quote data", exp);
				}
				throw new OmigaException("Unable to get accepted or active quote data", exp);
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : An error occurred building the request for retrieving the accepted or active quote data", exp);
				}
				throw new OmigaException("An error occurred building the request for retrieving the accepted or active quote data", exp);
			}

			XmlDocumentEx responseDocument = new XmlDocumentEx();
			XmlElementEx responseElement;
			XmlElementEx mortgageSubQuote;
			
			try
			{
				responseDocument.LoadXml(strResponse);
				responseElement = (XmlElementEx) responseDocument.SelectMandatorySingleNode("RESPONSE"); 
				mortgageSubQuote = (XmlElementEx) responseElement.SelectMandatorySingleNode("MORTGAGESUBQUOTE");
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Invalid response returned from GetAcceptedOrActiveQuoteData", exp);
				}
				throw new OmigaException("Invalid response returned from GetAcceptedOrActiveQuoteData", exp);
			}

			try
			{

				sOldLTV = mortgageSubQuote.GetNodeText("LTV");
				XmlElementEx mortgageSubQuoteCopy = (XmlElementEx) mortgageSubQuote.CloneNode(true);
				XmlElementEx LTVElement = (XmlElementEx) mortgageSubQuoteCopy.SelectMandatorySingleNode("LTV");
				LTVElement.InnerText = strNewLTV;
			
				tempDocument = new XmlDocumentEx();
				requestElement = (XmlElementEx)tempDocument.CreateElement("REQUEST");
				tempDocument.AppendChild(requestElement);
				XmlElementEx tempElement =  (XmlElementEx) tempDocument.ImportNode(mortgageSubQuoteCopy, true);
				requestElement.AppendChild(tempElement);

				MortgageSubQuoteBOClass omCMBO = new MortgageSubQuoteBOClass();	//MAR538 GHun
				
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling MortgageSubQuoteBO.Update. Response: " + tempDocument.OuterXml);
				}
				// MAR1603 M Heys 10/04/2006 start
				try
				{
					strResponse = omCMBO.Update(tempDocument.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(omCMBO);
				}
				// MAR1603 M Heys 10/04/2006 end
				tempResponse = new Omiga4Response(strResponse);

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from MortgageSubQuoteBO.Update. Response: " + strResponse);
				}

			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Unable to update mortgage sub quote", exp);
				}
				throw new OmigaException("Unable to update mortgage sub quote", exp);
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : An error occurred building the request for updating mortgage sub quote", exp);
				}
				throw new OmigaException("An error occurred building the request for updating mortgage sub quote", exp);
			}

			return(sOldLTV);
		}

		private string CalcCostModelLTV(string strApplicationNumber) //MAR818 GHun removed parameter string strPresentValuation
		{
			const string functionName = "CalcCostModelLTV";
			string sNewLTV = "";
			string strResponse = string.Empty;
			string amountRequested;
			XmlElementEx tempElement;
			Omiga4Response tempResponse;
			omAQ._ApplicationQuoteBO AppQuoteBO = null;
			// MAR1603 M Heys 10/04/2006 start
			try
			{
			// MAR1603 M Heys 10/04/2006
				try
				{
					//MAR818 GHun retrieve AmountRequested
					XmlDocumentEx tempDocument = new XmlDocumentEx();
					XmlElementEx requestElement = (XmlElementEx)tempDocument.CreateElement("REQUEST");
					tempDocument.AppendChild(requestElement);
					XmlElementEx applicationElement = (XmlElementEx)tempDocument.CreateElement("APPLICATION");
					requestElement.AppendChild(applicationElement);
					tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONNUMBER");
					tempElement.InnerText = strApplicationNumber;
					applicationElement.AppendChild(tempElement);
					tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONFACTFINDNUMBER");
					tempElement.InnerText = "1";
					applicationElement.AppendChild(tempElement);

					AppQuoteBO = new ApplicationQuoteBOClass();
					
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling GetAcceptedOrActiveQuoteData. Request: " + tempDocument.OuterXml);
					}

					strResponse = AppQuoteBO.GetAcceptedOrActiveQuoteData(tempDocument.OuterXml);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from GetAcceptedOrActiveQuoteData. Response: " + strResponse);
					}

					tempResponse = new Omiga4Response(strResponse);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to get accepted or active quote data", exp);
					}
					throw new OmigaException("Unable to get accepted or active quote data", exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request for retrieving the accepted or active quote data", exp);
					}
					throw new OmigaException("An error occurred building the request for retrieving the accepted or active quote data", exp);
				}

				XmlDocumentEx responseDocument = new XmlDocumentEx();
				XmlElementEx responseElement;
					
				try
				{
					responseDocument.LoadXml(strResponse);
					responseElement = (XmlElementEx) responseDocument.SelectMandatorySingleNode("RESPONSE"); 
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from GetAcceptedOrActiveQuoteData", exp);
					}
					throw new OmigaException("Invalid response returned from GetAcceptedOrActiveQuoteData", exp);
				}

				try
				{
					amountRequested = responseElement.GetNodeText("MORTGAGESUBQUOTE/AMOUNTREQUESTED");
					//MAR818 End
					XmlDocumentEx tempDocument = new XmlDocumentEx();
					XmlElementEx LTVRequest;
					LTVRequest = (XmlElementEx)tempDocument.CreateElement("REQUEST");
					tempDocument.AppendChild(LTVRequest);
					XmlElementEx LTVElement = (XmlElementEx)tempDocument.CreateElement("LTV");
					LTVRequest.AppendChild(LTVElement);
					tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONNUMBER");
					tempElement.InnerText = strApplicationNumber;
					LTVElement.AppendChild(tempElement);
					tempElement = (XmlElementEx)tempDocument.CreateElement("APPLICATIONFACTFINDNUMBER");
					tempElement.InnerText = "1";
					LTVElement.AppendChild(tempElement);
					tempElement = (XmlElementEx)tempDocument.CreateElement("AMOUNTREQUESTED");
					tempElement.InnerText = amountRequested;	//MAR818 GHun
					LTVElement.AppendChild(tempElement);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Calling CalcCostModelLTV. Request: " + tempDocument.OuterXml);
					}
					strResponse = AppQuoteBO.CalcCostModelLTV(tempDocument.OuterXml);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug(functionName + " : Returned from CalcCostModelLTV. Response: " + strResponse);
					}

					tempResponse = new Omiga4Response(strResponse);
				}
				catch (OmigaException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Unable to calculate the LTV", exp);
					}
					throw new OmigaException("Unable to calculate the LTV", exp);
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : An error occurred building the request to calculate the LTV", exp);
					}
					throw new OmigaException("An error occurred building the request to calculate the LTV", exp);
				}

				try
				{
					responseDocument.LoadXml(strResponse);
					responseElement = (XmlElementEx) responseDocument.SelectSingleNode("RESPONSE");
					sNewLTV = responseElement.GetNodeText("LTV");
				}
				catch (XmlException exp)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error(functionName + " : Invalid response returned from CalcCostModelLTV", exp);
					}
					throw new OmigaException("Invalid response returned from CalcCostModelLTV", exp);
				}

				return (sNewLTV);
			// MAR1603 M Heys 10/04/2006
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(AppQuoteBO);
			}
			// MAR1603 M Heys 10/04/2006 end
		}

		private void getValuationReportSummaryNode(XmlNode TopNode, XmlNode ValRepSummaryNode)
		{
			XmlNode TempNode;
			TempNode = TopNode.SelectSingleNode(".//Recommendation");
			// BC MAR1475 24/03/2006
			_xmlMgr.xmlSetAttributeValue(ValRepSummaryNode, "DATERECEIVED", _xmlMgr.xmlGetNodeText(TempNode, "DateofInspection"));
		}

		private void getValuationReportValuationNode(string sReportType, XmlNode TopNode, XmlNode ValRepNode)
		{
			XmlNode TempNode;
			TempNode = TopNode.SelectSingleNode(".//PropertyDetails");
			
			// RF 11/02/2006 MAR1252 Start - Change the Esurv response  so that the buildings certification type is converted
			//_xmlMgr.xmlSetAttributeValue(ValRepNode, "CERTIFICATIONTYPE", _xmlMgr.xmlGetNodeText(TempNode, "BuildingIndemnityType"));
			//JD MAR1650 BuildingIndemnityType is not returned in an External Appraisal
			if(sReportType == "Valuation Report")
			{
				string validationType = _xmlMgr.xmlGetNodeText(TempNode, "BuildingIndemnityType");
				string valueID = _cdMgr.GetComboValueIDForValidationType("BuildingsCertificationType", validationType);
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "CERTIFICATIONTYPE", valueID);
			}
			// RF 11/02/2006 MAR1252 End

			TempNode = TopNode.SelectSingleNode(".//Recommendation");
			string strValue = _xmlMgr.xmlGetNodeText(TempNode, "Saleability");
			
			// PSC 01/03/2006 MAR1343 - Start
			switch(strValue)
			{
				case "1": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationSaleability", "AA"); 
					break;
				}
				case "2": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationSaleability", "A"); 
					break;
				}
				case "3": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationSaleability", "BA"); 
					break;
				}
				default: 
				{
					strValue = string.Empty;
					break;
				}
			}
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "SALEABILITY", strValue);
			// PSC 01/03/2006 MAR1343 - End

			//JD MAR1767 11/05/2006
			if(sReportType == "Valuation Report")
			{
				//BC MAR1714 05/05/2006
				strValue = _xmlMgr.xmlGetNodeText(TempNode, "OverallCondition");
			}
			else
			{
				strValue = _xmlMgr.xmlGetNodeText(TempNode, "ExternalAppearance");
			}
			
			// PSC 01/03/2006 MAR1343 - Start
			switch(strValue)
			{
				case "1": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationOverallCondition", "AA"); 
					break;
				}
				case "2": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationOverallCondition", "A"); 
					break;
				}
				case "3": 
				{
					strValue = _cdMgr.GetComboValueIDForValidationType("ValuationOverallCondition", "BA"); 
					break;
				}
				default: 
				{
					strValue = string.Empty;
					break;
				}
			}
			// PSC 01/03/2006 MAR1343 - End
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "OVERALLCONDITION", strValue);

			strValue = _xmlMgr.xmlGetNodeText(TempNode, "PresentConditionValuation");
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "PRESENTVALUATION", strValue);

			//JD MAR1650 BuildingInsuranceReInStatement is not returned in an External Appraisal
			if(sReportType == "Valuation Report")
			{
				strValue = _xmlMgr.xmlGetNodeText(TempNode, "BuildingInsuranceReInStatement");
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "REINSTATEMENTVALUE", strValue);
			}

			strValue = _xmlMgr.xmlGetNodeText(TempNode, "Signature");
			if(strValue != "")
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "SIGNATURERETURNED", "1");
			else
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "SIGNATURERETURNED", "0");
            
		}

		private void getValuationReportPropertyRisksNode(string sReportType, XmlNode TopNode, XmlNode ValRepNode)
		{
			XmlNode TempNode;
			TempNode = TopNode.SelectSingleNode(".//PotentialRisk");

			//JD MAR1650 only Essential Matters is returned in an External Appraisal
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "ESSENTIALMATTERS", _xmlMgr.xmlGetNodeText(TempNode, "EssentialMatters"));

			if(sReportType == "Valuation Report")
			{
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "PROPERTYSUBSIDENCE", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "StructuralMovement")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "LONGSTANDINGSUBIDENCE", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "StructuralMovementSignificant")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "TIMBERDAMPREPORTREQ", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "SeriousRot")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "ASBESTOSPOORCONDITION", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "AsbestosPoorCondition")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "CAVITYWALLTIEFAILURE", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "CavityWallTiefailure")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "LARGEPANELSYSTEMAPPRAISED", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "LargePanelSystem")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "HISTORICBUILDINGREPAIRSREQ", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "HistoricBuildingRepairRequired")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "RETYPEIND", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "ReTypeIndicator")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "RETYPEORIGINALLENDERNAME", _xmlMgr.xmlGetNodeText(TempNode, "ReTypeOriginalLender"));
			}

			TempNode = TopNode.SelectSingleNode(".//LegalIssues");
			if(TempNode != null)
			{
				//external driveby has no legal issues node
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "SOLICITORNOTES", _xmlMgr.xmlGetNodeText(TempNode, "Comments"));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "EXTENSIONSORALTERATIONS", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "Extensions")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "NONRESIDENTIALLANDIND", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "NonResidential")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "UNADOPTEDSHAREDACCESSISSUES", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "Roads")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "TENANTEDPROPERTYIND", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "Tenanted")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "DEVELOPMENTPROPOSALS", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "DevelopmentProposal")));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "MININGREPORTSREQ", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "MiningArea")));
			}
		}

		private void getValuationReportPropertyServicesNode( XmlNode TopNode, XmlNode ValRepNode)
		{

			XmlNode TempNode;
			TempNode = TopNode.SelectSingleNode(".//Accommodation");
			if(TempNode == null)
			{
				//external driveby node is called AssumedAccommodation. Try that one
				TempNode = TopNode.SelectSingleNode(".//AssumedAccommodation");
			}
			if(TempNode != null)
			{
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "LIVINGROOMS", _xmlMgr.xmlGetNodeText(TempNode, "LivingRooms"));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "NUMBEROFBEDROOMS", _xmlMgr.xmlGetNodeText(TempNode, "BedRooms"));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "BATHROOMS", _xmlMgr.xmlGetNodeText(TempNode, "BathRooms"));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "GARAGES", _xmlMgr.xmlGetNodeText(TempNode, "Garages"));
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "PARKINGSPACES", _xmlMgr.xmlGetNodeText(TempNode, "ParkingSpaces"));
			}       			
		}

		private void getValuationReportPropertyDetailsNode(string sReportType, XmlNode TopNode, XmlNode ValRepNode)
		{
			XmlNode TempNode;
			string strValue;
			TempNode = TopNode.SelectSingleNode(".//PropertyDetails/Address");
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "BUILDINGORHOUSENUMBER", _xmlMgr.xmlGetNodeText(TempNode, "HouseOrBuildingNumber"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "FLATNUMBER", _xmlMgr.xmlGetNodeText(TempNode, "FlatNameOrNumber"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "STREET", _xmlMgr.xmlGetNodeText(TempNode, "Street"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "BUILDINGORHOUSENAME", _xmlMgr.xmlGetNodeText(TempNode, "HouseOrBuildingName"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "DISTRICT", _xmlMgr.xmlGetNodeText(TempNode, "District"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "TOWN", _xmlMgr.xmlGetNodeText(TempNode, "TownOrCity"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "COUNTY", _xmlMgr.xmlGetNodeText(TempNode, "County"));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "POSTCODE", _xmlMgr.xmlGetNodeText(TempNode, "PostCode"));

			TempNode = TopNode.SelectSingleNode(".//PropertyDetails");
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "INSTRUCTIONADDRESSCORRECT", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "InstructionAddressCorrect")));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "NEWPROPERTYINDICATOR", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "NewPropertyIndicator")));
			//JD MAR1650 Tenure, Leasehold is not returned in an External Appraisal
			if(sReportType == "Valuation Report")
			{
				strValue = _xmlMgr.xmlGetNodeText(TempNode, "Tenure");
				switch(strValue)
				{
					case "Freehold" : strValue = _cdMgr.GetComboValueIDForValidationType("PropertyTenure", "F"); break;
					case "Commonhold" : strValue = _cdMgr.GetComboValueIDForValidationType("PropertyTenure", "C"); break;
					case "Leasehold" : strValue = _cdMgr.GetComboValueIDForValidationType("PropertyTenure", "L"); break;
					default : strValue = _cdMgr.GetComboValueIDForValidationType("PropertyTenure", "FE"); break;
				}
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "TENURE", strValue);
			
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "UNEXPIREDLEASE", _xmlMgr.xmlGetNodeText(TempNode, "Leasehold"));
			}

			strValue = _xmlMgr.xmlGetNodeText(TempNode, "PropertyType");
			string strPropType = "";
			string strPropDesc = "";
			switch(strValue)
			{
				//BC MAR1473 24/03/2006 Begin
				case "HT" :
					strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "H");
					    strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "HT");
					break;

				case "HS" :
						strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "H");
					    strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "HS");
						break;

				case "HD" :
						strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "H");
					    strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "HD");
						break;

				case "F" :
						strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "F"); 
					    strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "F");
						break;

				case "B" :
						strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "B");
					    strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "B");
						break;

				default :
						strPropType = _cdMgr.GetComboValueIDForValidationType("PropertyType", "O");
						strPropDesc = _cdMgr.GetComboValueIDForValidationType("PropertyDescription", "O");
						break;
				//BC MAR1473 24/03/2006 End
			}
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "TYPEOFPROPERTY", strPropType);
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "PROPERTYDESCRIPTION", strPropDesc);
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "EXLOCALAUTHORITY", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "ExLocalAuthority")));
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "YEARBUILT", _xmlMgr.xmlGetNodeText(TempNode, "ApproxYearPropertyBuilt"));
			//JD MAR1650 GrossExternalFloorArea is not returned in an External Appraisal
			if(sReportType == "Valuation Report")
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "RESIDENCEAREA", _xmlMgr.xmlGetNodeText(TempNode, "GrossExternalFloorArea"));

			TempNode = TopNode.SelectSingleNode(".//Construction");
			if(TempNode == null)
			{
				//external driveby node name is AssumedConstruction
				TempNode = TopNode.SelectSingleNode(".//AssumedConstruction");
			}
			if(TempNode != null)
			{
				strValue = _xmlMgr.xmlGetNodeText(TempNode, "TraditionalConstructionIndicator");
				if(strValue == "Y")
				{
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "STRUCTURE",_cdMgr.GetComboValueIDForValidationType("BuildingConstruction", "S"));
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "CONSTRUCTIONTYPEDETAILS",_cdMgr.GetComboValueIDForValidationType("BuildingConstruction", "S"));
				}
				else
				{
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "STRUCTURE",_cdMgr.GetComboValueIDForValidationType("BuildingConstruction", "NS"));
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "CONSTRUCTIONTYPEDETAILS",_cdMgr.GetComboValueIDForValidationType("BuildingConstruction", "NS"));
				}
				_xmlMgr.xmlSetAttributeValue(ValRepNode, "NONSTANDARDCONSTRUCTIONTYPE", _xmlMgr.xmlGetNodeText(TempNode, "TraditionalConstructionType"));
				//JD MAR1650 singleSkinConstruction is not returned in an External Appraisal
				if(sReportType == "Valuation Report") //MAR1670
				{
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "ISSINGLESKINCONSTRUCTION", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "SingleSkinConstructionIndicator")));
					_xmlMgr.xmlSetAttributeValue(ValRepNode, "ISSINGLESKINTWOSTOREY", GetBoolFromYN(_xmlMgr.xmlGetNodeText(TempNode, "SingleSkin2storyIndicator")));
				}
				//_xmlMgr.xmlSetAttributeValue(ValRepNode, "GENERALOBSERVATIONS", _xmlMgr.xmlGetNodeText(TempNode, "GeneralObservations"));
			}
			//MAR1396 GeneralObservations is directly under MortgageValuation not in Construction
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "GENERALOBSERVATIONS", TopNode.SelectSingleNode(".//GeneralObservations").InnerXml);
			
			//MAR1459 BC 27/03/2006 Begin

			TempNode = TopNode.SelectSingleNode(".//Accommodation");

			if(TempNode == null)
			{
				//external driveby node is called AssumedAccommodation. Try that one
				TempNode = TopNode.SelectSingleNode(".//AssumedAccommodation");
			}
			_xmlMgr.xmlSetAttributeValue(ValRepNode, "NUMBEROFKITCHENS", _xmlMgr.xmlGetNodeText(TempNode, "Kitchens"));
			
			//MAR1459 BC 27/03/2006 End
		}

		private string GetBoolFromYN(string YNString)
		{
			string strBool = "0";

			if(YNString == "Y")
				strBool = "1";
			return strBool;
		}

		private string CreateValuationReport(string strApplicationNumber)
		{
			return CreateValuationReport(strApplicationNumber, string.Empty);
		}

		private string CreateValuationReport(string strApplicationNumber, string strInstNo)
		{
			//If valuation report created successfully the inst no is returned otherwise
			//empty string is returned.
			const string functionName = "CreateValuationReport";
			string strRetInstNo = "";
			string strResponse = string.Empty;
			string strOperation = "";
			Omiga4Response tempResponse;

			XmlDocumentEx requestDocument;
			XmlElementEx requestElement;

			try
			{
				requestDocument = new XmlDocumentEx();
				requestElement = (XmlElementEx)requestDocument.CreateElement("REQUEST");
				requestDocument.AppendChild(requestElement);
            
				if(strInstNo == "")
					strOperation = "CreateValuationReportNoInst";
				else
					strOperation = "CreateValuationReport";

				requestElement.SetAttribute("OPERATION", strOperation);
				XmlElementEx valuationElement = (XmlElementEx)requestDocument.CreateElement("VALUATION");
				requestElement.AppendChild(valuationElement);
				valuationElement.SetAttribute("APPLICATIONNUMBER", strApplicationNumber);
				valuationElement.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
				if(strInstNo != "")
					valuationElement.SetAttribute("INSTRUCTIONSEQUENCENO", strInstNo);

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling " + strOperation + ". Request: " + requestDocument.OuterXml);
				}
				//MAR1786 M Heys 15/05/2006 interop marshalling

				omAppProc.omAppProcBO valBO = new omAppProcBOClass();
				try
				{
					strResponse = valBO.OmAppProcRequest(requestDocument.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(valBO);
				}
				
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from " + strOperation + ". Response: " + strResponse);
				}

				tempResponse = new Omiga4Response(strResponse);

			
			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Unable to create valuation report", exp);
				}
				throw new OmigaException("Unable to create valuation report", exp);
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : An error occurred building the request for creating the valuation report", exp);
				}
				throw new OmigaException("An error occurred building the request for creating the valuation report", exp);
			}

			try
			{
				XmlDocumentEx responseDocument = new XmlDocumentEx();
				responseDocument.LoadXml(strResponse);
				XmlElementEx responseElement = (XmlElementEx) responseDocument.SelectMandatorySingleNode(".//RESPONSE");
			
				if(strInstNo != "")
				{
					strRetInstNo = strInstNo;
				}
				else
				{
					strRetInstNo = responseElement.GetAttribute("INSTRUCTIONSEQUENCENO");
				}
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " : Invalid response returned from " + strOperation, exp);
				}
				throw new OmigaException("Invalid response returned from " + strOperation, exp);
			}

			return(strRetInstNo);
		}

		private bool CreateApplicationLock(string ApplicationNumber)
		{
			const string functionName = "CreateApplicationLock";
	
			try
			{
				//Build Request for Application Lock
				XmlDocumentEx createLockRequest = new XmlDocumentEx();
				XmlElementEx tempElement;

				XmlElementEx rootElement = (XmlElementEx) createLockRequest.CreateElement("REQUEST");
				createLockRequest.AppendChild(rootElement);

				string sChannelID = new GlobalParameter("DefaultChannelId").String;
					
				rootElement.SetAttribute("UNITID", _unitId);
				rootElement.SetAttribute("USERID", _userId);
				rootElement.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
				rootElement.SetAttribute("CHANNELID", _channelId);

				XmlElementEx appLockElement = (XmlElementEx)createLockRequest.CreateElement("APPLICATIONLOCK");
				rootElement.AppendChild(appLockElement);
				
				//MAR538 GHun Fix CreateLock request XML
				tempElement = (XmlElementEx)createLockRequest.CreateElement("APPLICATIONNUMBER");
				tempElement.InnerText = ApplicationNumber;
				appLockElement.AppendChild(tempElement);
				tempElement = (XmlElementEx)createLockRequest.CreateElement("UNITID");
				tempElement.InnerText = _unitId;
				appLockElement.AppendChild(tempElement);
				tempElement = (XmlElementEx)createLockRequest.CreateElement("USERID");
				tempElement.InnerText = _userId;
				appLockElement.AppendChild(tempElement);
				tempElement = (XmlElementEx)createLockRequest.CreateElement("MACHINEID");
				tempElement.InnerText = System.Environment.MachineName;
				appLockElement.AppendChild(tempElement);
				tempElement = (XmlElementEx)createLockRequest.CreateElement("CHANNELID");
				tempElement.InnerText = _channelId;
				appLockElement.AppendChild(tempElement);
				//MAR538 End

				//Call ApplicationManangerBO.CreateLock()
				ApplicationManagerBOClass AppManagerBO = new ApplicationManagerBOClass();

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Calling CreateLock. Request: " + createLockRequest.OuterXml);
				}
				// MAR1655 M Heys 21/04/2006 start
				string response = "";
				try
				{
					response = AppManagerBO.CreateLock(createLockRequest.OuterXml); 
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppManagerBO);
				}
				// MAR1655 M Heys 21/04/2006 end
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(functionName + " : Returned from CreateLock. Response: " + response);
				}

				Omiga4Response createLockResponse = new Omiga4Response(response);	
			}
			catch (OmigaException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " Failed to lock application " + ApplicationNumber, exp);
				}
				return false;
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + " An error occurred building the request to lock the application " + ApplicationNumber, exp);
				}
				return false;
			}
			catch (Exception exp)
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Debug(functionName + " An unexpected error has occurred while trying to lock application " + ApplicationNumber, exp);
				}
				return false;	//MAR538 GHun return false on failure
			}
		
			return true;
		}

		private bool UnlockApplication(string ApplicationNumber)
		{
			const string cFunctionName = "UnlockApplication";
			try
			{
				//Build Request for Application unLock
				XmlDocumentEx unlockRequest = new XmlDocumentEx();

				XmlElementEx  rootElement = (XmlElementEx)unlockRequest.CreateElement("REQUEST");
				unlockRequest.AppendChild(rootElement);
							
				rootElement.SetAttribute("UNITID", _unitId);
				rootElement.SetAttribute("USERID", _userId);
				rootElement.SetAttribute("USERAUTHORITYLEVEL", _userAuthorityLevel);
				rootElement.SetAttribute("CHANNELID", _channelId);
				rootElement.SetAttribute("MACHINEID", _machineId);    // MAR1713

				//MAR538 GHun fix XML request
				XmlElementEx appLockElement = (XmlElementEx)unlockRequest.CreateElement("APPLICATIONLOCK");
				rootElement.AppendChild(appLockElement);
		
				XmlElementEx  appNumberElement =  (XmlElementEx)unlockRequest.CreateElement("APPLICATIONNUMBER");
				appNumberElement.InnerText = ApplicationNumber;
				appLockElement.AppendChild(appNumberElement);
					
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(cFunctionName + ": Calling UnlockApplicationAndCustomers. Request: " + unlockRequest.OuterXml);
				}

				ApplicationManagerBOClass AppManagerBO = new ApplicationManagerBOClass();
				// MAR1655 M Heys 21/04/2006 start
				string response = "";
				try
				{
					response = AppManagerBO.UnlockApplicationAndCustomers(unlockRequest.OuterXml); 
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppManagerBO);
				}
				// MAR1655 M Heys 21/04/2006 end
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(cFunctionName + ": Returned from UnlockApplicationAndCustomers. Response: " + response);
				}

				Omiga4Response unlockResponse = new Omiga4Response(response);
			}
			catch (OmigaException exp)
			{
				if (_logger.IsWarnEnabled)
				{
					_logger.Warn(cFunctionName + " Failed to unlock application " + ApplicationNumber, exp);
				}
				return false;
			}
			catch (XmlException exp)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Warn(cFunctionName + " An error occurred building the request to unlock the application " + ApplicationNumber, exp);
				}
				return false;
			}
			catch (Exception exp)
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Warn(cFunctionName + " An unexpected error has occurred while trying to unlock application " + ApplicationNumber, exp);
				}
				return false;	//MAR538 GHun return false on failure
			}

			return true;
		}

		//MAR538 GHun
		private int GetMaxCaseTaskInstance(string stageId, string taskId, string caseActivityGuid, string caseStageSeqNo)
		{
			string connString = Global.DatabaseConnectionString;
			int result = 1;
			
			try 
			{
				using (SqlConnection conn = new SqlConnection(connString))
				{
					string sql = "SELECT MAX(TaskInstance) FROM CaseTask";
					sql += " WHERE CaseActivityGuid = 0x" + caseActivityGuid;
					sql += " AND StageId = '" + stageId + "'";
					sql += " AND CaseStageSequenceNo = " + caseStageSeqNo;
					sql += " AND TaskId = '" + taskId + "'";
					
					SqlCommand cmd = new SqlCommand(sql, conn);
					cmd.CommandType = CommandType.Text;

					conn.Open();
					
					SqlDataReader sdr = cmd.ExecuteReader(CommandBehavior.SingleResult);
					if (sdr.HasRows)
					{
						sdr.Read();
						if (sdr.IsDBNull(0))
							throw new OmigaException("Unable to find maximum CaseTask TaskInstance");
						else
							result = (int) sdr.GetDecimal(0);
					}
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("GetMaxCaseTaskInstance error: " + ex.Message);
				}
			}

			return result;
		}
		//MAR538 End

		//MAR1300 GHun
		private void UpdateChangeOfProperty(string applicationNumber, string applicationFactFindNumber)
		{
			string strConnectionString = Global.DatabaseConnectionString;

			using(SqlConnection oConn = new SqlConnection(strConnectionString))
			{
				SqlCommand oComm = new SqlCommand("USP_ClearChangeOfProperty", oConn);
				oComm.CommandType = CommandType.StoredProcedure;

				SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, applicationNumber.Length);
				oParam.Value = applicationNumber;

				oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
				oParam.Value = Convert.ToInt16(applicationFactFindNumber, 10);
			
				oConn.Open();
				oComm.ExecuteNonQuery();
			}
		}
		//MAR1300 End
	}
}

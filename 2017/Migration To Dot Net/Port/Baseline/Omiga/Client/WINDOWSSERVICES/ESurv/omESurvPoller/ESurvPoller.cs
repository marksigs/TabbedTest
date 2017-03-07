/*
--------------------------------------------------------------------------------------------
Workfile:			ESurvPoller.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		The ESurv interface is an external interface to a evaluation bureau.
					Applicant and propery details are sent to ESurv via ING Direct Gateway along with the type of valuation required.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		10/10/2005	MAR296 Created
JD		08/11/2005	MAR488	corrected ProgID in SendToQueue
GHun	28/11/2005	MAR708 applied fixes from FirstTitlePoller
GHun	01/12/2005	MAR772 sequence numbers being skipped
GHun	12/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
GHun	21/12/2005	MAR929 changed ServiceTimer_Tick to handle exceptions from GetMessage
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
GHun	20/04/2006	MAR1626 Changed SendToQueue to check for errors and retry the message
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Xml;
using System.Xml.Serialization; 
using System.Threading;
//NET Ref
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.Core;

using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
//COM Ref
using omMQ;
//WEB Ref
using Vertex.Fsd.Omiga.Windows.Service.omESurvPoller.PollValuer;

namespace Vertex.Fsd.Omiga.Windows.Service.omESurvPoller
{
	public class ESurvPoller : System.ServiceProcess.ServiceBase
	{
		/// Required designer variable.
		private bool _pausing;
		private object _processing;		//specific for a Thread

		private long _PollerSleepPeriod;
		private long _CurrentMSN;
		private string _InboundQueue;
		private string _UserId;
		private string _UnitId;
		private string _ChannelId;
		private string _UserAuthorityLevel;
		private Timer _timer = null;
		private ValuationPollingRequestType _req;
		private DirectGatewaySoapService _service = new DirectGatewaySoapService();
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
		
		private System.ComponentModel.Container components = null;
		
		private const string _sequenceName = "ESurvInboundMSN";

		public ESurvPoller()
		{
			InitializeComponent();
		}
		// The main entry point for the process
		static void Main()
		{
			System.ServiceProcess.ServiceBase[] ServicesToRun;
			ServicesToRun = new System.ServiceProcess.ServiceBase[] { new ESurvPoller() };
			System.ServiceProcess.ServiceBase.Run(ServicesToRun);
		}

		private void InitializeComponent()
		{
			this.ServiceName = "ESurvPoller";
		}
		
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		protected override void OnStart(string[] args)
		{	// TODO: Add code here to start your service.
			string funcName = "OnStart: ";
			if(_logger.IsInfoEnabled) 
				_logger.Info(funcName + "Service starting");
			try
			{
				_pausing = false;
				_processing = false;

				GetGlobParams();
				if(_logger.IsDebugEnabled) 
					_logger.Debug(funcName + "Got global parameters");

				_CurrentMSN = GetCurrSequenceNumber();
				if(_logger.IsDebugEnabled) 
					_logger.Debug(funcName + "Got current sequence number: " + _CurrentMSN.ToString());

				BuildPollRequest();		// prepare in advance

				if(_logger.IsDebugEnabled) 
					_logger.Debug(funcName + "Got Poll Request");
				if(_logger.IsInfoEnabled) 
					_logger.Info(funcName + "Service started successfully");

				Scheduler();			// init timer
			}
			catch (Exception e)
			{
				if(_logger.IsErrorEnabled) _logger.Error(funcName + "Error occurred", e); 
			}
		}
		
		private void Scheduler() 
		{
			try 
			{
				_timer = new Timer(new TimerCallback(ServiceTimer_Tick), null, 0, (long) _PollerSleepPeriod);
			}
			catch(Exception e) 
			{ 
				if(_logger.IsErrorEnabled) _logger.Error("Scheduler: Error occurred ", e);
			}
		}

		private void GetGlobParams()
		{	//get glob params
			string funcName = "GetGlobParams: ";
			_PollerSleepPeriod = GetPollerSleepPeriod();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_PollerSleepPeriod=" + Convert.ToString (_PollerSleepPeriod)); 

			_InboundQueue = GetInQueueName();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_InboundQueue=" + _InboundQueue); 

			_UserId = GetUserId();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_UserId=" + _UserId); 
			
			_UnitId = GetUnitId();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_UnitId=" + _UnitId); 
			
			_ChannelId = GetChannelId();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_ChannelId=" + _ChannelId); 
			
			_UserAuthorityLevel = GetAuthorityLevel();
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "_UserAuthorityLevel=" + _UserAuthorityLevel); 
		}

		private long GetPollerSleepPeriod()	//MAR708 GHun
		{	// glob param (10); in minutes so convert to millisecs
			long defvalue = 10;
			GlobalParameter gp = new GlobalParameter("EsurvInterfacPollerSleepPeriod");
			if (gp.Amount > 0)
				return ((long) gp.Amount) * 1000 * 60;	//MAR708 GHun
			else
				return defvalue * 1000 * 60;
		}

		private string GetInQueueName()
		{	//"ESurvInboundMSN";
			GlobalParameter gp = new GlobalParameter("EsurvInboundQueueName");
			return gp.String; 
		}

		private string GetUserId()
		{
			GlobalParameter gp = new GlobalParameter("DefaultUserId");
			return gp.String; 
		}

		private string GetUnitId()
		{
			GlobalParameter gp = new GlobalParameter("DefaultUnitId");
			return gp.String; 
		}

		private string GetChannelId()
		{
			GlobalParameter gp = new GlobalParameter("DefaultChannelId");
			return gp.String; 
		}

		private string GetAuthorityLevel()
		{
			GlobalParameter gp = new GlobalParameter("DefaultAuthorityLevel");
			return Convert.ToString (gp.Amount); 
		}

		private bool InterfaceAvailable()
		{	//get Interface Availability from InterfaceAvailability table

			string intefaceValidationType = "ES";
			bool reply = false;
			string conn = Global.DatabaseConnectionString;
			
			try
			{
				using (SqlConnection oConn = new SqlConnection(conn))
				{
					SqlCommand oComm = new SqlCommand("USP_GetInterfaceAvailability", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					SqlParameter oParam = oComm.Parameters.Add("@p_InterfaceValType", SqlDbType.NVarChar, intefaceValidationType.Length);
					oParam.Value = intefaceValidationType;

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
				//send poller to sleep for period time as defined in global parameter
				_pausing = true;
				//throw new OmigaException("Unable to retrieve interface availably", ex);
				if (_logger.IsErrorEnabled)	
					_logger.Error ("Unable to retrieve interface availably"); 
			}

			return reply;
		}

		public void GetMessage()
		{	
			string funcName = "GetMessage: ";
			if(_logger.IsDebugEnabled) 
				_logger.Debug (funcName + "Calling InterfaceAvailable"); 
			if (InterfaceAvailable() == true)
			{
				//see if any messages exist
				if(_logger.IsDebugEnabled) 
					_logger.Debug (funcName + "Calling GetPollMessage"); 
				ValuationPollingResponseType pollMsg = GetPollMessage(); 
				if (pollMsg.ErrorCode == "0")
				{	//NO ERRORS
					if(_logger.IsDebugEnabled) 
						_logger.Debug ("GetMessage: No Errors in response"); 
					string msgOrgRef = string.Empty;
					string respItemType = Convert.ToString (pollMsg.Item.GetType());
					respItemType = respItemType.Substring (respItemType.LastIndexOf(".") + 1);
					
					switch (respItemType) 
					{
						case "NoMessageDetailsType":
							if(_logger.IsDebugEnabled) 
								_logger.Debug (funcName + "No More Messages"); 
							_pausing = true;
							break;
						default:
							//proper message
							//send queue to be processed by omESsurv.InboundBO.OnMessage
							if(_logger.IsDebugEnabled) 
								_logger.Debug (funcName + "Proper Message"); 
					
							string respPollMsg = dg.GetXmlFromGatewayObject (typeof (ValuationPollingResponseType),pollMsg );
						
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug (funcName + respPollMsg); 
								_logger.Debug (funcName + "SendToQueue"); 
							}
							respItemType = respItemType.Substring(0,respItemType.LastIndexOf("Type"));
							msgOrgRef = SendToQueue (respItemType, respPollMsg);
							break;
					}
					if (!_pausing)
					{
						//set the currentMSN to the one sent back with response + 1
						long nextSeqNumber = Convert.ToInt64 (msgOrgRef) + 1; 
						//update the MessageSeuenceNumber in the SEQNextNumber
						SetNextMessageSequenceNumber (nextSeqNumber);
						_CurrentMSN = GetCurrSequenceNumber();
						if(_logger.IsDebugEnabled) 
							_logger.Debug (funcName + "_CurrentMSN=" + Convert.ToString (_CurrentMSN));
					}
				}
				else
				{	//ERROR at GetPollMessage 
					//send poller to sleep for period time as defined in global parameter
					_pausing = true;
	
					if (_logger.IsErrorEnabled)
					{
						_logger.Error (funcName + "ErrorCode " + pollMsg.ErrorCode);
						_logger.Error (funcName + "ErrorMessage " + pollMsg.ErrorMessage);
						_logger.Error (funcName + "Poller error cannot call Gateway"); 
						//_logger.Error (funcName + oe.Omiga4Error.Description);
					}
					throw new OmigaException (8514);
				}
			}
			else
			{
				_pausing = true;
				if(_logger.IsDebugEnabled) 
					_logger.Debug (funcName + "Service is not available"); 
			}
		}

		private void BuildPollRequest()
		{
			if(_logger.IsDebugEnabled) 
				_logger.Debug("Creating Polling Request");
			try
			{	
				_service.Url = dg.GetServiceUrl();
				_service.Proxy = dg.GetProxy();

				_req = new ValuationPollingRequestType();
				// init details and update back empty details
				ValuationPollingRequestDetailsType reqDetails = new ValuationPollingRequestDetailsType();
				_req.ValuationPollingRequestDetails = reqDetails;
			
				string password = String.Empty;				//GetPassword();
				//prepare request
				_req.ClientDevice =	dg.GetClientDevice();	//"Omiga";
	
				if (_UserId.Length > 0 && password.Length > 0)
				{
					_req.TellerID = _UserId;				//"9920";	
					_req.TellerPwd = password;				//"9920";
				}
				_req.ProxyID = dg.GetProxyId();
				_req.ProxyPwd = dg.GetProxyPwd();

				_req.ProductType = dg.GetProductType();
				_req.CommunicationChannel = dg.GetCommunicationChannel((_UserId.Length > 0));
				_req.ServiceName =	dg.GetServiceName(); 	
				
				_req.Operator =	dg.GetOperator();	//MAR852 GHun
				_req.SessionID = "";
				_req.CommunicationDirection = "";
				_req.CustomerNumber = "NA";			//PERSONID in PERSON CRM TABLE
				
				if(_logger.IsDebugEnabled) 
					_logger.Debug("Polling Request has been created");
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled) 
					_logger.Error("An error occurred trying to create Polling Request", e); 
			}
		}

		private ValuationPollingResponseType GetPollMessage()
		{
			string funcName = "GetPollMessage: ";
			ValuationPollingResponseType resp;
			try
			{
				resp = new  ValuationPollingResponseType();
			}
			catch
			{
				if (_logger.IsErrorEnabled)
					_logger.Error("An error occurred trying to create Polling Response Node"); 
				throw new OmigaException("ERROR: CANNOT allocate Response Node!");
			}
			try
			{
				// REQ DETAILS: set Sequence Number to check
				_req.ValuationPollingRequestDetails.SequenceNumber = Convert.ToInt64(_CurrentMSN);	//MAR772 GHun
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug (funcName + "SequenceNumber=" + Convert.ToString(_CurrentMSN));	//MAR772 GHun
					_logger.Debug (funcName + "Calling service");
				}

				// CALL SERVICE
				resp = _service.execute(_req);
				if(_logger.IsDebugEnabled) 
					_logger.Debug (funcName + "Got service response");
			}
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "SOAP exception occurred: " + ex.Message.ToString();
				if (_logger.IsErrorEnabled)
					_logger.Error("SOAP exception occurred: " + ex.Message.ToString()); 
			}
			catch(Exception e)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Exception occurred: " + e.Message.ToString();
				if (_logger.IsErrorEnabled)
					_logger.Error("Exception occurred: " + e.Message.ToString()); 
			}
			catch
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Unknown exception occurred";
				if (_logger.IsErrorEnabled)
					_logger.Error("Unknown exception occurred"); 
			} 
			return resp;
		}

		private string SendToQueue(string respItemType, string respESurvMsg)
		{	//send to queue to be processed by omESurv.InboundBO.OmMessage

			// REQUEST OPERATION...
			//	MESSAGEQUEUE QUEUENAME... PROGID...
			//	  XML
			//      REQUEST USERAUTHORITYLEVEL... MACHINEID... USERID... UNITID... OPERATION ...
			
			string funcName = "SendToQueue: ";
			//convert response to xml doc	
			XmlDocument doc = new XmlDocument();
			doc.LoadXml (respESurvMsg);
			string msgOrgRef = doc.SelectSingleNode("//MessageOriginatorReference").InnerXml;
			if(_logger.IsDebugEnabled) 
				_logger.Debug (funcName + "msgOrgRef=" + msgOrgRef);

			XmlDocument reqDoc = new XmlDocument();
			string response = string.Empty;
			try
			{	// Prepare the request for Queue
				XmlElement rootElem = reqDoc.CreateElement("REQUEST");
				rootElem.SetAttribute("OPERATION", "SendToQueue");
				reqDoc.AppendChild(rootElem);
			
				// Set up message element
				XmlElement msgElem = reqDoc.CreateElement("MESSAGEQUEUE");
				msgElem.SetAttribute("QUEUENAME", _InboundQueue);
				msgElem.SetAttribute("PROGID", "omESurv.InboundBO"); //JD 
				rootElem.AppendChild(msgElem);

				// Set up data element
				XmlElement dataElem = reqDoc.CreateElement("XML");
				msgElem.AppendChild(dataElem);

				XmlElement dirReq = reqDoc.CreateElement("REQUEST");
				dirReq.SetAttribute("USERAUTHORITYLEVEL", _UserAuthorityLevel);
				dirReq.SetAttribute("MACHINEID", System.Environment.MachineName);
				dirReq.SetAttribute("USERID", _UserId);
				dirReq.SetAttribute("UNITID", _UnitId);
				dirReq.SetAttribute("CHANNELID", _ChannelId);	//MAR708 GHun
				dirReq.SetAttribute("OPERATION", "OnMessage");
				//dirReq.SetAttribute("INTERFACENAME", "RUNFTINITIALREQUESTINTERFACE");
			
				XmlNode dataNode = dirReq.OwnerDocument.ImportNode(doc.SelectSingleNode("//" + respItemType),true);
				dirReq.AppendChild(dataNode);
				dataElem.AppendChild (dirReq);
			
				omMQBOClass mq = new omMQBOClass();
				// MAR1603 M Heys 10/04/2006 start
				try
				{
					response = mq.omRequest(reqDoc.OuterXml);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(mq);
				}
				// MAR1603 M Heys 10/04/2006 end
				
				//MAR1626 GHun
				Omiga4Response omResponse = new Omiga4Response(response);
				if (!omResponse.isSuccess())
				{
					if(_logger.IsErrorEnabled)
						_logger.Error(funcName + "Error occurred. Response =" + omResponse.Message); 

					_pausing = true;
				}
				//MAR1626 End

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(funcName + "Request=" + reqDoc.OuterXml); 
					_logger.Debug(funcName + "Response=" + response);
				}
			}
			catch ( Exception e)
			{
				_pausing = true;	//MAR1626 GHun
				if (_logger.IsErrorEnabled)
					_logger.Error(funcName + "Error occurred " + e.Message); 
			}
			finally
			{
				doc = null;
				reqDoc = null;
			}
			return msgOrgRef;
		}

		//Retrieve a sequence number from the SEQNEXTNUMBER table without updating
		private long GetCurrSequenceNumber()
		{
			string conn = Global.DatabaseConnectionString;
			long currNumber = 0;
			try
			{
				using(SqlConnection oConn = new SqlConnection(conn))
				{
					SqlCommand oComm = new SqlCommand("USP_GetCurrSequenceNumber", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
			
					SqlParameter paramSName = oComm.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, _sequenceName.Length); //MAR708 GHun
					paramSName.Value = _sequenceName;	//MAR708 GHun

					SqlParameter paramNextNumber = oComm.Parameters.Add("@p_NextNumber", SqlDbType.NVarChar, 12);
					paramNextNumber.Direction = ParameterDirection.Output;

					oConn.Open();
					oComm.ExecuteNonQuery();

					currNumber = Convert.ToInt64 (paramNextNumber.Value);
				}
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
					_logger.Error("GetCurrSequenceNumber: Error occurred " + e.Message); 
			}
			return currNumber;
		}

		//Set next sequence number at SEQNEXTNUMBER table 
		private void SetNextMessageSequenceNumber (long nextSeqNumber)
		{
			string conn = Global.DatabaseConnectionString;
			string seqNo = nextSeqNumber.ToString();

			try
			{
				using (SqlConnection oConn = new SqlConnection(conn))
				{
					SqlCommand oComm = new SqlCommand("USP_SetCurrSequenceNumber", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
				
					SqlParameter oParam = oComm.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, _sequenceName.Length);	//MAR708 GHun
					oParam.Value = _sequenceName;	//MAR708 GHun

					oParam = oComm.Parameters.Add("@p_NextNumber", SqlDbType.NVarChar, 12);
					oParam.Value = Convert.ToString (nextSeqNumber);

					oConn.Open();
					oComm.ExecuteNonQuery();
				}
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
					_logger.Error("SetNextMessageSequenceNumber: Error occurred " + e.Message); 
			}
		}

		//STOP Service Here
		protected override void OnStop()
		{
			if(_logger.IsInfoEnabled) 
				_logger.Info("Service stopping");
			_pausing = false;
			_processing = false;
			_timer.Change(Timeout.Infinite, Timeout.Infinite);

			if(_logger.IsInfoEnabled) 
				_logger.Info("Service stopped successfully");
		}

		private void ServiceTimer_Tick(object state) 
		{	//IMPLEMENT YOUR TIMER TICK HERE
			string funcName = "ServiceTimer_Tick: ";
			if(_logger.IsDebugEnabled) 
				_logger.Debug(funcName + "Processing");

			if ( !(bool)_processing )
			{
				Interlocked.Exchange(ref _processing, true);
				while ( ! _pausing )
				{
					try	//MAR929 GHun
					{
						if(_logger.IsDebugEnabled) 
							_logger.Debug(funcName + "Calling GetMessage");

						GetMessage();

						if(_logger.IsDebugEnabled) 
							_logger.Debug(funcName + "Got Message");
					}
					//MAR929 GHun
					catch (Exception ex)
					{
						if(_logger.IsDebugEnabled)
							_logger.Debug(funcName + "An exception occurred in GetMessage(): " + ex.Message);
						_pausing = true;
					}
					//MAR929 End
				}
				_pausing = false;
				Interlocked.Exchange(ref _processing, false);
				if(_logger.IsDebugEnabled) 
					_logger.Debug(funcName + "Finish processing. Sleeping");
			}
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			omFirstTitlePoller.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		The FirstTitle interface is an external interface to a evaluation bureau.
					Applicant and propery details are sent to FirstTitle via ING Direct Gateway along with the type of valuation required.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		10/10/2005	MAR182 Created
GHun	29/11/2005	MAR531
GHun	12/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
GHun	21/12/2005	MAR929 changed ServiceTimer_Tick to handle exceptions from GetMessage
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
MHeys	12/04/2006	MAR1603	Build errors rectified
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
using Vertex.Fsd.Omiga.Windows.Service.omFirstTitlePoller.PollConveyancer;

namespace Vertex.Fsd.Omiga.Windows.Service.omFirstTitlePoller
{
	public class FirstTitlePoller : System.ServiceProcess.ServiceBase
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
		private PollConveyancerRequestType _req;
		private DirectGatewaySoapService _service = new DirectGatewaySoapService();
		
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
		
		private System.ComponentModel.Container components = null;
		private const string _sequenceName = "FirstTitleInboundMSN";

		public FirstTitlePoller()
		{
			InitializeComponent();
		}

		// The main entry point for the process
		static void Main()
		{
			System.ServiceProcess.ServiceBase[] ServicesToRun;
			ServicesToRun = new System.ServiceProcess.ServiceBase[] { new FirstTitlePoller() };
			System.ServiceProcess.ServiceBase.Run(ServicesToRun);
		}

		private void InitializeComponent()
		{
			this.ServiceName = "FirstTitlePoller";
		}

		protected override void Dispose( bool disposing )
		{
			if(disposing)
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
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

				BuildPollRequest();		// prepeare in advance
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
				_timer = new Timer(new TimerCallback(ServiceTimer_Tick), null, 0, _PollerSleepPeriod);
			}
			catch(Exception e)
			{
				if(_logger.IsErrorEnabled)
					_logger.Error("Scheduler: Error occurred ", e);
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

		private long GetPollerSleepPeriod()
		{	// glob param (10); in minutes so convert to millisecs
			long defvalue = 10;
			GlobalParameter gp = new GlobalParameter("FirstTitleIntPollerSleepPeriod");
			if (gp.Amount > 0)
				return ((long) gp.Amount) * 1000 * 60;
			else
				return defvalue * 1000 * 60;
		}

		private string GetInQueueName()
		{	//"FirstTitleInboundMSN";
			GlobalParameter gp = new GlobalParameter("FirstTitleInboundQueueName");
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

			string intefaceValidationType = "FT";
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

				PollConveyancerResponseType pollMsg = GetPollMessage();
				if (pollMsg.ErrorCode == "0")
				{	//NO ERRORS
					if(_logger.IsDebugEnabled)
						_logger.Debug (funcName + "No Errors in response");

					ConveyancerHeaderType pollHeader = (ConveyancerHeaderType) pollMsg.Header;

					if (pollMsg.Header.MessageType == "NOM")
					{	//send poller to sleep for period time as defined in global parameter
						if(_logger.IsDebugEnabled)
							_logger.Debug (funcName + "No More Messages");
						_pausing = true;
					}
					else
					{	//proper message
						//send queue to be processed by omFirstTitle.InboundBO.OnMessage
						if(_logger.IsDebugEnabled)
							_logger.Debug (funcName + "Proper Message");
						
						string respPollMsg = dg.GetXmlFromGatewayObject (typeof (PollConveyancerResponseType),pollMsg );
						
						if(_logger.IsDebugEnabled)
						{
							_logger.Debug (funcName + respPollMsg);
							_logger.Debug (funcName + "SendToQueue");
						}
						if (SendToQueue (respPollMsg))	//MAR1626 GHun check return value
						{
							//set the currentMSN to the one sent back with response
							long nextSeqNumber = Convert.ToInt64 (pollMsg.Header.ConveyancerSequenceNo);
							//update the MessageSeuenceNumber in the SEQNextNumber
							if(_logger.IsDebugEnabled)
							{
								_logger.Debug (funcName + "Updating MSN to " + nextSeqNumber.ToString());
							}
							SetNextMessageSequenceNumber (nextSeqNumber);
							_CurrentMSN = GetCurrSequenceNumber();
						}
						//MAR1626 GHun
						else
						{
							//send poller to sleep, and the same message will be retried later
							_pausing = true;
						}
						//MAR1626 End

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
					throw new OmigaException (8533);
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
			if (_logger.IsDebugEnabled)
				_logger.Debug("Creating Polling Request");
			try
			{
				_service.Url = dg.GetServiceUrl();
				_service.Proxy = dg.GetProxy();

				_req = new  PollConveyancerRequestType();
				// init details
				PollConveyancerRequestTypeAcknowledgement reqAcknowledgement = new PollConveyancerRequestTypeAcknowledgement();
				_req.Acknowledgement = reqAcknowledgement;
				ConveyancerHeaderType reqHeader = new ConveyancerHeaderType();
				_req.Header = reqHeader;
			
				string password = String.Empty;				//GetPassword();
				//prepare request
				_req.ClientDevice =	dg.GetClientDevice();
	
				if (_UserId.Length > 0 && password.Length > 0)
				{
					_req.TellerID = _UserId;
					_req.TellerPwd = password;
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

		private PollConveyancerResponseType GetPollMessage()
		{
			string funcName = "GetPollMessage: ";
			PollConveyancerResponseType resp;
			try
			{
				resp = new PollConveyancerResponseType();
			}
			catch
			{
				if (_logger.IsErrorEnabled)
					_logger.Error("An error occurred trying to create Polling Response Node");
				throw new Exception("ERROR: CANNOT allocate Response Node!");
			}
			try
			{
				// REQ DETAILS: set Sequence Number to check
				if(_logger.IsDebugEnabled)
					_logger.Debug (funcName + "MSN Seq Number " + Convert.ToString(_CurrentMSN));
				_req.Acknowledgement.PreviousMessageSequenceNo = Convert.ToString(_CurrentMSN);
				_req.Header.Timestamp = System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss");
				// CALL SERVICE
				if(_logger.IsDebugEnabled)
					_logger.Debug (funcName + "Calling service");
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
			catch(Exception ex)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Exception occurred: " + ex.Message.ToString();
				if (_logger.IsErrorEnabled)
					_logger.Error("Exception occurred: " + ex.Message.ToString());
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

		private bool SendToQueue(string respPollMsg)	//MAR1626 GHun return bool
		{	//send to queue to be processed by omFirstTitle.InboundBO.OnMessage

			// REQUEST OPERATION...
			//	MESSAGEQUEUE QUEUENAME... PROGID...
			//	  XML
			//      REQUEST USERAUTHORITYLEVEL... MACHINEID... USERID... UNITID... OPERATION ...
			string funcName = "SendToQueue: ";
			bool success = false;	//MAR1626 GHun
			//convert eserv response to xml doc	
			XmlDocument doc = new XmlDocument();
			doc.LoadXml (respPollMsg);

			XmlDocument reqDoc = new XmlDocument();
			string response = string.Empty;
			omMQBOClass mq = new omMQBOClass();			//MAR1603 M Heys 12/04/2006
			try
			{
				// Prepare the reequest for Queue
				XmlElement rootElem = reqDoc.CreateElement("REQUEST");
				rootElem.SetAttribute("OPERATION", "SendToQueue");
				reqDoc.AppendChild(rootElem);
			
				// Set up message element
				XmlElement msgElem = reqDoc.CreateElement("MESSAGEQUEUE");
				msgElem.SetAttribute("QUEUENAME", _InboundQueue);
				msgElem.SetAttribute("PROGID", "omFirstTitle.FirstTitleInboundBO");	//MAR531 GHun
				rootElem.AppendChild(msgElem);

				// Set up data element
				XmlElement dataElem = reqDoc.CreateElement("XML");
				msgElem.AppendChild(dataElem);

				XmlElement dirReq = reqDoc.CreateElement("REQUEST");
				dirReq.SetAttribute("USERAUTHORITYLEVEL", _UserAuthorityLevel);
				dirReq.SetAttribute("MACHINEID", System.Environment.MachineName);
				dirReq.SetAttribute("USERID", _UserId);
				dirReq.SetAttribute("UNITID", _UnitId);
				dirReq.SetAttribute("CHANNELID", _ChannelId);	//MAR531 GHun
				dirReq.SetAttribute("OPERATION", "OnMessage");
				//dirReq.SetAttribute("INTERFACENAME", "RUNFTINITIALREQUESTINTERFACE");
			
				XmlNode dataNode = null;
				//MAR531 GHun
				if (doc.SelectSingleNode ("//Header") != null)
				{
					dataNode = dirReq.OwnerDocument.ImportNode(doc.SelectSingleNode("//Header"), true);
					dirReq.AppendChild(dataNode);
				}
				//MAR531 End

				if (doc.SelectSingleNode ("//StatusMessage") != null)
					dataNode = dirReq.OwnerDocument.ImportNode(doc.SelectSingleNode("//StatusMessage"), true);
				if (doc.SelectSingleNode ("//Miscellaneous") != null)
					dataNode = dirReq.OwnerDocument.ImportNode(doc.SelectSingleNode("//Miscellaneous"), true);
				
				dirReq.AppendChild(dataNode);

				dataElem.AppendChild (dirReq);
			
				//omMQBOClass mq = new omMQBOClass();			//MAR1603 M Heys 12/04/2006
				response = mq.omRequest(reqDoc.OuterXml);
				
				if(_logger.IsDebugEnabled)
				{
					_logger.Debug(funcName + "Request=" + reqDoc.OuterXml);
					_logger.Debug(funcName + "Response=" + response);
				}

				//MAR1626 GHun
				Omiga4Response omResponse = new Omiga4Response(response);
				if (omResponse.isSuccess())
					success = true;
				else
				{
					if (_logger.IsErrorEnabled)
						_logger.Error(funcName + "Error occurred. Response = " + omResponse.Message);
				}
				//MAR1626 End
			}
			catch ( Exception e)
			{
				if (_logger.IsErrorEnabled)
					_logger.Error(funcName + "Error occurred " + e.Message);
			}
			finally
			{
				doc = null;
				reqDoc = null;
				Marshal.ReleaseComObject(mq);  //MAR1603 M Heys 10/04/2006
			}
			return(success);	//MAR1626 GHun
		}

		//Retrieve a sequence number from the SEQNEXTNUMBER table without updating
		private long GetCurrSequenceNumber()
		{
			string conn = Global.DatabaseConnectionString;
			long currNumber = 0;
			
			try
			{
				using (SqlConnection oConn = new SqlConnection(conn))
				{
					SqlCommand oComm = new SqlCommand("USP_GetCurrSequenceNumber", oConn);
					oComm.CommandType = CommandType.StoredProcedure;
			
					SqlParameter paramSName = oComm.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, _sequenceName.Length);
					paramSName.Value = _sequenceName;	//MAR531 GHun;

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
				
					SqlParameter oParam = oComm.Parameters.Add("@p_SequenceName", SqlDbType.NVarChar, _sequenceName.Length);
					oParam.Value = _sequenceName;	//MAR531 GHun
					oParam = oComm.Parameters.Add("@p_NextNumber", SqlDbType.NVarChar, seqNo.Length);
					oParam.Value = seqNo;

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

			if (!(bool)_processing)
			{
				Interlocked.Exchange(ref _processing, true);
				while (!_pausing)
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
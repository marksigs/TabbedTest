/*
--------------------------------------------------------------------------------------------
Workfile:			OutBoundBO.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
JD		28/10/2005	MAR342/MAR366 Continued work. Get config correct
JD		09/11/2005	MAR488 change to call to inboundBO.HandleInitialResponse
JD		16/11/2005	MAR597 add XML correctly to call to sendtoqueue
JD		20/11/2005  MAR608 add Vendor details as vendor may be the agent.
JD		22/11/2005  MAR608 use Vendor name instead of thirdparty.companyname
GHun	17/11/2005	MAR612
GHun	29/11/2005	MAR720 Restrict extension numbers to 5 digits amd phone numbers to 17
GHun	30/11/2005	MAR744 Changed AddMessageToQueue to fix 1 hour delay
GHun	30/11/2005	MAR744 Changed AddMessageToQueue to fix EXECUTEAFTERDATE format
GHun	12/12/2005	MAR852 set Operator in generic request
GHun	15/12/2005	MAR821
RF		20/12/2005	MAR914 Lenderifdifferent should be empty. Reject telephone numbers if too long.
GHun	22/12/2005	MAR931 Fix messageType and get request attributes before InterfaceAvailability call, in case it fails
PSC		03/01/2006	MAR961 Use omLogging wrapper
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
HMA     12/06/2006  MAR527 Add code to check Interface Availability

--------------------------------------------------------------------------------------------
*/

using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Globalization;

using omMQ;
using omTM;
using MsgTm;

using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.Core;

using MQL = MESSAGEQUEUECOMPONENTVCLib;
using WebRef = Vertex.Fsd.Omiga.omESurv.WebRefValuerValuation;
using XMLManager = Vertex.Fsd.Omiga.omESurv.XML.XMLManager;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 03/01/2006 MAR961

namespace Vertex.Fsd.Omiga.omESurv
{
	/// <summary>
	/// Summary description for OutBoundBO.
	/// </summary>
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[ProgId("omESurv.OutboundBO")]
	[ComVisible(true)]
	[Guid("8A7D2ABA-4599-4c83-9AED-3A467AEF4072")]
	public class OutboundBO : MQL.IMessageQueueComponentVC2 
	{
		string _UserID;
		string _UnitID;
		string _UserAuthorityLevel;
		string _ChannelID;
		string _Password;
		XMLManager _xmlMgr = new XMLManager();
		string _strApplicationNumber;
		string _strApplicationFactFindNumber;
		// PSC 03/01/2006 MAR961 - Start
		private  omLogging.Logger _logger = null;
		private string _processContext = "Default";
		// PSC 03/01/2006 MAR961 - End

		public OutboundBO()
		{
			// PSC 03/01/2006 MAR961
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}

		// PSC 03/01/2006 MAR961 - Start
		public OutboundBO(string processContext)
		{
			if (processContext == null || processContext.Length == 0)
			{
				throw new ArgumentNullException("Process context has not been supplied");
			}

			_processContext = processContext;
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
		}
		// PSC 03/01/2006 MAR961 - End

		int MQL.IMessageQueueComponentVC2.OnMessage(string config, string xmlData)
		{
			const string cFunctionName = "OnMessage";
			int result = (int) MQL.MESSQ_RESP.MESSQ_RESP_SUCCESS;
			
			// PSC 03/01/2006 MAR961 - Start
			_processContext = "MQL";
			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, _processContext);
			// PSC 03/01/2006 MAR961 - End

			if(_logger.IsDebugEnabled) 
			{
				_logger.Error("OutboundBO." + cFunctionName + ": config=[" + config + "] xmlData=[" + xmlData + "]");
			}

			try
			{
				//MAR821 GHun Call HandleOutboundMessage through ESurvNoTxBO to escape the transaction
				ESurvNoTxBO noTxBO = new ESurvNoTxBO();
				// PSC 03/01/2006 MAR961
				noTxBO.HandleOutboundMessage(xmlData, _processContext);
				//MAR821 End
			}
			catch (Exception exp)
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Error("OutboundBO." + cFunctionName + ": An error occurred processing the message", exp);
				}

				return result;	//MAR821 GHun return success regardless of any exceptions
			}

			return result;
		}
		
		internal void HandleOutboundMessage(string xmlData)
		{
			const string cFunctionName = "HandleOutboundMessage";

			try
			{
				//MAR931 GHun Get user details before call to InterfaceAvailability
				XmlDocument xdRequest = new XmlDocument();
				xdRequest.LoadXml(xmlData);

				XmlNode requestNode = xdRequest.DocumentElement;
				_UserID				= _xmlMgr.xmlGetAttributeText(requestNode, "USERID");
				_Password			= _xmlMgr.xmlGetAttributeText(requestNode, "PASSWORD");
				_UnitID				= _xmlMgr.xmlGetAttributeText(requestNode, "UNITID");
				_UserAuthorityLevel = _xmlMgr.xmlGetAttributeText(requestNode, "USERAUTHORITYLEVEL");
				//MAR931 End

				if (InterfaceAvailability())
				{
					InboundBO _InBoundBO = new InboundBO();
					
					WebRef.RealtimeValuationResponseType response;
					
					//send request to gateway
					response = SendToDG(xmlData);
					
					//Check errorcode.
					if(response.ErrorCode == "0" || response.ErrorCode == "1" || response.ErrorCode == "2")
					{
						//HandleInitialResponse
						//MAR612 GHun
						string result = "0";
						switch (response.RealtimeValuationResponseDetails.Result)
						{
							case WebRef.RealtimeValuationResponseDetailsTypeResult.Item0:
								result = "0";
								break;
							case WebRef.RealtimeValuationResponseDetailsTypeResult.Item1:
								result = "1";
								break;
							case WebRef.RealtimeValuationResponseDetailsTypeResult.Item2:
								result = "2";
								break;
							case WebRef.RealtimeValuationResponseDetailsTypeResult.Item3:
								result = "3";
								break;
							case WebRef.RealtimeValuationResponseDetailsTypeResult.Item4:
								result = "4";
								break;
						}

						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug("OutBoundBO." + cFunctionName + ": calling HandleInitialResponse: " + _strApplicationNumber + ":" + _strApplicationFactFindNumber + ":" + _UserID + ":" + _UnitID + ":" + _UserAuthorityLevel + ":" + result);
						}

						_InBoundBO.HandleInitialResponse(_strApplicationNumber, _strApplicationFactFindNumber, _UserID, _UnitID, _UserAuthorityLevel, result, "");
						//MAR612 End

						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug("OutBoundBO." + cFunctionName + ": HandleInitialResponse completed");
						}
					}
					else
					{
						if(_logger.IsDebugEnabled) 
						{
							_logger.Debug("OutBoundBO." + cFunctionName + ": failure from ESurv (adding message back onto message queue): " + response.ErrorMessage);
						}
						
						//a failure from ESurv has occurred so we need to send the entire message again
						AddMessageToQueue(xmlData, true, true);
					}
				}
				else
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug("OutboundBO." + cFunctionName + ": Interface not available - adding message back onto the queue to try again later");
					}

					//add the message to the queue again to try later.
					AddMessageToQueue(xmlData, true, true);
				}
			}
			catch(Exception ex)
			{
				throw new OmigaException(ex);
			}
		}

		public WebRef.RealtimeValuationResponseType SendToDG(string XmlData)
		{
			string functionName = "SendToDG";
			
			WebRef.RealtimeValuationRequestType Request;
			Request = new WebRef.RealtimeValuationRequestType();
			WebRef.RealtimeValuationResponseType response = new WebRef.RealtimeValuationResponseType();
			WebRef.DirectGatewaySoapService service = new WebRef.DirectGatewaySoapService();

			try
			{
				XmlDocument xdRequest = new XmlDocument();
				xdRequest.LoadXml(XmlData);

				XmlNode appNode = xdRequest.SelectSingleNode(".//APPLICATIONFACTFIND");
				XmlNode newPropNode = xdRequest.SelectSingleNode(".//NEWPROPERTY");
				XmlNode newPropRoomType = xdRequest.SelectSingleNode(".//NEWPROPERTYROOMTYPE");
				XmlNode newPropAddress = xdRequest.SelectSingleNode(".//NEWPROPERTYADDRESS");
				XmlNode mortSubQuote = xdRequest.SelectSingleNode(".//MORTGAGESUBQUOTE");
				//XmlNode custVersion = xdRequest.SelectSingleNode(".//CUSTOMERVERSION");
				XmlNode custHomePhone = xdRequest.SelectSingleNode(".//CUSTOMERHOMETELEPHONENUMBER");
				XmlNode custWorkPhone = xdRequest.SelectSingleNode(".//CUSTOMERWORKTELEPHONENUMBER");
				XmlNode custMobilePhone = xdRequest.SelectSingleNode(".//CUSTOMERMOBILETELEPHONENUMBER");
				XmlNode custAddress = xdRequest.SelectSingleNode(".//CUSTOMERADDRESS/ADDRESS");
				XmlNode unit = xdRequest.SelectSingleNode(".//UNIT");
				XmlNode NPA = xdRequest.SelectSingleNode(".//NEWPROPERTYADDRESS/NPA");
				//JD MAR608 
				XmlNode Vendor = xdRequest.SelectSingleNode(".//VENDOR");
				XmlNode agentNode = xdRequest.SelectSingleNode(".//VENDOR/AGENTNAME");

				//Set up header info from omDG
				DirectGatewayBO dg = new DirectGatewayBO();
			
				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();

				// Get common items from omDG
				Request.ClientDevice = dg.GetClientDevice();
				Request.ServiceName = dg.GetServiceName();
				if ((_UserID.Length > 0) && (_Password.Length > 0))
				{
					Request.TellerID = _UserID;
					Request.TellerPwd = _Password;
				}
				Request.ProxyID = dg.GetProxyId();
				Request.ProxyPwd = dg.GetProxyPwd();

				Request.ProductType = dg.GetProductType();

				Request.Operator = dg.GetOperator();	//MAR852 GHun
				Request.SessionID = "";
				Request.CommunicationChannel = dg.GetCommunicationChannel(false);	//MAR612 GHun
				Request.CommunicationDirection = "";
				Request.CustomerNumber = "NA";	//OTHERSYSTEMCUSTOMERNUMBER

				//Set up RealtimeValuationRequestDetails
				//MAR612 GHun
				_strApplicationNumber = _xmlMgr.xmlGetAttributeText(appNode, "APPLICATIONNUMBER");
				_strApplicationFactFindNumber = _xmlMgr.xmlGetAttributeText(appNode ,"APPLICATIONFACTFINDNUMBER");
				//MAR612 End

				WebRef.RealtimeValuationRequestDetailsType RealtimeValuationRequestDetails = new WebRef.RealtimeValuationRequestDetailsType();
				RealtimeValuationRequestDetails.MessageType = WebRef.RealtimeValuationRequestDetailsTypeMessageType.SurveyInstruction;
				RealtimeValuationRequestDetails.MessageTimeStamp = DateTime.Today.ToString("yyyy/MM/dd hh:mm:ss");
				//This is the MSN
				RealtimeValuationRequestDetails.MessageOriginatorReference = dg.GetNextMessageSequenceNumber("ESurvOutboundMSN");

				Request.RealtimeValuationRequestDetails = RealtimeValuationRequestDetails;

				//Set up FeedInData
				WebRef.RealtimeValuationRequestDetailsTypeFeedInData FeedInData;
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetails InstructionDetails;
				FeedInData = new WebRef.RealtimeValuationRequestDetailsTypeFeedInData();
				InstructionDetails = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetails();
				FeedInData.InstructionDetails = InstructionDetails;

				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetailsInstructionSource InstructionSource;
				InstructionSource = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetailsInstructionSource();
				InstructionSource = WebRef.RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetailsInstructionSource.INGDirectNV;
				FeedInData.InstructionDetails.InstructionSource = InstructionSource;
				WebRef.PhoneDetailsType PhoneDetailsType;
				PhoneDetailsType = new WebRef.PhoneDetailsType();
				PhoneDetailsType.AreaCode = "0";
				PhoneDetailsType.LocalNumber = "0";
				PhoneDetailsType.Extension = "0";
				FeedInData.InstructionDetails.InstTelephone = PhoneDetailsType;
				FeedInData.InstructionDetails.InstFax = PhoneDetailsType;
				FeedInData.InstructionDetails.InstructionSourceDX = "0";
			
				switch (_xmlMgr.xmlGetAttributeText(newPropNode, "TYPEOFINSTRUCTION"))
				{
					case "1" : FeedInData.InstructionDetails.TypeOfInstruction = WebRef.qstInstructionType.Item1; break;
					case "2" : FeedInData.InstructionDetails.TypeOfInstruction = WebRef.qstInstructionType.Item2; break;
					case "3" : FeedInData.InstructionDetails.TypeOfInstruction = WebRef.qstInstructionType.Item3; break;
					default : FeedInData.InstructionDetails.TypeOfInstruction = WebRef.qstInstructionType.Item4; break;
				}
				switch (_xmlMgr.xmlGetAttributeText(appNode, "TYPEOFAPPLICATION"))
				{
					case "1" : FeedInData.InstructionDetails.Applicationtype = WebRef.qstApplicationType.HM; break;
					case "2" : FeedInData.InstructionDetails.Applicationtype = WebRef.qstApplicationType.FT; break;
					default : FeedInData.InstructionDetails.Applicationtype = WebRef.qstApplicationType.RM; break;
				}

				Request.RealtimeValuationRequestDetails.FeedInData = FeedInData;

				//Set up LenderDetails
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataLenderDetails LenderDetails;
				LenderDetails = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataLenderDetails();

				LenderDetails.LenderIfDifferent = ""; // MAR914
				LenderDetails.LenderReference = _xmlMgr.xmlGetAttributeText(appNode, "APPLICATIONNUMBER");
				LenderDetails.LenderBranch = _xmlMgr.xmlGetAttributeText(unit, "UNITNAME");
				LenderDetails.LenderBranchTelephone	= PhoneDetailsType;
				LenderDetails.LenderStaff = "0";
				LenderDetails.LenderStaffTelephone = PhoneDetailsType;
				LenderDetails.AdvanceAmount = _xmlMgr.xmlGetAttributeText(mortSubQuote, "AMOUNTREQUESTED");
				LenderDetails.GrossFee = "0";
				WebRef.StructuredAddressDetailsType address;
				address = new WebRef.StructuredAddressDetailsType();
				LenderDetails.ReturnAddress = address;

				Request.RealtimeValuationRequestDetails.FeedInData.LenderDetails = LenderDetails;

				//Set up Property details
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataPropertyDetails PropertyDetails;
				PropertyDetails = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataPropertyDetails();

				switch(_xmlMgr.xmlGetAttributeText(newPropNode, "PROPERTYTYPE"))
				{
					case "1" : PropertyDetails.PropertyType = WebRef.qstPropertyType.HT; break;
					case "2" : PropertyDetails.PropertyType = WebRef.qstPropertyType.HS; break;
					case "3" : PropertyDetails.PropertyType = WebRef.qstPropertyType.HD; break;
					case "4" : PropertyDetails.PropertyType = WebRef.qstPropertyType.F; break;
					case "14" : PropertyDetails.PropertyType = WebRef.qstPropertyType.B; break;
					default : PropertyDetails.PropertyType = WebRef.qstPropertyType.Item; break;
				}
				PropertyDetails.NumberOfBedrooms = _xmlMgr.xmlGetAttributeText(newPropRoomType, "NUMBEROFROOMS");
				switch (_xmlMgr.xmlGetAttributeText(newPropNode, "TENURETYPE"))
				{
					case "1" : PropertyDetails.Tenure = WebRef.qstTenure.Freehold; break;
					case "2" : PropertyDetails.Tenure = WebRef.qstTenure.Leasehold; break;
					case "3" : PropertyDetails.Tenure = WebRef.qstTenure.Feuhold; break;
					case "4" : PropertyDetails.Tenure = WebRef.qstTenure.Commonhold; break;
					default : PropertyDetails.Tenure = WebRef.qstTenure.Item; break;
				}
				switch (_xmlMgr.xmlGetAttributeText(newPropNode, "NEWPROPERTYINDICATOR"))
				{
					case "Y" : PropertyDetails.NewProperty = WebRef.qstYN.Y; break;
					default : PropertyDetails.NewProperty = WebRef.qstYN.N; break;
				}
				//Plot number has to be there for validation reasons
				PropertyDetails.PlotNumber = "";
				switch(_xmlMgr.xmlGetAttributeText(newPropNode, "HOUSEBUILDERSGUARANTEE"))
				{
					case "N" : PropertyDetails.BuildingIndemnityType = WebRef.qstIndemnityType.N; break;
					case "Z" : PropertyDetails.BuildingIndemnityType = WebRef.qstIndemnityType.Z; break;
					case "P" : PropertyDetails.BuildingIndemnityType = WebRef.qstIndemnityType.P; break;
					default : PropertyDetails.BuildingIndemnityType = WebRef.qstIndemnityType.Item; break;
				}
			
				PropertyDetails.PriceOfProperty = _xmlMgr.xmlGetAttributeText(appNode, "PURCHASEPRICEORESTIMATEDVALUE");
				//set as Purchase Price if it is a new loan and Extimated Value if it is a remortgage
				if (_xmlMgr.xmlGetAttributeText(appNode, "TYPEOFAPPLICATION") == "1")
					PropertyDetails.PurchasePriceOrEstimatedValue = WebRef.qstPriceType.EV;
				else
					PropertyDetails.PurchasePriceOrEstimatedValue = WebRef.qstPriceType.PP;
				WebRef.qstName Name;
				Name = new WebRef.qstName();
				Name.Title = "";
				Name.Initials = "";
				Name.Surname = "";
				PropertyDetails.OccupierName = Name;
				WebRef.StructuredAddressDetailsType PropertyAddress = new WebRef.StructuredAddressDetailsType();
				PropertyDetails.PropertyAddress = PropertyAddress;
				if (newPropAddress != null)	//MAR931 GHun
				{
					PropertyDetails.PropertyAddress.HouseOrBuildingNumber = _xmlMgr.xmlGetAttributeText(newPropAddress, "BUILDINGORHOUSENUMBER");
					PropertyDetails.PropertyAddress.FlatNameOrNumber = _xmlMgr.xmlGetAttributeText(newPropAddress, "FLATNUMBER");
					PropertyDetails.PropertyAddress.HouseOrBuildingName = _xmlMgr.xmlGetAttributeText(newPropAddress, "BUILDINGORHOUSENAME");
					PropertyDetails.PropertyAddress.Street = _xmlMgr.xmlGetAttributeText(newPropAddress, "STREET");
					PropertyDetails.PropertyAddress.District = _xmlMgr.xmlGetAttributeText(newPropAddress, "DISTRICT");
					PropertyDetails.PropertyAddress.TownOrCity = _xmlMgr.xmlGetAttributeText(newPropAddress, "TOWN");
					PropertyDetails.PropertyAddress.County = _xmlMgr.xmlGetAttributeText(newPropAddress, "COUNTY");
					PropertyDetails.PropertyAddress.PostCode = _xmlMgr.xmlGetAttributeText(newPropAddress, "POSTCODE");
				}
				PropertyDetails.PropertyTelDay = PhoneDetailsType;
				PropertyDetails.PropertyTelEve = PhoneDetailsType;

				Request.RealtimeValuationRequestDetails.FeedInData.PropertyDetails = PropertyDetails;

				//Set up applicant details
				XmlNode app1 = xdRequest.SelectNodes(".//CUSTOMERVERSION[CUSTOMERROLE[@CUSTOMERORDER=1]]").Item(0);
				XmlNode app2 = xdRequest.SelectNodes(".//CUSTOMERVERSION[CUSTOMERROLE[@CUSTOMERORDER=2]]").Item(0);
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataApplicantDetails ApplicantDetails;
				ApplicantDetails = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataApplicantDetails();
			
				WebRef.qstName Applicant1 = new WebRef.qstName();
				ApplicantDetails.Applicant1 = Applicant1;
				ApplicantDetails.Applicant1.Title = _xmlMgr.xmlGetAttributeText(app1, "TITLE");
				string strInitials = _xmlMgr.xmlGetAttributeText(app1, "FIRSTFORENAME").Substring(0,1);
				if(_xmlMgr.xmlGetAttributeText(app1, "SECONDFORENAME") != "")
					strInitials += _xmlMgr.xmlGetAttributeText(app1, "SECONDFORENAME").Substring(0,1);
				ApplicantDetails.Applicant1.Initials = strInitials;
				ApplicantDetails.Applicant1.Surname = _xmlMgr.xmlGetAttributeText(app1, "SURNAME");
				WebRef.qstName Applicant2 = new WebRef.qstName();
				ApplicantDetails.Applicant2 = Applicant2;
				if(app2 != null)
				{
					ApplicantDetails.Applicant2.Title = _xmlMgr.xmlGetAttributeText(app2, "TITLE");
					strInitials = _xmlMgr.xmlGetAttributeText(app2, "FIRSTFORENAME").Substring(0,1);
					if(_xmlMgr.xmlGetAttributeText(app2, "SECONDFORENAME") != "")
						strInitials += _xmlMgr.xmlGetAttributeText(app2, "SECONDFORENAME").Substring(0,1);
					ApplicantDetails.Applicant2.Initials = strInitials;
					ApplicantDetails.Applicant2.Surname = _xmlMgr.xmlGetAttributeText(app2, "SURNAME");
				}
				else
				{
					ApplicantDetails.Applicant2.Title = "";
					ApplicantDetails.Applicant2.Initials = "";
					ApplicantDetails.Applicant2.Surname = "";
				}
				WebRef.StructuredAddressDetailsType ApplicantAddress = new WebRef.StructuredAddressDetailsType();
				ApplicantDetails.ApplicantAddress = ApplicantAddress;
				ApplicantDetails.ApplicantAddress.HouseOrBuildingNumber = _xmlMgr.xmlGetAttributeText(custAddress, "BUILDINGORHOUSENUMBER");
				ApplicantDetails.ApplicantAddress.FlatNameOrNumber = _xmlMgr.xmlGetAttributeText(custAddress, "FLATNUMBER");
				ApplicantDetails.ApplicantAddress.HouseOrBuildingName = _xmlMgr.xmlGetAttributeText(custAddress, "BUILDINGORHOUNSENAME");
				ApplicantDetails.ApplicantAddress.Street = _xmlMgr.xmlGetAttributeText(custAddress, "STREET");
				ApplicantDetails.ApplicantAddress.District = _xmlMgr.xmlGetAttributeText(custAddress, "DISTRICT");
				ApplicantDetails.ApplicantAddress.TownOrCity = _xmlMgr.xmlGetAttributeText(custAddress, "TOWN");
				ApplicantDetails.ApplicantAddress.County = _xmlMgr.xmlGetAttributeText(custAddress, "COUNTY");
				ApplicantDetails.ApplicantAddress.PostCode = _xmlMgr.xmlGetAttributeText(custAddress, "POSTCODE");
			
				if(custHomePhone != null)
				{
					WebRef.PhoneDetailsType ApplicantTelEve = new WebRef.PhoneDetailsType();
					ApplicantDetails.ApplicantTelEve = ApplicantTelEve;
					//MAR720 GHun
					string areaCode = _xmlMgr.xmlGetAttributeText(custHomePhone, "AREACODE");
					ApplicantDetails.ApplicantTelEve.AreaCode = areaCode;
					// MAR914 
					ApplicantDetails.ApplicantTelEve.LocalNumber = GetValidPhoneNumber(areaCode, _xmlMgr.xmlGetAttributeText(custHomePhone, "TELEPHONENUMBER"));
					//MAR720 End
					ApplicantDetails.ApplicantTelEve.Extension = "0";
					ApplicantDetails.ApplicantTelEve.PhoneType = WebRef.PhoneDetailsTypePhoneType.R;
				}
				else
					ApplicantDetails.ApplicantTelEve = PhoneDetailsType;

				if(custWorkPhone != null)
				{
					WebRef.PhoneDetailsType ApplicantTelDay = new WebRef.PhoneDetailsType();
					ApplicantDetails.ApplicantTelDay = ApplicantTelDay;
					//MAR720 GHun
					string areaCode = _xmlMgr.xmlGetAttributeText(custWorkPhone, "AREACODE");
					ApplicantDetails.ApplicantTelDay.AreaCode = areaCode;
					// MAR914 
					ApplicantDetails.ApplicantTelDay.LocalNumber = GetValidPhoneNumber(areaCode, _xmlMgr.xmlGetAttributeText(custWorkPhone, "TELEPHONENUMBER"));
					ApplicantDetails.ApplicantTelDay.Extension = TrimExtension(_xmlMgr.xmlGetAttributeText(custWorkPhone, "EXTENSIONNUMBER"));
					//MAR720 End
					ApplicantDetails.ApplicantTelDay.PhoneType = WebRef.PhoneDetailsTypePhoneType.B;
				}
				else
					ApplicantDetails.ApplicantTelDay = PhoneDetailsType;

				if(custMobilePhone != null)
				{
					WebRef.PhoneDetailsType ApplicantMobile = new WebRef.PhoneDetailsType();
					ApplicantDetails.ApplicantMobile = ApplicantMobile;
					//MAR720 GHun
					string areaCode = _xmlMgr.xmlGetAttributeText(custMobilePhone,"AREACODE");
					ApplicantDetails.ApplicantMobile.AreaCode = areaCode;
					// MAR914 
					ApplicantDetails.ApplicantMobile.LocalNumber = GetValidPhoneNumber(areaCode, _xmlMgr.xmlGetAttributeText(custMobilePhone, "TELEPHONENUMBER"));
					ApplicantDetails.ApplicantMobile.Extension = TrimExtension(_xmlMgr.xmlGetAttributeText(custMobilePhone, "EXTENSIONNUMBER"));
					//MAR720 End
					ApplicantDetails.ApplicantMobile.PhoneType = WebRef.PhoneDetailsTypePhoneType.C;
				}
				else
					ApplicantDetails.ApplicantMobile = PhoneDetailsType;


				ApplicantDetails.ApplicantEmail = _xmlMgr.xmlGetAttributeText(app1, "CONTACTEMAILADDRESS");
            
				Request.RealtimeValuationRequestDetails.FeedInData.ApplicantDetails = ApplicantDetails;

				//Set up agent details
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataAgentDetails agentDetails;
				agentDetails = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataAgentDetails();
				WebRef.PhoneDetailsType AgentTel = new WebRef.PhoneDetailsType();
				agentDetails.AgentTel = AgentTel;
				//JD MAR608 agent may be the vendor.
				if ((NPA != null) && _xmlMgr.xmlGetAttributeText(NPA, "ACCESSCONTACTNAME") != "")	//MAR931 GHun
				{
					agentDetails.Agent = _xmlMgr.xmlGetAttributeText(NPA, "ACCESSCONTACTNAME");
					//MAR720 GHun
					string areaCode = _xmlMgr.xmlGetAttributeText(NPA, "AREACODE");
					agentDetails.AgentTel.AreaCode = areaCode;
					// MAR914 
					agentDetails.AgentTel.LocalNumber = GetValidPhoneNumber(areaCode, _xmlMgr.xmlGetAttributeText(NPA, "ACCESSTELEPHONENUMBER"));
					agentDetails.AgentTel.Extension = TrimExtension(_xmlMgr.xmlGetAttributeText(NPA, "EXTENSIONNUMBER"));
					//MAR720 End
					agentDetails.AgentTel.PhoneType = WebRef.PhoneDetailsTypePhoneType.B;
				}
				else if (agentNode != null)
				{
					//JD MAR608 use the contact details name not the company name
					//agentDetails.Agent = _xmlMgr.xmlGetAttributeText(Vendor, "COMPANYNAME");
					string agentName = _xmlMgr.xmlGetAttributeText(agentNode, "CONTACTTITLE");
					agentName += " " + _xmlMgr.xmlGetAttributeText(agentNode, "CONTACTFORENAME");
					agentName += " " + _xmlMgr.xmlGetAttributeText(agentNode, "CONTACTSURNAME");
					if(agentName.Length > 40)
						agentDetails.Agent = agentName.Substring(0,40) ; // max length is 40
					else
						agentDetails.Agent = agentName;
					XmlNode CTD = agentNode.SelectSingleNode("CTD");
					//MAR720 GHun
					string areaCode = _xmlMgr.xmlGetAttributeText(CTD,"AREACODE");
					agentDetails.AgentTel.AreaCode = areaCode;
					// MAR914 
					agentDetails.AgentTel.LocalNumber = GetValidPhoneNumber(areaCode, _xmlMgr.xmlGetAttributeText(CTD, "TELENUMBER"));
					agentDetails.AgentTel.Extension = TrimExtension(_xmlMgr.xmlGetAttributeText(CTD, "EXTENSIONNUMBER"));
					//MAR720 End
					agentDetails.AgentTel.PhoneType = WebRef.PhoneDetailsTypePhoneType.R;
				}
				else
				{
					agentDetails.Agent ="NONE";
					agentDetails.AgentTel.AreaCode = "0";
					agentDetails.AgentTel.LocalNumber = "0";
					agentDetails.AgentTel.Extension = "0";
					agentDetails.AgentTel.PhoneType = WebRef.PhoneDetailsTypePhoneType.B;
				}

				Request.RealtimeValuationRequestDetails.FeedInData.AgentDetails = agentDetails;

				//Set up OtherInformation
				WebRef.RealtimeValuationRequestDetailsTypeFeedInDataOtherInformation otherInfo;
				otherInfo = new WebRef.RealtimeValuationRequestDetailsTypeFeedInDataOtherInformation();

				otherInfo.Access = "";
				otherInfo.InstructionNote1 = "";
				otherInfo.InstructionNote2 = "";
				otherInfo.InstructionNote3 = "";

				Request.RealtimeValuationRequestDetails.FeedInData.OtherInformation = otherInfo;

				//send to SOAP
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(WebRef.RealtimeValuationRequestType), Request);
					_logger.Debug(functionName + " : Service Request : " + requestString);
				}

				response = service.execute(Request);
			
				string responseString = dg.GetXmlFromGatewayObject(typeof(WebRef.RealtimeValuationResponseType), response);

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + " : Service Response : " + responseString );
				}

				//response = service.execute(Request);
			}
			// MAR914 Start
			catch(ArgumentException ex)
			{
				response.ErrorCode = "APPERR";
				response.ErrorMessage = ex.Message;	
			}
			// MAR914 End
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

			return response;
			
		}

		private bool InterfaceAvailability()
		{
			//MAR527 Check the InterfaceAvailabilityHours table for availability.
			string interfaceValidationType = "ES";
			bool reply = false;
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

			return reply;
		}

		public string RunESurvValuation(string xmlRequest)
		{
			const string cFunctionName = "RunESurvValuation";

			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("OutboundBO.RunESurvValuation called: xmlRequest=[" + xmlRequest + "]");
			}

			SqlCommand oComm;
			SqlParameter oParam;
			XmlDocument xdResponseDoc = new XmlDocument();
			StringWriter sw = new StringWriter();
			XmlDocument xmlDoc = new XmlDocument();
			XmlDocument xnRequest = new XmlDocument();
			XmlDocument xdResponse = new XmlDocument();
			XmlNode requestNode;
			
			string respType;
			
			xnRequest.LoadXml(xmlRequest);
			requestNode = _xmlMgr.xmlGetMandatoryNode(xnRequest, "//REQUEST");
				
			_UserID				= _xmlMgr.xmlGetAttributeText(requestNode, "USERID");
			_UnitID				= _xmlMgr.xmlGetAttributeText(requestNode, "UNITID");
			_UserAuthorityLevel = _xmlMgr.xmlGetAttributeText(requestNode, "USERAUTHORITYLEVEL");
			_ChannelID			= _xmlMgr.xmlGetAttributeText(requestNode, "CHANNELID");
			_Password			= _xmlMgr.xmlGetAttributeText(requestNode, "PASSWORD");

			//XmlDocument xdRequestDoc = new XmlDocument();
					
			//xnRootNode = xdRequestDoc.CreateElement("REQUEST");
			//xdRequestDoc.AppendChild(xnRootNode);

			//set up the response
			XmlNode xnResponse = xdResponseDoc.CreateElement("RESPONSE");
			XmlNode xdResponseNode = xdResponseDoc.AppendChild(xnResponse);

			//Get the data from the database
			try
			{
				XmlNode xnAppNode = xnRequest.SelectSingleNode("//APPLICATION");

				_strApplicationNumber = _xmlMgr.xmlGetAttributeText(xnAppNode ,"APPLICATIONNUMBER");
				_strApplicationFactFindNumber = _xmlMgr.xmlGetAttributeText(xnAppNode ,"APPLICATIONFACTFINDNUMBER");
			
				string strConnectionString =  Global.DatabaseConnectionString;

				using(SqlConnection oConn = new SqlConnection(strConnectionString))
				{
					oComm = new SqlCommand("USP_GETESURVINITIALREQUESTDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, _strApplicationNumber.Length);
					oParam.Value = _strApplicationNumber;

					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = Convert.ToInt16(_strApplicationFactFindNumber, 10);
			
					oConn.Open();
				
					SqlDataReader oReader = oComm.ExecuteReader();
					StringBuilder sb = new StringBuilder();
				
					sb.Append ("<XML>"); //MAR597
					do 
					{
						while (oReader.Read())
						{
							sb.Append (oReader.GetString(0));
						}
					} while (oReader.NextResult());

					sb.Append ("</XML>");
				
					oReader.Close();
				
					xmlDoc.LoadXml(sb.ToString());
				}

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("OutboundBO." + cFunctionName +": adding message to queue: " + xmlDoc.OuterXml);
				}
				//Add to queue
				AddMessageToQueue(xmlDoc.OuterXml, false, true);

				//Update the task to complete.
				//Update CaseTask as completed
				XmlNode xnCaseTaskRootNode;
				XmlDocument xdCaseTask = new XmlDocument();
				XmlNode xnCaseTask = xnRequest.SelectSingleNode("//CASETASK");

				xnCaseTaskRootNode = _xmlMgr.xmlGetRequestNode(xnRequest);
				xnCaseTaskRootNode = xdCaseTask.ImportNode(xnCaseTaskRootNode,true);
				_xmlMgr.xmlSetAttributeValue(xnCaseTaskRootNode,"OPERATION","UPDATECASETASK"); 
				
				xdCaseTask.AppendChild (xnCaseTaskRootNode);
				
				_xmlMgr.xmlSetAttributeValue(xnCaseTask,"TASKSTATUS", "40"); // Complete
				
				xnCaseTask = xdCaseTask.ImportNode(xnCaseTask,true);
				xnCaseTaskRootNode.AppendChild(xnCaseTask);
		
				MsgTmBOClass MsgTmBO = new MsgTmBOClass();
				try
				{
					xdResponse.LoadXml(MsgTmBO.TmRequest(xdCaseTask.OuterXml));
				}
				// MAR1603 M Heys 10/04/2006 start
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(MsgTmBO);
				}
				// MAR1603 M Heys 10/04/2006 end
				respType = _xmlMgr.xmlGetAttributeText(xdResponse.DocumentElement,"TYPE");
				if (respType == "APPERR")
				{
					throw new OmigaException("Error in Updating casetask :" + cFunctionName);
				}

				// Create the response node
				XmlAttribute typeAttrib;
				typeAttrib = xdResponseDoc.CreateAttribute("TYPE");
				typeAttrib.InnerText = "SUCCESS";
				
				xdResponseNode.Attributes.SetNamedItem(typeAttrib);
			}
			catch(Exception ex)
			{
				return new OmigaException(ex).ToOmiga4Response(); 
			}
			finally
			{
				oComm = null;
				oParam = null;
			}
			
			return xdResponseDoc.InnerXml;
		}

		//MAR720 GHun Truncate the extension number to 5 digits
		private string TrimExtension(string extension)
		{
			if (extension.Length > 5)
				return extension.Substring(0, 5);
			else
				return extension;
		}

		// MAR914 Start 
		// Disallow phone number (area code + " " + local number) > 17 digits 
//		//Truncate the phone number (area code + local number) to 17 digits
//		private string TrimPhoneNumber(string areaCode, string localNumber)
//		{
//			if ((areaCode.Length + localNumber.Length) > 17)
//				return localNumber.Substring(0, 17 - areaCode.Length);
//			else
//				return localNumber;
//		}
		private string GetValidPhoneNumber(string areaCode, string localNumber)
		{
			if ((areaCode.Length + localNumber.Length) > 16) 
			{
				throw new ArgumentException("Phone number exceeds 16 characters");
			}
			return localNumber;
		}
		//MAR720 End

		//MAR821 GHun
		private void AddMessageToQueue(string strData, bool addDelay, bool addToOutbound)
		{
			ESurvNTxBO ntxBO = new ESurvNTxBO();
			ntxBO.AddMessageToQueue(strData, addDelay, addToOutbound, _UserID, _UnitID, _UserAuthorityLevel, _logger);
		}
		//MAR821 End
	}
}

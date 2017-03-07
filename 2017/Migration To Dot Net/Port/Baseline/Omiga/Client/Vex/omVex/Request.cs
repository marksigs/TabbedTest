// --------------------------------------------------------------------------------------------
// Workfile:			OmigaData.cs
// Copyright:			Copyright ©2006 Vertex Financial Services
// 
// Description:		
// --------------------------------------------------------------------------------------------
// History:
// 
// Prog	Date		Description
// PE	08/11/2006	EP2_24 - Created
// PE	06/12/2006	EP2_280 - Unit test
// PE	06/12/2006	EP2_345 - Save VexInstructionSystemID	
// PE	31/01/2007	EP2_1029 - Various data mapping issues.
// PE	06/02/2007	EP2_1246 - Modification required for Gemini
// PE	08/02/2007	EP2_1024 - Add Application_ApplicationPriority property.
// PE	27/02/2007	EP2_1547 - ValuerInstruction not updating on receiving the Valuation Report
// PE	12/03/2007	EP2_1620 - Set ValuerInstruction to use NEWPROPERTYVALUATIONTYPE attribute.
// PE	21/03/2007	EP2_1578 - Modify location ValuationSatisfied retrieves xml from.
// PE	21/03/2007	EP2_2006 - Add new property Valuation_DateReceived.
// --------------------------------------------------------------------------------------------
using System;
using System.Xml;
using System.Reflection;

using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{

	public enum TaskState
	{
		Unknown = 0,
		Incomplete = 10,
		Pending = 20,
		NotApplicable = 30,
		Complete = 40,
		CarriedForward = 50,
		ChasedUp = 60,
		Cancelled = 70,
		FailedTAR = 80,
		FailedTAS = 90,
		RetryTAS = 100
	}

	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class OmigaMessageClass : XmlDocument
	{
		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		#region Setup

		public OmigaMessageClass(string xml)
		{
			base.LoadXml(xml);

			//Format DateReceived			
			String DateReceived = GetAttributeValue("//VALUATION","DATERECEIVED");

			if(DateReceived.Length != 0)
			{
				SetAttributeValue("//VALUATION","DATERECEIVED", DateTime.Parse(DateReceived).ToShortDateString());
			}			
		}

		public OmigaMessageClass()
		{
		}

		public void DefaultUserCredentials()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();

			try
			{
				this.UserID = GetDefaultValue("DefaultUserID");	
				this.UnitID = GetDefaultValue("DefaultUnitID");	
				this.UserAuthorityLevel = GetDefaultValue("DefaultAuthorityLevel");	
				this.ChannelID = GetDefaultValue("DefaultChannelID");	
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}
		}

		#endregion Setup

		#region CreateNode

		public XmlNode CreateRequestNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode = null;

			try
			{				
				RequestNode = base.SelectSingleNode("/REQUEST");
				if(RequestNode == null)
				{
					RequestNode = CreateNode("REQUEST");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(RequestNode);
		}

		public XmlNode CreateApplicationNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode ApplicationNode = null;

			try
			{
				ApplicationNode = base.SelectSingleNode("/REQUEST/APPLICATION");
				if(ApplicationNode == null)
				{
					RequestNode = CreateRequestNode();
					ApplicationNode = CreateNode(RequestNode, "APPLICATION");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(ApplicationNode);
		}

		public XmlNode CreateCaseTaskNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode CaseTaskNode = null;

			try
			{
				CaseTaskNode = base.SelectSingleNode("/REQUEST/CASETASK");
				if(CaseTaskNode == null)
				{
					RequestNode = CreateRequestNode();
					CaseTaskNode = CreateNode(RequestNode, "CASETASK");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(CaseTaskNode);
		}

		public XmlNode CreateStageTaskNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode StageTaskNode = null;

			try
			{
				StageTaskNode = base.SelectSingleNode("/REQUEST/STAGETASK");
				if(StageTaskNode == null)
				{
					RequestNode = CreateRequestNode();
					StageTaskNode = CreateNode(RequestNode, "STAGETASK");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(StageTaskNode);
		}

		public XmlNode CreateCaseStageNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode CaseStageNode = null;
			
			try
			{
				CaseStageNode = base.SelectSingleNode("/REQUEST/CASESTAGE");
				if(CaseStageNode == null)
				{
					RequestNode = CreateRequestNode();
					CaseStageNode = CreateNode(RequestNode, "CASESTAGE");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(CaseStageNode);
		}

		public XmlNode CreateCaseActivityNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode CaseActivityNode = null;

			try
			{
				CaseActivityNode = base.SelectSingleNode("/REQUEST/CASEACTIVITY");
				if(CaseActivityNode == null)
				{
					RequestNode = CreateRequestNode();
					CaseActivityNode = CreateNode(RequestNode, "CASEACTIVITY");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(CaseActivityNode);
		}

		public XmlNode CreateStageCaseNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode CaseStageNode = null;

			try
			{
				CaseStageNode = base.SelectSingleNode("/REQUEST/CASESTAGE");
				if(CaseStageNode == null)
				{
					RequestNode = CreateRequestNode();
					CaseStageNode = CreateNode(RequestNode, "CASESTAGE");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(CaseStageNode);
		}

		public XmlNode CreateInterfaceNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode InterfaceNode = null;

			try
			{
				InterfaceNode = base.SelectSingleNode("/REQUEST/INTERFACE");
				if(InterfaceNode == null)
				{
					RequestNode = CreateRequestNode();
					InterfaceNode = CreateNode(RequestNode, "INTERFACE");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(InterfaceNode);
		}

		
		public XmlNode CreateMessageSubTypeNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode InterfaceNode;
			XmlNode SubTypeListNode = null;
			XmlNode SubTypeNode = null;

			try
			{
				SubTypeNode = base.SelectSingleNode("/REQUEST/INTERFACE/MESSAGESUBTYPELIST/MESSAGESUBTYPE");
				if(SubTypeNode == null)
				{
					InterfaceNode = CreateInterfaceNode();
					SubTypeListNode = CreateNode(InterfaceNode, "MESSAGESUBTYPELIST");
					SubTypeNode = CreateNode(SubTypeListNode, "MESSAGESUBTYPE");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(SubTypeNode);
		}

		public XmlNode CreateVexNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode VexNode = null;

			try
			{
				VexNode = base.SelectSingleNode("/REQUEST/VEX");
				if(VexNode == null)
				{
					RequestNode = CreateRequestNode();
					VexNode = CreateNode(RequestNode, "VEX");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(VexNode);
		}

		public XmlNode CreateValuerInstructionsNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode InstructuctionsNode = null;

			try
			{
				InstructuctionsNode = base.SelectSingleNode("/REQUEST/VALUERINSTRUCTIONS");
				if(InstructuctionsNode == null)
				{
					RequestNode = CreateRequestNode();
					InstructuctionsNode = CreateNode(RequestNode, "VALUERINSTRUCTIONS");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(InstructuctionsNode);
		}

		public XmlNode CreateValuerInstructionNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode InstructuctionNode = null;

			try
			{
				InstructuctionNode = base.SelectSingleNode("/REQUEST/VALUERINSTRUCTION");
				if(InstructuctionNode == null)
				{
					RequestNode = CreateRequestNode();
					InstructuctionNode = CreateNode(RequestNode, "VALUERINSTRUCTION");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(InstructuctionNode);
		}

		public XmlNode CreateNewPropertyNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode NewPropertyNode = null;

			try
			{
				NewPropertyNode = base.SelectSingleNode("/REQUEST/NEWPROPERTY");
				if(NewPropertyNode == null)
				{
					RequestNode = CreateRequestNode();
					NewPropertyNode = CreateNode(RequestNode, "NEWPROPERTY");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(NewPropertyNode);
		}
		
		public XmlNode CreateValutionNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode ValuationNode = null;

			try
			{
				ValuationNode = base.SelectSingleNode("/REQUEST/VALUATION");
				if(ValuationNode == null)
				{
					RequestNode = CreateRequestNode();
					ValuationNode = CreateNode(RequestNode, "VALUATION");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(ValuationNode);
		}

		public XmlNode CreateContactDetailsNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode ContactDetailsNode = null;

			try
			{
				ContactDetailsNode = base.SelectSingleNode("/REQUEST/CONTACTDETAILS");
				if(ContactDetailsNode == null)
				{
					RequestNode = CreateRequestNode();
					ContactDetailsNode = CreateNode(RequestNode, "CONTACTDETAILS");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(ContactDetailsNode);
		}

		public XmlNode CreateLTVNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode LTVNode = null;

			try
			{
				LTVNode = base.SelectSingleNode("/REQUEST/LTV");
				if(LTVNode == null)
				{
					RequestNode = CreateRequestNode();
					LTVNode = CreateNode(RequestNode, "LTV");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(LTVNode);
		}

		public XmlNode CreateMortgageSubQuoteNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode SubQuoteNode = null;

			try
			{
				SubQuoteNode = base.SelectSingleNode("/REQUEST/MORTGAGESUBQUOTE");
				if(SubQuoteNode == null)
				{
					RequestNode = CreateRequestNode();
					SubQuoteNode = CreateNode(RequestNode, "MORTGAGESUBQUOTE");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(SubQuoteNode);		
		}

		public XmlNode CreateNewPropertyAddressNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode PropertyAddressNode = null;

			try
			{
				PropertyAddressNode = base.SelectSingleNode("/REQUEST/NEWPROPERTYADDRESS");
				if(PropertyAddressNode == null)
				{
					RequestNode = CreateRequestNode();
					PropertyAddressNode = CreateNode(RequestNode, "NEWPROPERTYADDRESS");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(PropertyAddressNode);		
		}		

		public XmlNode CreateRiskAssessmentNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode RiskAssesmentNode = null;

			try
			{
				RiskAssesmentNode = base.SelectSingleNode("/REQUEST/RISKASSESSMENT");
				if(RiskAssesmentNode == null)
				{
					RequestNode = CreateRequestNode();
					RiskAssesmentNode = CreateNode(RequestNode, "RISKASSESSMENT");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(RiskAssesmentNode);		
		}		

		public XmlNode CreatePrintDocumentDataNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode PrintDocumentDataNode = null;

			try
			{
				PrintDocumentDataNode = base.SelectSingleNode("/REQUEST/PRINTDOCUMENTDATA");
				if(PrintDocumentDataNode == null)
				{
					RequestNode = CreateRequestNode();
					PrintDocumentDataNode = CreateNode(RequestNode, "PRINTDOCUMENTDATA");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(PrintDocumentDataNode);		
		}		

		public XmlNode CreateControlDataNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode ControlDataNode = null;

			try
			{
				ControlDataNode = base.SelectSingleNode("/REQUEST/CONTROLDATA");
				if(ControlDataNode == null)
				{
					RequestNode = CreateRequestNode();
					ControlDataNode = CreateNode(RequestNode, "CONTROLDATA");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(ControlDataNode);		
		}		

		public XmlNode CreatePrintDataNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode RequestNode;
			XmlNode PrintDataNode = null;

			try
			{
				PrintDataNode = base.SelectSingleNode("/REQUEST/PRINTDATA");
				if(PrintDataNode == null)
				{
					RequestNode = CreateRequestNode();
					PrintDataNode = CreateNode(RequestNode, "PRINTDATA");
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return(PrintDataNode);		
		}		

		#endregion CreateNode

		#region Request

		public string UserID
		{			
			get
			{
				return GetAttributeValue("/REQUEST", "USERID");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "USERID", value);
			}
		}

		public string UnitID
		{			
			get
			{
				return GetAttributeValue("/REQUEST", "UNITID");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "UNITID", value);
			}
		}

		public string UserAuthorityLevel
		{
			get
			{
				return GetAttributeValue("/REQUEST", "USERAUTHORITYLEVEL");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "USERAUTHORITYLEVEL", value);
			}
		}

		public string ChannelID
		{
			get
			{
				return GetAttributeValue("/REQUEST", "CHANNELID");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "CHANNELID", value);
			}		
		}

		public string MachineID
		{
			get
			{
				return GetAttributeValue("/REQUEST", "MACHINEID");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "MACHINEID", value);
			}		
		}

		public string Operation
		{
			get
			{
				return GetAttributeValue("/REQUEST", "OPERATION");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "OPERATION", value);
				SetAttributeValue("/REQUEST", "ACTION", value);
			}
		}

		public string InstructionSequenceNo
		{
			get
			{
				return GetAttributeValue("/RESPONSE", "INSTRUCTIONSEQUENCENO");
			}
		}

		public string Context
		{
			get
			{
				return GetAttributeValue("/REQUEST", "CONTEXT");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/REQUEST", "CONTEXT", value);
			}
		}

		public bool ValuationSatisfied
		{
			get
			{
				string Status;
				bool Satisfied;

				Status = GetAttributeValue("/RESPONSE", "STATUS");
				Satisfied = (Status == "SATISIFIED");

				return Satisfied;
			}
		}

		#endregion Request

		#region Application

		public string Application_ApplicationNumber
		{
			get
			{
				return GetAttributeValue("//APPLICATION", "APPLICATIONNUMBER");
			}
			set
			{
				CreateApplicationNode();
				SetAttributeValue("/REQUEST/APPLICATION", "APPLICATIONNUMBER", value);
				SetNodeValue("/REQUEST/APPLICATION", "APPLICATIONNUMBER", value);
				CaseTask_CaseID = value;
			}		
		}

		public string Application_ApplicationFactFindNumber
		{
			get
			{
				string NodeValue;

				NodeValue = GetAttributeValue("//APPLICATION", "APPLICATIONFACTFINDNUMBER");
				return NodeValue;								
			}
			set
			{
				CreateApplicationNode();
				SetAttributeValue("/REQUEST/APPLICATION", "APPLICATIONFACTFINDNUMBER", value);
				SetNodeValue("/REQUEST/APPLICATION", "APPLICATIONFACTFINDNUMBER", value);
			}		
		}

		public string Application_ApplicationPriority
		{
			get
			{
				string NodeValue;

				NodeValue = GetAttributeValue("//APPLICATION", "APPLICATIONPRIORITY");
				return NodeValue;								
			}
			set
			{
				CreateApplicationNode();
				SetAttributeValue("/REQUEST/APPLICATION", "APPLICATIONPRIORITY", value);
			}		
		}

		#endregion Application

		#region Response

		public string Result
		{
			get
			{
				return GetAttributeValue("/RESPONSE", "TYPE");
			}
			set
			{
				CreateRequestNode();
				SetAttributeValue("/RESPONSE", "TYPE", value);
			}		
		}

		#endregion Response

		#region CaseTask

		public string CaseTask_TaskID
		{
			get
			{
				return GetAttributeValue("//CASETASK", "TASKID");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "TASKID", value);
			}
		}

		public TaskState CaseTask_TaskStatus
		{
			get
			{	
				try
				{
					return(TaskState) int.Parse(GetAttributeValue("//CASETASK", "TASKSTATUS"));
				}
				catch
				{
					return(TaskState.Unknown);
				}
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "TASKSTATUS", (int)value);
			}
		}


		public string CaseTask_CaseID
		{
			get
			{
				return GetAttributeValue("//CASETASK", "CASEID");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "CASEID", value);
			}
		}

		public string CaseTask_StageID
		{
			get
			{
				return GetAttributeValue("//CASETASK", "STAGEID");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "STAGEID", value);
			}
		}

		public string CaseTask_CaseStageSequenceNo
		{
			get
			{
				return GetAttributeValue("//CASETASK", "CASESTAGESEQUENCENO");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "CASESTAGESEQUENCENO", value);
			}
		}

		public string CaseTask_DueDate
		{
			get
			{
				return GetAttributeValue("//CASETASK", "TASKDUEDATEANDTIME");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "TASKDUEDATEANDTIME", value);
			}
		}

		public string CaseTask_TaskInstance
		{
			get
			{
				return GetAttributeValue("//CASETASK", "TASKINSTANCE");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "TASKINSTANCE", value);
			}
		}

		public string CaseTask_SourceApplication
		{
			get
			{
				return GetAttributeValue("//CASETASK", "SOURCEAPPLICATION");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASETASK", "SOURCEAPPLICATION", value);
			}
		}

		public string CaseTask_ActivityInstance
		{
			get
			{
				return GetAttributeValue("//CASETASK", "ACTIVITYINSTANCE");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASETASK", "ACTIVITYINSTANCE", value);
			}
		}

		public string CaseTask_CaseActivityGUID
		{
			get
			{
				return GetAttributeValue("//CASETASK", "CASEACTIVITYGUID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASETASK", "CASEACTIVITYGUID", value);
			}
		}

		public string CaseTask_ActivityID
		{
			get
			{
				return GetAttributeValue("//CASETASK", "ACTIVITYID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASETASK", "ACTIVITYID", value);
			}
		}

		public string CaseTask_Context
		{
			get
			{
				return GetAttributeValue("//CASETASK", "CONTEXT");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASETASK", "CONTEXT", value);
			}
		}

		#endregion CaseTask

		#region StageTask

		public string StageTask_StageID
		{
			get
			{
				return GetAttributeValue("//STAGETASK", "STAGEID");
			}
			set
			{
				CreateStageTaskNode();
				SetAttributeValue("/REQUEST/STAGETASK", "STAGEID", value);
			}
		}

		#endregion StageTask

		#region CaseStage

		public string CaseStage_StageID
		{
			get
			{
				return GetAttributeValue("//CASESTAGE", "STAGEID");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASESTAGE", "STAGEID", value);
			}
		}

		public string CaseStage_CaseStageSequenceNo
		{
			get
			{
				return GetAttributeValue("//CASESTAGE", "CASESTAGESEQUENCENO");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASESTAGE", "CASESTAGESEQUENCENO", value);
			}
		}

		public string CaseStage_SourceApplication
		{
			get
			{
				return GetAttributeValue("//CASESTAGE", "SOURCEAPPLICATION");
			}
			set
			{
				CreateCaseTaskNode();
				SetAttributeValue("/REQUEST/CASESTAGE", "SOURCEAPPLICATION", value);
			}
		}

		#endregion CaseStage

		#region Interface

		public string Interface_InterfaceType
		{
			get
			{
				return GetAttributeValue("//INTERFACE", "INTERFACETYPE");
			}
			set
			{
				CreateInterfaceNode();
				SetAttributeValue("/REQUEST/INTERFACE", "INTERFACETYPE", value);
			}		
		}

		public string Interface_MessageType
		{
			get
			{
				return GetAttributeValue("//INTERFACE", "MESSAGETYPE");
			}
			set
			{
				CreateInterfaceNode();
				SetAttributeValue("/REQUEST/INTERFACE", "MESSAGETYPE", value);
			}		
		}

		public string Interface_MessageSubType
		{
			get
			{
				return GetAttributeValue("//INTERFACE", "MESSAGESUBTYPE");
			}
			set
			{
				if(value!="0" && value.Length>0)
				{
					CreateMessageSubTypeNode();
					SetAttributeValue("/REQUEST/INTERFACE/MESSAGESUBTYPELIST/MESSAGESUBTYPE", "MESSAGESUBTYPE", value);
				}
			}		
		}

		public string Interface_TaskNotes
		{
			get
			{
				return GetAttributeValue("//INTERFACE", "TASKNOTES");
			}
			set
			{
				CreateInterfaceNode();
				SetAttributeValue("/REQUEST/INTERFACE", "TASKNOTES", value);
			}		
		}

		public string Interface_CreateTaskFlag
		{
			get
			{
				return GetAttributeValue("//INTERFACE", "CREATETASKFLAG");
			}
			set
			{
				CreateInterfaceNode();
				SetAttributeValue("/REQUEST/INTERFACE", "CREATETASKFLAG", value);
			}		
		}

		public string Interface_ResponseReason
		{
			get
			{
				return GetAttributeValue("/REQUEST/INTERFACE", "RESPONSEREASON");
			}
			set
			{
				CreateInterfaceNode();
				SetAttributeValue("/REQUEST/INTERFACE", "RESPONSEREASON", value);
			}		
		}

		#endregion Interface

		#region Vex

		public void CopyVexValuationNode()
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode VexValuationNode;
			XmlAttribute NewAttribute;
			
			VexValuationNode = SelectSingleNode("/REQUEST/VEX/VALUATION");
			if(VexValuationNode != null)
			{
				foreach(XmlAttribute ValuationAttribute in VexValuationNode.Attributes)
				{
					NewAttribute = (XmlAttribute)ValuationAttribute.Clone();
					CreateValutionNode().Attributes.Append(NewAttribute);
				}
			}
		}

		public void CopyVexContactDetailsNode()
		{
			XmlNode VexContactNode;
			XmlAttribute NewAttribute;
			
			VexContactNode = SelectSingleNode("/REQUEST/VEX/CONTACTDETAILS");
			if(VexContactNode != null)
			{
				foreach(XmlAttribute ContactAttribute in VexContactNode.Attributes)
				{
					NewAttribute = (XmlAttribute)ContactAttribute.Clone();
					CreateContactDetailsNode().Attributes.Append(NewAttribute);
				}
			}
		}

		public StatusCode Vex_ValuationStatus
		{
			get
			{
				try
				{
					return(StatusCode) int.Parse(GetAttributeValue("//VEX", "VALSTATUS"));
				}
				catch
				{
					return(StatusCode.Unknown);
				}
			}
			set
			{
				CreateVexNode();
				SetAttributeValue("/REQUEST/VEX", "VALSTATUS", (int)value);
			}		
		}

		public string Vex_ValuationReason
		{
			get
			{
				return GetAttributeValue("//VEX", "REASON");
			}
			set
			{
				CreateVexNode();
				SetAttributeValue("/REQUEST/VEX", "REASON", value);
			}		
		}

		public string Vex_AppointmentDate
		{
			get
			{
				return GetAttributeValue("//VEX", "APPOINTMENTDATE");
			}
			set
			{
				CreateVexNode();
				SetAttributeValue("/REQUEST/VEX", "APPOINTMENTDATE", value);
			}		
		}

		public Xit2ResultType Vex_ResponseCode
		{
			get
			{
				try
				{
					return(Xit2ResultType) int.Parse(GetAttributeValue("//VEX", "RESPONSECODE"));
				}
				catch
				{
					return(Xit2ResultType.ErrorXML);
				}
			}
			set
			{
				CreateVexNode();
				SetAttributeValue("REQUEST/VEX", "RESPONSECODE", (int)value);
			}		
		}

		public string Vex_NewPropertyIndicator
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "NEWPROPERTYINDICATOR");
			}
		}

		public string Vex_PresentValuation
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "PRESENTVALUATION");
			}
		}

		public string Vex_BuildingOrHouseName
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "BUILDINGORHOUSENAME");
			}
		}

		public string Vex_BuildingOrHouseNumber
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "BUILDINGORHOUSENUMBER");
			}
		}

		public string Vex_FlatNumber
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "FLATNUMBER");
			}
		}

		public string Vex_Street
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "STREET");
			}
		}

		public string Vex_District
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "DISTRICT");
			}
		}

		public string Vex_Town
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "TOWN");
			}
		}

		public string Vex_County
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "COUNTY");
			}
		}

		public string Vex_Country
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "COUNTRY");
			}
		}

		public string Vex_PostCode
		{
			get
			{
				return GetAttributeValue("//VEX/VALUATION", "POSTCODE");
			}
		}

		public string Vex_DocumentName
		{
			get
			{
				return GetAttributeValue("//VEX/PRINTDOCUMENT", "DOCUMENTNAME");
			}
		}
		public string Vex_DocumentType
		{
			get
			{
				return GetAttributeValue("//VEX/PRINTDOCUMENT", "DOCUMENTTYPE");
			}
		}

		public string Vex_DocumentContents
		{
			get
			{
				return GetAttributeValue("//VEX/PRINTDOCUMENT", "DOCUMENTCONTENTS");
			}
		}

		public string Vex_InstructionSystemID
		{
			get
			{
				return GetAttributeValue("//VEX", "SYSTEMID");
			}
		}

		#endregion Vex

		#region Activity

		public void RemoveActivityNode()
		{
			XmlNode ActivityNode;

			ActivityNode = base.SelectSingleNode("/REQUEST/CASEACTIVITY");
			if(ActivityNode != null)
			{
				try
				{
					CreateRequestNode().RemoveChild(ActivityNode);
				}
				catch(Exception e)
				{
					string x = e.Message;
				}
			}
		}

		public string Activity_CaseID
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "CASEID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "CASEID", value);
			}
		}

		public string Activity_CaseActivityGUID
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "CASEACTIVITYGUID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "CASEACTIVITYGUID", value);
			}
		}

		public string Activity_SourceApplication
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "SOURCEAPPLICATION");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "SOURCEAPPLICATION", value);
			}
		}

		public string Activity_SourceActivityID
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "ACTIVITYID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "ACTIVITYID", value);
			}
		}

		public string Activity_ActivityInstance
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "ACTIVITYINSTANCE");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "ACTIVITYINSTANCE", value);
			}
		}

		public string Activity_ActivityName
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "ACTIVITYNAME");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "ACTIVITYNAME", value);
			}
		}

		public string Activity_ActivityID
		{
			get
			{
				return GetAttributeValue("//CASEACTIVITY", "ACTIVITYID");
			}
			set
			{
				CreateCaseActivityNode();
				SetAttributeValue("/REQUEST/CASEACTIVITY", "ACTIVITYID", value);
			}
		}

		#endregion Activity

		#region ValuerInstructions

		public string ValuerInstructions_ApplicationNumber
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "APPLICATIONNUMBER");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "APPLICATIONNUMBER", value);
			}
		}

		public string ValuerInstructions_ApplicationFactFindNumber
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string ValuerInstructions_ValuationType
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "NEWPROPERTYVALUATIONTYPE");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "NEWPROPERTYVALUATIONTYPE", value);
			}
		}

		public string ValuerInstructions_AppointmentDate
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "APPOINTMENTDATE");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "APPOINTMENTDATE", value);
			}
		}

		public string ValuerInstructions_DateOfInstruction
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "DATEOFINSTRUCTION");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "DATEOFINSTRUCTION", value);
			}
		}

		public string ValuerInstructions_ValuerPanelNo
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "VALUERPANELNO");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "VALUERPANELNO", value);
			}
		}

		public string ValuerInstructions_ValuerType
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "VALUERTYPE");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "VALUERTYPE", value);
			}
		}

		public string ValuerInstructions_ValuationStatus
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "VALUATIONSTATUS");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "VALUATIONSTATUS", value);
			}
		}

		public string ValuerInstructions_InstructionSequenceNo
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "INSTRUCTIONSEQUENCENO");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "INSTRUCTIONSEQUENCENO", value);
			}
		}

		public string ValuerInstructions_VexInstructionSystemID
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTIONS", "VEXINSTRUCTIONSYSTEMID");
			}
			set
			{
				CreateValuerInstructionsNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTIONS", "VEXINSTRUCTIONSYSTEMID", value);
			}
		}

		#endregion ValuerInstructions

		#region ValuerInstruction

		public string ValuerInstruction_ApplicationNumber
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "APPLICATIONNUMBER");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "APPLICATIONNUMBER", value);
			}
		}

		public string ValuerInstruction_ApplicationFactFindNumber
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string ValuerInstruction_ValuationType
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "VALUATIONTYPE");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "VALUATIONTYPE", value);
			}
		}

		public string ValuerInstruction_AppointmentDate
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "APPOINTMENTDATE");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "APPOINTMENTDATE", value);
			}
		}

		public string ValuerInstruction_DateOfInstruction
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "DATEOFINSTRUCTION");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "DATEOFINSTRUCTION", value);
			}
		}

		public string ValuerInstruction_ValuerPanelNo
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "VALUERPANELNO");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "VALUERPANELNO", value);
			}
		}

		public string ValuerInstruction_ValuerType
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "VALUERTYPE");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "VALUERTYPE", value);
			}
		}

		public string ValuerInstruction_ValuationStatus
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "VALUATIONSTATUS");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "VALUATIONSTATUS", value);
			}
		}

		public string ValuerInstruction_InstructionSequenceNo
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "INSTRUCTIONSEQUENCENO");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "INSTRUCTIONSEQUENCENO", value);
			}
		}

		public string ValuerInstruction_VexInstructionSystemID
		{
			get
			{
				return GetAttributeValue("//VALUERINSTRUCTION", "VEXINSTRUCTIONSYSTEMID");
			}
			set
			{
				CreateValuerInstructionNode();
				SetAttributeValue("/REQUEST/VALUERINSTRUCTION", "VEXINSTRUCTIONSYSTEMID", value);
			}
		}

		#endregion ValuerInstruction

		#region Valuation

		public string Valuation_ApplicationNumber
		{
			get
			{
				return GetAttributeValue("//VALUATION", "APPLICATIONNUMBER");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "APPLICATIONNUMBER", value);
			}
		}

		public string Valuation_ApplicationFactFindNumber
		{
			get
			{
				return GetAttributeValue("//VALUATION", "APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string Valuation_InstructionSequenceNo
		{
			get
			{
				return GetAttributeValue("//VALUATION", "INSTRUCTIONSEQUENCENO");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "INSTRUCTIONSEQUENCENO", value);
			}
		}

		public string Valuation_Country
		{
			get
			{
				return GetAttributeValue("//VALUATION", "COUNTRY");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "COUNTRY", value);
			}
		}

		public string Valuation_TypeOfProperty
		{
			get
			{
				return GetAttributeValue("//VALUATION", "TYPEOFPROPERTY");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "TYPEOFPROPERTY", value);
			}
		}

		public string Valuation_PropertyDescription
		{
			get
			{
				return GetAttributeValue("//VALUATION", "PROPERTYDESCRIPTION");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "PROPERTYDESCRIPTION", value);
			}
		}

		public string Valuation_Structure
		{
			get
			{
				return GetAttributeValue("//VALUATION", "STRUCTURE");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "STRUCTURE", value);
			}
		}

		public string Valuation_MainRoof
		{
			get
			{
				return GetAttributeValue("//VALUATION", "MAINROOF");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "MAINROOF", value);
			}
		}

		public string Valuation_Tenure
		{
			get
			{
				return GetAttributeValue("//VALUATION", "TENURE");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "TENURE", value);
			}
		}

		public string Valuation_Drainage
		{
			get
			{
				return GetAttributeValue("//VALUATION", "DRAINAGE");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "DRAINAGE", value);
			}
		}

		public string Valuation_Water
		{
			get
			{
				return GetAttributeValue("//VALUATION", "WATER");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "WATER", value);
			}
		}

		public string Valuation_CertificationType
		{
			get
			{
				return GetAttributeValue("//VALUATION", "CERTIFICATIONTYPE");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "CERTIFICATIONTYPE", value);
			}
		}

		public string Valuation_OverallCondition
		{
			get
			{
				return GetAttributeValue("//VALUATION", "OVERALLCONDITION");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "OVERALLCONDITION", value);
			}
		}

		public string Valuation_RentalDemand
		{
			get
			{
				return GetAttributeValue("//VALUATION", "RENTALDEMAND");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "RENTALDEMAND", value);
			}
		}

		public string Valuation_DemandInArea
		{
			get
			{
				return GetAttributeValue("//VALUATION", "DEMANDINAREA");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "DEMANDINAREA", value);
			}
		}

		public string Valuation_Status
		{
			get
			{
				return GetAttributeValue("//VALUATION", "VALUATIONSTATUS");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "VALUATIONSTATUS", value);
			}
		}

		public string Valuation_Saleability
		{
			get
			{
				return GetAttributeValue("//VALUATION", "SALEABILITY");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "SALEABILITY", value);
			}
		}

		public string Valuation_DateReceived
		{
			get
			{
				return GetAttributeValue("//VALUATION", "DATERECEIVED");
			}
			set
			{
				CreateValutionNode();
				SetAttributeValue("/REQUEST/VALUATION", "DATERECEIVED", value);
			}
		}
		#endregion Valuation

		#region NewProperty

		public string NewProperty_ApplicationNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/APPLICATIONNUMBER");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "APPLICATIONNUMBER", value);
			}
		}

		public string NewProperty_ApplicationFactFindNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string NewProperty_ValuationType
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/VALUATIONTYPE");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "VALUATIONTYPE", value);
			}
		}

		public string NewProperty_PropertyLocation
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/PROPERTYLOCATION");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "PROPERTYLOCATION", value);
			}
		}

		public string NewProperty_NewPropertyIndicator
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/NEWPROPERTYINDICATOR");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "NEWPROPERTYINDICATOR", value);
			}
		}

		public string NewProperty_ChangeOfProperty
		{
			get
			{
				return GetNodeValue("//NEWPROPERTY/CHANGEOFPROPERTY");
			}
			set
			{
				CreateNewPropertyNode();
				SetNodeValue("/REQUEST/NEWPROPERTY", "CHANGEOFPROPERTY", value);
			}
		}

		#endregion NewProperty

		#region NewPropertyAddress
		
		public string NewPropertyAddress_ApplicationNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/APPLICATIONNUMBER");
			}
			set
			{
				CreateNewPropertyAddressNode();
				SetNodeValue("/REQUEST/NEWPROPERTYADDRESS", "APPLICATIONNUMBER", value);
			}
		}

		public string NewPropertyAddress_ApplicationFactFindNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateNewPropertyAddressNode();
				SetNodeValue("/REQUEST/NEWPROPERTYADDRESS", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string NewPropertyAddress_BuildingOrHouseName
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/BUILDINGORHOUSENAME");
			}
		}

		public string NewPropertyAddress_BuildingOrHouseNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/BUILDINGORHOUSENUMBER");
			}
		}

		public string NewPropertyAddress_FlatNumber
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/FLATNUMBER");
			}
		}

		public string NewPropertyAddress_Street
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/STREET");
			}
		}

		public string NewPropertyAddress_District
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/DISTRICT");
			}
		}

		public string NewPropertyAddress_Town
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/TOWN");
			}
		}

		public string NewPropertyAddress_County
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/COUNTY");
			}
		}

		public string NewPropertyAddress_Country
		{
			get
			{
				return GetAttributeValue("//NEWPROPERTYADDRESS/ADDRESS/COUNTRY", "TEXT");
			}
		}

		public string NewPropertyAddress_PostCode
		{
			get
			{
				return GetNodeValue("//NEWPROPERTYADDRESS/ADDRESS/POSTCODE");
			}
		}

		#endregion NewPropertyAndAddressDetails

		#region GetValuationReport

		public string GetValuationReport_InstructionSequenceNo
		{
			get
			{
				return GetAttributeValue("//GETVALUATIONREPORT", "INSTRUCTIONSEQUENCENO");
			}
		}

		#endregion GetValuationReport

		#region MortgageSubQuote

		public string MortgageSubQuote_ApplicationNumber
		{
			get
			{
				return GetNodeValue("//MORTGAGESUBQUOTE/APPLICATIONNUMBER");
			}
			set
			{
				CreateMortgageSubQuoteNode();
				SetNodeValue("REQUEST/MORTGAGESUBQUOTE", "APPLICATIONNUMBER", value);
			}
		}

		public string MortgageSubQuote_ApplicationFactFindNumber
		{
			get
			{
				return GetNodeValue("//MORTGAGESUBQUOTE/APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateMortgageSubQuoteNode();
				SetNodeValue("REQUEST/MORTGAGESUBQUOTE", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string MortgageSubQuote_AmountRequested
		{
			get
			{
				return GetNodeValue("//MORTGAGESUBQUOTE/AMOUNTREQUESTED");
			}
			set
			{
				CreateMortgageSubQuoteNode();
				SetNodeValue("REQUEST/MORTGAGESUBQUOTE", "AMOUNTREQUESTED", value);
			}
		}
		
		public string MortgageSubQuote_MortgageSubQuoteNumber
		{
			get
			{
				return GetNodeValue("//MORTGAGESUBQUOTE/MORTGAGESUBQUOTENUMBER");
			}
			set
			{
				CreateMortgageSubQuoteNode();
				SetNodeValue("REQUEST/MORTGAGESUBQUOTE", "MORTGAGESUBQUOTENUMBER", value);
			}
		}

		public string MortgageSubQuote_LTV
		{
			get
			{
				return GetNodeValue("//MORTGAGESUBQUOTE/LTV");
			}
			set
			{
				CreateMortgageSubQuoteNode();
				SetNodeValue("REQUEST/MORTGAGESUBQUOTE", "LTV", value);
			}
		}

		#endregion MortgageSubQuote

		#region LTV

		public string LTV
		{
			get
			{
				return GetNodeValue("//LTV");
			}
		}

		public string LTV_ApplicationNumber
		{
			get
			{
				return GetNodeValue("//LTV/APPLICATIONNUMBER");
			}
			set
			{
				CreateLTVNode();
				SetNodeValue("/REQUEST/LTV", "APPLICATIONNUMBER", value);
			}
		}

		public string LTV_ApplicationFactFindNumber
		{
			get
			{
				return GetNodeValue("//LTV/APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateLTVNode();
				SetNodeValue("/REQUEST/LTV", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string LTV_AmountRequested
		{
			get
			{
				return GetNodeValue("//LTV/AMOUNTREQUESTED");
			}
			set
			{
				CreateLTVNode();
				SetNodeValue("/REQUEST/LTV", "AMOUNTREQUESTED", value);
			}
		}
		
		#endregion LTV

		#region Quotation

		public string Quotation_QuotationNumber
		{
			get
			{
				return GetNodeValue("//QUOTATION/QUOTATIONNUMBER");
			}
		}

		#endregion Quotation

		#region RiskAssesment

		public string RiskAssessment_ApplicationNumber
		{
			get
			{
				return GetNodeValue("//RISKASSESSMENT/APPLICATIONNUMBER");
			}
			set
			{
				CreateRiskAssessmentNode();
				SetNodeValue("/REQUEST/RISKASSESSMENT", "APPLICATIONNUMBER", value);
			}
		}

		public string RiskAssessment_ApplicationFactFindNumber
		{
			get
			{
				return GetNodeValue("//RISKASSESSMENT/APPLICATIONFACTFINDNUMBER");
			}
			set
			{
				CreateRiskAssessmentNode();
				SetNodeValue("/REQUEST/RISKASSESSMENT", "APPLICATIONFACTFINDNUMBER", value);
			}
		}

		public string RiskAssessment_StageNumber
		{
			get
			{
				return GetNodeValue("//RISKASSESSMENT/STAGENUMBER");
			}
			set
			{
				CreateRiskAssessmentNode();
				SetNodeValue("/REQUEST/RISKASSESSMENT", "STAGENUMBER", value);
			}
		}

		#endregion RiskAssesment

		#region PrintDocumentData

		public string PrintDocumentData_CreatedBy
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "CREATEDBY");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "CREATEDBY", value);
			}
		}

		public string PrintDocumentData_FileName
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILENAME");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILENAME", value);
			}
		}

		public string PrintDocumentData_Compressed
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "COMPRESSED");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "COMPRESSED", value);
			}
		}

		public string PrintDocumentData_CompressionMethod
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "COMPRESSIONMETHOD");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "COMPRESSIONMETHOD", value);
			}
		}

		public string PrintDocumentData_DeliveryType
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "DELIVERYTYPE");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "DELIVERYTYPE", value);
			}
		}

		public string PrintDocumentData_FileGuid
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILEGUID");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILEGUID", value);
			}
		}

		public string PrintDocumentData_Fileversion
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILEVERSION");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILEVERSION", value);
			}
		}

		public string PrintDocumentData_Fileize
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILESIZE");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILESIZE", value);
			}
		}

		public string PrintDocumentData_FileContentsType
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILECONTENTS_TYPE");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILECONTENTS_TYPE", value);
			}
		}

		public string PrintDocumentData_FileContents
		{
			get
			{
				return GetAttributeValue("//PRINTDOCUMENTDATA", "FILECONTENTS");
			}
			set
			{
				CreatePrintDocumentDataNode();
				SetAttributeValue("/REQUEST/PRINTDOCUMENTDATA", "FILECONTENTS", value);
			}
		}

		#endregion PrintDocumentData

		#region ControlData

		public string ControlData_DocumentID
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "DOCUMENTID");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "DOCUMENTID", value);
			}
		}

		public string ControlData_DPSDocumentID
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "DPSDOCUMENTID");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "DPSDOCUMENTID", value);
			}
		}

		public string ControlData_HostTemplateName
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "HOSTTEMPLATENAME");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "HOSTTEMPLATENAME", value);
			}
		}

		public string ControlData_HostTemplateDescription
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "HOSTTEMPLATEDESCRIPTION");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "HOSTTEMPLATEDESCRIPTION", value);
			}
		}

		public string ControlData_DestinationType
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "DESTINATIONTYPE");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "DESTINATIONTYPE", value);
			}
		}

		public string ControlData_GeminiPrintMode
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "GEMINIPRINTMODE");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "GEMINIPRINTMODE", value);
			}
		}

		public string ControlData_DeliveryType
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "DELIVERYTYPE");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "DELIVERYTYPE", value);
			}
		}

		public string ControlData_TemplateGroupID
		{
			get
			{
				return GetAttributeValue("//CONTROLDATA", "TEMPLATEGROUPID");
			}
			set
			{
				CreateControlDataNode();
				SetAttributeValue("/REQUEST/CONTROLDATA", "TEMPLATEGROUPID", value);
			}
		}

		#endregion ControlData

		#region PrintData

		public string PrintData_ApplicationNumber
		{
			get
			{
				return GetAttributeValue("//PRINTDATA", "APPLICATIONNUMBER");
			}
			set
			{
				CreatePrintDataNode();
				SetAttributeValue("/REQUEST/PRINTDATA", "APPLICATIONNUMBER", value);
			}
		}

		#endregion PrintData

		#region Error

		public string Error_Number
		{
			get
			{
				return GetNodeValue("//ERROR/NUMBER");
			}
		}

		public string Error_Source
		{
			get
			{
				return GetNodeValue("//ERROR/SOURCE");
			}
		}

		public string Error_Description
		{
			get
			{
				return GetNodeValue("//ERROR/DESCRIPTION");
			}
		}

		#endregion Error

		#region Helpers

		public void RequestAppend(XmlNode Child)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			try
			{
				XmlNode NewNode = base.ImportNode(Child, true);
				CreateRequestNode().AppendChild(NewNode);
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}
		}

		public void RemoveChild(string XPath)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode Child;

			Child = this.SelectSingleNode(XPath);
			if(Child != null)
			{
				try
				{
					base.RemoveChild(Child);
				}
				catch(Exception NewError)
				{
					_Logger.LogError(MethodName, NewError.Message);
				}
			}
		}

		public void RemoveAttribute(string XPath, string AttributeName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode Child;

			Child = this.SelectSingleNode(XPath);
			if(Child != null)
			{
				try
				{
					Child.Attributes.RemoveNamedItem(AttributeName);
				}
				catch(Exception NewError)
				{
					_Logger.LogError(MethodName, NewError.Message);
				}
			}
		}

		public void SetAttributeValue(XmlNode ParentNode, string AttributeName, string AttributeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();

			try
			{
				XmlAttribute NewAttribute = base.CreateAttribute(AttributeName);
				NewAttribute.Value = AttributeValue;
				ParentNode.Attributes.Append(NewAttribute);				
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

		}

		public void SetAttributeValue(XmlNode ParentNode, string AttributeName, int AttributeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();

			try
			{
				XmlAttribute NewAttribute = base.CreateAttribute(AttributeName);
				NewAttribute.Value = AttributeValue.ToString();
				ParentNode.Attributes.Append(NewAttribute);				
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

		}

		public void SetAttributeValue(string XPath, string AttributeName, string AttributeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode ParentNode;

			try
			{
				ParentNode = base.SelectSingleNode(XPath);
				XmlAttribute NewAttribute = base.CreateAttribute(AttributeName);
				NewAttribute.Value = AttributeValue;
				ParentNode.Attributes.Append(NewAttribute);				
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

		}

		public void SetAttributeValue(string XPath, string AttributeName, int AttributeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode ParentNode;

			try
			{
				ParentNode = base.SelectSingleNode(XPath);
				XmlAttribute NewAttribute = base.CreateAttribute(AttributeName);
				NewAttribute.Value = AttributeValue.ToString();
				ParentNode.Attributes.Append(NewAttribute);				
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

		}

		public string GetAttributeValue(string XPath, string AttributeName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode ParentNode;
			string ReturnValue = "";

			try
			{
				ParentNode = base.SelectSingleNode(XPath);				
				if(ParentNode != null)
				{
					try
					{
						ReturnValue = ParentNode.Attributes.GetNamedItem(AttributeName).Value;
					}
					catch
					{
						ReturnValue = "";
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return ReturnValue;
		}

		public void SetNodeValue(string XPath, string NodeName, string NodeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();

			try
			{
				XmlNode ParentNode;
				XmlNode ValueNode;
				ParentNode = base.SelectSingleNode(XPath);
				
				if(ParentNode != null)
				{
					ValueNode = ParentNode.SelectSingleNode(NodeName);
					if(ValueNode == null)
					{
						ValueNode = CreateNode(ParentNode, NodeName);
					}

					ValueNode.InnerText = NodeValue;
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

		}

		public string GetNodeValue(string XPath)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			string ReturnValue = "";

			try
			{
				XmlNode ParentNode;				
				ParentNode = base.SelectSingleNode(XPath);
				
				if(ParentNode != null)
				{
					try
					{
						ReturnValue = ParentNode.InnerText;
					}
					catch
					{
						ReturnValue = "";
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return ReturnValue;
		}

		public XmlNode CreateNode(XmlNode ParentNode, string NodeName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode NewNode = null;

			try
			{
				NewNode = base.CreateElement(NodeName);
				ParentNode.PrependChild(NewNode);
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return NewNode;
		}

		public XmlNode CreateNode(string NodeName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode NewNode = null;

			try
			{
				NewNode = base.CreateElement(NodeName);
				base.AppendChild(NewNode);
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return NewNode;
		}

		public XmlNode CreateNode(XmlNode ParentNode, string NodeName, string NodeValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			XmlNode NewNode = null;

			try
			{
				NewNode = CreateNode(ParentNode, NodeName);
				NewNode.Value = NodeValue;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			return NewNode;
		}

		private string GetDefaultValue(string DefaultName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			string DefaultValue = "";

			try
			{
				GlobalParameter Default = new GlobalParameter(DefaultName);
				DefaultValue = Default.String;
				if(DefaultValue == "")
				{
					DefaultValue = Default.Amount.ToString();
				}
			}
			catch
			{
				DefaultValue = "";
			}

			return(DefaultValue);		
		}

		#endregion Helpers
	}
}
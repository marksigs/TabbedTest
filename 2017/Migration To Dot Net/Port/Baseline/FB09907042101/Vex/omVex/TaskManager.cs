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
// PE	13/12/2006	EP2_447 - Valuer Details to be captured from Xit2 Valn Report
// PE	08/01/2007	EP2_694 - xit2 - tasks not being created on receipt of messages from xit2
// PE	01/01/2007	EP2_997 - Set the appointment date & time set by the valuer on the Xit2 appointment task note.
// PE	01/02/2007	EP2_1059 - Error codes
// PE	06/02/2007	EP2_1246 - Modification required for Gemini
// PE	07/02/2007	EP2_1067 - Update CreateValuerInstruction and UpdateValuerInstruction
// PE	08/02/2007	EP2_1024 - Change task creation method to use CreateAdhocCaseTask
// PE	27/02/2007	EP2_1547 - ValuerInstruction not updating on receiving the Valuation Report
// PE	28/02/2007	EP2_1625 - ValuationStatus not set to Cancelled
// PE	05/03/2007	EP2_1741 - Extra tracing to help debug problem
// PE	09/03/2007	EP2_1741 - Set InstructionSequenceNumber correctly
// PE	12/03/2007	EP2_1620 - Set ValuerInstruction.ValuationType
// PE	19/03/2007	EP2_1620 - Set ValuerInstruction.ValuationType
// PE	21/03/2007	EP2_1578 - Add new method ProcessValuationReport.
// PE	27/03/2007	EP2_2110 - Remove UniversalTime methods calls.
// PE	27/03/2007	EP2_1889 - PropertyChanged functionality overridable by fixed global parameter.
// --------------------------------------------------------------------------------------------
using System;
using System.Xml;
using System.Reflection;

using omTM;
using omPM;
using MsgTm;
using omAppProc;
using omApp;
using Vertex.Fsd.Omiga.omLogging;
using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{

	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class TaskManagerClass
	{
		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		public bool HandleInterfaceResponse(ref OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = true;
		
			try
			{
				OmigaMessageClass TaskRequest = new OmigaMessageClass();
				Response = new OmigaMessageClass();

				TaskRequest.LoadXml(Request.InnerXml);
				TaskRequest.CaseTask_TaskStatus = TaskState.Complete;
				TaskRequest.Operation = "HandleInterfaceResponse";

				TaskRequest.Interface_InterfaceType = "60";
				TaskRequest.Interface_MessageType = Request.Vex_ResponseCode.GetHashCode().ToString();
				TaskRequest.Interface_MessageSubType = Request.Vex_ValuationStatus.GetHashCode().ToString();
				TaskRequest.Interface_CreateTaskFlag = "1";
				if(TaskRequest.Interface_TaskNotes == String.Empty)
				{
					TaskRequest.Interface_TaskNotes = TaskRequest.Vex_ValuationReason;
				}

				_Logger.Debug("TaskRequest = " + TaskRequest.InnerXml);

				OmTmBO OmigaTask = new OmTmBOClass();
				try 
				{
					Response.LoadXml(OmigaTask.OmTmRequest(TaskRequest.OuterXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		private bool UpdateTask(OmigaMessageClass Request, TaskState NewState)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				OmigaMessageClass TaskRequest = new OmigaMessageClass();
				OmigaMessageClass ResponseDocument = new OmigaMessageClass();

				TaskRequest.LoadXml(Request.InnerXml);
				TaskRequest.Operation = "UpdateCaseTask";

				ProcessOK = GetTaskList(Request, ref ResponseDocument);	
						
				TaskRequest.CaseTask_TaskStatus = NewState;			
				TaskRequest.CaseTask_CaseID = Request.Application_ApplicationNumber;
				TaskRequest.CaseTask_StageID = ResponseDocument.CaseTask_StageID;
				TaskRequest.CaseTask_CaseStageSequenceNo = ResponseDocument.CaseTask_CaseStageSequenceNo;			
				TaskRequest.CaseTask_TaskInstance = ResponseDocument.CaseTask_TaskInstance;
				TaskRequest.CaseTask_SourceApplication = ResponseDocument.CaseTask_SourceApplication;
				TaskRequest.CaseTask_DueDate = DateTime.Now.ToString();

				TaskRequest.Interface_TaskNotes = TaskRequest.Vex_ValuationReason;

				//Note that UpdateCaseTask expects to find The CaseActivity as the first node.
				TaskRequest.RemoveActivityNode();
				TaskRequest.Activity_CaseActivityGUID = ResponseDocument.CaseTask_CaseActivityGUID;
				TaskRequest.Activity_SourceApplication = ResponseDocument.CaseTask_SourceApplication;
				TaskRequest.Activity_ActivityInstance = ResponseDocument.CaseTask_ActivityInstance;
				TaskRequest.Activity_ActivityID = ResponseDocument.CaseTask_ActivityID;
				TaskRequest.Activity_CaseID = ResponseDocument.CaseTask_CaseID;
			
				MsgTm.MsgTmBO OmigaTask = new MsgTmBOClass();
				try 
				{
					ResponseDocument.LoadXml(OmigaTask.TmRequest(TaskRequest.OuterXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, ResponseDocument);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		private bool GetTaskList(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				Request.Operation = "FINDCASETASKLIST";

				MsgTm.MsgTmBO OmigaTask = new MsgTmBOClass();
				try 
				{				
					Response = new OmigaMessageClass((OmigaTask.TmRequest(Request.InnerXml)));
					ProcessOK = (Response.Result == "SUCCESS");
					if(!ProcessOK)
					{
						// Return an OK status if we couldn't find the record as this is not an error.
						if(Response.Error_Number == "-2147220492" || Response.Error_Description == "Record not found")
						{
							ProcessOK = true;
						}
						else
						{
							_Logger.ErrorCheck(MethodName, Response);
						}
					}
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		private bool GetCaseActivity(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				Request.Operation = "GETCASEACTIVITY";
				Request.Activity_CaseID = Request.Application_ApplicationNumber;

				MsgTm.MsgTmBO OmigaTask = new MsgTmBOClass();
				try 
				{				
					Response = new OmigaMessageClass((OmigaTask.TmRequest(Request.InnerXml)));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);	
		}
	
		private bool GetCaseStage(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				Request.Operation = "GETCASESTAGE";
				Request.CreateCaseStageNode();

				MsgTm.MsgTmBO OmigaTask = new MsgTmBOClass();
				try 
				{				
					Response = new OmigaMessageClass((OmigaTask.TmRequest(Request.InnerXml)));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);	
		}

		private bool GetCurrentStage(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				ProcessOK = GetCaseActivity(Request, ref Response);
				
				if(ProcessOK)
				{
					Request.Operation = "GetCurrentStage";
					Request.Activity_SourceApplication = Response.Activity_SourceApplication;
					Request.Activity_CaseID = Response.Activity_CaseID;
					Request.Activity_ActivityID = Response.Activity_ActivityID;
					Request.Activity_CaseActivityGUID = Response.Activity_CaseActivityGUID;
					Request.Activity_ActivityInstance = Response.Activity_ActivityInstance;

					MsgTm.MsgTmBO OmigaTask = new MsgTmBOClass();
					try 
					{				
						Response = new OmigaMessageClass((OmigaTask.TmRequest(Request.InnerXml)));
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
					}
				}

			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);	
		}

		public bool CreateTask(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			bool ProcessOK = false;
		
			try
			{
				OmigaMessageClass ResponseDocument = new OmigaMessageClass();								
				
				ProcessOK = GetCurrentStage(Request, ref ResponseDocument);
				if(ProcessOK)
				{
					Request.Operation = "CREATEADHOCCASETASK";
					Request.CaseTask_TaskStatus = TaskState.Incomplete;
					Request.CaseTask_StageID = ResponseDocument.CaseStage_StageID;
					Request.CaseTask_CaseStageSequenceNo = ResponseDocument.CaseStage_CaseStageSequenceNo;
					Request.CaseTask_ActivityID = Request.Activity_ActivityID;
					Request.CaseTask_ActivityInstance = Request.Activity_ActivityInstance;
					Request.CaseTask_SourceApplication = Request.Activity_SourceApplication;
					Request.Application_ApplicationPriority = "";
					
					omTM.OmTmBOClass OmigaTask = new OmTmBOClass();
					try 
					{
						ResponseDocument.LoadXml(OmigaTask.OmTmRequest(Request.OuterXml));
						ProcessOK = _Logger.ErrorCheck(MethodName, ResponseDocument);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(OmigaTask);
					}
				}
			
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool TaskExists(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Exists)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			Exists = false;

			try
			{				
				ProcessOK = GetTaskList(Request, ref Response);

				if(ProcessOK)
				{	
					if(Response.SelectNodes("/RESPONSE/CASETASK").Count > 0)
					{
						Exists = true;
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CreateTask(OmigaMessageClass Request, ref OmigaMessageClass Response, bool Complete)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				ProcessOK = CreateTask(Request, ref Response);
			
				if(ProcessOK && Complete)
				{
					ProcessOK = UpdateTask(Request, TaskState.Complete);
				}
				else
				{
					ProcessOK = UpdateTask(Request, TaskState.Incomplete);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CompleteTask(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			bool ProcessOK = false;
			bool Exists;

			try
			{
				ProcessOK = TaskExists(Request, ref Response, out Exists);
				if(ProcessOK)
				{
					if(Exists)
					{
						ProcessOK = UpdateTask(Request, TaskState.Complete);
					}
					else
					{
						ProcessOK = CreateTask(Request, ref Response, true);
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool IsTaskComplete(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Complete)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			Complete = false;

			try
			{				
				Request.CaseTask_TaskStatus = TaskState.Incomplete;
				ProcessOK = GetTaskList(Request, ref Response);

				if(ProcessOK)
				{	
					if(Response.SelectNodes("/RESPONSE/CASETASK").Count == 0)
					{
						Complete = true;
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool GetValuationTypeAndLocation(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{

			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omApp.NewPropertyBOClass NewProperty;

			try
			{
				Request.NewProperty_ApplicationNumber = Request.Application_ApplicationNumber;
				Request.NewProperty_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;

				NewProperty = new NewPropertyBOClass();

				try
				{
					Response.LoadXml(NewProperty.GetValuationTypeAndLocation(Request.InnerXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(NewProperty);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CreateValuerInstructions(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omAppProc.omAppProcBO ApplicationProcessing;

			try
			{
				ProcessOK = GetValuationTypeAndLocation(Request, ref Response);

				if(ProcessOK)
				{
					Request.Operation = "CreateValuerInstructions";
					Request.ValuerInstruction_ApplicationNumber = Request.Application_ApplicationNumber;
					Request.ValuerInstruction_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
					Request.ValuerInstruction_ValuerType = OmigaCombo.GetComboID("ValuerType", "E");										
					Request.ValuerInstruction_ValuationStatus = OmigaCombo.GetComboID("ValuationStatus", "N");
					Request.ValuerInstruction_ValuationType = Response.NewProperty_ValuationType;

					ApplicationProcessing = new omAppProc.omAppProcBOClass();
					try 
					{
						Response.LoadXml(ApplicationProcessing.OmAppProcRequest(Request.InnerXml));
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationProcessing);
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool GetValuerInstruction(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Found)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omAppProc.omAppProcBO ApplicationProcessing;
			Found = false;

			try
			{				
				Request.Operation = "GetValuerInstructions";
				Request.ValuerInstructions_ApplicationNumber = Request.Application_ApplicationNumber;
				Request.ValuerInstructions_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;		
			
				ApplicationProcessing = new omAppProc.omAppProcBOClass();
				try 
				{
					Response.LoadXml(ApplicationProcessing.OmAppProcRequest(Request.InnerXml));
					ProcessOK = (Response.Result == "SUCCESS");
					if(!ProcessOK)
					{
						// Return an OK status if we couldn't find the record as this is not an error.
						if(Response.Error_Number == "-2147220492" || Response.Error_Description == "Record not found")
						{
							ProcessOK = true;
						}
						else
						{
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
					}
					else
					{
						Found = true;
					}
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationProcessing);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CancelValuerInstruction(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omAppProc.omAppProcBO ApplicationProcessing;
			bool Found;

			try
			{
				ProcessOK = GetValuerInstruction(Request, ref Response, out Found);

				if(ProcessOK && Found)
				{				
					Request.Operation = "UpdateValuerInstructions";
					Request.ValuerInstruction_ApplicationNumber = Response.ValuerInstructions_ApplicationNumber;
					Request.ValuerInstruction_ApplicationFactFindNumber = Response.ValuerInstructions_ApplicationFactFindNumber;
					Request.ValuerInstruction_InstructionSequenceNo = Response.ValuerInstructions_InstructionSequenceNo;
					Request.ValuerInstruction_ValuationStatus = OmigaCombo.GetComboID("ValuationStatus", "CA");				

					ApplicationProcessing = new omAppProc.omAppProcBOClass();
					try 
					{
						Response.LoadXml(ApplicationProcessing.OmAppProcRequest(Request.InnerXml));
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationProcessing);
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool GetValuationReport(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Found)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			Found = false;
			OmigaMessageClass ValuationRequest;

			try
			{
				ValuationRequest = new OmigaMessageClass();

				ValuationRequest.Operation = "GetValuationReport";
				ValuationRequest.Valuation_ApplicationNumber = Request.Application_ApplicationNumber;
				ValuationRequest.Valuation_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
				ValuationRequest.Valuation_InstructionSequenceNo = Request.Valuation_InstructionSequenceNo;

				ProcessOK = RunAppProc(ValuationRequest, ref Response);
				if(ProcessOK)
				{
					ProcessOK = (Response.Result == "SUCCESS");
					if(!ProcessOK)
					{
						// Return an OK status if we couldn't find the record as this is not an error.
						if(Response.Error_Number == "-2147220492" || Response.Error_Description == "Record not found")
						{
							ProcessOK = true;
						}
						else
						{
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
					}
					else
					{
						Found = true;
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CreateValuationReport(OmigaMessageClass Request, ref OmigaMessageClass Response, out string ValuerInstructionNo)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			bool FoundInstruction;			
			bool FoundReport;
			string InstructionSequenceNo;
			ValuerInstructionNo = "";

			try
			{				
				ProcessOK = GetValuerInstruction(Request, ref Response, out FoundInstruction);

				if(ProcessOK)
				{				
					Request.Valuation_ApplicationNumber = Request.Application_ApplicationNumber;
					Request.Valuation_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
					InstructionSequenceNo = Response.ValuerInstructions_InstructionSequenceNo;

					if(!FoundInstruction)
					{
						Request.Operation = "CreateValuationReportNoInst";
					}
					else
					{

						Request.Valuation_InstructionSequenceNo = InstructionSequenceNo;
						ProcessOK = GetValuationReport(Request, ref Response, out FoundReport);
						if(ProcessOK)
						{						
							if(FoundReport)
							{
								Request.RemoveAttribute("//VALUATION", "INSTRUCTIONSEQUENCENO");
								Request.Operation = "CreateValuationReportNoInst";
							}
							else
							{
								Request.Valuation_InstructionSequenceNo = InstructionSequenceNo;
								Request.Operation = "CreateValuationReport";
							}				
						}
					}

					if(ProcessOK)
					{
						ProcessOK = RunAppProc(Request, ref Response);
					}

					if(ProcessOK)
					{
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}

					if(ProcessOK)
					{
						if(Request.Operation == "CreateValuationReportNoInst")
						{
							ValuerInstructionNo = Response.InstructionSequenceNo;
						}
						else
						{
							ValuerInstructionNo = InstructionSequenceNo;
						}
					}					

				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool UpdateValuationReport(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			bool FoundReport;

			try
			{
				ProcessOK = GetValuationReport(Request, ref Response, out FoundReport);
				if(ProcessOK)
				{
					Request.Operation = "UpdateValuationReport";

					ProcessOK = RunAppProc(Request, ref Response);
					if(ProcessOK)
					{
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool UpdateContactDetails(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			/*
			bool ProcessOK = false;
			bool FoundReport;
			omTP.ContactBOClass Contact;
			string ResponseXml;

			ProcessOK = GetValuationReport(Request, ref Response, out FoundReport);
			if(ProcessOK)
			{
				Request.CopyVexContactDetailsNode();

				Contact = new omTP.ContactBOClass();
				try 
				{
					ResponseXml = Contact.CreateContact(
					Response.LoadXml(string()Contact.CreateContact(Request.InnerXml.xml());
					ProcessOK = (Response.Result == "SUCCESS");
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(Contact);
				}
			}

			return(ProcessOK);
			*/
			return(true);
		}

		public bool GetNewProperty(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Found)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omApp.NewPropertyBOClass NewProperty;
			Found = false;

			try
			{				
				Request.NewProperty_ApplicationNumber = Request.Application_ApplicationNumber;
				Request.NewProperty_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;

				NewProperty = new NewPropertyBOClass();
				try
				{
					Response.LoadXml(NewProperty.GetValuationTypeAndLocation(Request.InnerXml));
					ProcessOK = (Response.Result == "SUCCESS");
					if(!ProcessOK)
					{
						// Return an OK status if we couldn't find the record as this is not an error.
						if(Response.Error_Number == "-2147220492" || Response.Error_Description == "Record not found")
						{
							ProcessOK = true;
						}
						else
						{
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
					}
					else
					{
						Found = true;
					}
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(NewProperty);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool PropertyChanged(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool Changed)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			Changed = false;
			bool ProcessOK = false;
			omApp.NewPropertyBOClass NewProperty;

			try
			{
				Request.NewPropertyAddress_ApplicationNumber = Request.Application_ApplicationNumber;
				Request.NewPropertyAddress_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;

				NewProperty = new NewPropertyBOClass();
				try
				{
					Response.LoadXml(NewProperty.GetNewPropertyAddress(Request.InnerXml));
					ProcessOK = (Response.Result == "SUCCESS");
					if(!ProcessOK)
					{
						// Return an OK status if we couldn't find the record as this is not an error.
						if(Response.Error_Number == "-2147220492" || Response.Error_Description == "Record not found")
						{
							ProcessOK = true;
						}
						else						
						{
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
					}
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(NewProperty);
				}

				if(ProcessOK)
				{
					//if any of the address details have changed
					if(AddressValueHasChanged(Response.NewPropertyAddress_BuildingOrHouseName, Request.Vex_BuildingOrHouseName))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_BuildingOrHouseNumber, Request.Vex_BuildingOrHouseNumber))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_FlatNumber, Request.Vex_FlatNumber))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_Street, Request.Vex_Street))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_District, Request.Vex_District))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_Town, Request.Vex_Town))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_County, Request.Vex_County))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_Country, Request.Vex_Country))
						Changed = true;

					if(AddressValueHasChanged(Response.NewPropertyAddress_PostCode, Request.Vex_PostCode))
						Changed = true;
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		private bool AddressValueHasChanged(string OldValue, string NewValue)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool Changed = false;

			try
			{
				if(OldValue.Trim() != "")
				{
					if(OldValue.ToUpper().Trim() != NewValue.ToUpper().Trim())
					{
						Changed = true;
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(Changed);
		}

		public bool UpdateNewProperty(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			bool Found = false;
			bool Changed = false;
			omApp.NewPropertyBOClass NewProperty;
			bool RunPropertyChange;

			try
			{

				ProcessOK = GetNewProperty(Request, ref Response, out Found);
				if(ProcessOK && Found)
				{

					try
					{
						GlobalParameter RunPropertyChangeGP = new GlobalParameter("Xit2RunPropertyChange");
						RunPropertyChange = RunPropertyChangeGP.Boolean;
					}
					catch
					{
						RunPropertyChange = false;
					}

					if(RunPropertyChange)
					{
						ProcessOK = PropertyChanged(Request, ref Response, out Changed);
						Request.NewProperty_ChangeOfProperty = Changed ? "1" : "0";
					}
					
					if(ProcessOK)
					{
						Request.NewProperty_ApplicationNumber = Request.Application_ApplicationNumber;
						Request.NewProperty_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
						Request.NewProperty_NewPropertyIndicator = Request.Vex_NewPropertyIndicator;						

						NewProperty = new NewPropertyBOClass();
						try
						{
							if(Found)
							{
								Response.LoadXml(NewProperty.UpdateNewProperty(Request.InnerXml));					
							}
							else
							{
								Response.LoadXml(NewProperty.CreateNewPropertyDetails(Request.InnerXml));					
							}							
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
						finally
						{
							System.Runtime.InteropServices.Marshal.ReleaseComObject(NewProperty);
						}
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool UpdateValuationInstruction(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omAppProc.omAppProcBO ApplicationProcessing;
			bool Found = false;
			string ValuationType = "";
		
			try
			{
				ProcessOK = GetValuationTypeAndLocation(Request, ref Response);

				if(ProcessOK)
				{
					ValuationType = Response.NewProperty_ValuationType;
					ProcessOK = GetValuerInstruction(Request, ref Response, out Found);
					if(ProcessOK && Found)
					{
						Request.ValuerInstruction_InstructionSequenceNo = Response.ValuerInstructions_InstructionSequenceNo;						
					}
				}

				if(ProcessOK)
				{
					Request.Operation = "UpdateValuerInstructions";
					Request.ValuerInstruction_ApplicationNumber = Request.Application_ApplicationNumber;
					Request.ValuerInstruction_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
					Request.ValuerInstruction_VexInstructionSystemID = Request.Vex_InstructionSystemID;					
					Request.ValuerInstruction_ValuationType = ValuationType;

					if(Request.Vex_AppointmentDate.Length > 0)
						Request.ValuerInstruction_AppointmentDate = DateTime.Parse(Request.Vex_AppointmentDate).ToString();

					if(Request.ValuerInstruction_DateOfInstruction.Length > 0)
						Request.ValuerInstruction_DateOfInstruction = DateTime.Parse(Request.Vex_AppointmentDate).ToString();

					GlobalParameter ValuerPanelNo = new GlobalParameter("Xit2Valuer");
					Request.ValuerInstruction_ValuerPanelNo = ValuerPanelNo.String;
					Request.ValuerInstruction_VexInstructionSystemID = Request.Vex_InstructionSystemID;

					ApplicationProcessing = new omAppProc.omAppProcBOClass();
					try 
					{
						Response.LoadXml(ApplicationProcessing.OmAppProcRequest(Request.InnerXml));
						ProcessOK = _Logger.ErrorCheck(MethodName, Response);
					}
					finally
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationProcessing);
					}
				}
			}
			catch(Exception NewError) 
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8807);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool GetAcceptedOrActiveQuoteData(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omAQ.ApplicationQuoteBOClass AppQuote;
		
			try
			{
				AppQuote = new omAQ.ApplicationQuoteBOClass();
				try 
				{
					Response.LoadXml(AppQuote.GetAcceptedOrActiveQuoteData(Request.InnerXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppQuote);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool CalculateLTV(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = true;
			string NewLTV;
			string OldLTV;
			string AmountRequested;
			string SubQuoteNumber;
			omAQ.ApplicationQuoteBOClass AppQuote;
			omCM.MortgageSubQuoteBOClass SubQuote;

			try
			{
				if(decimal.Parse(Request.Vex_PresentValuation) == 0)
				{
					NewLTV = "999";
				}
				else
				{
					ProcessOK = GetAcceptedOrActiveQuoteData(Request, ref Response);
					if(ProcessOK)
					{
						SubQuoteNumber = Response.MortgageSubQuote_MortgageSubQuoteNumber;
						OldLTV = Response.MortgageSubQuote_LTV;
						AmountRequested = Response.MortgageSubQuote_AmountRequested;
					
						Request.Operation = "CalcCostModelLTV";				
						Request.LTV_ApplicationNumber = Request.Application_ApplicationNumber;
						Request.LTV_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
						Request.LTV_AmountRequested = AmountRequested;
					
						AppQuote = new omAQ.ApplicationQuoteBOClass();
						try 
						{
							Response.LoadXml(AppQuote.CalcCostModelLTV(Request.InnerXml));
							ProcessOK = _Logger.ErrorCheck(MethodName, Response);
						}
						finally
						{
							System.Runtime.InteropServices.Marshal.ReleaseComObject(AppQuote);
						}

						if(ProcessOK)
						{
							NewLTV = Response.LTV;

							Request.MortgageSubQuote_ApplicationNumber = Request.Application_ApplicationNumber;
							Request.MortgageSubQuote_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
							Request.MortgageSubQuote_MortgageSubQuoteNumber = SubQuoteNumber;
							Request.MortgageSubQuote_LTV = NewLTV;

							SubQuote = new omCM.MortgageSubQuoteBOClass();
							try 
							{
								Response.LoadXml(SubQuote.Update(Request.InnerXml));
								ProcessOK = _Logger.ErrorCheck(MethodName, Response);
							}
							finally
							{
								System.Runtime.InteropServices.Marshal.ReleaseComObject(SubQuote);
							}
						}
					}
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8808);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool RunRiskAssesmentRules(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool ValuationSatisfied)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;
			omRA.RiskAssessmentBOClass RiskAssessment;
			ValuationSatisfied = false;

			try
			{				
				Request.Context = "VEX";
				Request.RiskAssessment_ApplicationNumber = Request.Application_ApplicationNumber;
				Request.RiskAssessment_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
				Request.Operation = "RunRiskAssessment";

				RiskAssessment = new omRA.RiskAssessmentBOClass();
				try 
				{
					Response.LoadXml(RiskAssessment.RunRiskAssessment(Request.InnerXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(RiskAssessment);
				}

				if(ProcessOK)
				{
					ValuationSatisfied = Response.ValuationSatisfied;
				}
			}
			catch(Exception NewError)
			{				
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8809);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		private bool RunAppProc(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("Request.InnerXml = " + Request.InnerXml);
			bool ProcessOK = true;
			omAppProc.omAppProcBO ApplicationProcessing;

			try
			{
				ApplicationProcessing = new omAppProc.omAppProcBOClass();
				try 
				{
					Response.LoadXml(ApplicationProcessing.OmAppProcRequest(Request.InnerXml));
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationProcessing);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool SaveValuationPDF(OmigaMessageClass Request, ref OmigaMessageClass Response)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = true;
			omPM.PrintManagerBOClass PrintManager;

			try
			{
				Request.Operation = "SavePrintDocument";
				Request.PrintDocumentData_CreatedBy = Request.UserID;

				PrintManager = new omPM.PrintManagerBOClass();
				try
				{
					Response.LoadXml(PrintManager.omRequest(Request.InnerXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);					
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(PrintManager);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool ProcessValuationReport(OmigaMessageClass Request, ref OmigaMessageClass Response, out bool ValuationSatisfied)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = true;
			omValuationRules.ValuationRulesBOClass ValuationRules;
			ValuationSatisfied = false;

			try
			{
				Request.Operation = "ProcessValuationReport";

				ValuationRules = new omValuationRules.ValuationRulesBOClass();
				try
				{
					Response.LoadXml(ValuationRules.RunValuationRules(Request.InnerXml));
					ProcessOK = _Logger.ErrorCheck(MethodName, Response);					
					_Logger.Debug("Response = " + Response.InnerXml.ToString());
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ValuationRules);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message); 
			}			

			if(ProcessOK)
			{
				ValuationSatisfied = Response.ValuationSatisfied;
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

	} //classTaskManagerClass

}
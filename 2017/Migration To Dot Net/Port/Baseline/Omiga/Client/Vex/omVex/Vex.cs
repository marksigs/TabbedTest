/// <summary>
/// Workfile:		omVex.cs
/// Copyright:		Copyright © 2006 Vertex FSD
/// 
/// Description:	
/// The Xit2 interface is an external two-way interface to a valuation bureau LMS via the Xit2 
/// software created by the ‘Valuation Exchange’ (Vex) Company. Details of the Applicant, property 
/// and valuation details are sent to LMS the panel valuer company via Xit2’s web service. LMS 
/// then instruct the valuer and send back the valuation report via Xit2 when the valuer has 
/// completed the valuation. The status of the valuation can be sent back to Omiga at any time 
/// along with updates to the inspection date or details of any problems holding up the report.
///
/// There are two Omiga objects dealing with the interface:
///  
/// •	An Omiga web service, (Omiga Vex Inbound ws) receiving incoming Valuation reports and 
/// 	incoming status messages. When incoming messages are received (status update or a 
/// 	valuation report) it validates the data, stores it in a transitionary table as an XML BLOB, 
/// 	sends a response to Vex and (if the Omiga reponse was “success”) calls the task manager 
/// 	component to generate a task in the workflow.
/// 
/// •	OmVexBO is the object that creates the initial valuation request to Xit2 and also 
/// 	deals with the creation of any subsequent outbound status messages.  It will call the 
/// 	Vex web service to interface valuation instructions or outbound status updates.  
/// 	It also has the methods to deal with incoming the status responses and the actual 
/// 	valuation report that comes back.
/// 
/// History:
/// 
/// Prog		Date		Description
/// PEdney		08/11/2006	Created - E2CR34 : Omiga : Xit2 interface for on-line valuations.
/// PEdney		29/11/2006	EP2_232 - Logging
/// PEdney		05/12/2006	EP2_321 - Environment element
/// PEdney		06/12/2006	EP2_280 - Unit test
/// PBuck		02/01/2007	EP2_527 - WP8 - Xit2 Interface - Amend Status message handling onhold, offhold, cancel (Removed)
/// PBuck		02/01/2007	EP2_528 - WP8 - Xit2 Interface - Error codes as spec'd to be added into OmVEx
/// PEdney		04/01/2007	EP2_449	- WP8 - Xit2 Interface - new SP to purge VexInboundMessage table
/// PEdney		25/01/2007	EP2_950 - Set the correct valuation status in RunVexStatusUpdate.
/// PEdney		31/01/2007	EP2_950 - Create failure task on Xit2 failure.
/// PEdney		31/01/2007	EP2_1029 - Various data mapping issues.
/// PEdney		01/02/2007	EP2_997 - Set the appointment date & time set by the valuer on the Xit2 appointment task note.
/// PEdney		01/02/2007	EP2_1166 - Handle Interface Response error for outbound Status updates
/// PEdney		01/02/2007	EP21059 - More error codes
/// PEdney		06/02/2007	EP2_1246 - Modification required for Gemini
/// PEdney		07/02/2007	EP2_1246 - Need to set the destination type
/// PEdney		07/02/2007	EP2_1067 - Create Valuer Instruction on successful Vex valuation instruction and update valuation instruction on status update = Appointment 
/// PEdney		27/02/2007	EP2_1547 - ValuerInstruction not updating on receiving the Valuation Report
/// PEdney		28/02/2007	EP2_1625 - ValuationStatus not set to Cancelled
/// PEdney		28/02/2007	EP2_1728 - Call HandleInterfaceResponse when valuation report received
/// PEdney		21/03/2007	EP2_1578 - Call TaskManager.ProcessValuationReport instead of TaskManager.RunRiskAssesmentRules.
/// PEdney		21/03/2007	EP2_2006 - Set DateReceived when in ProcessValuationReport.
/// PEdney		27/03/2007	EP2_2110 - Remove UniversalTime methods calls.
/// PEdney		04/04/2007	EP2_1578 - Valuation tasks no longer created as they are created in omValuationRules
/// PEdney		07/04/2007	EP2_2317 - Populate HOSTTEMPLATEID and TEMPLATEID from global parameter.
/// </summary>

using System;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Runtime.InteropServices;
using System.Reflection;

using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{

	#region declarations

	public enum Xit2ResultType
	{
		SuccessFailure = 0,
		SuccessInstruction = 1,
		SuccessStatus = 2,
		SuccessReport = 3,
		ErrorDatabase = -1,
		ErrorNetwork = -2,
		ErrorXML = -3,
		ErrorNoMessage = -4,
		ErrorNoSourceGuid = -5,
		ErrorNoVersionID = -6,
		ErrorNoXSD = -7
	}

	public enum StatusCode
	{
		Unknown = 0,
		Accept = 10,
		Appointment = 20,
		Submit = 30,
		ValCancel = 40,
		ValOnHold = 50,
		ValOffHold = 60,
		ValuationReceived = 70,
		MessageSent = 80
	}

	#endregion declarations

	/// <summary>
	/// OmVexBO is the object that creates the initial valuation request to Xit2 and also deals with 
	/// the creation of any subsequent outbound status messages.  It will call the Vex web service to 
	/// interface valuation instructions or outbound status updates.  It also has the methods to deal 
	///	with incoming the status responses and the actual valuation report that comes back.
	/// </summary>
	[ProgId("omVex.omVexBO")]
	[Guid("C04382B5-C5A2-485a-A4CE-667B5E098051")]
	[ComVisible(true)]
	public class omVexBOClass
	{

		private OmigaDataClass _OmigaData = new OmigaDataClass();
		private Xit2WSClass _Xit2 = new Xit2WSClass();
		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		#region public interface

		/// <summary>
		/// omVexBO constructor
		/// </summary>
		public omVexBOClass()
		{
		}
		
		/// <summary>
		/// This method collates the interface request data for an outgoing Valuation Instruction and calls the Vex web service.
		/// </summary>
		public bool RunVexValuationInstruction(string RequestXML, ref string ResponseXML)
		{			
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("RequestXML = " + RequestXML);
			bool ProcessOK = false;
			Xit2ResultType Xit2Result;
			XmlDocument VexInstruction = null;
			OmigaMessageClass Request = new OmigaMessageClass(RequestXML);
			OmigaMessageClass Response = new OmigaMessageClass();
			TaskManagerClass TaskManager = new TaskManagerClass();
			XmlNode InstructingSource;

			try
			{
				//	Generate Data
				ProcessOK = _OmigaData.GetVexInstruction(Request, ref VexInstruction);

				// Set the environment node
				try
				{
					GlobalParameter Environment = new GlobalParameter("Xit2Environment");
					InstructingSource = VexInstruction.SelectSingleNode("/Instruction/InsertInstruction/InstructingSource");					
					XmlHelper.SetNodeValue(VexInstruction, "/Instruction/InsertInstruction", "Environment", Environment.String, InstructingSource);
				}
				catch
				{
					_Logger.LogError(MethodName, "Unable to set the environment value");
				}

				// Validate Vex instruction against the schema 
				if(ProcessOK)
				{
					ProcessOK = ValidateXML(VexInstruction.InnerXml, "VexInstruction.xsd");
				}
				//PB 02/01/2007 EP2_528 Begin
				else
				{
					throw new OmigaException(8801); // Unable to generate data
				}
				//EP2_528 End

				//	Call Vex web service 
				if(ProcessOK)
				{
					_Logger.Debug("VexInstruction = " + VexInstruction.InnerXml);
					//PB 02/01/2007 EP2_528 Begin
					try
					{
						ProcessOK = _Xit2.SubmitIncomingInstructionMessage(VexInstruction, out Xit2Result);
					}
					catch(Exception NewException)
					{
						Request.Vex_ResponseCode = Xit2ResultType.SuccessFailure;
						Request.Vex_ValuationStatus = StatusCode.Unknown;
						Request.Vex_ValuationReason = NewException.Message;
						HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);
						throw new OmigaException(8800); // possible problems connecting to webservice
					}
					//EP2_528 End
					Request.Vex_ResponseCode = Xit2Result;
				}

				//	Process Response
				if(!ProcessOK)
				{
					Request.Vex_ResponseCode = Xit2ResultType.SuccessFailure;
					Request.Vex_ValuationStatus = StatusCode.Unknown;
				}
				ProcessOK = HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);

				// Create the valuation instruction
				if(ProcessOK)
				{
					ProcessOK = TaskManager.CreateValuerInstructions(Request, ref Response);
				}

				// Complete the task
				if(ProcessOK)
				{
					ProcessOK = TaskManager.CompleteTask(Request, ref Response);
				}

			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.Debug("ResponseXML = " + ResponseXML);
			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}

		/// <summary>
		///This method collates the interface request data for an outgoing Valuation Status update and calls the Vex web service.
		/// </summary>
		public bool RunVexStatusUpdate(string RequestXML, ref string ResponseXML)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("RequestXML = " + RequestXML);
			bool ProcessOK = false;
			string Status = "";
			XmlDocument VexStatusInstruction = null;
			OmigaMessageClass Request = new OmigaMessageClass(RequestXML);			
			OmigaMessageClass Response = new OmigaMessageClass();
			TaskManagerClass TaskManager = new TaskManagerClass();
			string TaskID = Request.CaseTask_TaskID;
			Xit2ResultType Result;
			XmlNode InstructingSource;			
			bool TaskComplete;
			GlobalParameter ApptTaskID;

			try
			{
				GlobalParameter TMVexValOnHold = new GlobalParameter("TMVexValOnHold");
				GlobalParameter TMVexValOffHold = new GlobalParameter("TMVexValOffHold");
				GlobalParameter TMVexValCancel = new GlobalParameter("TMVexValCancel");

				if(TaskID == TMVexValOnHold.String)
				{
					Status = "ONHOLD";
					Request.Vex_ValuationStatus = StatusCode.ValOnHold;				
				}

				else if (TaskID == TMVexValOffHold.String)
				{
					Status = "OFFHOLD";
					Request.Vex_ValuationStatus = StatusCode.ValOffHold;

					ApptTaskID = new GlobalParameter("TMVexValApptAwaited");
					Request.CaseTask_TaskID = ApptTaskID.String;
							
					ProcessOK = TaskManager.IsTaskComplete(Request, ref Response, out TaskComplete);

					if(ProcessOK && TaskComplete)
					{
						TaskManager.CreateTask(Request, ref Response);
					}

				}

				else if (TaskID == TMVexValCancel.String)
				{
					Status = "CANCEL";
					Request.Vex_ValuationStatus = StatusCode.ValCancel;

					ProcessOK = TaskManager.CancelValuerInstruction(Request, ref Response);
				}

				if(ProcessOK)
				{
					//	Generate Data
					ProcessOK = _OmigaData.GetVexStatus(Request, ref VexStatusInstruction);
				}


				if(ProcessOK)
				{
					// Set the status code
					XmlHelper.SetNodeValue(VexStatusInstruction, "/Instruction/Status", "StatusCode", Status, null);

					// Set the environment node
					try
					{
						GlobalParameter Environment = new GlobalParameter("Xit2Environment");
						InstructingSource = VexStatusInstruction.SelectSingleNode("/Instruction/InsertInstruction/InstructingSource");					
						XmlHelper.SetNodeValue(VexStatusInstruction, "/Instruction/InsertInstruction", "Environment", Environment.String, InstructingSource);
					}
					catch
					{
						_Logger.Debug("Unable to set the environment value");
					}
				}

				// Validate Vex instruction against the schema 
				if(ProcessOK)
				{
					ProcessOK = ValidateXML(VexStatusInstruction.InnerXml, "VexStatus.xsd");
				}

				//	Call Vex web service 
				if(ProcessOK)
				{
					_Logger.Debug("VexStatusInstruction = " + VexStatusInstruction.InnerXml);
					//PB 02/01/2007 EP2_528 Begin
					try
					{
						ProcessOK = _Xit2.SubmitIncomingStatusMessage(VexStatusInstruction, out Result);
					}
					catch(Exception NewException)
					{
						Request.Vex_ResponseCode = Xit2ResultType.SuccessFailure;
						Request.Vex_ValuationStatus = StatusCode.Unknown;
						Request.Vex_ValuationReason = NewException.Message;
						HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);
						throw new OmigaException(8800); // Error connecting to webservice
					}
					//EP2_528 End
					Request.Vex_ResponseCode = Result;
				}

				//	Process Response
				if(ProcessOK)
				{
					Request.Vex_ValuationStatus = StatusCode.MessageSent;
				}
				else
				{									
					Request.Vex_ResponseCode = Xit2ResultType.SuccessFailure;
					Request.Vex_ValuationStatus = StatusCode.Unknown;					
				}

				ProcessOK = HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);

				// Complete the task
				if(ProcessOK)
				{
					ProcessOK = TaskManager.CompleteTask(Request, ref Response);
				}

			}
			catch(Exception NewError) 
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.Debug("ResponseXML = " + ResponseXML);
			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}

		/// <summary>
		///	This private function is called to handle the webservice responses 
		///
		///	•	When the Vex response message from the Omiga web service is generated confirming 
		///		whether the incoming Valuation Report or incoming status message was successfully 
		///		interfaced, or providing the reason for a failure.		
		///
		///	•	When the Vex response message from the Vex web service is received confirming 
		///		whether the outgoing Valuation Instruction or status message was successfully 
		///		interfaced, or the reason for a failure.  This Vex web service response will 
		///		be a response code of:
		///
		///		This Omiga web service response as returned to Vex will be:
		///
		///		1 = SUCCESS (Valuation Instruction)
		///		2 = SUCCESS (Valuation Status incoming or outgoing Message) 
		///		3 = SUCCESS (Valuation Report)
		///
		///		- 1 = Database Error 
		///		- 2 = Network Error
		///		- 3 = XML Error
		///		- 4 = No Message
		///		- 5 = No Source GUID
		///		- 6 = No Version ID
		///		- 7 = NO XSD
		///	
		///	Tasks need to be configured in Supervisor where it is necessary to generate a task from the response
		///
		///	Once the response is received from the Vex web service, a task will attempt 
		///	to be created if configured in Supervisor as part of the interface tasks
		/// </summary>
		//public bool HandleInterfaceResponse(string ApplicationNumber, int ApplicationFactFindNumber, XmlDocument VexResponse)
		public bool HandleInterfaceResponse(string RequestXML, ref string ResponseXML)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("RequestXML = " + RequestXML);
			bool ProcessOK = false;

			try
			{	
				_Logger.Debug("RequestXML = " + RequestXML);		
				OmigaMessageClass Request = new OmigaMessageClass(RequestXML);
				OmigaMessageClass Response = new OmigaMessageClass();
				TaskManagerClass TaskManager = new TaskManagerClass();
						
				ProcessOK = TaskManager.HandleInterfaceResponse(ref Request, ref Response);
				ResponseXML = Response.InnerXml;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8804); // EP2_528 - Vex: Failed to process the initial response - the required tasks that indicate the response type have not been created
			}

			_Logger.Debug("ResponseXML = " + ResponseXML);
			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}

		/// <summary>
		///	This method is called from the Vex web site when an incoming status message or valuation report is available.  
		/// ???is it needed or can webservice decide which method to call – Process status message or Process valuation report?
		///
		///	The responses are defined in the VexInterfaceDefinition.doc.  The response from the Vex web service will
		///	be either an initial response or the valuation statuses or the valuation report.
		///
		///	Valuation Status
		///	The status that will be returned from the Vex web service at various points will be:
		///
		///	StatusCode
		///
		///		•	ACCEPT
		///		•	SUBMIT
		///		•	APPOINTMENT
		///		•	VALCANCEL
		///		•	VALONHOLD
		///		•	VALOFFHOLD
		///
		///	Valuation Report
		///	The Valuation report will be returned at a given point.
		/// </summary>
		public bool HandleInboundStatusUpdate(string RequestXML, ref string ResponseXML)
		{
			bool ProcessOK = false;

			//ProcessOK = ProcessVexStatusCode(RequestXML, ref ResponseXML);

			return ProcessOK;
		}

		/// <summary>
		/// Checks the Vex Incoming status code received from Vex during the Valuation processing period 
		/// and handles it accordingly.   This will be one of the following statuses:
		///		10 = ACCEPT
		///		20 = APPOINTMENT
		///		30 = SUBMIT 
		///		40 = VALCANCEL
		///		50 = VALONHOLD 
		///		60 = VALOFFHOLD
		///		
        ///	If Status message is Appointment, this method stores the Appt date and updates 
        ///	the relevant task for Awaiting appointment to completed.
		///
        ///	If Status is VALCANCEL, will need to update Valuation Status to cancelled.
		///
		///	This method also makes a call to the rules component to process codes and generate tasks
		/// </summary>
		public string ProcessVexStatusCode(string RequestXML, ref string ResponseXML)
		{
			Xit2ResultType Result = Xit2ResultType.ErrorXML;
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("RequestXML = " + RequestXML);
			bool ProcessOK = false;

			try
			{							
				OmigaMessageClass Request = new OmigaMessageClass(RequestXML);
				OmigaMessageClass Response = new OmigaMessageClass();			
				TaskManagerClass TaskManager = new TaskManagerClass();
				GlobalParameter TaskID;

				Request.DefaultUserCredentials();
				Request.Application_ApplicationFactFindNumber = "1";

				// Set the appointment date and time for the Xit2 appointment task note.
				if(Request.Vex_ValuationStatus == StatusCode.Appointment)
				{
					Request.Interface_TaskNotes = String.Concat("Appointment date: ", DateTime.Parse(Request.Vex_AppointmentDate).ToString());					
				}				

				ProcessOK = HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);

				if(ProcessOK)
				{

					switch(Request.Vex_ValuationStatus)
					{
						case StatusCode.Accept:

							// Store IncomingStatusUpdateXML\InstructionSystemID (Vex’s unique key) 
							// for future use (if needed when sending outgoing status updates)
							// Awaiting table structure from Michelle

							break;

						/* PB 02/01/2007 EP2_527 (Commented out)
						 case StatusCode.ValOnHold:
						
							// Set Task in GP TMVexValOnHold to complete or create and complete it
							TaskID = new GlobalParameter("TMVexValOnHold");
							Request.CaseTask_TaskID = TaskID.String;	
							ProcessOK = TaskManager.CompleteTask(Request, ref Response);
							break;

						case StatusCode.ValCancel:
						
							// Set Task in GP TMVexValCancel to complete or create and complete it
							TaskID = new GlobalParameter("TMVexValCancel");
							Request.CaseTask_TaskID = TaskID.String;
							ProcessOK = TaskManager.CompleteTask(Request, ref Response);
						
							// Call method to update the valuation instruction status to cancelled
							if(ProcessOK)
							{
								ProcessOK = TaskManager.CancelValuerInstruction(Request, ref Response);
							}
							break;

						case StatusCode.ValOffHold:
						
							// Set Task in GP TMVexValOffHold to complete or create and complete it
							TaskID = new GlobalParameter("TMVexValOffHold");
							Request.CaseTask_TaskID = TaskID.String;
							ProcessOK = TaskManager.CompleteTask(Request, ref Response);

							// If AwaitingAppointment task is already complete (task held in GP TMVexValApptAwaited).
							// Generate another as the next stage in the process will be for Vex to send an APPOINTMENT
							// message.
							if(ProcessOK) 
							{
								bool TaskComplete;

								TaskID = new GlobalParameter("TMVexValApptAwaited");
								Request.CaseTask_TaskID = TaskID.String;
							
								ProcessOK = TaskManager.IsTaskComplete(Request, ref Response, out TaskComplete);

								if(ProcessOK && TaskComplete)
								{
									TaskManager.CreateTask(Request, ref Response);
								}
							}
							break;
						*/

						case StatusCode.Appointment:
						
							// Create the ValuerInstruction record
														
							Request.ValuerInstruction_ValuationStatus = OmigaCombo.GetComboID("ValuationStatus", "A");
							ProcessOK = TaskManager.UpdateValuationInstruction(Request, ref Response);

							// Update the “Awaiting appointment” task to completed.
							if(ProcessOK)
							{
								TaskID = new GlobalParameter("TMVexValApptAwaited");
								Request.CaseTask_TaskID = TaskID.String;
								ProcessOK = TaskManager.CompleteTask(Request, ref Response);
							}
							break;

						case StatusCode.MessageSent:
							break;

						case StatusCode.Submit:
							break;

						case StatusCode.ValCancel:
							break;

						case StatusCode.ValOffHold:
							break;

						case StatusCode.ValOnHold:
							break;

						case StatusCode.ValuationReceived:						
							break;

						//PB EP2_528 02/01/2007 Begin
						default:

							// No valid response received
							throw new OmigaException(8803); //Vex: Unknown Response code received from the Vex Web Service

						//EP2_528 End
					}

				}
				else
				{
					ProcessOK = false;
					Result = Xit2ResultType.ErrorDatabase;
				}

				if(ProcessOK)
				{
					Result = Xit2ResultType.SuccessStatus;
				}
				else
				{
					Result = Xit2ResultType.ErrorDatabase;
				}

				// EP2_449 - Peter Edney - 04/01/2007
				if(ProcessOK)
				{
					Request.Vex_ResponseCode = Result;
					ProcessOK = _OmigaData.PurgeVexMessage(Request);
				}

			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.Debug("ResponseXML = " + ResponseXML);
			_Logger.LogFinsh(MethodName);
			return(Result.GetHashCode().ToString());
		}

		/// <summary>
		/// Saves the valuation report received from Vex and runs the rules against it and then 
		///	generates a valuation report returned task
		/// </summary>
		public string ProcessValuationReport(string RequestXML, ref string ResponseXML)
		{
			Xit2ResultType Result = Xit2ResultType.SuccessReport;
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("RequestXML = " + RequestXML);

			try
			{				

				TaskManagerClass TaskManager = new TaskManagerClass();
				OmigaMessageClass Request = new OmigaMessageClass(RequestXML);
				OmigaMessageClass Response = new OmigaMessageClass();
				GlobalParameter TaskID;
				bool ProcessOK = false;
				string ValuerInstructionNo;
				bool ValuationSatisfied;
			
				Request.DefaultUserCredentials();
				Request.Application_ApplicationFactFindNumber = "1";

				//TODO: Save the report


				// Step 1: Check for existing valuer instruction and set up structure ready for update
				// Check to see if an existing valuer instruction has been created since it maybe 
				// possible that a report can be received without an instruction, if so an 
				// instruction will firstly need to be created.
				// Create a valuation report and default instruction, the sequence number from 
				// the instruction needs to be set against the task created from global 
				// parameter TMVexValReportReturned  ????
				ProcessOK = TaskManager.CreateValuationReport(Request, ref Response, out ValuerInstructionNo);

				// Step 2: Now fill in the valuation report with details (save the details in 
				// IN ValuationReportXML to the valuation report tables)
				if(ProcessOK)
				{
					Request.CopyVexValuationNode();				
					Request.Valuation_Country = OmigaCombo.GetComboID("Country", "UK");
					Request.Valuation_TypeOfProperty = OmigaCombo.GetComboID("PropertyType", Request.Valuation_TypeOfProperty);
					Request.Valuation_PropertyDescription = OmigaCombo.GetComboID("PropertyDescription", Request.Valuation_PropertyDescription);
					Request.Valuation_Structure = OmigaCombo.GetComboID("BuildingConstruction", Request.Valuation_Structure);
					Request.Valuation_MainRoof = OmigaCombo.GetComboID("RoofConstruction", Request.Valuation_MainRoof);
					Request.Valuation_Tenure = OmigaCombo.GetComboID("PropertyTenure", Request.Valuation_Tenure);
					Request.Valuation_Drainage = OmigaCombo.GetComboID("PropertyDrainage", Request.Valuation_Drainage);
					Request.Valuation_Water = OmigaCombo.GetComboID("WaterSupply", Request.Valuation_Water);
					Request.Valuation_CertificationType = OmigaCombo.GetComboID("BuildingsCertificationType", Request.Valuation_CertificationType);
					Request.Valuation_OverallCondition = OmigaCombo.GetComboID("ValuationOverallCondition", Request.Valuation_OverallCondition);
					Request.Valuation_RentalDemand = OmigaCombo.GetComboIDByName("RentalDemand", Request.Valuation_RentalDemand);
					Request.Valuation_DemandInArea = OmigaCombo.GetComboIDByName("AreaDemand", Request.Valuation_DemandInArea);
					Request.Valuation_Saleability = OmigaCombo.GetComboID("ValuationSaleability", Request.Valuation_Saleability);
					Request.Valuation_InstructionSequenceNo = ValuerInstructionNo;					
					Request.Valuation_DateReceived = DateTime.Now.ToString();
					
					ProcessOK = TaskManager.UpdateValuationReport(Request, ref Response);				
				}

				if(ProcessOK)
				{
					ProcessOK = TaskManager.UpdateNewProperty(Request, ref Response);				
				}

				if(ProcessOK)
				{
					ProcessOK = TaskManager.UpdateContactDetails(Request, ref Response);				
				}

				if(ProcessOK) 
				{
					ProcessOK = TaskManager.CalculateLTV(Request, ref Response);
				}

				// The change of property is process allows for a change of property after at 
				// least one valuation has been returned and the existing property valuation is 
				// present.   This will set a parameter against the application called 
				// ‘Change Of Property’ and allows a task with an input process of CM010 to 
				// create and accept a new quote and enter a revised purchase price.  
				// The following allows for the fact that this valuation report may be a 
				// valuation for a change of property.

				// Step3: Update Valuer Status, and set awaiting valuation report and upload 
				// valuation report type tasks to complete and Call Valuation Rules.

				// Now the report has been received and created for the valuer instruction 
				// previously set up, the awaiting valuation report task needs to be completed 
				// if it exists and additionally the valuation status needs to be set to 
				// ‘completed’. Once this is completed the valuation rules can be called
				if(ProcessOK)
				{
					Request.ValuerInstruction_ValuationStatus = OmigaCombo.GetComboID("ValuationStatus", "C");				
					Request.ValuerInstruction_InstructionSequenceNo = ValuerInstructionNo;
					ProcessOK = TaskManager.UpdateValuationInstruction(Request, ref Response);
				}

				if(ProcessOK)
				{
					TaskID = new GlobalParameter("TMVexValReportAwaited");
					Request.CaseTask_TaskID = TaskID.String;
					Request.CaseTask_Context = ValuerInstructionNo;
					ProcessOK = TaskManager.CompleteTask(Request, ref Response);
				}

				if(ProcessOK)
				{
					TaskID = new GlobalParameter("TMVexValReportReturned");
					Request.CaseTask_TaskID = TaskID.String;
					Request.CaseTask_Context = ValuerInstructionNo;
					ProcessOK = TaskManager.CompleteTask(Request, ref Response);
				}

				if(ProcessOK)
				{
					ProcessOK = TaskManager.ProcessValuationReport(Request, ref Response, out ValuationSatisfied);					
				}

				if(!ProcessOK)
				{
					Result = Xit2ResultType.ErrorDatabase;
				}

				// EP2_449 - Peter Edney - 04/01/2007
				if(ProcessOK)
				{
					Request.Vex_ResponseCode = Result;
					ProcessOK = _OmigaData.PurgeVexMessage(Request);
				}

				// EP2_1728 - Peter Edney - 28/02/2007
				if(ProcessOK)
				{
					ProcessOK = HandleInterfaceResponse(Request.InnerXml, ref ResponseXML);
				}				

			}
			catch(Exception NewError)
			{				
				_Logger.LogError(MethodName, NewError.Message);
				throw new OmigaException(8806);
			}

			_Logger.Debug("ResponseXML = " + ResponseXML);
			_Logger.LogFinsh(MethodName);
			return(Result.GetHashCode().ToString());
		}
		
		public bool TransformVexMessage(string VexXML, string SchemaName, ref string OmigaXML)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{	
				OmigaMessageClass Request;

				//Validate Vex XML
				ProcessOK = ValidateXML(VexXML, SchemaName);

				//Transform Vex XML to Omiga format
				if(ProcessOK)
				{
					ProcessOK = TransformXML(VexXML, ref OmigaXML, "TransformVexMessage.xslt");				
					Request = new OmigaMessageClass(OmigaXML);
				
					Request.DefaultUserCredentials();
					Request.Application_ApplicationNumber = Request.Application_ApplicationNumber.Substring(0, Request.Application_ApplicationNumber.Length - 2); 
					Request.Application_ApplicationFactFindNumber = Request.Application_ApplicationFactFindNumber;
					OmigaXML = Request.InnerXml;
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		public bool SaveValuationPDF(string RequestString)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{
				TaskManagerClass TaskManager = new TaskManagerClass();
				OmigaMessageClass Request = new OmigaMessageClass(RequestString);
				OmigaMessageClass Response = new OmigaMessageClass();
				
				Request.PrintDocumentData_Compressed = "0";
				Request.PrintDocumentData_FileContentsType = "BIN.BASE64";
				Request.PrintDocumentData_FileContents = Request.Vex_DocumentContents;
				Request.PrintDocumentData_FileName = Request.Vex_DocumentName;
				Request.PrintDocumentData_DeliveryType = OmigaCombo.GetComboID("DocumentDeliveryType", "pdf");

				Request.ControlData_HostTemplateName = Request.Vex_DocumentName;
				Request.ControlData_HostTemplateDescription = "Xit2 Valuation Report";				
				Request.ControlData_TemplateGroupID = "20";
				Request.ControlData_GeminiPrintMode = "20";
				Request.ControlData_DeliveryType = OmigaCombo.GetComboID("DocumentDeliveryType", "pdf");
				Request.ControlData_DestinationType = "W";
				
				GlobalParameter TemplateIDParameter = new GlobalParameter("Xit2ValuationTemplateID");
				Request.ControlData_DPSDocumentID = TemplateIDParameter.String;
				Request.ControlData_DocumentID = TemplateIDParameter.String;

				Request.PrintData_ApplicationNumber = Request.Application_ApplicationNumber;

				ProcessOK = TaskManager.SaveValuationPDF(Request, ref Response);
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		#endregion public interface

		#region private interface

		private bool ValidateXML(string xmlIn, string SchemaName) 
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			try
			{				
				XmlTextReader xtr = new XmlTextReader(XmlHelper.GetSchema(SchemaName));					
				XmlValidatingReader xvr = new XmlValidatingReader(xmlIn,XmlNodeType.Document,null);
				xvr.Schemas.Add(null,xtr);
				xvr.ValidationType = ValidationType.Schema;
					
				while(xvr.Read());
				xvr.Close();

				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				throw new Exception(NewError.Message, NewError);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}
			
		private bool TransformXML(string XMLin, ref string XMLout, string XSLTFileName)
		{
			bool ProcessOK = false;
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);

			try
			{				
				XmlDocument SourceXMLDOM = new XmlDocument();					
				SourceXMLDOM.LoadXml(XMLin);

				StringWriter AppDataWriter = new StringWriter();

				XslTransform Transformer = new XslTransform();
				Transformer.Load(XmlHelper.GetXML(XSLTFileName));
				Transformer.Transform(SourceXMLDOM, null, AppDataWriter, null);

				XMLout = AppDataWriter.ToString();
				ProcessOK = true;
			}
			catch(Exception NewError)
			{
				ProcessOK = false;
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(ProcessOK);
		}

		#endregion private interface

	} //class omVexBOClass

}
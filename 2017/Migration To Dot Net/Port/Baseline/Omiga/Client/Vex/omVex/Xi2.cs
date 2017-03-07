/*
--------------------------------------------------------------------------------------------
Workfile:			Xit2.cs
Copyright:			Copyright ©2006 Vertex Financial Services

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PE		08/11/2006	EP2_24 - Created
PE		08/12/2006	EP2_345 - XML Error
PE		11/12/2006	EP2_405 - Configurable Web Service address
PE		12/12/2006	EP2_426 - Version ID Validation
PB		02/01/2007	EP2_528 - WP8 - Xit2 Interface - Error codes as spec'd to be added into OmVEx
PE		01/02/2007	EP21059 - More error codes
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using System.Reflection;

using Vertex.Fsd.Omiga.omVex.vexWebService;
using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Xit2WSClass
	{
		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		public bool SubmitIncomingInstructionMessage(XmlDocument VexInstruction, out Xit2ResultType Result)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			bool ProcessOK = false;

			string ErrorMessage = "";
			ValuationMessage MessageProxy;
			System.Net.NetworkCredential Credentials;

			GlobalParameter Xit2URL = new GlobalParameter("Xit2WebServiceURL");
			GlobalParameter UserName = new GlobalParameter("Xit2WebServiceUserName");
			GlobalParameter Password = new GlobalParameter("Xit2WebServicePassword");
			GlobalParameter SourceID = new GlobalParameter("Xit2WebServiceSourceID");
			GlobalParameter VersionID = new GlobalParameter("Xit2VersionIDInstruction");

			MessageProxy = new ValuationMessage();
			Credentials = new System.Net.NetworkCredential(UserName.String, Password.String);
			MessageProxy.PreAuthenticate = true;
			MessageProxy.Credentials = Credentials;
			MessageProxy.Url = Xit2URL.String;

			try
			{
				Result = (Xit2ResultType) int.Parse(MessageProxy.SubmitIncomingInstructionMessage(VexInstruction.OuterXml, (int) VersionID.Amount, SourceID.String));				
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				// PB 02/01/2007 EP2_528 begin
				//throw new OmigaException(8801);
				throw new OmigaException(8805); // Vex: Failed to process the incoming Vex status message – tasks configured against the status message responses have not been created
				// EP2_528 End
			}

			if(Result == Xit2ResultType.SuccessInstruction)
			{
				ProcessOK = true;
			}
			else
			{
				switch(Result)
				{
					case Xit2ResultType.ErrorDatabase:
						ErrorMessage = "Database error";
						break;
					case Xit2ResultType.ErrorNetwork:
						ErrorMessage = "Network error";
						break;
					case Xit2ResultType.ErrorXML:
						ErrorMessage = "XML error";
						break;
					case Xit2ResultType.ErrorNoSourceGuid:
						ErrorMessage = "No source guid";
						break;
					case Xit2ResultType.ErrorNoVersionID:
						ErrorMessage = "No version ID";
						break;
					case Xit2ResultType.ErrorNoXSD:
						ErrorMessage = "No XSD";
						break;		
					case Xit2ResultType.ErrorNoMessage:
						ErrorMessage = "No message";
						break;		
					default:
						ErrorMessage = "Undefined error";
						break;						
				}

				_Logger.LogError(MethodName, ErrorMessage);
				// PB 02/01/2007 EP2_528 begin
				//throw new OmigaException(8801);
				throw new OmigaException(8805); // Vex: Failed to process the incoming Vex status message – tasks configured against the status message responses have not been created
				// EP2_528 End
			}

			_Logger.LogFinsh(MethodName);
			return ProcessOK;			
		}
	
		public bool SubmitIncomingStatusMessage(XmlDocument VexStatus, out Xit2ResultType Result)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			const string UserName = "dbWSuser";
			const string Password = "db69mnSP62";
			const string SourceID = "DEUTSCHEBANK";
			const int VersionID = 1;

			bool ProcessOK = false;
			string ErrorMessage = "";
			ValuationMessage MessageProxy;
			System.Net.NetworkCredential Credentials;

			MessageProxy = new ValuationMessage();
			Credentials = new System.Net.NetworkCredential(UserName, Password);
			MessageProxy.PreAuthenticate = true;
			MessageProxy.Credentials = Credentials;

			try
			{
				Result = (Xit2ResultType) int.Parse(MessageProxy.SubmitIncomingStatusMessage(VexStatus.OuterXml, VersionID, SourceID));
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				// PB 02/01/2007 EP2_528 begin
				//throw new OmigaException(8801);
				throw new OmigaException(8805); // Vex: Failed to process the incoming Vex status message – tasks configured against the status message responses have not been created
				// EP2_528 End
			}

			if(Result == Xit2ResultType.SuccessInstruction || Result == Xit2ResultType.SuccessReport || Result == Xit2ResultType.SuccessStatus)
			{
				ProcessOK = true;
			}
			else
			{
				switch(Result)
				{
					case Xit2ResultType.ErrorDatabase:
						ErrorMessage = "Database error";
						break;
					case Xit2ResultType.ErrorNetwork:
						ErrorMessage = "Network error";
						break;
					case Xit2ResultType.ErrorXML:
						ErrorMessage = "XML error";
						break;
					case Xit2ResultType.ErrorNoSourceGuid:
						ErrorMessage = "No source guid";
						break;
					case Xit2ResultType.ErrorNoVersionID:
						ErrorMessage = "No version ID";
						break;
					case Xit2ResultType.ErrorNoXSD:
						ErrorMessage = "No XSD";
						break;						
					case Xit2ResultType.ErrorNoMessage:
						ErrorMessage = "No message";
						break;		
					default:
						ErrorMessage = "Undefined error";
						break;						
				}

				_Logger.LogError(MethodName, ErrorMessage);
				// PB 02/01/2007 EP2_528 Begin
				//throw new OmigaException(ErrorMessage);
				throw new OmigaException(8805); // Vex: Failed to process the incoming Vex status message – tasks configured against the status message responses have not been created
				// EP2_528 End
			}

			_Logger.LogFinsh(MethodName);
			return ProcessOK;
		}
	} //class Xit2Class

}

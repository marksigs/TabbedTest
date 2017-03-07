//--------------------------------------------------------------------------------------------
//Workfile:			LenderMessage.asmx.cs
//Copyright:			Copyright ©2006 Vertex Financial Services
//
//Description:		
//--------------------------------------------------------------------------------------------
//History:
//
//Prog	Date		Description
//PE	08/11/2006	EP2_24 - Created
//PE	29/11/2006	EP2_232 - Logging
//PE	04/12/2006	EP2_293 - Incorrect casing of web methods.
//PE	04/12/2006	EP2_297 - VersionID should be of type int.
//PE	04/12/2006	EP2_306 - Incorrectly cased web method names.
//PE	05/12/2006	EP2_311 - Web service should use default namespace (tempuri.org)
//PE	05/12/2006	EP2_318 - Reformat header
//PE	05/12/2006	EP2_314 - Incorrect VersionID
//PE	12/12/2006	EP2_426 - Version ID Validation
//PE	08/01/2007	EP2_713 - Change case of web method SubmitVExClientValuationMessage
//--------------------------------------------------------------------------------------------
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Xml.Schema;
using System.Reflection;
using System.Threading;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Messaging;

using Microsoft.Win32;

using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omVex;

namespace Vertex.Fsd.Omiga.Web.Services.Xit2Interface
{
	public delegate string VexCallBackDelegate(string RequestXML, ref string ResponseXML);

	public enum MessageType
	{
		Status,
		Report
	}

	/// <summary>
	/// The ProcessVexCallClass provides the callback functionality, allowing the
	/// web service to return a result before the process has completed.
	/// </summary>
	public class ProcessVexCallClass
	{

		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

		[OneWayAttribute()]
		public void ProcessResults(IAsyncResult Result)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);

			try
			{
				string Response = "";

				ManualResetEvent Reset = (ManualResetEvent )Result.AsyncState;

				// Extract the delegate from the AsyncResult  
				VexCallBackDelegate VexCallBack = (VexCallBackDelegate)((AsyncResult)Result).AsyncDelegate;

				// Obtain the result
				VexCallBack.EndInvoke(ref Response, Result);			
				_Logger.Debug("Response = " + Response);
				Reset.Set();
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
		}
	}

	public class LenderMessage : System.Web.Services.WebService
	{

		LoggingClass _Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());
		private omVexBOClass _Vex = new omVexBOClass();

		public LenderMessage()
		{
			//CODEGEN: This call is required by the ASP.NET Web Services Designer
			InitializeComponent();
		}

		#region Component Designer generated code
		
		//Required by the Web Services Designer 
		private IContainer components = null;
				
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if(disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);		
		}
		
		#endregion

		#region WebMethods

		[WebMethod]
		public string SubmitVExClientStatusMessage(string mESSAGE, int vERSIONID, string sOURCEGUID)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("Message = " + mESSAGE);
			_Logger.Debug("VersionID = " + vERSIONID);
			_Logger.Debug("SourceGUID = " + sOURCEGUID);

			string Result = "";

			try
			{				
				string Response = "";
				string OmigaRequest = "";			
				OmigaMessageClass Request = new OmigaMessageClass();
				omVex.OmigaDataClass OmigaData = new OmigaDataClass();
				VexCallBackDelegate VexCallBack;

				Result = ValidateCall(mESSAGE, vERSIONID, sOURCEGUID, MessageType.Status, ref OmigaRequest);			

				if(!OmigaData.StoreVexInboundMessage(mESSAGE, Result))
				{
					Result = Xit2ResultType.ErrorDatabase.GetHashCode().ToString();
				}
									
				if(Result == Xit2ResultType.SuccessStatus.GetHashCode().ToString())
				{					
					Request.LoadXml(OmigaRequest);
					Request.Vex_ResponseCode = (Xit2ResultType) int.Parse(Result);					

					// Initialise the delegate with the target method
					VexCallBack = new VexCallBackDelegate(_Vex.ProcessVexStatusCode);

					// Create an instance of the class which is going 
					// to be called when the call completes
					ProcessVexCallClass ProcessVexCall = new ProcessVexCallClass();

					// Define the AsyncCallback delegate
					AsyncCallback CallBack = new AsyncCallback(ProcessVexCall.ProcessResults);

					// Can stuff any object as the state object
					ManualResetEvent Reset = new ManualResetEvent(false);
					Object State = Reset;

					// Asynchronously invoke the Vex Call
					VexCallBack.BeginInvoke(Request.InnerXml, ref Response, CallBack, State);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				Result = Xit2ResultType.ErrorNetwork.GetHashCode().ToString();
			}

			_Logger.LogFinsh(MethodName);
			return(Result);
		}

		[WebMethod]
		public string SubmitVexClientValuationMessage(string mESSAGE, int vERSIONID, string sOURCEGUID)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);
			_Logger.Debug("Message = " + mESSAGE);
			_Logger.Debug("VersionID = " + vERSIONID);
			_Logger.Debug("SourceGUID = " + sOURCEGUID);

			string Result = "";
			bool ProcessOK = true;
			

			try
			{				
				string Response = "";
				string OmigaRequest = "";
				
				OmigaMessageClass Request = new OmigaMessageClass();		
				omVex.OmigaDataClass OmigaData = new OmigaDataClass();
				VexCallBackDelegate VexCallBack;

				Result = ValidateCall(mESSAGE, vERSIONID, sOURCEGUID, MessageType.Report, ref OmigaRequest);
				OmigaData.StoreVexInboundMessage(mESSAGE, Result);

				if(Result == Xit2ResultType.SuccessReport.GetHashCode().ToString())
				{
					ProcessOK = _Vex.SaveValuationPDF(OmigaRequest);
					if(!ProcessOK)
					{
						Result = Xit2ResultType.ErrorDatabase.GetHashCode().ToString();
					}
				}

				if(Result == Xit2ResultType.SuccessReport.GetHashCode().ToString())
				{
					Request.LoadXml(OmigaRequest);
					Request.Vex_ResponseCode = (Xit2ResultType) int.Parse(Result);		

					//Result = _Vex.ProcessValuationReport(Request.InnerXml,ref Response);
					
					// Initialise the delegate with the target method
					VexCallBack = new VexCallBackDelegate(_Vex.ProcessValuationReport);

					// Create an instance of the class which is going 
					// to be called when the call completes
					ProcessVexCallClass ProcessVexCall = new ProcessVexCallClass();

					// Define the AsyncCallback delegate
					AsyncCallback CallBack = new AsyncCallback(ProcessVexCall.ProcessResults);

					// Can stuff any object as the state object
					ManualResetEvent Reset = new ManualResetEvent(false);
					Object State = Reset;

					// Asynchronously invoke the Vex Call
					VexCallBack.BeginInvoke(Request.InnerXml, ref Response, CallBack, State);
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
				Result = Xit2ResultType.ErrorNetwork.GetHashCode().ToString();
			}
				
			_Logger.LogFinsh(MethodName);
			return(Result);		
		}

		#endregion WebMethods

		private string ValidateCall(string Message, int VersionID, string SourceGUID, MessageType CallType, ref string OmigaXML)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			_Logger.LogStart(MethodName);

			bool ProcessOK = true;
			string Result = "";

			try
			{
				if(Message == "")
				{
					Result = Xit2ResultType.ErrorNoMessage.GetHashCode().ToString();
					_Logger.LogError(MethodName, "Message is blank");
					ProcessOK = false;
				}

				if(ProcessOK)
				{

					GlobalParameter GPXit2VersionID;
					int Xit2VersionID = 0;

					if(CallType==MessageType.Status)
					{
						GPXit2VersionID = new GlobalParameter("Xit2VersionIDStatus");
						Xit2VersionID = (int) GPXit2VersionID.Amount;
					}

					if(CallType==MessageType.Report)
					{
						GPXit2VersionID = new GlobalParameter("Xit2VersionIDReport");
						Xit2VersionID = (int) GPXit2VersionID.Amount;
					}					

					if(VersionID != Xit2VersionID)
					{
						Result = Xit2ResultType.ErrorNoVersionID.GetHashCode().ToString();
						_Logger.LogError(MethodName, "Incorrect VersionID");
						ProcessOK = false;
					}
				}

				if(ProcessOK)
				{
					if(SourceGUID == "")
					{
						Result = Xit2ResultType.ErrorNoSourceGuid.GetHashCode().ToString();
						_Logger.LogError(MethodName, "SourceGUID is blank");
						ProcessOK = false;
					}
				}

				if(ProcessOK)
				{
					switch(CallType)
					{
						case MessageType.Status:
							if(!_Vex.TransformVexMessage(Message, "VexStatus.xsd", ref OmigaXML))
							{
								Result = Xit2ResultType.ErrorXML.GetHashCode().ToString();
								_Logger.LogError(MethodName, "Unable to transform status message.");
								ProcessOK = false;
							}
							else
							{
								Result = Xit2ResultType.SuccessStatus.GetHashCode().ToString();
							}
							break;
						case MessageType.Report:
							if(!_Vex.TransformVexMessage(Message, "VexReport.xsd", ref OmigaXML))
							{
								Result = Xit2ResultType.ErrorXML.GetHashCode().ToString();
								_Logger.LogError(MethodName, "Unable to transform report message.");
								ProcessOK = false;
							}
							else
							{
								Result = Xit2ResultType.SuccessReport.GetHashCode().ToString();
							}
							break;
					}				
				}
			}
			catch(Exception NewError)
			{
				_Logger.LogError(MethodName, NewError.Message);
			}

			_Logger.LogFinsh(MethodName);
			return(Result);
		}

	}
}

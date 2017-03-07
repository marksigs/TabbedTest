using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.Xml.Serialization; // for XmlSerializer
using System.IO;	// for reading file
using System.Xml; // for XMLReader
using System.Web.Services.Protocols; 

namespace Vertex.FSD.Omiga.Web.Services.SubmitStopAndSaveAiPWS
{
	/// <summary>
	/// Summary description for OmigaSubmitStopAndSaveAiPService.
	/// </summary>
	[WebService(Namespace="http://SubmitStopAndSaveAiP.IDUK.Omiga.marlboroughstirling.com")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://SubmitStopAndSaveAiP.IDUK.Omiga.marlboroughstirling.com",
		 Location="http://localhost/OmigaWebServices/Schemas/SubmitStopAndSaveAiP.wsdl"
		 )]
	public class OmigaSubmitStopAndSaveAiPService : Vertex.FSD.Omiga.Web.Services.SubmitStopAndSaveAiPWS.OmigaSubmitStopAndSaveAiP
	{
		public OmigaSubmitStopAndSaveAiPService()
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
		[XmlStreamSoapExtension]
		[WebMethodAttribute()]
		[SoapDocumentMethodAttribute( 
			 "http://Request.SubmitStopAndSaveAiP.IDUK.Omiga.marlboroughstirling.com", 
			 Use=System.Web.Services.Description.SoapBindingUse.Literal, 
			 ParameterStyle=SoapParameterStyle.Bare,
			 Binding="OmigaSoapRpcSoap")]
		[return: XmlElementAttribute( 
					 "RESPONSE", 
					 Namespace="http://Response.SubmitStopAndSaveAiP.IDUK.Omiga.marlboroughstirling.com")]
		public override RESPONSEType submitStopAndSaveAiP()
		{
			try
			{		
				RESPONSEType resp = new RESPONSEType();
				return resp;
			}
			catch(Exception ex)	
			{
				// create error response
				string msg = "Failed to deserialize response: \n";
				msg += GetExceptionMessages(ex);
				RESPONSEType errResp = new RESPONSEType();
				errResp.TYPE = RESPONSEAttribType.SYSERR;
				ERRORType errElem = new ERRORType();
				errElem.DESCRIPTION = msg;
				errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				errResp.ERROR = errElem;
				return errResp;
			}
			catch
			{
				// create error response
				RESPONSEType errResp = new RESPONSEType();
				errResp.TYPE = RESPONSEAttribType.SYSERR;
				ERRORType errElem = new ERRORType();
				errElem.DESCRIPTION = "Unknown error";
				errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				errResp.ERROR = errElem;
				return errResp;
			}
		}
		
		private string GetExceptionMessages(Exception ex)		
		{
			string msg = ex.Message.ToString();

			Exception exNew = ex.InnerException;
			while (null != exNew)
			{
				msg += "\n" + exNew.Message.ToString();
				exNew = exNew.InnerException;
			}
			return msg;
		}
	}
}

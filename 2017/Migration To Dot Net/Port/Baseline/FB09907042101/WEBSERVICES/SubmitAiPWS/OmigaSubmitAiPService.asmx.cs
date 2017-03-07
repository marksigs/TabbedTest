//
//------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/SubmitAiPWS/OmigaSubmitAiPService.asmx.cs $
//Workfile:			$Workfile: OmigaSubmitAiPService.asmx.cs $
//Current Version	$Revision: 2 $
//Last Modified		$Modtime: 26/10/06 16:13 $
//Modified By		$Author: Srao $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//IK	19/09/2006	EP2_1 remove Epsom identifiers from namespace directives
//IK	25/10/2006	associate with EP2_10
//--------------------------------------------------------------------------------------------
//
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

namespace Vertex.Fsd.Omiga.Web.Services.SubmitAiPWS
{
	/// <summary>
	/// Summary description for OmigaSubmitAiPService.
	/// </summary>
	[WebService(Namespace="http://SubmitAiP.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap", 
		 Namespace="http://SubmitAiP.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/SubmitAiP.wsdl"
		 )]
	public class OmigaSubmitAiPService : OmigaSubmitAiP
	{
		public OmigaSubmitAiPService()
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
		[System.Web.Services.WebMethodAttribute()]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.SubmitAiP.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.SubmitAiP.Omiga.vertex.co.uk")]
		public override RESPONSEType submitAiP([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.SubmitAiP.Omiga.vertex.co.uk")] REQUESTType REQUEST)
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
				errResp.ERRORTEXT = msg;
				//ERRORType errElem = new ERRORType();
				//errElem.DESCRIPTION = msg;
				//errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				//errResp.e = errElem;
				return errResp;
			}
			catch
			{
				// create error response
				RESPONSEType errResp = new RESPONSEType();
				errResp.ERRORTEXT= "Unknown error";
				//ERRORType errElem = new ERRORType();
				//errElem.DESCRIPTION = "Unknown error";
				//errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				//errResp.ERROR = errElem;
				return errResp;
			}
		}
		

		[XmlStreamSoapExtension]
		[System.Web.Services.WebMethodAttribute()]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.submitAiPAddressTargetResolver.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.SubmitAiP.Omiga.vertex.co.uk")]
		public override RESPONSEType submitAiPAddressTargetResolver([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.SubmitAiP.Omiga.vertex.co.uk")] atrREQUESTType ATR_REQUEST)
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
				errResp.ERRORTEXT = msg;
				//ERRORType errElem = new ERRORType();
				//errElem.DESCRIPTION = msg;
				//errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				//errResp.e = errElem;
				return errResp;
			}
			catch
			{
				// create error response
				RESPONSEType errResp = new RESPONSEType();
				errResp.ERRORTEXT= "Unknown error";
				//ERRORType errElem = new ERRORType();
				//errElem.DESCRIPTION = "Unknown error";
				//errElem.SOURCE = System.Reflection.MethodInfo.GetCurrentMethod().Name;
				//errResp.ERROR = errElem;
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

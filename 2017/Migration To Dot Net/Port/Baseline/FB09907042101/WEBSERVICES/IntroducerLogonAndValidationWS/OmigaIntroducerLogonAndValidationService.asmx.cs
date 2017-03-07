//
//------------------------------------------------------------------------------------------
//Archive			$Archive: /Epsom Phase2/2 INT Code/WebServices/IntroducerLogonAndValidationWS/OmigaIntroducerLogonAndValidationService.asmx.cs $
//Workfile:			$Workfile: OmigaIntroducerLogonAndValidationService.asmx.cs $
//Current Version	$Revision: 1 $
//Last Modified		$Modtime: 2/11/06 11:45 $
//Modified By		$Author: Dbarraclough $

//--------------------------------------------------------------------------------------------
//History:

//Prog	Date		Description
//IK	01/11/2006	created for EP2_22
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
using System.Web.Services.Protocols; 

namespace Vertex.Fsd.Omiga.Web.Services.IntroducerLogonAndValidationWS
{
	/// <summary>
	/// Summary description for Service1.
	/// </summary>
	[WebService(Namespace="http://IntroducerLogonAndValidation.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://IntroducerLogonAndValidation.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/Internal/IntroducerLogonAndValidation.wsdl"
		 )]
	public class OmigaIntroducerLogonAndValidationService : OmigaIntroducerLogonAndValidation
	{
		public OmigaIntroducerLogonAndValidationService()
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
		[WebMethodAttribute(Description ="Validates the introducer and returns a list of firms with which the introducer is associated")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
		public override GETFIRMSRESPONSEType getIntroducerUserFirms([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk")] GETFIRMSREQUESTType REQUEST)
		{
			GETFIRMSRESPONSEType resp = new GETFIRMSRESPONSEType();
			resp.TYPE = RESPONSEAttribType.SUCCESS;
			return resp;
		}

		[XmlStreamSoapExtension]
		[WebMethodAttribute(Description ="Authenticates the introducer against the firm and if successful returns the firms permissions ")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.AuthenticateFirm.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.AuthenticateFirm.Omiga.vertex.co.uk")]
		public override AUTHENTICATEFIRMRESPONSEType authenticateFirm([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.AuthenticateFirm.Omiga.vertex.co.uk")] AUTHENTICATEFIRMREQUESTType REQUEST)
		{
			AUTHENTICATEFIRMRESPONSEType resp = new AUTHENTICATEFIRMRESPONSEType();
			resp.TYPE = RESPONSEAttribType.SUCCESS;
			return resp;
		}


		[XmlStreamSoapExtension]
		[WebMethodAttribute(Description ="Validates a broker and if successful returns the firms the broker is associated with together with their permissions")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.ValidateBroker.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
		public override VALIDATEBROKERRESPONSEType validateBroker([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.ValidateBroker.Omiga.vertex.co.uk")] VALIDATEBROKERREQUESTType REQUEST)
		{			
			try
			{
				VALIDATEBROKERRESPONSEType resp = new VALIDATEBROKERRESPONSEType();
				resp.TYPE = RESPONSEAttribType.SUCCESS;
				return resp;
			}
			catch(Exception ex)	
			{
				// create error response
				string msg = "Failed to deserialize response: \n";
				msg += GetExceptionMessages(ex);
				VALIDATEBROKERRESPONSEType errResp = new VALIDATEBROKERRESPONSEType();
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
				VALIDATEBROKERRESPONSEType errResp = new VALIDATEBROKERRESPONSEType();
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

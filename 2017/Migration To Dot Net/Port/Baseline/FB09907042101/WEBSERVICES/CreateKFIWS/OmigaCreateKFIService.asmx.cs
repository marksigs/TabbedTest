//
//--------------------------------------------------------------------------------------------
//Workfile:			OmigaCreateKFIService.asmx.cs
//Copyright:			Copyright © 2005 Marlborough Stirling
//
//Description:		Create KFI web service
//--------------------------------------------------------------------------------------------
//History:
//
//Prog	Date		Description
//GHun	25/10/2005	MAR301 Created
//SAB	20/09/2006	EP2_1 - Updated the namespaces.
//PSC	07/11/2006	EP2_41 - Added new methods
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

namespace Vertex.Fsd.Omiga.Web.Services.CreateKFIWS
{
	/// <summary>
	/// Create KFI Web Service.
	/// </summary>
	
	[WebService(Namespace="http://CreateKFIWS.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://CreateKFIWS.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/CreateKFI.wsdl"
		 )]
	public class OmigaCreateKFIService : OmigaCreateKFI	//System.Web.Services.WebService
	{

		public OmigaCreateKFIService()
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
			 "http://Request.CreateKFI.Omiga.vertex.co.uk", 
			 Use=System.Web.Services.Description.SoapBindingUse.Literal, 
			 ParameterStyle=SoapParameterStyle.Bare,
			 Binding="OmigaSoapRpcSoap")]
		[return: XmlElementAttribute( 
					 "RESPONSE", 
					 Namespace="http://Response.CreateKFI.Omiga.vertex.co.uk")]
		public override RESPONSEType createKFI([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.CreateKFI.Omiga.vertex.co.uk")]REQUESTType REQUEST)
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

		[XmlStreamSoapExtension]
		[WebMethodAttribute(Description ="Calculates the fees that are chargeable for a KFI")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.CalculateKFIFees.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.CalculateKFIFees.Omiga.vertex.co.uk")]
		public override CALCKFIFEESRESPONSEType calculateKFIFees([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.CalculateKFIFees.Omiga.vertex.co.uk")]CALCKFIFEESREQUESTType REQUEST)
		{
			try
			{		
				CALCKFIFEESRESPONSEType resp = new CALCKFIFEESRESPONSEType();
				return resp;
			}
			catch(Exception ex)	
			{
				// create error response
				string msg = "Failed to deserialize response: \n";
				msg += GetExceptionMessages(ex);
				CALCKFIFEESRESPONSEType errResp = new CALCKFIFEESRESPONSEType();
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
				CALCKFIFEESRESPONSEType errResp = new CALCKFIFEESRESPONSEType();
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

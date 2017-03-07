//------------------------------------------------------------------------------
// SR   19/09/2006      EP2_1 : Modified namespaces for Epsom
// SR   22/09/2006      EP2_1 : use 'Fsd' instead of 'Fsd'  
// SR   23/10/2006      EP2_1 : Modified to check in
//------------------------------------------------------------------------------
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

namespace Vertex.Fsd.Omiga.Web.Services.GetComboListWS
{
	/// <summary>
	/// Summary description for OmigaGetComboListService.
	/// </summary>
	[WebService(Namespace="http://GetComboList.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://GetComboList.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/GetComboList.wsdl"
		 )]
	public class OmigaGetComboListService : Vertex.Fsd.Omiga.Web.Services.GetComboListWS.OmigaGetComboList
	{
		public OmigaGetComboListService()
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
			 "http://Request.GetComboList.Omiga.vertex.co.uk", 
			 Use=System.Web.Services.Description.SoapBindingUse.Literal, 
			 ParameterStyle=SoapParameterStyle.Bare,
			 Binding="OmigaSoapRpcSoap")]
		[return: XmlElementAttribute( 
					 "RESPONSE", 
					 Namespace="http://Response.GetComboList.Omiga.vertex.co.uk")]
		public override RESPONSEType getComboList()
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

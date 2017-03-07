using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
//	testbed only
using System.Xml;				
using System.Xml.Serialization;

namespace Vertex.Fsd.Omiga.Web.Services.OmigaFlexiCalcWS
{
	/// <summary>
	/// Summary description for OmigaFlexiCalcService.
	/// </summary>
	[WebService(Namespace="http://OmigaFlexiCalc.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://OmigaFlexiCalc.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/OmigaFlexiCalc.wsdl"
		 )]
	public class OmigaFlexiCalcService : Vertex.Fsd.Omiga.Web.Services.OmigaFlexiCalcWS.OmigaFlexiCalc
	{
		public OmigaFlexiCalcService()
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
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.FlexiCalc.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("OmigaFlexiCalcResponse", Namespace="http://Response.FlexiCalc.Omiga.vertex.co.uk")]
		public override OmigaFlexiCalcResponseType omigaFlexiCalc([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.FlexiCalc.Omiga.vertex.co.uk")] OmigaFlexiCalcRequestType OmigaFlexiCalcRequest)
		{
			OmigaFlexiCalcResponseType resp = new OmigaFlexiCalcResponseType();
			return resp;

		}
	}
}

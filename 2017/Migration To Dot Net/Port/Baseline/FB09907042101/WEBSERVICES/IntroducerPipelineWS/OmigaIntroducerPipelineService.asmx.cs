using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;

using omLogging = Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.Web.Services.IntroducerPipelineWS
{
	/// <summary>
	/// Summary description for OmigaIntroducerPipelineService.
	/// </summary>
	[WebService(Namespace="http://IntroducerPipeline.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://IntroducerPipeline.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/Internal/IntroducerPipeline.wsdl"
		 )]
	public class OmigaIntroducerPipelineService : OmigaIntroducerPipeline
	{
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public OmigaIntroducerPipelineService()
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

		// WEB SERVICE EXAMPLE
		// The HelloWorld() example service returns the string Hello World
		// To build, uncomment the following lines then save and build the project
		// To test this web service, press F5
		[XmlStreamSoapExtension]
		[WebMethodAttribute(Description ="Returns a list of pipeline cases for an introducer")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
		public override GETINTRODUCERPIPELINERESPONSEType getIntroducerPipeline([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk")]GETINTRODUCERPIPELINEREQUESTType REQUEST)
		{
			string functionName = System.Reflection.MethodInfo.GetCurrentMethod().Name; 
			string userId = string.Empty;
			string introducerId = string.Empty;

			if (REQUEST.USERID != null)
			{
				userId = REQUEST.USERID;
			}

			if (REQUEST.PIPELINESEARCH != null && REQUEST.PIPELINESEARCH.INTRODUCERID != null)
			{
				introducerId = REQUEST.PIPELINESEARCH.INTRODUCERID;
			}
				
			omLogging.ThreadContext.Properties["UserId"] = userId;
			omLogging.ThreadContext.Properties["IntroducerId"] = introducerId;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);

			}
			try
			{
				GETINTRODUCERPIPELINERESPONSEType resp = new GETINTRODUCERPIPELINERESPONSEType();
				resp.TYPE = RESPONSEAttribType.SUCCESS;
				return resp;
			}
			catch(Exception ex)	
			{
				// create error response
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": An error occurred deserializing the response.", ex);
				}

				string msg = "Failed to deserialize response: \n";
				msg += GetExceptionMessages(ex);
				GETINTRODUCERPIPELINERESPONSEType errResp = new GETINTRODUCERPIPELINERESPONSEType();
				errResp.TYPE = RESPONSEAttribType.SYSERR;
				ERRORType errElem = new ERRORType();
				errElem.DESCRIPTION = msg;
				errElem.SOURCE = functionName;
				errResp.ERROR = errElem;

				return errResp;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);

				}
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

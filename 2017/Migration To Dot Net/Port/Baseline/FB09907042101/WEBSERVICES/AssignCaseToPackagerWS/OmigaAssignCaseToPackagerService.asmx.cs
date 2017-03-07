using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using omLogging = Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.Web.Services.AssignCaseToPackagerWS
{
	/// <summary>
	/// Summary description for Service1.
	/// </summary>
	[WebService(Namespace="http://AssignCaseToPackager.Omiga.vertex.co.uk")]
	[WebServiceBinding(
		 Name="OmigaSoapRpcSoap",
		 Namespace="http://AssignCaseToPackager.Omiga.vertex.co.uk",
		 Location="http://localhost/OmigaWebServices/Schemas/Internal/AssignCaseToPackager.wsdl"
		 )]
	public class OmigaAssignCaseToPackagerService : OmigaAssignCaseToPackager
	{
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public OmigaAssignCaseToPackagerService()
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
		[WebMethodAttribute(Description ="Assigns a case to a packager")]
		[System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.AssignCaseToPackager.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
		[return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.AssignCaseToPackager.Omiga.vertex.co.uk")]
		public override ASSIGNCASETOPACKAGERRESPONSEType assignCaseToPackager([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.AssignCaseToPackager.Omiga.vertex.co.uk")]ASSIGNCASETOPACKAGERREQUESTType REQUEST)
		{
			string functionName = System.Reflection.MethodInfo.GetCurrentMethod().Name;
			string applicationNumber = string.Empty;

			if (REQUEST.ASSIGNCASE != null && REQUEST.ASSIGNCASE.APPLICATIONNUMBER != null)
			{
				applicationNumber = REQUEST.ASSIGNCASE.APPLICATIONNUMBER;
			}
			
			omLogging.ThreadContext.Properties["ApplicationNumber"] = applicationNumber;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);

			}

			try
			{
				ASSIGNCASETOPACKAGERRESPONSEType resp = new ASSIGNCASETOPACKAGERRESPONSEType();
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
				ASSIGNCASETOPACKAGERRESPONSEType errResp = new ASSIGNCASETOPACKAGERRESPONSEType();
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

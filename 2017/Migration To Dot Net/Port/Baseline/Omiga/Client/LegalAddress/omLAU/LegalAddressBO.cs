/*
--------------------------------------------------------------------------------------------
Workfile:			LegalAddressBO.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		The Legal Address interface passes the other system number (mortgage 
					account number) to the client Web to trigger change of address 
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		28/10/2005	MAR340 Created
HMA     01/12/2005  MAR750 Updated
GHun	13/12/2005	MAR852 set Operator on generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
HMA     13/12/2006  MAR1122 Remove hard coded Mortgage Account Number. Remove unused code.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using System.Runtime.InteropServices;
//COM
//NET Ref
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.Core;
//web Ref
using Vertex.Fsd.Omiga.omLAU.ChangeLegalAddress;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.omLAU
{
	[Guid("AE5CF898-6E03-44e2-8306-F4C95C4BB0FE")]
	[ComVisible(true)]
	[ProgId("omLAU.LegalAddressBO")]
	public class LegalAddressBO  //: ILegalAddressBO
	{
		private string _otherSystemNumber = string.Empty;

		DirectGatewayBO dg = new DirectGatewayBO();
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		
		public LegalAddressBO()
		{
		}
		
		// interface method to pass delayed completion date information to Profile
		public string RunLegalAddressTriggerInterface(string strRequest)
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST>
			//     <APPLICATION APPLICATIONNUMBER... APPLICATIONFACTFINDNUMBER.../>
			// </REQUEST>
			string functionName = "RunLegalAddressTriggerInterface";
			omLogging.ThreadContext.Properties["ApplicationNumber"] = "Unknown";

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + functionName + ". Request: " + strRequest); 
			}

			string respRun = String.Empty;
			XmlElement respRunXml=null;

			ChangeLegalAddressResponseType resp;
			DirectGatewaySoapService service;
			ChangeLegalAddressRequestType req;
			XmlDocument docVBReq = new XmlDocument();

			try
			{
				resp = new ChangeLegalAddressResponseType();
			}
			catch
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Error(functionName + "ERROR: CANNOT allocate Response Node");
				}

				throw new OmigaException(functionName + ": ERROR: CANNOT allocate Response Node: ");
			}
			try
			{
				service = new DirectGatewaySoapService();

				service.Url = dg.GetServiceUrl(); 
				service.Proxy = dg.GetProxy();

				req = new ChangeLegalAddressRequestType();

				docVBReq.LoadXml(strRequest);

				// Get the UserID from the request
				XmlElement RequestElement = (XmlElement) docVBReq.SelectSingleNode("REQUEST");
				string userID = RequestElement.GetAttribute("USERID");   
				string password = RequestElement.GetAttribute("PASSWORD");  

				//extract parameter
				_otherSystemNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@OTHERSYSTEMACCOUNTNUMBER").Value;

				//MAR1122 Get the Application Number from the request
				string strApplicationNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@APPLICATIONNUMBER").Value;
				omLogging.ThreadContext.Properties["ApplicationNumber"] = strApplicationNumber;

				// Get common items from omDG
				req.ClientDevice = dg.GetClientDevice();
				req.ServiceName = dg.GetServiceName();

				if (userID.Length > 0 && password.Length > 0)
				{
					req.TellerID = userID;
					req.TellerPwd = password;
				}
				
				req.ProxyID = dg.GetProxyId();
				req.ProxyPwd = dg.GetProxyPwd();

				req.ProductType = dg.GetProductType();
				req.CommunicationChannel = dg.GetCommunicationChannel((userID.Length > 0));

				req.Operator = dg.GetOperator();	//MAR852 GHun
				req.SessionID = "";
				req.CommunicationDirection = "";
				req.CustomerNumber = "NA";

				// add info from VBRequest
				//MAR1122 Remove hard coding : omTmBO checks that OtherSystemAccountNumber is present.
				//if (_otherSystemNumber.ToString().Length==0) _otherSystemNumber="80002587";
				req.MortgageAccountNumber = _otherSystemNumber;

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(ChangeLegalAddressRequestType), req);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				// call service
				resp = service.execute(req);
				
				if(_logger.IsDebugEnabled) 
				{
					string responseString = dg.GetXmlFromGatewayObject(typeof(ChangeLegalAddressResponseType), resp);
					_logger.Debug(functionName + ": Service Response: " + responseString );
				}
			}	
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Error(functionName + ": SOAP exception occurred: " + ex.Message.ToString(), ex);
				}

				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "SOAP exception occurred: " + ex.Message.ToString();
			}
			catch(Exception ex)
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Error(functionName + ": Exception occurred: " + ex.Message.ToString(), ex);
				}

				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Exception occurred: " + ex.Message.ToString();
			}
			catch
			{
				if(_logger.IsErrorEnabled) 
				{
					_logger.Error(functionName + ": Unknown exception occurred");
				}

				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Unknown exception occurred";
			} 
		
			//update output
			if (resp.ErrorCode == "0")
			{
				respRunXml = BuildResponse();
				respRunXml.SetAttribute("TYPE", "SUCCESS"); 
				respRun = respRunXml.OuterXml;
			}
			else
			{	
				respRunXml = BuildErrorResponse();
				if (resp.ErrorCode == "SYSERR" || resp.ErrorCode == "APPERR")
				{	//comes from exeption
					respRunXml.SetAttribute("TYPE", resp.ErrorCode); 
					respRunXml.SelectSingleNode("//NUMBER").InnerText="-1";
				}
				else
				{	//comes from exeption
					respRunXml.SetAttribute("TYPE","APPERR"); 
					respRunXml.SelectSingleNode("//NUMBER").InnerText= resp.ErrorCode;
				}
				respRunXml.SelectSingleNode("//DESCRIPTION").InnerText=resp.ErrorMessage;
				respRunXml.SelectSingleNode("//SOURCE").InnerText="RunLegalAddressTriggerInterface";
				respRun = respRunXml.OuterXml;
			}
			req = null;
			resp = null;
			docVBReq = null;
			respRunXml = null;
			return respRun;
		}

		private XmlElement BuildResponse()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement resp = doc.CreateElement ("RESPONSE");
			resp.SetAttribute("TYPE", "SUCCESS");
			return resp;
		}

		private XmlElement BuildErrorResponse()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement resp = doc.CreateElement ("RESPONSE");
			resp.SetAttribute("TYPE","");
			XmlNode errNode = doc.CreateElement("ERROR");
			errNode.AppendChild(doc.CreateElement("NUMBER"));
			errNode.AppendChild(doc.CreateElement("SOURCE"));
			errNode.AppendChild(doc.CreateElement("DESCRIPTION"));
			resp.AppendChild(errNode);
			return resp;
		}	
	}
}

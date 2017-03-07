/*
--------------------------------------------------------------------------------------------
Workfile:			UniservBO.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		Web service client to call Uniserve to validation addresses (PAFSearch)
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
TLiu	05/09/2005	MAR38 Created
GHun	08/10/2005	MAR137 Integrated with omDG and removed COM interop
MV		27/10/2005	MAR278 Included logging 
GHun	17/11/2005	MAR330 Only treat error code 1 as a success, if addresses are returned
GHun	12/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using System.Xml.Serialization;
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.omUniserv.WebRefValidateAddress;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
namespace Vertex.Fsd.Omiga.omUniserv
{
	public class UniservBO
	{
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public UniservBO()
		{
		}

		public string FindPAFAddress(string strRequest)
		{
			string functionName = "FindPAFAddress";
			DirectGatewaySoapService service = new DirectGatewaySoapService();
			ValidateAddressRequestType req = new ValidateAddressRequestType();
			ValidateAddressResponseType resp = new ValidateAddressResponseType();
			DirectGatewayBO dg = new DirectGatewayBO();

			try
			{
				//MAR137 GHun
				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();
				//MAR137 End

				XmlDocument doc = new XmlDocument();
				doc.LoadXml(strRequest);

				XmlNodeReader nodereader = new XmlNodeReader(doc.SelectSingleNode("ValidateAddressRequestType"));

				XmlSerializer serOut = new XmlSerializer(typeof(ValidateAddressRequestType));
				req = (ValidateAddressRequestType)(serOut.Deserialize(nodereader));

				//MAR137 GHun
				req.ClientDevice = dg.GetClientDevice();
				req.ServiceName = dg.GetServiceName();

				XmlElement requestElement = (XmlElement) doc.SelectSingleNode("REQUEST");
				string userId = (requestElement != null) ? requestElement.GetAttribute("USERID") : string.Empty;
				string password = (requestElement != null) ? requestElement.GetAttribute("PASSWORD") : string.Empty;

				if (userId.Length > 0 && password.Length > 0)
				{
					req.TellerID = userId;
					req.TellerPwd = password;
				}

				req.ProxyID = dg.GetProxyId();
				req.ProxyPwd = dg.GetProxyPwd();

				req.Operator = dg.GetOperator();	//MAR852 GHun

				req.ProductType = dg.GetProductType();
				req.CommunicationChannel = dg.GetCommunicationChannel((userId.Length > 0));
				req.CustomerNumber = "NA";
				//MAR137 End
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(ValidateAddressRequestType), req);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				resp = service.execute(req);

				if(_logger.IsDebugEnabled) 
				{
					string responseString  = dg.GetXmlFromGatewayObject(typeof(ValidateAddressResponseType), resp);
					_logger.Debug(functionName + ": Service Response: " + responseString);
				}

			}
			catch(System.Web.Services.Protocols.SoapException ex)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "SOAP exception occurred: " + ex.Code.ToString();
			}
			catch(Exception ex)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Exception occurred: " + ex.Message.ToString();
			}
			catch
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Unknown exception occurred";
			} 

			//MAR137 GHun
			string strResponse = dg.GetXmlFromGatewayObject(typeof(ValidateAddressResponseType), resp);

			string strType = resp.ErrorCode;
			//MAR330 GHun Only treat error code 1 as a success, if addresses are returned
			if (strType == "0" || ((strType == "1") && (strResponse.IndexOf("<Address", 1) > 0)))
				strType = "SUCCESS";
			else 
			{
				strType = "APPERR";
				strResponse = strResponse.Replace("<ErrorCode", "<ERROR><NUMBER");
				strResponse = strResponse.Replace("</ErrorCode>", "</NUMBER>");
				strResponse = strResponse.Replace("<ErrorMessage", "<DESCRIPTION");
				strResponse = strResponse.Replace("</ErrorMessage>", "</DESCRIPTION>");
				strResponse = strResponse.Replace("</ValidateAddressResponseType>", "<SOURCE>omUniserv.ValidateAddress</SOURCE><VERSION>1.0.0.0</VERSION></ERROR></ValidateAddressResponseType>");
			}

			strResponse = strResponse.Replace("<ValidateAddressResponseType", "<RESPONSE TYPE=\"" + strType + "\"");
			strResponse = strResponse.Replace("</ValidateAddressResponseType>", "</RESPONSE>");

			return strResponse;
		}
	}
}

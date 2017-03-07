/*
--------------------------------------------------------------------------------------------
Workfile:			omCardPayment.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HMA		02/11/2005	MAR397 Included logging 
MV		03/11/2005	MAR397 Fixed the reference
MV		03/11/2005	MAR397 Typo Error
HMA     01/12/2005  MAR750 Correct logging
GHun	13/12/2005	MAR852 set Operator in generic request
JD		10/05/2006	MAR1754 set CustomerNumber to passed in CIF number
--------------------------------------------------------------------------------------------
*/
using System;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.omCardPayment.WebTakeCardPayment;
using System.Xml;
using System.Xml.Serialization; // for XmlSerializer
using Vertex.Fsd.Omiga.omDg;
using log4net;

namespace Vertex.Fsd.Omiga.omCardPayment
{
//	[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
//	public interface ICardPaymentBO
//	{
//		[DispId(600)]string RunTakeCardPayment(string strRequest);
//	}

	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[ProgId("omCardPayment.CardPaymentBO")]
	[Guid("D23D38D2-61D2-4778-88D2-B279A533F215")]
	[ComVisible(true)]
	public class CardPaymentBO 
	{
		DirectGatewayBO dg = new DirectGatewayBO();
		private static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public string RunTakeCardPayment(string strRequest)
		{
			string functionName = "RunTakeCardPayment";

			TakeCardPaymentResponseType resp;
			DirectGatewaySoapService service;
			TakeCardPaymentRequestType req;

			try
			{
				resp = new TakeCardPaymentResponseType();
			}
			catch
			{
				return "ERROR: CANNOT allocate Response Node!";
			}

			try
			{
				service = new DirectGatewaySoapService();

				service.Url = dg.GetServiceUrl();
				service.Proxy = dg.GetProxy();

				req = new TakeCardPaymentRequestType();

				XmlDocument doc = new XmlDocument();
				
				doc.LoadXml(strRequest);

				// Get the UserID from the request
				XmlElement RequestElement = (XmlElement) doc.SelectSingleNode("REQUEST");
				string userID = RequestElement.GetAttribute("USERID");   
				string password = RequestElement.GetAttribute("PASSWORD");

				//MAR1754 get the CIF number from the request
				XmlNode CustomerElement = RequestElement.SelectSingleNode("CUSTOMER/CUSTOMERNUMBER");
				string sCIFNumber = CustomerElement.InnerText;

				XmlNodeReader nodereader = new XmlNodeReader(doc.SelectSingleNode(".//TakeCardPaymentRequestType"));
				XmlSerializer serOut = new XmlSerializer(typeof(TakeCardPaymentRequestType));
				req = (TakeCardPaymentRequestType)(serOut.Deserialize(nodereader));


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
				//req.CustomerNumber = "NA";
				req.CustomerNumber = sCIFNumber; //MAR1754

				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(TakeCardPaymentRequestType), req);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				resp = service.execute(req);
				
				if(_logger.IsDebugEnabled) 
				{
					string responseString = dg.GetXmlFromGatewayObject(typeof(TakeCardPaymentResponseType), resp);
					_logger.Debug(functionName + ": Service Response: " + responseString );
				}
			}

			catch(System.Web.Services.Protocols.SoapException ex)
			{
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "SOAP exception occurred: " + ex.Message.ToString();
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

			string strResponse = null;

			// MAR397 Use omDG function to get XML response
			strResponse = dg.GetXmlFromGatewayObject(typeof(TakeCardPaymentResponseType), resp);

			string strType = resp.ErrorCode;
			if (strType == "0")
				strType = "SUCCESS";
			else
			{
				strType = "ERROR";
				strResponse = strResponse.Replace("<ErrorCode", "<ERROR><NUMBER");
				strResponse = strResponse.Replace("</ErrorCode>", "</NUMBER>");
				strResponse = strResponse.Replace("<ErrorMessage", "<DESCRIPTION");
				strResponse = strResponse.Replace("</ErrorMessage>", "</DESCRIPTION>");
				strResponse = strResponse.Replace("</TakeCardPaymentResponseType>", "<SOURCE>omTakeCardPayment.TakeCardPayment</SOURCE><VERSION>1.0.0.0</VERSION></ERROR></TakeCardPaymentResponseType>");
			}

			strResponse = strResponse.Replace("<TakeCardPaymentResponseType", "<RESPONSE TYPE=\"" + strType + "\"");
			strResponse = strResponse.Replace("</TakeCardPaymentResponseType>", "</RESPONSE>");

			return strResponse;
		}
	}
}

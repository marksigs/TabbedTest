/*
--------------------------------------------------------------------------------------------
Workfile:			PinMailerBO.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		The Pin Mailer interface passes the other system customer numbers 
					to the client Web service
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		28/10/2005	MAR340 Created
HMA     30/11/2005  MAR750  Updated
HMA     06/12/2005  MAR750  Further changes
GHun	13/12/2005	MAR852	set Operator on generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
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
using Vertex.Fsd.Omiga.omPinMailer.ProcessPinMailerFulfilmentRequest;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.omPinMailer
{
	//[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
	//public interface IPinMailerBO
	//{
	//	[DispId(600)]string RunPinNumberTriggerInterface (string strRequest);
	//}
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	[Guid("48A5D003-377B-466e-AA7F-18CD9C76BA27")]
	[ComVisible(true)]
	[ProgId("omPinMailer.PinMailerBO")]
	public class PinMailerBO  //: IPinMailerBO
	{
		//private string _UserID;
		private string _ApplicationNumber;
		private string _ApplicationFactFindNumber;
		//private string[] _CustomerNumber;
		//private string[] _CustomerVersion;
		private string[] _OtherCustomerNumber;

		DirectGatewayBO dg = new DirectGatewayBO();
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		
		public PinMailerBO()
		{
		}

		//[STAThread]
		//static void Main() 
		//{
		//	PinMailerBO app = new PinMailerBO();
		//		
		//	//debug purposes
		//	XmlDocument doc = new XmlDocument();
		//	doc.Load ("../../PinMailer.xml");
		//	Console.Write(app.RunPinNumberTriggerInterface(doc.InnerXml));
		//}

		// interface method to pass delayed completion date information to Profile
		public string RunPinNumberTriggerInterface(string strRequest)
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST>
			//     <APPLICATION APPLICATIONNUMBER... APPLICATIONFACTFINDNUMBER.../>
			//	   <CUSTOMERLIST>
			//		   <CUSTOMER>
			//				<OTHERSYSTEMCUSTOMERNUMBER>12121212</OTHERSYSTEMCUSTOMERNUMBER>
			//		   </CUSTOMER>
			//	   </CUSTOMERLIST>
			// </REQUEST>
			string functionName = "RunPinNumberTriggerInterface";

			string respRun = string.Empty;
			XmlElement respRunXml=null;
			int custCount = 0;

			ProcessPinMailerFulfilmentResponseType resp;
			DirectGatewaySoapService service;
			ProcessPinMailerFulfilmentRequestType req;
			XmlDocument docVBReq = new XmlDocument();

			try
			{
				resp = new ProcessPinMailerFulfilmentResponseType();
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

				req = new  ProcessPinMailerFulfilmentRequestType();

				docVBReq.LoadXml(strRequest);

				//extract params
				XmlElement RequestElement = (XmlElement) docVBReq.SelectSingleNode("REQUEST");
				string userID = RequestElement.GetAttribute("USERID");   
				string password = RequestElement.GetAttribute("PASSWORD");  

				_ApplicationNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@APPLICATIONNUMBER").Value;
				_ApplicationFactFindNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@APPLICATIONFACTFINDNUMBER").Value;

				XmlNodeList reqlist = docVBReq.SelectNodes("//CUSTOMER");
				XmlNode custnode = null;
				//init arrays
				_OtherCustomerNumber = new string [reqlist.Count];
				custCount = reqlist.Count;
				for (int i=0; i<reqlist.Count; i++)
				{
					custnode = reqlist[i];
					_OtherCustomerNumber[i]= custnode.SelectSingleNode("OTHERSYSTEMCUSTOMERNUMBER").InnerText;
				}
				custnode = null;
				reqlist= null;

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

				// add info from VBRequest and call service per customer record
				for (int i=0; i< custCount; i++) 
				{
					req.CustomerNumber = _OtherCustomerNumber[i];

					if(_logger.IsDebugEnabled) 
					{
						string requestString  = dg.GetXmlFromGatewayObject(typeof(ProcessPinMailerFulfilmentRequestType), req);
						_logger.Debug(functionName + ": Service Request: " + requestString);
					}

					// call service
					resp = service.execute(req);

					if(_logger.IsDebugEnabled) 
					{
						string responseString  = dg.GetXmlFromGatewayObject(typeof(ProcessPinMailerFulfilmentResponseType), resp);
						_logger.Debug(functionName + ": Service Response: " + responseString);
					}

					if (resp.ErrorCode != "0") break;
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
				respRunXml.SelectSingleNode("//SOURCE").InnerText="RunPinNumberTriggerInterface";   // MAR750
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
			resp.SetAttribute ("TYPE","SUCCESS");
			return resp;
		}

		private XmlElement BuildErrorResponse()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement resp = doc.CreateElement("RESPONSE");
			resp.SetAttribute ("TYPE", "");
			XmlNode errNode = doc.CreateElement("ERROR");
			errNode.AppendChild(doc.CreateElement("NUMBER"));
			errNode.AppendChild(doc.CreateElement("SOURCE"));
			errNode.AppendChild(doc.CreateElement("DESCRIPTION"));
			resp.AppendChild(errNode);
			return resp;
		}	
	}
}
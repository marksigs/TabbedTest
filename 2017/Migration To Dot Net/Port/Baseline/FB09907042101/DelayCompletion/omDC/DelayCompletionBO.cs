/*
--------------------------------------------------------------------------------------------
Workfile:			omDelayCompletion.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		The DelayCompletion interface passes delayed completion date to Profile
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		24/10/2005	MAR29 Created
HMA     01/12/2005  MAR750  Update
GHun	13/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
MHeys	12/04/2006	MAR1603	Build errors rectified

--------------------------------------------------------------------------------------------
*/

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Text;
using System.Runtime.InteropServices;
//COM
using omApp;
using omROT;
using omPayProc;
//NET Ref
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.Core;
//web Ref
using Vertex.Fsd.Omiga.omDC.DelayCompletion;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.omDC
{
//	[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
//	public interface IDelayCompletionBO
//	{
//		[DispId(600)]string RunDelayedCompletionInterface (string strRequest);
//	}
//	[ClassInterface(ClassInterfaceType.AutoDual)]
	[Guid("626DCB4D-2CB3-4560-B4AF-7EF58607BB4D")]
	[ComVisible(true)]
	[ProgId("omDC.DelayCompletionBO")]
	public class DelayCompletionBO // : IDelayCompletionBO
	{
		private string _ApplicationNumber;
		private string _ApplicationFactFindNumber;
		private string _CompletionDate = String.Empty;
		private string _otherSystemNumber = String.Empty;

		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		DirectGatewayBO dg = new Vertex.Fsd.Omiga.omDg.DirectGatewayBO();
		omPayProc.PaymentProcessingBO _pp = new omPayProc.PaymentProcessingBOClass();
		
		public DelayCompletionBO()
		{
		}

//		[STAThread]
//		static void Main() 
//		{
//			DelayCompletionBO app = new DelayCompletionBO();
//			
//			//debug purposes
//			XmlDocument doc = new XmlDocument();
//			doc.Load ("../../RunDelayedCompletionInterface.xml");
//			Console.Write(app.RunDelayedCompletionInterface(doc.InnerXml));
//		}

		// interface method to pass delayed completion date information to Profile
		public string RunDelayedCompletionInterface(string strRequest)
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST>
			//     <APPLICATION APPLICATIONNUMBER... APPLICATIONFACTFINDNUMBER.../>
			// </REQUEST>

			string functionName = "RunDelayedCompletionInterface";

			string respRun = string.Empty;
			XmlElement respRunXml=null;
			bool bSuccess = true;
			XmlNode respDB = null;

			DelayCompletionResponseType resp;
			DirectGatewaySoapService service;
			DelayCompletionRequestType req;
			XmlDocument docVBReq = new XmlDocument();

			try
			{
				resp = new DelayCompletionResponseType();
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

				req = new  DelayCompletionRequestType();

				docVBReq.LoadXml(strRequest);

				//extract params
				// Get the UserID from the request
				XmlElement RequestElement = (XmlElement) docVBReq.SelectSingleNode("REQUEST");
				string userID = RequestElement.GetAttribute("USERID");   
				string password = RequestElement.GetAttribute("PASSWORD");  

				_ApplicationNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@APPLICATIONNUMBER").Value;
				_ApplicationFactFindNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@APPLICATIONFACTFINDNUMBER").Value;
				_CompletionDate = docVBReq.SelectSingleNode("//REQUEST/CASETASK/@TASKDUEDATE").Value;
				_otherSystemNumber = docVBReq.SelectSingleNode("//REQUEST/APPLICATION/@OTHERSYSTEMACCOUNTNUMBER").Value;

				//prepare web service request
				req.ClientDevice =	dg.GetClientDevice();	
				
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
				
				if (_CompletionDate.Length > 0)
				{
					req.RevisedCompletionDate = _CompletionDate.Substring(6,4) + "/" +
						_CompletionDate.Substring(3,2) + "/" +
						_CompletionDate.Substring(0,2);
				}

				req.MortgageAccountNumber = _otherSystemNumber;
				
				// call service
				if(_logger.IsDebugEnabled) 
				{
					string requestString  = dg.GetXmlFromGatewayObject(typeof(DelayCompletionRequestType), req);
					_logger.Debug(functionName + ": Service Request: " + requestString);
				}

				resp = service.execute(req);
				
				if(_logger.IsDebugEnabled) 
				{
					string responseString  = dg.GetXmlFromGatewayObject(typeof(DelayCompletionResponseType), resp);
					_logger.Debug(functionName + ": Service Response: " + responseString );
				}
			}	
			catch(System.Web.Services.Protocols.SoapException ex)
			{	bSuccess = false;
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "SOAP exception occurred: " + ex.Message.ToString();
			}
			catch(Exception ex)
			{	bSuccess = false;
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Exception occurred: " + ex.Message.ToString();
			}
			catch
			{	bSuccess = false;
				resp.ErrorCode = "SYSERR";
				resp.ErrorMessage = "Unknown exception occurred";
			} 
		
			XmlDocument docWebResp = new XmlDocument();
			// MAR1603 M Heys 10/04/2006 start
			try
			{
			// MAR1603 M Heys 10/04/2006
				if (bSuccess)
				{	//continue 
					string strResponse = dg.GetXmlFromGatewayObject (typeof (DelayCompletionResponseType),resp);
					docWebResp.LoadXml (strResponse);
					if ( docWebResp.SelectSingleNode("//ErrorCode").InnerText == "0")
					{
						// read payment summary
						bSuccess=false;
						respDB=null;
						GetPaymentSummary (out respDB); 
						if (respDB.SelectSingleNode("RESPONSE").Attributes.GetNamedItem("TYPE").InnerText=="SUCCESS")
						{							
							respRun = UpdateDisbursement (docVBReq, docWebResp, respDB);
						}
						else
						{	//return what we got
							respRun = respDB.OuterXml;
						}
					}
					else
					{	//return an error from Web
						respRunXml = BuildErrorResponse();
						respRunXml.SetAttribute("TYPE","APPERR");
						respRunXml.SelectSingleNode("//NUMBER").InnerText=resp.ErrorCode;
						respRunXml.SelectSingleNode("//DESCRIPTION").InnerText=resp.ErrorMessage;
						respRunXml.SelectSingleNode("//SOURCE").InnerText="RunDelayedCompletionInterface";
						respRun = respRunXml.OuterXml;
					}
				}
				else
				{	//return what we got
					if (resp.ErrorMessage.Length >0)
					{	//comes from exeption
						respRunXml = BuildErrorResponse();
						respRunXml.SetAttribute("TYPE",resp.ErrorCode); 
						respRunXml.SelectSingleNode("//NUMBER").InnerText="-1";
						respRunXml.SelectSingleNode("//DESCRIPTION").InnerText=resp.ErrorMessage;
						respRunXml.SelectSingleNode("//SOURCE").InnerText="RunDelayedCompletionInterface";
						respRun = respRunXml.OuterXml;
					}
					else
						respRun = respDB.OuterXml;
				}
			}
			// MAR1603 M Heys 10/04/2006
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(_pp);
			}
			// MAR1603 M Heys 10/04/2006 end
			req = null;
			resp = null;
			docVBReq = null;
			respRunXml = null;
			respDB = null;
			return respRun;
		}

		private void GetReportOnTitleData(out XmlNode respNodeOut )
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST OPERATION="GETREPORTONTITlEDATA">
			//     <REPORTONTITLE APPLICATIONNUMBER... APPLICATIONFACTFINDNUMBER.../>
			// </REQUEST>

			XmlDocument doc = new XmlDocument();
			string response = string.Empty;
			// Prepare the reequest for Queue
			XmlElement rootElem = doc.CreateElement("REQUEST");
			rootElem.SetAttribute("OPERATION", "GETREPORTONTITLEDATA");
			doc.AppendChild(rootElem);
		
			// Set up message element
			XmlElement msgElem = doc.CreateElement("REPORTONTITLE");
			msgElem.SetAttribute("APPLICATIONNUMBER", _ApplicationNumber);
			msgElem.SetAttribute("APPLICATIONFACTFINDNUMBER", _ApplicationFactFindNumber);
			rootElem.AppendChild(msgElem);

			string respROT = String.Empty;
			omROT.omRotBO rot = new omROT.omRotBOClass();//MAR1603 M Heys 12/04/2006
			try
			{
				//omROT.omRotBO rot = new omROT.omRotBOClass();//MAR1603 M Heys 12/04/2006
				respROT = rot.OmRotRequest (doc.OuterXml);
			}
			catch(Exception ex)
			{
				string errmsg = ex.Message.ToString();
			}
			// MAR1603 M Heys 10/04/2006 start
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(rot);
			}
			// MAR1603 M Heys 10/04/2006 end

			doc.LoadXml (respROT);
			respNodeOut = (XmlNode) doc.SelectSingleNode(".");
		}

		private void GetApplicationData (out XmlNode respNodeOut)
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST OPERATION="GETAPPLICATIONDATA">
			//     <APPLICATION> 
			//		  <APPLICATIONNUMBER>
			//		  <APPLICATIONFACTFINDNUMBER>
			//     </APPLICATION> 
			// </REQUEST>

			XmlDocument doc = new XmlDocument();
			string response = string.Empty;
			// Prepare the reequest for Queue
			XmlElement rootElem = doc.CreateElement("REQUEST");
			rootElem.SetAttribute("OPERATION", "GETAPPLICATIONDATA");
			doc.AppendChild(rootElem);
		
			XmlElement appElem = doc.CreateElement("APPLICATION");
			XmlElement childElem = doc.CreateElement("APPLICATIONNUMBER");
			childElem.InnerText = _ApplicationNumber;
			appElem.AppendChild (childElem.CloneNode(true));

			childElem = doc.CreateElement("APPLICATIONFACTFINDNUMBER");
			childElem.InnerText = _ApplicationFactFindNumber;
			appElem.AppendChild (childElem.CloneNode(true));

			rootElem.AppendChild(appElem);
			
			string respApp = String.Empty;
			omApp.ApplicationBO ap = new omApp.ApplicationBOClass();//MAR1603 M Heys 12/04/2006
			try
			{
				//omApp.ApplicationBO ap = new omApp.ApplicationBOClass();//MAR1603 M Heys 12/04/2006
				respApp = ap.GetApplicationData (doc.OuterXml);
			}
			catch(Exception ex)
			{
				string errmsg = ex.Message.ToString();
			}
			// MAR1603 M Heys 10/04/2006 start
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ap);
			}
			// MAR1603 M Heys 10/04/2006 end
			doc.LoadXml (respApp);
			respNodeOut = doc.SelectSingleNode(".");
		}

		private void GetPaymentSummary (out XmlNode respNodeOut)
		{	
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST OPERATION="GETPAYMENTSUMMARY">
			//     <PAYMENTRECORD APPLICATIONNUMBER=... _COMBOLOOKUP_=.../>
			// </REQUEST>

			XmlDocument doc = new XmlDocument();
			XmlElement req = doc.CreateElement("REQUEST");
			req.SetAttribute("OPERATION", "GETPAYMENTSUMMARY");

			XmlElement elem = doc.CreateElement("PAYMENTRECORD");
			elem.SetAttribute("APPLICATIONNUMBER", _ApplicationNumber);
			elem.SetAttribute("_COMBOLOOKUP_", "1");
			req.AppendChild(elem);
			
			string respPP=String.Empty;
			try
			{
				respPP = _pp.omPayProcRequest (req.OuterXml);
			}
			catch(Exception ex)
			{
				string errmsg = ex.Message.ToString();
			}
			doc.LoadXml (respPP);
			respNodeOut = (XmlNode) doc.SelectSingleNode(".");
		}

		private string UpdateDisbursement (XmlDocument docVBReq, XmlDocument docWebResp, XmlNode respDB)
		{
			// pass:         vstrXMLRequest  xml Request data stream
			// <REQUEST OPERATION="UPDATEDISBURSEMENT">
			//     <PAYMENTRECORD APPLICATIONNUMBER=... >
			//		  <DISBURSEMENTPAYMENT FIRSTPAYMENTDATE=... />
			//     </PAYMENTRECORD>
			// </REQUEST>

			XmlElement resp = BuildResponse();
			string response = resp.OuterXml;
			string respPP = String.Empty;

			XmlNode nodePay = null;
			XmlElement elemDisb = null;

			XmlNodeList listPay = respDB.SelectNodes ("//PAYMENTRECORD[DISBURSEMENTPAYMENT/@DELAYEDCOMPLETION='1']");
			
			if ( listPay.Count == 0 )
			{	//nothing to update - return type=success
				resp.SetAttribute ("TYPE","SUCCESS");
				response = resp.OuterXml;
			}
			else
			{	//create requests for disbursement
				XmlDocument doc = new XmlDocument();

				XmlNode nodeTemp = (docVBReq.SelectSingleNode("//REQUEST").CloneNode(false));
				XmlElement req = (XmlElement) doc.ImportNode(nodeTemp,true);
				nodeTemp=null;

				req.SetAttribute ("OPERATION","UPDATEDISBURSEMENT");
				
				string payDate = docWebResp.SelectSingleNode("//DateOfFirstPayment").InnerText;
				string payAmount = docWebResp.SelectSingleNode("//AmountOfFirstPayment").InnerText;
				string payAmountSub = docWebResp.SelectSingleNode("//AmountOfSubsequentPayment").InnerText;

				for (int i = 0; i < listPay.Count ; i++)
				{
					//update disbursement
					nodeTemp = listPay.Item(i).CloneNode(true);
					nodePay = doc.ImportNode(nodeTemp, true);

					elemDisb = (XmlElement) nodePay.SelectSingleNode("DISBURSEMENTPAYMENT");
					elemDisb.RemoveChild (elemDisb.SelectSingleNode("THIRDPARTY"));
					elemDisb.SetAttribute("FIRSTPAYMENTDATE", payDate);
					elemDisb.SetAttribute("FIRSTPAYMENT", payAmount);
					elemDisb.SetAttribute("REGULARPAYMENT", payAmountSub);
					elemDisb.SetAttribute("DELAYEDCOMPLETION", "0");
					nodePay.AppendChild (elemDisb);
					
					req.AppendChild (nodePay);
					
					try
					{
						respPP = _pp.omPayProcRequest (req.OuterXml);
					}
					catch(Exception ex)
					{
						string errmsg = ex.Message.ToString();
					}
					response = respPP;
				}
			}
			return response;
		}

		private XmlElement BuildResponse()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement resp = doc.CreateElement("RESPONSE");
			resp.SetAttribute("TYPE", "SUCCESS");
			return resp;
		}

		private XmlElement BuildErrorResponse()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement resp = doc.CreateElement("RESPONSE");
			resp.SetAttribute("TYPE", "");
			XmlNode errNode = doc.CreateElement("ERROR");
			errNode.AppendChild(doc.CreateElement("NUMBER"));
			errNode.AppendChild(doc.CreateElement("SOURCE"));
			errNode.AppendChild(doc.CreateElement("DESCRIPTION"));
			resp.AppendChild(errNode);
			return resp;
		}
	}
}

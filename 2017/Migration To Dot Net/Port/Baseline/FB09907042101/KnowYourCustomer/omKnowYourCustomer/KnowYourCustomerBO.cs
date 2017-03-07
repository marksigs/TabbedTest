/*
--------------------------------------------------------------------------------------------
Workfile:			KnowYourCustomerBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		26/10/2005	MAR278 Included logging
GHun	07/11/2005	MAR281 Use omCore to get database connection string
GHun	16/11/2005	MAR467 Pass OtherSystemCustomerNumber
GHun	13/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
PSC		20/01/2006	MAR964 Amend RunKnowYourCustomer to check KYC error too
LDM		13/03/2006	EP16   Epsom. New process for Epsom to make a call to experian, and save away the results
PE		25/07/2006	EP974/MAR1848 Mars merge/Prevent transaction enlistment
--------------------------------------------------------------------------------------------
*/
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization; // for XmlSerializer
using System.Xml.Xsl; 
using System.IO;	// for reading file
using System.Data;
using System.Data.SqlClient;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omExperianWrapper; //LDM 06/02/2006	Epsom
using XmlManager = Vertex.Fsd.Omiga.omKnowYourCustomer.XML.XMLManager;

namespace Vertex.Fsd.Omiga.omKnowYourCustomer
{
	[ProgId("omKnowYourCustomer.KnowYourCustomerBO")]
	[Guid("B4879894-DE10-49fb-A4F5-7C56B8CE5A54")]
	[ComVisible(true)]
	public class KnowYourCustomerBO
	{
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		private string _AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
		private XmlManager _xmlMgr = new XmlManager();

		public KnowYourCustomerBO()
		{
		}

		public string RunKnowYourCustomer(string strRequest)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(MethodInfo.GetCurrentMethod().Name + " : Request passed in : " + strRequest);
			}

			XmlDocument requestXML = null;
			XmlDocument responseXML = null;
			XmlDocument experianRespXML = null;
			XmlDocument tempXML = null;
			XmlNode requestNode = null;
			XmlNode testNode = null;
			XmlNode tempNode = null;
			XmlNode responseNode = null;	
			XmlNode typeAttrib = null;
			ExperianWrapperBO expWrapper = null; 
			XslTransform xslt = null;
			StringWriter sw = null;
			string newRequest;
			string strResponse;

			try
			{
				//build up response
				responseXML = new XmlDocument();
				tempNode = responseXML.CreateElement("RESPONSE");
				responseNode = responseXML.AppendChild(tempNode);	
				typeAttrib = responseXML.CreateAttribute("TYPE");
				responseNode.Attributes.SetNamedItem(typeAttrib);

				requestXML = new XmlDocument();
				requestXML.LoadXml(strRequest);
				requestNode = _xmlMgr.xmlGetMandatoryNode(requestXML, "//RESPONSE", true);

				// Get the Customer numbers from the request
				string customerNumber = _xmlMgr.xmlGetMandatoryNodeText(
					_xmlMgr.xmlGetMandatoryNode(requestXML, "//RESPONSE/CUSTOMER[1]/CUSTOMERVERSION", true), "CUSTOMERNUMBER", true);
				string customerVersionNumber = _xmlMgr.xmlGetMandatoryNodeText(
					_xmlMgr.xmlGetMandatoryNode(requestXML, "//RESPONSE/CUSTOMER[1]/CUSTOMERVERSION", true), "CUSTOMERVERSIONNUMBER", true);

				newRequest = GetKYCDataRequest(customerVersionNumber, customerNumber);
				tempXML = new XmlDocument();
				tempXML.LoadXml(newRequest);
				string applicationNumber = tempXML.SelectSingleNode("EXPERIAN/APPLICATION/APPLICATIONNUMBER").InnerText;

				// Restructure the xml for Experian ...
				sw = new StringWriter();
				xslt = new XslTransform();
				xslt.Load(GetXMLPath() + "\\" + "CreateExperianKYCBlocks.xslt");
				xslt.Transform(tempXML.FirstChild,null,sw,null);
				tempXML.LoadXml(sw.ToString());

				experianRespXML = new XmlDocument();
				expWrapper = new ExperianWrapperBO();	//LDM 02/02/2006 Epsom
				expWrapper.CallExperian("KYC", tempXML.DocumentElement, ref experianRespXML, applicationNumber);

				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Return from Experian:" + experianRespXML.OuterXml);
				}

				// Has it performed a know your customer?
				testNode = experianRespXML.SelectSingleNode("//REQUEST");
				if (testNode == null)
				{
					// Unexpected data
					throw new OmigaException("Unexpected response whilst submitting a know your customer.");
				}
				else
				{
					SaveXMLKYCData(experianRespXML.DocumentElement, requestNode, tempXML.DocumentElement);
				}

				typeAttrib.InnerText = "SUCCESS";
				strResponse = responseXML.OuterXml;
			}

			catch(Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				strResponse = new OmigaException(ex).ToOmiga4Response();
			}
			finally
			{
				requestXML = null;
				responseXML = null;
				experianRespXML = null;
				tempXML = null;
				requestNode = null;
				testNode = null;
				tempNode = null;
				responseNode = null;	
				typeAttrib = null;
				expWrapper = null; 
				sw.Flush();
				sw.Close();
				sw = null;
				xslt = null;
			}

			return strResponse;
		}

		private string GetKYCDataRequest(string strCustomerVersionNumber, string strCustomerNumber)
		{
			string strRequest = "";

			try
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Parameters to USP_GetKnowYourCustomerDataRequest. Customer No: " 
						+ strCustomerNumber + " Customer version Number: " + strCustomerVersionNumber);
				}

				//MAR182 GHun use omCore to get database connection string, and make sure connection gets closed			
				//EP974/MAR1848 - prevent transaction enlistment
				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString + "Enlist=false;"))

				{
					SqlCommand requestCMD = new SqlCommand("USP_GetKnowYourCustomerDataRequest", conn);
					requestCMD.CommandType = CommandType.StoredProcedure;

					SqlParameter myParm = requestCMD.Parameters.Add("@CustomerNumber", SqlDbType.NVarChar, strCustomerNumber.Length);
					myParm.Value = strCustomerNumber;
					myParm = requestCMD.Parameters.Add("@CustomerVersionNumber", SqlDbType.Int, 5);
					myParm.Value = Convert.ToInt16(strCustomerVersionNumber, 10);
					
					conn.Open();
					SqlDataReader myReader = requestCMD.ExecuteReader();
					strRequest = "<EXPERIAN>";
					do 
					{
						if (myReader.Read())
						{
							strRequest += myReader.GetString(0);
						}
					} while (myReader.NextResult());

					myReader.Close();
					strRequest += "</EXPERIAN>";

					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name + ": Return from USP_GetKnowYourCustomerDataRequest:" + strRequest);
					}
				}
			}
			catch(Exception ex)
			{
				throw new OmigaException(ex);
			}
			
			return strRequest;
		}

		private string GetXMLPath()
		{
			string thePath = _AppPath.ToUpper();
			return(thePath.Replace("\\DLLDOTNET", "\\XML"));
		}

		private void SaveXMLKYCData(XmlNode gatewayResponse, XmlNode methodRequest, 
			XmlNode omigaData)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			XslTransform xslt = null;
			StringWriter sw = null;
			XmlDocument transformedXML = null;
			XmlDocument headerXML = null;
			SqlCommand cmd = null;
			SqlParameter param = null;
			XmlNode newNode = null;
			XmlNode newAttrib = null;
			XmlNode dataNode = null;
			XmlNode appNode = null;
			try
			{
				// Restructure the xml for Experian ...
				xslt = new XslTransform();
				sw = new StringWriter();
				xslt.Load(GetXMLPath() + "\\" + "saveXMLKYC.xslt");
				xslt.Transform(gatewayResponse,null,sw,null);
				transformedXML = new XmlDocument();
				transformedXML.LoadXml(sw.ToString());

				// Save the data to the Omiga database ...
				// Get the Customer numbers from the request
				string customerNumber = _xmlMgr.xmlGetMandatoryNodeText(
					_xmlMgr.xmlGetMandatoryNode(methodRequest, "//RESPONSE/CUSTOMER[1]/CUSTOMERVERSION", true), "CUSTOMERNUMBER", true);
				string customerVersionNumber = _xmlMgr.xmlGetMandatoryNodeText(
					_xmlMgr.xmlGetMandatoryNode(methodRequest, "//RESPONSE/CUSTOMER[1]/CUSTOMERVERSION", true), "CUSTOMERVERSIONNUMBER", true);
				string userID = _xmlMgr.xmlGetMandatoryNodeText(
					_xmlMgr.xmlGetMandatoryNode(methodRequest, "//RESPONSE", true), "USERID", true);

				appNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//EXP", true);
				string kycRefNo = _xmlMgr.xmlGetAttributeText(appNode, "EXPERIANREF");
				
				headerXML = new XmlDocument();
				newNode = headerXML.CreateElement("HEADER");
				newNode = headerXML.AppendChild(newNode);	
				newAttrib = headerXML.CreateAttribute("CUSTOMERNUMBER");
				newAttrib.InnerText = customerNumber;
				newNode.Attributes.SetNamedItem(newAttrib);
				newAttrib = headerXML.CreateAttribute("CUSTOMERVERSIONNUMBER");
				newAttrib.InnerText = customerVersionNumber;
				newNode.Attributes.SetNamedItem(newAttrib);
				newAttrib = headerXML.CreateAttribute("USERID");
				newAttrib.InnerText = userID;
				newNode.Attributes.SetNamedItem(newAttrib);
				newAttrib = headerXML.CreateAttribute("KYCREFERENCENUMBER");
				newAttrib.InnerText = kycRefNo;
				newNode.Attributes.SetNamedItem(newAttrib);

				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString))
				{
					cmd = new SqlCommand("USP_SAVEKYCDATA", conn);
					cmd.CommandType = CommandType.StoredProcedure;

					param = cmd.Parameters.Add("@p_Header", SqlDbType.NText);
					param.Value = headerXML.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//AUTP", true);
					param = cmd.Parameters.Add("@p_Autp", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetNode(transformedXML.DocumentElement, "//POLICYBLOCK", false);
					param = cmd.Parameters.Add("@p_Policy", SqlDbType.NText);
					if (dataNode != null)
					{
						param.Value = dataNode.OuterXml;
					}
					else
					{
						param.Value = null;
					}

					conn.Open();
					cmd.ExecuteNonQuery();
				}
			}
			catch(Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				throw ex;
			}
			finally
			{
				sw.Flush();
				sw.Close();
				sw = null;
				xslt = null;
				transformedXML = null;
				headerXML = null;
				cmd.Dispose();
				param = null;
				newNode = null;
				newAttrib = null;
				dataNode = null;
				appNode = null;
			}
		}
	}
}

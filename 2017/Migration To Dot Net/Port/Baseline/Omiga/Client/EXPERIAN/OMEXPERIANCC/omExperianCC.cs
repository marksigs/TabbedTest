/*
--------------------------------------------------------------------------------------------
Workfile:			omExperianCC.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Prepares and executes an Experian credit check call.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
SAB		11/10/2005	MAR13 Created
GHun	18/10/2005	MAR184 Made changes to get calling from VB to work
SAB		28/10/2005	MAR245 Updated to allow address targeting
MV		01/11/2005	MAR368	Amended Logging
MV		03/11/2005	MAR368	Amended getGateWayResponse()
INR		27/11/2005	MAR577	SecondApplicant not set up correctly
GHun	13/12/2005	MAR852	set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
GHun	14/12/2005	MAR871 Provide a better deserialization error message
RF      13/02/2006  MAR1243 Extra logging added for address targeting investigation
PSC		03/03/2006	MAR1348 Amend updateAddressData for DETAILCODE, RMCCODE and REGIONCODE

------------------------------ CUT FOR EPSOM -----------------------------------------------

LDM		22/03/2006	EP6		Epsom  Put in a stub to represent the calls to Experian
PSC		29/03/2006	MAR1554 Amend RunXMLCreditCheck so it doesn't enlist in any
					encompassing transaction when getting the request data
PSC		31/03/2006	MAR1463	Improve error reporting --LDM not put in, dont use gateway.
HMA     04/04/2006  MAR1375 Corrected timestamp format.
LDM		13/04/2006  EP352   Changes for address targetting
SAB		03/05/2006	EP478	Updated to pass the Experian reference passed as part of an address
							targetted response.
LDM		30/05/2006  EP627	Apply Mars changes to Epsom (GHun	06/04/2006	MAR1409 Handle validation errors in request XML)
SAB		21/06/2006  EP688	Only populate DETAILCODE for a targeted address if the REGION and
							RMC codes have been passed.
SR		16/02/2007	EP2_659	implemented MAR1896. Return standard response if UseExperianCCStub global parameter is set.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.Core;
using XmlManager = Vertex.Fsd.Omiga.omExperianCC.XML.XMLManager;
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.omExperianWrapper; //LDM 31/01/2006	EP6 Epsom
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.omExperianCC
{
	[ProgId("omExperianCC.ExperianCCBO")]
	[Guid("565E3A2D-AE1A-4620-B2BB-5FAAC5F6A9CB")]
	[ComVisible(true)]
	public class ExperianCCBO
	{
		private XmlManager _xmlMgr = new XmlManager();
		private string _AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public ExperianCCBO()
		{
		}

		public string Execute(string request)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name + 
					". Request: " + request); 
			}

			XmlDocument requestXML = new XmlDocument();
			XmlNode requestNode;
			XmlDocument responseXML = new XmlDocument();
			XmlNode tempNode = null;
			XmlNode responseNode = null;	
			XmlNode typeAttrib = null;
			XmlNodeList operationNodeList;
			string response = "";
			// Create the response node ...
			tempNode = responseXML.CreateElement("RESPONSE");
			responseNode = responseXML.AppendChild(tempNode);	
			typeAttrib = responseXML.CreateAttribute("TYPE");
			responseNode.Attributes.SetNamedItem(typeAttrib);
			try
			{
				requestXML.LoadXml(request);
				requestNode = _xmlMgr.xmlGetMandatoryNode(requestXML, "//REQUEST", true);

				if (requestNode.Attributes.GetNamedItem("OPERATION") == null)
				{
					operationNodeList = requestNode.SelectNodes("OPERATION");
					foreach (XmlNode xndOperation in operationNodeList)
					{
						foreach (XmlAttribute xmlAttrib in requestNode.Attributes)
						{
							xndOperation.Attributes.SetNamedItem(xmlAttrib.CloneNode(true));
						}
						ForwardRequest(xndOperation, ref responseNode);
					}
				}
				else
				{
					ForwardRequest(requestNode, ref responseNode);
				}
							
				typeAttrib.InnerText = "SUCCESS";
				response = responseXML.OuterXml;
			
			}
			catch(Exception ex)
			{	
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				// Create the response node ...
				typeAttrib.InnerText = "ERROR";
				response = new OmigaException(282, ex.Message).ToOmiga4Response();
			}
			finally
			{
				requestXML = null;
				requestNode = null;
				responseXML = null;
				tempNode = null;
				typeAttrib = null;
				responseNode = null;	
				operationNodeList = null;
			}
			return response;
		}
		
		private void ForwardRequest(XmlNode requestNode, ref XmlNode responseNode)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			string operation = "";
			try
			{
				operation = (requestNode.Name == "REQUEST") ? 
					requestNode.Attributes.GetNamedItem("OPERATION").InnerText.ToUpper() :
					requestNode.Attributes.GetNamedItem("NAME").InnerText.ToUpper();

				switch (operation)
				{
					case ("RUNXMLCREDITCHECK"):
						RunXMLCreditCheck(requestNode, ref responseNode);
						break;
					case ("RUNXMLUPGRADECREDITCHECK")://LDM 22/03/06 EP6 upgrade experian call
						RunXMLUpgradeCreditCheck(requestNode, ref responseNode);
						break;
					default:
						break;
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
			}
		}	
		
		private void RunXMLCreditCheck(XmlNode requestNode, ref XmlNode responseNode)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			// Get data for the Experian call
			XmlDocument omigaDataXML = new XmlDocument();
			XmlDocument experianXML = new XmlDocument();
			XmlDocument responseXML = new XmlDocument();
			ExperianWrapperBO expWrapper = new ExperianWrapperBO();	//LDM 02/02/2006 EP6 Epsom
			SqlCommand cmd = null;
			SqlParameter param = null;
			SqlDataReader sqlDR = null;
			XmlNode ccNode = null;
			XmlNode newElement = null;
			XmlNode testNode = null;
			XmlNode AUKTest = null;
			XmlNode BUKTest = null;
			XmlNode requestRoot = null;
			XmlNode targetingReqNode = null;
			XmlNode targetingNode = null;
			XmlNode experianRefNode = null;
			XmlNodeList newNodeList = null;
			XslTransform xslt = new XslTransform();
			StringWriter sw = new StringWriter();
			XmlNode validationNode = null;	//MAR1409 GHun
			try
			{
				// Extract the relevant data from Omiga ...
				//MAR184 GHun Find application attributes on the APPLICATION node
				ccNode = _xmlMgr.xmlGetMandatoryNode(requestNode, "//APPLICATION", true);
				string applicationNumber = _xmlMgr.xmlGetMandatoryAttributeText(ccNode, "APPLICATIONNUMBER");
				string applicationFFN = _xmlMgr.xmlGetMandatoryAttributeText(ccNode, "APPLICATIONFACTFINDNUMBER");

				// PSC 29/03/2006 MAR1554 - Don't enlist in transaction for the getting of data
				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString + "Enlist=false;"))
				{
					cmd = new SqlCommand("USP_GETEXPERIANREQUESTDATA", conn);
					cmd.CommandType = CommandType.StoredProcedure;
					param = cmd.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, 12);
					param.Value = applicationNumber;
					param = cmd.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int, 5);
					param.Value = Convert.ToInt16(applicationFFN, 10);		

					conn.Open();
					StringBuilder spXMLresponse = new StringBuilder("<RESPONSE");

					//MAR184 GHun Copy the reprocess and rescore attributes from the request
					if (_xmlMgr.xmlGetAttributeText(requestNode, "RESCORE") == "1")
						spXMLresponse.Append(" RESCORE='1'");
					if (_xmlMgr.xmlGetAttributeText(requestNode, "REPROCESS") == "1")
						spXMLresponse.Append(" REPROCESS='1'");
					spXMLresponse.Append(">");

					sqlDR = cmd.ExecuteReader();
					do 
					{
					while (sqlDR.Read()) // DataReader doesn't always fit the XML in one row.
					{
						spXMLresponse.Append(sqlDR.GetString(0));
					}
					} while (sqlDR.NextResult());
					spXMLresponse.Append("</RESPONSE>");
					omigaDataXML.LoadXml(spXMLresponse.ToString());			
				}

				// Restructure the xml for Experian ...
				xslt.Load(GetXMLPath() + "\\" + "runXMLCreditCheck.xslt");
				
				xslt.Transform(omigaDataXML.FirstChild,null,sw,null);
				experianXML.LoadXml(sw.ToString());

				//MAR1409 GHun
				testNode = experianXML.DocumentElement;
				if (testNode != null)
				{	
					validationNode = testNode.SelectSingleNode("VALIDATION");
					if (validationNode != null)
					{
						if (validationNode.HasChildNodes)
						{
							string message = "A credit check cannot be performed due to the following missing information:\r\n";
							foreach(XmlNode error in validationNode.SelectNodes("ERROR"))
							{
								message += error.InnerText + "\r\n";
							}
							throw new OmigaException(message);
						}
						testNode.RemoveChild(validationNode);
					}
				}
				//MAR1409 End

				// Check if the request is being resubmitted due to address targetting ...
				targetingReqNode = requestNode.SelectSingleNode("//TARGETINGDATA[ADDRESSTARGETING='YES']");
				if (targetingReqNode != null)
				{
					requestRoot = experianXML.DocumentElement;
					updateAddressData(ref requestRoot, targetingReqNode);

					// Add the Experian reference to the request
					experianRefNode = targetingReqNode.SelectSingleNode("//EXPERIANREF");
					addExperianRefToRequest(ref requestRoot, experianRefNode);
				}

				//SR EP2_659 : MAR1896 If UseExperianCCStub exists and is set to TRUE, load a standard response.
				bool bUseStub;
				try
				{
					bUseStub = new GlobalParameter("UseExperianCCStub").Boolean;
				}
				catch
				{
					bUseStub = false;
				}
				
				if (bUseStub == true)
				{
					// Log use of stub.
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(MethodInfo.GetCurrentMethod().Name +  
							": Using Experian stub.");
					}

					// Delay for a given time (read from the registry)
					BlockThread();

					responseXML.Load(_AppPath + "//StandardCreditCheckResponse.xml");
				}
				else   //SR EP2_659 :  End
				{
					// Load into the response ...
					//getGatewayResponse(experianXML, requestNode, ref responseXML);
					// LDM 30/1/6 EP6 Epsom. Put in the option of using a stub rather than a real Experian call
					expWrapper.CallExperian("CreditCheck", experianXML.DocumentElement, ref responseXML, applicationNumber);
				}				

				// Has it performed a full credit check ?
				testNode = responseXML.SelectSingleNode("//REQUEST");
				
				// Is address targetting required ?  LDM 13/04/2006 EP352
				AUKTest = responseXML.SelectSingleNode("//ADDR-UK");
				BUKTest = responseXML.SelectSingleNode("//ADDR-UKBUREA");

				if (testNode != null && AUKTest == null && BUKTest == null )
				{
					string functionID = _xmlMgr.xmlGetMandatoryNodeText(
						_xmlMgr.xmlGetMandatoryNode(experianXML.DocumentElement, "AC01", true), "FUNCTION", true);
					SaveXMLCreditCheck(functionID, responseXML.DocumentElement, requestNode, 
						omigaDataXML.DocumentElement, targetingReqNode);
				}

				
				testNode = (AUKTest != null) ? AUKTest : BUKTest;

				if (testNode != null)
				{
					targetingNode = responseXML.CreateElement("TARGETINGDATA");
			
					// Add the Experian Request Data
					newElement = responseXML.ImportNode(experianXML.DocumentElement, true);
					targetingNode.AppendChild(newElement);
					
					// Add the Customer and ApplicationFactFind elements from the Omiga Data						
					newElement = responseXML.ImportNode(omigaDataXML.SelectSingleNode("//APPLICATIONFACTFIND"), true);
					targetingNode.AppendChild(newElement);
					
					newNodeList = omigaDataXML.SelectNodes("RESPONSE/CUSTOMER");
					foreach (XmlNode omigaCust in newNodeList)
					{
						newElement = responseXML.ImportNode(omigaCust, true);
						targetingNode.AppendChild(newElement);
					}
					responseXML.DocumentElement.AppendChild(targetingNode);

					// Determine what should happen now ...
					formatTargettingData(responseXML.DocumentElement, ref requestNode);
					targetingReqNode = requestNode.SelectSingleNode("//TARGETINGDATA");

					if (targetingReqNode != null)
					{
						testNode = targetingReqNode.SelectSingleNode("ADDRESSTARGET[@ADDRESSRESOLVED='N']");
						// Extract the Experian Reference as this will need to be passed with the
						// resolved addresses
						experianRefNode = responseXML.SelectSingleNode("//REQUEST/@EXP_ExperianRef");

						if (testNode != null)
						{
							// User intervention required so return the TARGETINGDATA element to the GUI
							if(_logger.IsDebugEnabled) 
							{
								_logger.Debug("In " + MethodInfo.GetCurrentMethod().Name + 
									". Address targeting is required."); 
							}
							newElement = requestNode.SelectSingleNode("//TARGETINGDATA");
							addExperianRefToRequest(ref newElement, experianRefNode);
							responseNode.AppendChild(responseNode.OwnerDocument.ImportNode(newElement, true));
						}
						else
						{
							// Update the address data, add the experian reference and automatically resubmit ...
							requestRoot = experianXML.DocumentElement;
							updateAddressData(ref requestRoot, targetingReqNode);
							addExperianRefToRequest(ref requestRoot, experianRefNode);

							//getGatewayResponse(experianXML, requestNode, ref responseXML);
							// LDM 30/1/6 EP6 Epsom. Put in the option of using a stub rather than a real Experian call
							expWrapper.CallExperian("AddressTarget", experianXML.DocumentElement, ref responseXML, applicationNumber);

							testNode = responseXML.SelectSingleNode("//REQUEST");
							if (testNode == null)
							{
								// Unexpected data
								throw new OmigaException("Unexpected response whilst resubmitting a credit check.");
							}
							else
							{
								// Save the data to the Omiga DB ...
								string functionID = _xmlMgr.xmlGetMandatoryNodeText(
									_xmlMgr.xmlGetMandatoryNode(experianXML.DocumentElement, "AC01", true), "FUNCTION", true);
								SaveXMLCreditCheck(functionID, responseXML.DocumentElement, requestNode, 
									omigaDataXML.DocumentElement, targetingReqNode);
							}
						}
					}
					else
					{
						throw new OmigaException("Invalid address targeting data.");
					}
				}
				// So far, the TARGETINGDATA element will only have been created if address
				// targeting is required. If it is not required an element still however needs to be created.
				targetingNode = responseNode.SelectSingleNode("//TARGETINGDATA");
				if (targetingNode == null)
				{
					targetingNode = responseNode.OwnerDocument.CreateElement("TARGETINGDATA");
					newElement = responseNode.OwnerDocument.CreateElement("ADDRESSTARGETING");
					newElement.InnerText = "NO";
					targetingNode.AppendChild(newElement);
					responseNode.AppendChild(targetingNode);				
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
				omigaDataXML = null;
				experianXML = null;
				responseXML = null;
				cmd = null;
				param = null;
				sqlDR = null;
				ccNode = null;
				newElement = null;
				AUKTest = null;
				BUKTest = null;
				requestRoot = null;
				targetingReqNode = null;
				targetingNode = null;
				experianRefNode = null;
				newNodeList = null;
				xslt = null;
				expWrapper = null;
			}
		}
		
		//		LDM 22/03/06 EP6
		//		New procedure to upgrade an enquiry "footprint" to a full application "footprint"		
		private void RunXMLUpgradeCreditCheck(XmlNode requestNode, ref XmlNode responseNode)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			// Get data for the Experian call
			XmlDocument omigaDataXML = new XmlDocument();
			XmlDocument experianXML = new XmlDocument();
			XmlDocument responseXML = new XmlDocument();
			ExperianWrapperBO expWrapper = new ExperianWrapperBO();
			SqlCommand cmd = null;
			SqlParameter param = null;
			SqlDataReader sqlDR = null;
			XmlNode ccNode = null;
			XslTransform xslt = new XslTransform();
			StringWriter sw = new StringWriter();
			try
			{
				// Extract the relevant data from Omiga ...
				ccNode = _xmlMgr.xmlGetMandatoryNode(requestNode, "//APPLICATION", true);
				string applicationNumber = _xmlMgr.xmlGetMandatoryAttributeText(ccNode, "APPLICATIONNUMBER");
				string applicationFFN = _xmlMgr.xmlGetMandatoryAttributeText(ccNode, "APPLICATIONFACTFINDNUMBER");

				// LDM 05/04/2006 MAR1554 - Don't enlist in transaction for the getting of data
				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString + "Enlist=false;"))
				{
					cmd = new SqlCommand("Usp_GetExperianUpgradeData", conn);
					cmd.CommandType = CommandType.StoredProcedure;
					param = cmd.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, 12);
					param.Value = applicationNumber;
					param = cmd.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int, 5);
					param.Value = Convert.ToInt16(applicationFFN, 10);		

					conn.Open();
					StringBuilder spXMLresponse = new StringBuilder("<RESPONSE>");

					sqlDR = cmd.ExecuteReader();
					do 
					{
					while (sqlDR.Read()) // DataReader doesn't always fit the XML in one row.
					{
						spXMLresponse.Append(sqlDR.GetString(0));
					}
					} while (sqlDR.NextResult());
					spXMLresponse.Append("</RESPONSE>");
					omigaDataXML.LoadXml(spXMLresponse.ToString());			
				}

				// Restructure the xml for Experian ...
				xslt.Load(GetXMLPath() + "\\" + "RunXMLUpgrade.xslt");
				
				xslt.Transform(omigaDataXML.FirstChild,null,sw,null);
				experianXML.LoadXml(sw.ToString());

				expWrapper.CallExperian("UpgradeCC", experianXML.DocumentElement, ref responseXML, applicationNumber);

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
				omigaDataXML = null;
				experianXML = null;
				responseXML = null;
				cmd = null;
				param = null;
				sqlDR = null;
				ccNode = null;
				xslt = null;
				expWrapper = null;
			}
		}
		

		private void updateAddressData(ref XmlNode experianRequest, XmlNode targetingNode)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			XmlNodeList targetAddrList = null;
			XmlNode srcAddr = null;
			XmlNode dataItem = null;			
			string detailCode = string.Empty;
			string rmcCode = string.Empty;
			string regionCode = string.Empty;

			try
			{
				targetAddrList = targetingNode.SelectNodes("ADDRESSTARGET[@ADDRESSRESOLVED='Y']");
				foreach (XmlNode targetAddr in targetAddrList)
				{
					switch (_xmlMgr.xmlGetMandatoryAttributeText(targetAddr, "BLOCKSEQNUMBER"))
					{
						case "01":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[1]", true);
							break;
						case "11":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[2]", true);
							break;
						case "21":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[3]", true);
							break;
						case "02":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[4]", true);
							break;
						case "12":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[5]", true);
							break;
						case "22":
							srcAddr	= _xmlMgr.xmlGetMandatoryNode(experianRequest, "ADDR[6]", true);
							break;
					}

					// PSC 03/03/2006 MAR1348 - Start
					rmcCode = _xmlMgr.xmlGetAttributeText(targetAddr, "RMCCODE"); // EP688
					regionCode = _xmlMgr.xmlGetAttributeText(targetAddr, "REGIONCODE"); // EP688					
					
					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "DETAILCODE", true);
					
					// EP688 - Only populate the detail code if REGIONCODE and RMCCODE are populated
					if (rmcCode != "" && regionCode != "")
                        dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "DETAILCODE");

					dataItem = _xmlMgr.xmlGetNode(srcAddr, "DATAITEM[1]",false);
					if (dataItem == null)
					{
						dataItem = _xmlMgr.CreateXmlNode("DATAITEM", string.Empty);
						dataItem = srcAddr.OwnerDocument.ImportNode(dataItem, true);
						srcAddr.AppendChild(dataItem);
					}
					dataItem.InnerText = rmcCode;
						
					dataItem = _xmlMgr.xmlGetNode(srcAddr, "DATAITEM[2]",false);
					if (dataItem == null)
					{
						dataItem = _xmlMgr.CreateXmlNode("DATAITEM", string.Empty);
						dataItem = srcAddr.OwnerDocument.ImportNode(dataItem, true);
						srcAddr.AppendChild(dataItem);
					}
					dataItem.InnerText = regionCode;

					// PSC 03/03/2006 MAR1348 - End
						
					//LDM 13/04/2006 EP352
					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "FLAT", true); 
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "FLAT");

					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "HOUSENAME", true);
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "HOUSENAME");

					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "HOUSENUMBER", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "HOUSENUMBER");

					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "STREET", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "STREET");

					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "DISTRICT", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "DISTRICT");
					
					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "TOWN", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "TOWN");
					
					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "COUNTY", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "COUNTY");

					dataItem = _xmlMgr.xmlGetMandatoryNode(srcAddr, "POSTCODE", true);			
					dataItem.InnerText = _xmlMgr.xmlGetAttributeText(targetAddr, "POSTCODE");
				}
			}
			catch (Exception ex)
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
				targetAddrList = null;
				srcAddr = null;
				dataItem = null;
			}
		}
		
		private void SaveXMLCreditCheck(string functionID, XmlNode gatewayResponse, XmlNode methodRequest, 
			XmlNode omigaData, XmlNode targetAddresses)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
			}

			XslTransform xslt = new XslTransform();
			StringWriter sw = new StringWriter();
			XmlDocument transformedXML = new XmlDocument();
			SqlCommand cmd = null;
			SqlParameter param = null;
			XmlNode ccRootNode = null;
			XmlNode newNode = null;
			XmlNode newAttrib = null;
			XmlNode omigaNode = null;		
			XmlNode dec1RCDNode = null;
			XmlNode dataNode = null;
			XmlNode appNode = null;
			XmlNode customerNode = null;
			XmlNode additionalData = null;
			try
			{
				ccRootNode = gatewayResponse.SelectSingleNode("//REQUEST");

				additionalData = gatewayResponse.OwnerDocument.CreateElement("OMIGADATA");
		
				// Add the aplication and customer details to the response...
				appNode = _xmlMgr.xmlGetMandatoryNode(methodRequest,"//APPLICATION", true);
				newNode = gatewayResponse.OwnerDocument.ImportNode(appNode.CloneNode(true), true);
				newNode = additionalData.AppendChild(newNode);
				newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("USERID");
				newAttrib.InnerText = _xmlMgr.xmlGetAttributeText(methodRequest, "USERID");
				gatewayResponse.Attributes.SetNamedItem(newAttrib);
				customerNode = _xmlMgr.xmlGetMandatoryNode(omigaData, "CUSTOMER[1]", false);
				newNode = gatewayResponse.OwnerDocument.CreateElement("PRIMARYAPPLICANT");
				newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("CUSTOMERNUMBER");
				newAttrib.InnerText = _xmlMgr.xmlGetMandatoryAttributeText(customerNode, "CUSTOMERNUMBER");
				newNode.Attributes.SetNamedItem(newAttrib);
				newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("CUSTOMERVERSIONNUMBER");
				newAttrib.InnerText = _xmlMgr.xmlGetMandatoryAttributeText(
					_xmlMgr.xmlGetMandatoryNode(customerNode,"CUSTOMERVERSION", true), "CUSTOMERVERSIONNUMBER");
				newNode.Attributes.SetNamedItem(newAttrib);
				additionalData.AppendChild(newNode);
				customerNode = _xmlMgr.xmlGetMandatoryNode(omigaData, "CUSTOMER[2]", false);
				if (customerNode != null)
				{
					newNode = gatewayResponse.OwnerDocument.CreateElement("SECONDAPPLICANT");
					newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("CUSTOMERNUMBER");
					newAttrib.InnerText = _xmlMgr.xmlGetMandatoryAttributeText(customerNode, "CUSTOMERNUMBER");
					//MAR577
					newNode.Attributes.SetNamedItem(newAttrib);
					newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("CUSTOMERVERSIONNUMBER");
					newAttrib.InnerText = _xmlMgr.xmlGetMandatoryAttributeText(
						_xmlMgr.xmlGetMandatoryNode(customerNode,"CUSTOMERVERSION", true), "CUSTOMERVERSIONNUMBER");
					newNode.Attributes.SetNamedItem(newAttrib);
					additionalData.AppendChild(newNode);
				}

				// Add the Experian function that was called ...
				newNode = gatewayResponse.OwnerDocument.CreateElement("FUNCTION");
				newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("ID");
				newAttrib.InnerText = functionID;
				newNode.Attributes.SetNamedItem(newAttrib);
				additionalData.AppendChild(newNode);

				// If it exists, add the targetAddresses element passed into the method
				if (targetAddresses != null)
				{
					additionalData.AppendChild(additionalData.OwnerDocument.ImportNode(targetAddresses, true));
				}
				
				// Add the new element to the response XML ...	
				ccRootNode.AppendChild(additionalData);

				// Add a new node to the Experian data.  This will be used to store some additional
				// values required by Omiga ...
				omigaNode = gatewayResponse.SelectSingleNode("//OMIGADATA");
				
				// As System.XML doesn't support XSLT v2.0 I have to reconfigure
				// certain elements of the credit check XML.
				dec1RCDNode = ccRootNode.SelectSingleNode("//DEC1/REASONCODEDATA");
				char[] commaSep = {','};
				string[] separatedList = dec1RCDNode.InnerText.Split(commaSep);
				foreach (string reasonCode in separatedList)
				{
					newNode = gatewayResponse.OwnerDocument.CreateElement("CREDITCHECKREASONCODE");
					newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("REASONCODE");
					newAttrib.InnerText = reasonCode;
					newNode.Attributes.SetNamedItem(newAttrib);
					omigaNode.AppendChild(newNode);
				}

				// The date and time that the credit check was performed needs to be reformatted and passed
				string expTimeStamp = _xmlMgr.xmlGetMandatoryAttributeText(ccRootNode, "timestamp");
				char[] spaceSep = {' '};
				separatedList = expTimeStamp.Split(spaceSep);
				if (separatedList.Length >= 5)
				{
					// MAR1375 Display as 24hr time by leaving AM/PM on timestamp for conversion.
					string dateStamp = separatedList[1] + "-" + separatedList[2] + "-" + separatedList[3] +
						" " + separatedList[5].PadLeft(5,'0') + separatedList[6];
					newNode = gatewayResponse.OwnerDocument.CreateElement("CREDITCHECKDATE");
					newAttrib = gatewayResponse.OwnerDocument.CreateAttribute("DATE");
					newAttrib.InnerText = dateStamp;
					newNode.Attributes.SetNamedItem(newAttrib);
					omigaNode.AppendChild(newNode);
				}
				else
				{	
					throw new ApplicationException("Invalid credit check date stamp");
				}

				// Restructure the xml for Experian ...
				xslt.Load(GetXMLPath() + "\\" + "saveXMLCreditCheck.xslt");
				xslt.Transform(gatewayResponse,null,sw,null);				
				transformedXML.LoadXml(sw.ToString());

				// Save the data to the Omiga database ...
				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString))
				{
					cmd = new SqlCommand("USP_SAVEEXPERIANCREDITCHECK", conn);
					cmd.CommandType = CommandType.StoredProcedure;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//MISC", true);
					param = cmd.Parameters.Add("@p_Miscellaneous", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//CREDITCHECKREASONCODEBLOCK", true);
					param = cmd.Parameters.Add("@p_CreditCheckReason", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//CREDITCHECKLETTERCODEBLOCK", true);
					param = cmd.Parameters.Add("@p_CreditCheckLetter", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//CREDITCHECKSCOREBLOCK", true);
					param = cmd.Parameters.Add("@p_CreditCheckScore", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//BESPOKEDECISIONREASONCODESBLOCK", true);
					param = cmd.Parameters.Add("@p_BespokeDecisionReason", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//BXCUSTOMERBUREAU1BLOCK", true);
					param = cmd.Parameters.Add("@p_BxCustomerBureau1", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//BXCUSTOMERBUREAU2BLOCK", true);
					param = cmd.Parameters.Add("@p_BxCustomerBureau2", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//BXCUSTOMTERCAISBLOCK", true);
					param = cmd.Parameters.Add("@p_BxCustomerCAIS", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//DETECTRULES", true);
					param = cmd.Parameters.Add("@p_DetectResonCodes", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

					dataNode = _xmlMgr.xmlGetMandatoryNode(transformedXML.DocumentElement, "//BESPOKEBUREAUDATABLOCK", true);
					param = cmd.Parameters.Add("@p_BespokeBureau", SqlDbType.NText);
					param.Value = dataNode.OuterXml;

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
				xslt = null;
				sw.Flush();
				sw.Close();
				sw = null;
				cmd.Dispose();

				xslt = null;
				transformedXML = null;
				param = null;
				ccRootNode = null;
				newNode = null;
				newAttrib = null;
				omigaNode = null;		
				dec1RCDNode = null;
				dataNode = null;
				appNode = null;
				customerNode = null;
				additionalData = null;
			}
		}
		
		private void formatTargettingData(XmlNode inputNode, ref XmlNode responseNode)
		{
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
				_logger.Debug("inputNode: " + inputNode.OuterXml);
				_logger.Debug("responseNode: " + inputNode.OuterXml);
			}

			XmlDocument tempDoc = new XmlDocument();
			XmlNode newNode = null;
			XmlNode oldTargetingData = null;
			XslTransform xslt = new XslTransform();
			StringWriter sw = new StringWriter();
			try
			{
				if (inputNode != null)
				{	
					// Restructure the data ...
					xslt.Load(GetXMLPath() + "\\"  + "GenerateTargettingXML.xslt");
					xslt.Transform(inputNode,null,sw,null);
					tempDoc.LoadXml(sw.ToString());
					newNode = responseNode.OwnerDocument.ImportNode(tempDoc.DocumentElement, true);
					oldTargetingData = responseNode.SelectSingleNode("//TARGETINGDATA");
					if (oldTargetingData == null)
					{
						responseNode.AppendChild(newNode);
					}
					else
					{
						responseNode.ReplaceChild(newNode, oldTargetingData);
					}
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
				xslt = null;
				sw.Flush();
				sw.Close();
				sw = null;
				tempDoc = null;
				newNode = null;
				oldTargetingData = null;
				xslt = null;
			}		
		}
		
		private string GetXMLPath()
		{
			string thePath = _AppPath.ToUpper();
			return(thePath.Replace("\\DLLDOTNET", "\\XML"));
		}

		private void addExperianRefToRequest(ref XmlNode targetNode, XmlNode experianRefNode)
		{
			if (experianRefNode != null && targetNode != null)
			{
				string experianRef = experianRefNode.InnerText;
				XmlNode expNode = targetNode.OwnerDocument.CreateElement("EXP");
				XmlNode expRefNode = targetNode.OwnerDocument.CreateElement("EXPERIANREF");
				expRefNode.InnerText = experianRefNode.InnerText;
				expNode.AppendChild(expRefNode);
				targetNode.AppendChild(expNode);
			}
		}
	
		//  EP2_659 / MAR1896
		private void BlockThread()
		{
			//Read delay in milliseconds from registry. Default value = 1000ms
			string sDelay;
			int nWaitTime;

			Microsoft.Win32.RegistryKey omExperianDelayKey;
		
			sDelay = "1000";

			try
			{
				omExperianDelayKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Omiga4\\System Configuration\\Stubs\\omExperianCC", false);
					
				if (omExperianDelayKey != null)
				{
					sDelay = (string)omExperianDelayKey.GetValue("Delay");
				}
			}
			catch
			{
			}

			//Delay for the given time
			nWaitTime = Convert.ToInt16(sDelay);
			TimeSpan waitTime = new TimeSpan(0,0,0,0,nWaitTime);

			System.Threading.Thread.Sleep(waitTime);		
		}
	}
} 

/*
--------------------------------------------------------------------------------------------
Workfile:			ExperianHIBO.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MV		26/10/2005	MAR278 Included logging 
HMA     08/11/2005  MAR339 Add extra error checking.
INR		15/11/2005	MAR533 Need to pass OTHERSYSTEMCUSTOMERNUMBER. Added getCommunicationChannel.
GHun	13/12/2005	MAR852 set Operator in generic request
PSC		13/12/2005	MAR457 Use omLogging wrapper
PSC		26/01/2006	MAR1133 Amend RunExperianHI to default Customer Number to NA

------------------------------ CUT FOR EPSOM -----------------------------------------------

PCT		06/03/2006  EP2 Amended to call Experian not the MARS gateway
SAB		14/03/2006  EP17 Amended to call full Experian wrapper
PSC		29/03/2006	MAR1554 Amend GetHunterRequestData so it doesn't enlist in any
					encompassing transaction when getting the request data			
SAB		28/04/2006	EP346 Revised call for direct hunter call
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
using Vertex.Fsd.Omiga.omDg;
using Vertex.Fsd.Omiga.omExperianHI.FraudCheckService;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

using Vertex.Fsd.Omiga.omExperianWrapper; //pct 06/03/2006	EP2 Epsom

namespace Vertex.Fsd.Omiga.omExperianHI
{
	/// <summary>
	/// Summary description for ExperianHIBO.
	/// </summary>
	/// 
	
	//	[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
	//	public interface IExperianHIBO
	//	{
	//		//[DispId(600)]
	//		string RunExperianHI(string strRequest);
	//	}
	[Guid("F928E364-7CE4-490b-9DF0-96D302D7D23A")]
	[ComVisible(true)]
	[ProgId("omExperianHI.ExperianHIBO")]
	//[ClassInterface(ClassInterfaceType.AutoDual)]
	public class ExperianHIBO //: IExperianHIBO
	{
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		private string _AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

		//private DirectGatewayBO dg = new DirectGatewayBO();
		//private string _AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
		
		// PSC 13/12/2005 MAR457
		//private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public ExperianHIBO()
		{
		}

		public string RunExperianHI(string strRequest)
		{
			ExperianWrapperBO expWrapper = new ExperianWrapperBO();
			XmlDocumentEx requestXML = new XmlDocumentEx();
			XmlDocumentEx responseXML = new XmlDocumentEx();
			XmlDocumentEx hunterRequestXML = new XmlDocumentEx();
			XmlDocument hunterResponseXML = new XmlDocument();			
			XmlElementEx requestNode = null;	
			XmlElementEx requestAppNode = null;				
			XmlElementEx hunterReqNode = null;
			XmlNode hunoNode = null;
			XmlNode tempNode = null;
			XmlNode typeAttrib = null;
			XmlNode responseNode = null;
			XmlNode experianRefNode = null;
			XmlNode hunterReturnNode = null;
			XslTransform xslt = null;
			StringWriter sw = null;

			SqlCommand cmd = null;
			SqlParameter param = null;
			SqlDataReader sqlDR = null;

			string strResponse = String.Empty;
			
			try
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name + "- Request XML: " + strRequest);
				}
				
				requestXML.LoadXml(strRequest);
				xslt = new XslTransform();
				sw = new StringWriter();
				expWrapper = new ExperianWrapperBO();
			
				// Extract parameters from the request
				requestNode = (XmlElementEx) requestXML.SelectMandatorySingleNode("//REQUEST");
				requestAppNode = (XmlElementEx) requestNode.SelectMandatorySingleNode("APPLICATION");
				string userID = requestNode.GetAttribute("USERID");
				string password = requestNode.GetAttribute("PASSWORD");
				string applicationNumber = requestAppNode.GetAttribute("APPLICATIONNUMBER");
				string applicationFactFindNumber = requestAppNode.GetAttribute("APPLICATIONFACTFINDNUMBER");
					
				if(_logger.IsDebugEnabled) 
				{
					if (applicationNumber != null && applicationNumber.Length > 0)
					{
						omLogging.ThreadContext.Properties["ApplicationNumber"] = applicationNumber;	
					}
				}

				// Extract the relevant data from Omiga
				using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString + "Enlist=false;"))
				{
					cmd = new SqlCommand("USP_GetHunterRequestData", conn);
					cmd.CommandType = CommandType.StoredProcedure;

					param = cmd.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, applicationNumber.Length);
					param.Value = applicationNumber;
					param = cmd.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					param.Value = Convert.ToInt16(applicationFactFindNumber, 10);
					
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
					hunterRequestXML.LoadXml(spXMLresponse.ToString());			
				}
				
				// Restructure the xml for hunter ...
				xslt.Load(GetXMLPath() + "\\" + "hunterrequest.xslt");
				xslt.Transform(hunterRequestXML.FirstChild,null,sw,null);
				hunterRequestXML.LoadXml(sw.ToString());
				
				// Call Experian
				hunterReqNode = (XmlElementEx) hunterRequestXML.SelectMandatorySingleNode("//EXPERIAN");
				expWrapper.CallExperian("hunter", hunterReqNode, ref hunterResponseXML, applicationNumber);

				// Save the response to the Omiga database
				hunoNode = hunterResponseXML.SelectSingleNode("//HUNO");
				
				if (hunoNode == null)
				{
					throw new OmigaException("HUNO element not returned by Experian.");
				}
				else
				{
					experianRefNode = hunoNode.SelectSingleNode("EXPERIANREF");
					hunterReturnNode = hunoNode.SelectSingleNode("RETURNCODE");

					if (experianRefNode == null || hunterReturnNode == null)
					{
						throw new OmigaException("The hunter response does not contain EXPERIANREF or RETURNCODE.");
					}
					else
					{
						using (SqlConnection conn = new SqlConnection(Global.DatabaseConnectionString))
						{
							cmd = new SqlCommand("USP_SetHunterDataReply", conn);
							cmd.CommandType = CommandType.StoredProcedure;

							param = cmd.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, applicationNumber.Length);
							param.Value = applicationNumber;
							param = cmd.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
							param.Value = Convert.ToInt16(applicationFactFindNumber, 10);
							param = cmd.Parameters.Add("@p_ExperianRef", SqlDbType.NVarChar, 10);
							param.Value = experianRefNode.InnerText;
							param = cmd.Parameters.Add("@p_ReturnCode",SqlDbType.NVarChar, 4);
							param.Value = hunterReturnNode.InnerText;
							param = cmd.Parameters.Add("@p_UserId", SqlDbType.NVarChar, 10);
							param.Value = userID; 
						
							conn.Open();
							cmd.ExecuteNonQuery();
						}

						// Create the response node
						tempNode = responseXML.CreateElement("RESPONSE");
						responseNode = responseXML.AppendChild(tempNode);	
						typeAttrib = responseXML.CreateAttribute("TYPE");
						typeAttrib.Value = "SUCCESS";
						responseNode.Attributes.SetNamedItem(typeAttrib);
						
						tempNode = responseXML.ImportNode(hunoNode, true);
						responseNode.AppendChild(tempNode);

						strResponse = responseXML.OuterXml;				
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
				// Create the response node ...
				strResponse = new OmigaException(ex.Message).ToOmiga4Response();
			}
			finally
			{
				expWrapper = null;
				requestXML = null;
				responseXML = null;
				hunterRequestXML = null;
				hunterResponseXML = null;			
				requestNode = null;	
				requestAppNode = null;				
				hunterReqNode = null;
				hunoNode = null;
				tempNode = null;
				typeAttrib = null;
				responseNode = null;
				experianRefNode = null;
				hunterReturnNode = null;
				xslt = null;
				sw = null;
				cmd = null;
				param = null;
				sqlDR = null;
			}
			return strResponse;
		}

		private string GetXMLPath()
		{
			string thePath = _AppPath.ToUpper();
			return(thePath.Replace("\\DLLDOTNET", "\\XML"));
		}
	}
}

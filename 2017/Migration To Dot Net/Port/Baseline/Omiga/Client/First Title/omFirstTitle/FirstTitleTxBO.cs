/*
--------------------------------------------------------------------------------------------
Workfile:			FirstTitleTxBO.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		Class to require a transaction
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		09/03/2006  MAR1392 Created
RF		09/03/2006  MAR1392 Ensure link to Clearquest
--------------------------------------------------------------------------------------------
*/
using System;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;  
using XMLManager = Vertex.Fsd.Omiga.omFirstTitle.XML.XMLManager;


namespace Vertex.Fsd.Omiga.omFirstTitle
{
	/// <summary>
	/// This class is used to require a transaction
	/// </summary>

	[ProgId("omFirstTitle.FirstTitleTxBO")]
	[ComVisible(true)]
	[Guid("1DAF4130-E9D8-4d68-A0DE-DDD30DCE373E")]
	[Transaction(TransactionOption.Required)]
	public class FirstTitleTxBO : ServicedComponent
	{
		private  omLogging.Logger _logger = null;

		public FirstTitleTxBO()
		{
		}
		
		public void SaveAppFirstTitleOutbound(XmlNode xnAppFirstTitle, string processContext)
		{
			const string cFunctionName = "SaveAppFirstTitleOutbound";

			_logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, 
				processContext);

			if(_logger.IsDebugEnabled)
			{
				_logger.Debug("Starting " + cFunctionName);
				//_logger.Debug("Transaction ID: " + ContextUtil.TransactionId.ToString()); 
			}

			try 
			{
				XMLManager xmlMgr = new XMLManager();

				//Check the mandatory Attributes
				string ApplicationNumber =
					xmlMgr.xmlGetMandatoryNodeText(xnAppFirstTitle, "APPLICATIONNUMBER");
				
				string MessageType = xmlMgr.xmlGetMandatoryNodeText(xnAppFirstTitle, "MESSAGETYPE");

				string MessageSubType =	xmlMgr.xmlGetNodeText(xnAppFirstTitle, "MESSAGESUBTYPE");

				string ConnectionString = Global.DatabaseConnectionString;
				
				using(SqlConnection oConn = new SqlConnection(ConnectionString))
				{
					SqlCommand oComm = new SqlCommand("USP_SAVEAPPFIRSTTITLEDATA", oConn);
					oComm.CommandType = CommandType.StoredProcedure;

					//ApplicationNumber
					SqlParameter oParam = oComm.Parameters.Add("@p_ApplicationNumber", SqlDbType.NVarChar, ApplicationNumber.Length);
					oParam.Value = ApplicationNumber;

					//ApplicationFactFindNumber
					oParam = oComm.Parameters.Add("@p_ApplicationFactFindNumber", SqlDbType.Int);
					oParam.Value = 1;

					//MessageType
					oParam = oComm.Parameters.Add("@p_MessageType", SqlDbType.NVarChar, MessageType.Length);
					if (MessageType.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MessageType;
				
					//MessageSubType
					oParam = oComm.Parameters.Add("@p_MessageSubType", SqlDbType.NVarChar, MessageSubType.Length);
					if (MessageSubType.Length == 0)
						oParam.Value = DBNull.Value;
					else
						oParam.Value = MessageSubType;
				
					//ConveyancerRef 
					oParam = oComm.Parameters.Add("@p_ConveyancerRef", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//HubRef
					oParam = oComm.Parameters.Add("@p_HubRef", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//DateTime
					oParam = oComm.Parameters.Add("@p_DateTime", SqlDbType.DateTime);
					oParam.Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
				
					//ReasonForRejection 
					oParam = oComm.Parameters.Add("@p_ReasonForRejection", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//TitleNumber1
					oParam = oComm.Parameters.Add("@p_TitleNumber1", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//TitleNumber2
					oParam = oComm.Parameters.Add("@p_TitleNumber2", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//TitleNumber3
					oParam = oComm.Parameters.Add("@p_TitleNumber3", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//RedemptionStatementExpDate
					oParam = oComm.Parameters.Add("@p_RedempStatementExpiryDate", SqlDbType.DateTime);
					oParam.Value = DBNull.Value;

					//CompletionDate
					oParam = oComm.Parameters.Add("@p_CompletionDate", SqlDbType.DateTime);
					oParam.Value = DBNull.Value;
				
					//MessageText
					oParam = oComm.Parameters.Add("@p_MessageText", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					//MortgageAdvance
					oParam = oComm.Parameters.Add("@p_MortgageAdvance", SqlDbType.NVarChar);
					oParam.Value = DBNull.Value;
				
					oConn.Open();
				
					int iRowsAffected = oComm.ExecuteNonQuery();
				
					if (iRowsAffected <= 0)
					{
						if(_logger.IsWarnEnabled) 
						{
							_logger.Warn(cFunctionName + " error: Error in inserting data");
						}
						throw new OmigaException("Error in " + cFunctionName + ": failed to save data");
					}
				}
			}
			catch(SqlException ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("SqlException in " + cFunctionName, ex);
					for (int i=0; i < ex.Errors.Count; i++)
					{
						_logger.Warn("Index #" + i + 
							" Error: " + ex.Errors[i].ToString());
					}
				}
				ContextUtil.SetAbort();
				throw new OmigaException(ex);
			}
			catch(Exception ex)
			{
				if(_logger.IsWarnEnabled) 
				{
					_logger.Warn("Error in " + cFunctionName, ex);
				}
				ContextUtil.SetAbort();
				throw new OmigaException(ex);
			}
			ContextUtil.SetComplete();
		}
	}
}

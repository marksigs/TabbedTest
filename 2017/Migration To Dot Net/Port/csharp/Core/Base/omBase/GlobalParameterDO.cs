/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Global parameters data object
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		20/07/1999  Created
MCS		02/09/1999  GetMaximumPasswords added
RF		29/09/1999  Added GetCurrentParameterByType
RF		22/12/1999  Fix GetCurrentParameter.
Improve performance of GetCurrentParameterByType.
RF		08/05/2000  Improved error handling.
MS		23/06/2000  Removed additional un-needed timing output
APS		27/07/2000  Stress testing errors in retrieving from SPM
MC		07/08/2000  SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
PSC		11/08/2000  SYS1430 Back out SYS1409
LD		07/11/2000  Explicity close recordsets
LD		07/11/2000  Explicity destroy command objects
APS		29/11/2000  CORE000022
LD		19/06/2001  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
DM		15/10/2001  SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
AS		13/11/2003  CORE1 Removed GENERIC_SQL.
TK		22/11/2004  BBG1821 - Performance related fixes.
AS		22/03/2006  CORE175: BMIDS858 SQL queries for getting global parameters should be improved
PSC		26/06/2006  CORE281 - Amend GetCurrentParameterByType to return correct value if column is empty
--------------------------------------------------------------------------------------------
BMIDS History:
Prog	Date		Description
SR		21/11/02	BMIDS01050 - Modified method 'GetSearchString'
MDC		10/03/2003  BM0493 Performance Enhancements
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		25/07/2007	First .Net version. Ported from GlobalParameterDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.GlobalParameterDO")]
	[Guid("2761194A-5C2C-438C-9C14-2EF0FA34C200")]
	[Transaction(TransactionOption.Supported)]
	public class GlobalParameterDO : ServicedComponent
	{
		private const string _tableName = "GLOBALPARAMETER";

		// header ----------------------------------------------------------------------------------
		// description:
		// Get the data for a single instance of the persistant data associated with
		// this data object, selected the current instance of the data as determined
		// by GLOBALPARAMETERSTARTDATE
		// pass:
		// strParamName
		// name of the parameter to be retrieved
		// return:
		// string containing XML data stream representation of data retrieved
		// ------------------------------------------------------------------------------------------
		// RF 22/12/99 Fix and improve performance.
		// APS 29/11/2000 : CORE000022 Added SPM functionality
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameter(string name)
		{
			string currentParameterXml = null;

			try
			{
				currentParameterXml = new GlobalParameter(name).ToString();
				ContextUtility.SetComplete();
			}
			catch (InvalidOperationException)
			{
				throw new RecordNotFoundException(TransactionAction.SetOnErrorType);
			}
			catch (ErrAssistException exception)
			{
				exception.SetOnErrorType();
				throw;
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return currentParameterXml;
		}

		public string GetCurrentParameterByType(string name, string fieldName)
		{
			string parameterValue = "";

			// APS 29/11/2000 : CORE000022
			if (fieldName.Length == 0)
			{
				throw new MissingParameterException("GlobalParameter Field Name");
			}

			// AS 22/03/2006 CORE175
			XmlDocument xmlDoc = new XmlDocument();
			xmlDoc.LoadXml(GetCurrentParameter(name));
			XmlNode xmlNode = xmlDoc.SelectSingleNode(".//" + fieldName.ToUpper());
			if (xmlNode != null)
			{
				parameterValue = xmlNode.InnerText;
			}
			// AS 22/03/2006 CORE175 End

			return parameterValue;
		}

		// BM0493 MDC 10/04/2003
		// header ----------------------------------------------------------------------------------
		// description:
		// Get the current global parameter values for a list of parameter names.
		// pass:
		// strParamName
		// name of the parameter to be retrieved
		// return:
		// string containing XML data stream representation of data retrieved
		// NB:
		// This method contains SQL Server specific code to retrieve XML data!
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameterListEx(string request)
		{
			string response = "";

			try
			{
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(request);
				string cmdText = "SELECT NAME, MAXIMUMAMOUNT, PERCENTAGE, AMOUNT, BOOLEAN, STRING FROM GLOBALPARAMETER";
				string sqlWhere = "";
				XmlNodeList xmlNameList = xmlDoc.SelectNodes(".//GLOBALPARAMETER/NAME");
				if (xmlNameList.Count > 0)
				{
					sqlWhere += " WHERE NAME IN (";
					int nameCount = 0;
					foreach (XmlNode xmlName in xmlNameList)
					{
						if (nameCount > 0)
						{
							sqlWhere += ", ";
						}
						sqlWhere += "'" + xmlName.InnerText + "'";
						nameCount++;
					}
					sqlWhere += ") AND";
				}
				else
				{
					sqlWhere = " WHERE";
				}
				sqlWhere += " GLOBALPARAMETERSTARTDATE = (SELECT MAX(GLOBALPARAMETERSTARTDATE)";
				sqlWhere += " from GLOBALPARAMETER GP2 WHERE GP2.NAME = GLOBALPARAMETER.NAME ";
				sqlWhere += " AND GLOBALPARAMETERSTARTDATE <= GETDATE() GROUP BY NAME) ";
				cmdText += sqlWhere + "FOR XML AUTO, ELEMENTS";
				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
				{
					using (SqlCommand sqlCommand = new SqlCommand(cmdText, sqlConnection))
					{
						sqlConnection.Open();
						using (XmlReader xmlReader = sqlCommand.ExecuteXmlReader())
						{
							xmlReader.Read();

							while (xmlReader.ReadState != ReadState.EndOfFile)
							{
								response += xmlReader.ReadOuterXml();
							}
						}
					}
				}
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}

		/*
		// AS 02/08/2007 No longer used - replaced by GlobalParameter.ToXmlString
		 
		// header ----------------------------------------------------------------------------------
		// description:
		// return a value from enumerated list (DBDATATYPE in SQLAssist) which indicates
		// the data 'type' for the value associated with a particular (xml) element
		// pass:
		// elementName the xml element name
		// return:
		// GetDbType	   enumerated value
		// Raise Errors:	 n/a
		// ------------------------------------------------------------------------------------------
		private static DBDATATYPE GetDbType(string columnName)
		{
			DBDATATYPE dbDataType = DBDATATYPE.dbdtNotStored;
			switch (columnName)
			{
				case "NAME":
				case "DESCRIPTION":
				case "STRING":
					dbDataType = DBDATATYPE.dbdtString;
					break;
				case "AMOUNT":
				case "MAXIMUMAMOUNT":
				case "PERCENTAGE":
					dbDataType = DBDATATYPE.dbdtDouble;
					break;
				case "BOOLEAN":
					dbDataType = DBDATATYPE.dbdtBoolean;
					break;
				case "GLOBALPARAMETERSTARTDATE":
					dbDataType = DBDATATYPE.dbdtDate;
					break;
				default:
					dbDataType = DBDATATYPE.dbdtNotStored;
					break;
			}
			return dbDataType;
		}
		*/

		/* AS 26/07/2007 These methods are not used?
		 
		// header ----------------------------------------------------------------------------------
		// description:
		// create an instance of the persistant data associated with this data object
		// for each set of data in the request
		// pass:
		// request  xml Request data stream containing data to be persisted
		// ------------------------------------------------------------------------------------------
		// XML Request Format:
		// <GLOBALPARAMETER>
		// <NAME>string</NAME>
		// <GLOBALPARAMETERSTARTDATE>date</GLOBALPARAMETERSTARTDATE>
		// <DESCRIPTION>string</DESCRIPTION>
		// <AMOUNT>double</AMOUNT>
		// <MAXIMUMAMOUNT>double</MAXIMUMAMOUNT>
		// <PERCENTAGE>double</PERCENTAGE>
		// <BOOLEAN>boolean</BOOLEAN>
		// <STRING>string</STRING>
		// </GLOBALPARAMETER>
		// ------------------------------------------------------------------------------------------
		public void Create(string request) 
		{
			try 
			{
				XmlDocument xmlDocument = XmlAssist.Load(request);
				if (xmlDocument.GetElementsByTagName(_tableName).Count != 1)
				{
					throw new MissingPrimaryTagException();
				}
				XmlElement xmlDataElement = (XmlElement)xmlDocument.GetElementsByTagName(_tableName)[0];
				// GetKeyString will raise an error if primary key(s) not found
				GetKeyString(request);
				// Build the SQL insert string for this row.
				// Extra processing will be required if sequence numbers are involved.
				// 
				// e.g.
				// strCustomerNumber = xmlNodeList.Item(intRow - 1).ChildNodes(0).Text
				// e.g
				// cmdText = "INSERT INTO ADDRESS "
				// cmdText = cmdText & "(ADDRESSGUID, BUILDINGORHOUSENAME, BUILDINGORHOUSENUMBER, MAILSORTCODE) VALUES ("
				// cmdText = cmdText & SqlAssist.FormatGuid(CreateGUID())
				// cmdText = cmdText & ", 'MSG Towers'"
				// cmdText = cmdText & ", '66'"
				// cmdText = cmdText & ", 123456)"
				// *****************************************************************************
				// SQL insert statement
				// SQL insert statement, column list
				// SQL insert statement, value list

				bool hasData = false;
				string sqlColumns = "";
				string sqlValues = "";
				foreach (XmlNode xmlChildNode in xmlDataElement.ChildNodes)
				{
					DBDATATYPE eDbDataType = GetDbType(xmlChildNode.Name);
					if (eDbDataType != DBDATATYPE.dbdtNotStored)
					{
						if (hasData)
						{
							sqlColumns = sqlColumns + ", ";
							sqlValues += ", ";
						}
						hasData = true;
						sqlColumns = sqlColumns + xmlChildNode.Name;
						switch (eDbDataType)
						{
							case DBDATATYPE.dbdtGuid:
								sqlValues += SqlAssist.FormatGuid(xmlChildNode.InnerText, GUIDSTYLE.guidLiteral);
								break;
							case DBDATATYPE.dbdtString:
								sqlValues += SqlAssist.FormatString(xmlChildNode.InnerText, false);
								break;
							case DBDATATYPE.dbdtInt:
							case DBDATATYPE.dbdtBoolean:
							case DBDATATYPE.dbdtComboId:
							case DBDATATYPE.dbdtCurrency:
							case DBDATATYPE.dbdtDouble:
								sqlValues += xmlChildNode.InnerText;
								break;
							case DBDATATYPE.dbdtDate:
								sqlValues += SqlAssist.FormatDateString(xmlChildNode.InnerText, DATETIMEFORMAT.dtfDateTime);
								break; 
						}
					}
				}
				if (!hasData)
				{
					throw new NoDataForCreateException();
				}
				string cmdText = "INSERT INTO " + _tableName + " (" + sqlColumns + ") VALUES (" + sqlValues + ")";
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetAbort);
			}
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Update a single instance of the persistant data associated with this
		// data object
		// pass:
		// request  xml Request data stream containing data to be persisted
		// ------------------------------------------------------------------------------------------
		public void Update(string request)
		{
			try
			{
				XmlDocument xmlDocument = XmlAssist.Load(request);
				if (xmlDocument.GetElementsByTagName(_tableName).Count != 1)
				{
					throw new MissingPrimaryTagException();
				}
				XmlElement xmlDataElement = (XmlElement)xmlDocument.GetElementsByTagName(_tableName)[0];
				// build the key string
				string keyString = GetKeyString(request);
				// build the full SQL string
				// SQL insert statement
				// SQL insert statement, column list

				bool hasData = false;
				string sqlSets = "";
				foreach (XmlNode xmlChildNode in xmlDataElement.ChildNodes)
				{
					// do not update primary keyString
					if (xmlChildNode.Name != "NAME" &&
						xmlChildNode.Name != "GLOBALPARAMETERSTARTDATE")
					{
						DBDATATYPE eDbDataType = eDbDataType = GetDbType(xmlChildNode.Name);
						if (eDbDataType != DBDATATYPE.dbdtNotStored)
						{
							if (hasData)
							{
								sqlSets += ", ";
							}
							hasData = true;
							sqlSets += xmlChildNode.Name + " = ";
							switch (eDbDataType)
							{
								case DBDATATYPE.dbdtGuid:
									sqlSets += SqlAssist.FormatGuid(xmlChildNode.InnerText, GUIDSTYLE.guidLiteral);
									break;
								case DBDATATYPE.dbdtString:
									sqlSets += SqlAssist.FormatString(xmlChildNode.InnerText, false);
									break;
								case DBDATATYPE.dbdtInt:
								case DBDATATYPE.dbdtDouble:
								case DBDATATYPE.dbdtCurrency:
								case DBDATATYPE.dbdtComboId:
									if (xmlChildNode.InnerText.Length == 0)
									{
										sqlSets += "Null";
									}
									else
									{
										sqlSets += xmlChildNode.InnerText;
									}
									break;
							}
						}
					}
				}
				if (!hasData)
				{
					throw new MissingPrimaryTagException();
				}
				string cmdText = "UPDATE " + _tableName + " SET " + sqlSets + " WHERE " + keyString;

				// do the update
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetAbort);
			}
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Delete a single instance of the persistant data associated with this
		// data object
		// pass:
		// request  xml Request data stream containing data to which identifies
		// the instance to be deleted
		// Raise Errors:
		// omiga4RecordNotFound
		// omiga4InvalidKeyString
		// parser errors
		// ------------------------------------------------------------------------------------------
		public void Delete(string request)
		{
			try
			{
				string keyString = GetKeyString(request);
				// check the record exists on the database
				AdoAssist.CheckSingleRecordExists(_tableName, keyString);
				// build the full SQL string
				string cmdText = "delete from " + _tableName + " where " + keyString;
				// do the delete
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetAbort);
			}
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// HasValue all keyString are present in the XML, and return them with their
		// associated values in a format ready for including in a SQL execution
		// string.
		// Raise an error if not all keyString have values specified.
		// pass:
		// request			  xml data
		// return:
		// A string in the format "TableName.KeyFieldA = ValueX and
		// TableName.KeyFieldB = ValueY and TableName.KeyFieldC = ValueZ"
		// Raise Errors:
		// omiga4InvalidKeyString
		// parser errors
		// ------------------------------------------------------------------------------------------
		private static string GetKeyString(string request)
		{
			XmlDocument xmlDocument = XmlAssist.Load(request);
			if (xmlDocument.GetElementsByTagName("NAME").Count == 0 ||
				xmlDocument.GetElementsByTagName("GLOBALPARAMETERSTARTDATE").Count == 0 ||
				xmlDocument.GetElementsByTagName("NAME")[0].InnerText.Length == 0 ||
				xmlDocument.GetElementsByTagName("GLOBALPARAMETERSTARTDATE")[0].InnerText.Length == 0)
			{
				throw new InvalidKeyStringException();
			}
			string keyString = _tableName + ".NAME = " +
				SqlAssist.FormatString(xmlDocument.GetElementsByTagName("NAME")[0].InnerText, false) +
				" AND " + _tableName + ".GLOBALPARAMETERSTARTDATE = " +
				SqlAssist.FormatDate(Convert.ToDateTime(xmlDocument.GetElementsByTagName("GLOBALPARAMETERSTARTDATE")[0].InnerText), DATETIMEFORMAT.dtfDate);

			return keyString;
		}
		*/

		protected override bool CanBePooled()
		{
			return true;
		}
	}

}

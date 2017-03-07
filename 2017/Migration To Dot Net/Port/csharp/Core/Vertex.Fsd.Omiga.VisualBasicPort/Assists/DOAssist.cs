/*
--------------------------------------------------------------------------------------------
Workfile:			DOAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		.Net version of DOAssist.cls.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MCS		26/08/1999	Initial Creation
MCS		24/09/1999  Modification following code review
AY		13/10/1999  FindList changed to allow an "order by" setting
 					Improved error message for incorrect combination of parameters to BuildSQLString
RF		20/10/1999  Improved error handling in GetDbType
MCS		19/11/1999  Findlist/BuildSQLString modified. To use the specific fields needed
					to be returned in the SQL search
PSC		15/12/1999  Make the xml node obptional on GetXMLFromRecordsetEx and
					GetXMLFromTableSchema. Remove node from GetDataEx and FindListEx
PSC		17/12/1999  Put all Ex functions onto new interface and remove Ex from name
PSC		20/12/1999  Amend to use new IADOAssist interface
JLD		07/12/2000  Changed GetNextSequenceNumberEx to use GetElementsByTagName()
					instead of selectNodes() to guarantee the node being found anywhere in the xml.
RF		25/01/2000  GetNextSequenceNumberEx put in interface as GetNextSequenceNumber.
PSC		01/02/2000  Don't clone nodes on return
SR		03/02/2000  New method - GetComponentData
SR		14/03/2000  New parameters to method - FindListMultiple
MH		05/04/2000  Added more diag information to GetXMLFromTableSchema
MS		23/06/2000  Removed additional un-needed timing output
LD		07/11/2000  Explicitly close recordsets
LD		07/11/2000  Explicitly destroy command objects
LD		29/11/2000  Explicitly close recordsets
PSC		23/01/2001  SYS1871 - Amend GetNextSequenceNumber to close the recordset once and also
					to close it in the error handler
DM		15/10/2001	SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
AS		13/11/2003  CORE1 Removed GENERIC_SQL.
SDS		20/01/2004  LIVE00009653 - NOLOCK is used only for SQLServer sql statements
--------------------------------------------------------------------------------------------
BMIDS History:

Prog	Date		Description
SR		02/01/2003  BM0209 - Modified methods 'FindListEx' and 'IDOAssist_FindList'
SR		08/01/2003  BM0209 - Modified methods 'FindListEx' and 'IDOAssist_FindList' (Added
					NOLOCK hint to SQl executed, whereever required.
MDC		14/01/2003  BM0249 - Handle SQLNOLOCK in GetXMLForTableSchema
RF		12/05/2003  BM0536 - Handle SQLNOLOCK in IDOAssist_FindListMultiple
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		23/08/2004  BBG1211 - omApp.ApplicationManagerBO.ImportAccountsIntoApplication (extracted from BMIDS788)
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		14/05/2007	First .Net version. Ported from DOAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

// Ambiguous reference in cref attribute.
#pragma warning disable 419

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Constants defining the different types of formatting for parts of a SQL string.
	/// </summary>
	public enum SQL_FORMAT_TYPE
	{
		/// <summary>
		/// Fields and values in two separate strings e.g. "FieldA,FieldB" and "Value1,Value2".
		/// </summary>
		sftFieldValueSeparated,
		/// <summary>
		/// Field=value pairs separated by commas, e.g. "FieldA = Value1, FieldB = Value2".
		/// </summary>
		sftCommaSeparated,
		/// <summary>
		/// Field=value pairs separated by " AND ", e.g. "FieldA = Value1 AND FieldB = Value2".
		/// </summary>
		sftAndSeparated
	}

	/// <summary>
	/// Constants defining which types of fields in the specified schema will be used when formatting a SQL string.
	/// </summary>
	public enum CLASS_DEF_KEY
	{
		/// <summary>
		/// Use all the primary key fields for the specified table.
		/// </summary>
		cdkPRIMARYKEY,
		/// <summary>
		/// Use all the non-primary key fields for the specified table.
		/// </summary>
		cdkOTHERS
	}

	/// <summary>
	/// Constants defining which key values are expected in the supplied xml data.
	/// </summary>
	public enum CLASS_DEF_KEY_AMOUNT
	{
		/// <summary>
		/// All key values are expected. A <see cref="MissingElementException"/> is thrown if a key value is missing.
		/// </summary>
		cdkaALLKEYS,
		/// <summary>
		/// Any of the key values are expected.
		/// </summary>
		cdkaANYKEYS
	}

	/// <summary>
	/// Database Object help methods.
	/// </summary>
	public static class DOAssist
	{
		/// <summary>
		/// Gets the data for a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="request">Xml request data stream which identifies the instance of the persistant data to be retrieved.</param>
		/// <param name="classDefinition">Xml class definition to parse <paramref name="request"/> against.</param>
		/// <returns>
		/// String containing an xml data stream representation of the data retrieved.
		/// </returns>
		public static string GetData(string request, string classDefinition)
		{
			string data = "";

			try
			{
				string tableName = "";
				string fieldValuePair = "";
				string fields = null;
				string values = null;
				string select = "";
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select);
				// build the full select statement
				string cmdText = "select * from " + tableName + " where " + fieldValuePair;
				WriteSQLToFile(cmdText);

				using (DataSet dataSet = new DataSet())
				{
					using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
					{
						using (SqlCommand sqlCommand = new SqlCommand(cmdText, sqlConnection))
						{
							sqlCommand.CommandType = CommandType.Text;

							using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
							{
								sqlDataAdapter.Fill(dataSet);
							}
						}
					}

					if (dataSet.Tables != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows != null && dataSet.Tables[0].Rows.Count > 0)
					{
						data = GetXMLFromRecordSet(GetDataRow(dataSet), classDefinition);
					}
					else
					{
						// raise application error to be interpreted by calling object
						throw new RecordNotFoundException();
					}
				}

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return data;
		}

		/// <summary>
		/// Get the data for a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> object containing the retrieved data.
		/// </returns>
		/// <remarks>
		/// Data is retrieved from the table named by the inner text of the first child element of 
		/// the TABLENAME element in <paramref name="xmlClassDefinition"/>. The same table name is 
		/// used to identify the node in <paramref name="xmlRequest"/> that is used when retrieving the data.
		/// </remarks>
		public static XmlNode GetDataEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return GetDataEx(xmlRequest, xmlClassDefinition, null, null);
		}

		/// <summary>
		/// Get the data for a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <param name="tableName">The table from which to retrieve the data. If null then defaults to the inner text of the first child element of the TABLENAME element in <paramref name="xmlClassDefinition"/>.</param>
		/// <param name="itemName">The node in <paramref name="xmlRequest"/> to use when retrieving the data. If null then defaults to <paramref name="tableName"/>.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> object containing the retrieved data.
		/// </returns>
		public static XmlNode GetDataEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string tableName, string itemName)
		{
			XmlNode xmlData = null;

			try
			{
				// The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
				if (itemName == null || itemName.Length == 0)
				{
					itemName = tableName;
				}
				// DM 14/11/02 BMIDS00935 BEGIN
				string sqlNoLock = "";
				if (xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK") != null)
				{
					if (AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
					{
						// SDS LIVE00009653 / 20/01/2004
						sqlNoLock = " with (nolock)";
					}
					xmlClassDefinition.DocumentElement.RemoveChild(xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK"));
				}
				// DM 14/11/02 BMIDS00935 END
				string fieldValuePair = "";
				string fields = null;
				string values = null;
				string select = "";
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select, itemName);
				// build the full select statement
				// DM 14/11/02 BMIDS00935 BEGIN
				string cmdText = "select * from " + tableName + sqlNoLock + " where " + fieldValuePair;
				// DM 14/11/02 BMIDS00935 END

				WriteSQLToFile(cmdText);

				using (DataSet dataSet = new DataSet())
				{
					using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
					{
						using (SqlCommand sqlCommand = new SqlCommand(cmdText, sqlConnection))
						{
							sqlCommand.CommandType = CommandType.Text;

							using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
							{
								sqlDataAdapter.Fill(dataSet);
							}
						}
					}
					if (dataSet.Tables != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows != null && dataSet.Tables[0].Rows.Count > 0)
					{
						xmlData = GetXMLFromRecordsetEx(GetDataRow(dataSet), xmlClassDefinition);
					}
					else
					{
						// raise application error to be interpreted by calling object
						throw new RecordNotFoundException();
					}
				}
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlData;
		}

		/// <summary>
		/// Gets component data.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> object containing the retrieved data.
		/// </returns>
		public static XmlNode GetComponentData(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			XmlNode xmlNode = null;

			try
			{
				// Clone of original class def
				// Class Def to be passed to GetData method
				// Create the top level nodes for new request
				XmlDocument xmlOut = new XmlDocument();
				XmlNode xmlNewRequest = xmlOut.CreateElement("REQUEST");
				xmlOut.AppendChild(xmlNewRequest);
				string tableName = xmlClassDefinition.FirstChild.FirstChild.InnerText;
				XmlNode xmlTableNode = xmlOut.CreateElement(tableName);
				xmlNewRequest.AppendChild(xmlTableNode);
				XmlDocument xmlNewClassDefinition = new XmlDocument();
				XmlNode xmlClassTableNode = xmlNewClassDefinition.CreateElement("TABLENAME");
				xmlClassTableNode.InnerText = tableName;
				xmlNewClassDefinition.AppendChild(xmlClassTableNode);
				// Build all the PrimaryKey node in NewRequest, with the values mentioned in request
				XmlDocument xmlClonedClassDef = new XmlDocument();
				xmlClonedClassDef.AppendChild(xmlClonedClassDef.ImportNode(xmlClassDefinition.DocumentElement, true));
				bool addAllNodes = XmlAssist.GetAttributeValue(xmlRequest, "OUTPUT", "ALLFIELDS") == "1";
				XmlNode xmlCriteria = xmlRequest.SelectSingleNode(".//CRITERIA");
				XmlNodeList xmlNodeList = xmlClonedClassDef.SelectNodes("//PRIMARYKEY");

				foreach (XmlNode xmlChildNode in xmlNodeList)
				{
					string fieldName = xmlChildNode.FirstChild.InnerText;
					// Add an element to the NewRequest
					XmlNode xmlTempNode = xmlOut.CreateElement(fieldName);
					if (xmlCriteria != null)
					{
						xmlTempNode.InnerText = XmlAssist.GetTagValue((XmlElement)xmlCriteria, fieldName);
					}
					xmlTableNode.AppendChild(xmlTempNode);
					// Add this element to New Class Definition, only if the AllNodes attrib is set to False
					if (!addAllNodes)
					{
						xmlClassTableNode.AppendChild(xmlNewClassDefinition.ImportNode(xmlChildNode, true));
					}
				}
				// Add the other fields to be queried to new class definition, only if the AllNodes
				// attribute is set to False
				if (!addAllNodes)
				{
					xmlNodeList = xmlRequest.SelectNodes(".//FIELD");
					foreach (XmlNode xmlChildNode in xmlNodeList)
					{
						string fieldName = xmlChildNode.FirstChild.InnerText;
						xmlClassTableNode.AppendChild(xmlNewClassDefinition.ImportNode(XmlAssist.GetNodeFromClassDefByNodeValue(xmlClonedClassDef, fieldName), true));
					}
				}
				xmlNode = GetDataEx((XmlElement)xmlNewRequest, xmlNewClassDefinition);
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}

			return xmlNode;
		}

		/// <summary>
		/// Creates a SQL string from a specified xml request and class definition.
		/// </summary>
		/// <param name="request">Xml request data from which to build the SQL string.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <param name="tableName">Base table name.</param>
		public static void BuildSQLString(string request, string classDefinition, ref string tableName)
		{
			string fieldValuePair = null;
			string fields = null;
			string values = null;
			string select = null;
			BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select);
		}

		/// <summary>
		/// Creates a SQL string from a specified xml request and class definition.
		/// </summary>
		/// <param name="request">Xml request data from which to build the SQL string.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <param name="tableName">Base table name.</param>
		/// <param name="sqlFormatType">
		/// Enumerated <see cref="SQL_FORMAT_TYPE"/> used to determine the 
		/// format of the fieldValuePair or fields and values that are returned. 
		/// sftFieldValueSeparated: As used in Create, e.g "insert into TABLE ( FieldA,FieldB )
		/// values ( Value1,Value2 )"; 
		/// sftCommaSeparated: e.g. "update TABLE set FieldA = Value1,FieldB = Value2 where"; 
		/// sftAndSeparated: e.g. "select * from TABLE where FieldA = Value1 AND FieldB = Value2".
		/// </param>
		/// <param name="classDefinitionKey">
		/// Enumerated <see cref="CLASS_DEF_KEY"/> constant. 
		/// This and <paramref name="classDefinitionKeyAmount"/> are used together to build the SQL string and raise 
		/// appropriate errors, e.g.,
		/// <code>
		/// GetData	 cdkaALLKEYS cdkPRIMARYKEY
		/// Delete	  cdkaALLKEYS cdkPRIMARYKEY
		/// DeleteAll   cdkaANYKEYS cdkPRIMARYKEY
		/// FindList	cdkaANYKEYS cdkPRIMARYKEY
		/// Update	  cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
		/// Create	  cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
		/// </code>
		/// </param>
		/// <param name="classDefinitionKeyAmount">
		/// Enumerated <see cref="CLASS_DEF_KEY_AMOUNT"/> constant. Used with <paramref name="classDefinitionKey"/> 
		/// to provide error handling depending on which combination is specified.
		/// </param>
		/// <param name="fieldValuePair">
		/// String created or appended to, which contains the SQL formatted values
		/// based on the settings of <paramref name="sqlFormatType"/>. 
		/// Used with sftAndSeparated and sftCommaSeparated; raises error if either of 
		/// these is specified and no parameter passed. Used by GetData, Delete, DeleteAll, FindList.
		/// </param>
		/// <param name="fields">
		/// String created or appended to, which contains SQL formatted values
		/// based on the settings of SQL_FORMAT_TYPE. Used with sftFieldValueSeparated; 
		/// raises error if this is specified and no parameter is passed. Used by Create.
		/// </param>
		/// <param name="values">
		/// String created or appended to, which contains SQL formatted values
		/// based on the settings of SQL_FORMAT_TYPE. Used with sftFieldValueSeparated; 
		/// raises error if this is specified and no parameter is passed. Used by Create.
		/// </param>
		/// <param name="select">
		/// Optional string used in Find, when used with a subset of the XML class/table
		/// Definition. Will return a string of Table.Fields for the provided XML.
		/// </param>
		/// <exception cref="InvalidKeyStringException"></exception>
		/// <exception cref="InvalidParameterException"></exception>
		/// <remarks>
		/// This method checks all keys are present in the XML, and returns them with their
		/// associated values in a format ready for including in a SQL execution string.
		/// An exception is thrown dependant on the <paramref name="classDefinitionKeyAmount"/> agrument 
		/// if some of the keys do not have values specified.
		/// </remarks>
		public static void BuildSQLString(string request, string classDefinition, ref string tableName, SQL_FORMAT_TYPE sqlFormatType, CLASS_DEF_KEY classDefinitionKey, CLASS_DEF_KEY_AMOUNT classDefinitionKeyAmount, ref string fieldValuePair, ref string fields, ref string values, ref string select)
		{
			string pairSeparator = "";

			XmlDocument xmlDocument = XmlAssist.Load(request);
			XmlDocument xmlClassDefinitionDocument = XmlAssist.Load(classDefinition);
			XmlElement xmlElement = (XmlElement)xmlClassDefinitionDocument.GetElementsByTagName("TABLENAME")[0];
			if (xmlElement == null)
			{
				throw new MissingXMLTableNameException();
			}
			tableName = xmlElement.FirstChild.InnerText;
			if (tableName == null || tableName.Length == 0)
			{
				// "Missing Table Description in XMLClass Definition"
				throw new MissingTableDescException();
			}
			if (xmlDocument.GetElementsByTagName(tableName).Count < 1)
			{
				throw new MissingPrimaryTagException("Expected " + tableName + " tag");
			}
			switch (sqlFormatType)
			{
				// add error checking
				case SQL_FORMAT_TYPE.sftFieldValueSeparated:
					if (fields == null)
					{
						throw new InvalidParameterException("Fields must be specified for this format type");
					}
					if (values == null)
					{
						throw new InvalidParameterException("Values must be specified for this format type");
					}
					break;
				case SQL_FORMAT_TYPE.sftAndSeparated:
					if (fieldValuePair == null)
					{
						throw new InvalidParameterException("FieldValuePair must be specified for this format type");
					}
					pairSeparator = " AND ";
					break;
				case SQL_FORMAT_TYPE.sftCommaSeparated:
					if (fieldValuePair == null)
					{
						throw new InvalidParameterException("FieldValuePair must be specified for " + SQL_FORMAT_TYPE.sftFieldValueSeparated);
					}
					pairSeparator = " , ";
					break;
			}

			XmlNodeList xmlNodeList = null;
			switch (classDefinitionKey)
			{
				case CLASS_DEF_KEY.cdkPRIMARYKEY:
					xmlNodeList = xmlClassDefinitionDocument.GetElementsByTagName("PRIMARYKEY");
					break;
				case CLASS_DEF_KEY.cdkOTHERS:
					xmlNodeList = xmlClassDefinitionDocument.GetElementsByTagName("OTHERS");
					break;
				default:
					throw new InValidKeyException("Of " + classDefinitionKey + " for " + tableName);
			}

			// Loop round all the Elements
			foreach (XmlNode xmlNode in xmlNodeList)
			{
				// reset variables
				string fieldName = xmlNode.FirstChild.InnerText;
				string valueText = "";

				if (fieldName == null || fieldName.Length == 0)
				{
					// check the element
					// No description in provided XML
					// "Missing Key Description in XMLClass Definition"
					throw new MissingKeyDescException();
				}
				// Combine the field elements to be used in the Select statement
				// MCS 19/11/99
				if (select != null && select.Length > 0)
				{
					select = select + "," + tableName + "." + fieldName;
				}
				else
				{
					select = tableName + "." + fieldName;
				}
				if (xmlDocument.GetElementsByTagName(fieldName).Count != 0)
				{
					// we have a field to process
					valueText = xmlDocument.GetElementsByTagName(fieldName)[0].InnerText;
					if (valueText.Trim() == "" && classDefinitionKey == CLASS_DEF_KEY.cdkPRIMARYKEY)
					{
						// Primary key value cannot be null
						throw new InValidKeyValueException("Of " + classDefinitionKeyAmount + " for " + valueText);
					}

					XmlElement xmlTypeElem = (XmlElement)xmlNode;

					// Get the Type Description from the XMLClass Description
					string dataTypeText = XmlAssist.GetTagValue(xmlTypeElem, "TYPE");
					if (dataTypeText.Length == 0)
					{
						// "Missing Type Description Value in XMLClass Definition"
						throw new MissingTypeDescException();
					}
					// format the data according to its data type
					if (valueText == null || valueText.Length == 0)
					{
						valueText = "Null";
					}
					else
					{
						valueText = SqlAssist.GetFormattedValue(valueText, GetDbType(dataTypeText));
					}

					switch (sqlFormatType)
					{
						case SQL_FORMAT_TYPE.sftFieldValueSeparated:
							if (fields != null && fields.Length > 0)
							{
								fields = fields + ", ";
								values = values + ", ";
							}
							fields = fields + fieldName;
							values = values + valueText;

							break;
						case SQL_FORMAT_TYPE.sftAndSeparated:
						case SQL_FORMAT_TYPE.sftCommaSeparated:
							if (fieldValuePair != null && fieldValuePair.Length > 0)
							{
								fieldValuePair = fieldValuePair + pairSeparator;
							}
							fieldValuePair = fieldValuePair + fieldName + "=" + valueText;
							break;
						default:
							// fixme
							// Raise error
							break;
					}
				}
				else
				{
					// No Element in XML passed in
					switch (classDefinitionKeyAmount)
					{
						case CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS:
							// Only an error if we want all of them
							// "Not ALL the Elements for the Keys specified have been supplied"
							throw new MissingElementException("Table name: " + tableName);
						case CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS:
							// do nothing no error (skip this field)
							break;
						default:
							throw new InValidKeyValueException("Of " + classDefinitionKeyAmount + " for " + valueText);
					}
				}
			}
		}

		/// <summary>
		/// Creates a SQL string from a specified xml request and class definition.
		/// </summary>
		/// <param name="xmlRequest">Xml request data from which to build the SQL string.</param>
		/// <param name="xmlClassDefinition">Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="xmlRequest"/>.</param>
		/// <param name="tableName">Base table name.</param>
		public static void BuildSQLStringEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, ref string tableName)
		{
			string fieldValuePair = null;
			string fields = null;
			string values = null;
			string select = null;
			BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select, null);
		}

		/// <summary>
		/// Creates a SQL string from a specified xml request and class definition.
		/// </summary>
		/// <param name="xmlRequest">Xml request data from which to build the SQL string.</param>
		/// <param name="xmlClassDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="xmlRequest"/>.
		/// </param>
		/// <param name="tableName">Base table name.</param>
		/// <param name="sqlFormatType">
		/// Enumerated <see cref="SQL_FORMAT_TYPE"/> used to determine the 
		/// format of the fieldValuePair or fields and values that are returned. 
		/// sftFieldValueSeparated: As used in Create, e.g "insert into TABLE ( FieldA,FieldB )
		/// values ( Value1,Value2 )"; 
		/// sftCommaSeparated: e.g. "update TABLE set FieldA = Value1,FieldB = Value2 where"; 
		/// sftAndSeparated: e.g. "select * from TABLE where FieldA = Value1 AND FieldB = Value2".
		/// </param>
		/// <param name="classDefinitionKey">
		/// Enumerated <see cref="CLASS_DEF_KEY"/> constant. 
		/// This and <paramref name="classDefinitionKeyAmount"/> are used together to build the SQL string and raise 
		/// appropriate errors, e.g.,
		/// <code>
		/// GetData	 cdkaALLKEYS cdkPRIMARYKEY
		/// Delete	  cdkaALLKEYS cdkPRIMARYKEY
		/// DeleteAll   cdkaANYKEYS cdkPRIMARYKEY
		/// FindList	cdkaANYKEYS cdkPRIMARYKEY
		/// Update	  cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
		/// Create	  cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
		/// </code>
		/// </param>
		/// <param name="classDefinitionKeyAmount">
		/// Enumerated <see cref="CLASS_DEF_KEY_AMOUNT"/> constant. Used with <paramref name="classDefinitionKey"/> 
		/// to provide error handling depending on which combination is specified.
		/// </param>
		/// <param name="fieldValuePair">
		/// String created or appended to, which contains the SQL formatted values
		/// based on the settings of <paramref name="sqlFormatType"/>. 
		/// Used with sftAndSeparated and sftCommaSeparated; raises error if either of 
		/// these is specified and no parameter passed. Used by GetData, Delete, DeleteAll, FindList.
		/// </param>
		/// <param name="fields">
		/// String created or appended to, which contains SQL formatted values
		/// based on the settings of SQL_FORMAT_TYPE. Used with sftFieldValueSeparated; 
		/// raises error if this is specified and no parameter is passed. Used by Create.
		/// </param>
		/// <param name="values">
		/// String created or appended to, which contains SQL formatted values
		/// based on the settings of SQL_FORMAT_TYPE. Used with sftFieldValueSeparated; 
		/// raises error if this is specified and no parameter is passed. Used by Create.
		/// </param>
		/// <param name="select">
		/// Optional string used in Find, when used with a subset of the XML class/table
		/// Definition. Will return a string of Table.Fields for the provided XML.
		/// </param>
		/// <param name="itemName">Item name.</param>
		/// <exception cref="InvalidKeyStringException"></exception>
		/// <exception cref="InvalidParameterException"></exception>
		/// <remarks>
		/// This method checks all keys are present in the XML, and returns them with their
		/// associated values in a format ready for including in a SQL execution string.
		/// An exception is thrown dependant on the <paramref name="classDefinitionKeyAmount"/> agrument 
		/// if some of the keys do not have values specified.
		/// </remarks>
		public static void BuildSQLStringEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, ref string tableName, SQL_FORMAT_TYPE sqlFormatType, CLASS_DEF_KEY classDefinitionKey, CLASS_DEF_KEY_AMOUNT classDefinitionKeyAmount, ref string fieldValuePair, ref string fields, ref string values, ref string select, string itemName)
		{
			XmlElement xmlElement = (XmlElement)xmlClassDefinition.GetElementsByTagName("TABLENAME")[0];
			if (xmlElement == null)
			{
				throw new MissingXMLTableNameException();
			}
			if (tableName == null || tableName.Length == 0)
			{
				tableName = xmlElement.FirstChild.InnerText;
			}
			// The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
			if (itemName == null || itemName.Length == 0)
			{
				itemName = tableName;
			}
			if (tableName == null || tableName.Length == 0)
			{
				// "Missing Table Description in XMLClass Definition"
				throw new MissingTableDescException();
			}
			// The first table specified in the table class definition will be used to ascertain the key of the table/view
			// being retrieved on
			XmlNode xmlPrimaryTable = xmlClassDefinition.DocumentElement.SelectSingleNode(".//PRIMARYKEY").ParentNode;
			XmlElement xmlTableElement = null;
			if (xmlRequest.Name == itemName)
			{
				xmlTableElement = xmlRequest;
			}
			else
			{
				xmlTableElement = (XmlElement)xmlRequest.SelectSingleNode(itemName);
			}
			if (xmlTableElement == null)
			{
				throw new MissingPrimaryTagException("Expected " + itemName + " tag");
			}

			string pairSeparator = "";
			switch (sqlFormatType)
			{
				// add error checking
				case SQL_FORMAT_TYPE.sftFieldValueSeparated:

					if (fields == null)
					{
						throw new InvalidParameterException("Fields must be specified for this format type");
					}
					if (values == null)
					{
						throw new InvalidParameterException("Values must be specified for this format type");
					}
					break;
				case SQL_FORMAT_TYPE.sftAndSeparated:
					if (fieldValuePair == null)
					{
						throw new InvalidParameterException("FieldValuePair must be specified for this format type");
					}
					pairSeparator = " AND ";
					break;
				case SQL_FORMAT_TYPE.sftCommaSeparated:
					if (fieldValuePair == null)
					{
						throw new InvalidParameterException("FieldValuePair must be specified for " + SQL_FORMAT_TYPE.sftFieldValueSeparated);
					}
					pairSeparator = " , ";
					break;
			}

			XmlNodeList xmlNodeList = null;
			switch (classDefinitionKey)
			{
				case CLASS_DEF_KEY.cdkPRIMARYKEY:
					xmlNodeList = xmlClassDefinition.DocumentElement.SelectNodes(".//PRIMARYKEY");
					break;
				case CLASS_DEF_KEY.cdkOTHERS:
					xmlNodeList = xmlClassDefinition.DocumentElement.SelectNodes(".//OTHERS");
					break;
				default:
					throw new InValidKeyException("Of " + classDefinitionKey + " for " + tableName);
			}

			// Loop round all the Elements
			foreach (XmlNode xmlChildNode in xmlNodeList)
			{
				// reset variables
				string valueText = "";
				// reset variables
				string dataTypeText = "";
				// reset variables
				// Get the field Description from the XMLClass Description
				string fieldName = xmlChildNode.FirstChild.InnerText;

				if (fieldName == null || fieldName.Length == 0)
				{
					// check the element
					// No description in provided XML
					// "Missing Key Description in XMLClass Definition"
					throw new MissingKeyDescException();
				}
				// Combine the field elements to be used in the Select statement
				// MCS 19/11/99
				if (select != null && select.Length > 0)
				{
					select = select + "," + tableName + "." + fieldName;
				}
				else
				{
					select = select + tableName + "." + fieldName;
				}
				XmlNode xmlNode = xmlTableElement.SelectSingleNode(fieldName);
				if (xmlNode != null)
				{
					// we have a field to process
					valueText = xmlNode.InnerText;
					if (valueText.Trim().Length == 0 && classDefinitionKey == CLASS_DEF_KEY.cdkPRIMARYKEY && xmlElement.ParentNode == xmlPrimaryTable)
					{
						// Primary key value cannot be null
						throw new InValidKeyValueException(" NULL value for " + fieldName);
					}
					// Get the Type Description from the XMLClass Description
					dataTypeText = XmlAssist.GetTagValue(xmlElement, "TYPE");
					if (dataTypeText.Length == 0)
					{
						// "Missing Type Description Value in XMLClass Definition"
						throw new MissingTypeDescException();
					}
					// format the data according to its data type
					if (valueText == null || valueText.Length == 0)
					{
						valueText = "Null";
					}
					else
					{
						valueText = SqlAssist.GetFormattedValue(valueText, GetDbType(dataTypeText));
					}

					switch (sqlFormatType)
					{
						case SQL_FORMAT_TYPE.sftFieldValueSeparated:
							if (fields != null && fields.Length > 0)
							{
								fields = fields + ", ";
								values = values + ", ";
							}
							fields = fields + fieldName;
							values = values + valueText;
							break;

						case SQL_FORMAT_TYPE.sftAndSeparated:
						case SQL_FORMAT_TYPE.sftCommaSeparated:
							if (fieldValuePair != null && fieldValuePair.Length > 0)
							{
								fieldValuePair = fieldValuePair + pairSeparator;
							}
							fieldValuePair = fieldValuePair + fieldName + "=" + valueText;
							break;

						default:
							// fixme
							// Raise error
							break;
					}
				}
				else
				{
					// No Element in XML passed in
					switch (classDefinitionKeyAmount)
					{
						case CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS:
							if (xmlElement.ParentNode == xmlPrimaryTable)
							{
								// Only an error if we want all of them
								// "Not ALL the Elements for the Keys specified have been supplied"
								throw new MissingElementException("Table name: " + tableName);
							}
							break;
						case CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS:
							// do nothing no error (skip this field)
							break;
						default:
							throw new InValidKeyValueException("Of " + classDefinitionKeyAmount + " for " + valueText);
					}
				}
			}
			tableName = xmlClassDefinition.DocumentElement.FirstChild.InnerText;
		}

		/// <summary>
		/// Delete a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="request">Xml request data identifying which instance to delete.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		public static void Delete(string request, string classDefinition) 
		{
			try 
			{
				string fieldValuePair = "";
				string tableName = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select);

				string cmdText = "delete from " + tableName + " where " + fieldValuePair;
				AdoAssist.ExecuteSQLCommand(cmdText);

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Delete all instances of the persistant data associated with this data object that match the 
		/// key values specified.
		/// </summary>
		/// <param name="request">Xml request data identifying which instances to delete.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		public static void DeleteAll(string request, string classDefinition) 
		{
			try 
			{
				string fieldValuePair = "";
				string tableName = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);

				string cmdText = "delete from " + tableName + " where " + fieldValuePair;
				try
				{
					AdoAssist.ExecuteSQLCommand(cmdText);
				}
				catch (NoRowsAffectedException)
				{
					throw new NoRowsAffectedByDeleteAllException();
				}

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Delete all instances of the persistant data associated with this 
		/// data object that match the key values specified.
		/// </summary>
		/// <param name="xmlRequest">Xml request data identifying which instances to delete.</param>
		/// <param name="xmlClassDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="xmlRequest"/>.
		/// </param>
		public static void DeleteAllEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			try
			{
				string fieldValuePair = "";
				string tableName = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select, "");

				string cmdText = "delete from " + tableName + " where " + fieldValuePair;
				try
				{
					AdoAssist.ExecuteSQLCommand(cmdText);
				}
				catch (NoRowsAffectedException)
				{
					throw new NoRowsAffectedByDeleteAllException();
				}

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Delete a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="xmlRequest">Xml request data identifying which instance to delete.</param>
		/// <param name="xmlClassDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="xmlRequest"/>.
		/// </param>
		public static void DeleteEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			try
			{
				string fieldValuePair = "";
				string tableName = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select, null);

				string cmdText = "delete from " + tableName + " where " + fieldValuePair;
				AdoAssist.ExecuteSQLCommand(cmdText);

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="request">Xml request data which identifies the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <returns>
		/// String containing the retrieved xml data.
		/// </returns>
		public static string FindList(string request, string classDefinition)
		{
			return FindList(request, classDefinition, null);
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="request">Xml request data identifying the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <param name="orderByField">The field name to order the search result by.</param>
		/// <returns>
		/// String containing the retrieved xml data.
		/// </returns>
		public static string FindList(string request, string classDefinition, string orderByField) 
		{
			string findListText = null;
			
			try 
			{
				string fieldValuePair = "";
				string tableName = "";
				string cmdText = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);

				if (select == null || select.Length == 0)
				{
					// should never get here but just in case
					select = "*";
				}
				if (fieldValuePair == null || fieldValuePair.Length == 0)
				{
					// Find All
					// MCS 19/11/99 - only slect the required fields
					cmdText = "select " + select + " from " + tableName;
				}
				else
				{
					cmdText = "select " + select + " from " + tableName + " where " + fieldValuePair;
				}
				// AY 13/10/1999 - if an order by field name has been passed in, add to the SQL string
				if (orderByField != null && orderByField.Length > 0)
				{
					cmdText = cmdText + " order by " + orderByField;
				}

				findListText = GetFindListXmlText(cmdText, tableName, classDefinition);
			
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return findListText;
		}

		private static string GetFindListXmlText(string cmdText, string tableName, string classDefinition)
		{
			XmlDocument xmlDocument = new XmlDocument();

			WriteSQLToFile(cmdText);
			using (DataSet dataSet = new DataSet())
			{
				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
				{
					using (SqlCommand cmd = new SqlCommand(cmdText, sqlConnection))
					{
						using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
						{
							sqlDataAdapter.Fill(dataSet);
						}
					}
				}

				if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows == null || dataSet.Tables[0].Rows.Count == 0)
				{
					// raise application error to be interpreted by calling object
					throw new RecordNotFoundException();
				}

				XmlNode xmlListNode = xmlDocument.AppendChild(xmlDocument.CreateElement(tableName + "LIST"));
				// Loop through the rows.
				foreach (DataRow dataRow in dataSet.Tables[0].Rows)
				{
					XmlDocument xmlRowDocument = XmlAssist.Load(GetXMLFromRecordSet(dataRow, classDefinition));
					xmlListNode.AppendChild(xmlDocument.ImportNode(xmlRowDocument.DocumentElement, true));
				}
			}

			return xmlDocument.OuterXml;
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with this data object 
		/// for the values supplied.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> containing the retrieved data.
		/// </returns>
		public static XmlNode FindListEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return FindListEx(xmlRequest, xmlClassDefinition, null);
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with this data object 
		/// for the values supplied.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <param name="orderByField">The field name to order the search result by.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> containing the retrieved data.
		/// </returns>
		public static XmlNode FindListEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string orderByField)
		{
			// History:
			// SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
			// ------------------------------------------------------------------------------------------
			XmlElement xmlListElement = null;
			
			try
			{
				XmlDocument xmlOut = new XmlDocument();
				string sqlNoLock = "";

				// SR 20/12/02 : BM0209 - Add 'With (NOLOCK)' to SQL statement, if required
				if (xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK") != null)
				{
					if (AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
					{
						// SDS LIVE00009653 / 20/01/2004
						sqlNoLock = " with (nolock)";
					}
					xmlClassDefinition.DocumentElement.RemoveChild(xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK"));
				}
				// SR 20/12/02 : BM0209 - End
				string fieldValuePair = "";
				string tableName = "";
				string cmdText = "";
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select, "");
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select, "");

				if (select == null || select.Length == 0)
				{
					// should never get here but just incase
					select = "*";
				}
				if (fieldValuePair == null || fieldValuePair.Length == 0)
				{
					// Find All
					// MCS 19/11/99 - only slect the required fields
					cmdText = "select " + select + " from " + tableName + sqlNoLock;
					// SR 20/12/02 : BM0209 - add nolock hint
				}
				else
				{
					cmdText = "select " + select + " from " + tableName + sqlNoLock + " where " + fieldValuePair;
					// SR 20/12/02 : BM0209 - add nolock hint
				}
				// AY 13/10/1999 - if an order by field name has been passed in, add to the SQL string
				if (orderByField != null && orderByField.Length > 0)
				{
					cmdText = cmdText + " order by " + orderByField;
				}

				xmlListElement = GetFindListExXml(cmdText, tableName, xmlClassDefinition);

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlListElement;
		}

		private static XmlElement GetFindListExXml(string cmdText, string tableName, XmlDocument xmlClassDefinition)
		{
			XmlElement xmlListElement = null;

			WriteSQLToFile(cmdText);
			using (DataSet dataSet = new DataSet())
			{
				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
				{
					using (SqlCommand cmd = new SqlCommand(cmdText, sqlConnection))
					{
						using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
						{
							sqlDataAdapter.Fill(dataSet);
						}
					}
				}

				if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows == null || dataSet.Tables[0].Rows.Count == 0)
				{
					// raise application error to be interpreted by calling object
					throw new RecordNotFoundException();
				}

				XmlDocument xmlOut = new XmlDocument();
				xmlListElement = xmlOut.CreateElement(tableName + "LIST");
				xmlOut.AppendChild(xmlListElement);
				// DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
				// dataSet.MoveFirst
				// loop through the record set
				foreach (DataRow dataRow in dataSet.Tables[0].Rows)
				{
					GetXMLFromRecordsetEx(dataRow, xmlClassDefinition, xmlListElement);
				}
			}

			return xmlListElement;
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with 
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="request">Xml request data identifying the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <returns>
		/// String containing an xml data stream representation of the retrieved data.
		/// </returns>
		/// <remarks>
		/// This method differs from <see cref="FindList"/> in that it can take multiple search criteria 
		/// 'blocks'. For example, if searching for GROUPCONNECTION records, the passed xml could include 
		/// more than one 'blocks' of GROUPCONNECTION xml.
		/// </remarks>
		public static string FindListMultiple(string request, string classDefinition)
		{
			return FindListMultiple(request, classDefinition, null);
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with 
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="request">Xml request data identifying the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <param name="orderByField">The field name to order the search result by.</param>
		/// <returns>
		/// String containing an xml data stream representation of the retrieved data.
		/// </returns>
		/// <remarks>
		/// This method differs from <see cref="FindList"/> in that it can take multiple search criteria 
		/// 'blocks'. For example, if searching for GROUPCONNECTION records, the passed xml could include 
		/// more than one 'blocks' of GROUPCONNECTION xml.
		/// </remarks>
		public static string FindListMultiple(string request, string classDefinition, string orderByField) 
		{
			string findListText = null;
			
			try 
			{
				string cmdText = "";
				string tableName = "";
				string where = "";
				XmlDocument xmlRequest = XmlAssist.Load(request);
				XmlDocument xmlClassDefinition = XmlAssist.Load(classDefinition);
				// Get table name
				XmlNodeList xmlNodeList = xmlClassDefinition.GetElementsByTagName("TABLENAME");
				if (xmlNodeList.Count == 0)
				{
					throw new MissingTableNameException("TABLENAME tag not found");
				}
				else
				{
					tableName = xmlNodeList[0].FirstChild.InnerText;
				}
				// Loop through each search criteria block in turn, adding the criteria to tbe WHERE clause
				xmlNodeList = xmlRequest.GetElementsByTagName(tableName);
				string select = null;
				foreach (XmlNode objXmlNode in xmlNodeList)
				{
					string fieldValuePair = "";
					// Get the part of the WHERE clause corresponding to this criteria block
					string fields = null;
					string values = null;
					BuildSQLString(objXmlNode.OuterXml, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);
					BuildSQLString(objXmlNode.OuterXml, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);
					fieldValuePair = "(" + fieldValuePair + ")";
					if (where.Length > 0)
					{
						where = where + " OR ";
					}
					where = where + fieldValuePair;
				}

				if (select == null || select.Length == 0)
				{
					// should never get here but just incase
					select = "*";
				}
				cmdText = "select " + select + 
					" from " + tableName + 
					(where != null && where.Length > 0 ? " where " + where : "") +
					(orderByField != null && orderByField .Length > 0 ? " order by " + orderByField : "");

				findListText = GetFindListXmlText(cmdText, tableName, classDefinition);

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return findListText;
		}

		/// <summary>
		/// Gets the data for all instances of the persistant data associated with
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <returns>
		/// An <see cref="XmlNode"/> containing the retrieved data.
		/// </returns>
		/// <remarks>
		/// This method differs from <see cref="FindList"/> in that it can take multiple search criteria 
		/// 'blocks'. For example, if searching for GROUPCONNECTION records, the passed xml could include 
		/// more than one 'blocks' of GROUPCONNECTION xml.
		/// </remarks>
		public static XmlNode FindListMultipleEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return FindListMultipleEx(xmlRequest, xmlClassDefinition, null, null, null, null);
		}

		/// <summary>
		/// Gets the data for all instances of the persistant data associated with
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="xmlRequest">An xml string containing data identifies the instance(s) of the persistant data to be retrieved.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		/// <param name="orderByField">The field name to order the search result by.</param>
		/// <param name="itemName"></param>
		/// <param name="additionalCondition"></param>
		/// <param name="andOrToCondition"></param>
		/// <returns>
		/// An <see cref="XmlNode"/> containing the retrieved data.
		/// </returns>
		/// <remarks>
		/// This method differs from <see cref="FindList"/> in that it can take multiple search criteria 
		/// 'blocks'. For example, if searching for GROUPCONNECTION records, the passed xml could include 
		/// more than one 'blocks' of GROUPCONNECTION xml.
		/// </remarks>
		public static XmlNode FindListMultipleEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string orderByField, string itemName, string additionalCondition, string andOrToCondition)
		{
			XmlElement xmlListElement = null;

			try
			{
				string fieldValuePair = "";
				string tableName = "";
				string where = "";
				string strTempSelect = "";

				// Get table name
				XmlNodeList xmlNodeList = xmlClassDefinition.GetElementsByTagName("TABLENAME");
				if (xmlNodeList.Count == 0)
				{
					throw new MissingTableNameException("TABLENAME tag not found");
				}
				else
				{
					tableName = xmlNodeList[0].FirstChild.InnerText;
				}
				// The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
				if (itemName == null || itemName.Length == 0)
				{
					itemName = tableName;
				}
				// RF 12/05/2003 BM0536 Start
				string sqlNoLock = null;
				if (xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK") != null)
				{
					if (AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
					{
						sqlNoLock = " with (nolock)";
					}
					xmlClassDefinition.DocumentElement.RemoveChild(xmlClassDefinition.SelectSingleNode("TABLENAME/SQLNOLOCK"));
				}
				// RF 12/05/2003 BM0536 End
				// Loop through each search criteria block in turn, adding the criteria to tbe WHERE clause
				XmlNode xmlNode = null;
				int intNodeListCount = 0;
				if (xmlRequest.Name == itemName)
				{
					xmlNode = xmlRequest;
					intNodeListCount = 0;
				}
				else
				{
					xmlNodeList = xmlRequest.GetElementsByTagName(itemName);
					intNodeListCount = xmlNodeList.Count;
				}

				int i = 0;
				string select = null;
				while (xmlNode != null || i < intNodeListCount)
				{
					if (xmlNode == null)
					{
						// This will happen every time except when the xmlNode is set prior to the While..Loop
						xmlNode = xmlNodeList[i];
					}
					fieldValuePair = "";
					// Get the part of the WHERE clause corresponding to this criteria block
					string fields = null;
					string values = null;
					BuildSQLStringEx((XmlElement)xmlNode, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref strTempSelect, itemName);
					BuildSQLStringEx((XmlElement)xmlNode, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref strTempSelect, itemName);
					fieldValuePair = "(" + fieldValuePair + ")";
					if (where != null && where.Length > 0)
					{
						where = where + " OR ";
					}
					where = where + fieldValuePair;
					if (select == null || select.Length == 0)
					{
						select = strTempSelect;
					}
					xmlNode = null;
					i++;
				}

				if (additionalCondition != null && additionalCondition.Length > 0)
				{
					if (where != null && where.Length != 0)
					{
						where = where + " " + andOrToCondition + " (" + additionalCondition + ")";
					}
					else
					{
						where = additionalCondition;
					}
				}

				if (select == null || select.Length == 0)
				{
					// should never get here but just incase
					select = "*";
				}
				else
				{
					select = "distinct " + select;
				}
				string cmdText = "select " + select +
					" from " + tableName + sqlNoLock +
					(where != null && where.Length > 0 ? " where " + where : "") +
					(orderByField != null && orderByField.Length > 0 ? " order by " + orderByField : "");

				xmlListElement = GetFindListExXml(cmdText, tableName, xmlClassDefinition);

				ContextUtility.SetComplete();

			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlListElement;
		}

		/// <summary>
		/// Creates an instance of the persistant data associated with this data object
		/// for each set of data in the request.
		/// </summary>
		/// <param name="request">Xml request data containing the data to be persisted.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		public static void Create(string request, string classDefinition) 
		{
			try 
			{
				string tableName = null;
				string fieldValuePair = null;
				string fields = "";
				string values = "";
				string select = null;

				// SQL insert statement
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftFieldValueSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select);
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftFieldValueSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select);

				string cmdText = "INSERT INTO " + tableName + " (" + fields + ") values ( " + values + ")";
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Creates an instance of the persistant data associated with this data object
		/// for each set of data in the request.
		/// </summary>
		/// <param name="xmlRequest">An <see cref="XmlElement"/> containing the data to be persisted.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		public static void CreateEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			try
			{
				string tableName = null;
				string fieldValuePair = null;
				string fields = "";
				string values = "";
				string select = null;

				// SQL insert statement

				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftFieldValueSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref fieldValuePair, ref fields, ref values, ref select, "");
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftFieldValueSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref fieldValuePair, ref fields, ref values, ref select, "");

				string cmdText = "INSERT INTO " + tableName + " (" + fields + ") values ( " + values + ")";
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Updates a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="request">Xml request data containing the data to be persisted.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		public static void Update(string request, string classDefinition) 
		{
			try 
			{
				string strAndSeparated = "";
				string strCommaSeparated = "";
				string tableName = "";

				// SQL insert statement

				string fields = null;
				string values = null;
				string select = null;
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref strAndSeparated, ref fields, ref values, ref select);
				BuildSQLString(request, classDefinition, ref tableName, SQL_FORMAT_TYPE.sftCommaSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref strCommaSeparated, ref fields, ref values, ref select);
				string cmdText = "UPDATE " + tableName + " SET " + strCommaSeparated + " where " + strAndSeparated;
				AdoAssist.ExecuteSQLCommand(cmdText);

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Updates a single instance of the persistant data associated with this data object.
		/// </summary>
		/// <param name="xmlRequest">An <see cref="XmlElement"/> containing the data to be persisted.</param>
		/// <param name="xmlClassDefinition">The xml class definition against which <paramref name="xmlRequest"/> is parsed.</param>
		public static void UpdateEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			try
			{
				string strAndSeparated = "";
				string strCommaSeparated = "";
				string tableName = "";

				// SQL insert statement
				string fields = null;
				string values = null;
				string select = null;
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftAndSeparated, CLASS_DEF_KEY.cdkPRIMARYKEY, CLASS_DEF_KEY_AMOUNT.cdkaALLKEYS, ref strAndSeparated, ref fields, ref values, ref select, "");
				BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, SQL_FORMAT_TYPE.sftCommaSeparated, CLASS_DEF_KEY.cdkOTHERS, CLASS_DEF_KEY_AMOUNT.cdkaANYKEYS, ref strCommaSeparated, ref fields, ref values, ref select, "");
				string cmdText = "UPDATE " + tableName + " SET " + strCommaSeparated + " where " + strAndSeparated;
				AdoAssist.ExecuteSQLCommand(cmdText);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}
		}

		/// <summary>
		/// Gets the <see cref="DataTable"/> object for the first table in a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <returns>
		/// The <see cref="DataTable"/> object for the first table in the <see cref="DataSet"/> object.
		/// </returns>
		/// <exception cref="ErrAssistException">The table does not exist.</exception>
		public static DataTable GetDataTable(DataSet dataSet)
		{
			return GetDataTable(dataSet, 0);
		}

		/// <summary>
		/// Gets the <see cref="DataTable"/> object for a specified table in a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <param name="tableIndex">The index the table in the <see cref="DataSet"/> object.</param>
		/// <returns>
		/// The <see cref="DataTable"/> object for the specified table in the <see cref="DataSet"/> object.
		/// </returns>
		/// <exception cref="ErrAssistException">The table does not exist.</exception>
		public static DataTable GetDataTable(DataSet dataSet, int tableIndex)
		{
			if (dataSet.Tables == null || dataSet.Tables.Count == 0)
			{
				throw new ErrAssistException("DataSet '" + dataSet.DataSetName + "' does not contain any tables.");
			}
			if (tableIndex >= dataSet.Tables.Count)
			{
				throw new ErrAssistException("DataSet '" + dataSet.DataSetName + "' does not contain table index " + tableIndex.ToString());
			}
			return dataSet.Tables[tableIndex];
		}

		/// <summary>
		/// Gets the <see cref="DataRow"/> object for the first row a specified <see cref="DataTable"/> object.
		/// </summary>
		/// <param name="dataTable">The <see cref="DataTable"/> object.</param>
		/// <returns>
		/// The <see cref="DataRow"/> object for the first row in the <see cref="DataTable"/> object.
		/// </returns>
		/// <exception cref="RecordNotFoundException">The row does not exist.</exception>
		public static DataRow GetDataRow(DataTable dataTable)
		{
			return GetDataRow(dataTable, 0);
		}

		/// <summary>
		/// Gets the <see cref="DataRow"/> object for the first row in the first table in a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <returns>
		/// The <see cref="DataRow"/> object for the first row in the first table in the <see cref="DataSet"/> object.
		/// </returns>
		/// <exception cref="RecordNotFoundException">The row does not exist.</exception>
		public static DataRow GetDataRow(DataSet dataSet)
		{
			return GetDataRow(dataSet, 0, 0);
		}

		/// <summary>
		/// Gets the <see cref="DataRow"/> object for a specified table and row in a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <param name="tableIndex">The index of the table in the <see cref="DataSet"/> object.</param>
		/// <param name="rowIndex">The index of the row in the table in the <see cref="DataSet"/> object.</param>
		/// <returns>
		/// The <see cref="DataRow"/> object for the specified table and row in the <see cref="DataSet"/> object.
		/// </returns>
		/// <exception cref="RecordNotFoundException">The row does not exist.</exception>
		public static DataRow GetDataRow(DataSet dataSet, int tableIndex, int rowIndex)
		{
			return GetDataRow(GetDataTable(dataSet, tableIndex), rowIndex);
		}

		/// <summary>
		/// Gets the <see cref="DataRow"/> object for a specified row in a specified <see cref="DataTable"/> object.
		/// </summary>
		/// <param name="dataTable">The <see cref="DataTable"/> object.</param>
		/// <param name="rowIndex">The index of the row in the <see cref="DataTable"/> object.</param>
		/// <returns>
		/// The <see cref="DataRow"/> object for the specified row in the <see cref="DataTable"/> object.
		/// </returns>
		/// <exception cref="RecordNotFoundException">The row does not exist.</exception>
		public static DataRow GetDataRow(DataTable dataTable, int rowIndex)
		{
			if (dataTable.Rows == null || dataTable.Rows.Count == 0)
			{
				throw new RecordNotFoundException("DataTable '" + dataTable.TableName + "' does not contain any rows.");
			}
			if (rowIndex >= dataTable.Rows.Count)
			{
				throw new RecordNotFoundException("DataTable '" + dataTable.TableName + "' does not contain row index " + rowIndex.ToString());
			}
			return dataTable.Rows[rowIndex];
		}

		/// <summary>
		/// Delegate for getting the <see cref="DBDATATYPE"/> constant for a specified column name.
		/// </summary>
		/// <param name="columnName">The column name.</param>
		/// <returns>
		/// The <see cref="DBDATATYPE"/> constant that defines the data type of the column.
		/// </returns>
		public delegate DBDATATYPE GetDbTypeDelegate(string columnName);

		/// <summary>
		/// Creates an xml data stream from field elements in an <see cref="DataRow"/>. 
		/// Add any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataRow">The <see cref="DataRow"/>.</param>
		/// <param name="tableName">The table name.</param>
		/// <param name="getDbTypeDelegate">Delegate for getting the <see cref="DBDATATYPE"/> constant for the specified column name.</param>
		/// <returns>
		/// The xml data stream.
		/// </returns>
		public static string GetXMLFromRecordSet(DataRow dataRow, string tableName, GetDbTypeDelegate getDbTypeDelegate)
		{
			string response = null;

			try
			{
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlElement = xmlOut.CreateElement(tableName);
				XmlNode xmlTableNode = xmlOut.AppendChild(xmlElement);

				DataTable dataTable = dataRow.Table;
				for (int columnIndex = 0; columnIndex < dataTable.Columns.Count && columnIndex < dataRow.ItemArray.Length; columnIndex++)
				{
					if (!dataRow.IsNull(columnIndex))
					{
						string columnName = dataTable.Columns[columnIndex].ColumnName;
						AppendFieldValue(dataRow, xmlTableNode, getDbTypeDelegate(columnName), columnName, null);
					}
				}
				response = xmlOut.OuterXml;
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return response;
		}

		/// <summary>
		/// Creates an xml data stream from field elements in an <see cref="DataRow"/>. 
		/// Add any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataRow">The <see cref="DataRow"/>.</param>
		/// <param name="classDefinition">
		/// Xml class definition describing the table, fields and formats
		/// to be used to parse the <paramref name="request"/>.
		/// </param>
		/// <returns>
		/// The xml data stream.
		/// </returns>
		public static string GetXMLFromRecordSet(DataRow dataRow, string classDefinition)
		{
			XmlDocument xmlDocumentOut = new XmlDocument();

			try
			{
				XmlDocument xmlClassDefinitionDocument = XmlAssist.Load(classDefinition);
				// get the number of elements for this class def
				XmlElement xmlClassDefinitionElement = (XmlElement)xmlClassDefinitionDocument.GetElementsByTagName("TABLENAME")[0];
				if (xmlClassDefinitionElement == null)
				{
					throw new MissingXMLTableNameException();
				}

				string tableName = xmlClassDefinitionElement.FirstChild.InnerText;
				if (tableName.Length == 0)
				{
					// "Missing Table Description in XMLClass Definition"
					throw new MissingTableDescException();
				}

				XmlNode xmlTableNode = xmlDocumentOut.AppendChild(xmlDocumentOut.CreateElement(tableName));
				if (xmlClassDefinitionElement.ChildNodes.Count < 1)
				{
					throw new NoFieldsFoundException("For " + tableName);
				}

				// Ignore first element - table name.
				for (int intField = 1; intField < xmlClassDefinitionElement.ChildNodes.Count; intField++)
				{
					try
					{
						XmlNode xmlChildNode = xmlClassDefinitionElement.ChildNodes[intField];
						// Get the field out
						if (xmlChildNode == null)
						{
							throw new NoFieldItemFoundException("For " + tableName);
						}
						// This is the field (PRIMARYKEY or OTHER as defined in ClassDef)
						string columnName = xmlChildNode.ChildNodes[0].InnerText;
						if (columnName.Length == 0)
						{
							throw new NoFieldItemNameException("For " + columnName);
						}
						// Get the Data type out, this is the TYPE as defined under (PRIMARYKEY or OTHER in ClassDef)
						string dataTypeText = XmlAssist.GetTagValue((XmlElement)xmlChildNode, "TYPE");

						AppendFieldValue(dataRow, xmlTableNode, GetDbType(dataTypeText), columnName, xmlChildNode);
					}
					catch (RecordNotFoundException)
					{
						// Ignore error and continue.
					}
					// add derived values to generated XML in base Class
				}
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlDocumentOut.OuterXml;
		}

		private static void AppendFieldValue(DataRow dataRow, XmlNode xmlTableNode, DBDATATYPE dbDataType, string columnName, XmlNode xmlChildNode)
		{
			// This is the element name for the type we want
			if (dbDataType != DBDATATYPE.dbdtNotStored)
			{
				XmlElement xmlElement = xmlTableNode.OwnerDocument.CreateElement(columnName);

				if (!dataRow.IsNull(columnName))
				{
					object field = dataRow[columnName];
					bool exists = false;

					// set element text; also set element attribute where applicable
					switch (dbDataType)
					{
						case DBDATATYPE.dbdtGuid:
							xmlElement.InnerText = SqlAssist.GuidToString((byte[])field, SqlDbType.Binary, GUIDSTYLE.guidBinary);
							break;
						case DBDATATYPE.dbdtCurrency:
							NumberFormatInfo nfi = new NumberFormatInfo();
							nfi.NumberDecimalDigits = 2;
							if (field is decimal)
							{
								xmlElement.InnerText = ((decimal)field).ToString("N", nfi);
							}
							else if (field is double)
							{
								xmlElement.InnerText = ((double)field).ToString("N", nfi);
							}
							else if (field is float)
							{
								xmlElement.InnerText = ((float)field).ToString("N", nfi);
							}
							else
							{
								xmlElement.InnerText = field.ToString();
							}
							break;
						case DBDATATYPE.dbdtDouble:
							xmlElement.SetAttribute("RAW", field.ToString());
							// Get the  format mask out
							string formatMask = xmlChildNode != null ? XmlAssist.GetTagValue((XmlElement)xmlChildNode, "FORMATMASK", out exists, true) : null;
							if (!exists || formatMask == null || formatMask.Length == 0)
							{
								// No FORMAT MASK defined so just use the default mask
								formatMask = "0.00";
							}
							if (field is decimal)
							{
								xmlElement.InnerText = ((decimal)field).ToString(formatMask);
							}
							else if (field is double)
							{
								xmlElement.InnerText = ((double)field).ToString(formatMask);
							}
							else if (field is float)
							{
								xmlElement.InnerText = ((float)field).ToString(formatMask);
							}
							else
							{
								xmlElement.InnerText = field.ToString();
							}
							break;
						case DBDATATYPE.dbdtComboId:
							// Get the  combo entry out
							if (xmlChildNode != null)
							{
								string comboGroupName = XmlAssist.GetTagValue((XmlElement)xmlChildNode, "COMBO", out exists, true);
								// trap no records found
								if (!exists)
								{
									// No combo entry
									throw new NoComboTagValueException("For " + columnName);
								}
								else
								{
									// This is the COMBO value as defined under (COMBO in ClassDef)
									// Add the defined COMBO value
									int valueId = Convert.ToInt32(field);
									string comboText = null;
									try
									{
										comboText = ComboAssist.GetComboText(comboGroupName, valueId);
									}
									catch
									{
										// Ignore any errors reading the combo text.
									}
									if (comboText != null)
									{
										xmlElement.SetAttribute("TEXT", comboText);
										xmlElement.InnerText = valueId.ToString();
									}
								}
							}
							break;
						case DBDATATYPE.dbdtDate:
							xmlElement.InnerText = SqlAssist.DateToString((DateTime)field);
							break;
						case DBDATATYPE.dbdtDateTime:
							xmlElement.InnerText = SqlAssist.DateTimeToString((DateTime)field);
							break;
						case DBDATATYPE.dbdtNotStored:
							// skip this field (shouldn't get here anyway)
							break;
						default:
							xmlElement.InnerText = Convert.ToString(field);
							break;
					}
				}
				// append the element whether or not the field has a null value
				xmlTableNode.AppendChild(xmlElement);
			}
			else
			{
				throw new InValidDataTypeValueException("Of " + dbDataType.ToString() + " for " + columnName);
			}
		}

		/// <summary>
		/// Creates an xml data stream from field elements in a <see cref="DataRow"/>. 
		/// Adds any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataRow">The <see cref="DataRow"/>.</param>
		/// <param name="xmlClassDefinition">Xml class definition describing the table, fields and formats.</param>
		/// <returns>
		/// The node added.
		/// </returns>
		public static XmlNode GetXMLFromRecordsetEx(DataRow dataRow, XmlDocument xmlClassDefinition)
		{
			return GetXMLFromRecordsetEx(dataRow, xmlClassDefinition, null);
		}

		/// <summary>
		/// Creates an xml data stream from field elements in a <see cref="DataRow"/>. 
		/// Adds any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataRow">The <see cref="DataRow"/>.</param>
		/// <param name="xmlClassDefinition">Xml class definition describing the table, fields and formats.</param>
		/// <param name="xmlInNode">The <see cref="XmlNode"/> to which the xml should be attached.</param>
		/// <returns>
		/// The node added.
		/// </returns>
		public static XmlNode GetXMLFromRecordsetEx(DataRow dataRow, XmlDocument xmlClassDefinition, XmlNode xmlInNode)
		{
			XmlNode xmlDataNode = null;

			try
			{
				xmlDataNode = GetXMLForTableSchema(dataRow, xmlClassDefinition.DocumentElement, xmlInNode);
				// 
				// Ascertain whether the data was retrieved via a view or not
				// 
				bool isFromView = true;
				foreach (XmlNode xmlChildNode in xmlClassDefinition.FirstChild.ChildNodes)
				{
					if (xmlChildNode.Name != "" && xmlChildNode.Name != "TABLENAME")
					{
						// Found a child node that does not represent a sub-table for a view - therefore the class definition
						// must represent a physical table rather than a view
						isFromView = false;
						break;
					}
				}
				// 
				// If the data was retrieved via a view then make sure that the 'primary' table data in the view replaces
				// the view node in the data XML schema
				// 
				if (isFromView)
				{
					// The first table specified in the table class definition will be used as the 'root' data node in the returned
					// XML. Any data for other tables in the schema will be created as child nodes of this primary data node
					XmlNode xmlPrimaryTableNode = xmlClassDefinition.DocumentElement.SelectSingleNode("TABLENAME/PRIMARYKEY").ParentNode;
					XmlNode xmlViewDataNode = xmlDataNode;
					if (xmlPrimaryTableNode != null)
					{
						XmlNode xmlPrimaryDataNode = xmlViewDataNode.SelectSingleNode(xmlPrimaryTableNode.FirstChild.InnerText);
						if (xmlPrimaryDataNode != null)
						{
							// Move the primary table data up one level in the XML to replace the node named after the view used in the
							// SQL query
							// Move all children of the data (i.e. view) node to be children of the primary data node instead
							foreach (XmlNode xmlChildNode in xmlViewDataNode.ChildNodes)
							{
								if (xmlChildNode != xmlPrimaryDataNode)
								{
									xmlPrimaryDataNode.AppendChild(xmlChildNode);
								}
							}
							if (xmlInNode == null)
							{
								xmlDataNode = xmlPrimaryDataNode;
							}
							else
							{
								// Move the primary data node up one level in the XML
								xmlDataNode = xmlDataNode.ParentNode.AppendChild(xmlPrimaryDataNode);
								// Remove the now empty view data node from the return XML
								xmlDataNode.ParentNode.RemoveChild(xmlViewDataNode);
							}
						}
					}
				}
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlDataNode;
		}

		/// <summary>
		/// Creates an xml data stream from field elements in a <see cref="DataRow"/>. 
		/// Adds any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataRow">The <see cref="DataRow"/>.</param>
		/// <param name="xmlClassDefinition">Xml class definition describing the table, fields and formats.</param>
		/// <param name="xmlInNode">The <see cref="XmlNode"/> to which the xml should be attached.</param>
		/// <returns>
		/// The node added.
		/// </returns>
		public static XmlNode GetXMLForTableSchema(DataRow dataRow, XmlElement xmlClassDefinition, XmlNode xmlInNode)
		{
			// ------------------------------------------------------------------------------------------
			// BMIDS History:
			// 
			// Prog  Date		Description
			// MDC   14/01/2003  BM0249 - Ignore SQLNOLOCK element in classdefs
			// ------------------------------------------------------------------------------------------
			// BBG History
			// TK	23/08/2004  BBG1211 (Extracted from BMIDS788) Only create ComboDO instance if actually needed
			// ------------------------------------------------------------------------------------------

			string tableName = null;
			string elementName = null;
			XmlNode xmlTableNode = null;
			DBDATATYPE dbDataType = DBDATATYPE.dbdtNotStored;

			try
			{
				// BM0249 MDC 14/01/2003
				foreach (XmlNode xmlNoLockNode in xmlClassDefinition.SelectNodes("//SQLNOLOCK"))
				{
					xmlNoLockNode.ParentNode.RemoveChild(xmlNoLockNode);
				}
				// BM0249 MDC 14/01/2003 - End
				tableName = xmlClassDefinition.FirstChild.InnerText;
				if (tableName.Length == 0)
				{
					// "Missing Table Description in XMLClass Definition"
					throw new MissingTableDescException();
				}

				// If a node hasn't been sent in then create a new document
				XmlDocument xmlDocument = null;
				if (xmlInNode == null)
				{
					xmlDocument = new XmlDocument();
					xmlTableNode = xmlDocument.AppendChild(xmlDocument.CreateElement(tableName));
				}
				else
				{
					xmlDocument = xmlInNode.OwnerDocument;
					xmlTableNode = xmlInNode.AppendChild(xmlDocument.CreateElement(tableName));
				}
				if (xmlClassDefinition.ChildNodes.Count < 1)
				{
					throw new NoFieldsFoundException();
				}
				bool isSkipField = true;
				// This flag is used to skip the first (i.e. tablename) node in each table definition
				foreach (XmlNode xmlNode in xmlClassDefinition.ChildNodes)
				{
					elementName = null;
					try
					{
						if (!isSkipField)
						{
							// This is the field (PRIMARYKEY or OTHER as defined in ClassDef)
							elementName = xmlNode.FirstChild.InnerText;
							if (elementName.Length == 0)
							{
								throw new NoFieldItemNameException();
							}
							if (xmlNode.Name == "TABLENAME")
							{
								// 
								// Node represents a 'sub-table' part of the schema
								// 
								GetXMLForTableSchema(dataRow, (XmlElement)xmlNode, xmlTableNode);
							}
							else
							{
								// 
								// Add the value for this field to the resulting XML
								// 

								// Get the Data type out, this is the TYPE as defined under (PRIMARYKEY or OTHER in ClassDef)
								string dataTypeText = xmlNode.ChildNodes[1].InnerText;
								dbDataType = GetDbType(dataTypeText);
								// This is the element name for the type we want
								if (dbDataType != DBDATATYPE.dbdtNotStored)
								{
									XmlElement xmlFieldDataElem = xmlDocument.CreateElement(elementName);
									object field = dataRow[elementName];
									if (field != DBNull.Value)
									{
										bool exists = false;

										// set element text; also set element attribute where applicable
										switch (dbDataType)
										{
											case DBDATATYPE.dbdtGuid:
												xmlFieldDataElem.InnerText = SqlAssist.GuidToString((byte[])field, SqlDbType.Binary, GUIDSTYLE.guidBinary);
												break;
											case DBDATATYPE.dbdtCurrency:
												NumberFormatInfo nfi = new NumberFormatInfo();
												nfi.NumberDecimalDigits = 2;
												if (field is decimal)
												{
													xmlNode.InnerText = ((decimal)field).ToString("N", nfi);
												}
												else if (field is double)
												{
													xmlNode.InnerText = ((double)field).ToString("N", nfi);
												}
												else if (field is float)
												{
													xmlNode.InnerText = ((float)field).ToString("N", nfi);
												}
												else
												{
													xmlNode.InnerText = field.ToString();
												}
												break;
											case DBDATATYPE.dbdtDouble:
												// Get the  format mask out
												string formatMask = XmlAssist.GetTagValue((XmlElement)xmlNode, "FORMATMASK", out exists, true);
												if (!exists || formatMask == null || formatMask.Length == 0)
												{
													// No FORMAT MASK defined so just use the default mask
													xmlFieldDataElem.InnerText = Convert.ToString(field);
												}
												else
												{
													// This is the FORMAT MASK  value as defined under (TYPE in ClassDef)
													// Add the defined FORMATMASK
													xmlFieldDataElem.SetAttribute("RAW", Convert.ToString(field));
													if (field is decimal)
													{
														xmlFieldDataElem.InnerText = ((decimal)field).ToString(formatMask);
													}
													else if (field is double)
													{
														xmlFieldDataElem.InnerText = ((double)field).ToString(formatMask);
													}
													else if (field is float)
													{
														xmlFieldDataElem.InnerText = ((float)field).ToString(formatMask);
													}
													else
													{
														xmlFieldDataElem.InnerText = field.ToString();
													}
												}
												break;
											case DBDATATYPE.dbdtComboId:
												// Get the  combo entry out
												string comboGroupName = XmlAssist.GetTagValue((XmlElement)xmlNode, "COMBO", out exists, true);
												// trap no records found
												if (!exists)
												{
													// No combo entry
													throw new NoComboTagValueException();
												}
												else
												{
													int valueId = Convert.ToInt32(field);
													string comboText = null;
													try
													{
														comboText = ComboAssist.GetComboText(comboGroupName, valueId);
													}
													catch
													{
														// Ignore any errors reading the combo text.
													}
													if (comboText != null)
													{
														xmlFieldDataElem.SetAttribute("TEXT", comboText);
														xmlFieldDataElem.InnerText = valueId.ToString();
													}
													// reset the error number
												}
												break;
											case DBDATATYPE.dbdtDate:
												xmlFieldDataElem.InnerText = SqlAssist.DateToString((DateTime)field);
												break;
											case DBDATATYPE.dbdtDateTime:
												xmlFieldDataElem.InnerText = SqlAssist.DateTimeToString((DateTime)field); 
												break;
											case DBDATATYPE.dbdtNotStored:
												// skip this field (shouldn't get here anyway)
												break;
											default:
												xmlFieldDataElem.InnerText = Convert.ToString(field);
												break;
										}
									}

									// append the element whether or not the field has a null value
									xmlTableNode.AppendChild(xmlFieldDataElem);
								}
								else
								{
									throw new InValidDataTypeValueException("Of " + dataTypeText);
								}
								// end if (dbDataType != dbdtNotStored)
							}
						}
						isSkipField = false;
					}
					catch (RecordNotFoundException)
					{
						// Ignore error and continue.
					}
				}
			}
			catch (Exception exception)
			{
				ErrAssistException errAssistException = new ErrAssistException(exception, TransactionAction.SetOnErrorType);
				if (dbDataType != DBDATATYPE.dbdtNotStored && elementName != null && elementName.Length != 0)
				{
					errAssistException.OmigaMessageText += "; Table=\"" + tableName + "\" Field=\"" + elementName + "\"";
				}
				else
				{
					errAssistException.OmigaMessageText += "; Table=\"" + tableName + "\"";
				}
				throw errAssistException;
			}

			return xmlTableNode;
		}

		/// <summary>
		/// Creates an xml data stream from field elements in a <see cref="DataTable"/>. 
		/// Adds any values derived from field elements to the xml.
		/// </summary>
		/// <param name="dataTable">The <see cref="DataTable"/>.</param>
		/// <param name="xmlClassDefinition">Xml class definition describing the table, fields and formats.</param>
		/// <param name="xmlNode">The <see cref="XmlNode"/> to which the xml should be attached.</param>
		public static void GetXMLFromWholeRecordset(DataTable dataTable, XmlDocument xmlClassDefinition, XmlNode xmlNode)
		{
			foreach (DataRow dataRow in dataTable.Rows)
			{
				GetXMLForTableSchema(dataRow, xmlClassDefinition.DocumentElement, xmlNode);
			}
		}

		private static DBDATATYPE GetDbType(string dataTypeText)
		{
			// header ----------------------------------------------------------------------------------
			// description:
			// Get the datatype (DBDATATYPE, an enumerated list defined in SqlAssist) associated
			// with a string description of the data type.
			// pass:
			// dataTypeText
			// String description of the required datatype, e.g. "dbdtString"
			// return:
			// enumerated value
			// ------------------------------------------------------------------------------------------

			DBDATATYPE dbType = DBDATATYPE.dbdtNotStored;

			switch (dataTypeText)
			{
				case "dbdtString":
					dbType = DBDATATYPE.dbdtString;
					break;
				case "dbdtInt":
					dbType = DBDATATYPE.dbdtInt;
					break;
				case "dbdtDouble":
					dbType = DBDATATYPE.dbdtDouble;
					break;
				case "dbdtDate":
					dbType = DBDATATYPE.dbdtDate;
					break;
				case "dbdtGuid":
					dbType = DBDATATYPE.dbdtGuid;
					break;
				case "dbdtCurrency":
					dbType = DBDATATYPE.dbdtCurrency;
					break;
				case "dbdtComboId":
					dbType = DBDATATYPE.dbdtComboId;
					break;
				case "dbdtBoolean":
					dbType = DBDATATYPE.dbdtBoolean;
					break;
				case "dbdtLong":
					dbType = DBDATATYPE.dbdtLong;
					break;
				case "dbdtDateTime":
					dbType = DBDATATYPE.dbdtDateTime;
					break;
				case "dbdtNotStored":
					dbType = DBDATATYPE.dbdtNotStored;
					break;
				default:
					throw new InvalidParameterException("Invalid datatype value: " + dataTypeText);
			}

			return dbType;
		}

		/// <summary>
		/// This function has been replaced with BuildSQLString.
		/// </summary>
		/// <param name="request"></param>
		/// <param name="classDefinition"></param>
		/// <param name="key"></param>
		/// <param name="type"></param>
		/// <param name="fieldValuePair"></param>
		/// <param name="tableName"></param>
		[Obsolete("Use BuildSQLString", true)]
		public static void GetKeyString(string request, string classDefinition, string key, string type, ref string fieldValuePair, ref string tableName) 
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Gets the next sequence number for a field given a set of primary key
		/// values to search on.
		/// </summary>
		/// <param name="request">The xml request.</param>
		/// <param name="classDefinition">The class definition.</param>
		/// <param name="sequenceFieldName">The sequence field name.</param>
		/// <returns>
		/// The next sequence number.
		/// </returns>
		/// <remarks>
		/// The primary key is determined from the <paramref name="xmlClassDefinition"/>parameter 
		/// and the values for that key from the <paramref name="xmlRequest"/> parameter. 
		/// A SQL statement is then built to find the maximum value for the specified sequence 
		/// field on the table given the specified key values.
		/// </remarks>
		public static int GetNextSequenceNumber(string request, string classDefinition, string sequenceFieldName) 
		{
			XmlDocument xmlRequest = XmlAssist.Load(request);
			XmlDocument xmlClassDefinition = XmlAssist.Load(classDefinition);
			XmlNodeList xmlNodeList = xmlClassDefinition.GetElementsByTagName("TABLENAME");
			string tableName = null;
			if (xmlNodeList.Count == 0)
			{
				throw new MissingTableNameException("TABLENAME tag not found");
			}
			else
			{
				tableName = xmlNodeList[0].FirstChild.InnerText;
			}

			return GetNextSequenceNumberEx(xmlRequest.DocumentElement, xmlClassDefinition, tableName, sequenceFieldName);
		}

		/// <summary>
		/// Gets the next sequence number for a field given a set of primary key
		/// values to search on.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <param name="tableName">The table name.</param>
		/// <param name="sequenceFieldName">The sequence field name.</param>
		/// <returns>
		/// The next sequence number.
		/// </returns>
		/// <remarks>
		/// The primary key is determined from the <paramref name="xmlClassDefinition"/>parameter 
		/// and the values for that key from the <paramref name="xmlRequest"/> parameter. 
		/// A SQL statement is then built to find the maximum value for the specified sequence 
		/// field on the table <paramref name="tableName"/> given the specified key values.
		/// </remarks>
		public static int GetNextSequenceNumberEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string tableName, string sequenceFieldName)
		{
			int nextSequenceNumber = 0;

			try 
			{
				string valueText = "";
				string sqlNoLock = "";

				// Get table name
				XmlNodeList xmlValueNodeList = xmlClassDefinition.GetElementsByTagName("TABLENAME");
				if (xmlValueNodeList.Count == 0)
				{
					throw new MissingTableNameException("TABLENAME tag not found");
				}
				else
				{
					tableName = xmlValueNodeList[0].FirstChild.InnerText;
				}
				// Get the primary key nodes from the class definition
				XmlNodeList xmlKeyNodeList = xmlClassDefinition.GetElementsByTagName("PRIMARYKEY");
				if (xmlKeyNodeList.Count == 0)
				{
					throw new MissingPrimaryTagException("No keys found");
				}
				// 
				// Build WHERE clause from primary key values
				// 
				string where = "";
				bool isEndOfKey = false;
				foreach (XmlNode xmlKeyNode in xmlKeyNodeList)
				{
					if (xmlKeyNode.FirstChild.InnerText == sequenceFieldName)
					{
						// Field in primary key matching sequence field has been found; so stop building WHERE clause here
						isEndOfKey = true;
						sqlNoLock = XmlAssist.GetTagValue((XmlElement)xmlKeyNode, "SQLNOLOCK");
						// DM 14/11/02 BMIDS00935
					}
					if (!isEndOfKey)
					{
						// Get type of field
						string fieldType = XmlAssist.GetTagValue((XmlElement)xmlKeyNode, "TYPE");
						// Get node in the data XML corresponding to this key
						xmlValueNodeList = xmlRequest.GetElementsByTagName(xmlKeyNode.FirstChild.InnerText);
						if (xmlValueNodeList.Count != 0)
						{
							if (xmlKeyNodeList.Count == 0)
							{
								throw new InvalidParameterException("Incomplete key specified");
							}
							valueText = xmlValueNodeList[0].InnerText;
							// AND value to WHERE clause
							if (where != null && where.Length > 0)
							{
								where = where + " AND ";
							}
							where = where + "(" + xmlKeyNode.FirstChild.InnerText +
								" = " + SqlAssist.GetFormattedValue(valueText, GetDbType(fieldType)) + ")";
						}
					}
				}
				// 
				// Execute the SQL
				// 
				string cmdText = "SELECT MAX(" + sequenceFieldName + ") FROM " + tableName;
				// DM 14/11/02 BMIDS00935 BEGIN
				if (sqlNoLock == "TRUE" && AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
				{
					// SDS LIVE00009653 / 20/01/2004
					cmdText = cmdText + " WITH (NOLOCK)";
				}
				// DM 14/11/02 BMIDS00935 END
				if (where != null && where.Length > 0)
				{
					cmdText = cmdText + " WHERE " + where;
				}

				WriteSQLToFile(cmdText);
				nextSequenceNumber = 1;
				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
				{
					using (SqlCommand cmd = new SqlCommand(cmdText, sqlConnection))
					{
						sqlConnection.Open();
						using (SqlDataReader sqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
						{
							if (sqlDataReader.Read() && sqlDataReader.HasRows && sqlDataReader[0] != DBNull.Value)
							{
								nextSequenceNumber = Convert.ToInt32(sqlDataReader[0]) + 1;
							}
						}
					}
				}
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}

			return nextSequenceNumber;
		}

		/// <summary>
		/// Gets the next sequence number for a field given a set of primary key
		/// values to search on.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="classDefinition">The class definition.</param>
		/// <param name="sequenceFieldName">The sequence field name.</param>
		/// <returns>The next sequence number.</returns>
		public static int GenerateSequenceNumber(XmlDocument xmlRequest, string classDefinition, string sequenceFieldName) 
		{
			return GenerateSequenceNumberEx(xmlRequest.DocumentElement, XmlAssist.Load(classDefinition), sequenceFieldName);
		}

		/// <summary>
		/// Gets the next sequence number for a field given a set of primary key
		/// values to search on.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <param name="sequenceFieldName">The sequence field name.</param>
		/// <returns>The next sequence number.</returns>
		/// <remarks>
		/// The primary key is determined from the <paramref name="xmlClassDefintion"/> parameter 
		/// and the values for that key from the <paramref name="xmlRequest"/> parameter. 
		/// A SQL statement is then built to find the maximum value for the specified sequence 
		/// field on the appropriate table given the specified key values.
		/// </remarks>
		public static int GenerateSequenceNumberEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string sequenceFieldName) 
		{
			int nextSequenceNumber = 0;

			try 
			{
				XmlNode xmlTableNode = null;			
				string tableName = "";

				// Get table name
				XmlNode xmlNode = xmlClassDefinition.SelectSingleNode("TABLENAME");
				if (xmlNode == null)
				{
					throw new MissingTableNameException("TABLENAME tag not found");
				}
				else
				{
					tableName = xmlNode.FirstChild.InnerText;
				}
				if (xmlRequest.Name == tableName)
				{
					xmlTableNode = xmlRequest;
				}
				else
				{
					xmlTableNode = xmlRequest.SelectSingleNode(tableName);
				}
				if (xmlTableNode == null)
				{
					throw new MissingPrimaryTagException(tableName + " tag not found");
				}
				// Generate the sequence number element if it does not already exist
				XmlNode xmlSequenceNode = xmlTableNode.SelectSingleNode(sequenceFieldName);
				if (xmlSequenceNode == null)
				{
					// No sequence number tag exists yet, so create one now
					xmlSequenceNode = xmlTableNode.AppendChild(xmlRequest.OwnerDocument.CreateElement(sequenceFieldName));
				}
				// Insert new sequence number into XML
				nextSequenceNumber = GetNextSequenceNumberEx(xmlRequest, xmlClassDefinition, tableName, sequenceFieldName);
				xmlSequenceNode.InnerText = Convert.ToString(nextSequenceNumber);
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}

			return nextSequenceNumber;
		}

		/// <summary>
		/// Writes the passed SQL to a file on the C drive.
		/// </summary>
		/// <param name="cmdText">The SQL to write to the file</param>
		private static void WriteSQLToFile(string cmdText)
		{
			AdoAssist.WriteSQLToFile(cmdText);
		}
	}
}

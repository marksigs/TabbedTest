/*
--------------------------------------------------------------------------------------------
Workfile:			AdoAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		.Net version of adoAssist.bas and ADOAssist.cls.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
SR		05/03/2001  New public function 'adoGetOmigaNumberForDatabaseError'
AW		20/03/2001  Added 'PARAMMODE' check to ConvertRecordSetToXML
AS		31/05/2001  Added GetValidRecordset for SQL Server Port
AS		11/06/2001  Fixed compile error with GENERIC_SQL=1
LD		12/06/2001  SYS2368 Modify code to get the connection string for Oracle and SQLServer
LD		11/06/2001  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
LD		15/06/2001  SYS2368 SQL Server Port - Guids to be generated as type adBinary in function GetParameter
LD		19/09/2001  SYS2722 SQL Server Port - Make function GuidToString public
IK		12/10/2001  SYS2803 Work Arounds for MSXML3 IXMLDOMNodeList bug
PSC		17/10/2001  SYS2815 Allow integrated security with SQL Server
SG		18/06/2002  SYS4889 Code error in ExecuteGetRecordSet
SDS		22/01/2004	LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/2004	BBG1821 Performance related fixes
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from adoAssist.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

// Ambiguous reference in cref attribute.
#pragma warning disable 419

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	// =============================================
	// Enumeration Declaration Section
	// =============================================
	/// <summary>
	/// Contains constants that indicate the type of database connection.
	/// </summary>
	public enum DBPROVIDER 
	{
		/// <summary>
		/// Unknown database type.
		/// </summary>
		omiga4DBPROVIDERUnknown,
		/// <summary>
		/// Native Oracle driver.
		/// </summary>
		omiga4DBPROVIDEROracle,
		/// <summary>
		/// Microsoft Oracle driver.
		/// </summary>
		omiga4DBPROVIDERMSOracle,
		/// <summary>
		/// SQL Server.
		/// </summary>
		omiga4DBPROVIDERSQLServer,
	}
	
	/// <summary>
	/// Contains constants that indicate the type of database lock.
	/// </summary>
	public enum LockType
	{
		/// <summary>
		/// Allow others to read table, but prevent them from updating it; 
		/// lock held until the end of the transaction.
		/// </summary>
		lckShare,
		/// <summary>
		/// Allow transaction to read data (without blocking other readers) and update it
		/// later with the assurance that the data will not have changed since it was read
		/// lock held until the end of the transaction.
		/// </summary>
		lckUpdate,
		/// <summary>
		/// Prevent others from reading or updating the table;
		/// lock held until the end of the transaction.
		/// </summary>
		lckExclusive,
	}

	/// <summary>
	/// ADO .Net helper class.
	/// </summary>
	/// <remarks>
	/// <para>
	/// <i>Do not use in new code - use the <see cref="System.Data"/> namespace and 
	/// child namespaces directly instead.</i>
	/// </para>
	/// <para>
	/// Methods have been ported from the VB6 AdoAssist module 
	/// (file name adoAssist.bas) and from the VB6 ADOAssist module 
	/// (file name ADOAssist.cls). 
	/// </para>
	/// <para>
	/// Most of the methods in this class are declared <b>static</b>, the remaining instance methods 
	/// require a schema xml document that is specific to the <see cref="AdoAssist"/> instance, i.e., 
	/// it depends on the location and name of the assembly which created this <see cref="AdoAssist"/> 
	/// instance, which is used to determine the location and name of the schema xml file.
	/// </para>
	/// <para>
	/// The LoadSchema (formerly adoLoadSchema) method is now declared <b>private</b>, as it is 
	/// automatically called by the <see cref="AdoAssist"/> constructor. Similarly, 
	/// the BuildDbConnectionString (formerly adoBuildDbConnectionString) method is also declared private 
	/// (and static) as it is automatically called by <see cref="GetDbConnectString"/>. Because 
	/// these two methods are now automatically called, there is generally no need to implement 
	/// the equivalent to the VB6 Main method, as this usually only calls adoLoadSchema and 
	/// adoBuildDbConnectionString.
	/// </para>
	/// </remarks>
	public sealed class AdoAssist
	{
		private enum GUIDSTYLE
		{
			guidBinary = 0,
			guidHyphen = 1,
			guidLiteral = 2,
		}

		private static Dictionary<string, XmlDocument> _xmlSchemaDocuments = new Dictionary<string, XmlDocument>();
		private static string _databaseConnectionString;
		private static int _databaseRetries;
		private static DBPROVIDER _databaseProvider;

		private const int SqlServerErrorNumberDuplicateKey = 2627;
		private const int SqlServerErrorNumberChildRecordsFound = 547;
		private const int OracleErrorNumberDuplicateKey = 1;
		private const int OracleErrorNumberChildRecordsFound = 2292;

		private Assembly _callingAssembly;

		/// <summary>
		/// Initializes a new instance of the <see cref="AdoAssist"/> class.
		/// </summary>
		/// <param name="callingAssembly">The assembly which is calling the <see cref="AdoAssist"/> class.</param>
		/// <remarks>
		/// <para>
		/// The <paramref name="callingAssembly"/> parameter is used to load an xml schema file, where 
		/// the location and name of the schema file are derived from the location and name of the 
		/// assembly. For example, if the assembly is "C:\Program Files\Marlborough Stirling\Omiga 4\DLL\omAAA.dll"
		/// then the schema file is "C:\Program Files\Marlborough Stirling\Omiga 4\XML\omAAA.xml". 
		/// Similarly, if the assembly is "C:\Program Files\Marlborough Stirling\Omiga 4\DLLDotNet\omBBB.dll"
		/// then the schema file is "C:\Program Files\Marlborough Stirling\Omiga 4\XML\omBBB.xml". 
		/// The <paramref name="callingAssembly"/> parameter can be defined as follows:
		/// <code>
		/// AdoAssist adoAssist = new AdoAssist(System.Reflection.Assembly.GetExecutingAssembly());
		/// </code>
		/// </para>
		/// </remarks>
		public AdoAssist(Assembly callingAssembly)
		{
			_callingAssembly = callingAssembly;
			LoadSchema();
		}

		#region adoAssistEx

		/// <summary>
		/// Loads the xml schema file for the assembly associated with this AdoAssist instance.
		/// </summary>
		/// <returns>
		/// The schema document.
		/// </returns>
		private XmlDocument LoadSchema()
		{
			XmlDocument xmlSchemaDocument = null;

			AppInstance appInstance = new AppInstance(_callingAssembly);
			string key = appInstance.Title;

			// Lock the global cache first before accessing it.
			lock (_xmlSchemaDocuments)
			{
				if (_xmlSchemaDocuments.ContainsKey(key))
				{
					// Schema already in the global cache, so do not load it again.
					xmlSchemaDocument = _xmlSchemaDocuments[key];
				}
				else
				{
					// pick up XML map from "...\Omiga 4\XML" directory
					// Only do the subsitution once to change DLL -> XML
					string schemaFileName = appInstance.Path + Path.DirectorySeparatorChar + appInstance.Title + ".xml";
					schemaFileName = schemaFileName.Replace("DLLDotNet", "XML");
					schemaFileName = schemaFileName.Replace("DLL", "XML");
					xmlSchemaDocument = new XmlDocument();
					xmlSchemaDocument.Load(schemaFileName);
					// Save the schema in the global cache.
					_xmlSchemaDocuments[key] = xmlSchemaDocument;
				}
			}

			return xmlSchemaDocument;
		}

		/// <summary>
		/// Get the schema xml document for a specified schema name.
		/// </summary>
		/// <param name="schemaName">The name of the schema.</param>
		/// <returns>
		/// The schema xml document for the schema name <paramref name="schemaName"/>.
		/// </returns>
		/// <remarks>
		/// Schema documents are stored in a global cache, so that they are only loaded once.
		/// </remarks>
		public XmlNode GetSchema(string schemaName)
		{
			XmlDocument xmlSchemaDocument = LoadSchema();
			string pattern = "//" + schemaName + "[@DATASRCE]";
			return xmlSchemaDocument.SelectSingleNode(pattern);
		}

		/// <summary>
		/// Calls <see cref="Create"/> for each node in an xml node list, with a specified schema.
		/// </summary>
		/// <param name="xmlRequestNodeList">The xml node list.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void CreateFromNodeList(XmlNodeList xmlRequestNodeList, string schemaName) 
		{
			// -------------------------------------------------------------------------------
			// BMIDS History
			// RF	 18/11/02	BMIDS00935 Applied Core Change SYS4752 (SYSMCP0734)
			// -------------------------------------------------------------------------------
			XmlNode xmlSchemaNode = GetSchema(schemaName);
			foreach (XmlNode xmlNode in xmlRequestNodeList)
			{
				Create(xmlNode, xmlSchemaNode);
			}
		}

		/// <summary>
		/// Calls <see cref="Create"/> for an xml request node with a specified schema.
		/// </summary>
		/// <param name="xmlRequestNode">The xml request node.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void CreateFromNode(XmlNode xmlRequestNode, string schemaName) 
		{
			Create(xmlRequestNode, GetSchema(schemaName));
		}

		/// <summary>
		/// Calls <see cref="Update"/> for each node in an xml node list, with a specified schema.
		/// </summary>
		/// <param name="xmlRequestNodeList">The xml node list.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void UpdateFromNodeList(XmlNodeList xmlRequestNodeList, string schemaName) 
		{
			XmlNode xmlSchemaNode = GetSchema(schemaName);
			foreach (XmlNode xmlNode in xmlRequestNodeList)
			{
				Update(xmlNode, xmlSchemaNode, false);
			}
		}

		/// <summary>
		/// Calls <see cref="Update"/> for an xml request node with a specified schema
		/// (non unique instances are not allowed).
		/// </summary>
		/// <param name="xmlRequestNode">The xml request node.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void UpdateFromNode(XmlNode xmlRequestNode, string schemaName) 
		{
			Update(xmlRequestNode, GetSchema(schemaName), false);
		}

		/// <summary>
		/// Calls <see cref="Update"/> for an xml request node with a specified schema 
		/// (non unique instances are allowed).
		/// </summary>
		/// <param name="xmlRequestNode">The xml request node.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void UpdateMultipleInstancesFromNode(XmlNode xmlRequestNode, string schemaName) 
		{
			Update(xmlRequestNode, GetSchema(schemaName), true);
		}

		/// <summary>
		/// For a specified request and schema, gets data as xml.
		/// </summary>
		/// <param name="xmlRequestNode">The request node.</param>
		/// <param name="xmlResponseNode">The response node in which the data is returned.</param>
		/// <param name="schemaName">The name of the schema.</param>
		public void GetAsXML(XmlNode xmlRequestNode, XmlNode xmlResponseNode, string schemaName)
		{
			GetAsXML(xmlRequestNode, xmlResponseNode, schemaName, null, null);
		}

		/// <summary>
		/// For a specified request, schema, filter and ordering, gets data as xml.
		/// </summary>
		/// <param name="xmlRequestNode">The request node.</param>
		/// <param name="xmlResponseNode">The response node in which the data is returned.</param>
		/// <param name="schemaName">The name of the schema.</param>
		/// <param name="filter">The filter to add to the SQL where clause.</param>
		/// <param name="orderBy">The ordering to add to the SQL order by clause.</param>
		public void GetAsXML(XmlNode xmlRequestNode, XmlNode xmlResponseNode, string schemaName, string filter, string orderBy) 
		{
			const string orderById = "_ORDERBY_";

			if (XmlAssist.AttributeValueExists(xmlRequestNode, orderById))
			{
				if (orderBy == null)
				{
					orderBy = "";
				}
				else if (orderBy.Length > 0)
				{
					orderBy = orderBy + ",";
				}
				orderBy = orderBy + XmlAssist.GetAttributeText(xmlRequestNode, orderById, "");
			}
			XmlNode xmlSchemaNode = GetSchema(schemaName);
			CloneBaseEntityRefs(xmlSchemaNode);
			if (XmlAssist.GetAttributeText(xmlSchemaNode, "ENTITYTYPE", "") == "PROCEDURE")
			{
				GetStoredProcAsXML(xmlRequestNode, xmlSchemaNode, xmlResponseNode);
			}
			else
			{
				GetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlResponseNode, filter, orderBy);
			}
		}

		/// <summary>
		/// Inserts a record for a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		public static void Create(XmlNode xmlDataNode, XmlNode xmlSchemaNode) 
		{
			string dataSource = XmlAssist.GetAttributeText(xmlSchemaNode, "DATASRCE", "");
			if (dataSource.Length == 0)
			{
				dataSource = xmlSchemaNode.Name;
			}
			// get any generated keys & default values
			foreach (XmlNode xmlSchemaFieldNode in xmlSchemaNode.ChildNodes)
			{
				if (xmlDataNode.Attributes.GetNamedItem(xmlSchemaFieldNode.Name) == null)
				{
					if (xmlSchemaFieldNode.Attributes.GetNamedItem("DEFAULT") != null)
					{
						XmlAttribute xmlAttribute = xmlDataNode.OwnerDocument.CreateAttribute(xmlSchemaFieldNode.Name);
						xmlAttribute.InnerText = xmlSchemaFieldNode.Attributes.GetNamedItem("DEFAULT").InnerText;
						xmlDataNode.Attributes.SetNamedItem(xmlAttribute);
					}
					if (xmlSchemaFieldNode.Attributes.GetNamedItem("KEYTYPE") != null)
					{
						if ((xmlSchemaFieldNode.Attributes.GetNamedItem("KEYTYPE").InnerText == "PRIMARY") || (xmlSchemaFieldNode.Attributes.GetNamedItem("KEYTYPE").InnerText == "SECONDARY"))
						{
							if (xmlSchemaFieldNode.Attributes.GetNamedItem("KEYSOURCE") != null)
							{
								if (xmlSchemaFieldNode.Attributes.GetNamedItem("KEYSOURCE").InnerText == "GENERATED")
								{
									GetGeneratedKey(xmlDataNode, xmlSchemaFieldNode);
								}
							}
						}
					}
				}
			}
			int databaseRetries = GetDbRetries();
			using (SqlCommand sqlCommand = new SqlCommand())
			{
				int parameterSuffix = 0;
				string columns = "";
				string values = "";

				foreach (XmlAttribute xmlDataAttribute in xmlDataNode.Attributes)
				{
					XmlNode xmlSchemaFieldNode = xmlSchemaNode.SelectSingleNode(xmlDataAttribute.Name);
					if (xmlSchemaFieldNode != null)
					{
						if (columns != null && columns.Length > 0)
						{
							columns = columns + ",";
						}
						columns = columns + xmlDataAttribute.Name;
						if (values.Length > 0)
						{
							values = values + ",";
						}
						SqlParameter sqlParameter = GetParameter(xmlSchemaFieldNode, xmlDataAttribute, true, ref parameterSuffix);
						sqlCommand.Parameters.Add(sqlParameter);
						values = values + "@" + sqlParameter.ParameterName;
					}
				}
				string cmdText = "INSERT INTO " + dataSource + " (" + columns + ") VALUES (" + values + ")";
				System.Diagnostics.Debug.WriteLine("AdoAssist.Create SQL: " + cmdText);
				if (values.Length > 0)
				{
					sqlCommand.CommandType = CommandType.Text;
					sqlCommand.CommandText = cmdText;
					int recordsAffected = ExecuteSqlCommandNonQuery(sqlCommand);
					System.Diagnostics.Debug.WriteLine("AdoAssist.Create records created: " + recordsAffected);
				}
			}
		}

		private static int ExecuteSqlCommandNonQuery(SqlCommand sqlCommand)
		{
			int recordsAffected = 0;

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection())
				{
					sqlConnection.ConnectionString = GetDbConnectString();
					sqlConnection.Open();
					sqlCommand.Connection = sqlConnection;
					recordsAffected = sqlCommand.ExecuteNonQuery();
				}
			}
			finally
			{
				sqlCommand.Dispose();
			}

			return recordsAffected;
		}

		/// <summary>
		/// Updates a record for a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The number of keys in <paramref name="xmlDataNode"/> does not match the number of keys in 
		/// <paramref name="xmlSchemaNode"/>.
		/// </exception>
		public static void Update(XmlNode xmlDataNode, XmlNode xmlSchemaNode) 
		{
			Update(xmlDataNode, xmlSchemaNode, false);
		}

		/// <summary>
		/// Updates a record for a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="isNonUniqueInstanceAllowed">
		/// If <b>false</b> an <see cref="XMLMissingAttributeException"/> exception is thrown if the number of 
		/// keys in <paramref name="xmlDataNode"/> does not match the number of keys in 
		/// <paramref name="xmlSchemaNode"/>.
		/// </param>
		/// <exception cref="XMLMissingAttributeException">
		/// The number of keys in <paramref name="xmlDataNode"/> does not match the number of keys in 
		/// <paramref name="xmlSchemaNode"/>, and <paramref name="isNonUniqueInstanceAllowed"/> 
		/// is <b>false</b>.
		/// </exception>
		public static void Update(XmlNode xmlDataNode, XmlNode xmlSchemaNode, bool isNonUniqueInstanceAllowed) 
		{
			string dataSource = XmlAssist.GetAttributeText(xmlSchemaNode, "DATASRCE", "");
			if (dataSource.Length == 0)
			{
				dataSource = xmlSchemaNode.Name;
			}
			int databaseRetries = GetDbRetries();
			using (SqlCommand sqlCommand = new SqlCommand())
			{
				string cmdText = "UPDATE " + dataSource;
				string sqlSetText = "";
				string sqlWhereText = "";
				int setCount = 0;
				int dataKeyCount = 0;
				bool setValues = false;

				XmlNodeList xmlSchemaNodeList = xmlSchemaNode.SelectNodes("*[@DATATYPE]");
				int parameterSuffix = 0;
				foreach (XmlNode xmlSchemaChildNode in xmlSchemaNodeList)
				{
					setValues = true;
					XmlAttribute xmlSchemaAttribute = (XmlAttribute)xmlSchemaChildNode.Attributes.GetNamedItem("KEYTYPE");
					if (xmlSchemaAttribute != null)
					{
						if (xmlSchemaAttribute.Value == "PRIMARY" || xmlSchemaAttribute.Value == "SECONDARY")
						{
							setValues = false;
						}
					}
					if (setValues)
					{
						XmlAttribute xmlDataAttribute =
							(XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlSchemaChildNode.Name);
						if (xmlDataAttribute != null)
						{
							if (sqlSetText.Length != 0)
							{
								sqlSetText = sqlSetText + ", ";
							}
							SqlParameter sqlParameter = GetParameter(xmlSchemaChildNode, xmlDataAttribute, false, ref parameterSuffix);
							sqlCommand.Parameters.Add(sqlParameter);
							sqlSetText = sqlSetText + xmlSchemaChildNode.Name + " = @" + sqlParameter.ParameterName;
							setCount++;
						}
					}
				}
				xmlSchemaNodeList = xmlSchemaNode.SelectNodes("*[@KEYTYPE='PRIMARY']");
				int schemaKeyCount = xmlSchemaNodeList.Count;
				foreach (XmlNode xmlSchemaChildNode in xmlSchemaNodeList)
				{
					XmlAttribute xmlDataAttribute =
						(XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlSchemaChildNode.Name);
					if (xmlDataAttribute != null)
					{
						if (xmlDataAttribute.Value.Length != 0)
						{
							if (sqlWhereText.Length != 0)
							{
								sqlWhereText = sqlWhereText + " AND ";
							}
							SqlParameter sqlParameter = GetParameter(xmlSchemaChildNode, xmlDataAttribute, false, ref parameterSuffix);
							sqlCommand.Parameters.Add(sqlParameter);
							sqlWhereText = sqlWhereText + xmlSchemaChildNode.Name + " = @" + sqlParameter.ParameterName;
							dataKeyCount++;
						}
					}
				}
				cmdText = cmdText + " SET " + sqlSetText;
				cmdText = cmdText + " WHERE " + sqlWhereText;
				System.Diagnostics.Debug.WriteLine("AdoAssist.Update SQL: " + cmdText);
				if (isNonUniqueInstanceAllowed == false && schemaKeyCount != dataKeyCount)
				{
					throw new XMLMissingAttributeException();
				}
				if (setCount > 0)
				{
					sqlCommand.CommandType = CommandType.Text;
					sqlCommand.CommandText = cmdText;
					int recordsAffected = ExecuteSqlCommandNonQuery(sqlCommand);
					System.Diagnostics.Debug.WriteLine("AdoAssist.Update records updated: " + recordsAffected);
				}
			}
		}

		/// <summary>
		/// Calls a stored procedure defined in a specified schema node, with arguments defined in a specified data node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="xmlResponseNode">The response node in which the data is returned.</param>
		public static void GetStoredProcAsXML(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlResponseNode) 
		{
			using (SqlCommand sqlCommand = new SqlCommand())
			{
				XmlNodeList xmlSchemaKeyNodeList = xmlSchemaNode.SelectNodes("*[@KEYTYPE='PRIMARY']");
				string sqlParameterClause = "";
				foreach (XmlNode xmlSchemaKeyNode in xmlSchemaKeyNodeList)
				{
					XmlAttribute xmlParameterName = (XmlAttribute)xmlSchemaKeyNode.Attributes.GetNamedItem("PARAMNAME");
					string parameterName = xmlParameterName != null ? xmlParameterName.Value : null;
					if (parameterName == null || parameterName.Length == 0)
					{
						throw new ErrAssistException("Missing PARAMNAME attribute for schema node " + xmlSchemaKeyNode.OuterXml + ".");
					}
					SqlParameter sqlParameter = 
						AddParameter(
							xmlSchemaKeyNode, (XmlAttribute)XmlAssist.GetAttributeNode(xmlDataNode, xmlSchemaKeyNode.Name),
							sqlCommand, parameterName);
					if (sqlParameterClause.Length != 0)
					{
						sqlParameterClause += ",";
					}
					sqlParameterClause += "@" + sqlParameter.ParameterName;
				}
				string cmdText = xmlSchemaNode.Attributes.GetNamedItem("DATASRCE").InnerText;
				System.Diagnostics.Debug.WriteLine("AdoAssist.GetStoredProcAsXML SQL: " + cmdText);
				sqlCommand.CommandType = CommandType.StoredProcedure;
				sqlCommand.CommandText = cmdText;
				using (DataSet dataSet = ExecuteGetRecordSet(sqlCommand))
				{
					if (dataSet != null && dataSet.Tables != null && dataSet.Tables.Count != 0)
					{
						ConvertRecordSetToXML(
							xmlSchemaNode,
							xmlResponseNode,
							dataSet, IsComboLookUpRequired(xmlDataNode));
					}
				}
			}
		}

		/// <summary>
		/// For a specified data node and schema node, gets data as xml.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="xmlResponseNode">The response node in which the data is returned.</param>
		public static void GetRecordSetAsXML(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlResponseNode)
		{
			GetRecordSetAsXML(xmlDataNode, xmlSchemaNode, xmlResponseNode, null, null);
		}

		/// <summary>
		/// For a specified data node, schema node, filter and ordering, gets data as xml.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="xmlResponseNode">The response node in which the data is returned.</param>
		/// <param name="filter">The filter to add to the SQL where clause.</param>
		/// <param name="orderBy">The ordering to add to the SQL order by clause.</param>	
		public static void GetRecordSetAsXML(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlResponseNode, string filter, string orderBy) 
		{
			using (DataSet dataSet = GetRecordSet(xmlDataNode, xmlSchemaNode, filter, orderBy))
			{
				if (dataSet != null)
				{
					ConvertRecordSetToXML(
						xmlSchemaNode,
						xmlResponseNode,
						dataSet, IsComboLookUpRequired(xmlDataNode));
				}
			}
		}

		/// <summary>
		/// Copies base entity references in a schema node.
		/// </summary>
		/// <param name="xmlSchemaNode">The schema node.</param>
		public void CloneBaseEntityRefs(XmlNode xmlSchemaNode) 
		{
			XmlNode xmlBaseSchemaEntityNode = null;

			foreach (XmlNode xmlSchemaElement in xmlSchemaNode.ChildNodes)
			{
				if (xmlSchemaElement.Attributes.GetNamedItem("DATASRCE") == null && xmlSchemaElement.Attributes.GetNamedItem("ENTITYREF") != null)
				{
					string entityReference = xmlSchemaElement.Attributes.GetNamedItem("ENTITYREF").InnerText;
					if (xmlBaseSchemaEntityNode == null)
					{
						xmlBaseSchemaEntityNode = GetSchema(entityReference);
					}
					else
					{
						if (xmlBaseSchemaEntityNode.Name != entityReference)
						{
							xmlBaseSchemaEntityNode = GetSchema(entityReference);
						}
					}
					if (xmlBaseSchemaEntityNode != null)
					{
						XmlNode xmlBaseSchemaNode = xmlBaseSchemaEntityNode.SelectSingleNode(xmlSchemaElement.Name);
						if (xmlBaseSchemaNode != null)
						{
							foreach (XmlAttribute xmlAttribute in xmlBaseSchemaNode.Attributes)
							{
								if ((xmlAttribute.Name).Substring(0, 3) != "KEY")
								{
									if (xmlSchemaElement.Attributes.GetNamedItem(xmlAttribute.Name) == null)
									{
										xmlSchemaElement.Attributes.SetNamedItem(xmlAttribute.CloneNode(true));
									}
								}
							}
							xmlSchemaElement.Attributes.RemoveNamedItem("ENTITYREF");
						}
					}
				}
			}
		}

		/// <summary>
		/// Gets a <see cref="DataSet"/> object for a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <returns>
		/// The populated <see cref="DataSet"/> object.
		/// </returns>
		public static DataSet GetRecordSet(XmlNode xmlDataNode, XmlNode xmlSchemaNode)
		{
			return GetRecordSet(xmlDataNode, xmlSchemaNode, null, null);
		}

		/// <summary>
		/// Gets a <see cref="DataSet"/> object for a specified data node, schema node, filter and ordering.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="filter">The filter to add to the SQL where clause.</param>
		/// <param name="orderBy">The ordering to add to the SQL order by clause.</param>
		/// <returns>
		/// The populated <see cref="DataSet"/> object.
		/// </returns>
		public static DataSet GetRecordSet(XmlNode xmlDataNode, XmlNode xmlSchemaNode, string filter, string orderBy) 
		{
			DataSet dataSet = null;

			// RF BMIDS00935 (SYS4752)
			string dataSource = XmlAssist.GetAttributeText(xmlSchemaNode, "DATASRCE", "");
			if (dataSource.Length == 0)
			{
				dataSource = xmlSchemaNode.Name;
			}
			// RF BMIDS00935 Start
			// SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
			// This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
			string sqlNoLockText = "";
			if (_databaseProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer && XmlAssist.GetAttributeText(xmlSchemaNode, "SQLNOLOCK", "") == "TRUE")
			{
				sqlNoLockText = " WITH (NOLOCK)";
			}
			// RF BMIDS00935 End
			int databaseRetries = GetDbRetries();
			using (SqlCommand sqlCommand = new SqlCommand())
			{
				string sqlWhereText = "";
				if (xmlDataNode != null)
				{
					int parameterSuffix = 0;
					foreach (XmlAttribute xmlDataAttribute in xmlDataNode.Attributes)
					{
						string xpathExpression = xmlDataAttribute.Name + "[@DATATYPE]";
						XmlNode xmlNode = xmlSchemaNode.SelectSingleNode(xpathExpression);
						if (xmlNode != null)
						{
							if (sqlWhereText.Length != 0)
							{
								sqlWhereText = sqlWhereText + " AND ";
							}
							SqlParameter sqlParameter = GetParameter(xmlNode, xmlDataAttribute, false, ref parameterSuffix);
							sqlCommand.Parameters.Add(sqlParameter);
							sqlWhereText = sqlWhereText + xmlNode.Name + " = @" + sqlParameter.ParameterName;
						}
					}
				}
				// used by comboAssist
				string cmdText = "";
				if (filter != null && filter.Length > 0 && (filter.Substring(0, 6)).ToUpper() == "SELECT")
				{
					cmdText = filter;
				}
				else
				{
					if (filter != null && filter.Length > 0)
					{
						if (sqlWhereText.Length != 0)
						{
							sqlWhereText = sqlWhereText + " AND ";
						}
						sqlWhereText = sqlWhereText + filter;
					}
					if (sqlCommand.Parameters.Count != 0)
					{
						// RF BMIDS00935 Start
						// SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
						cmdText = "SELECT * FROM " + dataSource + sqlNoLockText + " WHERE (" + sqlWhereText + ")";
						// RF BMIDS00935 End
					}
					else
					{
						// RF BMIDS00935 Start
						// SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
						cmdText = "SELECT * FROM " + dataSource + sqlNoLockText;
						// RF BMIDS00935 End
						if (sqlWhereText.Length != 0)
						{
							cmdText = cmdText + " WHERE (" + sqlWhereText + ")";
							// JLD SYS1774
						}
					}
					if (orderBy != null && orderBy.Length > 0)
					{
						cmdText = cmdText + " ORDER BY " + orderBy;
					}
				}

				dataSet = GetRecordSet(cmdText, sqlCommand);
			}

			return dataSet;
		}

		/// <summary>
		/// Gets a <see cref="DataSet"/> object for a specified SQL command text.
		/// </summary>
		/// <param name="cmdText">The SQL command text.</param>
		/// <returns>
		/// The populated <see cref="DataSet"/> object.
		/// </returns>
		public static DataSet GetRecordSet(string cmdText)
		{
			DataSet dataSet = null;

			using (SqlCommand sqlCommand = new SqlCommand())
			{
				dataSet = GetRecordSet(cmdText, sqlCommand);
			}

			return dataSet;
		}

		/// <summary>
		/// Gets a <see cref="DataSet"/> object for a specified SQL command text and <see cref="SqlCommand"/> object.
		/// </summary>
		/// <param name="cmdText">The SQL command text.</param>
		/// <param name="sqlCommand">The <see cref="SqlCommand"/> object.</param>
		/// <returns>
		/// The populated <see cref="DataSet"/> object.
		/// </returns>
		public static DataSet GetRecordSet(string cmdText, SqlCommand sqlCommand)
		{
			System.Diagnostics.Debug.WriteLine("AdoAssist.GetRecordSet SQL: " + cmdText);
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = cmdText;
			DataSet dataSet = ExecuteGetRecordSet(sqlCommand);

			if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0)
			{
				System.Diagnostics.Debug.WriteLine("AdoAssist.GetRecordSet records retrieved: 0");
			}
			else
			{
				System.Diagnostics.Debug.WriteLine("AdoAssist.GetRecordSet records retrieved: " + dataSet.Tables.Count.ToString());
			}

			return dataSet;
		}

		/// <summary>
		/// Converts a <see cref="DataSet"/> to xml, using a specified schema.
		/// </summary>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="xmlResponseNode">The response node which will contain the converted <see cref="DataSet"/>.</param>
		/// <param name="dataSet">The <see cref="DataSet"/>.</param>
		/// <param name="comboLookup">Indicates whether combo value ids should be converted into the corresponding combo values.</param>
		public static void ConvertRecordSetToXML(XmlNode xmlSchemaNode, XmlNode xmlResponseNode, DataSet dataSet, bool comboLookup)
		{
			XmlElement xmlElement = null;

			string nodeName = "";

			if (xmlSchemaNode.Attributes.GetNamedItem("OUTNAME") != null)
			{
				nodeName = xmlSchemaNode.Attributes.GetNamedItem("OUTNAME").InnerText;
			}
			else
			{
				nodeName = xmlSchemaNode.Name;
			}

			try
			{
				if (dataSet.Tables.Count > 0)
				{
					// AS 12/04/2007 Only convert first table.
					DataTable dataTable = dataSet.Tables[0];
					foreach (DataRow dataRow in dataTable.Rows)
					{
						if (xmlResponseNode.Name == nodeName)
						{
							xmlElement = (XmlElement)xmlResponseNode;
						}
						else
						{
							xmlElement = xmlResponseNode.OwnerDocument.CreateElement(nodeName);
						}
						foreach (XmlNode xmlThisSchemaNode in xmlSchemaNode.ChildNodes)
						{
							object fld = null;
							if (XmlAssist.GetAttributeText(xmlSchemaNode, "ENTITYTYPE", "") == "PROCEDURE")
							{
								if (XmlAssist.GetAttributeText(xmlThisSchemaNode, "PARAMMODE", "") != "IN")
								{
									fld = dataRow[xmlThisSchemaNode.Name];
								}
							}
							else
							{
								fld = dataRow[xmlThisSchemaNode.Name];
							}
							if (fld != null && !Convert.IsDBNull(fld))
							{
								FieldToXml(fld, xmlElement, xmlThisSchemaNode, comboLookup);
							}
						}
						if (xmlElement.Attributes.Count > 0)
						{
							xmlResponseNode.AppendChild(xmlElement);
						}
					}
				}
			}
			finally
			{
				dataSet.Dispose();
			}
		}

		/// <summary>
		/// Executes a SQL query represented by a <see cref="SqlCommand"/> object, returning the resulting <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="sqlCommand">The <see cref="SqlCommand"/>.</param>
		/// <returns>The <see cref="DataSet"/>.</returns>
		public static DataSet ExecuteGetRecordSet(SqlCommand sqlCommand) 
		{
			DataSet dataSet = new DataSet();

			using (SqlConnection sqlConnection = new SqlConnection(GetDbConnectString()))
			{
				sqlCommand.Connection = sqlConnection;
				using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
				{
					sqlDataAdapter.Fill(dataSet);
				}
			}

			return dataSet;
		}

		/// <summary>
		/// Executes a SQL query represented by a string, returning the number of rows affected.
		/// </summary>
		/// <param name="cmdText">The SQL query.</param>
		/// <returns>The number of rows affected.</returns>
		/// <remarks>The query will be retried if an exception occurs.</remarks>
		public static int ExecuteSQLCommand(string cmdText)
		{
			return ExecuteSQLCommand(cmdText, true);
		}

		/// <summary>
		/// Executes a SQL query represented by a string, returning the number of rows affected.
		/// </summary>
		/// <param name="cmdText">The SQL query.</param>
		/// <param name="retry">Indicates whether to retry the query if an exception occurs.</param>
		/// <returns>The number of rows affected.</returns>
		public static int ExecuteSQLCommand(string cmdText, bool retry)
		{
			int recordsAffected = 0;

			if (cmdText == null || cmdText.Length == 0)
			{
				throw new ArgumentNullException("cmdText");
			}

			using (SqlCommand sqlCommand = new SqlCommand())
			{
				sqlCommand.CommandText = cmdText;
				sqlCommand.CommandType = CommandType.Text;

				int maxRetries = retry ? GetDbRetries() : 1;
				int retryIndex = 0;
				bool success = false;
				while (!success && retryIndex < maxRetries)
				{
					try
					{
						recordsAffected = ExecuteSqlCommandNonQuery(sqlCommand);
						success = true;
					}
					catch (SqlException exception)
					{
						if (
								exception.Number == SqlServerErrorNumberDuplicateKey ||
								exception.Number == SqlServerErrorNumberChildRecordsFound)
						{
							throw exception;
						}
					}
					catch (System.Data.OleDb.OleDbException exception)
					{
						if (
								exception.Errors[0].NativeError == OracleErrorNumberDuplicateKey ||
								exception.Errors[0].NativeError == OracleErrorNumberChildRecordsFound)
						{
							throw exception;
						}
					}
					retryIndex++;
				}
			}

			if (recordsAffected == 0)
			{
				throw new NoRowsAffectedException();
			}

			return recordsAffected;
		}

		private static SqlParameter AddParameter(XmlNode xmlSchemaNode, XmlAttribute xmlAttribute, SqlCommand sqlCommand, string parameterName)
		{
			SqlParameter sqlParameter = GetParameter(xmlSchemaNode, xmlAttribute, false, parameterName);
			sqlCommand.Parameters.Add(sqlParameter);
			return sqlParameter;
		}

		private static SqlParameter GetParameter(XmlNode xmlSchemaNode, XmlAttribute xmlAttribute, bool isCreate, string parameterName)
		{
			int parameterSuffix = 0;
			return GetParameter(xmlSchemaNode, xmlAttribute, isCreate, parameterName, ref parameterSuffix);
		}

		private static SqlParameter GetParameter(XmlNode xmlSchemaNode, XmlAttribute xmlAttribute, bool isCreate, ref int parameterSuffix)
		{
			return GetParameter(xmlSchemaNode, xmlAttribute, isCreate, null, ref parameterSuffix);
		}

		private static SqlParameter GetParameter(XmlNode xmlSchemaNode, XmlAttribute xmlAttribute, bool isCreate, string parameterName, ref int parameterSuffix) 
		{
			SqlParameter sqlParameter = new SqlParameter(); 

			string dataType = xmlSchemaNode.Attributes.GetNamedItem("DATATYPE").InnerText;
			switch (dataType)
			{
				case "STRING":
					sqlParameter.SqlDbType = SqlDbType.NVarChar;
					sqlParameter.Size = xmlAttribute.InnerText.Length;
					break;
				case "GUID":
					sqlParameter.SqlDbType = SqlDbType.Binary;
					sqlParameter.Size = 16;
					break;
				case "SHORT":
				case "BOOLEAN":
				case "COMBO":
				case "LONG":
					sqlParameter.SqlDbType = SqlDbType.Int;
					break;
				case "DOUBLE":
				case "CURRENCY":
					sqlParameter.SqlDbType = SqlDbType.Float;
					break;
				case "DATE":
				case "DATETIME":
					sqlParameter.SqlDbType = SqlDbType.DateTime;
					break; 
			}
			sqlParameter.Direction = ParameterDirection.Input;
			sqlParameter.Value = System.DBNull.Value;
			if (parameterName != null && parameterName.Length != 0)
			{
				sqlParameter.ParameterName = parameterName.TrimStart(new char[] { '@' });
			}
			else 
			{
				sqlParameter.ParameterName = "parameter" + parameterSuffix.ToString();
				parameterSuffix++;
			}
			if (xmlAttribute != null)
			{
				if (xmlAttribute.InnerText.Length > 0)
				{
					if (dataType != "GUID")
					{
						sqlParameter.Value = xmlAttribute.InnerText;
					}
					else
					{
						sqlParameter.Value = SqlAssist.GuidStringToByteArray(xmlAttribute.InnerText);
					}
				}
			}

			return sqlParameter;
		}

		private static void GetGeneratedKey(XmlNode xmlDataNode, XmlNode xmlSchemaNode)
		{
			XmlAttribute xmlAttribute = null;
			string dataType = XmlAssist.GetAttributeText(xmlSchemaNode, "DATATYPE", "");
			switch (dataType)
			{
				case "GUID":
					xmlAttribute = xmlDataNode.OwnerDocument.CreateAttribute(xmlSchemaNode.Name);
					xmlAttribute.Value = GuidAssist.CreateGUID();
					xmlDataNode.Attributes.SetNamedItem(xmlAttribute);
					xmlAttribute = null;
					break;

				case "SHORT":
					xmlAttribute = xmlDataNode.OwnerDocument.CreateAttribute(xmlSchemaNode.Name);
					xmlAttribute.Value = Convert.ToString(GetNextKeySequence(xmlDataNode, xmlSchemaNode, xmlSchemaNode.ParentNode));
					xmlDataNode.Attributes.SetNamedItem(xmlAttribute);
					xmlAttribute = null;
					break; 
			}
		}

		private static int GetNextKeySequence(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlSchemaParentNode)
		{
			int thisSequence = 0;

			// RF BMIDS00935 (SYS4572)
			string dataSource = XmlAssist.GetAttributeText(xmlSchemaParentNode, "DATASRCE", "");
			if (dataSource.Length == 0)
			{
				dataSource = xmlSchemaParentNode.Name;
			}
			string sequenceFieldName = xmlSchemaNode.Name;
			// RF BMIDS00935 Start
			// SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
			// This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
			string sqlNoLockText = "";
			if ((_databaseProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer) && (XmlAssist.GetAttributeText(xmlSchemaNode, "SQLNOLOCK", "") == "TRUE"))
			{
				sqlNoLockText = " WITH (NOLOCK)";
			}
			// RF BMIDS00935 End
			using (SqlCommand sqlCommand = new SqlCommand())
			{
				int parameterSuffix = 0;
				string sqlWhereText = "";
				foreach (XmlNode xmlNode in xmlSchemaParentNode.ChildNodes)
				{
					if (xmlNode.Attributes.GetNamedItem("KEYTYPE") != null)
					{
						if (xmlNode.Attributes.GetNamedItem("KEYSOURCE") == null)
						{
							if (xmlDataNode.Attributes.GetNamedItem(xmlNode.Name) != null)
							{
								if (sqlWhereText.Length != 0)
								{
									sqlWhereText = sqlWhereText + " AND ";
								}
								SqlParameter sqlParameter = 
									GetParameter(
										xmlNode,
										(XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlNode.Name), 
										false, ref parameterSuffix);
								sqlCommand.Parameters.Add(sqlParameter);
								sqlWhereText = sqlWhereText + xmlNode.Name + " = @" + sqlParameter.ParameterName;
							}
						}
					}
				}
				if (sqlWhereText.Length > 0)
				{
					sqlWhereText = " WHERE (" + sqlWhereText + ")";
				}
				// RF BMIDS00935 Start
				// SYS4752 - Allow schema to specify the SQL-Server (NOLOCK) hint.
				string cmdText = "SELECT MAX(" + sequenceFieldName + ") FROM " + dataSource + sqlNoLockText + sqlWhereText;
				// RF BMIDS00935 End
				System.Diagnostics.Debug.WriteLine("AdoAssist.GetNextKeySequence SQL: " + cmdText);
				sqlCommand.CommandType = CommandType.Text;
				sqlCommand.CommandText = cmdText;
				using (DataSet dataSet = ExecuteGetRecordSet(sqlCommand))
				{
					if (dataSet != null && dataSet.Tables != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows != null && dataSet.Tables[0].Rows.Count > 0 && dataSet.Tables[0].Rows[0].ItemArray != null && dataSet.Tables[0].Rows[0].ItemArray.Length > 0)
					{
						object value = dataSet.Tables[0].Rows[0][0];
						if (Convert.IsDBNull(value) == false)
						{
							thisSequence = Convert.ToInt32(value);
						}
					}
				}
			}

			return ++thisSequence;
		}

		private static string DateToString(DateTime date) 
		{
			return date.ToString("dd/MM/yyyy");
		}

		private static string DateTimeToString(DateTime dateTime) 
		{
			return dateTime.ToString("dd/MM/yyyy HH:mm:ss");
		}

		private static void FieldToXml(object field, XmlElement xmlOutElement, XmlNode xmlSchemaNode, bool comboLookup) 
		{
			string value = null;

			switch (xmlSchemaNode.Attributes.GetNamedItem("DATATYPE").InnerText)
			{
				case "GUID":
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, GuidAssist.ToString((byte[])field));
					break;

				case "CURRENCY":
					NumberFormatInfo nfi = new NumberFormatInfo();
					nfi.NumberDecimalDigits = 2;
					if (field is decimal)
					{
						value = ((decimal)field).ToString("N", nfi);
					}
					else if (field is double)
					{
						value = ((double)field).ToString("N", nfi);
					}
					else if (field is float)
					{
						value = ((float)field).ToString("N", nfi);
					}
					else
					{
						value = field.ToString();
					}
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, value);
					break;

				case "DOUBLE":
					string formatMask = "";
					if (xmlSchemaNode.Attributes.GetNamedItem("FORMATMASK") != null)
					{
						formatMask = xmlSchemaNode.Attributes.GetNamedItem("FORMATMASK").InnerText;
					}
					if (formatMask.Length != 0)
					{
						value = field.ToString();
						xmlOutElement.SetAttribute(xmlSchemaNode.Name + "_RAW", value);
						if (field is decimal)
						{
							value = ((decimal)field).ToString(formatMask);
						}
						else if (field is double)
						{
							value = ((double)field).ToString(formatMask);
						}
						else if (field is float)
						{
							value = ((float)field).ToString(formatMask);
						}
						else
						{
							value = field.ToString();
						}
						xmlOutElement.SetAttribute(xmlSchemaNode.Name, value);
					}
					else
					{
						xmlOutElement.SetAttribute(xmlSchemaNode.Name, field.ToString());
					}
					break;

				case "COMBO":
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, field.ToString());
					if (comboLookup)
					{
						if (xmlSchemaNode.Attributes.GetNamedItem("COMBOGROUP") != null)
						{
							string strComboGroup = xmlSchemaNode.Attributes.GetNamedItem("COMBOGROUP").InnerText;
							string strComboValue = ComboAssist.GetComboText(strComboGroup, Convert.ToInt32(field));
							xmlOutElement.SetAttribute(xmlSchemaNode.Name + "_TEXT", strComboValue);
						}
					}
					break;

				case "DATE":
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, DateToString((DateTime)field));
					break;

				case "DATETIME":
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, DateTimeToString((DateTime)field));
					break;

				default:
					xmlOutElement.SetAttribute(xmlSchemaNode.Name, field.ToString());
					break; 
			}
		}

		/// <summary>
		/// Deletes a record based on a specified data node and schema.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="schemaName">The schema name.</param>
		/// <returns>
		/// The number of records deleted.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// <paramref name="xmlDataNode"/> does not 
		/// contain an attribute for every primary key node in <paramref name="schemaName"/>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// <paramref name="xmlDataNode"/> 
		/// contains an attribute with an empty value for a primary key node in <paramref name="schemaName"/>.
		/// </exception>
		public int DeleteFromNode(XmlNode xmlDataNode, string schemaName)
		{
			return Delete(xmlDataNode, GetSchema(schemaName), true);
		}
		
		/// <summary>
		/// Delete one or more records based on a specified data node and schema.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="schemaName">The schema name.</param>
		/// <param name="deleteSingleRecordOnly">Indicates whether only a single record should be deleted.</param>
		/// <returns>
		/// The number of records deleted.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// <paramref name="deleteSingleRecordOnly"/> is <b>true</b> and <paramref name="xmlDataNode"/> does not 
		/// contain an attribute for every primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// <paramref name="deleteSingleRecordOnly"/> is <b>true</b> and <paramref name="xmlDataNode"/> 
		/// contains an attribute with an empty value for a primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		public int DeleteFromNode(XmlNode xmlDataNode, string schemaName, bool deleteSingleRecordOnly) 
		{
			return Delete(xmlDataNode, GetSchema(schemaName), deleteSingleRecordOnly);
		}

		/// <summary>
		/// Deletes a record based on a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <returns>
		/// The number of records deleted.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// <paramref name="xmlDataNode"/> does not 
		/// contain an attribute for every primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// <paramref name="xmlDataNode"/> 
		/// contains an attribute with an empty value for a primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		public static int Delete(XmlNode xmlDataNode, XmlNode xmlSchemaNode)
		{
			return Delete(xmlDataNode, xmlSchemaNode, true);
		}

		/// <summary>
		/// Deletes one or more records based on a specified data node and schema node.
		/// </summary>
		/// <param name="xmlDataNode">The data node.</param>
		/// <param name="xmlSchemaNode">The schema node.</param>
		/// <param name="deleteSingleRecordOnly">Indicates whether only a single record should be deleted.</param>
		/// <returns>
		/// The number of records deleted.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// <paramref name="deleteSingleRecordOnly"/> is <b>true</b> and <paramref name="xmlDataNode"/> does not 
		/// contain an attribute for every primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// <paramref name="deleteSingleRecordOnly"/> is <b>true</b> and <paramref name="xmlDataNode"/> 
		/// contains an attribute with an empty value for a primary key node in <paramref name="xmlSchemaNode"/>.
		/// </exception>
		public static int Delete(XmlNode xmlDataNode, XmlNode xmlSchemaNode, bool deleteSingleRecordOnly) 
		{
			int recordsAffected = 0;

			using (SqlCommand sqlCommand = new SqlCommand())
			{
				// Get the table name from the schema
				string dataSource = XmlAssist.GetAttributeText(xmlSchemaNode, "DATASRCE", "");
				if (dataSource.Trim().Length == 0)
				{
					throw new XMLMissingAttributeException("DATASRCE");
				}
				// If specified that only a single record may be deleted, ensure that
				// the whole primary key has been provided. Otherwise raise an error.
				if (deleteSingleRecordOnly)
				{
					XmlNodeList xmlSchemaNodeList = xmlSchemaNode.SelectNodes("*[@KEYTYPE='PRIMARY']");
					foreach (XmlNode xmlNode in xmlSchemaNodeList)
					{
						XmlAttribute xmlDataAttribute = (XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlNode.Name);
						if (xmlDataAttribute == null)
						{
							throw new XMLMissingAttributeException(xmlNode.Name);
						}
						else if (xmlDataAttribute.Value.Length == 0)
						{
							throw new XMLInvalidAttributeValueException(xmlNode.Name);
						}
					}
				}
				// Get each condition in the WHERE clause
				string sqlWhereText = "";
				int parameterSuffix = 0;
				foreach (XmlNode xmlNode in xmlSchemaNode.ChildNodes)
				{
					XmlAttribute xmlDataAttribute = (XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlNode.Name);
					if (xmlDataAttribute != null && xmlDataAttribute.Value.Length != 0)
					{
						if (sqlWhereText.Length != 0)
						{
							sqlWhereText = sqlWhereText + " AND ";
						}
						SqlParameter sqlParameter = GetParameter(xmlNode, xmlDataAttribute, false, ref parameterSuffix);
						sqlCommand.Parameters.Add(sqlParameter);
						sqlWhereText = sqlWhereText + xmlNode.Name + " = @" + sqlParameter.ParameterName;
					}
				}
				// Run the SQL command
				if (sqlWhereText.Length > 0)
				{
					string cmdText = "DELETE FROM " + dataSource + " WHERE " + sqlWhereText;
					System.Diagnostics.Debug.WriteLine("AdoAssist.Delete SQL: " + cmdText);
					sqlCommand.CommandType = CommandType.Text;
					sqlCommand.CommandText = cmdText;
					recordsAffected = ExecuteSqlCommandNonQuery(sqlCommand);
					System.Diagnostics.Debug.WriteLine("AdoAssist.Delete records deleted: " + recordsAffected);
				}
			}

			return recordsAffected;
		}

		/// <summary>
		/// Populates child keys in destination node from a source node.
		/// </summary>
		/// <param name="xmlSourceNode">The source node.</param>
		/// <param name="xmlDestinationNode">The destination node.</param>
		public void PopulateChildKeys(XmlNode xmlSourceNode, XmlNode xmlDestinationNode) 
		{
			if (GetSchema(xmlSourceNode.Name) != null)
			{
				XmlNode xmlDestSchemaNode = GetSchema(xmlDestinationNode.Name);
				if (xmlDestSchemaNode != null)
				{
					XmlNodeList xmlPrimaryKeyNodeList = xmlDestSchemaNode.SelectNodes("*[@KEYTYPE='PRIMARY']");
					foreach (XmlNode xmlNode in xmlPrimaryKeyNodeList)
					{
						XmlAssist.CopyAttribute(xmlSourceNode, xmlDestinationNode, xmlNode.Name);
					}
				}
			}
		}

		private static void BuildDbConnectionString() 
		{
			_databaseConnectionString = GlobalProperty.DatabaseConnectionString;
			_databaseRetries = GlobalProperty.DatabaseRetries;
			string provider = GlobalProperty.DatabaseProvider;

			switch (provider)
			{
				case "MSDAORA":
					_databaseProvider = DBPROVIDER.omiga4DBPROVIDERMSOracle;
					break;
				case "ORAOLEDB.ORACLE":
					_databaseProvider = DBPROVIDER.omiga4DBPROVIDEROracle;
					break;
				case "SQLOLEDB":
					_databaseProvider = DBPROVIDER.omiga4DBPROVIDERSQLServer;
					break;
			}
		}

		/// <summary>
		/// Gets the database connection string from the registry.
		/// </summary>
		/// <returns>
		/// The database connection string.
		/// </returns>
		/// <remarks>
		/// The database connection string is configured by registry keys:
		/// <code>
		/// HKLM/SOFTWARE/Omiga4/Database Connection/Provider
		/// e.g. "MSDAORA" for Oracle or "SQLOLEDB" for SQL Server
		/// HKLM/SOFTWARE/Omiga4/Database Connection/Server
		/// database server, e.g. "MSGCH5209" or "OMDBSRV"
		/// HKLM/SOFTWARE/Omiga4/Database Connection/Database Name
		/// e.g. "OmigaSystemTest"
		/// HKLM/SOFTWARE/Omiga4/Database Connection/User ID
		/// e.g. "production"
		/// HKLM/SOFTWARE/Omiga4/Database Connection/Password
		/// e.g. "production"
		/// </code>
		/// </remarks>
		public static string GetDbConnectString() 
		{
			if (_databaseConnectionString == null || _databaseConnectionString.Length == 0)
			{
				// _databaseConnectionString is not initialised, so initialise it.
				BuildDbConnectionString();
			}
			return _databaseConnectionString;
		}

		/// <summary>
		/// Gets the number of database retries from the registry.
		/// </summary>
		/// <returns>
		/// The number of database retries.
		/// </returns>
		public static int GetDbRetries() 
		{
			GetDbConnectString();	// Ensure _databaseRetries is initialised.
			return _databaseRetries;
		}

		/// <summary>
		/// Gets the database provider from the registry.
		/// </summary>
		/// <returns>
		/// The database provider.
		/// </returns>
		public static DBPROVIDER GetDbProvider() 
		{
			GetDbConnectString();	// Ensure _databaseProvider is initialised.
			return _databaseProvider;
		}

		private static int GetErrorNumber(string errorDescription) 
		{
			string error = "";

			switch (_databaseProvider)
			{
				// SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					// oracle errors have format "ORA-nnnnn"
					if (errorDescription.Length > 10)
					{
						error = errorDescription.Substring(5 - 1, 5);
					}
					else
					{
						// "Cannot interpret database error"
						throw new UnspecifiedErrorException();
					}
					break;

				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					// SQL Server
					throw new NotImplementedException();

				case DBPROVIDER.omiga4DBPROVIDERUnknown:
					throw new NotImplementedException();
			}

			return Convert.ToInt32(error);
		}

		private static int GetOmigaNumberForDatabaseError(string errorDescription) 
		{
			int omigaErrorNumber = 0;
			// -----------------------------------------------------------------------------------
			// Description : Find the omiga equivalent number for a database error. This is used
			// to trap specific errors. Add to the list below, if you want trap a
			// a new error
			// Pass		: errorDescription : Description of the Error Message (from database )
			// ------------------------------------------------------------------------------------

			int errorNumber = GetErrorNumber(errorDescription);
			// 'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
			if (_databaseProvider == DBPROVIDER.omiga4DBPROVIDEROracle || _databaseProvider == DBPROVIDER.omiga4DBPROVIDERMSOracle)
			{
				switch (errorNumber)
				{
					case 1:
						omigaErrorNumber = (int)OMIGAERROR.DuplicateKey;
						break;
					case 2292:
						omigaErrorNumber = (int)OMIGAERROR.ChildRecordsFound;
						break;
					default:
						omigaErrorNumber = (int)OMIGAERROR.UnspecifiedError;
						break; 
				}
			}
			else
			{
				throw new NotImplementedException("Error trapping for this database engine");
			}
			return omigaErrorNumber;
		}

		private static bool GetValidRecordset(SqlDataReader sqlDataReader)
		{
			return GetValidRecordset(sqlDataReader, 1);
		}

		private static bool GetValidRecordset(SqlDataReader sqlDataReader, int resultIndex) 
		{
			bool success = false;

			if (_databaseProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer)
			{
				bool more = true;
				while (sqlDataReader != null && more)
				{
					if (!sqlDataReader.IsClosed && sqlDataReader.HasRows)
					{
						if (resultIndex > 0)
						{
							--resultIndex;
						}
						success = true;
						if (resultIndex == 0)
						{
							break;
						}
					}
					more = sqlDataReader.NextResult();
				} 
			}
			else if (sqlDataReader != null && !sqlDataReader.IsClosed && sqlDataReader.HasRows)
			{
				// Found an open recordset with records.
				success = true;
			}
			return success;
		}

		private static bool IsComboLookUpRequired(XmlNode xmlRequestNode) 
		{
			bool isComboLookUpRequired = false;
			if (xmlRequestNode.Attributes.GetNamedItem("_COMBOLOOKUP_") != null)
			{
				isComboLookUpRequired = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "_COMBOLOOKUP_", "");
			}
			else
			{
				if (xmlRequestNode.OwnerDocument.FirstChild != null)
				{
					isComboLookUpRequired = XmlAssist.GetAttributeAsBoolean(xmlRequestNode.OwnerDocument.FirstChild, "_COMBOLOOKUP_", "");
				}
				else
				{
					isComboLookUpRequired = false;
				}
			}
			return isComboLookUpRequired;
		}

		#endregion adoAssistEx


		#region ADOAssist

		/// <summary>
		/// Gets the first valid <see cref="DataTable"/> from a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <param name="dataTable">The returned <see cref="DataTable"/> object.</param>
		/// <returns><b>true</b> if a valid <see cref="DataTable"/> is returned, otherwise <b>false</b>.</returns>
		/// <remarks>
		/// This method is useful when a SQL query may return multiple tables in a data set,
		/// only one of which has rows. <i>This method is probably redundant in .Net.</i>
		/// </remarks>
		public static bool GetValidRecordset(DataSet dataSet, out DataTable dataTable)
		{
			return GetValidRecordset(dataSet, 1, out dataTable);
		}

		/// <summary>
		/// Gets a valid <see cref="DataTable"/> from a specified <see cref="DataSet"/> object.
		/// </summary>
		/// <param name="dataSet">The <see cref="DataSet"/> object.</param>
		/// <param name="tableIndex">The n'th valid <see cref="DataTable"/> to return. Valid means the <see cref="DataTable"/> has rows.</param>
		/// <param name="dataTable">The returned <see cref="DataTable"/> object.</param>
		/// <returns><b>true</b> if a valid <see cref="DataTable"/> is returned, otherwise <b>false</b>.</returns>
		/// <remarks>
		/// This method is useful when a SQL query may return multiple tables in a data set,
		/// only one of which has rows. <i>This method is probably redundant in .Net.</i>
		/// </remarks>
		public static bool GetValidRecordset(DataSet dataSet, int tableIndex, out DataTable dataTable)
		{
			dataTable = null;
			foreach (DataTable tempDataTable in dataSet.Tables)
			{
				if (tempDataTable.Rows != null && tempDataTable.Rows.Count > 0)
				{
					if (tableIndex > 0)
					{
						tableIndex--;
					}
					if (tableIndex == 0)
					{
						dataTable = tempDataTable;
						break;
					}
				}
			}

			return dataTable != null;
		}

		/// <summary>
		/// Indicates whether an exception object is for a unique key constraint violation.
		/// </summary>
		/// <param name="exception">The exception object.</param>
		/// <returns>
		/// <b>true</b> if <paramref name="exception"/> is for a unique key constraint violation.
		/// </returns>
		public static bool IsUniqueKeyConstraint(Exception exception)
		{
			bool isUniqueKeyConstraint = false;

			switch (GetDbProvider())
			{
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					if (exception is System.Data.OleDb.OleDbException)
					{
						isUniqueKeyConstraint = ((System.Data.OleDb.OleDbException)exception).Errors[0].NativeError == OracleErrorNumberDuplicateKey;
					}
					break;

				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					if (exception is SqlException)
					{
						isUniqueKeyConstraint = ((SqlException)exception).Number == SqlServerErrorNumberDuplicateKey;
					}
					break;

				default:
					throw new InvalidOperationException("Invalid DBPROVIDER: " + _databaseProvider.ToString() + ".", exception);
			}

			return isUniqueKeyConstraint;
		}

		/// <summary>
		/// Indicates whether an exception object is for a database connection error.
		/// </summary>
		/// <param name="exception">The exception object.</param>
		/// <returns>
		/// <b>true</b> if <paramref name="exception"/> is for a database connection error.
		/// </returns>
		public static bool IsDBConnectionError(Exception exception)
		{
			bool isDBConnectionError = false;

			switch (GetDbProvider())
			{
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					/*
					// DM SYS2321
					// We can use the GetErrorNumber function for Oracle but not for SQLServer
					errorNumber = GetErrorNumber(Convert.ToString(vstrErrDescription));

					// ORA-00018   maximum number of sessions exceeded
					// ORA-00019   maximum number of session licenses exceeded
					// ORA-00020   maximum number of processes num exceeded
					if (errorNumber >= 18 && errorNumber <= 20)
					{
						blnConnError = true;
					}
					*/
					break;
				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					throw new NotImplementedException();
				default:
					// "Invalid database engine type"
					throw new InternalErrorException();
			}

			return isDBConnectionError;
		}

		/// <summary>
		/// Applies a lock to a database table.
		/// </summary>
		/// <param name="tableName">The name of the table.</param>
		/// <param name="lockType">The type of lock to apply.</param>
		/// <remarks>
		/// See <see cref="LockType"/> for a description of the different types of lock.
		/// </remarks>
		public static void LockTable(string tableName, LockType lockType)
		{
			string cmdText = "";
			string lockStatement = "";

			switch (GetDbProvider())
			{
				case DBPROVIDER.omiga4DBPROVIDERSQLServer:

					switch (lockType)
					{
						case LockType.lckShare:
							// Allow others to read table, but prevent them from updating it;
							// lock held until end of transaction.
							lockStatement = "tablock holdlock";
							break;
						case LockType.lckUpdate:
							// Allow transaction to read data (without blocking other readers) and update it
							// later with the assurance that the data will not have changed since it was read
							// lock held until end of transaction.
							lockStatement = "updlock holdlock";
							break;
						case LockType.lckExclusive:
							// Prevent others from reading or updating the table;
							// lock held until end of transaction.
							lockStatement = "tablockx holdlock";
							break;
						default:
							throw new ArgumentException("Invalid lock type.", lockType.ToString());
					}
					// Dummy select on table to lock it (with no overhead of reading records from table).
					cmdText = "select count(*) from " + tableName + " (" + lockStatement + ")";
					break;

				case DBPROVIDER.omiga4DBPROVIDEROracle:
					switch (lockType)
					{
						case LockType.lckShare:
							// Allow others to read table, but prevent them from updating it;
							// lock held until end of transaction.
							lockStatement = "share";
							break;
						case LockType.lckUpdate:
							// Allow others to query, but not update the table;
							// lock held until end of transaction.
							lockStatement = "share row exclusive";
							break;
						case LockType.lckExclusive:
							// Prevent others from locking the table
							// lock held until end of transaction.
							lockStatement = "exclusive";
							break;
						default:
							throw new ArgumentException("Invalid lock type.", lockType.ToString());
					}
					// Explicit lock statement.
					cmdText = "lock table " + tableName + " in " + lockStatement + " mode";
					break;
				default:
					throw new InvalidOperationException("Invalid DBPROVIDER: " + _databaseProvider.ToString() + ".");
			}
			ExecuteSQLCommand(cmdText, true);
		}

		/// <summary>
		/// Checks that one and only one record exists in a specified table.
		/// </summary>
		/// <param name="tableName">The table to check.</param>
		/// <param name="criteria">The selection criteria for the record.</param>
		/// <remarks>
		/// The criteria should be in SQL where clause format, e.g., "column1 = value1 and column2 = value2". 
		/// The values should already be correctly formatted for the associated data type.
		/// </remarks>
		public static void CheckSingleRecordExists(string tableName, string criteria)
		{
			bool isSingleRecord = GetNumberOfRecords(tableName, criteria) == 1;

			if (!isSingleRecord)
			{
				throw new RecordNotFoundException();
			}
		}

		/// <summary>
		/// Gets the number of records in a table that match specified criteria.
		/// </summary>
		/// <param name="tableName">The name of the table.</param>
		/// <param name="criteria">
		/// The criteria. The format is: "FieldA = ValueX and FieldB = ValueY
		/// and FieldC = ValueZ". The field values are assumed to have
		/// been formatted appropriately for their associated datatype.
		/// </param>
		/// <returns>
		/// The number of records in <paramref name="tableName"/> that match <paramref name="criteria"/>.
		/// </returns>
		public static int GetNumberOfRecords(string tableName, string criteria)
		{
			int numberOfRecords = 0;

			string cmdText = "select count(*) Count from " + tableName + " with(nolock) where " + criteria;

			using (DataSet dataSet = GetRecordSet(cmdText))
			{
				if (
						dataSet != null &&
						dataSet.Tables != null &&
						dataSet.Tables.Count == 1 &&
						dataSet.Tables[0].Rows != null &&
						dataSet.Tables[0].Rows.Count == 1 &&
						dataSet.Tables[0].Rows[0].ItemArray != null &&
						dataSet.Tables[0].Rows[0].ItemArray.Length == 1)
				{
					numberOfRecords = Convert.ToInt32(dataSet.Tables[0].Rows[0].ItemArray[0]);
				}
			}

			return numberOfRecords;
		}

		/// <summary>
		/// Indicates whether a record exists on a specified table and matching specified criteria.
		/// </summary>
		/// <param name="tableName">The name of the table.</param>
		/// <param name="criteria">
		/// The criteria. The format is: "FieldA = ValueX and FieldB = ValueY
		/// and FieldC = ValueZ". The field values are assumed to have
		/// been formatted appropriately for their associated datatype.
		/// </param>
		/// <returns>
		/// <b>true</b> if the record exists, otherwise <b>false</b>.
		/// </returns>
		public static bool CheckRecordExists(string tableName, string criteria)
		{
			return GetNumberOfRecords(tableName, criteria) > 0;
		}

		/// <summary>
		/// Gets the value for a specified column in the first record in a specified table that matches the specified criteria.
		/// </summary>
		/// <param name="tableName">The name of the table.</param>
		/// <param name="criteria">
		/// The criteria. The format is: "FieldA = ValueX and FieldB = ValueY
		/// and FieldC = ValueZ". The field values are assumed to have
		/// been formatted appropriately for their associated datatype.
		/// </param>
		/// <param name="column">The name of the column.</param>
		/// <returns>
		/// The value in the column.
		/// </returns>
		public static object GetValueFromTable(string tableName, string criteria, string column)
		{
			object returnValue = null;
			bool recordExists = true;
			return GetValueFromTable(tableName, criteria, column, out returnValue, out recordExists);
		}

		/// <summary>
		/// Gets the value for a specified column in the first record in a specified table that matches the specified criteria.
		/// </summary>
		/// <param name="tableName">The name of the table.</param>
		/// <param name="criteria">
		/// The criteria. The format is: "FieldA = ValueX and FieldB = ValueY
		/// and FieldC = ValueZ". The field values are assumed to have
		/// been formatted appropriately for their associated datatype.
		/// </param>
		/// <param name="column">The name of the column.</param>
		/// <param name="returnValue">The value in the column.</param>
		/// <param name="recordExists">Indicates whether a matching record exists.</param>
		/// <returns>
		/// The value in the column.
		/// </returns>
		public static object GetValueFromTable(string tableName, string criteria, string column, out object returnValue, out bool recordExists)
		{
			object value = null;
			DataSet dataSet = null;
			recordExists = false;
			returnValue = null;

			try
			{
				string cmdText = "select " + column + " from " + tableName + " with(nolock) where " + criteria;
				dataSet = GetRecordSet(cmdText);
				if (
						dataSet != null &&
						dataSet.Tables != null &&
						dataSet.Tables.Count > 0 &&
						dataSet.Tables[0].Rows != null &&
						dataSet.Tables[0].Rows.Count > 0 &&
						dataSet.Tables[0].Rows[0].ItemArray != null && 
						dataSet.Tables[0].Rows[0].ItemArray.Length > 0)
				{
					recordExists = true;

					value = dataSet.Tables[0].Rows[0][0];
					if (value != System.DBNull.Value)
					{
						// SYS2695 - DRC added the check for adBinary
						if (dataSet.Tables[0].Columns[0].DataType == typeof(byte[]))
						{
							// If the value is a GUID convert to a string.
							value = GuidAssist.ToString((byte[])value);
						}
						returnValue = value;
					}
					else
					{
						value = null;
					}
				}
			}
			finally
			{
				if (dataSet != null) dataSet.Dispose();
			}

			return value;
		}

		/// <summary>
		/// Writes a specified SQL string to the file"C:\OutputSQL.log".
		/// </summary>
		/// <param name="cmdText">The SQL string.</param>
		public static void WriteSQLToFile(string cmdText)
		{
			WriteSQLToFile(cmdText, "");
		}

		/// <summary>
		/// Writes a specified SQL string to the file"C:\OutputSQL.log".
		/// </summary>
		/// <param name="cmdText">The SQL string.</param>
		/// <param name="functionName">The name of the function executing the SQL.</param>
		public static void WriteSQLToFile(string cmdText, string functionName)
		{
			if (IsSQLLoggingOn())
			{
				System.Diagnostics.Debug.WriteLine("AdoAssist.WriteSQLToFile SQL: " + cmdText);
				using (StreamWriter streamWriter = new StreamWriter("C:\\OutputSQL.log", true))
				{
					streamWriter.WriteLine("");
					streamWriter.WriteLine("");
					streamWriter.WriteLine("********************************************************************");
					streamWriter.WriteLine(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "\t(" + functionName + ")");
					streamWriter.WriteLine(cmdText);
					streamWriter.WriteLine("********************************************************************");
				}
			}
		}

		/// <summary>
		/// Starts SQL logging.
		/// </summary>
		public static void StartSQLLogging()
		{
			SetSQLLoggingFlag(true);
		}

		/// <summary>
		/// Stops SQL logging.
		/// </summary>
		public static void StopSQLLogging()
		{
			SetSQLLoggingFlag(false);
		}

		private static void SetSQLLoggingFlag(bool loggingOn)
		{
			GlobalProperty.SetSharedPropertyValue("SqlLogging", loggingOn);
		}

		private static bool IsSQLLoggingOn()
		{
			return Convert.ToBoolean(GlobalProperty.GetSharedPropertyValue("SqlLogging"));
		}

		#endregion ADOAssist


		#region Obsolete

		#pragma warning disable 1591

		[Obsolete("Use AdoAssist.ExecuteGetRecordSet instead")]
		public static DataSet executeGetRecordSet(SqlCommand sqlCommand)
		{
			return ExecuteGetRecordSet(sqlCommand);
		}

		[Obsolete("Use GuidAssist.ToString() instead")]
		public static string GuidToString(byte[] bytes)
		{
			return GuidAssist.ToString(bytes);
		}

		[Obsolete("Use SqlAssist.GuidStringToByteArray instead")]
		public static byte[] GuidStringToByteArray(string guidText)
		{
			return SqlAssist.GuidStringToByteArray(guidText);
		}

		[Obsolete("Use adoAssistEx.IsUniqueKeyConstraint instead")]
		public static bool IsThisUniqueKeyConstraint(Exception exception)
		{
			return IsUniqueKeyConstraint(exception);
		}

		#pragma warning restore 1591

		#endregion
	}
}

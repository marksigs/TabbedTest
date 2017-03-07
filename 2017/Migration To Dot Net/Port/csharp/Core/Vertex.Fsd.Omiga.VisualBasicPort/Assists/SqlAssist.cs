/*
--------------------------------------------------------------------------------------------
Workfile:			SqlAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Helper object for handling database-specific formatting when building SQL statements
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		16/07/1999	Created
MCS		02/09/1999	GetSystemDate added
MCS		04/10/1999	GetFormattedValue modified to raise error
RF		22/10/1999	Fixed DateTimeToString
DJP		25/04/2000	Added TruncateDateColumn
SR		02/05/2000	Modified method FormatWildCardedString, FormatString 
					(Added new parameter isInputToADOCommandObject)
APS		08/09/2000	Corrected the SQL Server CONVERT statement to convert dates into ANSI format
DJP		23/01/2001	Added GetColumnDayOfWeek
AS		05/03/2001	CC012 Added GUID support for SQL Server and Oracle
DM		14/03/2001	SYS1949 Removed raise error from SQLServer side of run time switch.
AS		13/11/2003	CORE1 Removed GENERIC_SQL.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from SqlAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Xml;
using omBase;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	#region Enumerated types

	// Missing XML comment for publicly visible type or member.
	#pragma warning disable 1591

	/// <summary>
	/// Constants that define different database types.
	/// </summary>
	public enum DBDATATYPE 
	{
		dbdtNotStored = 0,
		dbdtString = 1,
		dbdtInt = 2,
		dbdtDouble = 3,
		dbdtDate = 4,
		dbdtGuid = 5,
		dbdtCurrency = 6,
		dbdtComboId = 7,
		dbdtBoolean = 8,
		dbdtLong = 9,
		dbdtDateTime = 10,
	}

	public enum DATETIMEFORMAT
	{
		dtfDateTime = 0,
		// date and time
		dtfDate = 1,
		// date only
		dtfTime = 2,
		// time only
	}
	
	public enum GUIDSTYLE 
	{
		guidBinary = 0,
		guidHyphen = 1,
		guidLiteral = 2,
	}

	// Missing XML comment for publicly visible type or member.
	#pragma warning restore 1591

	#endregion Enumerated types


	/// <summary>
	/// Helper type for database-specific formatting when building SQL statements.
	/// </summary>
	/// <remarks>
	/// You should avoid using the SqlAssist class by either putting SQL into stored procedures, 
	/// or by creating parameter objects for the parameters in the SQL.
	/// </remarks>
	public static class SqlAssist
	{
		private const string _apostrophe = "'";
		private const string _wildCard = "*";

		/// <summary>
		/// Converts a guid string from one format to another.
		/// </summary>
		/// <param name="sourceGuid">The guid string to be converted.</param>
		/// <returns>The converted guid string.</returns>
		/// <remarks>
		/// Converts from either guidBinary or 
		/// guidHyphen format to a format that can be used as a literal input parameter for the
		/// database.
		/// </remarks>
		public static string FormatGuid(string sourceGuid)
		{
			return FormatGuid(sourceGuid, GUIDSTYLE.guidLiteral);
		}

		/// <summary>
		/// Converts a guid string from one format to another.
		/// </summary>
		/// <param name="sourceGuid">The guid string to be converted.</param>
		/// <param name="guidStyle">The style of the return string.</param>
		/// <returns>The converted guid string.</returns>
		/// <remarks>
		/// A <i>guidStyle</i> of GUIDSTYLE.guidHyphen converts from "DA6DA163412311D4B5FA00105ABB1680" to
		/// "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format. A <i>guidStyle</i> of GUIDSTYLE.guidBinary 
		/// converts from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680" 
		/// format. A <i>guidStyle</i> of GUIDSTYLE.guidLiteral converts from either guidBinary or 
		/// guidHyphen format to a format that can be used as a literal input parameter for the
		/// database. 
		/// </remarks>
		public static string FormatGuid(string sourceGuid, GUIDSTYLE guidStyle) 
		{
			// By default return string unchanged, i.e., if style is guidHyphen but the source string
			// is currently not in binary format, then it will be returned unchanged (perhaps it is
			// already in hyphenated format?).
			string targetGuid = sourceGuid;

			if (guidStyle == GUIDSTYLE.guidHyphen)
			{
				// Convert from "DA6DA163412311D4B5FA00105ABB1680" to "{63A16DDA-2341-D411-B5FA-00105ABB1680}"
				System.Diagnostics.Debug.Assert(sourceGuid.Length == 32);
				if (sourceGuid.Length == 32)
				{
					targetGuid = "{";
					for (int index1 = 0; index1 < 16; index1++)
					{
						int index2 = ConvertGuidIndex(index1);
						targetGuid = targetGuid + sourceGuid.Substring((index2 * 2) + 1 - 1, 2);
						if (index1 == 3 || index1 == 5 || index1 == 7 || index1 == 9)
						{
							targetGuid = targetGuid + "-";
						}
					}
					targetGuid = targetGuid + "}";
				}
				else
				{
				   throw new InvalidParameterException("sourceGuid length must be 32");
				}
			}
			else if (guidStyle == GUIDSTYLE.guidBinary)
			{
				// Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
				System.Diagnostics.Debug.Assert(sourceGuid.Length == 38);
				if (sourceGuid.Length == 38)
				{
					targetGuid = "";
					int intOffset = 2;
					for (int index1 = 0; index1 < 16; index1++)
					{
						int index2 = ConvertGuidIndex(index1);
						targetGuid = targetGuid + sourceGuid.Substring((index2 * 2) + intOffset - 1, 2);
						if (index1 == 3 || index1 == 5 || index1 == 7 || index1 == 9)
						{
							intOffset++;
						}
					}
				}
				else
				{
					throw new InvalidParameterException("sourceGuid length must be 38");
				}
			}
			else if (guidStyle == GUIDSTYLE.guidLiteral)
			{
				// Convert guid into a format that can be used as a literal input parameter to the database.
				// This assumes that the database type is raw for Oracle, or binary/varbinary for SQL Server.
				if (sourceGuid.Length == 38)
				{
					// Guid is in hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}",
					// so convert to binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
					sourceGuid = FormatGuid(sourceGuid, GUIDSTYLE.guidBinary);
				}
				System.Diagnostics.Debug.Assert(sourceGuid.Length == 32);
				if (sourceGuid.Length == 32)
				{
					switch (AdoAssist.GetDbProvider())
					{
						case DBPROVIDER.omiga4DBPROVIDERSQLServer:
							// e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
							targetGuid = "0x" + sourceGuid;
							break;
						case DBPROVIDER.omiga4DBPROVIDEROracle:
						case DBPROVIDER.omiga4DBPROVIDERMSOracle:
							// e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
							targetGuid = "HEXTORAW('" + sourceGuid + "')";
							break; 
					}
				}
			}
			else
			{
				System.Diagnostics.Debug.Assert(false);
				throw new InvalidParameterException("guidText length must be 32");
			}
			return targetGuid;
		}

		/// <summary>
		/// Formats a string into the appropriate SQL for a string datatype.
		/// </summary>
		/// <param name="cmdText">The string to format.</param>
		/// <returns>The formatted string.</returns>
		public static string FormatString(string cmdText)
		{
			return FormatString(cmdText, false);
		}

		/// <summary>
		/// Formats a string into the appropriate SQL for a string datatype.
		/// </summary>
		/// <param name="cmdText">The string to format.</param>
		/// <param name="isInputToADOCommandObject">true if the string should not be formatted because it will be bound to a parameter object.</param>
		/// <returns>The formatted string.</returns>
		public static string FormatString(string cmdText, bool isInputToADOCommandObject) 
		{
			string formatString = String.Empty;
			string strDelimiter = GetStringDelimiter();

			if (!isInputToADOCommandObject)
			{
				formatString = strDelimiter + CheckApostrophes(cmdText) + strDelimiter;
			}
			else
			{
				// Do not check for aposterphes and do not add delimiters
				formatString = cmdText;
			}

			return formatString;		   
		}

		/// <summary>
		/// Converts a string into appropriate SQL, including the substitution of any wildcard character into the correct SQL syntax.
		/// </summary>
		/// <param name="cmdText">The string to convert.</param>
		/// <param name="foundWildCard">On return, will be set to true if a wildcard was found.</param>
		/// <returns>The converted string.</returns>
		/// <remarks>
		/// For example, "Smit*" will be converted to "'Smit%'". 
		/// </remarks>
		public static string FormatWildcardedString(string cmdText, out bool foundWildCard)
		{
			return FormatWildcardedString(cmdText, out foundWildCard, false);
		}

		/// <summary>
		/// Converts a string into appropriate SQL, including the substitution of any wildcard character into the correct SQL syntax.
		/// </summary>
		/// <param name="cmdText">The string to convert.</param>
		/// <param name="foundWildCard">On return, will be set to true if a wildcard was found.</param>
		/// <param name="isInputToADOCommandObject">true if the string should not be formatted because it will be bound to a parameter object.</param>
		/// <returns>The converted string.</returns>
		/// <exception cref="ErrAssistException">OMIGAERROR.InvalidSearchString if wild card found at an invalid position.</exception>
		/// <remarks>
		/// For example, "Smit*" will be converted to "'Smit%'". 
		/// </remarks>
		public static string FormatWildcardedString(string cmdText, out bool foundWildCard, bool isInputToADOCommandObject) 
		{
			string temp = String.Empty;

			foundWildCard = false;
			// check if input string contains wildcard character;
			// this test will also cope with zero length input strings
			int index = cmdText.IndexOf(_wildCard, StringComparison.OrdinalIgnoreCase);
			if (index == -1)
			{
				// no wildcard
				if (!isInputToADOCommandObject)
				{
					temp = FormatString(cmdText, false);
					// HasValue for Aposterphes
				}
				else
				{
					temp = FormatString(cmdText, true);
					// Do not check for Apostrphes
				}
			}
			else
			{
				int textLength = cmdText.Length;
				if (index == textLength - 1)
				{
					// Wildcard is valid, i.e. at the end of the input string.
					temp = cmdText.Substring(0, textLength - 1) + "%";
					if (!isInputToADOCommandObject)
					{
						temp = FormatString(temp, false);
					}
					else
					{
						temp = FormatString(temp, true);
					}
					foundWildCard = true;
				}
				else
				{
					throw new InvalidSearchStringException("Wildcard found at a position other than the end of a string: " + cmdText);
				}
			}

			return temp;
		}

		/// <summary>
		/// Gets the current system date in the correct SQL format.
		/// </summary>
		/// <returns>The current system date.</returns>
		public static string GetSystemDate() 
		{
			string dateText = String.Empty;
			DBPROVIDER dbProvider = AdoAssist.GetDbProvider();
			switch (dbProvider)
			{
				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					// DLM SQLServer Port
					dateText = "GetDate()";
					break;
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					dateText = "SYSDATE";
					break;
				default:
					throw new NotImplementedException();
			}
			return dateText;
		}

		/// <summary>
		/// Gets the SQL like operator.
		/// </summary>
		/// <returns>The SQL like operator.</returns>
		public static string GetLikeOperator() 
		{
			return "LIKE";
		}

		/// <summary>
		/// Gets the SQL string delimiter.
		/// </summary>
		/// <returns>The SQL string delimiter.</returns>
		private static string GetStringDelimiter() 
		{
			return "'";
		}

		/// <summary>
		/// Doubles any apostrophes in a SQL string.
		/// </summary>
		/// <param name="cmdText">The SQL string.</param>
		/// <returns>The converted SQL string.</returns>
		/// <remarks>
		/// The double of apostrophes is particular important for names that contain apostrophes, e.g., 
		/// "'Connor" will be converted to "O''Conner".
		/// </remarks>
		private static string CheckApostrophes(string cmdText) 
		{
			string sqlStringNew = String.Empty;
			int sqlStringLength = cmdText.Length;
			for (int index = 0; index < sqlStringLength; index++)
			{
				string thisChar = cmdText.Substring(index, 1);
				sqlStringNew = sqlStringNew + thisChar;
				if (thisChar == _apostrophe)
				{
					// add an extra apostrophe
					sqlStringNew = sqlStringNew + _apostrophe;
				}
			}
			return sqlStringNew;
		}

		/// <summary>
		/// Converts a date column name into the SQL format for a date with the time portion set to 0's.
		/// </summary>
		/// <param name="dateColumnName">The date column name.</param>
		/// <returns>The formated date.</returns>
		public static string TruncateDateColumn(string dateColumnName) 
		{
			string truncatedDateText = String.Empty;

			if (dateColumnName.Length > 0)
			{
				dateColumnName = dateColumnName.ToUpper();
				DBPROVIDER dbProvider = AdoAssist.GetDbProvider();

				switch (dbProvider)
				{
					case DBPROVIDER.omiga4DBPROVIDERSQLServer:
						truncatedDateText = "CONVERT(CHAR, " + dateColumnName + " , 6 )";
						break;
					case DBPROVIDER.omiga4DBPROVIDEROracle:
					case DBPROVIDER.omiga4DBPROVIDERMSOracle:
						truncatedDateText = "trunc(" + dateColumnName + ")";
						break;
					default:
						throw new NotImplementedException();
				}
			}
			else
			{
				throw new MissingFieldDescException();
			}

			return truncatedDateText;
		}

		/// <summary>
		/// Formats a date object to the appropriate SQL for the database engine, excluding the time portion.
		/// </summary>
		/// <param name="dateTime">The date object.</param>
		/// <returns>The formatted SQL.</returns>
		public static string FormatDate(DateTime dateTime)
		{
			return FormatDate(dateTime, DATETIMEFORMAT.dtfDate);
		}

		/// <summary>
		/// Formats a date object to the appropriate SQL for the database engine.
		/// </summary>
		/// <param name="dateTime">The date object.</param>
		/// <param name="dateTimeFormat">The date time format for the SQL.</param>
		/// <returns>The formatted SQL.</returns>
		public static string FormatDate(DateTime dateTime, DATETIMEFORMAT dateTimeFormat) 
		{		
			string dateText = String.Empty;

			DBPROVIDER dbProvider = AdoAssist.GetDbProvider();

			switch (dateTimeFormat)
			{
				case DATETIMEFORMAT.dtfDateTime:
					// date and time format
					dateText = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
					switch (dbProvider)
					{
						case DBPROVIDER.omiga4DBPROVIDERSQLServer:
							dateText = "CONVERT(DATETIME,'" + dateText + "',102)";
							break;
						case DBPROVIDER.omiga4DBPROVIDEROracle:
						case DBPROVIDER.omiga4DBPROVIDERMSOracle:
							dateText = "TO_DATE('" + dateText + "','YYYY-MM-DD HH24:MI:SS')";
							break;
						default:
							throw new NotImplementedException();
					}
					break;
				case DATETIMEFORMAT.dtfDate:
					// date only but time defaulted to 00:00:00
					dateText = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
					switch (dbProvider)
					{
						case DBPROVIDER.omiga4DBPROVIDERSQLServer:
							// DLM SQLServer Port
							dateText = "CONVERT(DATETIME,'" + dateText + "',120)";
							break;
						case DBPROVIDER.omiga4DBPROVIDEROracle:
						case DBPROVIDER.omiga4DBPROVIDERMSOracle:
							dateText = "TO_DATE('" + dateText + "','YYYY-MM-DD HH24:MI:SS')";
							break;
					}
					break;
				case DATETIMEFORMAT.dtfTime:
					throw new NotImplementedException("Formatting as Time only has not been implemented");
				default:
					throw new NotImplementedException();
			}

			return dateText;
		}

		/// <summary>
		/// Formats a date string to the appropriate SQL for the database engine.
		/// </summary>
		/// <param name="dateText">The date string.</param>
		/// <returns>The formatted SQL.</returns>
		/// <remarks>
		/// The <paramref name="dateText"/> string should be in the format "DD/MM/YYYY HH:MM:SS".
		/// </remarks>
		public static string FormatDateString(string dateText)
		{
			return FormatDateString(dateText, DATETIMEFORMAT.dtfDateTime);
		}

		/// <summary>
		/// Formats a date string to the appropriate SQL for the database engine.
		/// </summary>
		/// <param name="dateText">The date string.</param>
		/// <param name="dateTimeFormat">The date time format for the SQL.</param>
		/// <returns>The formatted SQL.</returns>
		/// <remarks>
		/// The <paramref name="dateText"/> string should be in the format "DD/MM/YYYY" 
		/// (if <paramref name="dateTimeFormat"/> == DATETIMEFORMAT.dtfDateTime) or "DD/MM/YYYY HH:MM:SS" 
		/// (if <paramref name="dateTimeFormat"/> == DATETIMEFORMAT.dtfDate).
		/// </remarks>
		public static string FormatDateString(string dateText, DATETIMEFORMAT dateTimeFormat) 
		{
			DateTime dateTime = Convert.ToDateTime(dateText);

			switch (dateTimeFormat)
			{
				case DATETIMEFORMAT.dtfDateTime:
					if (dateText.Length < 19)
					{
						throw new InvalidDateTimeFormatException("Enumerated dateTimeFormat value of " + dateTimeFormat + "Requested for DateTime");
					}
					break;
				case DATETIMEFORMAT.dtfDate:
					if (dateText.Length != 10)
					{
						throw new InvalidDateTimeFormatException("Enumerated dateTimeFormat value of " + dateTimeFormat + " Requested for Date");
					}
					else
					{
						dateTime = Convert.ToDateTime(dateText + " 00:00:00");
						// default DATES to date + 00:00:00
					}
					break;
				case DATETIMEFORMAT.dtfTime:
					throw new NotImplementedException("Formatting as Time only has not been implemented");
				default:
					throw new NotImplementedException();
			}

			return FormatDate(dateTime, dateTimeFormat);
		}

		/// <summary>
		/// Converts a SqlGuid object to a guid string in binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
		/// </summary>
		/// <param name="sqlGuid">The SqlGuid object.</param>
		/// <returns>The guid string.</returns>
		public static string GuidToString(SqlGuid sqlGuid)
		{
			return GuidToString(sqlGuid, GUIDSTYLE.guidBinary);
		}

		/// <summary>
		/// Converts a SqlGuid object to a guid string in a required guid format.
		/// </summary>
		/// <param name="sqlGuid">The SqlGuid object.</param>
		/// <param name="guidStyle">The required guid format.</param>
		/// <returns>The guid string.</returns>
		public static string GuidToString(SqlGuid sqlGuid, GUIDSTYLE guidStyle)
		{
			return GuidToString(sqlGuid.Value.ToByteArray(), SqlDbType.Binary, guidStyle);
		}

		/// <summary>
		/// Converts a SqlBinary object to a guid string in binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
		/// </summary>
		/// <param name="sqlBinary">The SqlBinary object.</param>
		/// <returns>The guid string.</returns>
		public static string GuidToString(SqlBinary sqlBinary)
		{
			return GuidToString(sqlBinary, GUIDSTYLE.guidBinary);
		}

		/// <summary>
		/// Converts a SqlBinary object to a guid string in a required guid format.
		/// </summary>
		/// <param name="sqlBinary">The SqlBinary object.</param>
		/// <param name="guidStyle">The required guid format.</param>
		/// <returns>The guid string.</returns>
		public static string GuidToString(SqlBinary sqlBinary, GUIDSTYLE guidStyle)
		{
			// SqlDataReader.GetSqlBinary will truncate at null bytes.
			// Thus the guid 0x7F88653074204A54A02AF8E16F879E00 will be returned as 
			// a 15 byte SqlBinary (and not 16 bytes). The following workaround creates 
			// a SqlBinary object of the right length (16 bytes), with terminating null bytes.
			if (sqlBinary.Length < 16)
			{
				byte[] buffer = new byte[16];
				Array.Copy(sqlBinary.Value, buffer, sqlBinary.Length);
				sqlBinary = new SqlBinary(buffer);
			}
			return GuidToString(sqlBinary.Value, SqlDbType.Binary, guidStyle);
		}

		/// <summary>
		/// Converts a byte array to a guid string in binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
		/// </summary>
		/// <param name="guidBytes">The byte array.</param>
		/// <returns>The guid string.</returns>
		public static string GuidToString(byte[] guidBytes)
		{
			return GuidToString(guidBytes, SqlDbType.Binary, GUIDSTYLE.guidBinary);
		}

		/// <summary>
		/// Converts a byte array to a guid string in binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
		/// </summary>
		/// <param name="guidBytes">The byte array.</param>
		/// <param name="sqlDbType">The SqlDbType for the underlying database column for the byte array.</param>
		/// <returns>The guid string.</returns>
		/// <remarks>
		/// <paramref name="sqlDbType"/> should be one of SqlDbType.UniqueIdentifier, SqlDbType.Binary or SqlDbType.VarBinary.
		/// </remarks>
		public static string GuidToString(byte[] guidBytes, SqlDbType sqlDbType)
		{
			return GuidToString(guidBytes, sqlDbType, GUIDSTYLE.guidBinary);
		}

		/// <summary>
		/// Converts a byte array to a guid string in a required guid format.
		/// </summary>
		/// <param name="guidBytes">The byte array.</param>
		/// <param name="sqlDbType">The SqlDbType for the underlying database column for the byte array.</param>
		/// <param name="guidStyle">The required guid format.</param>
		/// <returns>The guid string.</returns>
		/// <remarks>
		/// <paramref name="sqlDbType"/> should be one of SqlDbType.UniqueIdentifier, SqlDbType.Binary or SqlDbType.VarBinary.
		/// </remarks>
		public static string GuidToString(byte[] guidBytes, SqlDbType sqlDbType, GUIDSTYLE guidStyle) 
		{
			// header ----------------------------------------------------------------------------------
			// description:  Converts a guid returned (as binary) from the database into a guid string.
			// pass:
			// guidBytes   The byte array to be converted. This can be an adBinary, adVarBinary, or
			// adGUID read from the database.
			// 
			// sqlDbType	   The ADO data type of the byte array. Unless you explicitly set the
			// data type, e.g., in ADODB.CommandCreateParameter, it defaults to:
			// 
			// Database Type			   ADO Type
			// ------------------------------------
			// Oracle raw				  adVarBinary
			// SQL Server binary		   adBinary
			// SQL Server varbinary		adVarBinary
			// SQL Server uniqueidentifier adGUID
			// 
			// The datatype on the database can be coerced to a different ADO data type
			// by explicitly setting the type, e.g., in ADODB.CommandCreateParameter.
			// 
			// If you explicitly set the data type to adGUID for SQL Server, you MUST
			// set sqlDbType to adGUID when calling GuidToString, otherwise you can
			// allow sqlDbType to default to adVarBinary.
			// 
			// The default of adVarBinary matches the data type returned for binary fields
			// in the Omiga 4 Phase 2 Oracle database.
			// 
			// guidStyle	  The required style for the resulting guid string, e.g.,
			// guidBinary = "DA6DA163412311D4B5FA00105ABB1680".
			// guidHyphen = "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
			// The default of guidBinary gives the same format as Omiga 4 Phase 2.
			// 
			// return:		   The guid string.
			// 
			// AS				28/02/01	First version
			// ------------------------------------------------------------------------------------------

			string guidText = String.Empty;

			if (sqlDbType == SqlDbType.UniqueIdentifier && AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
			{
				// The data type has explicitly been set to adGUID.
				// SQL Server adGUID data type is already in hyphenated format.
				guidText = ByteArrayToGuidString(GuidAssist.ToByteArray(new Guid(guidBytes)), guidStyle);
			}
			else if (sqlDbType == SqlDbType.Binary || sqlDbType == SqlDbType.VarBinary)
			{
				// For Oracle the data type defaults to adVarBinary.
				// Guids are always read from Oracle as binary arrays, irrespective of whether the
				// ADO data type is adGUID or adVarBinary.
				// For SQL Server the data type defaults to adBinary.
				// Convert binary array to guid string.
				guidText = ByteArrayToGuidString(guidBytes, guidStyle);
			}
			else
			{
				throw new InvalidParameterException("sqlDbType must be adGUID or adBinary or adVarBinary");
			}

			return guidText;
		}

		/// <summary>
		/// Converts a DateTime object into a date string in "dd/MM/yyyy" format.
		/// </summary>
		/// <param name="dateTime">The DateTime object.</param>
		/// <returns>The date string.</returns>
		public static string DateToString(DateTime dateTime) 
		{
			return dateTime.ToString("dd/MM/yyyy");
		}

		/// <summary>
		/// Converts a DateTime object into a date time string in "dd/MM/yyyy HH:mm:ss" format.
		/// </summary>
		/// <param name="dateTime">The DateTime object.</param>
		/// <returns>The date time string.</returns>
		public static string DateTimeToString(DateTime dateTime) 
		{
			return dateTime.ToString("dd/MM/yyyy HH:mm:ss");
		}

		/// <summary>
		/// Converts an expression into the appropriate SQL string for the expression's database type.
		/// </summary>
		/// <param name="expression">The expression string.</param>
		/// <param name="dbDataType">The expression's database type.</param>
		/// <returns></returns>
		public static string GetFormattedValue(string expression, DBDATATYPE dbDataType) 
		{
			string formattedValue = String.Empty;

			if (expression.Length == 0)
			{
				// If value is blank/null then always return the NULL value
				return "NULL";
			}
			switch (dbDataType)
			{
				case DBDATATYPE.dbdtGuid:
					formattedValue = FormatGuid(expression, GUIDSTYLE.guidLiteral);
					break;
				case DBDATATYPE.dbdtString:
					formattedValue = FormatString(expression, false);
					break;
				case DBDATATYPE.dbdtInt:
				case DBDATATYPE.dbdtComboId:
				case DBDATATYPE.dbdtCurrency:
				case DBDATATYPE.dbdtBoolean:
				case DBDATATYPE.dbdtDouble:
				case DBDATATYPE.dbdtLong:
					// value is unchanged
					formattedValue = expression;
					break;
				case DBDATATYPE.dbdtDate:
					formattedValue = FormatDateString(expression, DATETIMEFORMAT.dtfDate);
					break;
				case DBDATATYPE.dbdtDateTime:
					formattedValue = FormatDateString(expression, DATETIMEFORMAT.dtfDateTime);
					break;
				default:
					throw new InValidDataTypeValueException("Data type " + dbDataType + " is not valid");
			}

			return formattedValue;
		}

		/// <summary>
		/// Gets the required SQL to allow a SQL query to search on the day portion of a date for a specified column. 
		/// </summary>
		/// <param name="columnName">The column name.</param>
		/// <returns>The required SQL.</returns>
		/// <remarks>
		/// For example, for Oracle this returns "TO_CHAR(<paramref name="columnName"/>, 'D')". 
		/// For SQL Server this returns "DATEPART(WEEKDAY, <paramref name="columnName"/>)".
		/// </remarks>
		public static string GetColumnDayOfWeek(string columnName) 
		{
			string cmdText = String.Empty;
			
			if (columnName.Length == 0)
			{
				throw new InvalidParameterException("Column name must be entered");
			}
			DBPROVIDER dbProvider = AdoAssist.GetDbProvider();
			switch (dbProvider)
			{
				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					cmdText = "DATEPART(WEEKDAY, " + columnName + ")";
					break;
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					cmdText = "TO_CHAR(" + columnName + ",'D')";
					break;
				default:
					// not implemented
					throw new NotImplementedException();
			}
			return cmdText;
		}

		/// <summary>
		/// Converts a guid string into a byte array.
		/// </summary>
		/// <param name="guidText">The guid string.</param>
		/// <returns>The guid string converted into a byte array.</returns>
		/// <remarks>
		/// The <paramref name="guidText"/> can be in either binary format, 
		/// e.g., "DA6DA163412311D4B5FA00105ABB1680", or hyphenated format, e.g., 
		/// "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
		/// </remarks>
		public static byte[] GuidStringToByteArray(string guidText) 
		{
			byte[] guidBytes = new byte[16];
			if (guidText.Length == 38)
			{
				// Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
				guidText = FormatGuid(guidText, GUIDSTYLE.guidBinary);
			}

			if (guidText.Length == 32)
			{
				// Convert from "DA6DA163412311D4B5FA00105ABB1680" to byte array.
				for (int index = 0; index < guidBytes.Length; index++)
				{
					guidBytes[index] = Convert.ToByte(guidText.Substring(index * 2, 2), 16);
				}
			}
			else
			{
				System.Diagnostics.Debug.Assert(false);
				throw new InvalidParameterException("guidText length must be 32");
			}
			return guidBytes;
		}

		/// <summary>
		/// Converts a byte array into a guid string in binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
		/// </summary>
		/// <param name="guidBytes">The guid as a byte array.</param>
		/// <returns>The guid string.</returns>
		public static string ByteArrayToGuidString(byte[] guidBytes)
		{
			return ByteArrayToGuidString(guidBytes, GUIDSTYLE.guidBinary);
		}

		/// <summary>
		/// Converts a byte array into a guid string in a required format.
		/// </summary>
		/// <param name="guidBytes">The guid as a byte array.</param>
		/// <param name="guidStyle">The required guid format.</param>
		/// <returns>The guid string.</returns>
		/// <remarks>
		/// If <paramref name="guidStyle"/> is GUIDSTYLE.guidHyphen, then the guid string will be in 
		/// "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format. 
		/// If <paramref name="guidStyle"/> is GUIDSTYLE.guidBinary, then the guid string will be in 
		/// "DA6DA163412311D4B5FA00105ABB1680" format.
		/// </remarks>
		public static string ByteArrayToGuidString(byte[] guidBytes, GUIDSTYLE guidStyle) 
		{
			string guidText = String.Empty;

			System.Diagnostics.Debug.Assert(guidBytes.Length == 16);
			if (guidBytes.Length == 16)
			{
				if (guidStyle == GUIDSTYLE.guidHyphen)
				{
					// "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
					guidText = "{";
					for (int index1 = 0; index1 < 16; index1++)
					{
						int index2 = ConvertGuidIndex(index1);
						guidText = guidText + guidBytes[index2].ToString("X2");
						if (index1 == 3 || index1 == 5 || index1 == 7 || index1 == 9)
						{
							guidText = guidText + "-";
						}
					}
					guidText = guidText + "}";
				}
				else if (guidStyle == GUIDSTYLE.guidBinary)
				{
					// "DA6DA163412311D4B5FA00105ABB1680" format.
					for (int index = 0; index < 16; index++)
					{
						guidText = guidText + guidBytes[index].ToString("X2");
					}
				}
				else
				{
					System.Diagnostics.Debug.Assert(false);
					throw new InvalidParameterException("GUIDSTYLE must be guidHypen or guidBinary");
				}
			}
			else
			{
				System.Diagnostics.Debug.Assert(false);
				throw new InvalidParameterException("guidBytes size must be 16");
			}

			return guidText;
		}

		private static int ConvertGuidIndex(int index1) 
		{
			int index2 = index1;
			switch (index1)
			{
				case 0:
					index2 = 3;
					break;
				case 1:
					index2 = 2;
					break;
				case 2:
					index2 = 1;
					break;
				case 3:
					index2 = 0;
					break;
				case 4:
					index2 = 5;
					break;
				case 5:
					index2 = 4;
					break;
				case 6:
					index2 = 7;
					break;
				case 7:
					index2 = 6;
					break;
				default:
					index2 = index1;
					break; 
			}
			return index2;
		}
	}
}

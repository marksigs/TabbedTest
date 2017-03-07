/*
--------------------------------------------------------------------------------------------
Workfile:			ADOAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		.Net version of ADOAssist.cls. Wraps AdoAssist.cs.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		29/06/99    Created
RF		28/9/99     Added GetNumberOfRecords
SR		18/11/99    Added methods - CheckRecordExists, GetValueFromTable
SR		01/12/99    Modified method - GetValueFromTable
PSC		20/12/99    Do not use default interface
PSC		01/02/00    Add adExecuteNoRecords to ExecuteSQLCommand
RF		25/02/00    Enhanced error messaging to help fix AQR SYS0321.
MC		07/08/00    SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
PSC		11/08/00    SYS1430 Back ouut SYS1409
MC		31/08/00    Updated GetConnectionStringFromSPM to correct load testing error.
APS		07/09/00    Corrected the check for oeChildRecordsFound in ExecuteSQLCommand
DJP		02/10/00    P514 - get connection string from location if passed in.
LD		07/11/00    Explicity close database connections
LD		07/11/00    Explicity destroy command objects
APS		27/02/01    SYS1986
DM		22/05/01    SYS2321 Common handling of error numbers
LD		12/06/01    SYS2368 Modify code to get the connection string for Oracle and SQLServer
AS		13/06/01    CC012 Added GetValidRecordset
LD		29/06/01    Modify GetProviderSpecificConnStr
DC		27/09/01    SYS2748 Modify ExecuteSQLCommand to handle SQL Server error
DC		02/10/01    SYS2695 GetValuefromTable - check for adbinary added
PSC		17/10/01    SYS2815 Enable SQL Server Integrated Security
AS		13/11/03    CORE1 Removed GENERIC_SQL.
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from ADOAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete]
	public enum DBENGINETYPE
	{
		Undefined,
		SQLServer,
		Oracle,
	}

	[Obsolete("Use AdoAssist instead")]
	public static class ADOAssist
    {
		[Obsolete("Use AdoAssist.GetDbConnectString instead")]
        public static string GetConnStr()
        {
			return AdoAssist.GetDbConnectString();
        }

		[Obsolete("Use AdoAssist.IsUniqueKeyConstraint instead")]
		public static bool IsUniqueKeyConstraint(Exception exception) 
        {
			return AdoAssist.IsUniqueKeyConstraint(exception);
        }

		[Obsolete("Use AdoAssist.IsUniqueKeyConstraint instead")]
		public static bool IsThisUniqueKeyConstraint(Exception exception) 
        {
			return AdoAssist.IsUniqueKeyConstraint(exception);
        }

		[Obsolete("Use AdoAssist.IsDBConnectionError instead")]
		public static bool IsDBConnectionError(Exception exception)
		{
			return AdoAssist.IsDBConnectionError(exception);
		}

		[Obsolete("Use AdoAssist.GetDbProvider instead")]
		public static DBENGINETYPE GetDBEngine() 
        {
			DBENGINETYPE dbEngineType = DBENGINETYPE.Undefined;
			DBPROVIDER dbProvider = AdoAssist.GetDbProvider();

			switch (dbProvider)
			{
				case DBPROVIDER.omiga4DBPROVIDERSQLServer:
					dbEngineType = DBENGINETYPE.SQLServer;
					break;
				case DBPROVIDER.omiga4DBPROVIDEROracle:
				case DBPROVIDER.omiga4DBPROVIDERMSOracle:
					dbEngineType = DBENGINETYPE.Oracle;
					break;
            }

			return dbEngineType;
        }

		[Obsolete("Use AdoAssist.GetDbRetries instead")]
		public static int GetRetries()
		{
			return AdoAssist.GetDbRetries();
		}

		[Obsolete("Use AdoAssist.ExecuteSQLCommand instead")]
		public static int ExecuteSQLCommand(string cmdText) 
		{
			return AdoAssist.ExecuteSQLCommand(cmdText, true);
		}

		[Obsolete("Use AdoAssist.ExecuteSQLCommand instead")]
		public static int ExecuteSQLCommand(string cmdText, bool retry) 
        {
			return AdoAssist.ExecuteSQLCommand(cmdText, retry);
        }

		[Obsolete("Use AdoAssist.LockTable instead")]
		public static void LockTable(string tableName, LockType lockType) 
        {
			AdoAssist.LockTable(tableName, lockType);
        }

		[Obsolete("Use AdoAssist.CheckSingleRecordExists instead")]
		public static void CheckSingleRecordExists(string tableName, string keys) 
        {
			AdoAssist.CheckSingleRecordExists(tableName, keys);
        }

		[Obsolete("Not used")]
		public static DataSet GetFieldTypes(string tableName) 
        {
			return null;
        }

		[Obsolete("Use AdoAssist.GetNumberOfRecords instead")]
		public static int GetNumberOfRecords(string tableName, string criteria) 
        {
			return AdoAssist.GetNumberOfRecords(tableName, criteria);
        }

		[Obsolete("Use AdoAssist.CheckRecordExists instead")]
		public static bool CheckRecordExists(string tableName, string criteria) 
        {
			return AdoAssist.CheckRecordExists(tableName, criteria) ;
		}

		[Obsolete("Use AdoAssist.StartSQLLogging instead")]
		public static void StartSQLLogging() 
        {
			AdoAssist.StartSQLLogging();
        }

		[Obsolete("Use AdoAssist.StopSQLLogging instead")]
		public static void StopSQLLogging() 
        {
			AdoAssist.StopSQLLogging();
		}

		[Obsolete("Use AdoAssist.WriteSQLToFile instead")]
		public static void WriteSQLToFile(string cmdText, string functionName) 
        {
			AdoAssist.WriteSQLToFile(cmdText, functionName);
		}

		[Obsolete("Use AdoAssist.GetValueFromTable instead")]
		public static object GetValueFromTable(string tableName, string criteria, string columnName) 
		{
			object returnValue = null;
			bool recordExists = false;
			return AdoAssist.GetValueFromTable(tableName, criteria, columnName, out returnValue, out recordExists);
		}

		[Obsolete("Use AdoAssist.GetValueFromTable instead")]
		public static object GetValueFromTable(string tableName, string criteria, string columnName, out object returnValue, out bool recordExists) 
        {
			return AdoAssist.GetValueFromTable(tableName, criteria, columnName, out returnValue, out recordExists);
        }

		[Obsolete("Use AdoAssist.GetValidRecordset instead")]
		public static bool GetValidRecordset(DataSet dataSet, out DataTable dataTable)
		{
			return AdoAssist.GetValidRecordset(dataSet, 1, out dataTable);
		}

		[Obsolete("Use AdoAssist.GetValidRecordset instead")]
		public static bool GetValidRecordset(DataSet dataSet, int tableIndex, out DataTable dataTable)
        {
			return AdoAssist.GetValidRecordset(dataSet, tableIndex, out dataTable);
        }
	}
}

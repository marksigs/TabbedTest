/*
--------------------------------------------------------------------------------------------
Workfile:			IADOAssist.cs
Copyright:			Copyright © 2007 Vertex Financial Services
Description:		.Net version of IADOAssist.cls.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		19/06/2007	First .Net version. Ported from IADOAssist.cls.
					Implemented as wrappers around the equivalent ADOAssist methods after 
					merging any differences.
					Implemented as a class and not an interface because static classes can 
					not implement interfaces. 
					Marked as obsolete to discourage use - if possible use AdoAssist instead.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use AdoAssist instead")]
	public class IADOAssist
	{
		[Obsolete("Use AdoAssist.CheckRecordExists instead")]
		public bool CheckRecordExists(string tableName, string criteria)
		{
			return AdoAssist.CheckRecordExists(tableName, criteria);
		}

		[Obsolete("Use AdoAssist.ExecuteSQLCommand instead")]
		public int ExecuteSQLCommand(string cmdText, bool retry)
		{
			return AdoAssist.ExecuteSQLCommand(cmdText, retry);
		}

		[Obsolete("Use AdoAssist.GetDbConnectString instead")]
		public string GetConnStr()
		{
			return AdoAssist.GetDbConnectString(); ;
		}

		[Obsolete("Use AdoAssist.GetDbProvider instead")]
		public DBENGINETYPE GetDBEngine()
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

		[Obsolete("Not used")]
		private DataSet GetFieldTypes(string tableName)
		{
			return null;
		}

		[Obsolete("Use AdoAssist.GetNumberOfRecords instead")]
		public int GetNumberOfRecords(string tableName, string criteria)
		{
			return AdoAssist.GetNumberOfRecords(tableName, criteria);
		}

		[Obsolete("Use AdoAssist.GetValueFromTable instead")]
		public object GetValueFromTable(string tableName, string criteria, string columnName)
		{
			object returnValue = null;
			bool recordExists = false;
			return AdoAssist.GetValueFromTable(tableName, criteria, columnName, out returnValue, out recordExists);
		}

		[Obsolete("Use AdoAssist.GetValueFromTable instead")]
		public object GetValueFromTable(string tableName, string criteria, string columnName, out object returnValue, out bool recordExists)
		{
			return AdoAssist.GetValueFromTable(tableName, criteria, columnName, out returnValue, out recordExists);
		}

		[Obsolete("Use AdoAssist.IsUniqueKeyConstraint instead")]
		public bool IsUniqueKeyConstraint(Exception exception)
		{
			return AdoAssist.IsUniqueKeyConstraint(exception);
		}

		[Obsolete("Use AdoAssist.IsUniqueKeyConstraint instead")]
		public bool IsThisUniqueKeyConstraint(Exception exception)
		{
			return AdoAssist.IsUniqueKeyConstraint(exception);
		}

		[Obsolete("Use AdoAssist.LockTable instead")]
		public void LockTable(string tableName, LockType lockType)
		{
			AdoAssist.LockTable(tableName, lockType);
		}

		[Obsolete("Use AdoAssist.GetDbRetries instead")]
		public int GetRetries()
		{
			return AdoAssist.GetDbRetries();
		}

		[Obsolete("Use AdoAssist.IsDBConnectionError instead")]
		public bool IsDBConnectionError(Exception exception)
		{
			return AdoAssist.IsDBConnectionError(exception);
		}
	}
}

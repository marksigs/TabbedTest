/*
--------------------------------------------------------------------------------------------
Workfile:			DBHelper.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
SDS		22/01/2004	LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
MV		28/01/2004	LIVE00009659  Amended DBAssistBuildDbConnectionString()
TW		18/10/2005	MAR223
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		10/07/2007	First .Net version. Ported from DBHelper.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use the appropriate Vertex.Fsd.Omiga.VisualBasicPort class")]
	public static class DBHelper
	{
		[Obsolete("Use AdoAssist.ConvertRecordSetToXML instead")]
		public static void ConvertRecordSetToXML(XmlNode xmlSchemaNode, XmlNode xmlResponseNode, DataSet dataSet, bool comboLookUp)
		{
			AdoAssist.ConvertRecordSetToXML(xmlSchemaNode, xmlResponseNode, dataSet, comboLookUp);
		}

		[Obsolete("Use AdoAssist.executeGetRecordSet instead")]
		public static DataSet executeGetRecordSet(SqlCommand sqlCommand)
		{
			return AdoAssist.executeGetRecordSet(sqlCommand);
		}

		[Obsolete("Use SqlAssist.GuidStringToByteArray instead")]
		public static byte[] GuidStringToByteArray(string guidText)
		{
			return SqlAssist.GuidStringToByteArray(guidText);
		}

		[Obsolete("Use AdoAssist.GetDbConnectString instead")]
		public static string GetDbConnectString()
		{
			return AdoAssist.GetDbConnectString();
		}

		[Obsolete("Use GuidAssist.ToString() instead")]
		public static string GuidToString(byte[] bytes)
		{
			return GuidAssist.ToString(bytes);
		}

		[Obsolete("Use AdoAssist.GetDbConnectString instead")]
		public static void DBAssistBuildDbConnectionString()
		{
			AdoAssist.GetDbConnectString(); 
		}

		[Obsolete("Create an AdoAssist object instead")]
		public static void DBAssistLoadSchema()
		{
			;
		}

		[Obsolete("Use AdoAssist.GetRecordSetAsXML instead")]
		public static void GetRecordSetAsXML(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlResponseNode)
		{
			AdoAssist.GetRecordSetAsXML(xmlDataNode, xmlSchemaNode, xmlResponseNode);
		}

		[Obsolete("Use AdoAssist.GetRecordSetAsXML instead")]
		public static void GetRecordSetAsXML(XmlNode xmlDataNode, XmlNode xmlSchemaNode, XmlNode xmlResponseNode, string filter, string orderBy)
		{
			AdoAssist.GetRecordSetAsXML(xmlDataNode, xmlSchemaNode, xmlResponseNode, filter, orderBy);
		}

		[Obsolete("Use AdoAssist.GetRecordSet instead")]
		public static DataSet GetRecordSet(XmlNode xmlDataNode, XmlNode xmlSchemaNode)
		{
			return AdoAssist.GetRecordSet(xmlDataNode, xmlSchemaNode);
		}

		[Obsolete("Use AdoAssist.GetRecordSet instead")]
		public static DataSet GetRecordSet(XmlNode xmlDataNode, XmlNode xmlSchemaNode, string filter, string orderBy)
		{
			return AdoAssist.GetRecordSet(xmlDataNode, xmlSchemaNode, filter, orderBy);
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			IDOAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		.Net version of DOAssist.cls.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		25/01/00    Added GetNextSequenceNumberEx.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		18/06/2007	First .Net version. Ported from IDOAssist.cls.
					Implemented as wrappers around the equivalent DOAssist methods after 
					merging any differences (the original VB6 code did not use wrappers but 
					copied the underlying code with insignificant differences).
					Implemented as a class and not an interface because static classes can 
					not implement interfaces. 
					Marked as obsolete to discourage use - if possible use DOAssist instead.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use DOAssist instead")]
    public class IDOAssist
    {
		[Obsolete("Use DOAssist.BuildSQLStringEx instead")]
		public void BuildSQLString(XmlElement xmlRequest, XmlDocument xmlClassDefinition, ref string tableName)
		{
			DOAssist.BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName);
		}

		[Obsolete("Use DOAssist.BuildSQLStringEx instead")]
		public void BuildSQLString(XmlElement xmlRequest, XmlDocument xmlClassDefinition, ref string tableName, SQL_FORMAT_TYPE sqlFormatType, CLASS_DEF_KEY classDefKey, CLASS_DEF_KEY_AMOUNT classDefKeyAmount, ref string fieldValuePair, ref string fields, ref string values, ref string select, string itemName)
		{
			DOAssist.BuildSQLStringEx(xmlRequest, xmlClassDefinition, ref tableName, sqlFormatType, classDefKey, classDefKeyAmount, ref fieldValuePair, ref fields, ref values, ref select, itemName);
		}

		[Obsolete("Use DOAssist.CreateEx instead")]
		public void Create(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			DOAssist.CreateEx(xmlRequest, xmlClassDefinition);
		}

		[Obsolete("Use DOAssist.UpdateEx instead")]
		public void Update(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			DOAssist.UpdateEx(xmlRequest, xmlClassDefinition);
		}

		[Obsolete("Use DOAssist.DeleteEx instead")]
		public void Delete(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			DOAssist.DeleteEx(xmlRequest, xmlClassDefinition);
		}

		[Obsolete("Use DOAssist.DeleteAllEx instead")]
		public void DeleteAll(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			DOAssist.DeleteAllEx(xmlRequest, xmlClassDefinition);
		}

		[Obsolete("Use DOAssist.FindListEx instead")]
		public XmlNode FindList(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return DOAssist.FindListEx(xmlRequest, xmlClassDefinition, null);
		}

		[Obsolete("Use DOAssist.FindListEx instead")]
		public XmlNode FindList(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string orderByField)
		{
			return DOAssist.FindListEx(xmlRequest, xmlClassDefinition, orderByField);
		}

		[Obsolete("Use DOAssist.FindListMultipleEx instead")]
		public XmlNode FindListMultiple(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return DOAssist.FindListMultipleEx(xmlRequest, xmlClassDefinition, null, null, null, null);
		}

		[Obsolete("Use DOAssist.FindListMultipleEx instead")]
		public XmlNode FindListMultiple(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string orderByField, string itemName)
		{
			return DOAssist.FindListMultipleEx(xmlRequest, xmlClassDefinition, orderByField, itemName, null, null);
		}

		[Obsolete("Use DOAssist.FindListMultipleEx instead")]
		public XmlNode FindListMultiple(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string orderByField, string itemName, string additionalCondition, string andOrToCondition)
		{
			return DOAssist.FindListMultipleEx(xmlRequest, xmlClassDefinition, orderByField, itemName, additionalCondition, andOrToCondition);
		}

		[Obsolete("Use DOAssist.GetNextSequenceNumberEx instead")]
		public int GetNextSequenceNumber(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string tableName, string sequenceFieldName)
		{
			return DOAssist.GetNextSequenceNumberEx(xmlRequest, xmlClassDefinition, tableName, sequenceFieldName);
		}

		[Obsolete("Use DOAssist.GenerateSequenceNumberEx instead")]
		public int GenerateSequenceNumber(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string sequenceFieldName)
		{
			return DOAssist.GenerateSequenceNumberEx(xmlRequest, xmlClassDefinition, sequenceFieldName);
		}

		[Obsolete("Use DOAssist.GetXMLForTableSchema instead")]
		public XmlNode GetXMLFromRecordSet(DataRow dataRow, XmlDocument xmlClassDefinition, XmlNode xmlNode)
		{
			return DOAssist.GetXMLForTableSchema(dataRow, xmlClassDefinition.DocumentElement, xmlNode);
		}

		[Obsolete("Use DOAssist.GetXMLFromWholeRecordset instead")]
		public void GetXMLFromWholeRecordset(DataTable dataTable, XmlDocument xmlClassDefinition, XmlNode xmlNode)
		{
			DOAssist.GetXMLFromWholeRecordset(dataTable, xmlClassDefinition, xmlNode);
		}

		[Obsolete("Use DOAssist.GetDataEx instead")]
		public XmlNode GetData(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return DOAssist.GetDataEx(xmlRequest, xmlClassDefinition, null, null);
		}

		[Obsolete("Use DOAssist.GetDataEx instead")]
		public XmlNode GetData(XmlElement xmlRequest, XmlDocument xmlClassDefinition, string tableName, string itemName)
		{
			return DOAssist.GetDataEx(xmlRequest, xmlClassDefinition, tableName, itemName);
		}

		[Obsolete("Use DOAssist.GetComponentData instead")]
		public XmlNode GetComponentData(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return DOAssist.GetComponentData(xmlRequest, xmlClassDefinition);
		}
    }
}

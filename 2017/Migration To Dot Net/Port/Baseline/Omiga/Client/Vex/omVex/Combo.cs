/*
--------------------------------------------------------------------------------------------
Workfile:			Combo.cs
Copyright:			Copyright ©2006 Vertex Financial Services

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PE		27/11/2006	EP2_24 - Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Reflection;

using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omVex
{
	/// <summary>
	/// Summary description for Combo.
	/// </summary>
	public class OmigaCombo
	{
		
		public static string GetComboValue(string GroupName, string ValidationType)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

			string ComboValue = "";
			string ComboID = "";

			try
			{
				OmigaDataClass OmigaData = new OmigaDataClass();
				OmigaData.GetComboValue(GroupName, ValidationType, out ComboValue, out ComboID);
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			return(ComboValue);
		}

		public static string GetComboID(string GroupName, string ValidationType)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

			string ComboValue = "";
			string ComboID = "";

			try
			{
				OmigaDataClass OmigaData = new OmigaDataClass();
				OmigaData.GetComboValue(GroupName, ValidationType, out ComboValue, out ComboID);
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			return(ComboID);
		}

		public static string GetComboIDByName(string GroupName, string ValueName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

			string ValueID = "";

			try
			{
				OmigaDataClass OmigaData = new OmigaDataClass();
				OmigaData.GetComboIDByValue(GroupName, ValueName, out ValueID);
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			return(ValueID);
		}

	}
}

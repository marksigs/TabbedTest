/*
--------------------------------------------------------------------------------------------
Workfile:			ComboAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Combo values.
--------------------------------------------------------------------------------------------
BMIDS Specific History
Prog	Date		Description
GD	  10/07/2002	BMIDS00165 - New Function IsValidationTypeInValidationList
BS	  13/05/2003	BM0310 Amended test for nothing found in IsValidationTypeInValidationList
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/2004	BBG1821 Performance related fixes
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from comboAssist.bas and rewritten to remove
					unnecessary use of xml and adoAssist - now executes 10 times faster.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Supports reading combo values from the database.
	/// </summary>
	public static class ComboAssist
	{
		private static Dictionary<string, ComboCollection> _comboGroups = new Dictionary<string, ComboCollection>();

		/// <summary>
		/// Gets the combo text for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueId">The value id.</param>
		/// <returns>The matching combo text.</returns>
		/// <remarks>
		/// Equivalent to "select VALUENAME from COMBOVALUE where GROUPNAME=<i>groupName</i> and VALUEID=<i>valueId</i>".
		/// </remarks>
		public static string GetComboText(string groupName, int valueId)
		{
			return GetComboGroup(groupName).GetValueNameByValueId(valueId);
		}

		/// <summary>
		/// Returns true if a specified validation type exists for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueId">The value id.</param>
		/// <param name="validationType">The validation type.</param>
		/// <returns>true if validation type exists.</returns>
		/// <remarks>
		/// Equivalent to "select 1 where exists (select * from COMBOVALIDATION where GROUPNAME=<i>groupName</i> and VALUEID=<i>valueId</i> and VALIDATIONTYPE=<i>validationType</i>)"
		/// </remarks>
		public static bool IsValidationType(string groupName, int valueId, string validationType)
		{
			return GetComboGroup(groupName).GetComboValueByValueIdAndValidationType(valueId, validationType) != null;
		}

		// Missing XML comment for publicly visible type or member.
		#pragma warning disable 1591
		[Obsolete("Use IsValidationType instead as it is functionally equivalent (swap the order of the validationType and valueId parameters)")]
		public static bool IsValidationTypeInValidationList(string groupName, string validationType, int valueId)
		{
			// This method in the VB6 code seems to be functionally equivalent to IsValidationType, 
			// with the parameters in a different order. IsValidationType is called more often
			// so that is the primary implementation.
			return IsValidationType(groupName, valueId, validationType);
		}
		// Missing XML comment for publicly visible type or member.
		#pragma warning restore 1591

		/// <summary>
		/// Gets a list of value ids for a specified combo group and validation type.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="validationType">The validation type.</param>
		/// <returns>The matching value ids.</returns>
		/// <remarks>
		/// Equivalent to "select VALUEID from COMBOVALIDATION where GROUPNAME=<i>groupName</i> and VALIDATIONTYPE=<i>validationType</i>".
		/// </remarks>
		public static ReadOnlyCollection<int> GetValueIdsForValidationType(string groupName, string validationType)
		{
			return GetComboGroup(groupName).GetValueIdsByValidationType(validationType);
		}

		/// <summary>
		/// Gets the first validation type for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueId">The value id.</param>
		/// <returns>The first matching validation type.</returns>
		/// <remarks>
		/// There may be many validation types for a specific combo group and value id, but this 
		/// method only returns the first one. If you are testing whether a particular validation type 
		/// exists then use IsValidationType instead. If you want all the validation types then use 
		/// GetValidationTypesForValueID instead.
		/// </remarks>
		/// <seealso cref="IsValidationType"/>
		/// <seealso cref="GetValidationTypesForValueID"/>
		public static string GetValidationTypeForValueID(string groupName, int valueId)
		{
			return GetComboGroup(groupName).GetValidationTypeByValueId(valueId);
		}

		/// <summary>
		/// Gets the first value id for a specified combo group and validation type.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="validationType">The validation type.</param>
		/// <returns>The first matching value id.</returns>
		/// <exception cref="ErrAssistException">No matching value id.</exception>
		/// <remarks>
		/// Equivalent to "select TOP 1 VALUEID from COMBOVALIDATION where GROUPNAME=<i>groupName</i> and VALIDATIONTYPE=<i>validationType</i>".
		/// </remarks>
		/// <seealso cref="GetValueIdsForValidationType"/>
		public static int GetFirstComboValueId(string groupName, string validationType)
		{
			int valueId  = 0;
			ReadOnlyCollection<int> valueIds = GetValueIdsForValidationType(groupName, validationType);
			if (valueIds.Count > 0)
			{
				valueId = valueIds[0];
			}
			else
			{
				throw new ErrAssistException("Could not find validation type '" + validationType + "' for combo group '" + groupName + "'.");
			}
			return valueId;
		}

		/// <summary>
		/// Gets a list of value ids for a specified combo group and value name. 
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueName">The value name.</param>
		/// <returns>The list of matching value ids.</returns>
		/// <remarks>
		/// Equivalent to "select VALUEID from COMBOVALUE where GROUPNAME=<i>groupName</i> and VALUENAME=<i>valueName</i>"
		/// </remarks>
		public static ReadOnlyCollection<int> GetValueIdsForValueName(string groupName, string valueName)
		{
			return GetComboGroup(groupName).GetValueIdsByValueName(valueName);
		}

		/// <summary>
		/// Gets the combo value id for a specified combo group and value name. 
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueName">The value name.</param>
		/// <returns>The first matching value id.</returns>
		/// <remarks>
		/// Equivalent to "select TOP 1 VALUEID from COMBOVALUE where GROUPNAME=<i>groupName</i> and VALUENAME=<i>valueName</i>". 
		/// </remarks>
		/// <seealso cref="GetValueIdsForValueName"/>
		public static int GetFirstComboTextId(string groupName, string valueName)
		{
			int valueId = 0;
			ReadOnlyCollection<int> valueIds = GetValueIdsForValueName(groupName, valueName);
			if (valueIds.Count > 0)
			{
				valueId = valueIds[0];
			}
			return valueId;
		}

		/// <summary>
		/// Gets the validation types for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group.</param>
		/// <param name="valueId">The value id.</param>
		/// <returns>The matching validation types.</returns>
		/// <remarks>
		/// Equivalent to "select VALIDATIONTYPE from COMBOVALIDATION where GROUPNAME=<i>groupName</i> and VALUEID=<i>valueId</i>".
		/// </remarks>
		public static StringCollection GetValidationTypesForValueID(string groupName, int valueId)
		{
			return GetComboGroup(groupName).GetValidationTypesByValueId(valueId);
		}

		/// <summary>
		/// Gets a <see cref="ComboCollection"/> object for a specified group name.
		/// </summary>
		/// <param name="groupName">The group name.</param>
		/// <returns>The <see cref="ComboCollection"/> object.</returns>
		public static ComboCollection GetComboGroup(string groupName)
		{
			ComboCollection comboGroup = null;

			StringCollection groupNames = new StringCollection();
			groupNames.Add(groupName);
			List<ComboCollection> comboGroups = GetComboGroups(groupNames);

			if (comboGroups.Count > 0)
			{
				comboGroup = comboGroups[0];
			}

			if (comboGroup == null)
			{
				throw new ErrAssistException("Could not find combo group '" + groupName + "'.");
			}

			return comboGroup;
		}

		/// <summary>
		/// Gets the first combo value for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group name.</param>
		/// <param name="valueId">The value id.</param>
		/// <returns>The first combo value.</returns>
		public static ComboValue GetFirstComboValue(string groupName, int valueId)
		{
			return GetComboGroup(groupName).GetComboValueByValueId(valueId);
		}

		/// <summary>
		/// Gets the list of combo values for a specified combo group and value id.
		/// </summary>
		/// <param name="groupName">The combo group name.</param>
		/// <param name="valueId">The value id.</param>
		/// <returns>The list of combo values.</returns>
		public static ReadOnlyCollection<ComboValue> GetComboValues(string groupName, int valueId)
		{
			return GetComboGroup(groupName).GetComboValuesByValueId(valueId);
		}

		/// <summary>
		/// Gets a list of <see cref="ComboCollection"/> objects for a specified list of group names.
		/// </summary>
		/// <param name="groupNames">The comma separated list of group names.</param>
		/// <returns>A list of <see cref="ComboCollection"/> objects.</returns>
		public static List<ComboCollection> GetComboGroups(string groupNames)
		{
			StringCollection newGroupNames = new StringCollection();
			newGroupNames.AddRange(groupNames.Split(','));
			return GetComboGroups(newGroupNames);
		}

		/// <summary>
		/// Gets a list of <see cref="ComboCollection"/> objects for a specified list of group names.
		/// </summary>
		/// <param name="groupNames">The group names.</param>
		/// <returns>A list of <see cref="ComboCollection"/> objects.</returns>
		public static List<ComboCollection> GetComboGroups(StringCollection groupNames)
		{
			Dictionary<string, ComboCollection> comboGroups = new Dictionary<string, ComboCollection>();
			List<ComboCollection> newComboGroups = new List<ComboCollection>();

			bool containsKey = true;

			// Dictionary instances are only thread safe for multiple readers, so lock _comboGroups 
			// in case another thread is writing to it.
			StringCollection newGroupNames = new StringCollection();
			lock (_comboGroups)
			{
				foreach (string groupName in groupNames)
				{
					if (!_comboGroups.ContainsKey(groupName))
					{
						newGroupNames.Add(groupName);
						containsKey = false;
					}
				}
			}

			// Read combo values from the database outside the lock - this minimises the length of time
			// _comboGroups is locked.
			if (!containsKey)
			{
				// The key does not already exist in _comboGroups, so read the combo values from the 
				// database. Because this database read is not protected by a lock, it is possible for 
				// multiple threads to concurrently read the same combo values for the same group name, 
				// which is safe.
				DataSet dataSet = new DataSet();

				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString() + "Enlist=false;"))
				{
					using (SqlCommand sqlCommand = new SqlCommand("usp_GetComboList", sqlConnection))
					{
						sqlCommand.CommandType = CommandType.StoredProcedure;
						string listNames = GetGroupNamesWhereSql(newGroupNames);
						sqlCommand.Parameters.Add("@p_ListNames", SqlDbType.NVarChar, listNames.Length).Value = listNames;
						using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
						{
							sqlDataAdapter.Fill(dataSet);
						}
					}
				}

				ComboCollection comboGroup = null;
				string prevGroupName = null;
				foreach (DataTable dataTable in dataSet.Tables)
				{
					foreach (DataRow dataRow in dataTable.Rows)
					{
						string nextGroupName = Convert.ToString(dataRow[0]);
						if (comboGroup == null || nextGroupName != prevGroupName)
						{
							comboGroup = new ComboCollection(nextGroupName);
							comboGroups[nextGroupName] = comboGroup;
						}
						comboGroup.Add(Convert.ToInt32(dataRow[1]), Convert.ToString(dataRow[2]), Convert.ToString(dataRow[3]));
						prevGroupName = nextGroupName;
					}
				}
			}

			// Lock _comboGroups again in case we need to write to it. 
			// It is not necessary to lock any of the internal objects within _comboGroups, either here 
			// or any where else, as they are all thread safe for multiple readers, and the only writing 
			// that occurs is within this method.
			lock (_comboGroups)
			{
				foreach (string groupName in groupNames)
				{
					if (_comboGroups.ContainsKey(groupName))
					{
						// The key already exists in _comboGroups; either it was already there when we 
						// entered this method, or it has been added since by another thread whilst this 
						// thread was reading the combo values from the database. Use the existing key.
						newComboGroups.Add(_comboGroups[groupName]);
					}
					else if (comboGroups.ContainsKey(groupName))
					{
						// The key has not been added to _comboGroups by another thread since the last time 
						// we checked (at the beginning of this method), therefore add the key now, by
						// getting from the combo groups just read from the database.
						ComboCollection comboGroup = comboGroups[groupName];
						_comboGroups[groupName] = comboGroup;
						newComboGroups.Add(comboGroup);
					}
				}
			}

			return newComboGroups;
		}

		private static string GetGroupNamesWhereSql(StringCollection groupNames)
		{
			string whereSql = null;

			if (groupNames.Count == 1)
			{
				whereSql = "COMBOVALUE.GROUPNAME = '" + groupNames[0] + "'";
			}
			else
			{
				whereSql = "";
				foreach (string groupName in groupNames)
				{
					if (whereSql.Length > 0)
					{
						whereSql += ", ";
					}
					whereSql += "'" + groupName + "'";
				}
				whereSql = "COMBOVALUE.GROUPNAME IN (" + whereSql + ")";
			}

			return whereSql;
		}
	}
}

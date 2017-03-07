/*
--------------------------------------------------------------------------------------------
Workfile:			ComboCollection.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a combo group.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		01/08/2007	First .Net version. 
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// Represents a combo group.
	/// </summary>
	public class ComboCollection : IEnumerable<KeyValuePair<int, List<ComboValue>>>
	{
		private string _groupName;
		private Dictionary<int, List<ComboValue>> _values = new Dictionary<int, List<ComboValue>>();

		#region Public accessors

		/// <summary>
		/// Gets the name for the combo group. Corresponds to COMBOGROUP.GROUPNAME. 
		/// </summary>
		public string GroupName
		{
			get { return _groupName; }
		}

		#endregion

		/// <summary>
		/// Initializes a new instance of the <see cref="ComboCollection"/> class.
		/// </summary>
		/// <param name="groupName">The name for the <see cref="ComboCollection"/> instance. Corresponds to COMBOGROUP.GROUPNAME.</param>
		public ComboCollection(string groupName)
		{
			_groupName = groupName;
		}

		/// <summary>
		/// Gets the list of combo values for a valid id.
		/// </summary>
		/// <param name="valueId">The value id.</param>
		/// <returns>
		/// The list of combo values for <paramref name="valueId"/>.
		/// </returns>
		public ReadOnlyCollection<ComboValue> GetComboValuesByValueId(int valueId)
		{
			return _values.ContainsKey(valueId) ? new ReadOnlyCollection<ComboValue>(_values[valueId]) : null;
		}

		/// <summary>
		/// Adds a new combo value name and validation type to the list of existing combo values.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <param name="valueName">The name of the combo value as displayed in the Gui. Corresponds to COMBOVALUE.VALUENAME.</param>
		/// <param name="validationType">The validation type for the combo value. Corresponds to COMBOVALIDATION.VALIDATIONTYPE.</param>
		/// <remarks>
		/// If there is already a list of combo values for <paramref name="valueId"/>, then the 
		/// new combo value is added to the existing list; otherwise a new list is created and 
		/// associated with <paramref name="valueId"/>, and the new combo value is added to the new list.
		/// </remarks>
		public void Add(int valueId, string valueName, string validationType)
		{
			List<ComboValue> comboValues = _values.ContainsKey(valueId) ? _values[valueId] : null;
			if (comboValues == null)
			{
				comboValues = new List<ComboValue>();
				_values[valueId] = comboValues;
			}

			comboValues.Add(new ComboValue(valueId, valueName, validationType));
		}

		/// <summary>
		/// Gets the first combo value for a specified value id.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <returns>
		/// Generally there is only one combo value for a 
		/// specific value id, although it is possible to have more than one. 
		/// Only the first combo value for <paramref name="valueId"/> is returned. 
		/// Corresponds to "SELECT * FROM COMBOVALUE WHERE GROUPNAME = <i>groupName</i> AND VALUEID = <paramref name="valueId"/>".
		/// </returns>
		public ComboValue GetComboValueByValueId(int valueId)
		{
			ComboValue comboValue = null;
			ReadOnlyCollection<ComboValue> comboValues = GetComboValuesByValueId(valueId);
			if (comboValues != null && comboValues.Count > 0)
			{
				// Only interested in the first one - ignore others.
				comboValue = comboValues[0];
			}
			return comboValue;
		}

		/// <summary>
		/// Gets the first combo value for a specified value id and validation type.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <param name="validationType">The validation type for the combo value. Corresponds to COMBOVALIDATION.VALIDATIONTYPE.</param>
		/// <returns>
		/// The first combo value for value id <paramref name="valueId"/> and validation type <paramref name="validationType"/>. 
		/// Corresponds to 
		/// <code>
		/// SELECT COMBOVALUE.* FROM COMBOVALUE, COMBOVALIDATION 
		/// WHERE COMBOVALUE.GROUPNAME = <i>groupName</i> 
		/// AND COMBOVALUE.VALUEID = <paramref name="valueId"/>
		/// AND COMBOVALIDATION.GROUPNAME = COMBOVALUE.GROUPNAME
		/// AND COMBOVALIDATION.VALUEID = COMBOVALUE.VALUEID
		/// AND COMBOVALIDATION.VALIDATIONTYPE = <paramref name="validationType"/>
		/// </code>
		/// </returns>
		public ComboValue GetComboValueByValueIdAndValidationType(int valueId, string validationType)
		{
			ComboValue comboValue = null;
			ReadOnlyCollection<ComboValue> comboValues = GetComboValuesByValueId(valueId);
			if (comboValues != null)
			{
				foreach (ComboValue thisComboValue in comboValues)
				{
					if (string.Compare(thisComboValue.ValidationType, validationType, StringComparison.OrdinalIgnoreCase) == 0)
					{
						comboValue = thisComboValue;
						break;
					}
				}
			}
			return comboValue;
		}

		/// <summary>
		/// Gets the combo value for a specified value id.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <returns>
		/// The combo value for the value id <paramref name="valueId"/>. 
		/// Corresponds to "SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = <i>groupName</i> AND VALUEID = <paramref name="valueId"/>".
		/// </returns>
		public string GetValueNameByValueId(int valueId)
		{
			ComboValue comboValue = GetComboValueByValueId(valueId);
			return comboValue != null ? comboValue.ValueName : "";
		}

		/// <summary>
		/// Gets the combo validation type for a specified value id.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <returns>
		/// The combo validation type for the value id <paramref name="valueId"/>. 
		/// Corresponds to "SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE GROUPNAME = <i>groupName</i> 
		/// AND VALUEID = <paramref name="valueId"/>".
		/// </returns>
		public string GetValidationTypeByValueId(int valueId)
		{
			ComboValue comboValue = GetComboValueByValueId(valueId);
			return comboValue != null ? comboValue.ValidationType : "";
		}

		/// <summary>
		/// Gets the value ids for a specified validation type.
		/// </summary>
		/// <param name="validationType">The validation type for the combo value. Corresponds to COMBOVALIDATION.VALIDATIONTYPE.</param>
		/// <returns>
		/// Corresponds to "SELECT VALUEID FROM COMBOVALIDATION WHERE GROUPNAME = <i>groupName</i> 
		/// AND VALIDATIONTYPE = <paramref name="validationType"/>".
		/// </returns>
		public ReadOnlyCollection<int> GetValueIdsByValidationType(string validationType)
		{
			List<int> valueIds = new List<int>();
			foreach (KeyValuePair<int, List<ComboValue>> kvp in _values)
			{
				List<ComboValue> comboValues = kvp.Value;
				foreach (ComboValue comboValue in comboValues)
				{
					if (string.Compare(comboValue.ValidationType, validationType, StringComparison.OrdinalIgnoreCase) == 0 &&
						valueIds.IndexOf(comboValue.ValueId) == -1)
					{
						valueIds.Add(comboValue.ValueId);
					}
				}
			}
			return new ReadOnlyCollection<int>(valueIds);
		}

		/// <summary>
		/// Gets the value ids for a specified value name.
		/// </summary>
		/// <param name="valueName">The name of the combo value as displayed in the Gui. Corresponds to COMBOVALUE.VALUENAME.</param>
		/// <returns>
		/// Corresponds to "SELECT VALUEID FROM COMBOVALUE WHERE GROUPNAME = <i>groupName</i> 
		/// AND VALUENAME = <paramref name="valueName"/>".
		/// </returns>
		public ReadOnlyCollection<int> GetValueIdsByValueName(string valueName)
		{
			List<int> valueIds = new List<int>();
			foreach (KeyValuePair<int, List<ComboValue>> kvp in _values)
			{
				List<ComboValue> comboValues = kvp.Value;
				foreach (ComboValue comboValue in comboValues)
				{
					if (string.Compare(comboValue.ValueName, valueName, StringComparison.OrdinalIgnoreCase) == 0 &&
						valueIds.IndexOf(comboValue.ValueId) == -1)
					{
						valueIds.Add(comboValue.ValueId);
					}
				}
			}

			return new ReadOnlyCollection<int>(valueIds);
		}

		/// <summary>
		/// Gets the validation types for a specified value id.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <returns>
		/// Corresponds to "SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE GROUPNAME = <i>groupName</i> 
		/// AND VALUEID = <paramref name="valueId"/>".
		/// </returns>
		public StringCollection GetValidationTypesByValueId(int valueId)
		{
			StringCollection validationTypes = new StringCollection();
			ReadOnlyCollection<ComboValue> comboValues = GetComboValuesByValueId(valueId);
			if (comboValues != null)
			{
				foreach (ComboValue comboValue in comboValues)
				{
					validationTypes.Add(comboValue.ValidationType);
				}
			}
			return validationTypes;
		}

		#region IEnumerable<KeyValuePair<int,List<ComboValue>>> Members

		/// <summary>
		/// Returns an enumerator that iterates through the combo values for this combo group.
		/// </summary>
		/// <remarks>
		/// For the purposes of enumeration, each item is a 
		/// <b>System.Collections.Generic.KeyValuePair</b> structure 
		/// representing a value and its key. The value is a combo value id, and the key is a list of 
		/// the <see cref="ComboValue"/> objects for that combo value id. Therefore each 
		/// enumerated item represents a SQL query 
		/// "SELECT * FROM COMBOVALUE WHERE GROUPNAME = <i>groupName</i> AND VALUEID = <i>valueId</i>".
		/// </remarks>
		public IEnumerator<KeyValuePair<int, List<ComboValue>>> GetEnumerator()
		{			
			return _values.GetEnumerator();
		}

		#endregion

		#region IEnumerable Members

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return _values.GetEnumerator();
		}

		#endregion
	}
}

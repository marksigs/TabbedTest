/*
--------------------------------------------------------------------------------------------
Workfile:			ComboValue.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a combo value.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		01/08/2007	First .Net version. 
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// Represents a combo value.
	/// </summary>
	public class ComboValue
	{
		private int _valueId;
		private string _valueName = "";
		private string _validationType = "";

		/// <summary>
		/// Gets the identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.
		/// </summary>
		public int ValueId
		{
			get { return _valueId; }
		}

		/// <summary>
		/// Gets the name of the combo value as displayed in the Gui. Corresponds to COMBOVALUE.VALUENAME.
		/// </summary>
		public string ValueName
		{
			get { return _valueName; }
		}

		/// <summary>
		/// Gets the validation type for the combo value. Corresponds to COMBOVALIDATION.VALIDATIONTYPE.
		/// </summary>
		public string ValidationType
		{
			get { return _validationType; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ComboValue"/> class with default values.
		/// </summary>
		public ComboValue()
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ComboValue"/> class.
		/// </summary>
		/// <param name="valueId">The identifier for the combo value. Corresponds to COMBOVALUE.VALUEID.</param>
		/// <param name="valueName">The name of the combo value as displayed in the Gui. Corresponds to COMBOVALUE.VALUENAME.</param>
		/// <param name="validationType">The validation type for the combo value. Corresponds to SELECT VALIDATIONTYPE FROM COMBOVALIDATION.</param>
		public ComboValue(int valueId, string valueName, string validationType)
		{
			_valueId = valueId;
			_valueName = valueName;
			_validationType = validationType;
		}
	}
}

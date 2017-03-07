/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterTypeSafe.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a type safe global parameter.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		06/08/2007	First .Net version. 
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// The base class for type global parameters.
	/// </summary>
	/// <remarks>
	/// The classes that derive from <see cref="GlobalParameterTypeSafe"/> are type safe, i.e., 
	/// they have a <b>Value</b> property that returns a typed value for the global parameter. 
	/// Compare this to <see cref="GlobalParameter"/> which is not type safe because it 
	/// contains properties for all of the possible types of value 
	/// for the global parameter.
	/// </remarks>
	public abstract class GlobalParameterTypeSafe : GlobalParameterTypeSafeBase
	{
		/// <summary>
		/// Initializes a new instance of the GlobalParameterTypeSafe class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an exception if the global parameter does not exist.</param>
		/// <remarks>
		/// The GlobalParameterTypeSafe object is initialized by calling the 
		/// <see cref="GlobalParameter"/> constructor.
		/// </remarks>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the global parameter does not exist.
		/// </exception>
		protected GlobalParameterTypeSafe(string name, bool mandatory) 
			: base(new GlobalParameter(name), mandatory)
		{
		}

		/// <summary>
		/// Converts this instance to its equivalent <see cref="String"/> representation.
		/// </summary>
		/// <returns>
		/// A <see cref="String"/> representing the value of this instance.
		/// </returns>
		/// <remarks>
		/// See <see cref="Vertex.Fsd.Omiga.VisualBasicPort.Core.GlobalParameter.ToString"/>.
		/// </remarks>
		public override string ToString()
		{
			return GlobalParameterBase.ToString();
		}
	}

	/// <summary>
	/// Represents a type safe global parameter that maps onto the GLOBALPARAMETER.STRING field.
	/// </summary>
	public class GlobalParameterString : GlobalParameterTypeSafe
	{
		/// <summary>
		/// The value of the STRING field for the GLOBALPARAMETER record; may be null.
		/// </summary>
		public string Value
		{
			get { Initialize(); return GlobalParameterBase.String; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterString class; 
		/// does not throw an exception if the GLOBALPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterString(string name)
			: this(name, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterString class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterString(string name, bool mandatory)
			: base(name, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global parameter that maps onto the GLOBALPARAMETER.AMOUNT field.
	/// </summary>
	public class GlobalParameterAmount : GlobalParameterTypeSafe
	{
		/// <summary>
		/// The value of the AMOUNT field for the GLOBALPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.Amount; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterAmount class; 
		/// does not throw an exception if the GLOBALPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterAmount(string name)
			: this(name, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterAmount class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterAmount(string name, bool mandatory)
			: base(name, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global parameter that maps onto the GLOBALPARAMETER.MAXIMUMAMOUNT field.
	/// </summary>
	public class GlobalParameterMaximumAmount : GlobalParameterTypeSafe
	{
		/// <summary>
		/// The value of the MAXIMUMAMOUNT field for the GLOBALPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.MaximumAmount; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterMaximumAmount class; 
		/// does not throw an exception if the GLOBALPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterMaximumAmount(string name)
			: this(name, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterMaximumAmount class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterMaximumAmount(string name, bool mandatory)
			: base(name, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global parameter that maps onto the GLOBALPARAMETER.PERCENTAGE field.
	/// </summary>
	public class GlobalParameterPercentage : GlobalParameterTypeSafe
	{
		/// <summary>
		/// The value of the PERCENTAGE field for the GLOBALPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.Percentage; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterPercentage class; 
		/// does not throw an exception if the GLOBALPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterPercentage(string name)
			: this(name, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterPercentage class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterPercentage(string name, bool mandatory)
			: base(name, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global parameter that maps onto the GLOBALPARAMETER.BOOLEAN field.
	/// </summary>
	public class GlobalParameterBoolean : GlobalParameterTypeSafe
	{
		/// <summary>
		/// The value of the BOOLEAN field for the GLOBALPARAMETER record; may be null.
		/// </summary>
		public bool? Value
		{
			get { Initialize(); return GlobalParameterBase.Boolean; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterBoolean class; 
		/// does not throw an exception if the GLOBALPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterBoolean(string name)
			: this(name, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterBoolean class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		public GlobalParameterBoolean(string name, bool mandatory)
			: base(name, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}
}

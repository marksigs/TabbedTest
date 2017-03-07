/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalBandedParameterTypeSafe.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a type safe global banded parameter.
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
	/// The base class for type global banded parameters.
	/// </summary>
	/// <remarks>
	/// The classes that derive from <see cref="GlobalBandedParameterTypeSafe"/> are type safe, i.e., 
	/// they have a <b>Value</b> property that returns a typed value for the global banded parameter. 
	/// Compare this to <see cref="GlobalBandedParameter"/> which is not type safe because it 
	/// contains properties for all of the possible types of value 
	/// for the global banded parameter.
	/// </remarks>
	public abstract class GlobalBandedParameterTypeSafe : GlobalParameterTypeSafeBase
	{
		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterTypeSafe class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an exception if the global banded parameter does not exist.</param>
		/// <remarks>
		/// The GlobalBandedParameterTypeSafe object is initialized by calling the 
		/// <see cref="GlobalBandedParameter"/> constructor.
		/// </remarks>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the global banded parameter does not exist.
		/// </exception>
		protected GlobalBandedParameterTypeSafe(string name, double minimumHighBand, bool mandatory) 
			: base(new GlobalBandedParameter(name, minimumHighBand), mandatory)
		{
		}

		/// <summary>
		/// Converts this instance to its equivalent <see cref="String"/> representation.
		/// </summary>
		/// <returns>
		/// A <see cref="String"/> representing the value of this instance.
		/// </returns>
		/// <remarks>
		/// See <see cref="Vertex.Fsd.Omiga.VisualBasicPort.Core.GlobalBandedParameter.ToString"/>.
		/// </remarks>
		public override string ToString()
		{
			return GlobalParameterBase.ToString();
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.STRING field.
	/// </summary>
	public class GlobalBandedParameterString : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the STRING field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public string Value
		{
			get { Initialize(); return GlobalParameterBase.String; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterString class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterString(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterString class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterString(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.AMOUNT field.
	/// </summary>
	public class GlobalBandedParameterAmount : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the AMOUNT field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.Amount; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterAmount class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterAmount(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterAmount class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterAmount(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.MAXIMUMAMOUNT field.
	/// </summary>
	public class GlobalBandedParameterMaximum : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the MAXIMUMAMOUNT field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.MaximumAmount; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterMaximum class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterMaximum(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterMaximum class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterMaximum(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.PERCENTAGE field.
	/// </summary>
	public class GlobalBandedParameterPercentage : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the PERCENTAGE field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public double? Value
		{
			get { Initialize(); return GlobalParameterBase.Percentage; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterPercentage class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterPercentage(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterPercentage class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterPercentage(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.BOOLEAN field.
	/// </summary>
	public class GlobalBandedParameterBoolean : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the BOOLEAN field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public bool? Value
		{
			get { Initialize(); return GlobalParameterBase.Boolean; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterBoolean class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterBoolean(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterBoolean class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterBoolean(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}

	/// <summary>
	/// Represents a type safe global banded parameter that maps onto the GLOBALBANDEDPARAMETER.BOOLEANTEXT field.
	/// </summary>
	public class GlobalBandedParameterBooleanText : GlobalBandedParameterTypeSafe
	{
		/// <summary>
		/// The value of the BOOLEANTEXT field for the GLOBALBANDEDPARAMETER record; may be null.
		/// </summary>
		public string Value
		{
			get { Initialize(); return ((GlobalBandedParameter)GlobalParameterBase).BooleanText; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterBooleanText class; 
		/// does not throw an exception if the GLOBALBANDEDPARAMETER record does not exist.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterBooleanText(string name, double minimumHighBand)
			: this(name, minimumHighBand, false)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameterBooleanText class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <param name="mandatory">Indicates whether to throw an <see cref="ArgumentException"/> if the GLOBALBANDEDPARAMETER record does not exist.</param>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the GLOBALBANDEDPARAMETER record does not exist.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		public GlobalBandedParameterBooleanText(string name, double minimumHighBand, bool mandatory)
			: base(name, minimumHighBand, mandatory)
		{
		}

		/// <summary>
		/// Indicates whether the global banded parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global banded parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public override bool HasValue()
		{
			return Value != null;
		}
	}
}

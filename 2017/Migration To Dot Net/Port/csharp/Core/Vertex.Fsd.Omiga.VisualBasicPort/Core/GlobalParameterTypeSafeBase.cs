/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterTypeSafeBase.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		The base class for type safe global parameters and global banded parameters.
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
	/// Represents a type global parameter that maps onto a record either on the 
	/// GLOBALPARAMETER or GLOBALBANDEDPARAMETER database table.
	/// </summary>
	public abstract class GlobalParameterTypeSafeBase
	{
		private GlobalParameterBase _globalParameterBase;
		private bool _mandatory;
		private bool _initialized;

		/// <summary>
		/// Gets the internal <see cref="GlobalParameterBase"/> object.
		/// </summary>
		protected GlobalParameterBase GlobalParameterBase
		{
			get { Initialize(); return _globalParameterBase; }
		}

		/// <summary>
		/// The name of the global parameter; maps onto GLOBALPARAMETER.NAME or GLOBALBANDEDPARAMETER.NAME.
		/// </summary>
		public string Name
		{
			get { Initialize(); return _globalParameterBase.Name; }
		}

		/// <summary>
		/// The start date of the global parameter; maps onto GLOBALPARAMETER.GLOBALPARAMETERSTARTDATE or GLOBALBANDEDPARAMETER.GBPARAMSTARTDATE.
		/// </summary>
		public DateTime StartDate
		{
			get { Initialize(); return _globalParameterBase.StartDate; }
		}

		/// <summary>
		/// The description of the global parameter; maps onto GLOBALPARAMETER.DESCRIPTION or GLOBALBANDEDPARAMETER.DESCRIPTION; may be null.
		/// </summary>
		public string Description
		{
			get { Initialize(); return _globalParameterBase.Description; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterTypeSafeBase class.
		/// </summary>
		/// <param name="globalParameterBase">The <see cref="GlobalParameterBase"/> object of this GlobalParameterTypeSafeBase instance should encapsulate; maps onto GLOBALPARAMETER.NAME.</param>
		/// <param name="mandatory">Indicates whether to throw an exception if the global parameter does not exist.</param>
		/// <remarks>
		/// The GlobalParameterTypeSafeBase object wraps an object of a type derived from the 
		/// <see cref="GlobalParameterBase"/> type.
		/// </remarks>
		/// <exception cref="ArgumentException">
		/// <paramref name="mandatory"/> is <b>true</b> and the global parameter does not exist.
		/// </exception>
		protected GlobalParameterTypeSafeBase(GlobalParameterBase globalParameterBase, bool mandatory)
		{
			_globalParameterBase = globalParameterBase;
			_mandatory = mandatory;
		}

		/// <summary>
		/// Populate this instance from the GLOBALPARAMETER or GLOBALBANDEDPARAMETER table.
		/// </summary>
		/// <returns>
		/// This instance, enabling the following usage:
		/// <code>
		/// double amount = new GlobalParameterAmount("MyParameter").Initialize().Value;
		/// </code>
		/// </returns>
		/// <remarks>
		/// In public properties and methods in derived classes <b>Initialize</b> should be 
		/// called first to ensure that the global parameter has been initialized correctly. 
		/// For example, if the <b>Value</b> property calls <b>Initialize</b>, then the 
		/// following usage is possible:
		/// <code>
		/// double amount = new GlobalParameterAmount("MyParameter").Value;
		/// </code>
		/// </remarks>
		public virtual GlobalParameterTypeSafeBase Initialize()
		{
			if (!_initialized)
			{
				// Set flag first to prevent infinite recursion if HasValue override calls 
				// Initialize.
				_initialized = true;

				bool hasValue = true;

				try
				{
					_globalParameterBase.Initialize();

					if (_mandatory)
					{
						hasValue = HasValue();
					}
				}
				catch
				{
					// Initialization failed.
					_initialized = false;
					throw;
				}

				if (_mandatory && !hasValue)
				{
					_initialized = false;
					throw new ArgumentException("Missing mandatory global parameter '" + _globalParameterBase.Name + "'");
				}
			}

			return this;
		}
		
		/// <summary>
		/// Indicates whether the global parameter has a value.
		/// </summary>
		/// <returns>
		/// <b>true</b> if the global parameter has a value, otherwise <b>false</b>.
		/// </returns>
		public abstract bool HasValue();

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
			return _globalParameterBase.ToString();
		}
	}
}

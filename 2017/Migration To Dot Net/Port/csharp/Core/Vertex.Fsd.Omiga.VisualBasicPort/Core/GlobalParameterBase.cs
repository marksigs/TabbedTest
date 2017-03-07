/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterBase.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		The base class for global parameters and global banded parameters.
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
	/// Represents a global parameter that maps onto a record either on the 
	/// GLOBALPARAMETER or GLOBALBANDEDPARAMETER database table.
	/// </summary>
	public abstract class GlobalParameterBase
	{
		private bool _initialized;
		private string _name;
		private DateTime _startDate;
		private string _description;
		private double? _amount;
		private double? _maximumAmount;
		private double? _percentage;
		private bool? _boolean;
		private string _string;

		/// <summary>
		/// The name of the global parameter; maps onto GLOBALPARAMETER.NAME or GLOBALBANDEDPARAMETER.NAME.
		/// </summary>
		public string Name
		{
			get { Initialize(); return _name; }
			protected set { _name = value; }
		}

		/// <summary>
		/// The start date of the global parameter; maps onto GLOBALPARAMETER.GLOBALPARAMETERSTARTDATE or GLOBALBANDEDPARAMETER.GBPARAMSTARTDATE.
		/// </summary>
		public DateTime StartDate
		{
			get { Initialize(); return _startDate; }
			protected set { _startDate = value; }
		}

		/// <summary>
		/// The description of the global parameter; maps onto GLOBALPARAMETER.DESCRIPTION or GLOBALBANDEDPARAMETER.DESCRIPTION; may be null.
		/// </summary>
		public string Description
		{
			get { Initialize(); return _description; }
			protected set { _description = value; }
		}

		/// <summary>
		/// The amount of the global parameter; maps onto GLOBALPARAMETER.AMOUNT or GLOBALBANDEDPARAMETER.AMOUNT; may be null.
		/// </summary>
		public double? Amount
		{
			get { Initialize(); return _amount; }
			protected set { _amount = value; }
		}

		/// <summary>
		/// The maximum amount of the global parameter; maps onto GLOBALPARAMETER.MAXIMUMAMOUNT or GLOBALBANDEDPARAMETER.MAXIMUM; may be null.
		/// </summary>
		public double? MaximumAmount
		{
			get { Initialize(); return _maximumAmount; }
			protected set { _maximumAmount = value; }
		}

		/// <summary>
		/// The percentage of the global parameter; maps onto GLOBALPARAMETER.PERCENTAGE or GLOBALBANDEDPARAMETER.PERCENTAGE; may be null.
		/// </summary>
		public double? Percentage
		{
			get { Initialize(); return _percentage; }
			protected set { _percentage = value; }
		}

		/// <summary>
		/// The boolean of the global parameter; maps onto GLOBALPARAMETER.BOOLEAN or GLOBALBANDEDPARAMETER.BOOLEAN; may be null.
		/// </summary>
		public bool? Boolean
		{
			get { Initialize(); return _boolean; }
			protected set { _boolean = value; }
		}

		/// <summary>
		/// The string of the global parameter; maps onto GLOBALPARAMETER.STRING or GLOBALBANDEDPARAMETER.STRING; may be null.
		/// </summary>
		public string String
		{
			get { Initialize(); return _string; }
			protected set { _string = value; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalParameterBase class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME or GLOBALBANDEDPARAMETER.NAME.</param>
		/// <exception cref="ArgumentNullException">
		/// The <paramref name="name"/> parameter has not been specified.
		/// </exception>
		protected GlobalParameterBase(string name)
		{
			if (name == null || name.Length == 0)
			{
				throw new ArgumentNullException("name", "No global parameter name specified.");
			}

			_name = name;
		}

		/// <summary>
		/// Populate this instance from the GLOBALPARAMETER or GLOBALBANDEDPARAMETER table.
		/// </summary>
		/// <returns>
		/// This instance, enabling the following usage:
		/// <code>
		/// string response = new GlobalParameter("MyParameter").Initialize().ToString();
		/// </code>
		/// </returns>
		/// <remarks>
		/// In public properties and methods in derived classes <b>Initialize</b> should be 
		/// called first to ensure that the global parameter has been initialized correctly. 
		/// For example, if the <b>ToString</b> method override calls <b>Initialize</b>, then the 
		/// following usage is possible:
		/// <code>
		/// string response = new GlobalParameter("MyParameter").ToString();
		/// </code>
		/// </remarks>
		public virtual GlobalParameterBase Initialize()
		{
			if (!_initialized)
			{
				// Set flag first to prevent infinite recursion if Load override calls 
				// Initialize.
				_initialized = true;
				try
				{
					Load();
				}
				catch
				{
					// Initialization failed.
					_initialized = false;
					throw;
				}
			}

			return this;
		}

		/// <summary>
		/// Load the global parameter.
		/// </summary>
		/// <remarks>
		/// This method must be implemented in derived classes to load the global parameter, 
		/// e.g., from the database. Do not call <see cref="Initialize"/> in the <b>Load</b> 
		/// implementation.
		/// </remarks>
		protected abstract void Load();
	}
}

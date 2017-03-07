/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameter.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a global parameter.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		01/08/2007	First .Net version. 
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// Represents a global parameter that maps onto a record on the GLOBALPARAMETER database table.
	/// </summary>
	/// <remarks>
	/// The <see cref="GlobalParameter"/> class is not type safe: it contains properties 
	/// for all of the posssible types of value for a global parameter; generally only one 
	/// of these properties will contain a value, the rest will be null. 
	/// For type safe classes see <see cref="GlobalParameterAmount"/>, 
	/// <see cref="GlobalParameterMaximumAmount"/>, <see cref="GlobalParameterPercentage"/>, 
	/// <see cref="GlobalParameterBoolean"/> and <see cref="GlobalParameterString"/>.
	/// </remarks>
	public class GlobalParameter : GlobalParameterBase
	{
		/// <summary>
		/// Initializes a new instance of the GlobalParameter class.
		/// </summary>
		/// <param name="name">The name of the global parameter; maps onto GLOBALPARAMETER.NAME.</param>
		/// <remarks>
		/// The GlobalParameter object is initialized by executing the <b>usp_GetGlobalParameter</b> 
		/// stored procedure.
		/// </remarks>
		/// <exception cref="ArgumentNullException">
		/// The <paramref name="name"/> parameter has not been specified.
		/// </exception>
		public GlobalParameter(string name) : base(name)
		{
		}

		/// <summary>
		/// Populates this instance from the GLOBALPARAMETER table.
		/// </summary>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALPARAMETER record.
		/// </exception>
		protected override void Load()
		{
			try
			{
				string connectionString = GlobalProperty.DatabaseConnectionString + "Enlist=false;";

				using (SqlConnection sqlConnection = new SqlConnection(connectionString))
				{
					using (SqlCommand sqlCommand = new SqlCommand("usp_GetGlobalParameter", sqlConnection))
					{
						sqlCommand.CommandType = CommandType.StoredProcedure;
						sqlCommand.Parameters.Add("@p_Name", SqlDbType.NVarChar, Name.Length).Value = Name;

						sqlConnection.Open();
						using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
						{
							if (sqlDataReader.HasRows && sqlDataReader.Read())
							{
								if (!sqlDataReader.IsDBNull(1)) StartDate = sqlDataReader.GetDateTime(1);
								if (!sqlDataReader.IsDBNull(2)) String = sqlDataReader.GetString(2);
								if (!sqlDataReader.IsDBNull(3)) Description = sqlDataReader.GetString(3);
								if (!sqlDataReader.IsDBNull(4)) Amount = sqlDataReader.GetDouble(4);
								if (!sqlDataReader.IsDBNull(5)) MaximumAmount = sqlDataReader.GetDouble(5);
								if (!sqlDataReader.IsDBNull(6)) Percentage = sqlDataReader.GetDouble(6);
								if (!sqlDataReader.IsDBNull(7)) Boolean = sqlDataReader.GetByte(7) == 1;
							}
							else
							{
								throw new InvalidOperationException("Missing record for global parameter " + Name + ".");
							}
						}
					}
				}
			}
			catch (InvalidOperationException)
			{
				throw;
			}
			catch (Exception ex)
			{
				throw new InvalidOperationException("Unable to read global parameter " + Name + ".", ex);
			}
		}

		/// <summary>
		/// Converts this instance to its equivalent <see cref="String"/> representation.
		/// </summary>
		/// <returns>
		/// A <see cref="String"/> representing the value of this instance.
		/// </returns>
		/// <remarks>
		/// The <see cref="String"/> is in xml format and matches the format used by the 
		/// Omiga Visual Basic 6 method <b>omBase.GlobalParameterDO.GetCurrentParameter</b>.
		/// </remarks>
		public override string ToString()
		{
			Initialize();

			// Note: do not use xml serialization as this does not produce a backwards compatible format.
			string xmlString = null;

			using (StringWriter stringWriter = new StringWriter())
			{
				XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
				xmlWriterSettings.OmitXmlDeclaration = true;
				using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, xmlWriterSettings))
				{
					xmlWriter.WriteStartDocument();
					xmlWriter.WriteStartElement("GLOBALPARAMETER");
					xmlWriter.WriteElementString("NAME", Name);
					xmlWriter.WriteElementString("GLOBALPARAMETERSTARTDATE", StartDate.ToString("dd/MM/yyyy"));
					if (String != null) xmlWriter.WriteElementString("STRING", String);
					if (Description != null) xmlWriter.WriteElementString("DESCRIPTION", Description);
					if (Amount.HasValue) xmlWriter.WriteElementString("AMOUNT", Amount.Value.ToString());
					if (MaximumAmount.HasValue) xmlWriter.WriteElementString("MAXIMUMAMOUNT", MaximumAmount.Value.ToString());
					if (Percentage.HasValue) xmlWriter.WriteElementString("PERCENTAGE", Percentage.Value.ToString());
					if (Boolean.HasValue) xmlWriter.WriteElementString("BOOLEAN", Boolean.Value ? "1" : "0");
					xmlWriter.WriteEndElement();
					xmlWriter.WriteEndDocument();
					xmlWriter.Flush();
				}

				xmlString = stringWriter.ToString();
			}

			return xmlString;
		}
	}

}


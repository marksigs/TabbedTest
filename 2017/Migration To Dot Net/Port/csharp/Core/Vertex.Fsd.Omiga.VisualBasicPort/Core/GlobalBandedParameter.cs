/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalBandedParameter.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Represents a global banded parameter.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		02/08/2007	First .Net version. 
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
	/// Represents a global banded parameter that maps onto a record on the GLOBALBANDEDPARAMETER database table.
	/// </summary>
	public class GlobalBandedParameter : GlobalParameterBase
	{
		private double _minimumHighBand;
		private double _highBand;
		private string _booleanText;

		/// <summary>
		/// The high band of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.HIGHBAND.
		/// </summary>
		public double HighBand
		{
			get { Initialize(); return _highBand; }
		}

		/// <summary>
		/// The boolean text of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.BOOLEANTEXT; may be null.
		/// </summary>
		public string BooleanText
		{
			get { Initialize(); return _booleanText; }
		}

		/// <summary>
		/// Initializes a new instance of the GlobalBandedParameter class.
		/// </summary>
		/// <param name="name">The name of the global banded parameter; maps onto GLOBALBANDEDPARAMETER.NAME.</param>
		/// <param name="minimumHighBand">The minimum high band of this instance.</param>
		/// <remarks>
		/// The GlobalBandedParameter object is initialized by getting the GLOBALBANDEDPARAMETER record 
		/// where NAME = <paramref name="name"/> 
		/// and GBPARAMSTARTDATE is the maximum GBPARAMSTARTDATE &lt;= today 
		/// and HIGHBAND is &gt;= <paramref name="minimumHighBand"/>.
		/// </remarks>
		/// <exception cref="ArgumentNullException">
		/// The <paramref name="name"/> parameter has not been specified.
		/// </exception>
		public GlobalBandedParameter(string name, double minimumHighBand) : base(name)
		{
			_minimumHighBand = minimumHighBand;
		}

		/// <summary>
		/// Populate this instance from the GLOBALBANDEDPARAMETER table.
		/// </summary>
		/// <exception cref="InvalidOperationException">
		/// An exception occurred when trying to read the GLOBALBANDEDPARAMETER record.
		/// </exception>
		protected override void Load()
		{
			try
			{
				string connectionString = GlobalProperty.DatabaseConnectionString + "Enlist=false;";

				using (SqlConnection sqlConnection = new SqlConnection(connectionString))
				{
					const string cmdText =
						"SELECT * FROM GLOBALBANDEDPARAMETER GBP1 WITH(NOLOCK) " +
						"WHERE NAME = @name " +
						"AND GBP1.GBPARAMSTARTDATE = " +
						"(SELECT MAX(GBPARAMSTARTDATE) FROM GLOBALBANDEDPARAMETER WITH(NOLOCK) " +
						"WHERE NAME = @name " +
						"AND GBPARAMSTARTDATE <= @systemDate) " +
						"AND GBP1.HIGHBAND = " +
						"(SELECT MIN(HIGHBAND) FROM GLOBALBANDEDPARAMETER WITH(NOLOCK) " +
						"WHERE NAME = @name " +
						"AND HIGHBAND >= @minimumHighBand " +
						"AND GBPARAMSTARTDATE = GBP1.GBPARAMSTARTDATE)";
					using (SqlCommand sqlCommand = new SqlCommand(cmdText, sqlConnection))
					{
						sqlCommand.Parameters.Add("name", SqlDbType.NVarChar, Name.Length).Value = Name;
						sqlCommand.Parameters.Add("systemDate", SqlDbType.DateTime).Value = DateTime.Now;
						sqlCommand.Parameters.Add("minimumHighBand", SqlDbType.Float).Value = _minimumHighBand;

						sqlConnection.Open();
						using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
						{
							if (sqlDataReader.HasRows && sqlDataReader.Read())
							{
								if (!sqlDataReader.IsDBNull(1)) StartDate = sqlDataReader.GetDateTime(1);
								if (!sqlDataReader.IsDBNull(2)) _highBand = sqlDataReader.GetDouble(2);
								if (!sqlDataReader.IsDBNull(3)) Description = sqlDataReader.GetString(3);
								if (!sqlDataReader.IsDBNull(4)) Amount = sqlDataReader.GetDouble(4);
								if (!sqlDataReader.IsDBNull(5)) Percentage = sqlDataReader.GetDouble(6);
								if (!sqlDataReader.IsDBNull(6)) MaximumAmount = sqlDataReader.GetDouble(5);
								if (!sqlDataReader.IsDBNull(7)) Boolean = sqlDataReader.GetByte(7) == 1;
								if (!sqlDataReader.IsDBNull(8)) String = sqlDataReader.GetString(8);
								if (!sqlDataReader.IsDBNull(9)) _booleanText = sqlDataReader.GetString(9);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new InvalidOperationException("Unable to read global banded parameter " + Name + ".", ex);
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
		/// Omiga Visual Basic 6 method <b>omBase.GlobalBandedParameterDO.GetCurrentParameter</b>.
		/// </remarks>
		public override string ToString()
		{
			Initialize();

			string xmlString = null;

			// Note: do not use xml serialization as this does not produce a backwards compatible format.
			using (StringWriter stringWriter = new StringWriter())
			{
				XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
				xmlWriterSettings.OmitXmlDeclaration = true;
				using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, xmlWriterSettings))
				{
					xmlWriter.WriteStartDocument();
					xmlWriter.WriteStartElement("GLOBALBANDEDPARAMETER");
					xmlWriter.WriteElementString("NAME", Name);
					xmlWriter.WriteElementString("GBPARAMSTARTDATE", StartDate.ToString("dd/MM/yyyy"));
					WriteXmlElement(xmlWriter, "HIGHBAND", _highBand);
					WriteXmlElement(xmlWriter, "AMOUNT", Amount);
					WriteXmlElement(xmlWriter, "PERCENTAGE", Percentage);
					WriteXmlElement(xmlWriter, "MAXIMUM", MaximumAmount);
					xmlWriter.WriteElementString("BOOLEAN", Boolean.HasValue ? (Boolean.Value ? "1" : "0") : "");
					xmlWriter.WriteElementString("STRING", String != null ? String : "");
					xmlWriter.WriteEndElement();
					xmlWriter.WriteEndDocument();
					xmlWriter.Flush();
				}

				xmlString = stringWriter.ToString();
			}

			return xmlString;
		}

		private static void WriteXmlElement(XmlWriter xmlWriter, string name, double value)
		{
			xmlWriter.WriteStartElement(name);
			xmlWriter.WriteAttributeString("RAW", value.ToString());
			xmlWriter.WriteValue(value.ToString("0.00"));
			xmlWriter.WriteEndElement();
		}

		private static void WriteXmlElement(XmlWriter xmlWriter, string name, double? value)
		{
			if (value.HasValue)
			{
				WriteXmlElement(xmlWriter, name, (double)value);
			}
			else
			{
				xmlWriter.WriteElementString(name, "");
			}
		}
	}
}

/*
--------------------------------------------------------------------------------------------
Workfile:			DbXmlAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Provides routines to return data selected from the database as xml.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
TW		13/05/2005	Initial Creation
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		11/07/2007	First .Net version. Ported from dbxmlAssist.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Contains methods for executing select queries on the database and returning the records as xml.
	/// </summary>
	public static class DbXmlAssist
	{
		/// <summary>
		/// Gets a record from the GLOBALPARAMETER table as an <see cref="XmlNode"/> object.
		/// </summary>
		/// <param name="parameterName">The name of the global parameter (corresponds to GLOBALPARAMETER.NAME).</param>
		/// <returns>
		/// The GLOBALPARAMETER record as a <see cref="XmlNode"/>.
		/// </returns>
		public static XmlNode GetCurrentParameterXML(string parameterName)
		{
			return GetDataFromDatabase("SELECT NAME, CONVERT(varchar, GLOBALPARAMETERSTARTDATE, 103) AS GLOBALPARAMETERSTARTDATE, DESCRIPTION, AMOUNT, MAXIMUMAMOUNT, PERCENTAGE, BOOLEAN, STRING FROM GLOBALPARAMETER WHERE NAME = '" + parameterName + "' FOR XML AUTO, ELEMENTS");
		}

		/// <summary>
		/// Gets a record from the TEMPLATE table as an <see cref="XmlNode"/> object.
		/// </summary>
		/// <param name="templateId">The unique identifier for the template (corresponds to TEMPLATE.TEMPLATEID).</param>
		/// <returns>
		/// The TEMPLATE record as a <see cref="XmlNode"/>.
		/// </returns>
		public static XmlNode GetTemplateXML(string templateId)
		{
			return GetDataFromDatabase("SELECT * FROM TEMPLATE WHERE TEMPLATEID = " + templateId + " FOR XML AUTO, ELEMENTS");
		}

		private static XmlNode GetDataFromDatabase(string cmdText) 
		{
			XmlNode xmlNode = null;

			// Get the requested data as an xml node
			using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
			{
				using (SqlCommand sqlCommand = new SqlCommand(cmdText, sqlConnection))
				{
					sqlConnection.Open();
					using (XmlReader xmlReader = sqlCommand.ExecuteXmlReader())
					{					
						XmlDocument xmlDocument = new XmlDocument();
						try
						{
							if (xmlReader.Read())
							{
								string xml = null;
								while (xmlReader.ReadState != ReadState.EndOfFile)
								{
									xml += xmlReader.ReadOuterXml();
								}
								xmlDocument.LoadXml(xml);
								xmlNode = xmlDocument.DocumentElement;
							}
							else
							{
								throw new XmlException("No more nodes to read.");
							}
						}
						catch (XmlException exception)
						{
							XmlElement xmlElement = xmlDocument.CreateElement("XML_ERROR");
							xmlElement.SetAttribute("ERRORCODE", exception.GetType().ToString());
							xmlElement.SetAttribute("LINE", exception.LineNumber.ToString());
							xmlElement.SetAttribute("LINEPOS", exception.LinePosition.ToString());
							xmlElement.SetAttribute("REASON", exception.Message);
							xmlElement.InnerText = exception.Source;
							xmlDocument.AppendChild(xmlElement);
							xmlNode = xmlDocument.DocumentElement;
						}
					}
				}
			}

			return xmlNode;
		}
	}
}

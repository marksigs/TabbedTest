/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Reading parameters from the GLOBALPARAMETERS table.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from GlobalAssist.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// The GlobalAssist class contains methods for reading global parameters from the database.
	/// </summary>
	public static class GlobalAssist
	{
		/// <summary>
		/// Gets a BOOLEAN value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the BOOLEAN column where NAME = <i>name</i>.</returns>
		public static bool GetGlobalParamBoolean(string name) 
		{
			return new GlobalParameterBoolean(name).Value ?? false;
		}

		/// <summary>
		/// Gets a mandatory BOOLEAN value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the BOOLEAN column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static bool GetMandatoryGlobalParamBoolean(string name) 
		{
			return new GlobalParameterBoolean(name, true).Value ?? false;
		}

		/// <summary>
		/// Gets a STRING value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the STRING column where NAME = <i>name</i>.</returns>
		public static string GetGlobalParamString(string name)
		{
			return new GlobalParameterString(name).Value ?? "";
		}

		/// <summary>
		/// Gets a mandatory STRING value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the STRING column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static string GetMandatoryGlobalParamString(string name) 
		{
			return new GlobalParameterString(name, true).Value ?? "";
		}

		/// <summary>
		/// Gets an AMOUNT value from the GLOBALPARAMETER table as an int.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the AMOUNT column where NAME = <i>name</i>.</returns>
		public static int GetGlobalParamAmount(string name)
		{
			return Convert.ToInt32(new GlobalParameterAmount(name).Value ?? 0.0);
		}

		/// <summary>
		/// Gets a mandatory AMOUNT value from the GLOBALPARAMETER table as an int.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the AMOUNT column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static int GetMandatoryGlobalParamAmount(string name) 
		{
			return Convert.ToInt32(new GlobalParameterAmount(name, true).Value ?? 0);
		}

		/// <summary>
		/// Gets an AMOUNT value from the GLOBALPARAMETER table as a double.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the AMOUNT column where NAME = <i>name</i>.</returns>
		public static double GetGlobalParamAmountAsDouble(string name)
		{
			return new GlobalParameterAmount(name).Value ?? 0.0;
		}

		/// <summary>
		/// Gets a mandatory AMOUNT value from the GLOBALPARAMETER table as a double.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the AMOUNT column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static double GetMandatoryGlobalParamAmountAsDouble(string name)
		{
			return new GlobalParameterAmount(name, true).Value ?? 0.0;
		}

		/// <summary>
		/// Gets a MAXIMUMAMOUNT value from the GLOBALPARAMETER table as an int.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the MAXIMUMAMOUNT column where NAME = <i>name</i>.</returns>
		public static int GetGlobalParamMaximumAmount(string name)
		{
			return Convert.ToInt32(new GlobalParameterMaximumAmount(name).Value ?? 0.0);
		}

		/// <summary>
		/// Gets a mandatory MAXIMUMAMOUNT value from the GLOBALPARAMETER table as an int.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the MAXIMUMAMOUNT column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static int GetMandatoryGlobalParamMaximumAmount(string name)
		{
			return Convert.ToInt32(new GlobalParameterMaximumAmount(name, true).Value ?? 0.0);
		}

		/// <summary>
		/// Gets a MAXIMUMAMOUNT value from the GLOBALPARAMETER table as a double.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the MAXIMUMAMOUNT column where NAME = <i>name</i>.</returns>
		public static double GetGlobalParamMaximumAmountAsDouble(string name)
		{
			return new GlobalParameterMaximumAmount(name).Value ?? 0.0;
		}

		/// <summary>
		/// Gets a mandatory MAXIMUMAMOUNT value from the GLOBALPARAMETER table as a double.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the MAXIMUMAMOUNT column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static double GetMandatoryGlobalParamMaximumAmountAsDouble(string name)
		{
			return new GlobalParameterMaximumAmount(name, true).Value ?? 0.0;
		}

		/// <summary>
		/// Gets a PERCENTAGE value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the PERCENTAGE column where NAME = <i>name</i>.</returns>
		public static double GetGlobalParamPercentage(string name)
		{
			return new GlobalParameterPercentage(name).Value ?? 0.0;
		}

		/// <summary>
		/// Gets a mandatory PERCENTAGE value from the GLOBALPARAMETER table.
		/// </summary>
		/// <param name="name">The name of the global parameter.</param>
		/// <returns>The value of the PERCENTAGE column where NAME = <i>name</i>.</returns>
		/// <exception cref="ErrAssistException">
		/// The parameter does not exist on the GLOBALPARAMETER table. 
		/// ErrAssistException.OmigaError is set to OMIGAERROR.MissingParameter.
		/// </exception>
		public static double GetMandatoryGlobalParamPercentage(string name)
		{
			return new GlobalParameterPercentage(name, true).Value ?? 0.0;
		}

		/// <summary>
		/// Get parameters from the GLOBALBANDEDPARAMETER table as xml.
		/// </summary>
		/// <param name="name">The name of the GLOBALBANDEDPARAMETER.</param>
		/// <param name="parameterType">The type of the GLOBALBANDEDPARAMETER (one of "STRING", "AMOUNT", "PERCENTAGE", "MAXIMUM", or "BOOLEAN").</param>
		/// <param name="xmlResponseNode">The node in which to return the xml response.</param>
		public static void GetAllGlobalBandedParamValuesAsXml(string name, string parameterType, XmlNode xmlResponseNode)
		{
			XmlDocument xmlDocument = new XmlDocument();
			// Create BOGUS root node.
			XmlElement xmlElement = xmlDocument.CreateElement("BOGUS");
			XmlNode xmlRootNode = xmlDocument.AppendChild(xmlElement);
			xmlElement = xmlDocument.CreateElement("GLOBALBANDEDPARAM");
			xmlElement.SetAttribute("ENTITYTYPE", "LOGICAL");
			xmlElement.SetAttribute("DATASRCE", "OM_GLOBALBANDEDPARAM");
			XmlNode xmlSchemaNode = xmlRootNode.AppendChild(xmlElement);

			xmlElement = xmlDocument.CreateElement("NAME");
			xmlElement.SetAttribute("DATATYPE", "STRING");
			xmlElement.SetAttribute("LENGTH", "30");
			xmlSchemaNode.AppendChild(xmlElement);

			switch (parameterType.ToUpper())
			{
				case "STRING":
					xmlElement = xmlDocument.CreateElement("STRING");
					xmlElement.SetAttribute("DATATYPE", "STRING");
					xmlElement.SetAttribute("LENGTH", "30");
					xmlSchemaNode.AppendChild(xmlElement);
					break;
				case "AMOUNT":
					xmlElement = xmlDocument.CreateElement("AMOUNT");
					xmlElement.SetAttribute("DATATYPE", "DOUBLE");
					xmlSchemaNode.AppendChild(xmlElement);
					break;
				case "PERCENTAGE":
					xmlElement = xmlDocument.CreateElement("PERCENTAGE");
					xmlElement.SetAttribute("DATATYPE", "DOUBLE");
					xmlSchemaNode.AppendChild(xmlElement);
					break;
				case "MAXIMUM":
					xmlElement = xmlDocument.CreateElement("MAXIMUM");
					xmlElement.SetAttribute("DATATYPE", "DOUBLE");
					xmlSchemaNode.AppendChild(xmlElement);
					break;
				case "BOOLEAN":
					xmlElement = xmlDocument.CreateElement("BOOLEAN");
					xmlElement.SetAttribute("DATATYPE", "BOOLEAN");
					xmlSchemaNode.AppendChild(xmlElement);
					break; 
			}

			// Create REQUEST details.
			xmlElement = xmlDocument.CreateElement("GLOBALPARAM");
			xmlElement.SetAttribute("NAME", name);
			XmlNode xmlRequestNode = xmlRootNode.AppendChild(xmlElement);
			if (xmlResponseNode == null)
			{
				xmlResponseNode = xmlDocument.CreateElement("RESPONSE");
			}

			AdoAssist.GetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlResponseNode);
		}
	}
}

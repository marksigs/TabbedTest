/*
--------------------------------------------------------------------------------------------
Workfile:			SystemDatesDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Contains methods that perform date manipulation with regard to working and
					non-working days in relation to nonworking days specified at country and
					distribution channel level.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
DJP		17/01/01	Created.
DJP		23/01/01	Validate date passed into CheckNonWorkingOccurence.
LD		11/06/01	SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
CL		07/05/02	SYS4510  Modifications to class initialise
CL		10/05/02	SYS4510  Remove class initialize & class terminate
SDS		25/08/04	BBG369   This AQR is not BBG specific. There is in-consistency in evaluating
					the weekday in VB and SQL Server. Hence the fix
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		25/07/2007	First .Net version. Ported from SystemDatesDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.SystemDatesDO")]
	[Guid("B0D9062A-6487-45D4-99CD-C2C0245F6E14")]
	[Transaction(TransactionOption.Supported)]
	public class SystemDatesDO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:  Returns whether or not the date passed in is a non-working day or not.
		// pass:		 xmlRequestElement  xml Request data stream
		// return:	   xml Response Node
		// ------------------------------------------------------------------------------------------
		public XmlNode CheckNonWorkingOccurence(XmlElement xmlRequestElement) 
		{
			XmlNode xmlResponseNode = null;

			const string recurringValidation = "R";
			const string nonWorkingDayGroup = "NonWorkingDayType";

			try 
			{
				// HasValue the correct keys have been passed in
				string dateToCheckText = XmlAssist.GetMandatoryNodeText(xmlRequestElement, ".//DATE");
				// Validate Date
				DateTime dateToCheck = new DateTime();
				if (!DateTime.TryParse(dateToCheckText, out dateToCheck))
				{
					throw new InvalidParameterException("DATE must be a validate date or date/time");
				}
				string channelIdText = XmlAssist.GetMandatoryNodeText(xmlRequestElement, ".//CHANNELID");
				string cmdText = "SELECT BH.BANKHOLIDAYDATE " + 
					"FROM " + 
					"BANKHOLIDAY BH, DISTRIBUTIONCHANNEL DC " + 
					"WHERE " + 
					"DC.CHANNELID = @dcChannelId AND " + 
					"BH.COUNTRYNUMBER= DC.COUNTRYNUMBER AND " + 
					"BH.BANKHOLIDAYDATE = @bankHolidayDate " + 
					"UNION ";
				// SDS  BBG369  25/08/2004___START
				cmdText += "SELECT NWD.NONWORKINGDAYDATE " + 
					"FROM " + 
					"NONWORKINGDAY NWD " + 
					"WHERE " + 
					"NWD.CHANNELID = @nwdChannelId1 AND " + 
					"NWD.NONWORKINGDAYDATE = @nwdDate1 UNION ";
				cmdText += "SELECT NWD.NONWORKINGDAYDATE " + 
					"FROM " + 
					"NONWORKINGDAY NWD " + 
					"WHERE " + 
					"NWD.CHANNELID = @nwdChannelId2 AND " + 
					"NWD.NONWORKINGDAYDATE = @nwdDate2 AND " + 
					SqlAssist.GetColumnDayOfWeek("NWD.NONWORKINGDAYDATE") + " = " + SqlAssist.GetColumnDayOfWeek("'" + dateToCheck + "'");
				// SqlAssist.GetColumnDayOfWeek("NWD.NONWORKINGDAYDATE") & " = ? "
				// SDS  BBG369  25/08/2004___END
				// Set up the parameters to be substituted into the SQL
				string isNotWorkingDayText = "0";
				using (SqlCommand sqlCommand = new SqlCommand())
				{
					sqlCommand.Parameters.Add("dcChannelId", SqlDbType.NVarChar, channelIdText.Length).Value = channelIdText;
					sqlCommand.Parameters.Add("bankHolidayDate", SqlDbType.DateTime).Value = dateToCheck;
					sqlCommand.Parameters.Add("nwdChannelId1", SqlDbType.NVarChar, channelIdText.Length).Value = channelIdText;
					sqlCommand.Parameters.Add("nwdDate1", SqlDbType.DateTime).Value = dateToCheck;
					sqlCommand.Parameters.Add("nwdChannelId2", SqlDbType.NVarChar, channelIdText.Length).Value = channelIdText;
					sqlCommand.Parameters.Add("nwdDate2", SqlDbType.DateTime).Value = dateToCheck;

					// ''''''''''''' Get Combo ValueID's ''''''''''''''''''
					ReadOnlyCollection<int> valueIds = ComboAssist.GetValueIdsForValidationType(nonWorkingDayGroup, recurringValidation);
					// Need to add the comboid values returned to the SQL - we only want recurring dates to be checked if the
					// NonWorkingType field has a valueid who's validation type is "R"
					string sqlValueText = "";
					if (valueIds.Count > 0)
					{
						sqlValueText = " AND NWD.NONWORKINGTYPE IN (";
						for (int valueIdIndex = 0; valueIdIndex < valueIds.Count; valueIdIndex++)
						{
							if (valueIdIndex > 0)
							{
								sqlValueText += ", ";
							}
							string parameterName = "@nonWorkingType" + valueIdIndex.ToString();
							sqlValueText += parameterName;
							sqlCommand.Parameters.Add(parameterName, SqlDbType.Int).Value = valueIds[valueIdIndex];
						}
						sqlValueText += ") ";
					}
					else
					{
						throw new MissingPrimaryTagException("Value ID tag(s) not found");
					}
					cmdText += sqlValueText;

					sqlCommand.CommandText = cmdText;
					sqlCommand.CommandType = CommandType.Text;

					using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
					{
						sqlCommand.Connection = sqlConnection;
						int retry = 0;
						while (sqlConnection.State != ConnectionState.Open && retry < AdoAssist.GetDbRetries())
						{
							try
							{
								sqlConnection.Open();
							}
							catch (Exception)
							{
							}
							retry++;
						}

						if (sqlConnection.State == ConnectionState.Open)
						{
							isNotWorkingDayText = sqlCommand.ExecuteScalar() != null ? "1" : "0";
						}
						else
						{
							throw new UnableToConnectException();
						}
					}
				}

				XmlDocument xmlOut = new XmlDocument();
				xmlResponseNode = (XmlNode)xmlOut.CreateElement("NONWORKINGIND");
				xmlResponseNode.InnerText = isNotWorkingDayText;
				xmlOut.AppendChild(xmlResponseNode);
				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return xmlResponseNode;
		}
	}
}

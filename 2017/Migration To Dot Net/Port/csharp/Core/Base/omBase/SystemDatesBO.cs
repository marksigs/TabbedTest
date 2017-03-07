/*
--------------------------------------------------------------------------------------------
Workfile:			SystemDatesBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Contains methods that perform date manipulation with regard to working and
					non-working days in relation to nonworking days specified at country and
					distribution channel level.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
DJP		17/01/01	Created
DJP		23/01/01	Added ValidateTimePart to validate days and hours
DJP		29/01/01	In FindWorkingDay move 0 hours onto next day.
CL		07/05/02	SYS4510  Modifications to class initialise
CL		10/05/02	SYS4510  Remove class initialize & class terminate
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		30/07/2007	First .Net version. Ported from SystemDatesBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
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
	[ProgId("omBase.SystemDatesBO")]
	[Guid("7B34AC80-3148-499D-A5D7-EAEC78B2B7F8")]
	[Transaction(TransactionOption.Supported)]
	public class SystemDatesBO : ServicedComponent
	{
		private const string _rootNodeName = "SYSTEMDATE";

		// header ----------------------------------------------------------------------------------
		// description:  See ISystemDatesBO_FindWorkingDay
		// pass:		 vstrXmlRequest  xml Request data stream
		// return:					   xml Response data stream containing results of operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// ------------------------------------------------------------------------------------------
		public string FindWorkingDay(string request) 
		{
			string response = "";

			try
			{
				// Create default response block
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlDocument xmlIn = XmlAssist.Load(request);

				// Delegate to XmlDocument based method and attach returned data to our response
				XmlNode xmlTempResponseNode = FindWorkingDay(xmlIn.DocumentElement);
				ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElement, true);
				XmlAssist.AttachResponseData(xmlResponseElement, (XmlElement)xmlTempResponseNode);
				response = xmlResponseElement.OuterXml;
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}

		// header ----------------------------------------------------------------------------------
		// description:  Returns the next or previous working day, or if variation is specified, finds the n'th
		// non working day either before or after the specified date.
		// business object
		// pass:		 xmlRequest  xml Request data stream containing data to be persisted
		// DATE: The start date and optionally, time
		// DIRECTION: either + or - for future or past
		// HOURS: number of hours to move forwards or backwards
		// DAYS: number of days to move forwards or backwards
		// CHANNELID: distribution channel for which nonworking days should be found
		// return:	   xml Response Node
		// ------------------------------------------------------------------------------------------
		public XmlNode FindWorkingDay(XmlElement xmlRequest) 
		{
			XmlNode xmlResponse = null;

			try
			{
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");

				xmlResponseElement = (XmlElement)xmlResponseElement.AppendChild(xmlOut.CreateElement(_rootNodeName));
				
				XmlNode xmlRequestNode = null;
				if (xmlRequest.Name == _rootNodeName)
				{
					xmlRequestNode = xmlRequest;
				}
				else
				{
					xmlRequestNode = xmlRequest.GetElementsByTagName(_rootNodeName)[0];
				}

				if (xmlRequestNode == null)
				{
					throw new MissingPrimaryTagException(_rootNodeName + " tag not found");
				}

				// Read input XML Values
				string inputDateText = XmlAssist.GetMandatoryNodeText(xmlRequestNode, ".//DATE");
				string directionText = XmlAssist.GetMandatoryNodeText(xmlRequestNode, ".//DIRECTION");
				// Both the following are non-mandatory
				string daysText = XmlAssist.GetNodeText(xmlRequestNode, ".//DAYS");
				string hoursText = XmlAssist.GetNodeText(xmlRequestNode, ".//HOURS");
				XmlNode xmlChannelId = XmlAssist.GetMandatoryNode(xmlRequestNode, ".//CHANNELID").CloneNode(true);
				// Validation direction
				if (directionText != "+" && directionText != "-")
				{
					throw new InvalidParameterException("Direction must be '+' or '-'");
				}
				// Validate Hours
				bool limit = false;
				if (daysText.Length > 0)
				{
					limit = true;
					if (!ValidateTimePart(daysText))
					{
						throw new InvalidParameterException("DAYS must be a positive whole number");
					}
				}
				// Validate Minutes
				if (hoursText.Length > 0)
				{
					if (!ValidateTimePart(hoursText))
					{
						throw new InvalidParameterException("HOURS must be a positive whole number");
					}
				}
				// Validate Date
				DateTime inputDate;
				if (!DateTime.TryParse(inputDateText, out inputDate))
				{
					throw new InvalidParameterException("DATE must be a validate date or date/time");
				}
				// Create request block for SystemDatesDO
				XmlDocument xmlDateOccurence = new XmlDocument();
				XmlElement xmlDateElem = (XmlElement)xmlDateOccurence.AppendChild(xmlDateOccurence.CreateElement(_rootNodeName));
				xmlDateElem.InnerXml = xmlChannelId.OuterXml;
				xmlDateElem = (XmlElement)xmlDateElem.AppendChild(xmlDateOccurence.CreateElement("DATE"));
				int dayLimit = 0;
				// Decide if we're going forwards or backwards
				if (directionText == "-")
				{
					dayLimit = -1;
				}
				else
				{
					dayLimit = 1;
				}
				DateTime tempInputDate = inputDate;
				// Add the hours first
				if (hoursText.Length > 0)
				{
					if (IsNumeric(hoursText))
					{
						if (directionText == "-")
						{
							inputDate = inputDate.AddHours(-(Convert.ToInt32(hoursText)));
						}
						else
						{
							inputDate = inputDate.AddHours(Convert.ToInt32(hoursText));
						}
					}
					else
					{
						throw new InvalidParameterException("Hours must be a single integer value");
					}
				}
				// If we've not jumped over a day boundry, start on the next day
				int dayCount = 0;
				if (tempInputDate.Day == inputDate.Day)
				{
					dayCount = 1;
					// Add or subtract a day to the current date for our start point. If adding the hour onto
					// the date did this, we won't come in here so won't add the date
					inputDate = inputDate.AddDays((double)dayLimit);
				}

				// Need to loop from the data passed in to the specified number of days and hours, or if days and hours
				// are not specified, to the next non working day.
				using (SystemDatesDO systemDatesDO = new SystemDatesDO())
				{
					bool found = false;
					while (!found)
					{
						// See if this day is a non working day or not
						xmlDateElem.InnerText = inputDate.ToString("dd/MM/yyyy");
						// Is the day a working day?
						XmlNode xmlWorkingDayResp = null;
						xmlWorkingDayResp = systemDatesDO.CheckNonWorkingOccurence(xmlDateOccurence.DocumentElement);
						string nonWorkingDayText = XmlAssist.GetMandatoryNodeText(xmlWorkingDayResp, "//NONWORKINGIND");
						// Decide if we've hit a non-working day or not. If we have and there isn't a day count specified we've
						// found the right day. If not, we need to count up until the specified number of days has been hit.
						if (nonWorkingDayText == "0")
						{
							if (limit)
							{
								if (dayCount == Convert.ToInt32(daysText))
								{
									found = true;
								}
								else
								{
									dayCount++;
								}
							}
							else
							{
								found = true;
							}
						}
						if (!found)
						{
							inputDate = inputDate.AddDays((double)dayLimit);
						}
					}
				}
				XmlNode xmlDataNode = xmlResponseElement.AppendChild(xmlOut.CreateElement("DATE"));
				xmlDataNode.InnerText = inputDate.ToString("dd/MM/yyyy");

				xmlResponse = xmlOut.DocumentElement;
			}
			catch (Exception exception)
			{
				xmlResponse = new ErrAssistException(exception).CreateErrorResponseNode();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return xmlResponse;
		}

		private bool IsNumeric(string text)
		{
			bool isNumeric = false;

			if (text != null && text.Length > 0)
			{
				isNumeric = true;
				foreach (char ch in text)
				{
					if (!Char.IsDigit(ch))
					{
						isNumeric = false;
						break;
					}
				}
			}

			return isNumeric;
		}

		// header ----------------------------------------------------------------------------------
		// description:  ValidateTimePart - Validates that the number passed in is a) there at all,
		// b) a whole number and c) positive
		// pass:		 timePart	   Number to be validated
		// return:	   true if valid, false if not
		// ------------------------------------------------------------------------------------------
		private bool ValidateTimePart(string timePartText) 
		{
			bool isValid = false;

			if (IsNumeric(timePartText))
			{
				double timePart = Convert.ToDouble(timePartText);
				if (timePart >= 0.0 && Math.Floor(timePart) - timePart == 0.0)
				{
					isValid = true;
				}
			}

			return isValid;
		}

		// header ----------------------------------------------------------------------------------
		// description:  Returns whether or not the date passed in is a non-working day or not.
		// pass:		 xmlRequest  xml Request data stream
		// return:	   xml Response Node
		// ------------------------------------------------------------------------------------------
		public XmlNode CheckNonWorkingOccurence(XmlElement xmlRequest) 
		{
			XmlNode xmlResponse = null;
			try 
			{
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				xmlResponseElement = (XmlElement)xmlResponseElement.AppendChild(xmlOut.CreateElement(_rootNodeName));
				XmlNode xmlRequestNode = null;
				if (xmlRequest.Name == _rootNodeName)
				{
					xmlRequestNode = (XmlNode)xmlRequest;
				}
				else
				{
					xmlRequestNode = (XmlNode)xmlRequest.GetElementsByTagName(_rootNodeName)[0];
				}

				if (xmlRequestNode == null)
				{
					throw new MissingPrimaryTagException(_rootNodeName + " tag not found");
				}

				// Call the DO...
				XmlDocument xmlNonWorking = new XmlDocument();
				XmlNode xmlNonWorkingelem = xmlNonWorking.AppendChild(xmlNonWorking.ImportNode(xmlRequestNode, true));
				XmlNode xmlDataNode = null;
				using (SystemDatesDO systemDatesDO = new SystemDatesDO())
				{
					xmlDataNode = systemDatesDO.CheckNonWorkingOccurence(xmlNonWorking.DocumentElement);
				}
				xmlResponseElement.AppendChild(xmlOut.ImportNode(xmlDataNode, true));

				xmlResponse = xmlOut.DocumentElement;
			}
			catch (Exception exception)
			{
				xmlResponse = new ErrAssistException(exception).CreateErrorResponseNode();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return xmlResponse;
		}

		// header ----------------------------------------------------------------------------------
		// description:  See ISystemDatesBO_CheckNonWorkingOccurence
		// pass:		 vstrXmlRequest  xml Request data stream containing data to be persisted
		// return:					   xml Response data stream containing results of operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// ------------------------------------------------------------------------------------------
		public string CheckNonWorkingOccurence(string request) 
		{
			string response = "";
			try 
			{
				// Create default response block
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlDocument xmlIn = XmlAssist.Load(request);

				// Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
				XmlNode xmlTempResponseNode = CheckNonWorkingOccurence(xmlIn.DocumentElement);
				ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElement, true);
				XmlAssist.AttachResponseData(xmlResponseElement, (XmlElement)xmlTempResponseNode);
				response = xmlOut.OuterXml;
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}
	}
}

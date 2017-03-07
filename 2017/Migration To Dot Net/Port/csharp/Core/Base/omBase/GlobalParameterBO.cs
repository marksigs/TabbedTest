/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalParameterBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Global parameters business object
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		21/07/99	Created
MV		29/11/00	Created New Function FindCurrentParameterList
ASm		12/01/01	SYS1817: oeNoDataForParameter enum removed and replaced with specific error number
PSC		12/03/01	SYS2024: Added isTaskManager()
DRC		3/10/01	 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
------------------------------------------------------------------------------------------
BMids History:

Prog	Date		Description
MDC		10/04/2003  BM0493 Added GetCurrentParameterListEx
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
------------------------------------------------------------------------------------------
Core History

Prog   Date		Description
AS	 22/03/2006  CORE175: BMIDS609 GlobalParameterBO Functions not getting the object context
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		26/07/2007	First .Net version. Ported from GlobalParameterBO.cls.
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
using Microsoft.Win32;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.GlobalParameterBO")]
	[Guid("D69B7828-7EE3-4D9A-BAFF-6DF7F952AD50")]
	[Transaction(TransactionOption.Supported)]
	public class GlobalParameterBO : ServicedComponent
	{
		private const string _applicationName = "Omiga4";
		private const string _registrySection = "System Configuration";

		// header ----------------------------------------------------------------------------------
		// description:  Gets the current values for the parameter passed in
		// pass:		 name   Parameter for which current data is required
		// return:	   GetCurrentParameter xml Response data stream containing results of
		// operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameter(string name) 
		{
			string response = "";

			try 
			{
				string data = "";
				using (GlobalParameterDO globalParameterDO = new GlobalParameterDO())
				{
					data = globalParameterDO.GetCurrentParameter(name);
				}

				// if we get here, everything has completed OK
				response = "<RESPONSE TYPE='SUCCESS'>" + data + "</RESPONSE>";
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
		// description:  Get the data for all instances of the persistant data associated with
		// this data object
		// pass:		 request  xml Request data stream containing data to which identifies
		// instance of the persistant data to be retrieved
		// 
		// return:	   FindList		 xml Response data stream containing results of operation
		// either: TYPE="SUCCESS" and xml representation of data
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string FindCurrentParameterList(string request) 
		{
			string response = "";

			try
			{
				// Create Response Object
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				// append child
				XmlNode xmlDataNode = xmlOut.AppendChild(xmlResponseElement);
				// set type attribute and value as suscess
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				// load Request into an XML Document
				XmlDocument xmlIn = XmlAssist.Load(request);
				// HasValue for GLobalParameter Node
				XmlNode xmlRequestNode = xmlIn.GetElementsByTagName("GLOBALPARAMETER")[0];
				if (xmlRequestNode == null)
				{
					// If null then raise error
					throw new MissingPrimaryTagException("GLOBALPARAMETER tag not found");
				}
				// store all name nodes
				XmlNodeList xmlNameList = xmlIn.SelectNodes(".//GLOBALPARAMETER/NAME");
				// Create an elemetn for response object
				XmlElement xmlParamList = xmlOut.CreateElement("GLOBALPARAMETERLIST");
				xmlResponseElement.AppendChild(xmlParamList);
				// loop through the name list node
				using (GlobalParameterDO globalParameterDO = new GlobalParameterDO())
				{
					foreach (XmlNode xmlChildNode in xmlNameList)
					{
						XmlElement xmlElement = (XmlElement)xmlChildNode;
						// If a parameter is not found then raise an cutomised error message
						string data = "";
						try
						{
							data = globalParameterDO.GetCurrentParameter(xmlElement.InnerText);
						}
						catch (RecordNotFoundException)
						{
							throw new OmigaErrorException(OMIGAERROR.MissingGlobalParameter, xmlElement.InnerText);
						}
						// If the parameter exists then
						if (data != null && data.Length > 0)
						{
							// Load into an XMlDoc
							XmlDocument xmlDataDoc = XmlAssist.Load(data);
							// append to the response object
							xmlParamList.AppendChild(xmlOut.ImportNode(xmlDataDoc.DocumentElement, true));
						}
					}
				}
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

		// BM0493 MDC 10/04/2003
		// header ----------------------------------------------------------------------------------
		// description:  Gets the current values for the parameter passed in
		// pass:		 name   Parameter for which current data is required
		// return:	   GetCurrentParameter xml Response data stream containing results of
		// operation
		// either: TYPE="SUCCESS"
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameterListEx(string request)
		{
			string response = "";

			try
			{
				string data = "";
				using (GlobalParameterDO globalParameterDO = new GlobalParameterDO())
				{
					data = globalParameterDO.GetCurrentParameterListEx(request);
				}

				// if we get here, everything has completed OK
				response = "<RESPONSE TYPE='SUCCESS'>" + data + "</RESPONSE>";
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

		public bool IsTaskManager() 
		{
			bool isTaskManager = false;

			try
			{
				isTaskManager = IsSystemConfiguration("Task Manager");
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return isTaskManager;
		}

		public string IsMultipleLender()
		{
			string response = "";

			try
			{
				bool isMultipleLender = IsSystemConfiguration("Multiple Lender");
				response = 
					"<RESPONSE><MULTIPLELENDER>" + 
					(isMultipleLender ? "1" : "0") + 
					"</MULTIPLELENDER></RESPONSE>";
			}
			finally
			{
				ContextUtility.SetComplete();
			}

			return response;
		}

		private bool IsSystemConfiguration(string valueName)
		{
			bool isSystemConfiguration = false;

			string registryKeyName = "SOFTWARE\\" + _applicationName + "\\" + _registrySection;
			using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(registryKeyName))
			{
				string item = (string)registryKey.GetValue(valueName);
				isSystemConfiguration = item == "1";
			}

			return isSystemConfiguration;
		}

		protected override bool CanBePooled()
		{
			return true;
		}

		#region Not implemented

		/*
		// These methods were defined in the VB6 code as raising a "not implemented" error.
		// Comment out here so it is impossible to call them.
		public string Create(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Update(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Delete(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}
		*/

		#endregion
	}

}

/*
--------------------------------------------------------------------------------------------
Workfile:			DOAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Code for omiga4 Business Object - ComboBO
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		30/06/99	Created
DRC		03/10/01	SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
SR		13/04/04	BBG132 - New method GetRegSaleRelatedComboList
TK		22/11/04	BBG1821 - Performance related fixes.
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		24/07/2007	First .Net version. Ported from ComboBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.ComboBO")]
	[Guid("1DD9FA64-A34A-409C-97E8-B56F6CA5E96A")]
	[Transaction(TransactionOption.Supported)]
	public class ComboBO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:  Gets the data relating to a valueid of a combo list
		// pass:		 request  xml Request data stream containing data to which identifies
		// instance of the persistant data to be retrieved
		// 
		// return:	   GetComboValue   xml Response data stream containing results of operation
		// ither: TYPE="SUCCESS" and xml representation of data
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetComboValue(string request)
		{
			string response = null;

			try
			{
				XmlDocument xmlDocumentOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlDocumentOut.CreateElement("RESPONSE");
				xmlDocumentOut.AppendChild(xmlResponseElement);

				using (ComboDO comboDO = new ComboDO())
				{
					xmlResponseElement.InnerXml = comboDO.GetComboValue(request);
				}

				// if we get here, everything has completed OK
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				response = xmlDocumentOut.OuterXml;
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
		// description:  Gets all data relating to a particular combo group
		// pass:		 request  xml Request data stream containing data to which identifies
		// instance of the persistant data to be retrieved
		// 
		// return:	   GetComboList	xml Response data stream containing results of operation
		// either: TYPE="SUCCESS" and xml representation of data
		// or: TYPE="SYSERR" and <ERROR> element
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetComboList(string request)
		{
			string response = "";

			try
			{
				XmlDocument xmlDocumentOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlDocumentOut.CreateElement("RESPONSE");
				xmlDocumentOut.AppendChild(xmlResponseElement);

				using (ComboDO comboDO = new ComboDO())
				{
					xmlResponseElement.InnerXml = comboDO.GetComboList(request);
				}

				// if we get here, everything has completed OK
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");

				response = xmlDocumentOut.OuterXml;
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

		// AS 24/07/2007 Not implemented as specific to BBG MX.
		public string GetRegSaleRelatedComboList(string request)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.Assists.NotImplementedException();
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

		public string GetData(string request)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string FindList(string request)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string DeleteAll(string request)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}
		*/

		#endregion
	}

}

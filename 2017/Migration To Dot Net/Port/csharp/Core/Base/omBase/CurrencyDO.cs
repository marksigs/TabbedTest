/*
--------------------------------------------------------------------------------------------
Workfile:			CurrencyDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Collects all Currency data for use with the Currency Calculator
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PF		09/04/2001	Created
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		24/07/2007	First .Net version. Ported from CurrencyDO.cls.
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
	[ProgId("omBase.CurrencyDO")]
	[Guid("5CC43279-541E-4F2F-9B6B-AFDE958F6A23")]
	[Transaction(TransactionOption.Supported)]
	public class CurrencyDO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:
		// Get the data for all instances of the persistant data associated with
		// this data object for the values supplied
		// pass:
		// vxmlTableElement  xml element containing the request
		// return:				xml node containing retrieved data
		// ------------------------------------------------------------------------------------------
		public XmlNode FindList() 
		{
			XmlNode xmlResponseNode = null;
			try 
			{
				XmlDocument xmlClassDefDocument = omBaseClassDef.LoadCurrencyData();
				XmlElement xmlElement = xmlClassDefDocument.CreateElement("CURRENCY");
				xmlResponseNode = DOAssist.FindListEx(xmlElement, xmlClassDefDocument);
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

/*
--------------------------------------------------------------------------------------------
Workfile:			DOAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Code for omiga4 Data Object - ComboDO
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		01/07/99	Created.
RF		20/09/99	Added IsItemInValidation.
RF	  06/10/99	DOs shouldn't call SetAbort except on system error.
RF	  17/12/99	Added GetFirstComboValueId.
MS	  21/07/00	Moved methods from MP as part of performance rework.
MC	  07/08/2000  SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
PSC	 11/08/2000  SYS1430 Back out SYS1409
LD	  07/11/2000  Explicitly close recordsets
AS	  20/11/2000  SYS1670: Explicitly close recordsets
DM	  17/05/01	SYS2316
DM	  17/05/01	SYS2316 Move outer joins
DM	  31/07/01	SYS2539 Remove outer join code from Oracle side of Run Time switch.
					This is from GetComboValue, GetComboList.
AS	  13/11/03	CORE1 Removed GENERIC_SQL.
--------------------------------------------------------------------------------------------
BMids History:

Prog	Date		Description
MDC		12/09/2002  BMIDS00336 Added GetComboValueIdFromValueName and GetFirstComboValueIdFromValueName
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
SR		13/04/04 BBG132 - New method GetRegSaleRelatedComboList
MV		07/07/2004  BBG123 WP16 Add function to ascertain if an application is a TofE
TK		22/11/04 BBG1821 - Performance related fixes.
MS		21/07/00	performance rework.
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		23/07/2007	First .Net version. Ported from ComboDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.ComboDO")]
	[Guid("72240773-634F-46C8-97E9-D7D661940D9D")]
	[Transaction(TransactionOption.Supported)]
	public class ComboDO : ServicedComponent
	{
		// header ----------------------------------------------------------------------------------
		// description:  Get the combo data for all instances of the persistant data associated with
		//               this data object matching the request criteria
		//
		// pass:         vstrXMLRequest  xml Request data stream containing list which identifies an
		//                               instance of the persistant data to be retrieved. Received in
		//                               the format:
		// i.e.
		//   <LIST><LISTNAME>Title</LISTNAME></LIST>
		//
		// return:       GetComboList    string containing XML data stream representation of
		//                               data retrieved
		// i.e.
		// <LIST>
		//    <LISTNAME NAME=Title>
		//       <LISTENTRY>
		//           <VALUEID>1</VALUEID>
		//           <VALUENAME>Mr</VALUENAME>
		//           <VALIDATIONTYPELIST>
		//               <VALIDATIONTYPE>M</VALIDATIONTYPE>
		//               <VALIDATIONTYPE>O</VALIDATIONTYPE>
		//               ...
		//           </VALIDATIONTYPELIST>
		//       </LISTENTRY>
		//       ...
		//    </LISTNAME>
		//    <LISTNAME>
		//    ...
		//    </LISTNAME>
		// </LIST>
		// Raise Errors:     if record not found, raise omiga4RecordNotFound
		//                   if no combo group specified raise omiga4err108
		public string GetComboList(string request)
		{
			// maintenance -----------------------------------------------------------------------------
			// Date		  Developer   Comments
			// 30/06/1999	IK		  Initial Creation
			// 02/07/1999	AS		  Initial GetComboList implementation
			// 07/09/1999	AS		  Changed GetComboList to read multiple combo groups in one call
			// ------------------------------------------------------------------------------------------

			string response = "";

			try
			{
				XmlDocument xmlDocumentIn = XmlAssist.Load(request);
				XmlElement xmlElement = (XmlElement)xmlDocumentIn.GetElementsByTagName("LIST")[0];

				if (xmlElement == null)
				{
					throw new MissingPrimaryTagException("LIST tag not found");
				}

				string groupName = XmlAssist.GetTagValue(xmlElement, "LISTNAME");
				if (groupName == null || groupName.Length == 0)
				{
					throw new OmigaErrorException(OMIGAERROR.MissingComboGroup);
				}

				List<ComboCollection> comboGroups = ComboAssist.GetComboGroups(GetGroupNames(xmlDocumentIn));

				XmlDocument xmlDocumentOut = new XmlDocument();
				XmlNode xmlNode = xmlDocumentOut.CreateElement("LIST");
				XmlNode xmlListNode = xmlDocumentOut.AppendChild(xmlNode);
				foreach (ComboCollection comboGroup in comboGroups)
				{
					XmlElement xmlListNameElement = xmlDocumentOut.CreateElement("LISTNAME");
					xmlListNode.AppendChild(xmlListNameElement);
					xmlListNameElement.SetAttribute("NAME", comboGroup.GroupName);

					foreach (KeyValuePair<int, List<ComboValue>> value in comboGroup)
					{
						XmlElement xmlListEntryElement = xmlDocumentOut.CreateElement("LISTENTRY");
						xmlListNameElement.AppendChild(xmlListEntryElement);
						GetListEntry(xmlListEntryElement, comboGroup, value.Key);
					}
				}

				response = xmlDocumentOut.OuterXml;

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return response;
		}

		private static string GetGroupNames(XmlDocument xmlDocumentIn)
		{
			string groupNames = "";

			XmlNodeList xmlListNames = xmlDocumentIn.GetElementsByTagName("LISTNAME");
			foreach (XmlNode xmlListNode in xmlListNames)
			{
				if (groupNames.Length > 0)
				{
					groupNames += ",";
				}
				groupNames += xmlListNode.InnerText;
			}

			return groupNames;
		}

		private static XmlNode GetListEntry(XmlNode xmlListEntryNode, ComboCollection comboGroup, int valueId)
		{
			XmlNode xmlNode = xmlListEntryNode.OwnerDocument.CreateElement("GROUPNAME");
			xmlListEntryNode.AppendChild(xmlNode);
			xmlNode.InnerText = comboGroup.GroupName;
			xmlNode = xmlListEntryNode.OwnerDocument.CreateElement("VALUEID");
			xmlListEntryNode.AppendChild(xmlNode);
			xmlNode.InnerText = valueId.ToString();

			XmlNode xmlValueNameNode = null;
			XmlNode xmlValidationTypeListNode = xmlListEntryNode.OwnerDocument.CreateElement("VALIDATIONTYPELIST");
			ReadOnlyCollection<ComboValue> comboValues = comboGroup.GetComboValuesByValueId(valueId);
			if (comboValues.Count == 0)
			{
				// Ensure matches xml from VB6 version.
				xmlListEntryNode.AppendChild(xmlValidationTypeListNode);
			}
			else
			{
				foreach (ComboValue comboValue in comboValues)
				{
					if (xmlValueNameNode == null)
					{
						xmlValueNameNode = xmlListEntryNode.OwnerDocument.CreateElement("VALUENAME");
						xmlListEntryNode.AppendChild(xmlValueNameNode);
						xmlListEntryNode.AppendChild(xmlValidationTypeListNode);
						xmlValueNameNode.InnerText = comboValue.ValueName;
					}

					if (comboValue.ValidationType != null && comboValue.ValidationType.Length > 0)
					{
						xmlNode = xmlListEntryNode.OwnerDocument.CreateElement("VALIDATIONTYPE");
						xmlValidationTypeListNode.AppendChild(xmlNode);
						xmlNode.InnerText = comboValue.ValidationType;
					}
				}
			}

			return xmlListEntryNode;
		}

		// header ----------------------------------------------------------------------------------
		// description:  Get the data for a single instance of the persistant data associated with
		// this data object
		// pass:
		// request  xml Request data stream containing list which identifies an
		// instance of the persistant data to be retrieved. Received in
		// the format:
		// <LIST>
		// <GROUPNAME>Title</GROUPNAME>
		// <VALUEID>1</VALUEID>
		// </LIST>
		// return:
		// GetComboValue   string containing XML data stream representation of
		// data retrieved in the format:
		// <LISTENTRY>
		// <VALUEID>1</VALUEID>
		// <VALUENAME>Mr</VALUENAME>
		// <VALIDATIONTYPELIST>
		// <VALIDATIONTYPE>M</VALIDATIONTYPE>
		// ...
		// <VALIDATIONTYPE>O</VALIDATIONTYPE>
		// <VALIDATIONTYPELIST>
		// </LISTENTRY>
		// Raise Errors:
		// if record not found, raise omiga4RecordNotFound
		// if no group name specified raise omiga4err108
		// if no valueID specified raise omiga4err109
		// maintenance -----------------------------------------------------------------------------
		// Date		  Developer   Comments
		// 07/09/1999	AS		  Changed GetComboValue due to GetComboList reading
		// multiple combo groups in one call
		// ------------------------------------------------------------------------------------------
		public string GetComboValue(string request)
		{
			string response = "";

			try
			{
				XmlDocument xmlDocumentIn = XmlAssist.Load(request);
				XmlElement xmlElement = (XmlElement)xmlDocumentIn.GetElementsByTagName("LIST")[0];

				if (xmlElement == null)
				{
					throw new MissingPrimaryTagException("LIST tag not found");
				}

				string groupName = XmlAssist.GetTagValue(xmlElement, "GROUPNAME");
				string valueIdText = XmlAssist.GetTagValue(xmlElement, "VALUEID");

				if (groupName == null || groupName.Length == 0)
				{
					throw new OmigaErrorException(OMIGAERROR.MissingComboGroup);
				}
				if (valueIdText == null || valueIdText.Length == 0)
				{
					throw new OmigaErrorException(OMIGAERROR.MissingComboValueId);
				}

				ComboCollection comboGroup = ComboAssist.GetComboGroup(groupName);
				int valueId = Convert.ToInt32(valueIdText);
				XmlDocument xmlDocumentOut = new XmlDocument();
				XmlElement xmlListEntryElement = xmlDocumentOut.CreateElement("LISTENTRY");
				xmlDocumentOut.AppendChild(xmlListEntryElement);
				GetListEntry(xmlListEntryElement, comboGroup, valueId);

				response = xmlDocumentOut.OuterXml;

				ContextUtility.SetComplete();
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception, TransactionAction.SetOnErrorType);
			}

			return response;
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Get the combo text (i.e. ValueName) for a given combo name and ValueId
		// pass:
		// groupName	   value for ComboValue.ComboName
		// valueIdText		 value for ComboValue.ValueId
		// return:
		// GetComboText		required value of ComboValue.ValueName
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public string GetComboText(string groupName, string valueIdText)
		{
			return ComboAssist.GetComboText(groupName, Convert.ToInt32(valueIdText));
		}

		public bool IsItemInValidation(string groupName, string valueIdText, string validationType)
		{
			return ComboAssist.IsValidationType(groupName, Convert.ToInt32(valueIdText), validationType);
		}

		public string GetFirstComboValueId(string groupName, string validationType)
		{
			return ComboAssist.GetFirstComboValueId(groupName, validationType).ToString();
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// pass:
		// groupName	   Combo group name to be searched
		// validationType  Validation type to search for in group
		// return:
		// Response Format:
		// <VALUEIDLIST>
		// <VALUEID></VALUEID>
		// <VALUEID></VALUEID>
		// ...
		// <VALUEID></VALUEID>
		// </VALUEIDLIST>
		// ------------------------------------------------------------------------------------------
		public string GetComboValueId(string groupName, string validationType)
		{
			return CreateValueIdList(ComboAssist.GetValueIdsForValidationType(groupName, validationType));
		}

		// BMIDS00336 MDC 12/09/2002
		// header ----------------------------------------------------------------------------------
		// description:
		// Get a single combo value id as a simple string rather than as an xml response.
		// pass:
		// return:
		// ------------------------------------------------------------------------------------------
		public string GetFirstComboValueIdFromValueName(string groupName, string valueName)
		{
			return ComboAssist.GetFirstComboTextId(groupName, valueName).ToString();
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// pass:
		// groupName	   Combo group name to be searched
		// valueName	   Text for the combo value name
		// return:
		// Response Format:
		// <VALUEIDLIST>
		// <VALUEID></VALUEID>
		// <VALUEID></VALUEID>
		// ...
		// <VALUEID></VALUEID>
		// </VALUEIDLIST>
		// ------------------------------------------------------------------------------------------
		public string GetComboValueIdFromValueName(string groupName, string valueName)
		{
			return CreateValueIdList(ComboAssist.GetValueIdsForValueName(groupName, valueName));
		}

		private static string CreateValueIdList(ReadOnlyCollection<int> valueIds)
		{
			XmlDocument xmlDocumentOut = new XmlDocument();
			XmlElement xmlValueIdListElement = xmlDocumentOut.CreateElement("VALUEIDLIST");
			xmlDocumentOut.AppendChild(xmlValueIdListElement);

			foreach (int valueId in valueIds)
			{
				XmlElement xmlValueIdElement = xmlDocumentOut.CreateElement("VALUEID");
				xmlValueIdListElement.AppendChild(xmlValueIdElement);
				xmlValueIdElement.InnerText = valueId.ToString();
			}

			return xmlDocumentOut.OuterXml;
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Get the combo text (i.e. ValueName) for a given combo name and ValueId
		// pass:
		// groupName	   value for ComboValue.ComboName
		// valueIdText		 value for ComboValue.ValueId
		// return:
		// The first validtion for the valueid
		// ------------------------------------------------------------------------------------------
		public string GetFirstComboValidation(string groupName, string valueIdText)
		{
			return ComboAssist.GetValidationTypeForValueID(groupName, Convert.ToInt32(valueIdText));
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Return the first combo value for an entry found to be a New Loan
		// pass:
		// 
		// return:	   New Loan Value
		// ------------------------------------------------------------------------------------------
		public string GetNewLoanValue()
		{
			return ComboAssist.GetFirstComboValueId("TypeOfMortgage", "N").ToString();
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// HasValue if applicationType corresponds to a New Loan.
		// pass:
		// applicationType
		// return:
		// IsNewLoan
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public bool IsNewLoan(string applicationType)
		{
			return IsItemInValidation("TypeOfMortgage", applicationType, "N");
		}

		// header ----------------------------------------------------------------------------------
		// description:
		//   HasValue if vstrTypeOfApplication corresponds to a Further Advance
		// pass:
		//   vstrTypeOfApplication
		// return:
		//   IsFurtherAdvance
		// Raise Errors:
		//------------------------------------------------------------------------------------------
		public bool IsFurtherAdvance(string applicationType)
		{
			return IsItemInValidation("TypeOfMortgage", applicationType, "F");
		}

		// BBG123 Start
		public bool IsTransferOfEquity(string applicationType)
		{
			return IsItemInValidation("TypeOfMortgage", applicationType, "T");
		}
		// BBG123 End

		// header ----------------------------------------------------------------------------------
		// description:
		// HasValue if applicationType corresponds to a Remortgage.
		// pass:
		// applicationType
		// return:
		// IsRemortgage
		// Raise Errors:
		// ------------------------------------------------------------------------------------------
		public bool IsRemortgage(string applicationType)
		{
			return IsItemInValidation("TypeOfMortgage", applicationType, "R");
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// HasValue if valuationType corresponds to a Reinspection
		// 
		// pass:
		// valuationType	   Valuation type
		// 
		// return:				   True if Reinspection
		// ------------------------------------------------------------------------------------------
		public bool IsReInspection(string valuationType)
		{
			return IsItemInValidation("ValuationType", valuationType, "R");
		}

		protected override bool CanBePooled()
		{
			return true;
		}
	
		#region Ported from IComboDO

		// header ----------------------------------------------------------------------------------
		// description:
		// Return the combo value for Quick Quote location
		// pass:
		// 
		// return:   Quick Quote Location Combo Value
		// ------------------------------------------------------------------------------------------
		public string GetQuickQuoteLocationValueId()
		{
			return ComboAssist.GetFirstComboValueId("PropertyLocation", GlobalAssist.GetGlobalParamString("QQLocation")).ToString();
		}

		// header ----------------------------------------------------------------------------------
		// description:
		//   Return the combo value for Quick Quote location
		// pass:
		//
		// return:   Quick Quote Location Combo Value
		//------------------------------------------------------------------------------------------
		public string GetQuickQuoteValuationTypeValueId()
		{
			return ComboAssist.GetFirstComboValueId("ValuationType", GlobalAssist.GetGlobalParamString("QQValuationType")).ToString();
		}

		// header ----------------------------------------------------------------------------------
		// description:
		// Return the combo values for the entries found to be a dormant legal fee
		// pass:
		// 
		// return:	   Dormant Legal Fee Combo Value
		// ------------------------------------------------------------------------------------------
		public string GetDormantLegalFeeValueId()
		{
			return ComboAssist.GetFirstComboValueId("LegalFeeType", "D").ToString();
		}

		// AS 24/07/2007 Not implemented as specific to BBG MX.
		public XmlNode GetRegSaleRelatedComboList(XmlElement xmlElement)
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.Assists.NotImplementedException();
		}

		#endregion
	}
}

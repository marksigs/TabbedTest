/*
--------------------------------------------------------------------------------------------
Workfile:			XMLAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Helper object for XML parser. 
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		01/07/1999	Created
RF		21/12/1999	Use ErrAssist.ThrowError() not ErrAssist.RaiseError().
SR		29/12/1999	New parameter for the function GetTagValue
SR		02/02/2000	New Method - GetNodeFromClassDefByNodeValue
DJP		08/02/2000	Added GetNode and CheckNode
DJP		17/02/2000	Added GetRequestNodeEx
DJP		17/02/2000	Updated GetRequestNode and GetRequestNodeEx to create a new DOM document when cloning
SR		09/03/2000	Modified methods GetRequestNodeEx - HasValue the REQUEST node is the top
					level node in the Input to the function.
MH		17/03/2000	Changed so that WriteXMLtoFile uses SPM to make the decision
PSC		11/04/2000	AQR SYS0602: Amend GetNodeValue to take in a further optional parameter of
					isTextMandatory to check if the tag text is present
PSC		28/04/2000	AQR SYS0687: Amend GetNodeValue to only throw error if the node
					itself is mandatory as well
SR		03/07/2002 SYS2433 - Modified method GetRequestNodeEx
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/2004	BBG1821 Performance related fixes
--------------------------------------------------------------------------------------------
BMIDS Specific History:

Prog	Date		Description
MV		12/08/2002	BMIDS00323  Core AQR: SYS4754 - Performance. Replace all For...Each...Next with For...Next
					Modified AddNode(),CopyNode(),GroupNodesIntoList(),RemoveDuplicates(),DetermineOperation(),
					DetermineOperationEx(),AreAllNonKeyValuesNull(),AreAllNonKeyValuesNullEx(),
					AttachResponseData(),GetNodeFromClassDefByNodeValue()
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from ParserAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use XmlAssist instead")]
	public static class XMLAssist
    {
		[Obsolete("Use XmlAssist.Load instead")]
        public static XmlDocument load(string xmlText) 
        {
			return XmlAssist.Load(xmlText);
        }

		[Obsolete("Use XmlAssist.GetTagValue instead")]
		public static string GetTagValue(XmlElement xmlElementIn, string tagName)
		{
			bool exists = false;
			return XmlAssist.GetTagValue(xmlElementIn, tagName, out exists, true);
		}

		[Obsolete("Use XmlAssist.GetTagValue instead")]
		public static string GetTagValue(XmlElement xmlElementIn, string tagName, out bool exists, bool deep) 
        {
			return XmlAssist.GetTagValue(xmlElementIn, tagName, out exists, deep);
        }

		[Obsolete("Use XmlAssist.GetAttributeValue instead")]
		public static string GetAttributeValue(XmlElement xmlElementIn, string tagName, string attributeName)
		{
			bool exists = false;
			return XmlAssist.GetAttributeValue(xmlElementIn, tagName, attributeName, out exists);
		}

		[Obsolete("Use XmlAssist.GetAttributeValue instead")]
		public static string GetAttributeValue(XmlElement xmlElementIn, string tagName, string attributeName, out bool exists) 
        {
			return XmlAssist.GetAttributeValue(xmlElementIn, tagName, attributeName, out exists);
        }

		[Obsolete("Use XmlAssist.CopyNode instead")]
		public static bool CopyNode(string sourceNodeName, string targetNodeName, XmlElement xmlSourceElement, XmlElement xmlTargetElement, bool deep) 
        {
			return XmlAssist.CopyNode(sourceNodeName, targetNodeName, xmlSourceElement, xmlTargetElement, deep);
        }

		[Obsolete("Use XmlAssist.AddNode instead")]
		public static bool AddNode(XmlElement xmlElement, string targetNodeName, bool deep, XmlElement xmlDocumentElement)
		{
			return XmlAssist.AddNode(xmlElement, targetNodeName, deep, xmlDocumentElement);
		}

		[Obsolete("Use XmlAssist.GetRequestNodeEx instead")]
		public static XmlNode GetRequestNodeEx(XmlElement xmlElement)
		{
			return XmlAssist.GetRequestNodeEx(xmlElement);
		}

		[Obsolete("Use XmlAssist.GetRequestNode instead")]
		public static XmlNode GetRequestNode(XmlDocument xmlDocument) 
        {
			return XmlAssist.GetRequestNode(xmlDocument.DocumentElement);
		}

		[Obsolete("Use XmlAssist.AddNodeWithUCaseText instead")]
		public static void AddNodeWithUCaseText(XmlNode xmlNode, string tagName)
		{
			XmlAssist.AddNodeWithUCaseText(xmlNode, tagName);
		}

		[Obsolete("Use XmlAssist.ChangeNodeName instead")]
		public static void ChangeNodeName(XmlNode xmlNode, string oldName, string newName)
		{
			XmlAssist.ChangeNodeName(xmlNode, oldName, newName);
		}

		[Obsolete("Use XmlAssist.CountNonNullNodes instead")]
		public static int CountNonNullNodes(XmlElement xmlElement)
		{
			return XmlAssist.CountNonNullNodes(xmlElement);
		}

		[Obsolete("Use XmlAssist.GroupNodesIntoList instead")]
		public static void GroupNodesIntoList(XmlElement xmlElement, string tagName)
		{
			XmlAssist.GroupNodesIntoList(xmlElement, tagName);
		}

		[Obsolete("Use XmlAssist.RemoveDuplicates instead")]
		public static void RemoveDuplicates(XmlElement xmlElement)
		{
			XmlAssist.RemoveDuplicates(xmlElement, "");
		}

		[Obsolete("Use XmlAssist.RemoveDuplicates instead")]
		public static void RemoveDuplicates(XmlElement xmlElement, string compareTagName)
		{
			XmlAssist.RemoveDuplicates(xmlElement, compareTagName);
		}

		[Obsolete("Use XmlAssist.NodesMatch instead")]
		public static bool NodesMatch(XmlNode xmlNode1, XmlNode xmlNode2)
		{
			return XmlAssist.NodesMatch(xmlNode1, xmlNode2, "");
		}

		[Obsolete("Use XmlAssist.NodesMatch instead")]
		public static bool NodesMatch(XmlNode xmlNode1, XmlNode xmlNode2, string compareTagName)
		{
			return XmlAssist.NodesMatch(xmlNode1, xmlNode2, compareTagName);
		}

		[Obsolete("Use XmlAssist.GroupNodes instead")]
		public static void GroupNodes(XmlElement xmlElement, string groupName, string fields) 
		{
			XmlAssist.GroupNodes(xmlElement, groupName, fields);
		}

		[Obsolete("Use XmlAssist.GroupNodesEx instead")]
		public static void GroupNodesEx(XmlElement xmlElement, string groupName, string fields)
		{
			XmlAssist.GroupNodesEx(xmlElement, groupName, fields);
		}

		[Obsolete("Use XmlAssist.BuildComboValueList instead")]
		public static string BuildComboValueList(string groupName, string valueText)
		{
			return XmlAssist.BuildComboValueList(groupName, valueText);
		}

		[Obsolete("Use XmlAssist.DetermineOperation instead")]
		public static BOOPERATIONTYPE DetermineOperation(string request, string classDefinition)
		{
			return XmlAssist.DetermineOperation(request, classDefinition);
		}

		[Obsolete("Use XmlAssist.DetermineOperationEx instead")]
		public static BOOPERATIONTYPE DetermineOperationEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			return XmlAssist.DetermineOperationEx(xmlRequest, xmlClassDefinition);
		}

		[Obsolete("Use XmlAssist.AllChildNodesAreNull instead")]
		public static bool AllChildNodesAreNull(XmlElement xmlElement) 
        {
			return XmlAssist.AllChildNodesAreNull(xmlElement);
		}

		[Obsolete("Use XmlAssist.AreAllNonKeyValuesNull instead")]
		public static bool AreAllNonKeyValuesNull(string request, string classDefinition)
		{
			return XmlAssist.AreAllNonKeyValuesNull(request, classDefinition);
		}

		[Obsolete("Use XmlAssist.AreAllNonKeyValuesNullEx instead")]
		public static bool AreAllNonKeyValuesNullEx(XmlElement xmlElement, XmlDocument xmlClassDefinition)
		{
			return XmlAssist.AreAllNonKeyValuesNullEx(xmlElement, xmlClassDefinition);
		}

		[Obsolete("Use XmlAssist.GetMandatoryAttributeText instead")]
		public static string GetMandatoryAttribute(XmlNode xmlNode, string attributeName)
		{
			return XmlAssist.GetMandatoryAttributeText(xmlNode, attributeName);
		}

		[Obsolete("Use XmlAssist.GetTagValue instead")]
		public static string GetElementText(XmlNode xmlParentNode, string patternText)
		{
			bool exists = false;
			return XmlAssist.GetTagValue((XmlElement)xmlParentNode, patternText, out exists, false);
		}

		[Obsolete("Use XmlAssist.GetTagValue instead")]
		public static bool GetElementTextBln(XmlNode xmlParentNode, string patternText, ref string elementText)
		{
			bool exists = false;
			elementText = XmlAssist.GetTagValue((XmlElement)xmlParentNode, patternText, out exists, false);
			return elementText.Length > 0;
		}

		[Obsolete("Use XmlAssist.GetMandatoryNodeText instead")]
		public static string GetMandatoryElementText(XmlNode xmlParentNode, string patternText)
		{
			return XmlAssist.GetMandatoryNodeText(xmlParentNode, patternText);
		}

		[Obsolete("Use XmlAssist.GetMandatoryNode instead")]
		public static XmlNode GetMandatoryNode(XmlNode xmlParentNode, string patternText)
		{
			return XmlAssist.GetMandatoryNode(xmlParentNode, patternText);
		}

		[Obsolete("Use XmlAssist.AttachResponseData instead")]
		public static void AttachResponseData(XmlNode xmlNodeToAttachTo, XmlElement xmlResponse)
		{
			XmlAssist.AttachResponseData(xmlNodeToAttachTo, xmlResponse);
		}

		[Obsolete("Use XmlAssist.PromoteNode instead")]
		public static void PromoteNode(XmlNode xmlData, XmlNode xmlNodeToPromote)
		{
			XmlAssist.PromoteNode(xmlData, xmlNodeToPromote, true, true);
		}

		[Obsolete("Use XmlAssist.PromoteNode instead")]
		public static void PromoteNode(XmlNode xmlData, XmlNode xmlNodeToPromote, bool replaceParent, bool inheritSiblings)
		{
			XmlAssist.PromoteNode(xmlData, xmlNodeToPromote, replaceParent, inheritSiblings);
		}

		[Obsolete("Use XmlAssist.SelectSchemaFields instead")]
		public static void SelectSchemaFields(string[] varList, XmlNode xmlTableNode)
		{
			XmlAssist.SelectSchemaFields(varList, xmlTableNode);
		}

		[Obsolete("Use XmlAssist.GetAttributeText instead")]
		public static string GetAttributeFromNode(XmlNode xmlNode, string attributeName)
		{
			return XmlAssist.GetAttributeText(xmlNode, attributeName);
		}

		[Obsolete("Use XmlAssist.SetMandatoryChildText instead")]
		public static void SetMandatoryChildText(XmlNode xmlParentNode, string tagName, string childText)
		{
			XmlAssist.SetMandatoryChildText(xmlParentNode, tagName, childText);
		}

		[Obsolete("Use XmlAssist.GetNodeFromClassDefByNodeValue instead")]
		public static XmlNode GetNodeFromClassDefByNodeValue(XmlDocument xmlClassDefinition, string nodeValue)
		{
			return XmlAssist.GetNodeFromClassDefByNodeValue(xmlClassDefinition, nodeValue);
		}

		[Obsolete("Use XmlAssist.CheckNode instead")]
		public static void CheckNode(object xmlNode)
		{
			XmlAssist.CheckNode(xmlNode);
		}

		[Obsolete("Use XmlAssist.CheckNode instead")]
		public static void CheckNode(object xmlNode, string errorText) 
        {
			XmlAssist.CheckNode(xmlNode, errorText);
        }

		[Obsolete("Use XmlAssist.GetMandatoryNode instead")]
		public static XmlNode GetNode(XmlNode xmlNodeParent, string patternText)
		{
			return XmlAssist.GetMandatoryNode(xmlNodeParent, patternText);
		}

		[Obsolete("Use XmlAssist.GetMandatoryNode or XmlAssist.GetNode instead")]
		public static XmlNode GetNode(XmlNode xmlNodeParent, string patternText, bool isNodeMandatory)
		{
			return 
				isNodeMandatory ?
				XmlAssist.GetMandatoryNode(xmlNodeParent, patternText) : 
				XmlAssist.GetNode(xmlNodeParent, patternText);
		}

		[Obsolete("Use XmlAssist.GetMandatoryNodeText or XmlAssist.GetNodeText instead")]
		public static string GetNodeValue(XmlNode xmlNodeParent, string patternText, bool isNodeMandatory, bool isTextMandatory)
		{
			return
				isNodeMandatory ?
				XmlAssist.GetMandatoryNodeText(xmlNodeParent, patternText) :
				XmlAssist.GetNodeText(xmlNodeParent, patternText);
		}

		[Obsolete("Use XmlAssist.StartXMLLogging instead")]
		public static void StartXMLLogging() 
        {
			XmlAssist.StartXMLLogging();
        }

		[Obsolete("Use XmlAssist.StopXMLLogging instead")]
		public static void StopXMLLogging() 
        {
			XmlAssist.StopXMLLogging();
        }
    }
}

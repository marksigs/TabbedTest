/*
--------------------------------------------------------------------------------------------
Workfile:			XmlAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Helper functions for XML handling
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		01/11/2000  Initial creation
ASm	 15/01/2001  SYS1824: Rationalisation of methods following meeting with PC and IK
LD	  30/01/2001  Add optional default value to GetAttribute*
PSC	 22/02/2001  Add GetRequestNode
SR	  05/03/2001  New methods 'SetSysDateToNodeAttrib' and 'SetSysDateToNodeListAttribs'
PSC	 10/04/2001  Add extra parameter to the HasValue/Get mandatory attribut/node methods to
					enable the default message to be overridden
SR	  13/06/2001  SYS2362 - modified method 'MakeNodeElementBased'. Create attrib for
					combo descriptions
SR	  12/07/2001  SYS2412 : new method 'ChangeNodeName'
AW	  02/02/2004  Made 'ParserError' Public in line with Baseline changes
--------------------------------------------------------------------------------------------
BMIDS Specific History:

Prog	Date		Description
MV		12/08/2002  BMIDS00323 Core AQR: SYS4754 - Performance. Replace all For...Each...Next with For...Next
					Modified AttachResponseData(),CreateChildRequest(),CreateAttributeBasedResponse(),
					MakeNodeElementBased(),CreateElementRequestFromNode(),
					ChangeNodeName(),SetSysDateToNodeListAttribs()
--------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
MV		23/02/2004	BBG74 Created a new method xmlCreateDOMObject
TK		14/01/2005	BBG1891 Commented out all unused objects and set to nothing after using objects.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from xmlAssist.bas and ParserAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

// Ambiguous reference in cref attribute.
#pragma warning disable 419

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// The types of operations that can be performed by a business object.
	/// </summary>
	public enum BOOPERATIONTYPE
	{
		/// <summary>
		/// Not defined.
		/// </summary>
		booNone = 0,
		/// <summary>
		/// Insert.
		/// </summary>
		booCreate = 1,
		/// <summary>
		/// Delete.
		/// </summary>
		booDelete = 2,
		/// <summary>
		/// Update.
		/// </summary>
		booUpdate = 3,
	}

	/// <summary>
	/// Wrapper methods for <see cref="System.Xml"/>. 
	/// </summary>
	/// <remarks>
	/// <para>
	/// <i>Do not use in new code - use the <see cref="System.Xml"/> namespace directly instead.</i>
	/// </para>
	/// <para>
	/// Methods have been ported from the VB6 XmlAssist module 
	/// (file name xmlAssist.bas) and from the VB6 XMLAssist module 
	/// (file name ParserAssist.cls). 
	/// </para>
	/// </remarks>
	public static class XmlAssist
	{
		#region xmlAssistEx

		/// <summary>
		/// Creates a new <see cref="System.Xml.XmlDocument"/> object.
		/// </summary>
		/// <returns>
		/// The <see cref="System.Xml.XmlDocument"/> object.
		/// </returns>
		public static XmlDocument CreateDOMObject()
		{
			return new XmlDocument();
		}

		/// <summary>
		/// Sets a specified attribute to the current date and time if the attribute does not exist 
		/// or does not already have a value.
		/// </summary>
		/// <param name="xmlNode"></param>
		/// <param name="dateAttributeName"></param>
		public static void SetSysDateToNodeAttrib(XmlNode xmlNode, string dateAttributeName)
		{
			SetSysDateToNodeAttrib(xmlNode, dateAttributeName, false);
		}

		/// <summary>
		/// Sets a specified attribute to the current date and time.
		/// </summary>
		/// <param name="xmlNode">The node which contains the attribute.</param>
		/// <param name="dateAttributeName">The name of the attribute.</param>
		/// <param name="overWrite">Indicates whether to overwrite any existing value.</param>
		public static void SetSysDateToNodeAttrib(XmlNode xmlNode, string dateAttributeName, bool overWrite) 
		{
			if (GetAttributeText(xmlNode, dateAttributeName, "").Length > 0)
			{
				if (overWrite)
				{
					SetAttributeValue(xmlNode, dateAttributeName, DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
				}
			}
			else
			{
				SetAttributeValue(xmlNode, dateAttributeName, DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
			}
		}

		/// <summary>
		/// For each node in a list, sets a specified attribute on the node to the current date and time, 
		/// if the attribute does not exist or does not already have a value.
		/// </summary>
		/// <param name="xmlNodeList">The list of nodes.</param>
		/// <param name="dateAttributeName">The name of the attribute.</param>
		public static void SetSysDateToNodeListAttribs(XmlNodeList xmlNodeList, string dateAttributeName)
		{
			SetSysDateToNodeListAttribs(xmlNodeList, dateAttributeName, false);
		}

		/// <summary>
		/// For each node in a list, sets a specified attribute on the node to the current date and time.
		/// </summary>
		/// <param name="xmlNodeList">The list of nodes.</param>
		/// <param name="dateAttributeName">The name of the attribute.</param>
		/// <param name="overWrite">Indicates whether to overwrite any existing value.</param>
		public static void SetSysDateToNodeListAttribs(XmlNodeList xmlNodeList, string dateAttributeName, bool overWrite) 
		{
			foreach (XmlNode xmlNode in xmlNodeList)
			{
				SetSysDateToNodeAttrib(xmlNode, dateAttributeName, false);
			}
		}

		/// <summary>
		/// Throws a new <see cref="XMLParserErrorException"/> for a specified <see cref="XmlException"/>.
		/// </summary>
		/// <param name="exception">The <see cref="XmlException"/>.</param>
		public static void ParserError(XmlException exception) 
		{
			throw new XMLParserErrorException(FormatParserError(exception));
		}

		/// <summary>
		/// Formats a specified <see cref="XmlException"/>.
		/// </summary>
		/// <param name="exception">The <see cref="XmlException"/>.</param>
		/// <returns>The formatted exception.</returns>
		private static string FormatParserError(XmlException exception) 
		{
			return exception.ToString();
		}

		/// <summary>
		/// Creates and loads an xml Document from a string.
		/// </summary>
		/// <param name="xmlText">The xml string to be loaded.</param>
		/// <returns>An <see cref="XmlDocument"/> containing the xml.</returns>
		public static XmlDocument Load(string xmlText) 
		{
			XmlDocument xmlDocument = new XmlDocument();
			try
			{
				xmlDocument.LoadXml(xmlText);
			}
			catch (XmlException exception)
			{
				ParserError(exception);
			}
			return xmlDocument;
		}

		/// <summary>
		/// Searches a specified node for a child node that matches a specified XPath expression.
		/// </summary>
		/// <param name="xmlParentNode">The node to search.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="isMandatory">Indicates whether a exception should be thrown if the child node does not exist.</param>
		/// <param name="omigaError">The Omiga error to throw if the child node does not exist.</param>
		/// <returns>The matching child node.</returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and <paramref name="isMandatory"/> is <b>true</b>, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist, and <paramref name="isMandatory"/> is <b>true</b>, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		private static XmlNode GetNode(XmlNode xmlParentNode, string xpath, bool isMandatory, OMIGAERROR omigaError) 
		{
			if (xmlParentNode == null)
			{
				throw new MissingElementException("Missing parent node. Match xpath: " + xpath);
			}
			XmlNode xmlNode = xmlParentNode.SelectSingleNode(xpath);
			if (isMandatory == true && xmlNode == null)
			{
				if (omigaError == OMIGAERROR.UnspecifiedError)
				{
					throw new MissingElementException("Match xpath: " + xpath);
				}
				else
				{
					throw new ErrAssistException(omigaError);
				}
			}
			return xmlNode;
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The matching child node, or null if no node is found.
		/// </returns>
		public static XmlNode GetNode(XmlNode xmlParentNode, string xpath) 
		{
			return GetNode(xmlParentNode, xpath, false, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node, or the empty string 
		/// if no matching child node was found.
		/// </returns>
		public static string GetNodeText(XmlNode xmlParentNode, string xpath) 
		{
			XmlNode xmlNode = GetNode(xmlParentNode, xpath);
			return xmlNode != null ? xmlNode.InnerText : "";
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node as a bool.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a bool. 
		/// The following strings will return <b>true</b> (anything else returns <b>false</b>): 
		/// "1", "Y", "YES" (text comparisons are case insensitive).
		/// </returns>
		public static bool GetNodeAsBoolean(XmlNode xmlParentNode, string xpath) 
		{
			return IsBooleanText(GetNodeText(xmlParentNode, xpath));
		}

		private static bool IsBooleanText(string valueText)
		{
			valueText = valueText.ToUpper();
			return valueText == "1" || valueText == "Y" || valueText == "YES";
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node as a DateTime object.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a DateTime object. 
		/// </returns>
		public static DateTime GetNodeAsDate(XmlNode xmlParentNode, string xpath) 
		{
			return ConvertAssist.CSafeDate(GetNodeText(xmlParentNode, xpath));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node as a double.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a double. 0 is returned 
		/// if no matching node is found.
		/// </returns>
		public static double GetNodeAsDouble(XmlNode xmlParentNode, string xpath) 
		{
			return ConvertAssist.CSafeDbl(GetNodeText(xmlParentNode, xpath));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node as an int.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to an int. 0 is returned 
		/// if no matching node is found.
		/// </returns>
		public static int GetNodeAsInteger(XmlNode xmlParentNode, string xpath) 
		{
			return ConvertAssist.CSafeInt(GetNodeText(xmlParentNode, xpath));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and returns 
		/// the inner text in the matching child node as a long.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a long. 0 is returned 
		/// if no matching node is found.
		/// </returns>
		public static int GetNodeAsLong(XmlNode xmlParentNode, string xpath) 
		{
			return ConvertAssist.CSafeLng(GetNodeText(xmlParentNode, xpath));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// throws an <see cref="MissingElementException"/> exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		public static void CheckMandatoryNode(XmlNode xmlParentNode, string xpath)
		{
			CheckMandatoryNode(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// throws an exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CheckMandatoryNode(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			GetNode(xmlParentNode, xpath, true, omigaError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the matching child node, or throws an <see cref="MissingElementException"/> exception 
		/// if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The matching child node.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static XmlNode GetMandatoryNode(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNode(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the matching child node, or throws an exception 
		/// if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The matching child node.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static XmlNode GetMandatoryNode(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return GetNode(xmlParentNode, xpath, true, omigaError);
		}

		/// <summary>
		/// Searches in a specified node for child nodes that match a specified XPath expression, and 
		/// returns the matching child nodes, or throws an exception 
		/// if the child nodes do not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The matching child nodes.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child nodes do not exist.
		/// </exception>
		public static XmlNodeList GetMandatoryNodeList(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNodeList(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for child nodes that match a specified XPath expression, and 
		/// returns the matching child nodes, or throws an exception 
		/// if the child nodes do not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child nodes do not exist.</param>
		/// <returns>
		/// The matching child nodes.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child nodes do not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child nodes do not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static XmlNodeList GetMandatoryNodeList(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			XmlNodeList xmlMandatoryNodeList = xmlParentNode.SelectNodes(xpath);
			if (xmlMandatoryNodeList.Count == 0)
			{
				if (omigaError == OMIGAERROR.UnspecifiedError)
				{
					throw new MissingElementException("Match xpath: " + xpath);
				}
				else
				{
					throw new ErrAssistException(omigaError);
				}
			}
			return xmlMandatoryNodeList;
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static string GetMandatoryNodeText(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNodeText(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node.
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static string GetMandatoryNodeText(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			string valueText = GetMandatoryNode(xmlParentNode, xpath, omigaError).InnerText;
			if (valueText.Length == 0)
			{
				if (omigaError == OMIGAERROR.UnspecifiedError)
				{
					throw new MissingElementValueException("Match xpath:- " + xpath);
				}
				else
				{
					throw new ErrAssistException(omigaError);
				}
			}
			return valueText;
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a bool, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a bool. 
		/// The following strings will return <b>true</b> (anything else returns <b>false</b>): 
		/// "1", "Y", "YES" (text comparisons are case insensitive).
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static bool GetMandatoryNodeAsBoolean(XmlNode xmlParentNode, string xpath) 
		{
			return GetMandatoryNodeAsBoolean(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a bool, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a bool. 
		/// The following strings will return <b>true</b> (anything else returns <b>false</b>): 
		/// "1", "Y", "YES" (text comparisons are case insensitive).
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static bool GetMandatoryNodeAsBoolean(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return IsBooleanText(GetMandatoryNodeText(xmlParentNode, xpath, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a DateTime object, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a DateTime object. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static DateTime GetMandatoryNodeAsDate(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNodeAsDate(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a DateTime object, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a DateTime object. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static DateTime GetMandatoryNodeAsDate(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeDate(GetMandatoryNodeText(xmlParentNode, xpath, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a double, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a double. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static double GetMandatoryNodeAsDouble(XmlNode xmlParentNode, string xpath) 
		{
			return GetMandatoryNodeAsDouble(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a double, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a double. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static double GetMandatoryNodeAsDouble(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeDbl(GetMandatoryNodeText(xmlParentNode, xpath, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as an int, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to an int. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static int GetMandatoryNodeAsInteger(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNodeAsInteger(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as an int, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to an int. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static int GetMandatoryNodeAsInteger(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeInt(GetMandatoryNodeText(xmlParentNode, xpath, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a long, or throws an <see cref="MissingElementException"/> 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a long. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist.
		/// </exception>
		public static int GetMandatoryNodeAsLong(XmlNode xmlParentNode, string xpath)
		{
			return GetMandatoryNodeAsLong(xmlParentNode, xpath, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for a child node that matches a specified XPath expression, and 
		/// returns the inner text in the matching child node as a long, or throws an 
		/// exception if the child node does not exist.
		/// </summary>
		/// <param name="xmlParentNode">The node to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the child node does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching child node converted to a long. 
		/// </returns>
		/// <exception cref="MissingElementException">
		/// The child node does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The child node does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static int GetMandatoryNodeAsLong(XmlNode xmlParentNode, string xpath, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeLng(GetMandatoryNodeText(xmlParentNode, xpath, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The matching attribute node, or null if no matching attribute is found.
		/// </returns>
		public static XmlNode GetAttributeNode(XmlNode xmlNode, string attributeName) 
		{
			return xmlNode.Attributes.GetNamedItem(attributeName);
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, and returns true if the attribute exists.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// <b>true</b> if the attribute exists and its InnerText property is not null and is not the empty string, otherwise <b>false</b>.
		/// </returns>
		public static bool AttributeValueExists(XmlNode xmlNode, string attributeName) 
		{
			xmlNode = GetAttributeNode(xmlNode, attributeName);
			return xmlNode != null && xmlNode.InnerText != null && xmlNode.InnerText.Length > 0;
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, or the empty string if 
		/// the attribute does not exist.
		/// </returns>
		public static string GetAttributeText(XmlNode xmlNode, string attributeName) 
		{
			return GetAttributeText(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, or <paramref name="defaultText"/> if 
		/// the attribute does not exist.
		/// </returns>
		public static string GetAttributeText(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			xmlNode = GetAttributeNode(xmlNode, attributeName);
			return xmlNode != null ? xmlNode.InnerText : defaultText;
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a bool.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a bool, or <b>false</b> if 
		/// the attribute does not exist. 
		/// The following strings will return <b>true</b> (anything else returns <b>false</b>): 
		/// "1", "Y", "YES" (text comparisons are case insensitive).
		/// </returns>
		public static bool GetAttributeAsBoolean(XmlNode xmlNode, string attributeName)
		{
			return GetAttributeAsBoolean(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a bool.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return (as a bool) if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, or <paramref name="defaultText"/> if 
		/// the attribute does not exist, converted to a bool. 
		/// The following strings will return <b>true</b> (anything else returns <b>false</b>): 
		/// "1", "Y", "YES" (text comparisons are case insensitive).
		/// </returns>
		public static bool GetAttributeAsBoolean(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			return IsBooleanText(GetAttributeText(xmlNode, attributeName, defaultText));
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a DateTime object.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a DateTime object, 
		/// or DateTime.Now if the attribute does not exist. 
		/// </returns>
		public static DateTime GetAttributeAsDate(XmlNode xmlNode, string attributeName) 
		{
			return GetAttributeAsDate(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a DateTime object.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return (as a DateTime object) if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, or <paramref name="defaultText"/> if 
		/// the attribute does not exist, converted to a DateTime object. 
		/// </returns>
		public static DateTime GetAttributeAsDate(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			return ConvertAssist.CSafeDate(GetAttributeText(xmlNode, attributeName, defaultText));
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a double.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a double, 
		/// or 0 if the attribute does not exist. 
		/// </returns>
		public static double GetAttributeAsDouble(XmlNode xmlNode, string attributeName)
		{
			return GetAttributeAsDouble(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a double.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return (as a double) if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, 
		/// or <paramref name="defaultText"/> if the attribute does not exist, converted to a double. 
		/// </returns>
		public static double GetAttributeAsDouble(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			return ConvertAssist.CSafeDbl(GetAttributeText(xmlNode, attributeName, defaultText));
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as an int.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to an int, 
		/// or 0 if the attribute does not exist. 
		/// </returns>
		public static int GetAttributeAsInteger(XmlNode xmlNode, string attributeName)
		{
			return GetAttributeAsInteger(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as an int.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return (as an int) if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, 
		/// or <paramref name="defaultText"/> if the attribute does not exist, converted to an int. 
		/// </returns>
		public static int GetAttributeAsInteger(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			return ConvertAssist.CSafeInt(GetAttributeText(xmlNode, attributeName, defaultText));
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a long.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a long, 
		/// or 0 if the attribute does not exist. 
		/// </returns>
		public static int GetAttributeAsLong(XmlNode xmlNode, string attributeName)
		{
			return GetAttributeAsLong(xmlNode, attributeName, "");
		}

		/// <summary>
		/// Searches in a specified node for an attibute with a specified name, 
		/// and returns the inner text for the attribute as a long.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="defaultText">The text to return (as a long) if the attribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute, 
		/// or <paramref name="defaultText"/> if the attribute does not exist, converted to a long. 
		/// </returns>
		public static int GetAttributeAsLong(XmlNode xmlNode, string attributeName, string defaultText) 
		{
			return ConvertAssist.CSafeLng(GetAttributeText(xmlNode, attributeName, defaultText));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the matching attribute node, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The matching attribute node.
		/// </returns>
		public static XmlNode GetMandatoryAttribute(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttribute(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the matching attribute node, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The matching attribute node.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static XmlNode GetMandatoryAttribute(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			XmlNode xmlAttribute = xmlNode.Attributes.GetNamedItem(attributeName);
			if (xmlAttribute == null)
			{
				if (omigaError == OMIGAERROR.UnspecifiedError)
				{
					throw new XMLMissingAttributeException("[@" + attributeName + "]");
				}
				else
				{
					throw new ErrAssistException(omigaError);
				}
			}
			return xmlAttribute;
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static string GetMandatoryAttributeText(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttributeText(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static string GetMandatoryAttributeText(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			string text = null;
			XmlNode xmlAttribute = GetMandatoryAttribute(xmlNode, attributeName, omigaError);
			if (xmlAttribute.InnerText == null || xmlAttribute.InnerText.Length == 0)
			{
				if (omigaError == OMIGAERROR.UnspecifiedError)
				{
					throw new XMLInvalidAttributeValueException("[@" + attributeName + "]");
				}
				else
				{
					throw new ErrAssistException(omigaError);
				}
			}
			else
			{
				text = xmlAttribute.InnerText;
			}
			return text;
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a bool, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a bool.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static bool GetMandatoryAttributeAsBoolean(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttributeAsBoolean(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a bool, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a bool.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>. 
		/// </exception>
		public static bool GetMandatoryAttributeAsBoolean(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			return IsBooleanText(GetMandatoryAttributeText(xmlNode, attributeName, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a DateTime object, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a DateTime object.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static DateTime GetMandatoryAttributeAsDate(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttributeAsDate(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a DateTime object, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a DateTime object.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static DateTime GetMandatoryAttributeAsDate(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeDate(GetMandatoryAttributeText(xmlNode, attributeName, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a double, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a double.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static double GetMandatoryAttributeAsDouble(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttributeAsDouble(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a double, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a double.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static double GetMandatoryAttributeAsDouble(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeDbl(GetMandatoryAttributeText(xmlNode, attributeName, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to an int, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to an int.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static int MandatoryGetAttributeAsInteger(XmlNode xmlNode, string attributeName)
		{
			return MandatoryGetAttributeAsInteger(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to an int, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to an int.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static int MandatoryGetAttributeAsInteger(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeInt(GetMandatoryAttributeText(xmlNode, attributeName, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a long, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a long.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static int GetMandatoryAttributeAsLong(XmlNode xmlNode, string attributeName)
		{
			return GetMandatoryAttributeAsLong(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, and 
		/// returns the inner text of the matching attribute node converted to a long, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <returns>
		/// The InnerText property of the matching attribute converted to a long.
		/// </returns>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static int GetMandatoryAttributeAsLong(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			return ConvertAssist.CSafeLng(GetMandatoryAttributeText(xmlNode, attributeName, omigaError));
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static void CheckMandatoryAttribute(XmlNode xmlNode, string attributeName)
		{
			CheckMandatoryAttribute(xmlNode, attributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Searches in a specified node for an attribute with a specified name, or throws an exception 
		/// if the attribute does not exist.
		/// </summary>
		/// <param name="xmlNode">The node to be searched.</param>
		/// <param name="attributeName">The attribute name.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CheckMandatoryAttribute(XmlNode xmlNode, string attributeName, OMIGAERROR omigaError) 
		{
			GetMandatoryAttributeText(xmlNode, attributeName, omigaError);
		}

		/// <summary>
		/// Copies a specified attribute from a source node to a destination node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the attribute to copy.</param>
		/// <remarks>
		/// Any existing attribute in the <paramref name="xmlDestinationNode"/> will be overwritten. 
		/// Does nothing if the attribute does not exist in <paramref name="xmlSourceNode"/>, 
		/// or it has no value.
		/// </remarks>
		public static void CopyAttribute(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName) 
		{
			if (AttributeValueExists(xmlSourceNode, sourceAttributeName))
			{
				xmlDestinationNode.Attributes.SetNamedItem(
					xmlSourceNode.Attributes.GetNamedItem(sourceAttributeName).CloneNode(true));
			}
		}

		/// <summary>
		/// Copies a specified attribute from a source node to a destination node, 
		/// or throws an exception if the attribute does not exist in the source node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the attribute to copy.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static void CopyMandatoryAttribute(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName)
		{
			CopyMandatoryAttribute(xmlSourceNode, xmlDestinationNode, sourceAttributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Copies a specified attribute from a source node to a destination node, 
		/// or throws an exception if the attribute does not exist in the source node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the attribute to copy.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The source attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CopyMandatoryAttribute(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, OMIGAERROR omigaError) 
		{
			CheckMandatoryAttribute(xmlSourceNode, sourceAttributeName, omigaError);
			CopyAttribute(xmlSourceNode, xmlDestinationNode, sourceAttributeName);
		}

		/// <summary>
		/// Copies the value of a specified attribute in a source node to a specified attribute in a destination node, 
		/// overwriting any existing value in the destination attribute.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute.</param>
		public static void CopyAttributeValue(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName) 
		{
			if (AttributeValueExists(xmlSourceNode, sourceAttributeName))
			{
				XmlAttribute xmlAttribute = xmlDestinationNode.OwnerDocument.CreateAttribute(destinationAttributeName);
				xmlAttribute.InnerText = xmlSourceNode.Attributes.GetNamedItem(sourceAttributeName).InnerText;
				xmlDestinationNode.Attributes.SetNamedItem(xmlAttribute);
			}
		}

		/// <summary>
		/// Copies the value of a specified attribute in a source node to a specified attribute in a destination node,  
		/// overwriting any existing value in the destination attribute.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static void CopyMandatoryAttributeValue(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName)
		{
			CopyMandatoryAttributeValue(xmlSourceNode, xmlDestinationNode, sourceAttributeName, destinationAttributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Copies the value of a specified attribute in a source node to a specified attribute in a destination node,  
		/// overwriting any existing value in the destination attribute.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The source attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CopyMandatoryAttributeValue(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName, OMIGAERROR omigaError) 
		{
			CheckMandatoryAttribute(xmlSourceNode, sourceAttributeName, omigaError);
			XmlAttribute xmlAttribute = xmlDestinationNode.OwnerDocument.CreateAttribute(destinationAttributeName);
			xmlAttribute.InnerText = xmlSourceNode.Attributes.GetNamedItem(sourceAttributeName).InnerText;
			xmlDestinationNode.Attributes.SetNamedItem(xmlAttribute);
		}

		/// <summary>
		/// Copies a specified attribute in a source node to a destination node, 
		/// unless the attribute already exists in the destination node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		public static void CopyAttribIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName) 
		{
			if (!AttributeValueExists(xmlDestinationNode, sourceAttributeName))
			{
				CopyAttribute(xmlSourceNode, xmlDestinationNode, sourceAttributeName);
			}
		}

		/// <summary>
		/// Copies a specified attribute in a source node to a destination node, 
		/// unless the attribute already exists in the destination node; 
		/// throws an exception if the attribute does not exist in the source node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string.
		/// </exception>
		public static void CopyMandatoryAttribIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName)
		{
			CopyMandatoryAttribIfMissingFromDest(xmlSourceNode, xmlDestinationNode, sourceAttributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Copies a specified attribute in a source node to a destination node, 
		/// unless the attribute already exists in the destination node; 
		/// throws an exception if the attribute does not exist in the source node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute.</param>
		/// <param name="sourceAttributeName">The name of the source attribute to copy.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist and <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The source attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CopyMandatoryAttribIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, OMIGAERROR omigaError) 
		{
			CheckMandatoryAttribute(xmlSourceNode, sourceAttributeName, omigaError);
			CopyAttribIfMissingFromDest(xmlSourceNode, xmlDestinationNode, sourceAttributeName);
		}

		/// <summary>
		/// Copies a specified attribute value in a source node to an attribute in a destination node, 
		/// unless the attribute already exists in the destination node.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute from which to copy the value.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute to which to copy the value.</param>
		public static void CopyAttribValueIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName) 
		{
			if (!AttributeValueExists(xmlDestinationNode, sourceAttributeName))
			{
				CopyAttributeValue(xmlSourceNode, xmlDestinationNode, sourceAttributeName, destinationAttributeName);
			}
		}

		/// <summary>
		/// Copies a specified attribute value in a source node to an attribute in a destination node, 
		/// unless the attribute already exists in the destination node; 
		/// throws an exception if the attribute does not exist in the source node, 
		/// or it has no value.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute from which to copy the value.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute to which to copy the value.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The source attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CopyMandatoryAttribValueIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName)
		{
			CopyMandatoryAttribValueIfMissingFromDest(xmlSourceNode, xmlDestinationNode, sourceAttributeName, destinationAttributeName, OMIGAERROR.UnspecifiedError);
		}

		/// <summary>
		/// Copies a specified attribute value in a source node to an attribute in a destination node, 
		/// unless the attribute already exists in the destination node; 
		/// throws an exception if the attribute does not exist in the source node, 
		/// or it has no value.
		/// </summary>
		/// <param name="xmlSourceNode">The node from which to copy the attribute value.</param>
		/// <param name="xmlDestinationNode">The node to which to copy the attribute value.</param>
		/// <param name="sourceAttributeName">The name of the source attribute from which to copy the value.</param>
		/// <param name="destinationAttributeName">The name of the destination attribute to which to copy the value.</param>
		/// <param name="omigaError">The Omiga error to include in the <see cref="ErrAssistException"/> exception that is thrown if the atttribute does not exist.</param>
		/// <exception cref="XMLMissingAttributeException">
		/// The source attribute does not exist, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="XMLInvalidAttributeValueException">
		/// The source attribute inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		/// <exception cref="ErrAssistException">
		/// The source attribute does not exist, or the inner text is null or the inner text is the empty string, and 
		/// <paramref name="omigaError"/> is not <b>OMIGAERROR.UnspecifiedError</b>.
		/// </exception>
		public static void CopyMandatoryAttribValueIfMissingFromDest(XmlNode xmlSourceNode, XmlNode xmlDestinationNode, string sourceAttributeName, string destinationAttributeName, OMIGAERROR omigaError) 
		{
			CheckMandatoryAttribute(xmlSourceNode, sourceAttributeName, omigaError);
			if (!AttributeValueExists(xmlDestinationNode, sourceAttributeName))
			{
				CopyAttribValueIfMissingFromDest(
					xmlSourceNode, xmlDestinationNode, sourceAttributeName, destinationAttributeName);
			}
		}

		/// <summary>
		/// Creates a new attribute in a specified node, and sets the attribute value.
		/// </summary>
		/// <param name="xmlDestinationNode">The node where the attribute will be created.</param>
		/// <param name="attributeName">The name of the new attribute.</param>
		/// <param name="valueText">The value for the new attribute.</param>
		public static void SetAttributeValue(XmlNode xmlDestinationNode, string attributeName, string valueText) 
		{
			XmlAttribute xmlAttribute = xmlDestinationNode.OwnerDocument.CreateAttribute(attributeName);
			xmlAttribute.Value = valueText;
			xmlDestinationNode.Attributes.SetNamedItem(xmlAttribute.CloneNode(true));
		}

		/// <summary>
		/// Creates a new attribute in a specified node, and sets the attribute value, 
		/// unless the attribute already exists.
		/// </summary>
		/// <param name="xmlDestinationNode">The node where the attribute will be created.</param>
		/// <param name="attributeName">The name of the new attribute.</param>
		/// <param name="valueText">The value for the new attribute.</param>
		public static void SetAttributeValueIfMissingFromDest(XmlNode xmlDestinationNode, string attributeName, string valueText) 
		{
			if (!AttributeValueExists(xmlDestinationNode, attributeName))
			{
				XmlAttribute xmlAttribute = xmlDestinationNode.OwnerDocument.CreateAttribute(attributeName);
				xmlAttribute.Value = valueText;
				xmlDestinationNode.Attributes.SetNamedItem(xmlAttribute.CloneNode(true));
			}
		}

		/// <summary>
		/// For all child nodes in a specified parent node, if the child node has a specified name, then change it to a new name.
		/// </summary>
		/// <param name="xmlNode">The parent node.</param>
		/// <param name="oldName">The old name to replace.</param>
		/// <param name="newName">The new name.</param>
		public static void ChangeNodeName(XmlNode xmlNode, string oldName, string newName)
		{
			// If this node is to be renamed create a new node with the new name
			// and copy all the Attributes across

			XmlNode xmlNewNode = null;

			if (xmlNode.Name == oldName)
			{
				xmlNewNode = xmlNode.OwnerDocument.CreateNode(xmlNode.NodeType, newName, xmlNode.NamespaceURI);
				// Copy Attributes if this is an element node
				if (xmlNewNode.NodeType == XmlNodeType.Element)
				{
					foreach (XmlNode xmlAttribute in xmlNode.Attributes)
					{
						((XmlElement)xmlNewNode).SetAttributeNode((XmlAttribute)xmlAttribute.CloneNode(true));
					}
				}
			}

			// For all children of this node change their name if it is oldName and
			// append them to the new node
			for (int index = 0; index < xmlNode.ChildNodes.Count; index++)
			{
				XmlNode xmlChildNode = xmlNode.ChildNodes[index];
				ChangeNodeName(xmlChildNode, oldName, newName);
				if (xmlNewNode != null)
				{
					xmlNewNode.AppendChild(xmlChildNode.CloneNode(true));
				}
			}
			// Replace the original node with the new node
			if (xmlNewNode != null)
			{
				if (xmlNode.ParentNode != null)
				{
					xmlNode.ParentNode.ReplaceChild(xmlNewNode, xmlNode);
				}
				xmlNode = xmlNewNode;
			}
		}

		/// <summary>
		/// Creates an element based request node from an attribute based request node.
		/// </summary>
		/// <param name="xmlNode">The attribute based request node.</param>
		/// <param name="xpath">The XPath expression to use when searching for child nodes to convert.</param>
		/// <param name="recursive">Indicates whether to recursively convert child nodes.</param>
		/// <returns>
		/// The XmlDocument object containing the converted node.
		/// </returns>
		public static XmlDocument CreateElementRequestFromNode(XmlNode xmlNode, string xpath, bool recursive)
		{
			return CreateElementRequestFromNode(xmlNode, xpath, recursive, "");
		}

		/// <summary>
		/// Creates an element based request node from an attribute based request node.
		/// </summary>
		/// <param name="xmlNode">The attribute based request node.</param>
		/// <param name="xpath">The XPath expression to use when searching for child nodes to convert.</param>
		/// <param name="recursive">Indicates whether to recursively convert child nodes.</param>
		/// <param name="newTagName">The new name for the node.</param>
		/// <returns>
		/// The XmlDocument object containing the converted node.
		/// </returns>
		public static XmlDocument CreateElementRequestFromNode(XmlNode xmlNode, string xpath, bool recursive, string newTagName) 
		{
			XmlDocument xmlDocument = null;
			XmlNode xmlRequestNode = null;
			XmlNode xmlDestParentNode = null;

			if (xmlNode != null)
			{
				xmlDocument = new XmlDocument();
				// create a new request node
				XmlElement xmlElement = xmlDocument.CreateElement("REQUEST");
				if (xmlElement != null)
				{
					xmlRequestNode = xmlDocument.AppendChild(xmlElement);
				}
				// clone the request Attributes of the passed in request node
				foreach (XmlAttribute xmlAttribute in xmlNode.Attributes)
				{
					CopyAttribute(xmlNode, xmlRequestNode, xmlAttribute.Name);
				}
				// Extract the nodes to convert.
				XmlNodeList xmlNodeList = xmlNode.SelectNodes(xpath);
				if (xmlNodeList.Count == 0)
				{
					throw new XMLMissingElementTextException("No matching nodes found for: " + xpath);
				}
				foreach (XmlNode xmlChildNode in xmlNodeList)
				{
					xmlDestParentNode = 
						xmlRequestNode.AppendChild(xmlDocument.ImportNode(
							MakeNodeElementBased(xmlChildNode, recursive, newTagName), true));
				}
			}
			return xmlDocument;
		}

		/// <summary>
		/// Creates an element based node from an attribute based node.
		/// </summary>
		/// <param name="xmlSourceNode">The attribute based request node.</param>
		/// <param name="recursive">Indicates whether to recursively convert child nodes.</param>
		/// <param name="targetNodeName">The new name for the node.</param>
		/// <param name="attributeNames">Optional list of attributes to be converted. If omitted then all attributes are converted.</param>
		/// <returns>
		/// The converted node.
		/// </returns>
		public static XmlNode MakeNodeElementBased(XmlNode xmlSourceNode, bool recursive, string targetNodeName, params string[] attributeNames) 
		{
			return ConvertNodeToElementBased(null, xmlSourceNode, recursive, targetNodeName, attributeNames);
		}

		private static XmlNode ConvertNodeToElementBased(XmlNode xmlParentNode, XmlNode xmlSourceNode, bool recursive, string targetNodeName, params string[] attributeNames)
		{
			XmlNode xmlTargetNode = null;

			if (xmlSourceNode != null)
			{
				if (xmlSourceNode.NodeType == XmlNodeType.Element)
				{
					// Determine the name of the new node
					if (targetNodeName == null || targetNodeName.Trim().Length == 0)
					{
						targetNodeName = xmlSourceNode.Name;
					}

					if (xmlParentNode == null)
					{
						// No existing parent so create a new document.
						XmlDocument xmlDocument = new XmlDocument();
						xmlTargetNode = xmlDocument.CreateElement(targetNodeName);
						xmlDocument.AppendChild(xmlTargetNode);
					}
					else
					{
						xmlTargetNode = xmlParentNode.OwnerDocument.CreateElement(targetNodeName);
						xmlParentNode.AppendChild(xmlTargetNode);
					}

					// Recursively add child nodes.
					if (recursive && xmlSourceNode.HasChildNodes)
					{
						foreach (XmlNode xmlChildNode in xmlSourceNode.ChildNodes)
						{
							ConvertNodeToElementBased(xmlTargetNode, xmlChildNode, recursive, null);
						}
					}

					if (attributeNames != null && attributeNames.Length > 0)
					{
						// Convert name attributes to nodes.
						foreach (string attributeName in attributeNames)
						{
							string attributeText = GetAttributeText(xmlSourceNode, attributeName);
							if (attributeText != null && attributeText.Length > 0)
							{
								ConvertAttributeToNode(xmlTargetNode, attributeName, attributeText);
							}
							else
							{
								throw new XMLMissingAttributeException(attributeName);
							}
						}
					}
					else
					{
						// Convert all attributes to nodes.
						foreach (XmlAttribute xmlAttribute in xmlSourceNode.Attributes)
						{
							ConvertAttributeToNode(xmlTargetNode, xmlAttribute.Name, xmlAttribute.InnerText);
						}
					}
				}
				else if (xmlSourceNode.NodeType == XmlNodeType.Text)
				{
					xmlParentNode.AppendChild(xmlParentNode.OwnerDocument.CreateTextNode(xmlSourceNode.InnerText));
				}
			}

			return xmlTargetNode;
		}

		private static XmlNode ConvertAttributeToNode(XmlNode xmlParentNode, string attributeName, string attributeText)
		{
			XmlNode xmlConvertedNode = null;

			const string textName = "_TEXT";

			// SR 12/06/01 : if node name is description of any combo value, create a new attribute
			// in the respective node
			if (attributeName.Length > textName.Length && attributeName.Substring(attributeName.Length - textName.Length) == textName)
			{
				string comboValueNodeName = attributeName.Substring(0, attributeName.Length - textName.Length);
				xmlConvertedNode = xmlParentNode.SelectSingleNode("./" + comboValueNodeName);
				if (xmlConvertedNode != null)
				{
					SetAttributeValue(xmlConvertedNode, "TEXT", attributeText);
				}
			}

			if (xmlConvertedNode == null)
			{
				xmlConvertedNode = xmlParentNode.OwnerDocument.CreateElement(attributeName);
				xmlConvertedNode.InnerText = attributeText;
				xmlParentNode.AppendChild(xmlConvertedNode);
			}

			return xmlConvertedNode;
		}

		/// <summary>
		/// Creates a attribute based response node from an element based response node.
		/// </summary>
		/// <param name="xmlNode">The element based response node.</param>
		/// <param name="recursive">Indicates whether to recursively convert child nodes.</param>
		/// <returns>
		/// The new attribute based response node.
		/// </returns>
		public static XmlNode CreateAttributeBasedResponse(XmlNode xmlNode, bool recursive) 
		{
			XmlElement xmlParent = null;

			if (xmlNode != null)
			{
				// Clone the parent node (without child nodes) as a basis for the returned node
				xmlParent = (XmlElement)xmlNode.CloneNode(false);
				// Iterate through each child node
				foreach (XmlNode xmlChildNode in xmlNode.ChildNodes)
				{
					// HasValue if it is a 'text' node or genuinely has child elements
					bool noChildren = !xmlChildNode.HasChildNodes || xmlChildNode.NodeType == XmlNodeType.Text;
					if (noChildren && (xmlChildNode.InnerText).Trim().Length > 0)
					{
						// Create as an attribute of parent
						xmlParent.SetAttribute(xmlChildNode.Name, xmlChildNode.InnerText);
						// HasValue if combo item with TEXT attribute
						foreach (XmlAttribute xmlAttribute in xmlChildNode.Attributes)
						{
							if (xmlAttribute.Name.ToUpper() == "TEXT")
							{
								// Create a text attribute also for the combo text
								xmlParent.SetAttribute(xmlChildNode.Name + "_TEXT", xmlAttribute.InnerText);
							}
						}
					}
					else if (recursive && !noChildren)
					{
						// Process child node recursively
						xmlParent.AppendChild(CreateAttributeBasedResponse(xmlChildNode, recursive));
					}
					noChildren = false;
				}
			}
			return xmlParent;
		}


		/// <summary>
		/// Creates a request node that can be used to find all matching records in a child table.
		/// </summary>
		/// <param name="xmlDataNode">The element to be converted.</param>
		/// <param name="xmlSchemaNode">The schema entry for the parent table.</param>
		/// <param name="newName">The name of the child node to be created.</param>
		/// <returns>
		/// The new request node.
		/// </returns>
		/// <remarks>
		/// Combo look up is not performed.
		/// </remarks>
		public static XmlNode CreateChildRequest(XmlElement xmlDataNode, XmlNode xmlSchemaNode, string newName)
		{
			return CreateChildRequest(xmlDataNode, xmlSchemaNode, newName, false);
		}

		/// <summary>
		/// Creates a request node that can be used to find all matching records in a child table.
		/// </summary>
		/// <param name="xmlDataNode">The element to be converted.</param>
		/// <param name="xmlSchemaNode">The schema entry for the parent table.</param>
		/// <param name="newName">The name of the child node to be created.</param>
		/// <param name="comboLookup">Indicates whether combo look up should be performed.</param>
		/// <returns>
		/// The new request node.
		/// </returns>
		public static XmlNode CreateChildRequest(XmlElement xmlDataNode, XmlNode xmlSchemaNode, string newName, bool comboLookup) 
		{
			XmlElement xmlNewNode = null;
			if (xmlDataNode != null && xmlSchemaNode != null && newName.Trim().Length > 0)
			{
				// Create a new root node
				XmlDocument xmlDocument = new XmlDocument();
				xmlNewNode = xmlDocument.CreateElement(newName);
				if (comboLookup)
				{
					xmlNewNode.SetAttribute("_COMBOLOOKUP_", "1");
				}
				// Iterate through the schema node to get the primary key items
				foreach (XmlNode xmlChildNode in xmlSchemaNode.ChildNodes)
				{
					if (xmlChildNode.Attributes.GetNamedItem("KEYTYPE") != null)
					{
						string valueText = xmlChildNode.Attributes.GetNamedItem("KEYTYPE").InnerText;
						if (valueText == "PRIMARY")
						{
							// Get this attribute from the data element and add it to the new node
							XmlAttribute xmlAttribute = (XmlAttribute)xmlDataNode.Attributes.GetNamedItem(xmlChildNode.Name);
							if (xmlAttribute != null)
							{
								xmlNewNode.SetAttribute(xmlChildNode.Name, xmlAttribute.InnerText);
							}
						}
					}
				}
			}
			else
			{
				throw new InvalidParameterException();
			}

			return xmlNewNode;
		}

		/// <summary>
		/// Creates a new element with the same name and inner text 
		/// as an attribute in a source node, and appends the element to a destination node.
		/// </summary>
		/// <param name="xmlDestinationNode">The destination node.</param>
		/// <param name="xmlSourceNode">The source node.</param>
		/// <param name="attributeName">The name of the attribute in the source node.</param>
		public static void ElemFromAttrib(XmlNode xmlDestinationNode, XmlNode xmlSourceNode, string attributeName) 
		{
			XmlElement xmlElement = xmlDestinationNode.OwnerDocument.CreateElement(attributeName);
			if (xmlSourceNode.Attributes.GetNamedItem(attributeName) != null)
			{
				xmlElement.InnerText = xmlSourceNode.Attributes.GetNamedItem(attributeName).InnerText;
			}
			xmlDestinationNode.AppendChild(xmlElement);
		}

		/// <summary>
		/// Gets the REQUEST node from an element, without any children, 
		/// using the XPath expression "//REQUEST" to find the REQUEST node.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <returns>
		/// The REQUEST node.
		/// </returns>
		public static XmlNode GetRequestNode(XmlElement xmlElement)
		{
			return GetRequestNode(xmlElement, "//REQUEST");
		}

		/// <summary>
		/// Gets the REQUEST node from an element, without any children, 
		/// using the XPath expression ".//REQUEST" to find the REQUEST node.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <returns>
		/// The REQUEST node.
		/// </returns>
		public static XmlNode GetRequestNodeEx(XmlElement xmlElement)
		{
			return GetRequestNode(xmlElement, ".//REQUEST");
		}

		/// <summary>
		/// Gets the REQUEST node from an element, without any children.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <param name="xpath">The XPath expression to use to find the REQUEST node.</param>
		/// <returns>
		/// The REQUEST node.
		/// </returns>
		private static XmlNode GetRequestNode(XmlElement xmlElement, string xpath) 
		{
			XmlNode xmlRequestNode = null;

			if (xmlElement == null)
			{
				throw new MissingElementException("Node Empty");
			}

			// If the top most tag of the node is 'REQUEST', just return the node, else
			// search for 'REQUEST'
			if (xmlElement.Name == "REQUEST")
			{
				xmlRequestNode = xmlElement.CloneNode(false);
			}
			else
			{
				XmlNode xmlNode = xmlElement.SelectSingleNode(xpath);
				if (xmlNode == null)
				{
					throw new MissingPrimaryTagException("Expected REQUEST tag");
				}
				xmlRequestNode = xmlNode.CloneNode(false);
			}
			// Attach the cloned node to a new dom document to ensure safety of using selectsinglenode
			// with something like this: "/REQUEST/SOMETHING"
			// AS 18/04/2007 Commented out as causes an ArgumentException:
			// "The node to be inserted is from a different document context.".
			//XmlDocument xmlDocument = new XmlDocument();
			//xmlDocument.AppendChild(xmlRequestNode);
			return xmlRequestNode;
		}

		#endregion xmlAssistEx

		#region XMLAssist

		/// <summary>
		/// Gets the text value for a node with a specified name by searching a specified element and all its sub-nodes. 
		/// </summary>
		/// <param name="xmlElement">The element to search.</param>
		/// <param name="nodeName">The name to search for.</param>
		/// <returns>
		/// The text value of the matching node or the empty string if no match was found.
		/// </returns>
		public static string GetTagValue(XmlElement xmlElement, string nodeName)
		{
			bool exists = false;
			return GetTagValue(xmlElement, nodeName, out exists, true);
		}

		/// <summary>
		/// Gets the text value for a node with a specified name by searching a specified element. 
		/// </summary>
		/// <param name="xmlElement">The element to search.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="exists">Indicates whether the node was found.</param>
		/// <param name="deep">Indicates whether to search sub-nodes of the specified element.</param>
		/// <returns>
		/// The text value of the matching node or the empty string if no match was found.
		/// </returns>
		public static string GetTagValue(XmlElement xmlElement, string xpath, out bool exists, bool deep)
		{
			string tagValue = "";
			exists = false;

			if (xmlElement.Name != xpath)
			{
				if (!deep)
				{
					xmlElement = (XmlElement)xmlElement.SelectSingleNode(xpath);
				}
				else
				{
					xmlElement = (XmlElement)xmlElement.SelectSingleNode(".//" + xpath);
				}
			}

			if (xmlElement != null)
			{
				exists = true;
				tagValue = xmlElement.InnerText;
			}

			return tagValue;
		}

		/// <summary>
		/// Gets the inner text for an attribute with a specified name, in a child node with a specified name.
		/// </summary>
		/// <param name="xmlElement">The element to search.</param>
		/// <param name="nodeName">The name of the child node to search for.</param>
		/// <param name="attributeName">The name of the attribute to search for.</param>
		/// <returns>
		/// The InnerText property of the attribute, or the empty string if the child node or attribute are not found.
		/// </returns>
		public static string GetAttributeValue(XmlElement xmlElement, string nodeName, string attributeName)
		{
			bool exists = false;
			return GetAttributeValue(xmlElement, nodeName, attributeName, out exists);
		}

		/// <summary>
		/// Gets the inner text for an attribute with a specified name, in a child node with a specified name.
		/// </summary>
		/// <param name="xmlElement">The element to search.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="attributeName">The name of the attribute to search for.</param>
		/// <param name="exists">Indicates whether the node was found.</param>
		/// <returns>
		/// The InnerText property of the attribute, or the empty string if the child node or attribute are not found.
		/// </returns>
		public static string GetAttributeValue(XmlElement xmlElement, string xpath, string attributeName, out bool exists)
		{
			string attributeValue = "";
			exists = false;

			if (xmlElement.Name != xpath)
			{
				xmlElement = (XmlElement)xmlElement.SelectSingleNode(".//" + xpath);
			}

			if (xmlElement != null)
			{
				XmlAttribute xmlAttribute = xmlElement.GetAttributeNode(attributeName);
				if (xmlAttribute != null)
				{
					exists = true;
					attributeValue = xmlAttribute.InnerText;
				}
			}

			return attributeValue;
		}

		/// <summary>
		/// Adds a specified element to every node with a specified name that is a descendant of a specified element.
		/// </summary>
		/// <param name="xmlElement">The element to add.</param>
		/// <param name="targetNodeName">The name of the nodes to which the element should be added.</param>
		/// <param name="deep">Indicates whether the descendants of the added element should be included.</param>
		/// <param name="xmlTargetElement">The element to search for descendants with the name <paramref name=""/>.</param>
		/// <returns>
		/// Indicates whether the operation was successful.
		/// </returns>
		public static bool AddNode(XmlElement xmlElement, string targetNodeName, bool deep, XmlElement xmlTargetElement)
		{
			// History:
			// 
			// Prog  Date	 Description
			// RF	01/07/99 Fixed to work when targetNodeName is at top level of
			// xmlTargetElement
			// ------------------------------------------------------------------------------------------

			if (xmlTargetElement == null)
			{
				throw new ArgumentException("Parameter is null", "xmlTargetElement");
			}

			if (targetNodeName == null || targetNodeName.Length == 0)
			{
				throw new ArgumentException("Parameter is null or empty string", "targetNodeName");
			}

			if (xmlTargetElement.Name == targetNodeName)
			{
				XmlNode xmlNewNode = xmlElement.CloneNode(deep);
				xmlTargetElement.AppendChild(xmlNewNode);
			}
			else
			{
				XmlNodeList xmlNodeList = xmlTargetElement.SelectNodes(".//" + targetNodeName);
				foreach (XmlNode xmlNode in xmlNodeList)
				{
					XmlNode xmlNewNode = xmlElement.CloneNode(deep);
					xmlNode.AppendChild(xmlNewNode);
				}
			}

			return true;
		}

		/// <summary>
		/// Add a child node with the name "UCASE" + specified prefix to a specified node, 
		/// where the inner text of the child node is the upper case version of the inner text of 
		/// the specified node, and thus can be used for alphabetic sorting.
		/// </summary>
		/// <param name="xmlNode">The node to which the child node should be added.</param>
		/// <param name="nodeName">The name prefix for the new node.</param>
		/// <remarks>
		/// If <paramref name="nodeName"/> is <b>SURNAME</b>, then the xml
		///	<code>
		///	&lt;CUSTOMERVERSION&gt;
		///		&lt;SURNAME&gt;Smith&lt;/SURNAME&gt;
		///	&lt;/CUSTOMERVERSION&gt;
		/// </code>
		/// will be converted to:
		///	<code>
		///	&lt;CUSTOMERVERSION&gt;
		///		&lt;SURNAME&gt;Smith&lt;/SURNAME&gt;
		///			&lt;UCASESURNAME&gt;SMITH&lt;/UCASESURNAME&gt;
		///	&lt;/CUSTOMERVERSION&gt;
		/// </code>
		/// </remarks>
		public static void AddNodeWithUCaseText(XmlNode xmlNode, string nodeName)
		{
			bool exists = false;
			string tagValue = GetTagValue((XmlElement)xmlNode, nodeName, out exists, true);
			XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("UCASE" + nodeName);
			xmlElement.InnerText = tagValue.ToUpper();
			xmlNode.AppendChild(xmlElement);
		}

		/// <summary>
		/// Gets the data nodes out of a response element and appends them to a node.
		/// </summary>
		/// <param name="xmlNodeToAttachTo">The node to which the data nodes will be attached.</param>
		/// <param name="xmlResponse">The response element.</param>
		public static void AttachResponseData(XmlNode xmlNodeToAttachTo, XmlElement xmlResponse)
		{
			if (xmlNodeToAttachTo == null || xmlResponse == null)
			{
				throw new InvalidParameterException("Node to attach to or Response missing");
			}
			if (xmlResponse.Name != "RESPONSE")
			{
				throw new InvalidParameterException("RESPONSE must be top level tag");
			}

			foreach (XmlNode xmlChildNode in xmlResponse.ChildNodes)
			{
				if (xmlChildNode.Name != "MESSAGE")
				{
					xmlNodeToAttachTo.AppendChild(xmlNodeToAttachTo.OwnerDocument.ImportNode(xmlChildNode, true));
				}
			}
		}

		/// <summary>
		/// Throws an <see cref="MissingElementException"/> exception if a node is null.
		/// </summary>
		/// <param name="xmlNode">The node.</param>
		/// <exception cref="MissingElementException">
		/// <paramref name="xmlNode"/> is null.
		/// </exception>
		public static void CheckNode(object xmlNode)
		{
			CheckNode(xmlNode, "Node Empty");
		}

		/// <summary>
		/// Throws an <see cref="MissingElementException"/> exception if a node is null.
		/// </summary>
		/// <param name="xmlNode">The node.</param>
		/// <param name="errorText">The message to include in the exception.</param>
		/// <exception cref="MissingElementException">
		/// <paramref name="xmlNode"/> is null.
		/// </exception>
		public static void CheckNode(object xmlNode, string errorText)
		{
			if (xmlNode == null)
			{
				throw new MissingElementException(errorText);
			}
		}

		/// <summary>
		/// Copies a node from a specified source element to a specified target element, including all child nodes.
		/// </summary>
		/// <param name="sourceNodeName">The name of the source node.</param>
		/// <param name="targetNodeName">The name of the target node.</param>
		/// <param name="xmlSourceElement">The element containing the source node.</param>
		/// <param name="xmlTargetElement">The element containing the target node.</param>
		/// <param name="deep">Indicates whether to recursively copy child nodes.</param>
		/// <returns>
		/// Indicates whether the operation was successful.
		/// </returns>
		/// <remarks>
		/// The child nodes of the source node are always copied, but deeper descendants are only copied if 
		/// <paramref name="deep"/> is <b>true</b>.
		/// </remarks>
		public static bool CopyNode(string sourceNodeName, string targetNodeName, XmlElement xmlSourceElement, XmlElement xmlTargetElement, bool deep)
		{
			// SR 19-01-00
			// if sourceNodeName is at top level of xmlSourceElement, add all the children
			// to the target
			if (xmlSourceElement.Name == sourceNodeName)
			{
				XmlNode xmlNode = xmlTargetElement.OwnerDocument.CreateElement(targetNodeName);
				XmlNode xmlNewNode = xmlTargetElement.AppendChild(xmlNode);
				foreach (XmlNode xmlChildNode in xmlSourceElement.ChildNodes)
				{
					xmlNewNode.AppendChild(xmlChildNode.CloneNode(true));
				}
			}
			else
			{
				// obtain the list of nodes matching the sourceNodeName
				XmlNodeList xmlNodeList = xmlSourceElement.SelectNodes(".//" + sourceNodeName);
				// for every occurrence of a node matching the sourceNodeName create a
				// new node on the target and copy its children
				// SA - Core AQR : SYS4754
				foreach (XmlNode xmlChildNode in xmlNodeList)
				{
					XmlNode xmlNode = xmlTargetElement.OwnerDocument.CreateElement(targetNodeName);
					if (xmlNode != null)
					{
						XmlNode xmlNewNode = xmlTargetElement.AppendChild(xmlNode);
						foreach (XmlNode xmlChildNode2 in xmlChildNode.ChildNodes)
						{
							xmlNode = xmlNewNode.AppendChild(xmlChildNode2.CloneNode(deep));
						}
					}
				}
			}

			return true;
		}

		/// <summary>
		/// Gets the number of child nodes of a specified element whose inner text is not null or the empty string.
		/// </summary>
		/// <param name="xmlElement">The element whose child nodes should be counted.</param>
		/// <returns>
		/// The number of child nodes of <paramref name="xmlElement"/> whose InnerText property is not null or the empty string.
		/// </returns>
		public static int CountNonNullNodes(XmlElement xmlElement)
		{
			int nodes = 0;
			foreach (XmlNode xmlNode in xmlElement.ChildNodes)
			{
				if (xmlNode.InnerText != null && xmlNode.InnerText.Trim().Length > 0)
				{
					nodes++;
				}
			}

			return nodes;
		}

		/// <summary>
		/// Searches a specified element for child nodes that match a specified XPath expression, 
		/// and appends each matching child node to a new "LIST" node below the specified element.
		/// </summary>
		/// <param name="xmlElement">The element to be searched.</param>
		/// <param name="xpath">The XPath expression.</param>
		public static void GroupNodesIntoList(XmlElement xmlElement, string xpath)
		{
			XmlNodeList xmlNodeList = xmlElement.SelectNodes(xpath);
			if (xmlNodeList.Count > 0)
			{
				XmlNode xmlListNode = xmlElement.AppendChild(xmlElement.OwnerDocument.CreateElement(xpath + "LIST"));
				// SA Core AQR : SYS4754
				// Move each node in the node list under the new LIST tag
				foreach (XmlNode xmlNode in xmlNodeList)
				{
					xmlListNode.AppendChild(xmlNode);
				}
			}
		}

		/// <summary>
		/// Removes duplicate descendant nodes from a specified element.
		/// </summary>
		/// <param name="xmlElement">The element from which duplicate nodes should be removed.</param>
		/// <remarks>
		/// For all descendant nodes of <paramref name="xmlElement"/>, 
		/// if the InnerText property is the same for two matching descendant nodes, the second node is removed.
		/// If the second node has children of its own its data is saved by moving the data across
		/// to the first node.	
		/// </remarks>
		public static void RemoveDuplicates(XmlElement xmlElement)
		{
			RemoveDuplicates(xmlElement, "");
		}

		/// <summary>
		/// Removes duplicate descendant nodes matching a XPath expression from a specified element.
		/// </summary>
		/// <param name="xmlElement">The element from which duplicate nodes should be removed.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <remarks>
		/// The XPath expression is used to find descendant nodes of <paramref name="xmlElement"/>. 
		/// If the InnerText property is the same for two matching descendant nodes, the second node is removed.
		/// If the second node has children of its own its data is saved by moving the data across
		/// to the first node.	
		/// </remarks>
		public static void RemoveDuplicates(XmlElement xmlElement, string xpath)
		{
			for (int i = 0; i < xmlElement.ChildNodes.Count; i++)
			{
				XmlNode xmlThisNode = xmlElement.ChildNodes[i];
				int j = i + 1;
				while (j < xmlElement.ChildNodes.Count)
				{
					XmlNode xmlTestNode = xmlElement.ChildNodes[j];
					if (NodesMatch(xmlThisNode, xmlTestNode, xpath))
					{
						// The IDs of the node and the test node match
						if (xpath != null && xpath.Length > 0)
						{
							// SA - Core AQR : SYS4754
							// Copy the contents of the test node to the current node.
							foreach (XmlNode xmlNode in xmlTestNode.ChildNodes)
							{
								if (xmlNode.ChildNodes.Count > 1)
								{
									xmlThisNode.AppendChild(xmlNode.CloneNode(true));
								}
							}
						}
						// Contents of the test node have now been copied, so destroy it
						xmlElement.RemoveChild(xmlTestNode);
					}
					else
					{
						j++;
					}
				}
			}
		}

		/// <summary>
		/// Indicates whether two specified nodes match each other, 
		/// by comparing the inner text of each corresponding descendant node.
		/// </summary>
		/// <param name="xmlNode1">The first node to compare.</param>
		/// <param name="xmlNode2">The second node to compare.</param>
		/// <returns>
		/// <b>true</b> if the two nodes match, otherwise <b>false</b>.
		/// </returns>
		public static bool NodesMatch(XmlNode xmlNode1, XmlNode xmlNode2)
		{
			return NodesMatch(xmlNode1, xmlNode2, "");
		}

		/// <summary>
		/// Indicates whether two specified nodes match each other, 
		/// by comparing the inner text of descendant nodes that match an XPath expression.
		/// </summary>
		/// <param name="xmlNode1">The first node to compare.</param>
		/// <param name="xmlNode2">The second node to compare.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <returns>
		/// <b>true</b> if the two nodes match, otherwise <b>false</b>.
		/// </returns>
		public static bool NodesMatch(XmlNode xmlNode1, XmlNode xmlNode2, string xpath)
		{
			bool match = true;
			if (xpath == null || xpath.Length == 0)
			{
				// Compare the entire nodes
				if (xmlNode1.HasChildNodes && xmlNode1.ChildNodes.Count == xmlNode2.ChildNodes.Count && xmlNode1.Name == xmlNode2.Name)
				{
					// Both nodes have the same number of children and the same tag names
					for (int i = 0; i < xmlNode1.ChildNodes.Count; i++)
					{
						match &= NodesMatch(xmlNode1.ChildNodes[i], xmlNode2.ChildNodes[i], "");
					}
				}
				else if (!(xmlNode1.HasChildNodes || xmlNode2.HasChildNodes) && xmlNode1.Name == xmlNode2.Name)
				{
					// Nodes do not have children (i.e. are 'leaf' nodes), so compare the text of both nodes instead
					match = xmlNode1.InnerText == xmlNode2.InnerText;
				}
				else
				{
					match = false;
				}
			}
			else
			{
				// Compare the specified tag in the nodes
				XmlNode xmlValue1 = xmlNode1.SelectSingleNode(xpath);
				XmlNode xmlValue2 = xmlNode2.SelectSingleNode(xpath);
				if (xmlValue1 == null || xmlValue2 == null)
				{
					match = false;
				}
				else
				{
					match = xmlValue1.InnerText == xmlValue2.InnerText;
				}
			}

			return match;
		}

		/// <summary>
		/// For a specified element attempts to group all specified child nodes under a new child node.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <param name="groupName">The name for the new child node.</param>
		/// <param name="fields">A comma separated list of child node names.</param>
		/// <remarks>
		/// For each child node of <paramref name="xmlElement"/>, if its name appears in <paramref name="fields"/>
		/// then the child node is moved to become a child of a new child node named <paramref name="groupName"/>. 
		/// For example, if the following element is passed in:
		/// <code>
		/// &lt;TABLE&gt;
		///		&lt;FIELD1&gt;text1&lt;/FIELD1&gt;
		///	 &lt;FIELD2&gt;text2&lt;/FIELD2&gt;
		///	 &lt;FIELD3&gt;text3&lt;/FIELD3&gt;
		/// &lt;/TABLE&gt;
		/// </code>
		/// and <paramref name="groupName"/> is set to "MYGROUP", and <paramref name="fields"/> is set to "FIELD2,FIELD3", 
		/// then the resulting element would be:
		/// <code>
		/// &lt;TABLE&gt;
		///	 &lt;FIELD1&gt;text1&lt;/FIELD1&gt;
		///	 &lt;MYGROUP&gt;
		///		 &lt;FIELD2&gt;text2&lt;/FIELD2&gt;
		///		 &lt;FIELD3&gt;text3&lt;/FIELD3&gt;
		///	 &lt;/MYGROUP&gt;
		/// &lt;/TABLE&gt;
		/// </code>
		/// </remarks>
		public static void GroupNodes(XmlElement xmlElement, string groupName, string fields)
		{
			GroupNodes(xmlElement, groupName, fields, true);
		}

		/// <summary>
		/// For a specified element attempts to group all specified child nodes under a new child node.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <param name="groupName">The name for the new child node.</param>
		/// <param name="fields">A comma separated list of child node names.</param>
		/// <remarks>
		/// For each child node of <paramref name="xmlElement"/>, if its name appears in <paramref name="fields"/>
		/// then the child node is moved to become a child of a new child node named <paramref name="groupName"/>. 
		/// For example, if the following element is passed in:
		/// <code>
		/// &lt;TABLE&gt;
		///		&lt;FIELD1&gt;text1&lt;/FIELD1&gt;
		///	 &lt;FIELD2&gt;text2&lt;/FIELD2&gt;
		///	 &lt;FIELD3&gt;text3&lt;/FIELD3&gt;
		/// &lt;/TABLE&gt;
		/// </code>
		/// and <paramref name="groupName"/> is set to "MYGROUP", and <paramref name="fields"/> is set to "FIELD2,FIELD3", 
		/// then the resulting element would be:
		/// <code>
		/// &lt;TABLE&gt;
		///	 &lt;FIELD1&gt;text1&lt;/FIELD1&gt;
		///	 &lt;MYGROUP&gt;
		///		 &lt;FIELD2&gt;text2&lt;/FIELD2&gt;
		///		 &lt;FIELD3&gt;text3&lt;/FIELD3&gt;
		///	 &lt;/MYGROUP&gt;
		/// &lt;/TABLE&gt;
		/// </code>
		/// <para>
		/// This method differs from <see cref="GroupNodes"/> as the former creates a new <see cref="XmlDocument"/> to 
		/// create the new group element, whereas the latter uses the existing <see cref="XmlDocument"/>. Functionally the two 
		/// methods are equivalent.
		/// </para>
		/// </remarks>
		public static void GroupNodesEx(XmlElement xmlElement, string groupName, string fields)
		{
			GroupNodes(xmlElement, groupName, fields, false);
		}

		private static void GroupNodes(XmlElement xmlElement, string groupName, string fields, bool createNewDocument)
		{
			XmlElement xmlGroupElement = null;

			fields = "," + fields + ",";
			bool createdGroupTag = false;
			int i = 0;
			while (i < xmlElement.ChildNodes.Count)
			{
				XmlNode xmlThisNode = xmlElement.ChildNodes[i];
				if (fields.IndexOf("," + xmlThisNode.Name + ",", StringComparison.OrdinalIgnoreCase) != -1)
				{
					// The field represented by the tag is a member of the fields string so move the tag to
					// the specified grouping element
					if (!createdGroupTag)
					{
						if (createNewDocument)
						{
							XmlDocument xmlDocument = new XmlDocument();
							xmlGroupElement = (XmlElement)xmlElement.AppendChild(xmlDocument.CreateElement(groupName));
						}
						else
						{
							xmlGroupElement = (XmlElement)xmlElement.AppendChild(xmlElement.OwnerDocument.CreateElement(groupName));
						}
						createdGroupTag = true;
					}
					xmlGroupElement.AppendChild(xmlThisNode.CloneNode(true));
					xmlElement.RemoveChild(xmlThisNode);
				}
				else
				{
					i++;
				}
			}
		}

		/// <summary>
		/// Generates the xml request to pass to routines in the ComboDO such as GetComboValue.
		/// </summary>
		/// <param name="groupName">The combo group name.</param>
		/// <param name="valueText">The value for the VALUEID element.</param>
		/// <returns>The xml request.</returns>
		public static string BuildComboValueList(string groupName, string valueText)
		{
			XmlDocument xmlDocument = new XmlDocument();

			XmlElement xmlListElement = xmlDocument.CreateElement("LIST");
			xmlDocument.AppendChild(xmlListElement);
			XmlElement xmlElement = xmlDocument.CreateElement("GROUPNAME");
			xmlElement.InnerText = groupName;
			xmlListElement.AppendChild(xmlElement);
			xmlElement = xmlDocument.CreateElement("VALUEID");
			xmlElement.InnerText = valueText;
			xmlListElement.AppendChild(xmlElement);

			return xmlDocument.OuterXml;
		}

		/// <summary>
		/// Determines the appropriate operation to conduct on the passed xml request given
		/// the primary key field values.
		/// </summary>
		/// <param name="request">The xml request.</param>
		/// <param name="classDefinition">The class definition.</param>
		/// <returns>
		/// The <see cref="BOOPERATIONTYPE"/> value.
		/// </returns>
		/// <remarks>
		/// <para>
		/// If the value specified for the final key field in the paramarray is blank then
		/// BOOPERATIONTYPE.booInsert is returned.
		/// </para>
		/// <para>
		/// If all values for the key fields in the paramarray are non-blank and <i>all values 
		/// for non-key fields are blank</i> then BOOPERATIONTYPE.booDelete is returned.
		/// </para>
		/// <para>
		/// If all values for the key fields in the paramarray are non-blank and at least one
		/// value for a non-key field is non-blank then BOOPERATIONTYPE.booUpdate is returned.
		/// </para>
		/// </remarks>
		public static BOOPERATIONTYPE DetermineOperation(string request, string classDefinition)
		{
			XmlElement xmlRootElement = null;
			string tableName = "";
			BOOPERATIONTYPE botReturn = BOOPERATIONTYPE.booNone;

			// 
			// Initialise
			// 
			XmlDocument xmlIn = Load(request);
			XmlDocument xmlClassDefinition = Load(classDefinition);
			// 
			// Get table name
			// 
			XmlNodeList xmlValueNodeList = xmlClassDefinition.SelectNodes("TABLENAME");
			if (xmlValueNodeList.Count == 0)
			{
				throw new MissingTableNameException("TABLENAME tag not found");
			}
			else
			{
				tableName = xmlValueNodeList[0].FirstChild.InnerText;
			}
			// 
			// Find table tag in the request XML
			// 
			XmlNodeList xmlNodeList = xmlIn.SelectNodes(".//" + tableName);
			if (xmlNodeList.Count == 0)
			{
				return botReturn;
			}
			else
			{
				xmlRootElement = (XmlElement)xmlNodeList[0];
			}
			// 
			// Get the primary key nodes from the class definition
			// 
			XmlNodeList xmlKeyNodeList = xmlClassDefinition.SelectNodes(".//PRIMARYKEY");
			if (xmlKeyNodeList.Count == 0)
			{
				throw new MissingPrimaryTagException("No keys found");
			}
			// 
			// Ascertain values of primary key fields
			// 
			bool completeKey = true;
			// SA Core AQR : SYS4754
			// ++ SA SYS4754 code review
			foreach (XmlNode xmlKeyNode in xmlKeyNodeList)
			{
				string valueText = "";
				// Get node in the data XML corresponding to this key
				xmlValueNodeList = xmlRootElement.SelectNodes(".//" + xmlKeyNode.FirstChild.InnerText);
				if (xmlValueNodeList.Count != 0)
				{
					XmlNode xmlValueNode = xmlValueNodeList[0];
					valueText = xmlValueNode.InnerText;
					xmlRootElement.RemoveChild(xmlValueNode);
				}
				completeKey = completeKey && valueText.Length > 0;
			}
			// Determine operation
			// 
			bool hasNonPrimaryValues = !AllChildNodesAreNull(xmlRootElement);
			if (!completeKey && hasNonPrimaryValues)
			{
				// INSERT
				botReturn = BOOPERATIONTYPE.booCreate;
			}
			else if (completeKey && hasNonPrimaryValues)
			{
				// UPDATE
				botReturn = BOOPERATIONTYPE.booUpdate;
			}
			else if (completeKey && !hasNonPrimaryValues)
			{
				// DELETE
				botReturn = BOOPERATIONTYPE.booDelete;
			}
			else
			{
				// Could not ascertain process to perform
				botReturn = BOOPERATIONTYPE.booNone;
			}
			return botReturn;
		}

		/// <summary>
		/// Determines the appropriate operation to conduct on the passed xml request given
		/// the primary key field values.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <returns>
		/// The <see cref="BOOPERATIONTYPE"/> value.
		/// </returns>
		/// <remarks>
		/// <para>
		/// If the value specified for the final key field in the paramarray is blank then
		/// BOOPERATIONTYPE.booInsert is returned.
		/// </para>
		/// <para>
		/// If all values for the key fields in the paramarray are non-blank and <i>all values 
		/// for non-key fields are blank</i> then BOOPERATIONTYPE.booDelete is returned.
		/// </para>
		/// <para>
		/// If all values for the key fields in the paramarray are non-blank and at least one
		/// value for a non-key field is non-blank then BOOPERATIONTYPE.booUpdate is returned.
		/// </para>
		/// </remarks>
		public static BOOPERATIONTYPE DetermineOperationEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition)
		{
			XmlElement xmlRootElement = null;
			string tableName = "";
			BOOPERATIONTYPE botReturn = BOOPERATIONTYPE.booNone;

			// 
			// Get table name
			// 
			XmlNode xmlValueNode = xmlClassDefinition.SelectSingleNode("TABLENAME");
			if (xmlValueNode == null)
			{
				throw new MissingTableNameException("TABLENAME tag not found");
			}
			else
			{
				tableName = xmlValueNode.FirstChild.InnerText;
			}
			// 
			// Find table tag in the request XML
			// 
			if (xmlRequest.Name == tableName)
			{
				xmlRootElement = xmlRequest;
			}
			else
			{
				xmlRootElement = (XmlElement)xmlRequest.SelectSingleNode(".//" + tableName);
			}
			if (xmlRootElement == null)
			{
				return botReturn;
			}
			// 
			// Get the primary key nodes from the class definition
			// 
			XmlNodeList xmlKeyNodeList = xmlClassDefinition.DocumentElement.SelectNodes("PRIMARYKEY");
			if (xmlKeyNodeList.Count == 0)
			{
				throw new MissingPrimaryTagException("No keys found");
			}
			// 
			// Ascertain values of primary key fields
			// 
			bool completeKey = true;
			// SA Core AQR : SYS4754
			foreach (XmlNode xmlKeyNode in xmlKeyNodeList)
			{
				string valueText = "";
				// Get node in the data XML corresponding to this key
				xmlValueNode = xmlRootElement.SelectSingleNode(xmlKeyNode.FirstChild.InnerText);
				if (xmlValueNode != null)
				{
					valueText = xmlValueNode.InnerText;
				}
				completeKey = completeKey && valueText.Length > 0;
			}
			// 
			// Determine operation
			// 
			bool blnHasNonKeyValues = !AreAllNonKeyValuesNullEx(xmlRootElement, xmlClassDefinition);
			if (!completeKey)
			{
				// INSERT
				botReturn = BOOPERATIONTYPE.booCreate;
			}
			else if (completeKey && blnHasNonKeyValues)
			{
				// UPDATE
				botReturn = BOOPERATIONTYPE.booUpdate;
			}
			else if (completeKey && !blnHasNonKeyValues)
			{
				// DELETE
				botReturn = BOOPERATIONTYPE.booDelete;
			}
			else
			{
				// Could not ascertain process to perform
				botReturn = BOOPERATIONTYPE.booNone;
			}

			return botReturn;
		}

		/// <summary>
		/// Indicates whether all child nodes for a specified element are null values.
		/// </summary>
		/// <param name="xmlElement">The element.</param>
		/// <returns>
		/// <b>true</b> if all child nodes are null, otherwise <b>false</b>.
		/// </returns>
		public static bool AllChildNodesAreNull(XmlElement xmlElement)
		{
			bool allChildNodesAreNull = true;
			foreach (XmlNode xmlNode in xmlElement.ChildNodes)
			{
				if (xmlNode.InnerText != null && xmlNode.InnerText.Trim().Length > 0)
				{
					// Found a non-null node value
					allChildNodesAreNull = false;
					break;
				}
			}

			return allChildNodesAreNull;
		}

		/// <summary>
		/// Indicates whether all non-key values are null.
		/// </summary>
		/// <param name="request">The xml request.</param>
		/// <param name="classDefinition">The class defintion.</param>
		/// <returns>
		/// <b>true</b> if all non-key values are null, otherwise <b>false</b>.
		/// </returns>
		public static bool AreAllNonKeyValuesNull(string request, string classDefinition)
		{
			// Load Class Definition and the request
			XmlDocument xmlIn = XmlAssist.Load(request);
			XmlDocument xmlClassDefinition = XmlAssist.Load(classDefinition);
			return AreAllNonKeyValuesNullEx(xmlIn.DocumentElement, xmlClassDefinition);
		}

		/// <summary>
		/// Indicates whether all non-key values are null.
		/// </summary>
		/// <param name="xmlRequest">The xml request.</param>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <returns>
		/// <b>true</b> if all non-key values are null, otherwise <b>false</b>.
		/// </returns>
		public static bool AreAllNonKeyValuesNullEx(XmlElement xmlRequest, XmlDocument xmlClassDefinition) 
		{
			XmlNodeList xmlNonPKNodeList = xmlClassDefinition.SelectNodes(".//OTHERS");
			bool allNonKeyValuesAreNullEx = true;
			foreach (XmlNode xmlNonPKNode in xmlNonPKNodeList)
			{
				// Get node in the data XML corresponding to this key
				string nodeName = xmlNonPKNode.FirstChild.InnerText;
				if (GetTagValue(xmlRequest, nodeName).Length != 0)
				{
					allNonKeyValuesAreNullEx = false;
					break;
				}
			}

			return allNonKeyValuesAreNullEx;
		}

		/// <summary>
		/// Searches a specified node for a child node, and promotes the child node up on level in 
		/// the xml hierarchy.
		/// </summary>
		/// <param name="xmlData"></param>
		/// <param name="xmlNodeToPromote"></param>
		/// <remarks>
		/// The child node replaces its original parent node, 
		/// and inherits all the nodes that were its siblings before the promotion as its children.
		/// </remarks>
		public static void PromoteNode(XmlNode xmlData, XmlNode xmlNodeToPromote)
		{
			PromoteNode(xmlData, xmlNodeToPromote, true, true);
		}

		/// <summary>
		/// Searches a specified node for a child node, and promotes the child node up on level in 
		/// the xml hierarchy.
		/// </summary>
		/// <param name="xmlData">The node.</param>
		/// <param name="xmlNodeToPromote">The child node to promote.</param>
		/// <param name="replaceParent">Indicates whether the child node should replace its original parent node.</param>
		/// <param name="inheritSiblings">Indicates whether the child node will inherit all the nodes that were its siblings before the promotion as its children.</param>
		public static void PromoteNode(XmlNode xmlData, XmlNode xmlNodeToPromote, bool replaceParent, bool inheritSiblings)
		{
			// 
			// Initialise
			// 
			if (xmlNodeToPromote == null)
			{
				return;
			}
			if (xmlNodeToPromote.ParentNode == null || xmlNodeToPromote == xmlNodeToPromote.OwnerDocument)
			{
				return;
			}
			// Find parent node of the candidate node to promote
			XmlNode xmlParentNode = xmlNodeToPromote.ParentNode;
			if (xmlParentNode == null)
			{
				return;
			}
			// 
			// Inherit siblings if necessary
			// 
			if (inheritSiblings)
			{
				// Inherit all the siblings
				foreach (XmlNode xmlSiblingNode in xmlParentNode.ChildNodes)
				{
					if (xmlSiblingNode != xmlNodeToPromote)
					{
						xmlNodeToPromote.AppendChild(xmlSiblingNode);
					}
				}
			}
			// 
			// Promote the node, including the replacement of the parent node if necessary
			// 
			XmlNode xmlGrandparentNode = xmlParentNode.ParentNode;
			if (xmlGrandparentNode == null || xmlParentNode == xmlParentNode.OwnerDocument || xmlGrandparentNode == xmlParentNode.OwnerDocument)
			{
				// Parent node is the root node of the document
				if (replaceParent)
				{
					// Parent node is the root of the document so point the passed XML pointer to the promoted node itself
					xmlData = xmlNodeToPromote;
				}
				else
				{
					// Do nothing - since the parent node is the root of the document the node to be promoted cannot be moved
					// up to the same level in the XML hierarchy as the parent as that would mean that there were two root nodes
				}
			}
			else
			{
				if (replaceParent)
				{
					// Remove the original parent
					xmlGrandparentNode.RemoveChild(xmlParentNode);
				}
				else
				{
					// Do nothing - the node to be promoted can simply be moved up to be a sibling of its parent node
				}
				// Promote this node
				xmlGrandparentNode.AppendChild(xmlNodeToPromote);
			}
		}

		/// <summary>
		/// For each child node of a specified node, removes the child if the text of its
		/// first child is not in a array of permitted values.
		/// </summary>
		/// <param name="varList">The array of permitted values.</param>
		/// <param name="xmlTableNode">The node.</param>
		public static void SelectSchemaFields(string[] varList, XmlNode xmlTableNode)
		{
			XmlNode xmlNode = xmlTableNode.FirstChild;
			while (xmlNode != null)
			{
				XmlNode xmlNextSibling = xmlNode.NextSibling;

				if (xmlNode.NodeType == XmlNodeType.Element)
				{
					bool match = false;
					for (int i = 0; i < varList.Length; i++)
					{
						if (varList[i] == xmlNode.FirstChild.InnerText)
						{
							match = true;
							break;
						}
					}
					if (!match)
					{
						xmlTableNode.RemoveChild(xmlNode);
					}
				}
				xmlNode = xmlNextSibling;
			}
		}

		/// <summary>
		/// Searches a specified node for a descendant node matching a XPath expression, and sets the 
		/// descendant node inner text to a specified value; if the descendant node does not exist it is 
		/// created.
		/// </summary>
		/// <param name="xmlNode">The node to search.</param>
		/// <param name="xpath">The XPath expression.</param>
		/// <param name="text">The value for the inner text of the matching descendant node.</param>
		public static void SetMandatoryChildText(XmlNode xmlNode, string xpath, string text)
		{
			XmlNode xmlChildNode = xmlNode.SelectSingleNode(xpath);
			if (xmlChildNode == null)
			{
				xmlChildNode = xmlNode.OwnerDocument.CreateElement(xpath);
				xmlNode.AppendChild(xmlChildNode);
			}
			xmlChildNode.InnerText = text;
		}

		/// <summary>
		/// Finds within a specific class definition a node with a specified inner text.
		/// </summary>
		/// <param name="xmlClassDefinition">The class definition.</param>
		/// <param name="nodeValue">The inner text for which to search.</param>
		/// <returns>
		/// The matching node.
		/// </returns>
		public static XmlNode GetNodeFromClassDefByNodeValue(XmlDocument xmlClassDefinition, string nodeValue)
		{
			// First check whether the required node is part of Primary Key
			XmlNodeList xmlNodeList = xmlClassDefinition.DocumentElement.SelectNodes("PRIMARYKEY");
			// SA - Core AQR: SYS4754
			foreach (XmlNode xmlNode in xmlNodeList)
			{
				if (xmlNode.FirstChild.InnerText == nodeValue)
				{
					return xmlNode;
				}
			}
			// check whether the required node is amongst others
			xmlNodeList = xmlClassDefinition.DocumentElement.SelectNodes("OTHERS");
			// SA - Core AQR: SYS4754
			foreach (XmlNode xmlNode in xmlNodeList)
			{
				if (xmlNode.FirstChild.InnerText == nodeValue)
				{
					return xmlNode;
				}
			}

			return null;
		}

		/// <summary>
		/// Starts xml logging.
		/// </summary>
		public static void StartXMLLogging()
		{
			SetXMLLoggingFlag(true);
		}

		/// <summary>
		/// Stops xml logging.
		/// </summary>
		public static void StopXMLLogging()
		{
			SetXMLLoggingFlag(false);
		}

		private static void SetXMLLoggingFlag(bool loggingOn)
		{
			GlobalProperty.SetSharedPropertyValue("XmlLogging", loggingOn);
		}

		private static bool IsXmlLoggingOn()
		{
			return Convert.ToBoolean(GlobalProperty.GetSharedPropertyValue("XmlLogging"));
		}

		#endregion XMLAssist


		#region Obsolete

		#pragma warning disable 1591

		[Obsolete("Use XmlAssist.Load instead")]
		public static XmlDocument load(string xmlText)
		{
			return Load(xmlText);
		}

		#pragma warning restore 1591

		#endregion
	}
}

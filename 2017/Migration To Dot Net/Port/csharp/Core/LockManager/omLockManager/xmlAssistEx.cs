using System;
using MSXML2;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility.VB6;
namespace omLockManager
{
    public class xmlAssistEx
    {

        // header ----------------------------------------------------------------------------------
        // Workfile:      xmlAssist.bas
        // Copyright:     Copyright © 2001 Marlborough Stirling
        // 
        // Description:   Helper functions for XML handling
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // History:
        // 
        // Prog    Date        Description
        // IK      01/11/00    Initial creation
        // ASm     15/01/01    SYS1824: Rationalisation of methods following meeting with PC and IK
        // LD      30/01/01    Add optional default value to xmlGetAttribute*
        // PSC     22/02/01    Add xmlGetRequestNode
        // SR      05/03/01    New methods 'xmlSetSysDateToNodeAttrib' and 'xmlSetSysDateToNodeListAttribs'
        // PSC     10/04/01    Add extra parameter to the Check/Get mandatory attribut/node methods to
        // enable the default message to be overridden
        // SR      13/06/01    SYS2362 - modified method 'xmlMakeNodeElementBased'. Create attrib for
        // combo descriptions
        // SR      12/07/01    SYS2412 : new method 'xmlChangeNodeName'
        // AW      02/02/04    Made 'xmlParserError' Public in line with Baseline changes
        // 
        // ------------------------------------------------------------------------------------------
        // ------------------------------------------------------------------------------------------
        // BMIDS Specific History:
        // 
        // Prog  Date        AQR         Description
        // MV    12/08/2002  BMIDS00323  Core AQR: SYS4754 - Performance. Replace all For...Each...Next with For...Next
        // Modified AttachResponseData(),xmlCreateChildRequest(),xmlCreateAttributeBasedResponse(),
        // xmlMakeNodeElementBased(),xmlCreateElementRequestFromNode(),
        // xmlChangeNodeName(),xmlSetSysDateToNodeListAttribs()
        // ------------------------------------------------------------------------------------------
        // ------------------------------------------------------------------------------------------
        // BBG Specific History:
        // 
        // Prog   Date        AQR         Description
        // MV     23/02/2004  BBG74       Created a new method xmlCreateDOMObject
        // TK     14/01/2005  BBG1891     Commented out all unused objects and set to nothing after using objects.
        // ------------------------------------------------------------------------------------------
        private const int NO_MESSAGE_NUMBER = -1;
        public static void xmlSetSysDateToNodeAttrib(MSXML2.IXMLDOMNode vxmlNode, string strDateAttribName, bool blnOverWrite) 
        {
            // ----------------------------------------------------------------------------------
            // Description :
            // Set system date to the attrib. Overwrite the existing value based on the optional
            // parameter passed in
            // Pass:
            // vxmlNode           : Node to which the required attrib is attached
            // strDateAttribName  : Name of the attribute to which the value is to be set
            // blnOverWrite       : if Yes, OverWrite the existing value(if any), else no.
            // ----------------------------------------------------------------------------------
            if (Strings.Len(xmlGetAttributeText(vxmlNode, strDateAttribName, "")) > 0)
            {
                if (blnOverWrite)
                {
                    xmlSetAttributeValue(vxmlNode, strDateAttribName, Support.Format(DateTime.Now, "dd/mm/yyyy hh:mm:ss", (FirstDayOfWeek)1, (FirstWeekOfYear)1));
                }
            }
            else
            {
                xmlSetAttributeValue(vxmlNode, strDateAttribName, Support.Format(DateTime.Now, "dd/mm/yyyy hh:mm:ss", (FirstDayOfWeek)1, (FirstWeekOfYear)1));
            }
        }

        public static void xmlSetSysDateToNodeListAttribs(MSXML2.IXMLDOMNodeList vxmlNodeList, string strDateAttribName, bool blnOverWrite) 
        {
            int lngNodeCount = 0;
            int lngListLength = 0;
            // ----------------------------------------------------------------------------------
            // Description :
            // Set system date to the attrib. Overwrite the existing value based on the optional
            // parameter passed in
            // Pass:
            // vxmlNode           : Node to which the required attrib is attached
            // strDateAttribName  : Name of the attribute to which the value is to be set
            // blnOverWrite       : if Yes, OverWrite the existing value(if any), else no.
            // ----------------------------------------------------------------------------------
            lngListLength = vxmlNodeList.length - 1;
            for(lngNodeCount = 0; lngNodeCount <= lngListLength; lngNodeCount += 1)
            {
                xmlSetSysDateToNodeAttrib(vxmlNodeList.item[lngNodeCount], strDateAttribName, false);
            }
        }

        public static void xmlParserError(MSXML2.IXMLDOMParseError vobjParseError, string vstrCallingFunction) 
        {
            string strParserError = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Raise XML parser error
            // pass:         vobjParseError parser error
            // vstrCallingFunction name of the original calling function
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            strParserError = FormatParserError(vobjParseError);
            errAssistEx.errThrowError(vstrCallingFunction, (short)(OMIGAERROR.oeXMLParserError), );

        }

        private static string FormatParserError(MSXML2.IXMLDOMParseError objParseError) 
        {
            string strErrDesc = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Format parser error into user friendly error information
            // pass:         vobjParseError parser error
            // return:       Parser error converted to a string
            // ------------------------------------------------------------------------------------------
            // formatted parser error
            // TODO DOTNETPORT
            // strErrDesc = _
            // '        "XML parser error - " & vbCr & _
            // '        "Reason: " & objParseError.reason & vbCr & _
            // '        "Error code: " & str$(objParseError.errorCode) & vbCr & _
            // '        "Line no.: " & str$(objParseError.Line) & vbCr & _
            // '        "Character: " & str$(objParseError.linepos) & vbCr & _
            // '        "Source text: " & objParseError.srcText
            return strErrDesc;
        }

        public static MSXML2.FreeThreadedDOMDocument40 xmlLoad(string vstrXMLData, string vstrCallingFunction) 
        {
            string strFunctionName = String.Empty;
            MSXML2.FreeThreadedDOMDocument40 objXmlDoc = new MSXML2.FreeThreadedDOMDocument40();
            // header ----------------------------------------------------------------------------------
            // description:  Creates and loads an XML Document from a string
            // pass:         vstrXMLData as xml data stream to be loaded
            // vstrCallingFunction which is the name of the calling function
            // return:       XML DOM Document of the vstrXMLData string
            // ------------------------------------------------------------------------------------------
            strFunctionName = "xmlLoad";
            objXmlDoc.async = false;
            // wait until XML is fully loaded
            objXmlDoc.setProperty("NewParser", true);
            objXmlDoc.validateOnParse = false;
            objXmlDoc.loadXML(vstrXMLData);
            if (objXmlDoc.parseError.errorCode != 0)
            {
                xmlParserError(objXmlDoc.parseError, vstrCallingFunction);
            }
            objXmlDoc = null;
            return objXmlDoc;
        }

        private static MSXML2.IXMLDOMNode GetNode(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, bool vblnMandatoryNode, int vlngMessageNo) 
        {
            const string cstrFunctionName = "xmlGetNode";
            MSXML2.IXMLDOMNode xmlNode = null;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vblnMandatoryNode   Optional -
            // if true, an error is raised if the pattern is not met.
            // If false, nothing is returned.
            // return:       IXMLDOMNode         The node which matches the search pattern
            // ------------------------------------------------------------------------------------------
            if (vxmlParentNode == null)
            {
                errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeMissingElement), "Missing parent node");
            }
            xmlNode = vxmlParentNode.selectSingleNode(vstrMatchPattern);
            if (vblnMandatoryNode == true && xmlNode == null)
            {
                if (vlngMessageNo == NO_MESSAGE_NUMBER)
                {
                    errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeMissingElement), "Match pattern:- " + vstrMatchPattern);
                }
                else
                {
                    errAssistEx.errThrowError(cstrFunctionName, (short)(vlngMessageNo), );
                }
            }
            xmlNode = null;
            return xmlNode;
        }

        public static MSXML2.IXMLDOMNode xmlGetNode(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            // header ----------------------------------------------------------------------------------
            // description:      Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:             vxmlParentNode      Node to be searched
            // return:           The node which matches the search pattern
            // Nothing if the node is not found
            // ------------------------------------------------------------------------------------------
            return GetNode(vxmlParentNode, vstrMatchPattern, false, NO_MESSAGE_NUMBER);
        }

        public static string xmlGetNodeText(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            // header ----------------------------------------------------------------------------------
            // description:      Gets the node text
            // pass:             vxmlParentNode      xml Node to search for child node
            // vstrMatchPattern    XSL pattern match
            // return:           Node text of child node
            // empty string if node not found
            // ------------------------------------------------------------------------------------------
            xmlNode = xmlGetNode(vxmlParentNode, vstrMatchPattern);
            if (!(xmlNode == null))
            {
                xmlNode = null;
            }
            return xmlNode.text;
        }

        public static bool xmlGetNodeAsBoolean(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       True if node text = "1" or "Y" or "YES" else False
            // False if node not found
            // ------------------------------------------------------------------------------------------
            strValue = xmlGetNodeText(vxmlParentNode, vstrMatchPattern);
            strValue = strValue.ToUpper();
            if (strValue == "1" || strValue == "Y" || strValue == "YES")
            {
            }
            return true;
        }

        public static DateTime xmlGetNodeAsDate(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       The text value of the found node as a Date
            // empty Date if node not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDate(xmlGetNodeText(vxmlParentNode, vstrMatchPattern));
        }

        public static double xmlGetNodeAsDouble(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       The text value of the found node as a Double
            // 0 if node not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDbl(xmlGetNodeText(vxmlParentNode, vstrMatchPattern));
        }

        public static short xmlGetNodeAsInteger(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       The text value of the found node as a Integer
            // 0 if node not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeInt(xmlGetNodeText(vxmlParentNode, vstrMatchPattern));
        }

        public static int xmlGetNodeAsLong(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       The text value of the found node as a Long
            // 0 if node not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeLng(xmlGetNodeText(vxmlParentNode, vstrMatchPattern));
        }

        public static void xmlCheckMandatoryNode(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls GetNode to check that specified node exists
            // GetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // ------------------------------------------------------------------------------------------
            GetNode(vxmlParentNode, vstrMatchPattern, true, vlngMessageNo);
        }

        public static MSXML2.IXMLDOMNode xmlGetMandatoryNode(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       IXMLDOMNode         The node which matches the search pattern or
            // ------------------------------------------------------------------------------------------
            return GetNode(vxmlParentNode, vstrMatchPattern, true, vlngMessageNo);
        }

        public static MSXML2.IXMLDOMNodeList xmlGetMandatoryNodeList(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            const string cstrFunctionName = "xmlGetMandatoryNodeList";
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       IXMLDOMNode         The node which matches the search pattern or
            // ------------------------------------------------------------------------------------------
            if (xmlGetMandatoryNodeList().length == 0)
            {
                if (vlngMessageNo == NO_MESSAGE_NUMBER)
                {
                    errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeMissingElement), "Match pattern:- " + vstrMatchPattern);
                }
                else
                {
                    errAssistEx.errThrowError(cstrFunctionName, (short)(vlngMessageNo), );
                }
            }
            return vxmlParentNode.selectNodes(vstrMatchPattern);
        }

        public static string xmlGetMandatoryNodeText(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       String              Node text of child node
            // ------------------------------------------------------------------------------------------
            xmlNode = xmlGetMandatoryNode(vxmlParentNode, vstrMatchPattern, vlngMessageNo);
            strValue = xmlNode.text;
            xmlNode = null;
            if (Strings.Len(strValue) == 0)
            {
                if (vlngMessageNo == NO_MESSAGE_NUMBER)
                {
                    errAssistEx.errThrowError(
                        "xmlGetMandatoryNodeText", (short)(
                        OMIGAERROR.oeMissingElementValue), 
                        "Match pattern:- " + vstrMatchPattern);
                }
                else
                {
                    errAssistEx.errThrowError("xmlGetMandatoryNodeText", (short)(vlngMessageNo), );
                }
            }
            return strValue;
        }

        public static bool xmlGetMandatoryNodeAsBoolean(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       True if node text = "1" or "Y" or "YES" else False
            // ------------------------------------------------------------------------------------------
            strValue = xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo);
            if (strValue == "1" || strValue == "Y" || strValue == "YES")
            {
            }
            return true;
        }

        public static DateTime xmlGetMandatoryNodeAsDate(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       The text value of the found node as a Date
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDate(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo));
        }

        public static double xmlGetMandatoryNodeAsDouble(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       The text value of the found node as a Double
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDbl(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo));
        }

        public static short xmlGetMandatoryNodeAsInteger(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       The text value of the found node as a Integer
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeInt(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo));
        }

        public static int xmlGetMandatoryNodeAsLong(MSXML2.IXMLDOMNode vxmlParentNode, string vstrMatchPattern, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetNode to check that specified node exists and return node
            // xmlGetNode will raise an error if this condition is not met
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       The text value of the found node as a Long
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeLng(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo));
        }

        public static MSXML2.IXMLDOMNode xmlGetAttributeNode(MSXML2.IXMLDOMNode vxmlNode, string vstrAttributeName) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Finds the attribute node specified by vstrAttributeName
            // in the XML node vxmlNode
            // pass:         vxmlNode            Node to be searched
            // vstrAttributeName   Attribute to be found in the node
            // return:       IXMLDOMNode         The attribute node
            // ------------------------------------------------------------------------------------------
            return vxmlNode.attributes.getNamedItem(vstrAttributeName);
        }

        public static bool xmlAttributeValueExists(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the attribute node specified by vstrAttributeName
            // in the XML node vxmlNode
            // pass:         vxmlNode            Node to be searched
            // vstrAttributeName   Attribute to be found in the node
            // return:       True if attribute (with value) exists on node
            // ------------------------------------------------------------------------------------------
            xmlNode = xmlGetAttributeNode(vobjNode, vstrAttribName);
            if (!(xmlNode == null))
            {
                if (Strings.Len(xmlNode.text) > 0)
                {
                }
            }
            xmlNode = null;
            return true;
        }

        public static string xmlGetAttributeText(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            string xmlGetAttributeText = String.Empty;
            MSXML2.IXMLDOMNode xmlNode = null;
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node text
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // return:           Node text of attribute
            // empty string if attribute does not exist on Node
            // ------------------------------------------------------------------------------------------
            xmlNode = xmlGetAttributeNode(vobjNode, vstrAttribName);
            if (!(xmlNode == null))
            {
                xmlGetAttributeText = xmlNode.text;
                xmlNode = null;
            }
            else
            {
                xmlGetAttributeText = vstrDefault;
            }
            return xmlGetAttributeText;
        }

        public static bool xmlGetAttributeAsBoolean(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // return:       True if attribute value is "1" or "Y" or "YES"
            // otherwise False
            // ------------------------------------------------------------------------------------------
            strValue = (xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault)).ToUpper();
            if (strValue == "1" || strValue == "Y" || strValue == "YES")
            {
            }
            return true;
        }

        public static DateTime xmlGetAttributeAsDate(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node as a date
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // return:           Date - Node text as a date
            // empty date if attribute value not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDate(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault));
        }

        public static double xmlGetAttributeAsDouble(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node as a double
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // return:           Double - Node text as a double
            // 0 if attribute value not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDbl(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault));
        }

        public static short xmlGetAttributeAsInteger(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node as a integer
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // return:           Integer - Node text as a integer
            // 0 if attribute value not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeInt(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault));
        }

        public static int xmlGetAttributeAsLong(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, string vstrDefault) 
        {
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node as a long
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // return:           Long - Node text as a long
            // 0 if attribute value not found
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeLng(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault));
        }

        public static MSXML2.IXMLDOMNode xmlGetMandatoryAttribute(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            MSXML2.IXMLDOMNode xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the attribute node specified by vstrAttributeName
            // in the XML node vxmlNode
            // pass:         vxmlNode            Node to be searched
            // vstrAttributeName   Attribute to be found in the node
            // vlngMessageNo       Optional message number to override default message
            // return:       IXMLDOMNode         The attribute node
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // ------------------------------------------------------------------------------------------
            xmlAttrib = vobjNode.attributes.getNamedItem(vstrAttribName);
            if (xmlAttrib == null)
            {
                if (vlngMessageNo == NO_MESSAGE_NUMBER)
                {
                    errAssistEx.errThrowError(
                        "xmlGetMandatoryAttribute", (short)(
                        OMIGAERROR.oeXMLMissingAttribute), 
                        "[@" + vstrAttribName + "]");
                }
                else
                {
                    errAssistEx.errThrowError("xmlGetMandatoryAttribute", (short)(vlngMessageNo), );
                }
            }
            else
            {

                xmlAttrib = null;
            }
            return xmlAttrib;
        }

        public static string xmlGetMandatoryAttributeText(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            MSXML2.IXMLDOMNode xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:      Gets the attribute node text
            // pass:             vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // vlngMessageNo - Optional message number to override default message
            // return:           String - Node text of attribute
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlAttrib = xmlGetMandatoryAttribute(vobjNode, vstrAttribName, vlngMessageNo);
            if (Strings.Len(xmlAttrib.text) == 0)
            {
                if (vlngMessageNo == NO_MESSAGE_NUMBER)
                {
                    errAssistEx.errThrowError(
                        "xmlGetMandatoryAttributeText", (short)(
                        OMIGAERROR.oeXMLInvalidAttributeValue), 
                        "[@" + vstrAttribName + "]");
                }
                else
                {
                    errAssistEx.errThrowError("xmlGetMandatoryAttributeText", (short)(vlngMessageNo), );
                }
            }
            else
            {

            }
            xmlAttrib = null;
            return xmlAttrib.text;
        }

        public static bool xmlGetMandatoryAttributeAsBoolean(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
            // pass:         vxmlParentNode      Node to be searched
            // vstrMatchPattern    XSL search pattern
            // vlngMessageNo       Optional message number to override default message
            // return:       True if attribute value is "1" or "Y" or "YES"
            // otherwise False
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            strValue = (xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo)).ToUpper();
            if (strValue == "1" || strValue == "Y" || strValue == "YES")
            {
            }
            return true;
        }

        public static DateTime xmlGetMandatoryAttributeAsDate(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Gets the attribute node as a date
            // pass:         vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // vlngMessageNo - Optional message number to override default message
            // return:       Date - Node text as a date
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDate(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo));
        }

        public static double xmlGetMandatoryAttributeAsDouble(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Gets the attribute node as a double
            // pass:         vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // vlngMessageNo - Optional message number to override default message
            // return:       Double - Node text as a double
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeDbl(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo));
        }

        public static short xmlMandatoryGetAttributeAsInteger(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Gets the attribute node as a integer
            // pass:         vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // vlngMessageNo - Optional message number to override default message
            // return:       Integer - Node text as a integer
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeInt(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo));
        }

        public static int xmlGetMandatoryAttributeAsLong(MSXML2.IXMLDOMNode vobjNode, string vstrAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Gets the attribute node as a long
            // pass:         vobjNode - xml Node to search for attribute
            // vstrAttribName - attribute name to search for
            // vlngMessageNo - Optional message number to override default message
            // return:       Long - Node text as a long
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            return ConvertAssistEx.CSafeLng(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo));
        }

        public static void xmlCheckMandatoryAttribute(MSXML2.IXMLDOMNode vxmlNode, string vstrAttributeName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  calls xmlGetMandatoryAttributeText to check that specified attribute exists
            // and has value,
            // xmlGetMandatoryAttributeText will raise an error if these conditions
            // are not met
            // pass:         vxmlNode          Node to be searched
            // vstrAttributeName Attribute to be found in the node
            // vlngMessageNo     Optional message number to override default message
            // exceptions:
            // raises error oeXMLMissingAttribute if attribute not found
            // raises error oeXMLInvalidAttributeValue if attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlGetMandatoryAttributeText(vxmlNode, vstrAttributeName, vlngMessageNo);
        }

        public static void xmlCopyAttribute(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
            // will overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // ------------------------------------------------------------------------------------------
            if (xmlAttributeValueExists(vxmlSrceNode, vstrSrceAttribName) == true)
            {
                vxmlDestNode.attributes.setNamedItem(
                    vxmlSrceNode.attributes.getNamedItem(vstrSrceAttribName).cloneNode(true));
            }
        }

        public static void xmlCopyMandatoryAttribute(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
            // will overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // vlngMessageNo: Optional message number to override default message
            // exceptions:
            // raises error oeXMLMissingAttribute if source attribute not found
            // raises error oeXMLInvalidAttributeValue if source attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlCheckMandatoryAttribute(vxmlSrceNode, vstrSrceAttribName, vlngMessageNo);
            xmlCopyAttribute(vxmlSrceNode, vxmlDestNode, vstrSrceAttribName);
        }

        public static void xmlCopyAttributeValue(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, string vstrDestAttribName) 
        {
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
            // to attribute named as vstrDestAttribName on vxmlDestNode
            // will overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // ------------------------------------------------------------------------------------------
            if (xmlAttributeValueExists(vxmlSrceNode, vstrSrceAttribName) == true)
            {

                xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName);
                xmlAttrib.text = vxmlSrceNode.attributes.getNamedItem(vstrSrceAttribName).text;
                vxmlDestNode.attributes.setNamedItem(xmlAttrib);
                xmlAttrib = null;
            }
        }

        public static void xmlCopyMandatoryAttributeValue(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, string vstrDestAttribName, int vlngMessageNo) 
        {
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
            // to attribute named as vstrDestAttribName on vxmlDestNode
            // will overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // vlngMessageNo: Optional message number to override default message
            // exceptions:
            // raises error oeXMLMissingAttribute if source attribute not found
            // raises error oeXMLInvalidAttributeValue if source attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlCheckMandatoryAttribute(vxmlSrceNode, vstrSrceAttribName, vlngMessageNo);
            xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName);

            xmlAttrib.text = vxmlSrceNode.attributes.getNamedItem(vstrSrceAttribName).text;
            vxmlDestNode.attributes.setNamedItem(xmlAttrib);
            xmlAttrib = null;
        }

        public static void xmlCopyAttribIfMissingFromDest(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
            // will not overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // ------------------------------------------------------------------------------------------
            if (xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) == false)
            {
                xmlCopyAttribute(vxmlSrceNode, vxmlDestNode, vstrSrceAttribName);
            }
        }

        // 
        public static MSXML2.FreeThreadedDOMDocument40 xmlCreateDOMObject() 
        {
            MSXML2.FreeThreadedDOMDocument40 vxmlDOMDocument = null;

            vxmlDOMDocument = new MSXML2.FreeThreadedDOMDocument40();
            vxmlDOMDocument.setProperty("NewParser", true);
            vxmlDOMDocument.validateOnParse = false;
            vxmlDOMDocument.async = false;
            vxmlDOMDocument = null;
            return vxmlDOMDocument;
        }

        public static void xmlCopyMandatoryAttribIfMissingFromDest(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
            // will not overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // vlngMessageNo: Optional message number to override default message
            // exceptions:
            // raises error oeXMLMissingAttribute if source attribute not found
            // raises error oeXMLInvalidAttributeValue if source attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlCheckMandatoryAttribute(vxmlSrceNode, vstrSrceAttribName, vlngMessageNo);
            xmlCopyAttribIfMissingFromDest(vxmlSrceNode, vxmlDestNode, vstrSrceAttribName);
        }

        public static void xmlChangeNodeName(ref MSXML2.IXMLDOMNode rxmlNode, string vstrOldName, string vstrNewName) 
        {
            string strFunctionName = String.Empty;
            MSXML2.IXMLDOMNode objNewNode = null;
            short intListLength = 0;
            short intChildNodesLength = 0;
            short intChildIndex = 0;
            short intAttribIndex = 0;
            MSXML2.IXMLDOMElement objElement = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Changes all nodes present in rxmlNode that have a node name
            // of vstrOldName to vstrNewName
            // 
            // pass:
            // rxmlNode    Xml Node
            // vstrOldName Name of nodes to be changed
            // vstrNewName Name nodes are to be changed to
            // return:   n/a
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO xmlChangeNodeNameVbErr is not supported
            try 
            {

                strFunctionName = "xmlChangeNodeName";


                // If this node is to be renamed create a new node with the new name
                // and copy all the attributes across
                if (rxmlNode.nodeName == vstrOldName)
                {
                    objNewNode = rxmlNode.ownerDocument.createNode(rxmlNode.nodeType, vstrNewName, rxmlNode.namespaceURI);
                    // Copy attributes if this is an element node
                    if (objNewNode.nodeType == MSXML2.DOMNodeType.NODE_ELEMENT)
                    {

                        objElement = objNewNode;
                        intListLength = (short)(rxmlNode.attributes.length - 1);
                        for(intAttribIndex = (short)0;
                            intAttribIndex <= intListLength;
                            intAttribIndex = Convert.ToInt16(intAttribIndex + 1))
                        {
                            // ISSUE: Method or data member not found: 'cloneNode'
                            objElement.setAttributeNode(rxmlNode.attributes[intAttribIndex].cloneNode(true));
                        }
                    }
                }

                // For all children of this node change their name if it is vstrOldName and
                // append them to the new node
                intChildNodesLength = (short)(rxmlNode.childNodes.length - 1);
                for(intChildIndex = (short)0;
                    intChildIndex <= intChildNodesLength;
                    intChildIndex = Convert.ToInt16(intChildIndex + 1))
                {
                    xmlChangeNodeName(rxmlNode.childNodes[intChildIndex], vstrOldName, vstrNewName);
                    if (!(objNewNode == null))
                    {
                        // ISSUE: Method or data member not found: 'cloneNode'
                        objNewNode.appendChild(rxmlNode.childNodes[intChildIndex].cloneNode(true));
                    }
                }
                // Replace the original node with the new node
                if (!(objNewNode == null))
                {
                    if (!((rxmlNode.parentNode == null)))
                    {
                        rxmlNode.parentNode.replaceChild(objNewNode, rxmlNode);
                    }
                    rxmlNode = objNewNode;
                }
                objNewNode = null;
                objElement = null;
                return ;
                // WARNING: xmlChangeNodeNameVbErr: is not supported 
            }
            catch(Exception exc)
            {
                objNewNode = null;
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static void xmlCopyAttribValueIfMissingFromDest(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, string vstrDestAttribName) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
            // to attribute named as vstrDestAttribName on vxmlDestNode
            // will not overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // ------------------------------------------------------------------------------------------
            if (xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) == false)
            {
                xmlCopyAttributeValue(
                    vxmlSrceNode, vxmlDestNode, vstrSrceAttribName, vstrDestAttribName);
            }
        }

        public static void xmlCopyMandatoryAttribValueIfMissingFromDest(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode, string vstrSrceAttribName, string vstrDestAttribName, int vlngMessageNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
            // to attribute named as vstrDestAttribName on vxmlDestNode
            // will not overwrite existing attribute value on vxmlDestNode
            // will do nothing if source attribute not present or has no value
            // pass:
            // vxmlSrceNode: source node for attribute
            // vxmlDestNode: destination node for attribute
            // vstrSrceAttribName: name of attribute to be copied
            // vlngMessageNo: Optional message number to override default message
            // exceptions:
            // raises error oeXMLMissingAttribute if source attribute not found
            // raises error oeXMLInvalidAttributeValue if source attribute has no value
            // ------------------------------------------------------------------------------------------
            xmlCheckMandatoryAttribute(vxmlSrceNode, vstrSrceAttribName, vlngMessageNo);
            if (xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) == false)
            {
                xmlCopyAttribValueIfMissingFromDest(
                    vxmlSrceNode, vxmlDestNode, vstrSrceAttribName, vstrDestAttribName);
            }
        }

        public static void xmlSetAttributeValue(MSXML2.IXMLDOMNode vxmlDestNode, string vstrAttribName, string vstrAttribValue) 
        {
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Creates a xml attribute on vxmlDestNode,
            // will overwrite existing attribute value
            // pass:
            // vxmlDestNode - Destination Node for new attribute to be placed on
            // vstrAttribName - Attrbiute name of the new attribute
            // vstrAttribValue - Attribute value for new attribute on vxmlDestNode
            // ------------------------------------------------------------------------------------------
            xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName);
            xmlAttrib.value = vstrAttribValue;

            vxmlDestNode.attributes.setNamedItem(xmlAttrib.cloneNode(true));
            xmlAttrib = null;
        }

        public static void xmlSetAttributeValueIfMissingFromDest(MSXML2.IXMLDOMNode vxmlDestNode, string vstrAttribName, string vstrAttribValue) 
        {
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Creates a xml attribute on vxmlDestNode,
            // will not overwrite existing attribute value
            // pass:
            // vxmlDestNode - Destination Node for new attribute to be placed on
            // vstrAttribName - Attrbiute name of the new attribute
            // vstrAttribValue - Attribute value for new attribute on vxmlDestNode
            // ------------------------------------------------------------------------------------------
            if (xmlAttributeValueExists(vxmlDestNode, vstrAttribName) == false)
            {
                xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName);
                xmlAttrib.value = vstrAttribValue;

                vxmlDestNode.attributes.setNamedItem(xmlAttrib.cloneNode(true));
                xmlAttrib = null;
            }
        }

        public static MSXML2.FreeThreadedDOMDocument40 xmlCreateElementRequestFromNode(MSXML2.IXMLDOMNode vxmlNode, string vstrMasterTagName, bool blnConvertChildren, string vstrNewTagName) 
        {
            MSXML2.FreeThreadedDOMDocument40 xmlDOMDocument = null;
            MSXML2.IXMLDOMNode xmlRequestNode = null;
            MSXML2.IXMLDOMNode xmlDestParentNode = null;
            MSXML2.IXMLDOMNode xmlNode = null;
            MSXML2.IXMLDOMNodeList xmlNodeList = null;
            short intNodeCount = 0;
            short intListLength = 0;
            string strFunctionName = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Convert a Request node from a attribute based to element based.
            // pass:
            // vxmlNode            xml Request node to be converted
            // vstrMasterTagName   name of parent tag (can be pattern match)
            // bInConvertChildren  True/False - Convert all children of "MasterTagName" node
            // vstrNewTagName      New name for "Master Tag" (optional)
            // return:
            // FreeThreadedDOMDocument40         Converted Request
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            strFunctionName = "xmlCreateElementRequestFromNode";
            if (!(vxmlNode == null))
            {
                xmlDOMDocument = new MSXML2.FreeThreadedDOMDocument40();
                // create a new request node
                xmlNode = xmlDOMDocument.createElement("REQUEST");
                if (!(xmlNode == null))
                {
                    xmlRequestNode = xmlDOMDocument.appendChild(xmlNode);
                }
                // clone the request attributes of the passed in request node
                xmlNode = vxmlNode.ownerDocument.selectSingleNode("REQUEST");
                intListLength = (short)(vxmlNode.attributes.length - 1);
                for(intNodeCount = (short)0;
                    intNodeCount <= intListLength;
                    intNodeCount = Convert.ToInt16(intNodeCount + 1))
                {
                    // ISSUE: Method or data member not found: 'nodeName'
                    xmlCopyAttribute(vxmlNode, xmlRequestNode, Convert.ToString(vxmlNode.attributes[intNodeCount].nodeName));
                }
                // Extract the node to convert
                xmlNodeList = vxmlNode.selectNodes(vstrMasterTagName);
                if (xmlNodeList.length == 0)
                {
                    errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLMissingElementText), "No matching nodes found for: " + vstrMasterTagName);
                }
                foreach(MSXML2.IXMLDOMNode __each1 in xmlNodeList)
                {
                    xmlNode = __each1;
                    xmlNode = xmlMakeNodeElementBased(xmlNode, blnConvertChildren, vstrNewTagName, );
                    xmlDestParentNode = xmlRequestNode.appendChild(xmlNode);
                }
            }
        xmlCreateElementRequestFromNodeExit:
            xmlDOMDocument = null;
            xmlRequestNode = null;
            xmlDestParentNode = null;
            xmlNode = null;
            xmlNodeList = null;
            return xmlDOMDocument;
        }

        public static MSXML2.IXMLDOMNode xmlMakeNodeElementBased(MSXML2.IXMLDOMNode vxmlNode, bool blnConvertChildren, string vstrNewTagName, params object[] vstrAttributes) 
        {
            MSXML2.FreeThreadedDOMDocument40 xmlDOMDocument = null;
            MSXML2.IXMLDOMNode xmlDestParentNode = null;
            MSXML2.IXMLDOMNode xmlNode = null;
            MSXML2.IXMLDOMNode xmlNode2 = null;
            short intAttribsLength = 0;
            short intAttribsCount = 0;
            string strNodeName = String.Empty;
            string strComboValueNodeName = String.Empty;
            string strAttribName = String.Empty;
            short intIndex = 0;
            string strNodeText = String.Empty;
            short intNoOfSpecifiedTags = 0;
            string strFunctionName = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Convert a single node from a attribute based to element based.
            // pass:
            // vxmlNode            xml Request node to be converted
            // bInConvertChildren  True/False - Convert all children of "MasterTagName" node
            // vstrNewTagName      New name for "Master Tag" (optional)
            // vstrAttributes      ParamArray can be used to name individual attributes to be
            // converted. If not used all attributes are converted.
            // return:
            // Node                Converted Node
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            strFunctionName = "xmlMakeNodeElementBased";
            if (!(vxmlNode == null))
            {
                xmlDOMDocument = new MSXML2.FreeThreadedDOMDocument40();
                // Determine the name of the new node
                if (Strings.Len(vstrNewTagName.Trim()) > 0)
                {
                    strNodeName = vstrNewTagName;
                }
                else
                {
                    strNodeName = vxmlNode.nodeName;
                }
                // create the new node
                xmlNode = xmlDOMDocument.createElement(strNodeName);
                if (!(xmlNode == null))
                {
                    xmlDestParentNode = xmlDOMDocument.appendChild(xmlNode);
                }
                // Process each attribute
                // have specified a subset of the attributes to clone?
                // if we have not specified and specific tags the UBound = -1
                intNoOfSpecifiedTags = (short)(Information.UBound(vstrAttributes, 1));
                if (intNoOfSpecifiedTags >= 0)
                {
                    for(intIndex = (short)0;
                        intIndex <= Convert.ToDouble(intIndex > intNoOfSpecifiedTags);
                        intIndex = Convert.ToInt16(intIndex + 1))
                    {
                        strNodeText = xmlGetAttributeText(vxmlNode, Convert.ToString(vstrAttributes[intIndex]), "");
                        if (Strings.Len(strNodeText) > 0)
                        {
                            // SR 12/06/01 : if node name is description of any combo value, create a new attribute
                            // to the respective node
                            strAttribName = Convert.ToString(vstrAttributes[intIndex]);
                            if (strAttribName.Substring(strAttribName.Length - 5) == "_TEXT")
                            {
                                strComboValueNodeName = strAttribName.Substring(0, Strings.Len(strAttribName) - 5);
                                xmlNode2 = xmlDestParentNode.selectSingleNode("./" + strComboValueNodeName);
                                if (!(xmlNode2 == null))
                                {
                                    xmlSetAttributeValue(xmlNode2, "TEXT", strNodeText);
                                }
                                else
                                {
                                    xmlNode = xmlDOMDocument.createElement(Convert.ToString(vstrAttributes[intIndex]));
                                    xmlNode.text = strNodeText;
                                    xmlDestParentNode.appendChild(xmlNode);
                                }
                            }
                            else
                            {
                                xmlNode = xmlDOMDocument.createElement(Convert.ToString(vstrAttributes[intIndex]));
                                xmlNode.text = strNodeText;
                                xmlDestParentNode.appendChild(xmlNode);
                            }
                        }
                        else
                        {
                            errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLMissingAttribute), ": " + vstrAttributes[intIndex]);
                        }
                    }
                }
                else
                {
                    intAttribsLength = (short)(vxmlNode.attributes.length - 1);
                    // ++ SA SYS4754
                    for(intAttribsCount = (short)0;
                        intAttribsCount <= intAttribsLength;
                        intAttribsCount = Convert.ToInt16(intAttribsCount + 1))
                    {
                        // ++ SA SYS4754
                        // SR 12/06/01 : if node name is description of any combo value, create a new attribute
                        // to the respective node
                        // ISSUE: Method or data member not found: 'nodeName'
                        strAttribName = Convert.ToString(vxmlNode.attributes[intAttribsCount].nodeName);
                        if (strAttribName.Substring(strAttribName.Length - 5) == "_TEXT")
                        {
                            strComboValueNodeName = strAttribName.Substring(0, Strings.Len(strAttribName) - 5);
                            xmlNode2 = xmlDestParentNode.selectSingleNode("./" + strComboValueNodeName);
                            if (!(xmlNode2 == null))
                            {
                                // ISSUE: Method or data member not found: 'Text'
                                xmlSetAttributeValue(xmlNode2, "TEXT", Convert.ToString(vxmlNode.attributes[intAttribsCount].Text));
                            }
                            else
                            {
                                // ISSUE: Method or data member not found: 'nodeName'
                                xmlNode = xmlDOMDocument.createElement(Convert.ToString(vxmlNode.attributes[intAttribsCount].nodeName));
                                // ++ SA SYS4754
                                if (!(xmlNode == null))
                                {
                                    // ISSUE: Method or data member not found: 'Text'
                                    xmlNode.text = Convert.ToString(vxmlNode.attributes[intAttribsCount].Text);
                                    xmlDestParentNode.appendChild(xmlNode);
                                }
                            }
                        }
                        else
                        {
                            // ISSUE: Method or data member not found: 'nodeName'
                            xmlNode = xmlDOMDocument.createElement(Convert.ToString(vxmlNode.attributes[intAttribsCount].nodeName));
                            if (!(xmlNode == null))
                            {
                                // ISSUE: Method or data member not found: 'Text'
                                xmlNode.text = Convert.ToString(vxmlNode.attributes[intAttribsCount].Text);
                                xmlDestParentNode.appendChild(xmlNode);
                            }
                        }
                    }
                }
                // If required, process each child element recursively (converting all attributes)
                if (blnConvertChildren && vxmlNode.hasChildNodes())
                {
                    foreach(MSXML2.IXMLDOMNode __each2 in vxmlNode.childNodes)
                    {
                        xmlNode = __each2;
                        xmlNode = xmlMakeNodeElementBased(xmlNode, blnConvertChildren, "", );
                        xmlDOMDocument.createElement(xmlNode.nodeName);
                        xmlDestParentNode.appendChild(xmlNode);
                    }
                }
            }
        xmlMakeNodeElementBasedExit:
            xmlDOMDocument = null;
            xmlDestParentNode = null;
            xmlNode = null;
            xmlNode2 = null;
            return xmlDestParentNode;
        }

        public static MSXML2.IXMLDOMNode xmlCreateAttributeBasedResponse(MSXML2.IXMLDOMNode vxmlNode, bool blnConvertChild) 
        {
            MSXML2.IXMLDOMElement xmlParent = null;
            MSXML2.IXMLDOMNode xmlNode = null;
            bool blnNoChildren = false;
            short intChildCount = 0;
            short intChildLength = 0;
            short intAttribsCount = 0;
            short intAttribsLength = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // Convert a Response node from element based to attribute based.
            // pass:
            // vxmlNode            xml Response node to be converted
            // bInConvertChildren  True/False - Convert all children of parent node also
            // return:
            // Node                Converted Request
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            if (!(vxmlNode == null))
            {
                // Clone the parent node (without child nodes) as a basis for the returned node
                xmlParent = vxmlNode.cloneNode(false);
                // Iterate through each child node
                intChildLength = (short)(vxmlNode.childNodes.length - 1);
                for(intChildCount = (short)0;
                    intChildCount <= intChildLength;
                    intChildCount = Convert.ToInt16(intChildCount + 1))
                {

                    // Check if it is a 'text' node or genuinely has child elements
                    // ISSUE: Method or data member not found: 'hasChildNodes'
                    if (Convert.ToBoolean(vxmlNode.childNodes[intChildCount].hasChildNodes))
                    {
                        // ISSUE: Method or data member not found: 'childNodes'
                        if (Convert.ToString(vxmlNode.childNodes[intChildCount].childNodes(0).nodeName) == "#text")
                        {
                            blnNoChildren = true;
                        }
                    }
                    else
                    {
                        blnNoChildren = true;
                    }
                    // ISSUE: Method or data member not found: 'Text'
                    if (blnNoChildren && Strings.Len((Convert.ToString(vxmlNode.childNodes[intChildCount].Text)).Trim()) > 0)
                    {

                        // Create as an attribute of parent
                        // ISSUE: Method or data member not found: 'nodeName'
                        // ISSUE: Method or data member not found: 'Text'
                        xmlParent.setAttribute(Convert.ToString(vxmlNode.childNodes[intChildCount].nodeName), vxmlNode.childNodes[intChildCount].Text);
                        // Check if combo item with TEXT attribute
                        // ISSUE: Method or data member not found: 'Attributes'
                        intAttribsLength = (short)(Convert.ToInt32(vxmlNode.childNodes[intChildCount].Attributes.length) - 1);
                        for(intAttribsCount = (short)0;
                            intAttribsCount <= intAttribsLength;
                            intAttribsCount = Convert.ToInt16(intAttribsCount + 1))
                        {
                            // ISSUE: Method or data member not found: 'Attributes'
                            if ((Convert.ToString(vxmlNode.childNodes[intChildCount].Attributes(intAttribsCount).nodeName)).ToUpper() == "TEXT")
                            {

                                // Create a text attribute also for the combo text
                                // ISSUE: Method or data member not found: 'nodeName'
                                // ISSUE: Method or data member not found: 'Attributes'
                                xmlParent.setAttribute(vxmlNode.childNodes[intChildCount].nodeName + "_TEXT", vxmlNode.childNodes[intChildCount].Attributes(intAttribsCount).Text);
                            }
                        }
                    }
                    else if( blnConvertChild && !blnNoChildren)
                    {

                        // Process child node recursively
                        xmlNode = xmlCreateAttributeBasedResponse(vxmlNode.childNodes[intChildCount], blnConvertChild);
                        xmlParent.appendChild(xmlNode);
                    }
                    blnNoChildren = false;
                }
            }
        xmlCreateAttributeBasedResponseExit:
            xmlParent = null;
            xmlNode = null;
            return xmlParent;
        }

        public static MSXML2.IXMLDOMNode xmlCreateChildRequest(MSXML2.IXMLDOMElement vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, string vstrNewName, bool blnComboLookup) 
        {
            string strFunctionName = String.Empty;
            MSXML2.FreeThreadedDOMDocument40 xmlDoc = null;
            MSXML2.IXMLDOMElement xmlNewNode = null;
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            short intSchemaLength = 0;
            short intSchemaCount = 0;
            string strValue = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Create a Request node that can be used to find all matching records in a child table
            // pass:
            // vxmlDataNode        xml element to be converted
            // vxmlSchemaNode      Schema entry for the parent table
            // vstrNewName         Name of the child node to be created
            // return:
            // Node                Converted Request
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            strFunctionName = "xmlCreateChildRequest";
            if ((!(vxmlDataNode == null)) && (!(vxmlSchemaNode == null)) && Strings.Len(vstrNewName.Trim()) > 0)
            {
                // Create a new root node
                xmlDoc = new MSXML2.FreeThreadedDOMDocument40();
                xmlNewNode = xmlDoc.createElement(vstrNewName);
                if (blnComboLookup)
                {
                    xmlNewNode.setAttribute("_COMBOLOOKUP_", "1");
                }
                // Iterate through the schema node to get the primary key items
                intSchemaLength = (short)(vxmlSchemaNode.childNodes.length - 1);
                for(intSchemaCount = (short)0;
                    intSchemaCount <= intSchemaLength;
                    intSchemaCount = Convert.ToInt16(intSchemaCount + 1))
                {
                    // ISSUE: Method or data member not found: 'Attributes'
                    if (!(vxmlSchemaNode.childNodes[intSchemaCount].Attributes.getNamedItem("KEYTYPE") == null))
                    {
                        // ISSUE: Method or data member not found: 'Attributes'
                        // ISSUE: Don't know the object type: 'vxmlSchemaNode.childNodes[intSchemaCount].Attributes'
                        strValue = Convert.ToString(vxmlSchemaNode.childNodes[intSchemaCount].Attributes.getNamedItem("KEYTYPE").Text);
                        if (strValue == "PRIMARY")
                        {
                            // Get this attribute from the data element and add it to the new node
                            // ISSUE: Method or data member not found: 'nodeName'
                            xmlAttrib = vxmlDataNode.attributes.getNamedItem(Convert.ToString(vxmlSchemaNode.childNodes[intSchemaCount].nodeName));
                            if (!(xmlAttrib == null))
                            {
                                // ISSUE: Method or data member not found: 'nodeName'
                                xmlNewNode.setAttribute(Convert.ToString(vxmlSchemaNode.childNodes[intSchemaCount].nodeName), xmlAttrib.text);
                            }
                        }
                    }
                }
            }
            else
            {
                errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeInvalidParameter), );
            }
        xmlCreateChildRequestExit:
            xmlNewNode = null;
            xmlAttrib = null;
            xmlDoc = null;
            return xmlNewNode;
        }

        public static void xmlElemFromAttrib(MSXML2.IXMLDOMNode vxmlDestParentNode, MSXML2.IXMLDOMNode vxmlSrceNode, string vstrAttribName) 
        {
            MSXML2.IXMLDOMElement xmlElem = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Create a Element node from an attribute
            // with node name = attribute name and node value = attribute value
            // pass:
            // vxmlDestParentNode  parent node for element to be created
            // vxmlSrceNode        node that contains attribute that is source of new element
            // vstrAttribName      Name of attribute to be located on vxmlSrceNode
            // ------------------------------------------------------------------------------------------
            xmlElem = vxmlDestParentNode.ownerDocument.createElement(vstrAttribName);
            if (!(vxmlSrceNode.attributes.getNamedItem(vstrAttribName) == null))
            {
                xmlElem.text = vxmlSrceNode.attributes.getNamedItem(vstrAttribName).text;
            }
            vxmlDestParentNode.appendChild(xmlElem);
            xmlElem = null;
        }

        public static MSXML2.IXMLDOMNode xmlGetRequestNode(MSXML2.IXMLDOMElement vxmlElement) 
        {
            const string cstrFunctionName = "xmlGetRequestNode";
            MSXML2.FreeThreadedDOMDocument40 xmlDocument = null;
            MSXML2.IXMLDOMNode xmlRequestNode = null;
            MSXML2.IXMLDOMNode xmlNode = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Get the <REQUEST> node from an xml element, without any children. T
            // pass:
            // vxmlElement The element from which to get the request node
            // return:
            // Request node
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            if (vxmlElement == null)
            {
                errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeMissingElement), );
            }
            // If the top most tag of the node is 'REQUEST', just return the node, else
            // search for 'REQUEST'
            if (vxmlElement.nodeName == "REQUEST")
            {
                xmlRequestNode = vxmlElement.cloneNode(false);
            }
            else
            {
                xmlNode = vxmlElement.selectSingleNode("//REQUEST");
                if (xmlNode == null)
                {
                    errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeMissingPrimaryTag), "Expected REQUEST tag");
                }
                xmlRequestNode = xmlNode.cloneNode(false);
            }
            // Attach the cloned node to a new dom document to ensure safety of using selectsinglenode
            // with something like this: "/REQUEST/SOMETHING"
            xmlDocument = new MSXML2.FreeThreadedDOMDocument40();
            xmlDocument.documentElement = xmlRequestNode;
            return xmlRequestNode;
        xmlGetRequestNodeExit:
            xmlDocument = null;
            xmlRequestNode = null;
            xmlNode = null;
        }

        public static void AttachResponseData(MSXML2.IXMLDOMNode vxmlNodeToAttachTo, MSXML2.IXMLDOMElement vxmlResponse) 
        {
            string strFunctionName = String.Empty;
            MSXML2.IXMLDOMNodeList xmlChildList = null;
            short intChildListLength = 0;
            short intChildCount = 0;
            // Header ----------------------------------------------------------------------------------
            // Description:
            // Gets the data nodes out of vxmlResponse and appends them to vxmlNodeToAttachTo
            // Pass:
            // vxmlNodeToAttachTo  xml node to attach the data to
            // vxmlResponse        xml element containing the response whos data is to be
            // extracted
            // ------------------------------------------------------------------------------------------
            strFunctionName = "AttachResponseData";

            if (vxmlNodeToAttachTo == null || vxmlResponse == null)
            {
                errAssistEx.errThrowError(strFunctionName, (short)(
                    OMIGAERROR.oeInvalidParameter), 
                    "Node to attach to or Response missing");
            }
            if (vxmlResponse.nodeName != "RESPONSE")
            {
                errAssistEx.errThrowError(strFunctionName, (short)(
                    OMIGAERROR.oeInvalidParameter), 
                    "RESPONSE must be top level tag");
            }

            xmlChildList = vxmlResponse.childNodes;
            intChildListLength = (short)(xmlChildList.length - 1);
            for(intChildCount = (short)0;
                intChildCount <= intChildListLength;
                intChildCount = Convert.ToInt16(intChildCount + 1))
            {
                if (Convert.ToString(xmlChildList[intChildCount].nodeName) != "MESSAGE")
                {
                    vxmlNodeToAttachTo.appendChild(xmlChildList[intChildCount].cloneNode(true));
                }
            }
            xmlChildList = null;
        }


    }

}

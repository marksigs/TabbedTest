using System;
using MSXML2;
using Microsoft.VisualBasic;
using Vertex.Fsd.Omiga.VisualBasicPort;
namespace omLockManager
{
    public class globalAssist
    {

        private const string gstrMODULEPREFIX = "globalAssist.";
        // DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
        public static bool GetGlobalParamBoolean(string vstrParamName) 
        {
            bool varParamValue = false;
            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptBoolean, ref varParamValue))
            {
            }
            return varParamValue;
        }

        // DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
        public static bool GetMandatoryGlobalParamBoolean(string vstrParamName) 
        {
            const string cstrFunctionName = "GetMandatoryGlobalParamBoolean";
            string varParamValue = String.Empty;

            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptBoolean, ref varParamValue))
            {

            }
            else
            {

                errAssistEx.errThrowError(
                    gstrMODULEPREFIX + cstrFunctionName, (short)(
                    OMIGAERROR.oeMissingParameter), 
                    vstrParamName);
            }
            return varParamValue;
        }

        public static string GetMandatoryGlobalParamString(string vstrParamName) 
        {
            const string cstrFunctionName = "GetMandatoryGlobalParamString";
            string varParamValue = String.Empty;

            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptString, ref varParamValue))
            {

            }
            else
            {

                errAssistEx.errThrowError(
                    gstrMODULEPREFIX + cstrFunctionName, (short)(
                    OMIGAERROR.oeMissingParameter), 
                    vstrParamName);
            }
            return varParamValue;
        }

        // APS SYS1920 Added new method
        public static int GetMandatoryGlobalParamAmount(string vstrParamName) 
        {
            const string cstrFunctionName = "GetMandatoryGlobalParamAmount";
            string varParamValue = String.Empty;

            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptAmount, ref varParamValue))
            {

            }
            else
            {

                errAssistEx.errThrowError(
                    gstrMODULEPREFIX + cstrFunctionName, (short)(
                    OMIGAERROR.oeMissingParameter), 
                    vstrParamName);
            }
            return varParamValue;
        }

        public static string GetGlobalParamString(string vstrParamName) 
        {
            string varParamValue = String.Empty;
            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptString, ref varParamValue))
            {
            }
            return varParamValue;
        }

        // APS SYS1920 Added new method
        public static int GetGlobalParamAmount(string vstrParamName) 
        {
            string varParamValue = String.Empty;
            if (GetGlobalParam(vstrParamName, GLOBALPARAMTYPE.oegptAmount, ref varParamValue))
            {
            }
            return varParamValue;
        }

        private static bool GetGlobalParam(string vstrParamName, globalAssist.GLOBALPARAMTYPE enumParamType, ref object vvarParamValue) 
        {
            bool GetGlobalParam = false;
            const string cstrFunctionName = "GetGlobalParam";
            MSXML2.FreeThreadedDOMDocument40 xmlDoc = null;
            MSXML2.IXMLDOMElement xmlElem = null;
            MSXML2.IXMLDOMNode xmlRootNode = null;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNode xmlRequestNode = null;
            MSXML2.IXMLDOMNode xmlResponseNode = null;
            // create one FreeThreadedDOMDocument40 to contain schema, request & response
            xmlDoc = new MSXML2.FreeThreadedDOMDocument40();
            // create BOGUS root node
            xmlElem = xmlDoc.createElement("BOGUS");
            xmlRootNode = xmlDoc.appendChild(xmlElem);
            // create schema for OM_GLOBALPARAM view
            xmlElem = xmlDoc.createElement("GLOBALPARAM");
            xmlElem.setAttribute("ENTITYTYPE", "LOGICAL");
            xmlElem.setAttribute("DATASRCE", "OM_GLOBALPARAM");
            xmlSchemaNode = xmlRootNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("NAME");
            xmlElem.setAttribute("DATATYPE", "STRING");
            xmlElem.setAttribute("LENGTH", "30");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("GLOBALPARAMETERSTARTDATE");
            xmlElem.setAttribute("DATATYPE", "DATE");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("STRING");
            xmlElem.setAttribute("DATATYPE", "STRING");
            xmlElem.setAttribute("LENGTH", "30");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("DESCRIPTION");
            xmlElem.setAttribute("DATATYPE", "STRING");
            xmlElem.setAttribute("LENGTH", "255");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("AMOUNT");
            xmlElem.setAttribute("DATATYPE", "DOUBLE");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("AMOUNT");
            xmlElem.setAttribute("DATATYPE", "DOUBLE");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("MAXIMUMAMOUNT");
            xmlElem.setAttribute("DATATYPE", "DOUBLE");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("PERCENTAGE");
            xmlElem.setAttribute("DATATYPE", "DOUBLE");
            xmlSchemaNode.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("BOOLEAN");
            xmlElem.setAttribute("DATATYPE", "BOOLEAN");
            xmlSchemaNode.appendChild(xmlElem);

            // create REQUEST details
            xmlElem = xmlDoc.createElement("GLOBALPARAM");
            xmlElem.setAttribute("NAME", vstrParamName);
            xmlRequestNode = xmlRootNode.appendChild(xmlElem);

            // create RESPONSE node
            xmlElem = xmlDoc.createElement("RESPONSE");
            xmlResponseNode = xmlRootNode.appendChild(xmlElem);
            adoAssistEx.adoGetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlResponseNode, String.Empty, String.Empty);
            xmlResponseNode = xmlAssistEx.xmlGetNode(xmlResponseNode, "GLOBALPARAM");
            if (!(xmlResponseNode == null))
            {
                switch (enumParamType)
                {
                    case GLOBALPARAMTYPE.oegptString:
                        if (xmlAssistEx.xmlAttributeValueExists(xmlResponseNode, "STRING"))
                        {
                            vvarParamValue = xmlAssistEx.xmlGetAttributeText(xmlResponseNode, "STRING", "");
                            GetGlobalParam = true;
                        }
                        break;
                    case GLOBALPARAMTYPE.oegptAmount:
                        if (xmlAssistEx.xmlAttributeValueExists(xmlResponseNode, "AMOUNT"))
                        {
                            vvarParamValue = xmlAssistEx.xmlGetAttributeAsDouble(xmlResponseNode, "AMOUNT", "");
                            GetGlobalParam = true;
                        }
                        break;
                    case GLOBALPARAMTYPE.oegptMaximumAmount:
                        if (xmlAssistEx.xmlAttributeValueExists(xmlResponseNode, "MAXIMUMAMOUNT"))
                        {
                            vvarParamValue = xmlAssistEx.xmlGetAttributeAsDouble(xmlResponseNode, "MAXIMUMAMOUNT", "");
                            GetGlobalParam = true;
                        }
                        break;
                    case GLOBALPARAMTYPE.oegptPercentage:
                        if (xmlAssistEx.xmlAttributeValueExists(xmlResponseNode, "PERCENTAGE"))
                        {
                            vvarParamValue = xmlAssistEx.xmlGetAttributeAsDouble(xmlResponseNode, "PERCENTAGE", "");
                            GetGlobalParam = true;
                        }
                        // DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
                        break;
                    case GLOBALPARAMTYPE.oegptBoolean:
                        if (xmlAssistEx.xmlAttributeValueExists(xmlResponseNode, "BOOLEAN"))
                        {
                            // SYS3185 Changed this to call the xmlGetAttributeAsBoolean method
                            vvarParamValue = xmlAssistEx.xmlGetAttributeAsBoolean(xmlResponseNode, "BOOLEAN", "");
                            GetGlobalParam = true;
                        }
                        break; 
                }
            }
            xmlElem = null;
            xmlRootNode = null;
            xmlSchemaNode = null;
            xmlRequestNode = null;
            xmlResponseNode = null;
            xmlDoc = null;
            return GetGlobalParam;
        }

        // SR 25/01/2006 : BBGRb74 - New method 'GetAllGlobalBandedParamValuesAsXml'
        public static void GetAllGlobalBandedParamValuesAsXml(string vstrParamName, string strValueType, ref MSXML2.IXMLDOMNode vxmlResponse) 
        {
            const string cstrFunctionName = "GetAllGlobalBandedParamValuesAsXml";
            MSXML2.FreeThreadedDOMDocument40 xmlDoc = null;
            MSXML2.IXMLDOMElement xmlElem = null;
            MSXML2.IXMLDOMNode xmlRootNode = null;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNode xmlRequestNode = null;



            xmlDoc = new MSXML2.FreeThreadedDOMDocument40();
            // create BOGUS root node
            xmlElem = xmlDoc.createElement("BOGUS");
            xmlRootNode = xmlDoc.appendChild(xmlElem);
            xmlElem = xmlDoc.createElement("GLOBALBANDEDPARAM");
            xmlElem.setAttribute("ENTITYTYPE", "LOGICAL");
            xmlElem.setAttribute("DATASRCE", "OM_GLOBALBANDEDPARAM");
            xmlSchemaNode = xmlRootNode.appendChild(xmlElem);

            xmlElem = xmlDoc.createElement("NAME");
            xmlElem.setAttribute("DATATYPE", "STRING");
            xmlElem.setAttribute("LENGTH", "30");
            xmlSchemaNode.appendChild(xmlElem);

            switch (strValueType.ToUpper())
            {
                case "STRING":
                    xmlElem = xmlDoc.createElement("STRING");
                    xmlElem.setAttribute("DATATYPE", "STRING");
                    xmlElem.setAttribute("LENGTH", "30");
                    xmlSchemaNode.appendChild(xmlElem);
                    break;
                case "AMOUNT":
                    xmlElem = xmlDoc.createElement("AMOUNT");
                    xmlElem.setAttribute("DATATYPE", "DOUBLE");
                    xmlSchemaNode.appendChild(xmlElem);
                    break;
                case "PERCENTAGE":
                    xmlElem = xmlDoc.createElement("PERCENTAGE");
                    xmlElem.setAttribute("DATATYPE", "DOUBLE");
                    xmlSchemaNode.appendChild(xmlElem);
                    break;
                case "MAXIMUM":
                    xmlElem = xmlDoc.createElement("MAXIMUM");
                    xmlElem.setAttribute("DATATYPE", "DOUBLE");
                    xmlSchemaNode.appendChild(xmlElem);
                    break;
                case "BOOLEAN":
                    xmlElem = xmlDoc.createElement("BOOLEAN");
                    xmlElem.setAttribute("DATATYPE", "BOOLEAN");
                    xmlSchemaNode.appendChild(xmlElem);
                    break; 
            }

            // create REQUEST details
            xmlElem = xmlDoc.createElement("GLOBALPARAM");
            xmlElem.setAttribute("NAME", vstrParamName);
            xmlRequestNode = xmlRootNode.appendChild(xmlElem);
            if (vxmlResponse == null)
            {
                vxmlResponse = xmlDoc.createElement("RESPONSE");
            }
            adoAssistEx.adoGetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, vxmlResponse, String.Empty, String.Empty);

            xmlDoc = null;
            xmlElem = null;
            xmlRootNode = null;
            xmlSchemaNode = null;
            xmlRequestNode = null;
        }


        // DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
        // SR     BBGRb74 25/01/2006 New method GetAllGlobalBandedParamValuesAsXml
        private enum GLOBALPARAMTYPE {
            oegptString,
            oegptAmount,
            oegptMaximumAmount,
            oegptPercentage,
            oegptBoolean,
        }


    }

}

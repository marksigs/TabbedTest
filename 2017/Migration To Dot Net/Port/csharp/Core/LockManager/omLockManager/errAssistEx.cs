using System;
using MSXML2;
using Microsoft.VisualBasic;
namespace omLockManager
{
    public class errAssistEx
    {

        // header ----------------------------------------------------------------------------------
        // Workfile:      errAssist.bas
        // Copyright:     Copyright © 2001 Marlborough Stirling
        // 
        // Description:   Helper functions for error handling
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // History:
        // 
        // Prog    Date        Description
        // IK      01/11/00    Initial creation
        // ASm     10/01/01    SYS1817: Rationalisation of methods following meeting with PC and IK
        // PSC     25/02/02    SYS4097: Amend errGetMessage text to default message to not found
        // ------------------------------------------------------------------------------------------
        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        // BMIDS Specific History
        // 
        // Prog   Date        Description
        // DB     21/03/2003  BM0483  Performance upgrades - generate correct omiga error numbers
        // INR    28/03/2003  BM0056  Use Omiga Error No. to get message text.
        // 
        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        // ******************************************************************************************
        // BBG Specific History
        // 
        // Prog   Date        Description
        // MC     10/08/2004  E2EM00000558 - member function errAddWarningAndThrow() ADDED
        // PSC    12/08/2004  BBG1213 - Amend errThrowError to cater for extra details not being passed
        // 
        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        // ******************************************************************************************
        // Core Specific History
        // 
        // Prog   Date        Description
        // AS     20/05/2005  CORE135 Added support for getting error via CRUD
        // AS     07/03/2007  CORE9 Add ExceptionType support.
        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        private const string gstrMODULEPREFIX = "errAssist.";
        // AS 07/03/2007 CORE9 START Add ExceptionType
        private const string gstrEXCEPTIONTYPEPREFIX = "ExceptionType=";
        // AS 07/03/2007 CORE9 END Add ExceptionType
        public static void errThrowError(string vstrFunctionName, short vintOmigaErrorNo, params string[] vstrAdditionalOptions) 
        {
            const string strFunctionName = "errThrowError";
            string strText = String.Empty;
            string strLeftHandSide = String.Empty;
            string strMessage = String.Empty;
            int lngPosition = 0;
            short intIndex = 0;
            int lngErrNo = 0;
            // header ----------------------------------------------------------------------------------
            // description:  Raise a VB error
            // pass:         Function Name of calling component
            // Omiga Error number to throw
            // Parameter Array of addition arguments, if the error description
            // has subsitution parameters "%s" in it then you must specify the
            // these additional subsituation parameters in the parameter array
            // from 1 to n.
            // NOTE: The FIRST element in the parameter array (element 0_ is not
            // used in the errpor descprition subsitution but is used as
            // final addtion error text.
            // return:       n/a
            // ------------------------------------------------------------------------------------------

            lngErrNo = 0x80040000 + 512 + vintOmigaErrorNo;
            strText = errGetMessageText(vintOmigaErrorNo, MESSAGE.omMESSAGE_TEXT);
            // Find %s and substitute it with the substitution parameters
            lngPosition = Strings.InStr(1, strText, "%s", ((Microsoft.VisualBasic.CompareMethod)(CompareMethod.Text)));
            intIndex = (short)1;
            while(lngPosition != 0 && intIndex <= Information.UBound(vstrAdditionalOptions, 1))
            {
                strLeftHandSide = strText.Substring(0, lngPosition - 1);
                strMessage = strMessage + strLeftHandSide;
                // Substitute parameter if present
                if (!(Information.IsNothing(vstrAdditionalOptions[intIndex])))
                {
                    strMessage = strMessage + vstrAdditionalOptions[intIndex];
                }
                else
                {
                    strMessage = strMessage + "*** MISSING PARAMETER ***";
                }
                strText = strText.Substring(lngPosition + 2- 1);
                lngPosition = Strings.InStr(1, strText, "%s", ((Microsoft.VisualBasic.CompareMethod)(CompareMethod.Text)));
                intIndex = (short)(intIndex + 1);
            } 
            strMessage = strMessage + strText;
            // If we have additional parameters
            if (Information.UBound(vstrAdditionalOptions, 1) >= 0)
            {
                // First parameter is the additional error text
                // PSC 12/08/2004 BBG1213 - start
                if (!(Information.IsNothing(vstrAdditionalOptions[0])))
                {
                    strMessage = strMessage + ", Details: " + vstrAdditionalOptions[0];
                }
                // PSC 12/08/2004 BBG1213 - End
            }
            Information.Err().Raise(lngErrNo, vstrFunctionName, strMessage, null, null);
        }

        public static bool errIsWarning() 
        {
            string strErrorType = String.Empty;
            int lngOmigaErrorNo = 0;
            bool blnIsWarning = false;
            // header ----------------------------------------------------------------------------------
            // description:  Check if the raised error is a warning
            // pass:         n/a
            // return:       Boolean indicating if Err.Number corresponds to an Omiga4 Warning
            // ------------------------------------------------------------------------------------------
            blnIsWarning = false;
            if (errIsApplicationError())
            {

                lngOmigaErrorNo = errGetOmigaErrorNumber(Information.Err().Number);
                strErrorType = errGetMessageText(lngOmigaErrorNo, (short)(MESSAGE.omMESSAGE_TYPE));
                if (String.Compare(strErrorType, "Warning", Convert.ToBoolean(CompareMethod.Text)) == 0 || 
                    String.Compare(strErrorType, "Prompt", Convert.ToBoolean(CompareMethod.Text)) == 0)
                {
                    blnIsWarning = true;
                }
            }
            return blnIsWarning;
        }

        public static short errCheckError(string vstrFunctionName, string vstrObjectName, ref MSXML2.IXMLDOMElement vxmlResponseNode) 
        {
            short errCheckError = 0;
            string strExceptionType = String.Empty;
            int intErrNumber = 0;
            string strErrDescription = String.Empty;
            string strErrSource = String.Empty;
            string strHelpFile = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Re-raise the Application Error
            // pass:         Function Name of calling object
            // Object Name of calling object
            // xmlResponse node to append warning (only handles a single warning!) to
            // return:       Enumeration indicating if Err.Number corresponds to:
            // omWARNING or omNO_ERR see OMIGAMESSAGETYPE enum
            // ------------------------------------------------------------------------------------------
            errCheckError = (short)(OMIGAERRORTYPE.omNO_ERR);
            // AS 07/03/2007 CORE9 START Add ExceptionType
            strExceptionType = errGetExceptionType(Information.Err());
            if (Information.Err().Number != 0 || Strings.Len(strExceptionType) > 0)
            {
                // AS 07/03/2007 CORE9 END Add ExceptionType
                // AS 07/03/2007 CORE9 Add ExceptionType
                // Save err information as errIsWarning call MessageBO/DO which will reset
                // the error handler
                intErrNumber = Information.Err().Number;
                strErrDescription = Information.Err().Description;
                strErrSource = Information.Err().Source;
                strHelpFile = Information.Err().HelpFile;
                // AS 07/03/2007 CORE9 Add ExceptionType
                if (errIsWarning())
                {
                    if (vxmlResponseNode == null)
                    {
                        errThrowError(strErrSource, (short)(OMIGAERROR.oeXMLMissingElement), "Missing RESPONSE element");
                    }
                    errAddWarning(vxmlResponseNode);
                    // Add warning to the response node
                    errCheckError = (short)(OMIGAERRORTYPE.omWARNING);
                }
                else
                {
                    if (strErrSource != vstrFunctionName)
                    {
                        strErrSource = vstrFunctionName + "." + strErrSource;
                    }

                    if (Strings.Len(vstrObjectName) != 0)
                    {
                        strErrSource = vstrObjectName + "." + strErrSource;
                    }

                    Information.Err().Raise(intErrNumber, strErrSource, strErrDescription, strHelpFile, null);
                    // AS 07/03/2007 CORE9 Add ExceptionType
                }
            }
            return errCheckError;
        }

        public static int errCheckXMLResponseNode(MSXML2.IXMLDOMElement vxmlResponseNodeToCheck, MSXML2.IXMLDOMElement vxmlResponseNodeToAdd, bool vblnRaiseError) 
        {
            int lngErrNo = 0;
            string strErrSource = String.Empty;
            string strErrDesc = String.Empty;
            MSXML2.IXMLDOMNode xmlResponseTypeNode = null;
            MSXML2.IXMLDOMNodeList xmlMessageList = null;
            MSXML2.IXMLDOMElement xmlMessageElem = null;
            MSXML2.IXMLDOMElement xmlFirstChild = null;
            MSXML2.IXMLDOMNode xmlResponseErrorNumber = null;
            MSXML2.IXMLDOMNode xmlResponseErrorSource = null;
            MSXML2.IXMLDOMNode xmlResponseErrorDesc = null;
            // header ----------------------------------------------------------------------------------
            // description:  Re-raise the Application Error held in an IXMLDOMNode
            // pass:         xmlResponse node to check for error information
            // xmlResponse node to add warning(s) to
            // Raise Error boolean as to whether or not you require to re-raise the error
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            lngErrNo = OMIGAERROR.oeUnspecifiedError;
            strErrSource = App.Title;
            if (vxmlResponseNodeToCheck == null)
            {
                errThrowError(strErrSource, (short)(OMIGAERROR.oeXMLMissingElement), "Missing RESPONSE element");
            }
            if (vxmlResponseNodeToCheck.nodeName != "RESPONSE")
            {
                errThrowError(strErrSource, (short)(OMIGAERROR.oeMissingPrimaryTag), "RESPONSE must be the top level tag");
            }
            if (!(vxmlResponseNodeToAdd == null))
            {
                if (vxmlResponseNodeToAdd.nodeName != "RESPONSE")
                {
                    errThrowError(strErrSource, (short)(OMIGAERROR.oeMissingPrimaryTag), "RESPONSE must be the top level tag");
                }
            }
            xmlResponseTypeNode = vxmlResponseNodeToCheck.attributes.getNamedItem("TYPE");
            if (!(xmlResponseTypeNode == null))
            {
                if (xmlResponseTypeNode.text == "WARNING")
                {
                    if (!(vxmlResponseNodeToAdd == null))
                    {
                        vxmlResponseNodeToAdd.setAttribute("TYPE", "WARNING");
                        xmlFirstChild = vxmlResponseNodeToAdd.firstChild;
                        xmlMessageList = vxmlResponseNodeToCheck.selectNodes("MESSAGE");
                        // insert messages at the top of the response to add element
                        foreach(MSXML2.IXMLDOMElement __each1 in xmlMessageList)
                        {
                            xmlMessageElem = __each1;
                            vxmlResponseNodeToAdd.insertBefore(xmlMessageElem.cloneNode(true), xmlFirstChild);
                        }
                    }
                }
                else if( xmlResponseTypeNode.text != "SUCCESS")
                {


                    xmlResponseErrorNumber = vxmlResponseNodeToCheck.selectSingleNode("ERROR/NUMBER");
                    if (!(xmlResponseErrorNumber == null))
                    {
                        if (Information.IsNumeric(xmlResponseErrorNumber.text) == true)
                        {
                            lngErrNo = ConvertAssistEx.CSafeLng(xmlResponseErrorNumber.text);
                        }
                    }
                    xmlResponseErrorNumber = null;
                    if (vblnRaiseError)
                    {

                        xmlResponseErrorSource = vxmlResponseNodeToCheck.selectSingleNode("ERROR/SOURCE");
                        if (!(xmlResponseErrorSource == null))
                        {
                            if (Strings.Len(xmlResponseErrorSource.text) > 0)
                            {
                                strErrSource = xmlResponseErrorSource.text;
                            }
                        }
                        xmlResponseErrorSource = null;
                        xmlResponseErrorDesc = vxmlResponseNodeToCheck.selectSingleNode("ERROR/DESCRIPTION");
                        if (!(xmlResponseErrorDesc == null))
                        {
                            if (Strings.Len(xmlResponseErrorDesc.text) > 0)
                            {
                                strErrDesc = xmlResponseErrorDesc.text;
                            }
                        }
                        xmlResponseErrorDesc = null;
                        if (Strings.Len(strErrDesc) == 0)
                        {
                            strErrDesc = errGetMessageText(OMIGAERROR.oeUnspecifiedError, MESSAGE.omMESSAGE_TEXT);
                        }

                        // AS 07/03/2007 CORE9 START Add ExceptionType
                        Information.Err().Raise(lngErrNo, strErrSource, strErrDesc, errCreateHelpFileFromExceptionType(errGetExceptionTypeFromResponse(vxmlResponseNodeToCheck)), null);
                        // AS 07/03/2007 CORE9 END Add ExceptionType
                    }
                }
                else
                {
                    lngErrNo = 0;
                }
            }
            return lngErrNo;
        }

        public static MSXML2.IXMLDOMNode CreateErrorResponseNode(bool blnLogError) 
        {
            MSXML2.FreeThreadedDOMDocument40 xmlDoc = new MSXML2.FreeThreadedDOMDocument40();
            MSXML2.IXMLDOMElement xmlReponseElem = null;
            MSXML2.IXMLDOMElement xmlErrorElem = null;
            MSXML2.IXMLDOMElement xmlDescriptionElem = null;
            MSXML2.IXMLDOMElement xmlElement = null;
            string strExceptionType = String.Empty;
            string strId = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Creates an error node from the err object
            // pass:         n/a
            // return:       xml response node holding the error information for the error raised in the format:
            // <RESPONSE TYPE=APPERR or SYSERR>
            // <ERROR>
            // <NUMBER> </NUMBER>
            // <SOURCE> </SOURCE>
            // <DESCRIPTION> </DESCRIPTION>
            // </ERROR>
            // </RESPONSE>
            // AS 07/03/2007 CORE9 Note that the ordering of ERROR child elements has changed, so that
            // SOURCE (the longest string) is last.
            // ------------------------------------------------------------------------------------------
            xmlReponseElem = xmlDoc.createElement("RESPONSE");
            xmlDoc.appendChild(xmlReponseElem);
            xmlErrorElem = xmlDoc.createElement("ERROR");
            xmlReponseElem.appendChild(xmlErrorElem);
            xmlElement = xmlDoc.createElement("NUMBER");
            xmlElement.text = Convert.ToString(Information.Err().Number);
            xmlErrorElem.appendChild(xmlElement);
            // AS 07/03/2007 CORE9 START Add ExceptionType
            strExceptionType = errGetExceptionType(Information.Err());
            if (Strings.Len(strExceptionType) > 0)
            {
                // Only add EXCEPTIONTYPE and ID elements if Err.HelpFile contains "ExceptionType=" prefix,
                // i.e., this error has been raised by Vertex.Fsd.Omiga.Core.OmigaException.
                xmlElement = xmlDoc.createElement("EXCEPTIONTYPE");
                xmlElement.text = strExceptionType;
                xmlErrorElem.appendChild(xmlElement);
                strId = guidAssistEx.CreateGUID();
                xmlElement = xmlDoc.createElement("ID");
                xmlElement.text = strId;
                xmlErrorElem.appendChild(xmlElement);
            }
            // AS 07/03/2007 CORE9 END Add ExceptionType
            xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION");
            xmlDescriptionElem.text = Information.Err().Description;
            xmlErrorElem.appendChild(xmlDescriptionElem);
            xmlElement = xmlDoc.createElement("VERSION");
            xmlElement.text = ;
            xmlErrorElem.appendChild(xmlElement);
            xmlElement = xmlDoc.createElement("SOURCE");
            xmlElement.text = Information.Err().Source;
            xmlErrorElem.appendChild(xmlElement);
            if (errIsApplicationError() == true)
            {
                xmlReponseElem.setAttribute("TYPE", "APPERR");
                if (Strings.Len(xmlDescriptionElem.text) == 0)
                {
                    // INR BM0056
                    // xmlDescriptionElem.Text = errGetMessageText(Err.Number)
                    xmlDescriptionElem.text = errGetMessageText(errGetOmigaErrorNumber(Information.Err().Number), MESSAGE.omMESSAGE_TEXT);
                }
            }
            else
            {
                xmlReponseElem.setAttribute("TYPE", "SYSERR");
            }

            if (blnLogError)
            {
                if (errIsSystemError() == true)
                {
                    // AS 07/03/2007 CORE9 START Add ExceptionType
                    // ISSUE: Method or data member not found: 'LogEvent'
                    App.LogEvent(
                        "SYSERR: " + Information.Err().Number + ", " + 
                        Interaction.IIf(Strings.Len(strExceptionType) > 0, strExceptionType + ", ", "") + 
                        Interaction.IIf(Strings.Len(strId) > 0, strId + ", ", "") + 
                        Information.Err().Description + ", " + 
                        Information.Err().Source, 
                        1);
                    // AS 07/03/2007 CORE9 END Add ExceptionType
                }
            }
            xmlDoc = null;
            xmlReponseElem = null;
            xmlErrorElem = null;
            xmlDescriptionElem = null;
            xmlElement = null;
            return xmlReponseElem.cloneNode(true);
        }

        public static string errCreateErrorResponse() 
        {
            MSXML2.IXMLDOMElement xmlErrorElem = null;
            string strExceptionType = String.Empty;
            string strId = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Creates an xml error response based on the raised error
            // pass:         n/a
            // return:       The xml response string containing an the error information in the format
            // <RESPONSE TYPE=APPERR or SYSERR>
            // <ERROR>
            // <NUMBER> </NUMBER>
            // <SOURCE> </SOURCE>
            // <DESCRIPTION> </DESCRIPTION>
            // </ERROR>
            // </RESPONSE>
            // ------------------------------------------------------------------------------------------
            xmlErrorElem = CreateErrorResponseNode(false);
            if (errIsSystemError() == true)
            {
                // AS 07/03/2007 CORE9 START Add ExceptionType
                strExceptionType = errGetExceptionType(Information.Err());
                if (Strings.Len(strExceptionType) > 0)
                {
                    strId = guidAssistEx.CreateGUID();
                }
                // ISSUE: Method or data member not found: 'LogEvent'
                App.LogEvent(
                    "SYSERR: " + Information.Err().Number + ", " + 
                    Interaction.IIf(Strings.Len(strExceptionType) > 0, strExceptionType + ", ", "") + 
                    Interaction.IIf(Strings.Len(strId) > 0, strId + ", ", "") + 
                    Information.Err().Description + ", " + 
                    Information.Err().Source, 
                    1);
                // AS 07/03/2007 CORE9 END Add ExceptionType
            }

            xmlErrorElem = null;
            return xmlErrorElem.xml;
        }

        public static string errGetMessageText(int vlngMessageNo, short vintMessageField) 
        {
            string errGetMessageText = String.Empty;
            const string strFunctionName = "errGetMessageText";
            string strMessageText = String.Empty;
            MSXML2.FreeThreadedDOMDocument40 xmlDoc = null;
            MSXML2.IXMLDOMElement xmlElem = null;
            MSXML2.IXMLDOMNode xmlRootNode = null;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNode xmlRequestNode = null;
            MSXML2.IXMLDOMNode xmlResponseNode = null;
            // header ----------------------------------------------------------------------------------
            // description:  Gets the soft coded message information, the message can be an error or warning
            // pass:         Message number to use to find message text
            // Optional Message field to return either Text or Type see omMESSAGE enum
            // return:       Message text string in the following format:
            // ------------------------------------------------------------------------------------------
            // AS 20/05/2005 CORE135 Added support for getting error via CRUD
#if omCRUD
            // ISSUE: Function is not defined: 'adoCRUDGetMessageText'
            errGetMessageText = Convert.ToString(adoCRUDGetMessageText(vlngMessageNo, vintMessageField));
#else 
            // First see if the error is not on the database
            if (!(errGetDefaultMessageText(vlngMessageNo, ref strMessageText)))
            {

                // Get the message from the database
                // create one FreeThreadedDOMDocument40 to contain schema, request & response
                xmlDoc = new MSXML2.FreeThreadedDOMDocument40();
                // create BOGUS root node
                xmlElem = xmlDoc.createElement("BOGUS");
                xmlRootNode = xmlDoc.appendChild(xmlElem);
                // create schema for MESSAGE table
                xmlElem = xmlDoc.createElement("SCHEMA");
                xmlSchemaNode = xmlRootNode.appendChild(xmlElem);
                xmlElem = xmlDoc.createElement("MESSAGE");
                xmlElem.setAttribute("ENTITYTYPE", "PHYSICAL");
                xmlSchemaNode = xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlDoc.createElement("MESSAGENUMBER");
                xmlElem.setAttribute("DATATYPE", "SHORT");
                xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlDoc.createElement("MESSAGETEXT");
                xmlElem.setAttribute("DATATYPE", "STRING");
                xmlElem.setAttribute("LENGTH", "255");
                xmlSchemaNode.appendChild(xmlElem);
                // create REQUEST details
                xmlElem = xmlDoc.createElement("REQUEST");
                xmlRequestNode = xmlRootNode.appendChild(xmlElem);
                xmlElem = xmlDoc.createElement("MESSAGE");
                xmlElem.setAttribute("MESSAGENUMBER", Convert.ToString(vlngMessageNo));
                xmlRequestNode = xmlRequestNode.appendChild(xmlElem);
                // create RESPONSE node
                xmlElem = xmlDoc.createElement("RESPONSE");
                xmlResponseNode = xmlRootNode.appendChild(xmlElem);
                adoAssistEx.adoGetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlResponseNode, String.Empty, String.Empty);
                if (vintMessageField == MESSAGE.omMESSAGE_TEXT)
                {
                    if (!(xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETEXT") == null))
                    {
                        strMessageText = xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETEXT").text;
                        // Else
                        // strMessageText = "Error Message not found"
                    }
                }
                else
                {
                    if (!(xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETYPE") == null))
                    {
                        strMessageText = xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETYPE").text;
                    }
                }
                xmlDoc = null;
                xmlElem = null;
                xmlRootNode = null;
                xmlSchemaNode = null;
                xmlRequestNode = null;
                xmlResponseNode = null;
            }
            errGetMessageText = strMessageText;
#endif 
            return errGetMessageText;
        }

        private static bool errGetDefaultMessageText(int lngMessageNo, ref string strMsgText) 
        {
            bool errGetDefaultMessageText = false;
            // header ----------------------------------------------------------------------------------
            // description:  Gets all messages not specifed in the message table
            // pass:         message number and empty message text string
            // returns:      Boolean indicating whether the message number was found
            // message text if message is found
            // ------------------------------------------------------------------------------------------
            errGetDefaultMessageText = true;
            switch (lngMessageNo)
            {

                case 556:
                    strMsgText = "Unable to establish database connection";
                    break;
                case 901:
                    strMsgText = "Error message not found";
                    // ------------------------------------------------------------------------------------------
                    // catch all
                    // ------------------------------------------------------------------------------------------
                    break;
                default:
                    errGetDefaultMessageText = false;
                    break; 
            }
            return errGetDefaultMessageText;
        }

        public static bool errIsApplicationError() 
        {
            int lngOmigaErrorNo = 0;
            int lngErrNo = 0;
            bool blnIsApplicationError = false;
            // header ----------------------------------------------------------------------------------
            // description:  Check if error is an omiga 4 application error.
            // NOTE: Warnings are application errors!
            // pass:         n/a
            // return:       boolean indicating if the error is an applicaton error
            // ------------------------------------------------------------------------------------------
            blnIsApplicationError = false;
            lngErrNo = Information.Err().Number;

            // If lngErrNo <> 0 Then
            // lngOmigaErrorNo = errGetOmigaErrorNumber(lngErrNo)
            // 
            // ' AS 23/11/99 For the error to be an Omiga4 Application error it must also
            // ' not be in the ADO error number range see ADODB.ErrorValueEnum
            // If (lngOmigaErrorNo > 0 And _
            // '            lngOmigaErrorNo <= clngMAX_ERROR_NO And _
            // '            (lngOmigaErrorNo < clngADO_START_ERROR_NO Or _
            // '            lngOmigaErrorNo > clngADO_END_ERROR_NO)) Then
            // 
            // blnIsApplicationError = True
            // 
            // End If
            // End If
            // DB BM0483 - Re-worked to subtract 512 from the error number.
            if (lngErrNo != 0)
            {
                lngOmigaErrorNo = errGetOmigaErrorNumber(lngErrNo);
                if (lngOmigaErrorNo > 0 && 
                    lngOmigaErrorNo <= errConstants.clngMAX_ERROR_NO)
                {
                    lngErrNo = lngErrNo - 0x80040000;
                    if (lngErrNo < errConstants.clngADO_START_ERROR_NO || 
                        lngOmigaErrorNo > errConstants.clngADO_END_ERROR_NO)
                    {
                        blnIsApplicationError = true;
                    }
                }
            }
            // DB End
            return blnIsApplicationError;
        }

        public static bool errIsSystemError() 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Determines if the error is a system error
            // pass:         n/a
            // return:       boolean indicating a system error
            // ------------------------------------------------------------------------------------------
            return !(errIsApplicationError());
        }

        public static string errFormatMessageNode() 
        {
            MSXML2.FreeThreadedDOMDocument40 xmlMessageDoc = new MSXML2.FreeThreadedDOMDocument40();
            MSXML2.IXMLDOMElement xmlMessageElement = null;
            MSXML2.IXMLDOMElement xmlElement = null;
            int lngOmigaErrorNo = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // pass:
            // ------------------------------------------------------------------------------------------

            xmlMessageElement = xmlMessageDoc.createElement("MESSAGE");
            xmlMessageDoc.appendChild(xmlMessageElement);
            xmlElement = xmlMessageDoc.createElement("TEXT");
            xmlElement.text = Information.Err().Description;
            xmlMessageElement.appendChild(xmlElement);
            lngOmigaErrorNo = errGetOmigaErrorNumber(Information.Err().Number);
            xmlElement = xmlMessageDoc.createElement("TYPE");
            xmlElement.text = "Warning";
            xmlMessageElement.appendChild(xmlElement);
            xmlMessageDoc = null;
            xmlMessageElement = null;
            xmlElement = null;

            return xmlMessageDoc.xml;
        }

        public static void errAddWarning(MSXML2.IXMLDOMElement xmlResponse) 
        {
            MSXML2.IXMLDOMElement xmlMessageElement = null;
            MSXML2.IXMLDOMElement xmlElement = null;
            MSXML2.IXMLDOMNode xmlFirstChild = null;
            int lngOmigaErrorNo = 0;
            // header ----------------------------------------------------------------------------------
            // description:  Adds the warning into the xmlResponse
            // pass:         xmlResponse to add the warning to
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            if (xmlResponse.nodeName != "RESPONSE")
            {
                errThrowError(Information.Err().Source, (short)(OMIGAERROR.oeMissingPrimaryTag), "RESPONSE must be the top level tag");
            }
            xmlResponse.setAttribute("TYPE", "WARNING");
            xmlFirstChild = xmlResponse.firstChild;
            xmlMessageElement = xmlResponse.ownerDocument.createElement("MESSAGE");
            xmlResponse.insertBefore(xmlMessageElement, xmlFirstChild);

            xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETEXT");
            xmlElement.text = Information.Err().Description;
            xmlMessageElement.appendChild(xmlElement);
            lngOmigaErrorNo = errGetOmigaErrorNumber(Information.Err().Number);
            xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETYPE");
            xmlElement.text = errGetMessageText(lngOmigaErrorNo, (short)(MESSAGE.omMESSAGE_TYPE));
            xmlMessageElement.appendChild(xmlElement);
            xmlMessageElement = null;
            xmlElement = null;
            xmlFirstChild = null;
        }

        // [MC]E2EM00000558 - WARNING MESSAGE WITHOUT APPERROR
        public static void errAddOmigaWarning(MSXML2.IXMLDOMElement xmlResponse, short vintOmigaErrorNo) 
        {
            MSXML2.IXMLDOMElement xmlMessageElement = null;
            MSXML2.IXMLDOMElement xmlElement = null;
            MSXML2.IXMLDOMNode xmlFirstChild = null;
            int lngOmigaErrorNo = 0;
            int lngErrNo = 0;
            string strText = String.Empty;


            if (xmlResponse.nodeName != "RESPONSE")
            {
                errThrowError(Information.Err().Source, (short)(OMIGAERROR.oeMissingPrimaryTag), "RESPONSE must be the top level tag");
            }

            lngErrNo = 0x80040000 + 512 + vintOmigaErrorNo;
            strText = errGetMessageText(vintOmigaErrorNo, MESSAGE.omMESSAGE_TEXT);

            xmlResponse.setAttribute("TYPE", "WARNING");
            xmlFirstChild = xmlResponse.firstChild;
            xmlMessageElement = xmlResponse.ownerDocument.createElement("MESSAGE");
            xmlResponse.insertBefore(xmlMessageElement, xmlFirstChild);

            xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETEXT");
            xmlElement.text = strText;
            xmlMessageElement.appendChild(xmlElement);
            lngOmigaErrorNo = vintOmigaErrorNo;
            xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETYPE");
            xmlElement.text = Convert.ToString(vintOmigaErrorNo);
            xmlMessageElement.appendChild(xmlElement);
            xmlMessageElement = null;
            xmlElement = null;
            xmlFirstChild = null;

        }

        // *=[MC]E2EM00000558 - END
        public static int errGetOmigaErrorNumber(int vlngErrorNo) 
        {
            // header ----------------------------------------------------------------------------------
            // description:  Converts an  error number to an omiga number. When errors are raised by VB
            // they get added to them the VB constant 'vbObjectError + 512' so to get the
            // Omiga4 error number orginally raised we need to subtract them from the err
            // number.
            // NOTE: System error numbers will end up much larger than Omiga4 error numbers
            // pass:         Error number to calculate
            // return:       Calculated Omiga4 error number
            // ------------------------------------------------------------------------------------------
            return vlngErrorNo - 0x80040000 - 512;
        }

        public static int errCheckXMLResponse(string vstrXmlReponse, bool vblnRaiseError, MSXML2.IXMLDOMElement vxmlInResponseElement) 
        {
            string strFunctionName = String.Empty;
            MSXML2.FreeThreadedDOMDocument40 xmlXmlIn = null;
            MSXML2.IXMLDOMElement xmlResponseElement = null;
            // header ----------------------------------------------------------------------------------
            // description:  Checks the xml response, this method expects the xml response as a string
            // and designed to called when calling across component methods and then intergrating
            // the returned xml
            // pass:         xmlResponse containing xml data to be checked
            // boolean as to whether to re-raise the error
            // xmlResponseElement to copy the warnings to
            // ------------------------------------------------------------------------------------------
            strFunctionName = "CheckXMLResponse";
            xmlXmlIn = xmlAssistEx.xmlLoad(vstrXmlReponse, strFunctionName);
            xmlResponseElement = xmlXmlIn.selectSingleNode("RESPONSE");
            if (xmlResponseElement == null)
            {
                errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLMissingElement), "Missing RESPONSE element");
            }
            xmlXmlIn = null;
            xmlResponseElement = null;
            return errCheckXMLResponseNode(xmlResponseElement, vxmlInResponseElement, vblnRaiseError);
        }

        public static void errRaiseXMLResponse(string vstrXmlResponse) 
        {
            string strFunctionName = String.Empty;
            MSXML2.FreeThreadedDOMDocument40 xmlXmlIn = null;
            // header ----------------------------------------------------------------------------------
            // description:  Checks the XML response for error information and raises the error
            // pass:         xmlResponse containing xml data to be checked for an error
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            strFunctionName = "RaiseXMLResponseError";
            xmlXmlIn = xmlAssistEx.xmlLoad(vstrXmlResponse, strFunctionName);
            errRaiseXMLResponseNode(xmlXmlIn.documentElement);
        }

        public static void errRaiseXMLResponseNode(MSXML2.IXMLDOMElement vxmlResponseNode) 
        {
            MSXML2.IXMLDOMNode xmlResponseErrorNumber = null;
            MSXML2.IXMLDOMNode xmlResponseErrorSource = null;
            MSXML2.IXMLDOMNode xmlResponseErrorDesc = null;
            MSXML2.IXMLDOMNode xmlResponseErrorExceptionType = null;
            int lngErrNo = 0;
            string strExceptionType = String.Empty;
            string strHelpFile = String.Empty;
            string strErrSource = String.Empty;
            string strErrDesc = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Checks the XML response node for error information and raises the error
            // pass:         vxmlResponseNode containing xml data to be checked for an error
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            // AS 07/03/2007 CORE9 Add ExceptionType
            xmlResponseErrorNumber = vxmlResponseNode.selectSingleNode("ERROR/NUMBER");
            if (!(xmlResponseErrorNumber == null))
            {
                if (Information.IsNumeric(xmlResponseErrorNumber.text) == true)
                {
                    lngErrNo = ConvertAssistEx.CSafeLng(xmlResponseErrorNumber.text);
                }
            }
            xmlResponseErrorNumber = null;

            // AS 07/03/2007 CORE9 START Add ExceptionType
            strExceptionType = errGetExceptionTypeFromResponse(vxmlResponseNode);

            if (lngErrNo != 0 || Strings.Len(strExceptionType) > 0)
            {
                strHelpFile = errCreateHelpFileFromExceptionType(strExceptionType);
                // AS 07/03/2007 CORE9 END Add ExceptionType
                xmlResponseErrorSource = vxmlResponseNode.selectSingleNode("ERROR/SOURCE");
                if (!(xmlResponseErrorSource == null))
                {
                    if (Strings.Len(xmlResponseErrorSource.text) > 0)
                    {
                        strErrSource = xmlResponseErrorSource.text;
                    }
                }
                xmlResponseErrorSource = null;
                xmlResponseErrorDesc = vxmlResponseNode.selectSingleNode("ERROR/DESCRIPTION");
                if (!(xmlResponseErrorDesc == null))
                {
                    if (Strings.Len(xmlResponseErrorDesc.text) > 0)
                    {
                        strErrDesc = xmlResponseErrorDesc.text;
                    }
                }
                xmlResponseErrorDesc = null;
                if (Strings.Len(strErrDesc) == 0)
                {
                    strErrDesc = errGetMessageText(OMIGAERROR.oeUnspecifiedError, MESSAGE.omMESSAGE_TEXT);
                }

                Information.Err().Raise(lngErrNo, strErrSource, strErrDesc, strHelpFile, null);
                // AS 07/03/2007 CORE9 Add ExceptionType
            }
        }

        // AS 07/03/2007 CORE9 START Add ExceptionType
        public static string errGetExceptionType(ref System.Exception objErr) 
        {
            string strHelpFile = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Gets any exception type from a VB error object.
            // pass:         An optional VB error object (if not supplied then Err is used).
            // return:       The exception type, as a string.
            // Exception type information is held in the Err.HelpFile property, in the format, e.g.,
            // "ExceptionType=Vertex.Fsd.Omiga.Core.OmigaException".
            // ------------------------------------------------------------------------------------------
            if (objErr == null)
            {
                strHelpFile = Information.Err().HelpFile;
            }
            else
            {
                // ISSUE: Method or data member not found: 'HelpFile'
                strHelpFile = Convert.ToString(objErr.HelpFile);
            }

            if (Strings.Len(strHelpFile) >= Strings.Len(gstrEXCEPTIONTYPEPREFIX))
            {
                if ((strHelpFile.Substring(0, Strings.Len(gstrEXCEPTIONTYPEPREFIX))).ToLower() == gstrEXCEPTIONTYPEPREFIX.ToLower())
                {
                }
            }
            return strHelpFile.Substring(strHelpFile.Length - Strings.Len(strHelpFile) - Strings.Len(gstrEXCEPTIONTYPEPREFIX));
        }

        public static string errGetExceptionTypeFromResponse(MSXML2.IXMLDOMElement vxmlResponseNode) 
        {
            MSXML2.IXMLDOMNode xmlResponseErrorExceptionType = null;
            xmlResponseErrorExceptionType = vxmlResponseNode.selectSingleNode("ERROR/EXCEPTIONTYPE");
            if (!(xmlResponseErrorExceptionType == null))
            {
                if (Strings.Len(xmlResponseErrorExceptionType.text) > 0)
                {
                }
            }
            xmlResponseErrorExceptionType = null;
            return xmlResponseErrorExceptionType.text;
        }

        public static bool errIsExceptionType(string strCheckExceptionType, ref System.Exception objErr) 
        {
            string strExceptionType = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  Checks whether an error is of a specified type.
            // pass:         The exception type to check, and an optional VB error object
            // (if not supplied then Err is used).
            // return:       True if the exception type matches.
            // ------------------------------------------------------------------------------------------
            strExceptionType = errGetExceptionType(ref objErr);
            if (Strings.Len(strExceptionType) >= Strings.Len(strCheckExceptionType))
            {
                // Must be case sensitive as exception types names are case sensitive, i.e.,
                // TESTException is a different type to TestException.
                // Note: avoid type names that differ only by case.
                if (strExceptionType.Substring(strExceptionType.Length - Strings.Len(strCheckExceptionType)) == strCheckExceptionType)
                {
                }
            }
            return true;
        }

        private static string errCreateHelpFileFromExceptionType(string strExceptionType) 
        {
            return Interaction.IIf(Strings.Len(strExceptionType) > 0, gstrEXCEPTIONTYPEPREFIX + strExceptionType, "");
        }


    }

}

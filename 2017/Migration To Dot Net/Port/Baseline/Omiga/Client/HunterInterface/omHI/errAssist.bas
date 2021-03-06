Attribute VB_Name = "errAssistEx"
'header ----------------------------------------------------------------------------------
'Workfile:      errAssist.bas
'Copyright:     Copyright � 2001 Marlborough Stirling
'
'Description:   Helper functions for error handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog    Date        Description
'IK      01/11/00    Initial creation
'ASm     10/01/01    SYS1817: Rationalisation of methods following meeting with PC and IK
'PSC     25/02/02    SYS4097: Amend errGetMessage text to default message to not found
'------------------------------------------------------------------------------------------
'BMIDS Specific History
'
'Prog   Date        Description
'DB     21/03/2003  BM0483  Performance upgrades - generate correct omiga error numbers
'INR    28/03/2003  BM0056  Use Omiga Error No. to get message text.
'HMA    21/12/2004  BMIDS958  Updated errGetMessageText to reduce database queries.
'------------------------------------------------------------------------------------------
'MARS Specific History
'
'Prog   Date        Description
'PSC    27/10/2005  MARS300 Correct processing of warnings
'GHun   13/03/2006  MAR1300 Merge core changes
'------------------------------------------------------------------------------------------
Option Explicit
   
Public Sub errThrowError( _
    ByVal vstrFunctionName As String, _
    ByVal vintOmigaErrorNo As Integer, _
    ParamArray vstrAdditionalOptions())
    
    Const strFunctionName As String = "errThrowError"
    
' header ----------------------------------------------------------------------------------
' description:  Raise a VB error
' pass:         Function Name of calling component
'               Omiga Error number to throw
'               Parameter Array of addition arguments, if the error description
'               has subsitution parameters "%s" in it then you must specify the
'               these additional subsituation parameters in the parameter array
'               from 1 to n.
'               NOTE: The FIRST element in the parameter array (element 0_ is not
'               used in the errpor descprition subsitution but is used as
'               final addtion error text.
' return:       n/a
'------------------------------------------------------------------------------------------
    
    Dim strText As String
    Dim strLeftHandSide As String
    Dim strMessage As String
    Dim lngPosition As Long
    Dim intIndex As Integer
    Dim lngErrNo As Long
        
    lngErrNo = vbObjectError + 512 + vintOmigaErrorNo
    
    strText = errGetMessageText(vintOmigaErrorNo)
    
    ' Find %s and substitute it with the substitution parameters
    lngPosition = InStr(1, strText, "%s", vbTextCompare)
    
    intIndex = 1
    
    Do While lngPosition <> 0 And intIndex <= UBound(vstrAdditionalOptions)
        strLeftHandSide = Left$(strText, lngPosition - 1)
        strMessage = strMessage & strLeftHandSide
        
        ' Substitute parameter if present
        If Not IsMissing(vstrAdditionalOptions(intIndex)) Then
            strMessage = strMessage & vstrAdditionalOptions(intIndex)
        Else
            strMessage = strMessage & "*** MISSING PARAMETER ***"
        End If
        
        strText = Mid$(strText, lngPosition + 2)
        lngPosition = InStr(1, strText, "%s", vbTextCompare)
        intIndex = intIndex + 1
    Loop
    
    strMessage = strMessage & strText
    
    ' If we have additional parameters
    If UBound(vstrAdditionalOptions) >= 0 Then
        ' First parameter is the additional error text
        If Not IsMissing(vstrAdditionalOptions(0)) Then 'MAR1300 GHun
            strMessage = strMessage & ", Details: " & vstrAdditionalOptions(0)
        End If
    End If
    
    Err.Raise lngErrNo, vstrFunctionName, strMessage

End Sub

Public Function errIsWarning() As Boolean
' header ----------------------------------------------------------------------------------
' description:  Check if the raised error is a warning
' pass:         n/a
' return:       Boolean indicating if Err.Number corresponds to an Omiga4 Warning
'------------------------------------------------------------------------------------------
    
    Dim strErrorType As String
    Dim lngOmigaErrorNo As Long
    Dim blnIsWarning As Boolean
    
    blnIsWarning = False
    
    If errIsApplicationError() Then
        
        lngOmigaErrorNo = errGetOmigaErrorNumber(Err.Number)
        strErrorType = errGetMessageText(lngOmigaErrorNo, omMESSAGE_TYPE)
        
        If StrComp(strErrorType, "Warning", vbTextCompare) = 0 Or _
           StrComp(strErrorType, "Prompt", vbTextCompare) = 0 Then
            blnIsWarning = True
        End If
    End If
    
    errIsWarning = blnIsWarning
    
End Function

Public Function errCheckError( _
    ByVal vstrFunctionName As String, _
    Optional ByVal vstrObjectName As String, _
    Optional ByRef vxmlResponseNode As IXMLDOMElement) As Integer
' header ----------------------------------------------------------------------------------
' description:  Re-raise the Application Error
' pass:         Function Name of calling object
'               Object Name of calling object
'               xmlResponse node to append warning (only handles a single warning!) to
' return:       Enumeration indicating if Err.Number corresponds to:
'               omWARNING or omNO_ERR see OMIGAMESSAGETYPE enum
'------------------------------------------------------------------------------------------
    
    errCheckError = omNO_ERR
    
    If Err.Number <> 0 Then
    
        Dim intErrNumber As Long
        Dim strErrDescription As String
        Dim strErrSource As String
        
        ' Save err information as errIsWarning call MessageBO/DO which will reset
        ' the error handler
        intErrNumber = Err.Number
        strErrDescription = Err.Description
        strErrSource = Err.Source
        
        If errIsWarning() Then
            If vxmlResponseNode Is Nothing Then
                errThrowError strErrSource, oeXMLMissingElement, "Missing RESPONSE element"
            End If
            ' PSC 27/10/2005 MAR300
            errAddWarning vxmlResponseNode, intErrNumber, strErrSource, strErrDescription    ' Add warning to the response node
            errCheckError = omWARNING
        Else
        
            If strErrSource <> vstrFunctionName Then
                strErrSource = vstrFunctionName & "." & strErrSource
            End If
                
            If Len(vstrObjectName) <> 0 Then
                strErrSource = vstrObjectName & "." & strErrSource
            End If
                
            Err.Raise intErrNumber, strErrSource, strErrDescription
            
        End If
    
    End If

End Function

Public Function errCheckXMLResponseNode(ByVal vxmlResponseNodeToCheck As IXMLDOMElement, _
                            Optional ByVal vxmlResponseNodeToAdd As IXMLDOMElement, _
                            Optional ByVal vblnRaiseError As Boolean = False) As Long
' header ----------------------------------------------------------------------------------
' description:  Re-raise the Application Error held in an IXMLDOMNode
' pass:         xmlResponse node to check for error information
'               xmlResponse node to add warning(s) to
'               Raise Error boolean as to whether or not you require to re-raise the error
' return:       n/a
'------------------------------------------------------------------------------------------

    Dim lngErrNo As Long
    Dim strErrSource As String, _
        strErrDesc As String
        
    lngErrNo = oeUnspecifiedError
    strErrSource = App.Title

    If vxmlResponseNodeToCheck Is Nothing Then
        errThrowError strErrSource, oeXMLMissingElement, "Missing RESPONSE element"
    End If
    
    If vxmlResponseNodeToCheck.nodeName <> "RESPONSE" Then
        errThrowError strErrSource, oeMissingPrimaryTag, "RESPONSE must be the top level tag"
    End If
    
    If Not vxmlResponseNodeToAdd Is Nothing Then
        If vxmlResponseNodeToAdd.nodeName <> "RESPONSE" Then
            errThrowError strErrSource, oeMissingPrimaryTag, "RESPONSE must be the top level tag"
        End If
    End If
    
    Dim xmlResponseTypeNode As IXMLDOMNode
    Set xmlResponseTypeNode = vxmlResponseNodeToCheck.Attributes.getNamedItem("TYPE")
    
    If Not xmlResponseTypeNode Is Nothing Then
        If xmlResponseTypeNode.Text = "WARNING" Then
            If Not vxmlResponseNodeToAdd Is Nothing Then
                vxmlResponseNodeToAdd.setAttribute "TYPE", "WARNING"
    
                Dim xmlMessageList As IXMLDOMNodeList
                Dim xmlMessageElem As IXMLDOMElement
                Dim xmlFirstChild  As IXMLDOMElement
                
                Set xmlFirstChild = vxmlResponseNodeToAdd.firstChild
                Set xmlMessageList = vxmlResponseNodeToCheck.selectNodes("MESSAGE")
    
                ' insert messages at the top of the response to add element
                For Each xmlMessageElem In xmlMessageList
                    vxmlResponseNodeToAdd.insertBefore xmlMessageElem.cloneNode(True), xmlFirstChild
                Next
            End If
        ElseIf xmlResponseTypeNode.Text <> "SUCCESS" Then
            
            Dim xmlResponseErrorNumber As IXMLDOMNode, _
                xmlResponseErrorSource As IXMLDOMNode, _
                xmlResponseErrorDesc As IXMLDOMNode
            
            Set xmlResponseErrorNumber = vxmlResponseNodeToCheck.selectSingleNode("ERROR/NUMBER")
            If Not xmlResponseErrorNumber Is Nothing Then
                If IsNumeric(xmlResponseErrorNumber.Text) = True Then
                    lngErrNo = CSafeLng(xmlResponseErrorNumber.Text)
                End If
            End If
            Set xmlResponseErrorNumber = Nothing
            
            If vblnRaiseError Then
                
                Set xmlResponseErrorSource = vxmlResponseNodeToCheck.selectSingleNode("ERROR/SOURCE")
                If Not xmlResponseErrorSource Is Nothing Then
                    If Len(xmlResponseErrorSource.Text) > 0 Then
                        strErrSource = xmlResponseErrorSource.Text
                    End If
                End If
                Set xmlResponseErrorSource = Nothing
                
                Set xmlResponseErrorDesc = vxmlResponseNodeToCheck.selectSingleNode("ERROR/DESCRIPTION")
                If Not xmlResponseErrorDesc Is Nothing Then
                    If Len(xmlResponseErrorDesc.Text) > 0 Then
                        strErrDesc = xmlResponseErrorDesc.Text
                    End If
                End If
                Set xmlResponseErrorDesc = Nothing
                
                If Len(strErrDesc) = 0 Then
                    strErrDesc = errGetMessageText(oeUnspecifiedError)
                End If
                            
                Err.Raise lngErrNo, strErrSource, strErrDesc
                
            End If
        Else
            lngErrNo = 0
        End If
    End If
    
    errCheckXMLResponseNode = lngErrNo
    
End Function

Public Function CreateErrorResponseNode(Optional ByVal blnLogError As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates an error node from the err object
' pass:         n/a
' return:       xml response node holding the error information for the error raised in the format:
'               <RESPONSE TYPE=APPERR or SYSERR>
'                   <ERROR>
'                       <NUMBER> </NUMBER>
'                       <SOURCE> </SOURCE>
'                       <DESCRIPTION> </DESCRIPTION>
'                   </ERROR>
'               </RESPONSE>
'------------------------------------------------------------------------------------------
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlReponseElem As IXMLDOMElement
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlDescriptionElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlReponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlReponseElem
    
    Set xmlErrorElem = xmlDoc.createElement("ERROR")
    xmlReponseElem.appendChild xmlErrorElem
    
    Set xmlElement = xmlDoc.createElement("NUMBER")
    xmlElement.Text = Err.Number
    xmlErrorElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("SOURCE")
    xmlElement.Text = Err.Source
    xmlErrorElem.appendChild xmlElement
    
    Set xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION")
    xmlDescriptionElem.Text = Err.Description
    xmlErrorElem.appendChild xmlDescriptionElem
    
    Set xmlElement = xmlDoc.createElement("VERSION")
    xmlElement.Text = App.Comments
    xmlErrorElem.appendChild xmlElement
    
    If errIsApplicationError() = True Then
        xmlReponseElem.setAttribute "TYPE", "APPERR"
    
        If Len(xmlDescriptionElem.Text) = 0 Then
            'INR BM0056
'            xmlDescriptionElem.Text = errGetMessageText(Err.Number)
            xmlDescriptionElem.Text = errGetMessageText(errGetOmigaErrorNumber(Err.Number))
        End If
    Else
        xmlReponseElem.setAttribute "TYPE", "SYSERR"
    End If
        
    Set CreateErrorResponseNode = xmlReponseElem.cloneNode(True)
    
    If blnLogError Then
        If errIsSystemError() = True Then
            App.LogEvent "SYSERR: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError
        End If
    End If
    
    Set xmlDoc = Nothing
    Set xmlReponseElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlDescriptionElem = Nothing
    Set xmlElement = Nothing

End Function

Public Function errCreateErrorResponse() As String
' header ----------------------------------------------------------------------------------
' description:  Creates an xml error response based on the raised error
' pass:         n/a
' return:       The xml response string containing an the error information in the format
'               <RESPONSE TYPE=APPERR or SYSERR>
'                   <ERROR>
'                       <NUMBER> </NUMBER>
'                       <SOURCE> </SOURCE>
'                       <DESCRIPTION> </DESCRIPTION>
'                   </ERROR>
'               </RESPONSE>
'------------------------------------------------------------------------------------------
    Dim xmlErrorElem As IXMLDOMElement
    
    Set xmlErrorElem = CreateErrorResponseNode(False)
    
    If errIsSystemError() = True Then
        App.LogEvent "SYSERR: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError
    End If
        
    errCreateErrorResponse = xmlErrorElem.xml
    
    Set xmlErrorElem = Nothing

End Function

Public Function errGetMessageText(ByVal vlngMessageNo As Long, _
                                 Optional ByVal vintMessageField As Integer = omMESSAGE_TEXT) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the soft coded message information, the message can be an error or warning
' pass:         Message number to use to find message text
'               Optional Message field to return either Text or Type see omMESSAGE enum
' return:       Message text string in the following format:
'------------------------------------------------------------------------------------------
' MAR1300 - CORE135 Added support for getting error via CRUD
#If omCRUD Then
    errGetMessageText = adoCRUDGetMessageText(vlngMessageNo, vintMessageField)
#Else
    Const strFunctionName As String = "errGetMessageText"
    
    Dim strMessageText As String
    
    ' First see if the error is not on the database
    If Not errGetDefaultMessageText(vlngMessageNo, strMessageText) Then
        
    ' BMIDS958 If we are looking for a message text that has already been found, do not query the database.
        If (vintMessageField = omMESSAGE_TEXT) And _
           (vlngMessageNo = errGetOmigaErrorNumber(Err.Number)) And _
           (Err.Description <> "") Then
           
            strMessageText = Err.Description
        Else
        
           ' Get the message from the database
    
            Dim xmlDoc As FreeThreadedDOMDocument40
            Dim xmlElem As IXMLDOMElement
            Dim xmlRootNode As IXMLDOMNode
            Dim xmlSchemaNode As IXMLDOMNode
            Dim xmlRequestNode As IXMLDOMNode
            Dim xmlResponseNode As IXMLDOMNode
        
            ' create one DOMDocument to contain schema, request & response
            Set xmlDoc = New FreeThreadedDOMDocument40
            xmlDoc.validateOnParse = False
            xmlDoc.setProperty "NewParser", True
        
            ' create BOGUS root node
            Set xmlElem = xmlDoc.createElement("BOGUS")
            Set xmlRootNode = xmlDoc.appendChild(xmlElem)
        
            ' create schema for MESSAGE table
            Set xmlElem = xmlDoc.createElement("SCHEMA")
            Set xmlSchemaNode = xmlRootNode.appendChild(xmlElem)
            Set xmlElem = xmlDoc.createElement("MESSAGE")
            xmlElem.setAttribute "ENTITYTYPE", "PHYSICAL"
            Set xmlSchemaNode = xmlSchemaNode.appendChild(xmlElem)
            Set xmlElem = xmlDoc.createElement("MESSAGENUMBER")
            xmlElem.setAttribute "DATATYPE", "SHORT"
            xmlSchemaNode.appendChild xmlElem
            Set xmlElem = xmlDoc.createElement("MESSAGETEXT")
            xmlElem.setAttribute "DATATYPE", "STRING"
            xmlElem.setAttribute "LENGTH", "255"
            xmlSchemaNode.appendChild xmlElem
            Set xmlElem = xmlDoc.createElement("MESSAGETYPE")
            xmlElem.setAttribute "DATATYPE", "STRING"
            xmlElem.setAttribute "LENGTH", "20"
            xmlSchemaNode.appendChild xmlElem
            
        
            ' create REQUEST details
            Set xmlElem = xmlDoc.createElement("REQUEST")
            Set xmlRequestNode = xmlRootNode.appendChild(xmlElem)
            Set xmlElem = xmlDoc.createElement("MESSAGE")
            xmlElem.setAttribute "MESSAGENUMBER", CStr(vlngMessageNo)
            Set xmlRequestNode = xmlRequestNode.appendChild(xmlElem)
        
            ' create RESPONSE node
            Set xmlElem = xmlDoc.createElement("RESPONSE")
            Set xmlResponseNode = xmlRootNode.appendChild(xmlElem)
        
            adoGetRecordSetAsXML xmlRequestNode, xmlSchemaNode, xmlResponseNode
        
            If vintMessageField = omMESSAGE_TEXT Then
                If Not xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETEXT") Is Nothing Then
                    strMessageText = xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETEXT").Text
    '            Else
    '                strMessageText = "Error Message not found"
                End If
            Else
                    If Not xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETYPE") Is Nothing Then
                    strMessageText = xmlResponseNode.selectSingleNode("MESSAGE/@MESSAGETYPE").Text
                End If
            End If
        End If
        
        Set xmlDoc = Nothing
        Set xmlElem = Nothing
        Set xmlRootNode = Nothing
        Set xmlSchemaNode = Nothing
        Set xmlRequestNode = Nothing
        Set xmlResponseNode = Nothing
        
    End If
    
    errGetMessageText = strMessageText
#End If
End Function

Private Function errGetDefaultMessageText(ByVal lngMessageNo As Long, ByRef strMsgText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Gets all messages not specifed in the message table
' pass:         message number and empty message text string
' returns:      Boolean indicating whether the message number was found
'               message text if message is found
'------------------------------------------------------------------------------------------
          
    errGetDefaultMessageText = True
    
    Select Case lngMessageNo
        
        Case 556
            strMsgText = "Unable to establish database connection"
        
        Case 901
            strMsgText = "Error message not found"
            
        '------------------------------------------------------------------------------------------
        ' catch all
        '------------------------------------------------------------------------------------------
        Case Else
            errGetDefaultMessageText = False
    
    End Select

End Function

Public Function errIsApplicationError() As Boolean
' header ----------------------------------------------------------------------------------
' description:  Check if error is an omiga 4 application error.
'               NOTE: Warnings are application errors!
' pass:         n/a
' return:       boolean indicating if the error is an applicaton error
'------------------------------------------------------------------------------------------
    Dim lngOmigaErrorNo As Long, _
        lngErrNo As Long
        
    Dim blnIsApplicationError As Boolean
    
    blnIsApplicationError = False
    
    lngErrNo = Err.Number
        
'    If lngErrNo <> 0 Then
'        lngOmigaErrorNo = errGetOmigaErrorNumber(lngErrNo)
'
'        ' AS 23/11/99 For the error to be an Omiga4 Application error it must also
'        ' not be in the ADO error number range see ADODB.ErrorValueEnum
'        If (lngOmigaErrorNo > 0 And _
'            lngOmigaErrorNo <= clngMAX_ERROR_NO And _
'            (lngOmigaErrorNo < clngADO_START_ERROR_NO Or _
'            lngOmigaErrorNo > clngADO_END_ERROR_NO)) Then
'
'            blnIsApplicationError = True
'
'        End If
'    End If
    
    'DB BM0483 - Re-worked to subtract 512 from the error number.
    If lngErrNo <> 0 Then
        lngOmigaErrorNo = errGetOmigaErrorNumber(lngErrNo)

        If (lngOmigaErrorNo > 0 And _
            lngOmigaErrorNo <= clngMAX_ERROR_NO) Then
            lngErrNo = lngErrNo - vbObjectError
            If (lngErrNo < clngADO_START_ERROR_NO Or _
                lngOmigaErrorNo > clngADO_END_ERROR_NO) Then
     
                blnIsApplicationError = True
     
            End If
        End If
    End If
    'DB End
    
    errIsApplicationError = blnIsApplicationError
    
End Function


Public Function errIsSystemError() As Boolean
' header ----------------------------------------------------------------------------------
' description:  Determines if the error is a system error
' pass:         n/a
' return:       boolean indicating a system error
'------------------------------------------------------------------------------------------
    errIsSystemError = Not errIsApplicationError()
    
End Function

Public Function errFormatMessageNode() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    
    Dim xmlMessageDoc As FreeThreadedDOMDocument40
    Set xmlMessageDoc = New FreeThreadedDOMDocument40
    xmlMessageDoc.validateOnParse = False
    xmlMessageDoc.setProperty "NewParser", True
    Dim xmlMessageElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim lngOmigaErrorNo As Long
        
    Set xmlMessageElement = xmlMessageDoc.createElement("MESSAGE")
    xmlMessageDoc.appendChild xmlMessageElement
    Set xmlElement = xmlMessageDoc.createElement("TEXT")
    xmlElement.Text = Err.Description
    xmlMessageElement.appendChild xmlElement
    
    lngOmigaErrorNo = errGetOmigaErrorNumber(Err.Number)
    Set xmlElement = xmlMessageDoc.createElement("TYPE")
    xmlElement.Text = "Warning"
    xmlMessageElement.appendChild xmlElement
    
    errFormatMessageNode = xmlMessageDoc.xml
    
    Set xmlMessageDoc = Nothing
    Set xmlMessageElement = Nothing
    Set xmlElement = Nothing
        
End Function

' PSC 27/10/2005 MAR300
Public Sub errAddWarning(ByVal xmlResponse As IXMLDOMElement, ByVal lngErrorNo As Long, ByVal strSource As String, ByVal strDescription As String)
' header ----------------------------------------------------------------------------------
' description:  Adds the warning into the xmlResponse
' pass:         xmlResponse to add the warning to
' return:       n/a
'------------------------------------------------------------------------------------------

    Dim xmlMessageElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlFirstChild As IXMLDOMNode
    Dim lngOmigaErrorNo As Long
    
    If xmlResponse.nodeName <> "RESPONSE" Then
        ' PSC 27/10/2005 MAR300
        errThrowError strSource, oeMissingPrimaryTag, "RESPONSE must be the top level tag"
    End If
    
    xmlResponse.setAttribute "TYPE", "WARNING"
    
    Set xmlFirstChild = xmlResponse.firstChild
    
    Set xmlMessageElement = xmlResponse.ownerDocument.createElement("MESSAGE")
    
    xmlResponse.insertBefore xmlMessageElement, xmlFirstChild
        
    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETEXT")
    ' PSC 27/10/2005 MAR300
    xmlElement.Text = strDescription
    xmlMessageElement.appendChild xmlElement
    
    ' PSC 27/10/2005 MAR300
    lngOmigaErrorNo = errGetOmigaErrorNumber(lngErrorNo)
    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETYPE")
    xmlElement.Text = errGetMessageText(lngOmigaErrorNo, omMESSAGE_TYPE)
    xmlMessageElement.appendChild xmlElement
    
    Set xmlMessageElement = Nothing
    Set xmlElement = Nothing
    Set xmlFirstChild = Nothing
    
End Sub

'MAR1300 GHun
Public Sub errAddOmigaWarning(ByVal xmlResponse As IXMLDOMElement, ByVal vintOmigaErrorNo As Integer)
    
    Dim xmlMessageElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlFirstChild As IXMLDOMNode
    Dim lngOmigaErrorNo As Long
    Dim lngErrNo As Long
    Dim strText As String
    
    If xmlResponse.nodeName <> "RESPONSE" Then
        errThrowError Err.Source, oeMissingPrimaryTag, "RESPONSE must be the top level tag"
    End If
    
    lngErrNo = vbObjectError + 512 + vintOmigaErrorNo
    strText = errGetMessageText(vintOmigaErrorNo)
    xmlResponse.setAttribute "TYPE", "WARNING"

    Set xmlFirstChild = xmlResponse.firstChild
    Set xmlMessageElement = xmlResponse.ownerDocument.createElement("MESSAGE")

    xmlResponse.insertBefore xmlMessageElement, xmlFirstChild
        
    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETEXT")
    xmlElement.Text = strText

    xmlMessageElement.appendChild xmlElement
    lngOmigaErrorNo = vintOmigaErrorNo

    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETYPE")
    xmlElement.Text = CStr(vintOmigaErrorNo)
    xmlMessageElement.appendChild xmlElement

    Set xmlMessageElement = Nothing
    Set xmlElement = Nothing
    Set xmlFirstChild = Nothing
    
End Sub
'MAR1300 End

Public Function errGetOmigaErrorNumber(ByVal vlngErrorNo As Long) As Long
' header ----------------------------------------------------------------------------------
' description:  Converts an  error number to an omiga number. When errors are raised by VB
'               they get added to them the VB constant 'vbObjectError + 512' so to get the
'               Omiga4 error number orginally raised we need to subtract them from the err
'               number.
'               NOTE: System error numbers will end up much larger than Omiga4 error numbers
' pass:         Error number to calculate
' return:       Calculated Omiga4 error number
'------------------------------------------------------------------------------------------
    errGetOmigaErrorNumber = vlngErrorNo - vbObjectError - 512
End Function

Public Function errCheckXMLResponse( _
    ByVal vstrXmlReponse As String, _
    Optional ByVal vblnRaiseError As Boolean = False, _
    Optional ByVal vxmlInResponseElement As IXMLDOMElement = Nothing) _
    As Long

' header ----------------------------------------------------------------------------------
' description:  Checks the xml response, this method expects the xml response as a string
'               and designed to called when calling across component methods and then intergrating
'               the returned xml
' pass:         xmlResponse containing xml data to be checked
'               boolean as to whether to re-raise the error
'               xmlResponseElement to copy the warnings to
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "errCheckXMLResponse"

    Dim xmlXmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElement As IXMLDOMElement

    Set xmlXmlIn = xmlLoad(vstrXmlReponse, strFunctionName)
    Set xmlResponseElement = xmlXmlIn.selectSingleNode("RESPONSE")
    
    If xmlResponseElement Is Nothing Then
        errThrowError strFunctionName, oeXMLMissingElement, "Missing RESPONSE element"
    End If
    
    errCheckXMLResponse = errCheckXMLResponseNode(xmlResponseElement, vxmlInResponseElement, vblnRaiseError)

    Set xmlXmlIn = Nothing
    Set xmlResponseElement = Nothing
    
End Function

Public Sub errRaiseXMLResponse(ByVal vstrXmlResponse As String)
' header ----------------------------------------------------------------------------------
' description:  Checks the XML response for error information and raises the error
' pass:         xmlResponse containing xml data to be checked for an error
' return:       n/a
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "errRaiseXMLResponse"
    
    Dim xmlXmlIn As FreeThreadedDOMDocument40
    
    Set xmlXmlIn = xmlLoad(vstrXmlResponse, strFunctionName)
    
    errRaiseXMLResponseNode xmlXmlIn.documentElement
    
End Sub

Public Sub errRaiseXMLResponseNode(ByVal vxmlResponseNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Checks the XML response node for error information and raises the error
' pass:         vxmlResponseNode containing xml data to be checked for an error
' return:       n/a
'------------------------------------------------------------------------------------------
        
    Dim xmlResponseErrorNumber As IXMLDOMNode, _
        xmlResponseErrorSource As IXMLDOMNode, _
        xmlResponseErrorDesc As IXMLDOMNode
            
    Dim lngErrNo As Long
    
    Set xmlResponseErrorNumber = vxmlResponseNode.selectSingleNode("ERROR/NUMBER")
    If Not xmlResponseErrorNumber Is Nothing Then
        If IsNumeric(xmlResponseErrorNumber.Text) = True Then
            lngErrNo = CSafeLng(xmlResponseErrorNumber.Text)
        End If
    End If
    Set xmlResponseErrorNumber = Nothing
            
    If (lngErrNo <> 0) Then
                
        Dim strErrSource As String
        Set xmlResponseErrorSource = vxmlResponseNode.selectSingleNode("ERROR/SOURCE")
        If Not xmlResponseErrorSource Is Nothing Then
            If Len(xmlResponseErrorSource.Text) > 0 Then
                strErrSource = xmlResponseErrorSource.Text
            End If
        End If
        Set xmlResponseErrorSource = Nothing
        
        Dim strErrDesc As String
        Set xmlResponseErrorDesc = vxmlResponseNode.selectSingleNode("ERROR/DESCRIPTION")
        If Not xmlResponseErrorDesc Is Nothing Then
            If Len(xmlResponseErrorDesc.Text) > 0 Then
                strErrDesc = xmlResponseErrorDesc.Text
            End If
        End If
        Set xmlResponseErrorDesc = Nothing
        
        If Len(strErrDesc) = 0 Then
            strErrDesc = errGetMessageText(oeUnspecifiedError)
        End If
                    
        Err.Raise lngErrNo, strErrSource, strErrDesc
        
    End If
    
End Sub



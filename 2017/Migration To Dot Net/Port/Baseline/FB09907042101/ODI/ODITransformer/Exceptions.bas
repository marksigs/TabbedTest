Attribute VB_Name = "Exceptions"
'Workfile:      Exceptions.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     27/08/01    Expanded code created by LD.
'AS     12/10/01    Fixed handling of Optimus BUSINESSEXCEPTION.
'DS     24/04/02    Fixed handling of Optimus business exception even more (SYS4350)
'------------------------------------------------------------------------------------------
Option Explicit

' fixme - check error type values
Private Const cstrTYPE_ERROR = "0"

Public Sub AddExceptionsToResponse( _
    ByVal vnodeConverterResponse As IXMLDOMNode, _
    ByVal vnodeResponse As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Adds any business exceptions in vnodeConverterResponse to vnodeResponse
'   in the following format:
'   <BUSINESSEXCEPTION
'       MESSAGE=""
'       SOURCEOBJECTKEY=""
'       SOURCEPROPERTY=""
'       TYPE=""/>
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddExceptionsToResponseExit

    Const strFunctionName = "AddExceptionsToResponse"
    
    Dim nodelistBusinessExceptions As IXMLDOMNodeList
    Dim nodeBusinessException As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    
    Set nodelistBusinessExceptions = vnodeConverterResponse.selectNodes("EXCEPTIONS/BUSINESSEXCEPTION")
    
    For Each nodeBusinessException In nodelistBusinessExceptions
        Set nodeTemp = vnodeResponse.ownerDocument.createElement("BUSINESSEXCEPTION")
        xmlSetAttributeValue nodeTemp, "MESSAGE", _
            xmlGetNodeText(nodeTemp, ".//MESSAGE/@DATA")
        xmlSetAttributeValue nodeTemp, "SOURCEOBJECTKEY", _
            xmlGetNodeText(nodeTemp, ".//SOURCEOBJECTKEY/@DATA")
        xmlSetAttributeValue nodeTemp, "SOURCEPROPERTY", _
            xmlGetNodeText(nodeTemp, ".//SOURCEPROPERTY/@DATA")
        xmlSetAttributeValue nodeTemp, "TYPE", _
            xmlGetNodeText(nodeTemp, ".//TYPE/@DATA")
        vnodeResponse.appendChild nodeTemp
    Next
    
AddExceptionsToResponseExit:

    Set nodelistBusinessExceptions = Nothing
    Set nodeBusinessException = Nothing
    Set nodeTemp = Nothing
    
    errCheckError strFunctionName
    
End Sub

Public Function CheckConverterResponse( _
    ByVal vnodeConverterResponse As IXMLDOMNode, _
    ByVal blnRaiseOmigaError) As Integer
' header ----------------------------------------------------------------------------------
' description:
'   Checks nodeConverterResponse for any business exceptions of type ERROR. If none are
'   found it returns zero. If any are found it returns the type of the first it finds. If
'   blnRaiseOmigaError is true it raises an Omiga error.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CheckConverterResponseExit

    Const strFunctionName = "CheckConverterResponse"
    
    Dim nodelistBusinessExceptions As IXMLDOMNodeList
    Dim nodeBusinessException As IXMLDOMNode
    Dim strType As String
    Dim intErrorNum As Integer
    
    intErrorNum = 0
    
    Set nodelistBusinessExceptions = vnodeConverterResponse.selectNodes("EXCEPTIONS/BUSINESSEXCEPTION/BUSINESSEXCEPTION")
    
    For Each nodeBusinessException In nodelistBusinessExceptions
    
        strType = xmlGetMandatoryNodeText(nodeBusinessException, "TYPE/@DATA")
        If strType = cstrTYPE_ERROR Then
            If blnRaiseOmigaError = True Then
                ' "Administration System Error"
                errThrowError strFunctionName, 4500, xmlGetNodeText(nodeBusinessException, "MESSAGE/@DATA")
            End If
            ' fixme - get the correct warning number from the message
            intErrorNum = 1
        End If
    
    Next
    
    CheckConverterResponse = intErrorNum
    
CheckConverterResponseExit:

    Set nodelistBusinessExceptions = Nothing
    Set nodeBusinessException = Nothing
    
    errCheckError strFunctionName
    
End Function


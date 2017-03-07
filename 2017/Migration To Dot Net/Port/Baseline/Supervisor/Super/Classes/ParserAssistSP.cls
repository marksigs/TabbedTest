VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLAssistSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Workfile:      ParserAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Helper object for XML parser.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'SA     SYS4111 If there's a parser error in the Load method, write response string to event log.
'------------------------------------------------------------------------------------------

Option Explicit


Public Enum BOOPERATIONTYPE
    booNone = 0
    booCreate = 1
    booDelete = 2
    booUpdate = 3
End Enum
Public Function Load( _
    ByVal vstrXMLData As String, _
    ByVal vstrCallingObject As String, _
    ByVal vstrCallingFunction As String _
    ) As FreeThreadedDOMDocument
' header ----------------------------------------------------------------------------------
' description:
'   Load an xml document from a string. Raise any resultant parser error.
' pass:
'   vstrXMLData         xml data stream containing data to be loaded
'   vstrCallingObject    name of calling object
'   vstrCallingFunction  name of calling function
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "Load"
    
    Dim objXmlDoc As New FreeThreadedDOMDocument
    Dim strErrMsg As String
    
    objXmlDoc.async = False
    
    objXmlDoc.loadXML vstrXMLData
    
    If objXmlDoc.parseError.errorCode <> 0 Then
        'SYS4111 If there's a parser error, write the contents of vstrXMLData
        'to the event log as it will contain the html detailing the specific error.
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & " called from " & vstrCallingFunction
        strErrMsg = strErrMsg & vbCrLf & vstrXMLData
                
        App.LogEvent strErrMsg, vbLogEventTypeError
    
        g_clsErrorHandling.RaiseError errGeneralError, "Parser Error - See event log for further details."
    End If
    
    Set Load = objXmlDoc
    
End Function
Public Function GetAttributeFromNode(ByVal vobjNode As IXMLDOMNode, _
                                     ByVal vstrAttribName As String) _
                                     As String
    
    Dim strValue As String
    
    If Not vobjNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
        strValue = vobjNode.Attributes.getNamedItem(vstrAttribName).Text
    End If
    
    GetAttributeFromNode = strValue

End Function
Public Function CheckResponse(ByVal vxmlResponseToCheck As IXMLDOMElement, _
                              ByVal vxmlResponseToAddTo As IXMLDOMElement, _
                              Optional ByVal vblnRaiseError As Boolean = False) As Long
' header ----------------------------------------------------------------------------------
' description:
'   takes an xml response and returns the error number. Re raises the error if
'   vblnRaiseError is set to true
' pass:
'   vxmlResponseToCheck   xml Response to check for errors
'   vxmlResponseToAddTo   xml Element to add the warnings to
'   vblnRaiseError        Determines whether to raise the error
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "CheckResponse"
    
    Dim objXmlAssist As New XMLAssistSP
    Dim xmlErrNumElement As IXMLDOMElement
    Dim xmlErrSourceElement As IXMLDOMElement
    Dim xmlErrDescElement As IXMLDOMElement
    Dim xmlFirstChildElem As IXMLDOMNode
    Dim lngErrorNumber As Long
    
    Dim xmlMessageList As IXMLDOMNodeList
    Dim xmlMessageElem As IXMLDOMElement
    
    If vxmlResponseToCheck Is Nothing Or vxmlResponseToAddTo Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Response to Check or Response to add to missing"
    End If
    
    If vxmlResponseToCheck.nodeName <> "RESPONSE" Then
        g_clsErrorHandling.RaiseError errGeneralError, "RESPONSE must be top level tag"
    End If
    
    ' If the error is to be re-raised get the appropriate data out of the response
    ' and raise the error
    If vxmlResponseToCheck.getAttribute("TYPE") = "WARNING" Then
        vxmlResponseToAddTo.setAttribute "TYPE", "WARNING"
        
        Set xmlFirstChildElem = vxmlResponseToCheck.firstChild
        Set xmlMessageList = vxmlResponseToCheck.getElementsByTagName("MESSAGE")

        ' insert messages at the top of the response
        For Each xmlMessageElem In xmlMessageList
            vxmlResponseToAddTo.appendChild xmlMessageElem.cloneNode(True)
        Next
    ElseIf vxmlResponseToCheck.getAttribute("TYPE") <> "SUCCESS" Then
        Set xmlErrNumElement = vxmlResponseToCheck.getElementsByTagName("NUMBER").Item(0)
        lngErrorNumber = CLng(xmlErrNumElement.Text)
        
        If vblnRaiseError = True Then
            Set xmlErrSourceElement = vxmlResponseToCheck.getElementsByTagName("SOURCE").Item(0)
            Set xmlErrDescElement = vxmlResponseToCheck.getElementsByTagName("DESCRIPTION").Item(0)
            Err.Raise lngErrorNumber, xmlErrSourceElement.Text, xmlErrDescElement.Text
        End If
    End If

    CheckResponse = lngErrorNumber
    
    Set objXmlAssist = Nothing
    Set xmlErrNumElement = Nothing
    Set xmlErrSourceElement = Nothing
    Set xmlErrDescElement = Nothing
    Set xmlFirstChildElem = Nothing
    
End Function

Public Function CheckXMLResponse( _
    ByVal vstrXmlReponse As String, Optional ByVal vblnRaiseError As Boolean = False, _
    Optional ByVal vxmlInResponseElement As IXMLDOMElement = Nothing) As Long
' header ----------------------------------------------------------------------------------
' description:
'   takes an xml response and returns the error number. Re raises the error if
'   vblnRaiseError is set to true
' pass:
'   vstrXmlReponse        xml Response data stream containing data to be checked
'   vblnRaiseError        Determines whether to raise the error
'   vxmlInResponseElement xml Element to add the warnings to
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "CheckXMLResponse"
    
    Dim objXmlIn As FreeThreadedDOMDocument
    Dim objXmlAssist As New XMLAssistSP
    Dim objResponseElement As IXMLDOMElement
    Dim objErrNumElement As IXMLDOMElement
    Dim objErrSourceElement As IXMLDOMElement
    Dim objErrDescElement As IXMLDOMElement
    Dim objFirstChildElem As IXMLDOMNode
    Dim lngErrorNumber As Long
    
    Dim xmlMessageList As IXMLDOMNodeList
    Dim xmlMessageElem As IXMLDOMElement

    Set objXmlIn = objXmlAssist.Load(vstrXmlReponse, TypeName(Me), strFunctionName)

    Set objResponseElement = objXmlIn.getElementsByTagName("RESPONSE").Item(0)
    
    If objResponseElement Is Nothing Then
        g_clsErrorHandling.RaiseError errMissingPrimaryTag
    End If
    
    ' If the error is to be re-raised get the appropriate data out of the response
    ' and raise the error
    If objResponseElement.getAttribute("TYPE") = "WARNING" Then
        If Not vxmlInResponseElement Is Nothing Then
            vxmlInResponseElement.setAttribute "TYPE", "WARNING"
            
            Set objFirstChildElem = vxmlInResponseElement.firstChild
            Set xmlMessageList = objXmlIn.getElementsByTagName("MESSAGE")
    
            ' insert messages at the top of the response
            For Each xmlMessageElem In xmlMessageList
                vxmlInResponseElement.insertBefore xmlMessageElem.cloneNode(True), _
                                                   objFirstChildElem
            Next
        End If
    ElseIf objResponseElement.getAttribute("TYPE") <> "SUCCESS" Then
        Set objErrNumElement = objResponseElement.getElementsByTagName("NUMBER").Item(0)
        lngErrorNumber = CLng(objErrNumElement.Text)
        
        If vblnRaiseError = True Then
            Set objErrSourceElement = objResponseElement.getElementsByTagName("SOURCE").Item(0)
            Set objErrDescElement = objResponseElement.getElementsByTagName("DESCRIPTION").Item(0)
            Err.Raise lngErrorNumber, objErrSourceElement.Text, objErrDescElement.Text
        End If
    End If

    CheckXMLResponse = lngErrorNumber
    
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objResponseElement = Nothing
    Set objErrNumElement = Nothing
    Set objErrSourceElement = Nothing
    Set objErrDescElement = Nothing
    Set objFirstChildElem = Nothing

    
End Function




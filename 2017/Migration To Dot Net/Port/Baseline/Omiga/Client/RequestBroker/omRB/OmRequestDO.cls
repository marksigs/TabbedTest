VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmRequestDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------
'EPSOM Specific History:
'Prog   Date        AQR         Description
'LH     10/01/2007  EP2_647     added task man. rules (TMRBTemplateEx) interface
'-----------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl

Public Function OmDataRequest(ByVal vstrXMLRequest As String) As String
    On Error GoTo OmDataRequestVbErr
    Const cstrFunctionName As String = "OmDataRequest"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim strTemplate As String
    
    GetObjectContext.SetComplete
    
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.async = False
    xmlIn.setProperty "NewParser", True
    xmlIn.validateOnParse = False
    
    xmlIn.loadXML vstrXMLRequest
    If xmlIn.parseError.errorCode <> 0 Then
        xmlParserError xmlIn.parseError
    End If
    
    If xmlIn.documentElement.nodeName <> "REQUEST" Then
        errRaise "invalid REQUEST: " & vstrXMLRequest
    End If
    
    If xmlIn.documentElement.childNodes.length = 0 Then
        errRaise "invalid REQUEST: " & vstrXMLRequest
    End If
    
    Select Case xmlIn.documentElement.getAttribute("RB_TEMPLATE")
    
        Case "RARulesTemplate"
            
            OmDataRequest = omCrudInterface(xmlIn.documentElement, "raRules")
    
        Case "CDRBTemplate"
            
            OmDataRequest = omCrudInterface(xmlIn.documentElement, "cdRules")
    
        Case "CCRBTemplate"
            
            OmDataRequest = omCrudInterface(xmlIn.documentElement, "ccRules")
    
        'IK_11/05/2006_EP357
        Case "AdminRulesTemplate"
            
            OmDataRequest = omCrudInterface(xmlIn.documentElement, "completions")
    
        'LH_10/01/2007_EP2_647
        Case "TMRBTemplateEx", "TMRBTemplate"
            
            OmDataRequest = omCrudInterface(xmlIn.documentElement, "tmRules")
            
        Case Else
            
            OmDataRequest = omRbInterface(vstrXMLRequest)
        
    End Select

OmDataRequestExit:
    
    Set xmlIn = Nothing
    Exit Function

OmDataRequestVbErr:
    
    OmDataRequest = createErrorResponse
    Resume OmDataRequestExit

End Function

Private Function omCrudInterface( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vstrSchemaName As String) As String

    Const cstrMethodName As String = "omCrudInterface"
    On Error GoTo omCrudInterfaceExit
    
    Dim omCRUD As Object
    
    vxmlRequest.removeAttribute "RB_TEMPLATE"
    vxmlRequest.setAttribute "CRUD_OP", "READ"
    vxmlRequest.setAttribute "SCHEMA_NAME", vstrSchemaName
    vxmlRequest.setAttribute "COMBOLOOKUP", "Y"
    vxmlRequest.setAttribute "COMBOTYPELOOKUP", "EX"
    vxmlRequest.setAttribute "ENTITY_REF", vxmlRequest.firstChild.nodeName
    vxmlRequest.firstChild.Attributes.removeNamedItem "_SCHEMA_"
    
    Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    omCrudInterface = omCRUD.omRequest(vxmlRequest.xml)
    
omCrudInterfaceExit:

    Set omCRUD = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If

End Function

Private Function omRbInterface(ByVal vstrXMLRequest As String) As String

    Const cstrMethodName As String = "omRbInterface"
    On Error GoTo omRbInterfaceExit

    Dim omRBlg As Object
    Set omRBlg = GetObjectContext.CreateInstance("omRBlg.omRequestDO")
    omRbInterface = omRBlg.OmDataRequest(vstrXMLRequest)
    
omRbInterfaceExit:

    Set omRBlg = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If

End Function

Private Sub errRaise(ByVal vstrDesc As String)
    Err.Raise oeXMLParserError, "omRB.OmDataRequest", vstrDesc
End Sub

Private Sub xmlParserError(ByVal vobjParseError As IXMLDOMParseError)
    
    Dim strErrDesc As String    ' formatted parser error
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & vobjParseError.reason & vbCr & _
        "Error code: " & Str(vobjParseError.errorCode) & vbCr & _
        "Line no.: " & Str(vobjParseError.Line) & vbCr & _
        "Character: " & Str(vobjParseError.linepos) & vbCr & _
        "Source text: " & vobjParseError.srcText & vbCr & _
        "url: " & vobjParseError.url
        
    errRaise strErrDesc

End Sub

Private Function createErrorResponse() As String
    
    createErrorResponse = _
        "<RESPONSE TYPE='ERROR'>" & _
        "<ERROR>" & _
        "<NUMBER>" & Err.Number & "</NUMBER>" & _
        "<SOURCE>" & Err.Source & "</SOURCE>" & _
        "<DESCRIPTION>" & Err.Description & "</DESCRIPTION>" & _
        "</ERROR>" & _
        "</RESPONSE>"

End Function

Private Sub ObjectControl_Activate()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
End Sub

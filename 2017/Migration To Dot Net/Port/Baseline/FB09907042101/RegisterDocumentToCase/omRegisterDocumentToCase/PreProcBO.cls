VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "PreProcBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     24/10/2005  created for Project MARS (MAR232)
'BC     23/02/2006  MAR1294 - Create DOCUMENTGROUP element from VALUEID rather than VALUENAME

Option Explicit


Public Function OmRequest(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "OmRequest"
    On Error GoTo OmRequestVbErr
    
    Dim xmlInDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    
    GetObjectContext.SetComplete
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    Set xmlRequestNode = xmlInDoc.selectSingleNode("REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    ValidateRequest xmlRequestNode
    
    GetDocumentGroup xmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS")
    
    OmRequest = xmlInDoc.xml
    
OmRequestVbErr:

    Set xmlInDoc = Nothing
    
    If Err.Number <> 0 Then
        If Err.Source <> cstrMethodName Then
            If Err.Source = App.EXEName Then
                Err.Source = cstrMethodName
            Else
                Err.Source = cstrMethodName & "." & Err.Source
            End If
        End If
        Err.Source = App.EXEName & "." & TypeName(Me) & "." & Err.Source
        OmRequest = FormatError(vstrRequest)
    End If

End Function

Private Sub ValidateRequest(ByVal vxmlRequestNode As IXMLDOMNode)

    Const cstrMethodName As String = "ValidateRequest"
    On Error GoTo ValidateRequestExit
    
    If vxmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS[@WORKGROUPID]") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no valid  DOCUMENTAUDITDETAILS node on pre-processing REQUEST node"
    End If
    
ValidateRequestExit:
    
    CheckError cstrMethodName

End Sub

Private Sub GetDocumentGroup(ByVal vxmlDocumentAuditDetailsNode As IXMLDOMElement)

    Const cstrMethodName As String = "GetDocumentGroup"
    On Error GoTo GetDocumentGroupExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Dim strComponentResponse As String, _
        strWorkGroupId As String
        
    strWorkGroupId = vxmlDocumentAuditDetailsNode.getAttribute("WORKGROUPID")

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "COMBOGROUP"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("COMBOGROUP")
    xmlElem.setAttribute "GROUPNAME", "PrintDocumentType"
    xmlNode.appendChild xmlElem

    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    strComponentResponse = crudObj.OmRequest(xmlDoc.xml)
    Set crudObj = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error reading COMBOGROUP PrintDocumentType"
    End If
    
    Set xmlElem = xmlDoc.selectSingleNode("RESPONSE/COMBOGROUP[@GROUPNAME='PrintDocumentType']/COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" & strWorkGroupId & "']]")
    
    'BC MAR1294 23/02/2006 Begin
'    If Not xmlElem Is Nothing Then
'        If Not xmlElem.Attributes.getNamedItem("VALUENAME") Is Nothing Then
'            vxmlDocumentAuditDetailsNode.setAttribute "DOCUMENTGROUP", xmlElem.getAttribute("VALUENAME")
'        End If
'    End If

    If Not xmlElem Is Nothing Then
        If Not xmlElem.Attributes.getNamedItem("VALUEID") Is Nothing Then
            vxmlDocumentAuditDetailsNode.setAttribute "SEARCHKEY1", xmlElem.getAttribute("VALUEID")
            vxmlDocumentAuditDetailsNode.setAttribute "DOCUMENTGROUP", xmlElem.getAttribute("VALUEID")
        End If
    End If
    'BC MAR1294 23/02/2006 End
    
GetDocumentGroupExit:

    Set crudObj = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

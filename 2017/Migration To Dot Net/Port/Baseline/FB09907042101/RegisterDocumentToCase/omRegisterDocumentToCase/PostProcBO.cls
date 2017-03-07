VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "PostProcBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     24/10/2005  created for Project MARS (MAR232)

Option Explicit

Public Function OmRequest( _
    ByVal vstrCrudRequest As String, _
    ByVal vstrCrudResponse As String _
    ) As String

    Const cstrMethodName As String = "OmRequest"
    On Error GoTo OmRequestVbErr
    
    Dim xmlInDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    
    Dim strResponse As String
    
    ' default response - will be overwritten as required
    OmRequest = "<RESPONSE TYPE='SUCCESS'/>"
    
    GetObjectContext.SetComplete
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrCrudRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    Set xmlRequestNode = xmlInDoc.selectSingleNode("REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    ValidateRequest xmlRequestNode
    
    CallHandleInterfaceResponse xmlRequestNode
    
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
        OmRequest = FormatError(vstrCrudRequest)
    End If

End Function

Private Sub ValidateRequest(ByVal vxmlRequestNode As IXMLDOMNode)

    Const cstrMethodName As String = "ValidateRequest"
    On Error GoTo ValidateRequestExit
    
    If vxmlRequestNode.Attributes.getNamedItem("USERID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERID on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("UNITID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no UNITID on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERAUTHORITYLEVEL on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS[@APPLICATIONNUMBER][@WORKGROUPID]") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no valid  DOCUMENTAUDITDETAILS node on post-processing REQUEST node"
    End If
    
ValidateRequestExit:
    
    CheckError cstrMethodName

End Sub

Private Sub CallHandleInterfaceResponse(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "CallHandleInterfaceResponse"
    On Error GoTo CallHandleInterfaceResponseExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlCaseActivityDoc As DOMDocument40
    
    Dim omTmObj As Object
    
    Dim strComponentResponse As String, _
        strWorkGroupId As String, _
        strMessageType As String, _
        strApplicationNumber As String, _
        strCaseStage As String
        
    Set xmlElem = vxmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS[@APPLICATIONNUMBER][@WORKGROUPID]")
    strApplicationNumber = xmlElem.getAttribute("APPLICATIONNUMBER")
    strWorkGroupId = xmlElem.getAttribute("WORKGROUPID")
    strMessageType = strWorkGroupId
    
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.setProperty "NewParser", True
    xmlCaseActivityDoc.async = False
    
    GetCaseActivity strApplicationNumber, xmlCaseActivityDoc
    Set xmlElem = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGE")
    strCaseStage = xmlElem.getAttribute("STAGEID")
    
    If strWorkGroupId = "5" Then
    
        Set xmlElem = vxmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS[@APPLICATIONNUMBER]")
    
        If IsAfterOffer(strCaseStage) Then
            strMessageType = "6"
        End If
        
    End If
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
    xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
    xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
    xmlElem.setAttribute "OPERATION", "HANDLEINTERFACERESPONSE"
    Set xmlThisRequestNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    Set xmlRequestElem = vxmlRequestNode.selectSingleNode("OPERATION/DOCUMENTAUDITDETAILS[@APPLICATIONNUMBER]")
    xmlElem.setAttribute "APPLICATIONNUMBER", xmlRequestElem.getAttribute("APPLICATIONNUMBER")
    If xmlRequestElem.Attributes.getNamedItem("APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlRequestElem.getAttribute("APPLICATIONFACTFINDNUMBER")
    End If
    xmlThisRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("INTERFACE")
    xmlElem.setAttribute "INTERFACETYPE", GetInterfaceType
    xmlElem.setAttribute "MESSAGETYPE", strMessageType
    xmlElem.setAttribute "CREATETASKFLAG", "1"
    xmlThisRequestNode.appendChild xmlElem
    
    xmlThisRequestNode.appendChild xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY").cloneNode(False)
    
    Set omTmObj = GetObjectContext.CreateInstance("omTm.omTmBO")
    strComponentResponse = omTmObj.OmTmRequest(xmlDoc.xml)
    xmlDoc.loadXML strComponentResponse
    Set omTmObj = Nothing
        
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        gstrComponentId = "omTm.omTmBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error calling HandleInterfaceResponse"
    End If
    
CallHandleInterfaceResponseExit:
    
    Set omTmObj = Nothing
    Set xmlRequestElem = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlCaseActivityDoc = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub GetCaseActivity(ByVal vstrApplicationNumber As String, ByVal vxmlDoc As DOMDocument40)

    Const cstrMethodName As String = "GetCaseActivity"
    On Error GoTo GetCaseActivityExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Dim strComponentResponse As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "CASEACTIVITY"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "CASEID", vstrApplicationNumber
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASESTAGE")
    xmlNode.appendChild xmlElem

    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    strComponentResponse = crudObj.OmRequest(xmlDoc.xml)
    Set crudObj = Nothing
    
    vxmlDoc.loadXML strComponentResponse
    If vxmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/CASEACTIVITY/CASESTAGE[@STAGEID]") Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error reading CASEACTIVITY"
    End If
    
GetCaseActivityExit:
    
    Set crudObj = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Function GetInterfaceType() As String

    Const cstrMethodName As String = "GetInterfaceType"
    On Error GoTo GetInterfaceTypeExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Dim strComponentResponse As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "COMBOGROUP"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("COMBOGROUP")
    xmlElem.setAttribute "GROUPNAME", "InterfaceType"
    xmlNode.appendChild xmlElem

    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    strComponentResponse = crudObj.OmRequest(xmlDoc.xml)
    Set crudObj = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlDoc.parseError
    End If
    
    Set xmlElem = xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/COMBOGROUP/COMBOVALUE[@VALUEID][COMBOVALIDATION[@VALIDATIONTYPE='CAT']]")
    If xmlElem Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving InterfaceType COMBOVALUE"
    End If
    
    GetInterfaceType = xmlElem.getAttribute("VALUEID")
    
GetInterfaceTypeExit:
    
    Set crudObj = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function IsAfterOffer(ByVal vstrThisStage As String) As Boolean

    Const cstrMethodName As String = "IsAfterOffer"
    On Error GoTo IsAfterOfferExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Dim strComponentResponse As String, _
        strTMOfferStage As String, _
        strThisStage As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "GLOBALPARAMETER"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("GLOBALPARAMETER")
    xmlElem.setAttribute "NAME", "TMOfferStage"
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
    
    Set xmlElem = xmlDoc.selectSingleNode("RESPONSE/GLOBALPARAMETER[@AMOUNT]")
    If Not xmlElem Is Nothing Then
    
        strTMOfferStage = xmlElem.getAttribute("AMOUNT")
        
        If IsNumeric(strTMOfferStage) And IsNumeric(vstrThisStage) Then
        
            If CInt(vstrThisStage) > CInt(strTMOfferStage) Then
                IsAfterOffer = True
            End If
            
        End If
    
    End If
    
IsAfterOfferExit:
    
    Set crudObj = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

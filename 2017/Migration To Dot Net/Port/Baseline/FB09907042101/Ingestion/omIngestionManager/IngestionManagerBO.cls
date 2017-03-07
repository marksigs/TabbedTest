VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IngestionManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Prog   Date        Description
'IK     18/10/2005  created for Project MARS
'IK     25/10/2005  MAR232 - re-structuring, RemoveRedundantNodes moved from PreProcBO
'IK     09/11/2005  MAR498 - re-structuring, call critical data check before omCRUD
'                   CRUD operation & decision task in transaction boundary
'IK     25/11/2005  MAR695 gstrComponentId & gstrComponentResponse now private member variables
'IK     29/11/2005  MAR715
'IK     21/12/2005  MAR919 - certain nodes cannot be deleted by RemoveRedundantNodes
'                   xpath: values needed elsewhere
'IK     07/03/2006  EP198, EP199 stored procedures in WebServiceSchema
'IK     08/03/2006  EP195 LinkPackagerToApp now incorporated in ingestion CRUD call
'IK     11/10/2006  EP2_10 - address targeting response
'IK     18/11/2006  EP2_134 no critical data check required for epsom
'IK     11/12/2006  EP2_398 get default values in ValidatePreProcRequest
'IK     11/12/2006  EP2_398 no APPLICATION no post-proc task
'IK     11/12/2006  EP2_398 if MORTGAGEPRODUCTCODE, STARTDATE required
'IK     04/01/2007  EP2_494 validate (& substitute default values) AddressTargetResolver request
'IK     08/01/2007  EP2_352 remote underwriter

Option Explicit
Private gstrComponentId As String
Private gstrComponentResponse As String

Public Function OmRequest(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "OmRequest"
    On Error GoTo OmRequestVbErr

    Dim objPreProc As Object
    Dim objCRUD As Object
    
    Dim xmlDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    
    Dim strAfterPreProc As String, _
        strCrudResponse As String
        
    ' will be updated as required
    GetObjectContext.SetComplete
    
    ' ik_20051221_mar919
    LoadConfigFile
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML vstrRequest
    
    If xmlDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlDoc.parseError
    End If
    
    Set xmlRequestNode = xmlDoc.selectSingleNode("REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    ValidateRequest xmlRequestNode
    
    ' EP2_10
    If Not IsAddressTargetResolver(xmlRequestNode) Then
        PreProcRequest xmlRequestNode
    End If
    
    If xmlDoc.selectSingleNode("REQUEST") Is Nothing Then
        
        ' error from pre-proc
        OmRequest = xmlDoc.xml
            
        'EP2_398
        ' create general failure task
        ' xmlRequestNode.setAttribute "UPDATEABORT", "true"
        ' xmlRequestNode.setAttribute "PostProcTaskId", "GeneralFailure"
        ' ProcessPostProc xmlRequestNode
    
    Else

        Select Case xmlRequestNode.getAttribute("OPERATION")
        
            Case "SubmitAiP", "SubmitFMA"
                
                OmRequest = ProcessSubmit(xmlDoc.documentElement, xmlRequestNode.getAttribute("OPERATION"))
                
            Case Else
            
                If xmlDoc.selectSingleNode("REQUEST[@CRUD_OP]") Is Nothing Then
                    ' no changes
                    OmRequest = "<RESPONSE TYPE='SUCCESS'/>"
                Else
                    OmRequest = ProcessStopAndSave(xmlDoc.documentElement)
                End If
                
        End Select
            
    End If
    
OmRequestVbErr:

    Set xmlDoc = Nothing
        
    If Err.Number <> 0 Then
        If Err.Source = "OMIGAERROR" Then
            OmRequest = Err.Description
        Else
            If Err.Source <> cstrMethodName Then
                If Err.Source = App.EXEName Then
                    Err.Source = cstrMethodName
                Else
                    Err.Source = cstrMethodName & "." & Err.Source
                End If
            End If
            Err.Source = App.EXEName & "." & TypeName(Me) & "." & Err.Source
            OmRequest = FormatError(vstrRequest, gstrComponentId, gstrComponentResponse)
        End If
    End If

End Function

Private Sub ValidateRequest(ByVal vxmlRequestNode As IXMLDOMNode)

    Const cstrMethodName As String = "ValidateRequest"
    On Error GoTo ValidateRequestExit
    
    'EP2_398
    Dim xmlNode As IXMLDOMNode
    'EP2_352
    Dim xmlAttrib As IXMLDOMAttribute
    
    If vxmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no OPERATION on ingestion REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERID on ingestion REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("UNITID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no UNITID on ingestion REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERAUTHORITYLEVEL on ingestion REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("APPLICATION") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION node on ingestion REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION/APPLICATIONFACTFIND node on ingestion REQUEST node"
    End If
    
    'EP2_398
    For Each xmlNode In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGEPRODUCTCODE]")
        If xmlNode.Attributes.getNamedItem("STARTDATE") Is Nothing Then
            Err.Raise oeMissingPrimaryTag, cstrMethodName, "no STARTDATE for MORTGAGEPRODUCTCODE on ingestion LOANCOMPONENT node"
        End If
    Next
    
    'EP2_494
    If vxmlRequestNode.Attributes.getNamedItem("USERID").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERID]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("USERID").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERID]").Attributes.getNamedItem("USERID").Text
        End If
    End If
    If vxmlRequestNode.Attributes.getNamedItem("UNITID").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@UNITID]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("UNITID").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@UNITID]").Attributes.getNamedItem("UNITID").Text
        End If
    End If
    If vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERAUTHORITYLEVEL]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERAUTHORITYLEVEL]").Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
        End If
    End If
    'EP2_494_ends
    
    'EP2_352
    If vxmlRequestNode.Attributes.getNamedItem("omigaClient") Is Nothing Then
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("omigaClient")
        xmlAttrib.Value = "mars"
        vxmlRequestNode.Attributes.setNamedItem xmlAttrib
    End If
    'EP2_352_ends
    
ValidateRequestExit:

    Set xmlNode = Nothing
    'EP2_352
    Set xmlAttrib = Nothing
    
    CheckError cstrMethodName

End Sub

Private Function ProcessSubmit( _
    ByVal vxmlPreProcResponse As IXMLDOMElement, _
    ByVal vstrOperation As String) _
    As String
    
    Const cstrMethodName As String = "ProcessSubmit"
    On Error GoTo ProcessSubmitExit
    
    Dim xmlResponseDoc As DOMDocument40
    Dim strCrudResponse As String
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
                
    If vstrOperation = "SubmitAiP" Then
        
        vxmlPreProcResponse.setAttribute "PostProcTaskId", "RunXMLCreditCheck"
    
    Else
        'IK_18/11/2006_EP2_134
        Select Case vxmlPreProcResponse.getAttribute("omigaClient")
            Case "epsom"
                ' nothing to do
            Case Else
        
                If Not vxmlPreProcResponse.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
            
                    xmlResponseDoc.loadXML GetApplicationData(vxmlPreProcResponse)
                    
                    If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/APPLICATION") Is Nothing Then
                        gstrComponentId = "omCRUD.omCRUDBO"
                        gstrComponentResponse = xmlResponseDoc.xml
                        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving comparison data "
                    End If
                    
                    ' creates attribute 'PostProcTaskId' containing omTm OPERATION value
                    ' on REQUEST node
                    CriticalDataCheck vxmlPreProcResponse, xmlResponseDoc.xml
                
                End If
        
        End Select
        ''IK_18/11/2006_EP2_134_ends
    
    End If
    
    If vxmlPreProcResponse.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
        ProcessSubmit = ProcessPostProc(vxmlPreProcResponse)
    Else
        strCrudResponse = CrudCall(vxmlPreProcResponse)
        xmlResponseDoc.loadXML strCrudResponse
    
        If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            
            'IK_29/11/2005_MAR715
            If Not xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE") Is Nothing Then
                xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text = _
                    "IngestionManagerBO.ProcessSubmit." & xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text
                ProcessSubmit = xmlResponseDoc.xml
            Else
                ProcessSubmit = strCrudResponse
            End If
            
            vxmlPreProcResponse.setAttribute "UPDATEABORT", "true"
            vxmlPreProcResponse.setAttribute "PostProcTaskId", "GeneralFailure"
            ProcessPostProc vxmlPreProcResponse
        
        Else
            ProcessSubmit = ProcessPostProc(vxmlPreProcResponse)
        End If
    
    End If
    
ProcessSubmitExit:

    Set xmlResponseDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function ProcessStopAndSave(ByVal vxmlPreProcResponse As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "ProcessStopAndSave"
    On Error GoTo ProcessStopAndSaveExit
    
    Dim xmlResponseDoc As DOMDocument40
    Dim strCrudResponse As String
        
    'IK_18/11/2006_EP2_134
    Select Case vxmlPreProcResponse.getAttribute("omigaClient")
        Case "epsom"
            ' nothing to do
        Case Else
            CriticalDataTest vxmlPreProcResponse
    End Select
    ''IK_18/11/2006_EP2_134_ends
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    
    strCrudResponse = CrudCall(vxmlPreProcResponse)
    xmlResponseDoc.loadXML strCrudResponse
    
    If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            
        'IK_29/11/2005_MAR715
        If Not xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE") Is Nothing Then
            xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text = _
                "IngestionManagerBO.ProcessStopAndSave." & xmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text
            ProcessStopAndSave = xmlResponseDoc.xml
        Else
            ProcessStopAndSave = strCrudResponse
        End If
        
        vxmlPreProcResponse.setAttribute "UPDATEABORT", "true"
        vxmlPreProcResponse.setAttribute "PostProcTaskId", "GeneralFailure"
        ProcessPostProc vxmlPreProcResponse
    
    Else
        
        ProcessStopAndSave = ProcessPostProc(vxmlPreProcResponse)
    
    End If
    
ProcessStopAndSaveExit:

    Set xmlResponseDoc = Nothing
    
    CheckError cstrMethodName

End Function

'IK_18/11/2006_EP2_134
Private Sub CriticalDataTest(ByVal vxmlPreProcResponse As IXMLDOMElement)
    
    Const cstrMethodName As String = "CriticalDataTest"
    On Error GoTo CriticalDataTestExit
    
    Dim xmlResponseDoc As DOMDocument40
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    
    If vxmlPreProcResponse.selectSingleNode("APPLICATION[@CRUD_OP='CREATE']") Is Nothing Then
    
        ' i.e. not new application
        If AnyPreviousCreditCheck(vxmlPreProcResponse) Then
    
            xmlResponseDoc.loadXML GetApplicationData(vxmlPreProcResponse)
            
            If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/APPLICATION") Is Nothing Then
                gstrComponentId = "omCRUD.omCRUDBO"
                gstrComponentResponse = xmlResponseDoc.xml
                Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving comparison data "
            End If
            
            ' creates attribute 'PostProcTaskId' containing omTm OPERATION value
            ' on REQUEST node
            CriticalDataCheck vxmlPreProcResponse, xmlResponseDoc.xml
            
        End If
        
    End If
    
CriticalDataTestExit:

    Set xmlResponseDoc = Nothing
    
    CheckError cstrMethodName

End Sub
'IK_18/11/2006_EP2_134_ends

Private Function ProcessPostProc(ByVal vxmlPreProcResponse As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "ProcessPostProc"
    On Error GoTo ProcessPostProcExit
    ProcessPostProc = PostProcRequest(vxmlPreProcResponse)
    
ProcessPostProcExit:
    
    If vxmlPreProcResponse.Attributes.getNamedItem("UPDATEABORT") Is Nothing Then
        CheckError cstrMethodName
    '   otherwise return CRUD failure
    End If

End Function

' IK_09/11/2005_MAR498
Private Function AnyPreviousCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    
    Const cstrMethodName As String = "AnyPreviousCreditCheck"
    On Error GoTo AnyPreviousCreditCheckExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strComponentResponse As String
    
    strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    strApplicationFactFindNumber = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONCREDITCHECK")
    xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem

    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    strComponentResponse = crudObj.OmRequest(xmlDoc.xml)
    Set crudObj = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error checking for existing APPLICATIONCREDITCHECK"
    End If
    
    If xmlDoc.selectNodes("RESPONSE/APPLICATIONCREDITCHECK").length > 0 Then
        AnyPreviousCreditCheck = True
    End If
    
AnyPreviousCreditCheckExit:
    
    Set crudObj = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    CheckError cstrMethodName

End Function

' IK_09/11/2005_MAR498
Private Sub CriticalDataCheck( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vstrBeforeImage As String)
    
    Const cstrMethodName As String = "CriticalDataCheck"
    On Error GoTo CriticalDataCheckExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlBeforeNode As IXMLDOMNode
    Dim xmlAfterNode As IXMLDOMNode
    
    Dim objCDRules As Object
    
    Dim strComponentResponse As String

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML vstrBeforeImage

    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    For Each xmlAttrib In vxmlRequestNode.Attributes
        xmlElem.setAttribute xmlAttrib.Name, xmlAttrib.Value
    Next
    Set xmlRequestNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlRequestDoc.createElement("BEFORE")
    Set xmlBeforeNode = xmlRequestNode.appendChild(xmlElem)
    
    For Each xmlNode In vxmlRequestNode.childNodes
        xmlBeforeNode.appendChild xmlNode.cloneNode(True)
    Next
    
    Set xmlElem = xmlRequestDoc.createElement("AFTER")
    Set xmlAfterNode = xmlRequestNode.appendChild(xmlElem)
    
    For Each xmlNode In xmlDoc.selectSingleNode("RESPONSE").childNodes
        xmlAfterNode.appendChild xmlNode.cloneNode(True)
    Next
    
    xmlTransform xmlRequestDoc, "transformCriticalData.xslt"
    
    Set objCDRules = GetObjectContext.CreateInstance("omCDRules.omCDRulesBO")
    If objCDRules Is Nothing Then
        Err.Raise oeUnspecifiedError, cstrMethodName, "error creating Critical Data rules component: omCDRules.omCDRulesBO"
    End If
    
    strComponentResponse = objCDRules.RunRules(xmlRequestDoc.xml)
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        gstrComponentId = "omCDRules.omCDRulesBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error running Critical Data rules "
    End If
    
    If Not xmlDoc.selectSingleNode("RESPONSE/CASETASK[@TASKID = 'Exp_XMLReprocess']") Is Nothing Then
        vxmlRequestNode.setAttribute "PostProcTaskId", "RunXMLReprocessCreditCheck"
    ElseIf Not xmlDoc.selectSingleNode("RESPONSE/CASETASK[@TASKID = 'Exp_XMLRescore']") Is Nothing Then
        vxmlRequestNode.setAttribute "PostProcTaskId", "RunXMLRescoreCreditCheck"
    End If
    
CriticalDataCheckExit:
    
    Set objCDRules = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    
    Set xmlRequestNode = Nothing
    Set xmlBeforeNode = Nothing
    Set xmlAfterNode = Nothing
    Set xmlRequestDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Function CrudCall(ByVal vxmlRequestNode As IXMLDOMNode) As String
    
    Const cstrMethodName As String = "CrudCall"
    On Error GoTo CrudCallExit
    
    Dim xmlCrudRequestDoc As DOMDocument40
    Dim objCRUD As Object
    
    Set xmlCrudRequestDoc = New DOMDocument40
    xmlCrudRequestDoc.async = False
    xmlCrudRequestDoc.setProperty "NewParser", True
    xmlCrudRequestDoc.loadXML vxmlRequestNode.xml
        
    RemoveRedundantNodes xmlCrudRequestDoc.documentElement
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDTxBO")
    CrudCall = objCRUD.OmRequest(xmlCrudRequestDoc.xml)
    Set objCRUD = Nothing
    
CrudCallExit:
    
    Set objCRUD = Nothing
    Set xmlCrudRequestDoc = Nothing

End Function

Private Sub RemoveRedundantNodes(ByVal vxmlRequestNode As IXMLDOMNode)
    
    Const cstrMethodName As String = "RemoveRedundantNodes"
    On Error GoTo RemoveRedundantNodesExit

    Dim xmlChildNode As IXMLDOMNode
    Dim xmlAppFFChildNode As IXMLDOMNode
        
    For Each xmlChildNode In vxmlRequestNode.childNodes
        If AnyCrudOp(xmlChildNode) Then
            If Not xmlChildNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
                ' do nothing if CREATE, all child nodes are implicit CREATEs
                If xmlChildNode.Attributes.getNamedItem("CRUD_OP").Text <> "CREATE" Then
                    RemoveRedundantNodes xmlChildNode
                End If
            Else
                RemoveRedundantNodes xmlChildNode
            End If
        Else
            ' ik_20051221_mar919
            If CanDeleteNode(xmlChildNode) Then
                vxmlRequestNode.removeChild xmlChildNode
            Else
                RemoveRedundantNodes xmlChildNode
            End If
        End If
    Next
    
RemoveRedundantNodesExit:
    
    CheckError cstrMethodName
    
End Sub

Private Function AnyCrudOp(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    Const cstrMethodName As String = "AnyCrudOp"
    On Error GoTo AnyCrudOpExit

    Dim xmlChildNode As IXMLDOMNode

    If Not vxmlRequestNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
        If vxmlRequestNode.Attributes.getNamedItem("CRUD_OP").Text <> "iUPDATE" Then
            AnyCrudOp = True
        End If
    End If
    
    If Not AnyCrudOp Then
        For Each xmlChildNode In vxmlRequestNode.childNodes
            If AnyCrudOp(xmlChildNode) Then
                AnyCrudOp = True
            End If
        Next
    End If
    
AnyCrudOpExit:
    
    CheckError cstrMethodName

End Function

' ik_20051221_mar919
Private Function CanDeleteNode(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    If gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & vxmlRequestNode.nodeName & "'][@noTrim]") Is Nothing Then
        CanDeleteNode = True
    End If
End Function

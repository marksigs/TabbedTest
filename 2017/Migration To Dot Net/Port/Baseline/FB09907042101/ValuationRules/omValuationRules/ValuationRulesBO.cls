VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ValuationRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'PROG   Date            Description
'JJ     08-Nov-2005     omValuation Rules
'DRC    14-Nov-2005     MAR529 HOMETRACK Logic fix
'DRC    17-Nov-2005     MAR610 Change taskid for Hometrack->eSurv call
'PE     27-Mar-2007     EP2_1578 - Added to Epsom2
'PE     04-Apr-2007     EP2_1578 - Redeveloped to call omRARules.RARulesBO.RiskAssessmentRules
'PE     13-Apr-2007     EP2_2360 - Modified task creation to stop SQL Server timeout (???).
'-----------------------------------------------------------------------------------

Option Explicit

Public Function RunValuationRules(ByVal vstrgxmlRequest As String) As String

    On Error GoTo RunValuationRulesExit
    
    Const cstrMethodName As String = "RunValuationRules"
    Dim objRARules As OmRARules.RARulesBO
    Dim strXMLResponse As String
    Dim xmlDocRequest As FreeThreadedDOMDocument40
    Dim xmlDocRules As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlRulesRequestNode As IXMLDOMElement
    Dim xmlApplicationNode As IXMLDOMElement
    Dim XMLResponseDoc As FreeThreadedDOMDocument40
        
    Set xmlDocRequest = New FreeThreadedDOMDocument40
    xmlDocRequest.setProperty "NewParser", True
    xmlDocRequest.async = False
    xmlDocRequest.loadXML vstrgxmlRequest
       
    If xmlDocRequest.parseError.errorCode <> 0 Then
        xmlParseError xmlDocRequest.parseError
    End If
    
    Set xmlRequestNode = xmlDocRequest.selectSingleNode("REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    Set xmlApplicationNode = xmlRequestNode.selectSingleNode("APPLICATION")
    If xmlApplicationNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION node"
    End If
    
    Set xmlDocRules = New FreeThreadedDOMDocument40
    xmlDocRules.appendChild xmlRequestNode.cloneNode(False)
    
    Set xmlRulesRequestNode = xmlDocRules.selectSingleNode("REQUEST")
    GetApplicationData xmlApplicationNode, xmlRulesRequestNode
    xmlRulesRequestNode.setAttribute "CONTEXT", "VEX"
    
    Set objRARules = GetObjectContext.CreateInstance("omRARules.RARulesBO")
    strXMLResponse = objRARules.RiskAssessmentRules(xmlDocRules.xml)
    
    Set XMLResponseDoc = New FreeThreadedDOMDocument40
    XMLResponseDoc.setProperty "NewParser", True
    XMLResponseDoc.async = False
    XMLResponseDoc.loadXML strXMLResponse
    
    If XMLResponseDoc.selectSingleNode("/RESPONSE[@FAILLEVEL>0]") Is Nothing Then
        CreateTask xmlRequestNode, GetGlobalParameter("TMVexValReportSatisfied")
    Else
        CreateTask xmlRequestNode, GetGlobalParameter("TMVexValReportNotSatisfied")
    End If
    
    strXMLResponse = XMLResponseDoc.xml
    
RunValuationRulesExit:

    Set objRARules = Nothing
    RunValuationRules = strXMLResponse
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If
    
End Function

Private Sub GetApplicationData( _
    ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
    ByVal vxmlApplicationDestNode As IXMLDOMNode)
    
    On Error GoTo GetApplicationDataExit
    Const cstrMethodName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim XMLResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Text = "RARulesTemplate"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.Attributes.getNamedItem("APPLICATIONNUMBER").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Set XMLResponseDoc = New FreeThreadedDOMDocument40
    XMLResponseDoc.validateOnParse = False
    XMLResponseDoc.setProperty "NewParser", True
    XMLResponseDoc.async = False
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    XMLResponseDoc.loadXML omCrudInterface(xmlRequestDoc.documentElement, "raRules")
    
    Debug.Print "GetApplicationData response"
    Debug.Print XMLResponseDoc.xml
    
    For Each xmlNode In XMLResponseDoc.selectSingleNode("RESPONSE").childNodes
        vxmlApplicationDestNode.appendChild xmlNode.cloneNode(True)
    Next
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    Set XMLResponseDoc = Nothing
    
GetApplicationDataExit:

    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If
    
End Sub

Private Sub CreateTask(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vstrTaskId)
    
    Const cstrMethodName As String = "CreateTask"
    On Error GoTo CreateTaskByTaskRefExit
    
    Dim xmlCaseActivityDoc As DOMDocument40
    Dim xmlRequestDoc As DOMDocument40
    Dim XMLResponseDoc As DOMDocument40
    Dim xmlApplicationPriority As IXMLDOMElement
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objTm As omTM.OmTmBO
        
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    GetCurrentStage xmlCaseActivityDoc, xmlCaseActivityDoc
        
    Set xmlApplicationPriority = xmlCaseActivityDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY")
    'Set xmlCaseActivity = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASESTAGE")
    
    If Not OutstandingTaskExists(xmlCaseStage, vstrTaskId) Then
    
        Set xmlRequestDoc = New DOMDocument40
        xmlRequestDoc.async = False
        xmlRequestDoc.setProperty "NewParser", True
        
        Set xmlElem = xmlRequestDoc.createElement("REQUEST")
        xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
        xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
        xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
        xmlElem.setAttribute "OPERATION", "CreateAdhocCaseTask"
        Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("CASETASK")
        xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseStage.getAttribute("SOURCEAPPLICATION")
        xmlElem.setAttribute "CASEID", xmlCaseStage.getAttribute("CASEID")
        xmlElem.setAttribute "ACTIVITYID", xmlCaseStage.getAttribute("ACTIVITYID")
        xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseStage.getAttribute("ACTIVITYINSTANCE")
        xmlElem.setAttribute "STAGEID", xmlCaseStage.getAttribute("STAGEID")
        xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
        xmlElem.setAttribute "TASKID", vstrTaskId
        xmlElem.setAttribute "APPLICATIONPRIORITY", ""
        'xmlElem.setAttribute "TASKSTATUS", 10
        
        xmlNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
        xmlElem.setAttribute "APPLICATIONPRIORITY", ""
        xmlNode.appendChild xmlElem
    
        Set XMLResponseDoc = New DOMDocument40
        XMLResponseDoc.async = False
        XMLResponseDoc.setProperty "NewParser", True
    
        Set objTm = GetObjectContext.CreateInstance("omTm.omTmBO")
        XMLResponseDoc.loadXML objTm.OmTmRequest(xmlRequestDoc.xml)
        Set objTm = Nothing
            
        If XMLResponseDoc.parseError.errorCode <> 0 Then
            xmlParseError XMLResponseDoc.parseError
        End If
        
        If XMLResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            Err.Raise oeUnspecifiedError, cstrMethodName, "error creating task"
        End If
    
    End If
    
CreateTaskByTaskRefExit:
        
    Set objTm = Nothing
    
    Set xmlCaseActivityDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlApplicationPriority = Nothing
    Set xmlRequestDoc = Nothing
    Set XMLResponseDoc = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If
    
End Sub

Public Sub GetCurrentStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseDoc As DOMDocument40)
    
    Const cstrMethodName As String = "GetCurrentStage"
    On Error GoTo GetCurrentStageExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim msgTMobj As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "OPERATION", "GetCurrentStage"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "CASEACTIVITYGUID", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@CASEACTIVITYGUID").Text
    xmlElem.setAttribute "SOURCEAPPLICATION", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@SOURCEAPPLICATION").Text
    xmlElem.setAttribute "CASEID", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@CASEID").Text
    xmlElem.setAttribute "ACTIVITYID", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@ACTIVITYID").Text
    xmlElem.setAttribute "ACTIVITYINSTANCE", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@ACTIVITYINSTANCE").Text
    xmlElem.setAttribute "ACTIVITYNAME", vxmlRequestNode.selectSingleNode("RESPONSE/CASEACTIVITY/@ACTIVITYNAME").Text
    xmlNode.appendChild xmlElem
        
    Set msgTMobj = GetObjectContext.CreateInstance("MsgTm.MsgTmBO")
    vxmlResponseDoc.loadXML msgTMobj.TmRequest(xmlDoc.xml)
    Set msgTMobj = Nothing
    
    If vxmlResponseDoc.parseError.errorCode <> 0 Then
        xmlParseError vxmlResponseDoc.parseError
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        Err.Raise oeUnspecifiedError, cstrMethodName, vxmlResponseDoc.xml
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/CASESTAGE") Is Nothing Then
        Err.Raise oeXMLMissingElement, cstrMethodName, "error retrieving CASEASTAGE details"
    End If
    
GetCurrentStageExit:
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If

End Sub

Public Sub GetCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseDoc As DOMDocument40)
    
    Const cstrMethodName As String = "GetCaseActivity"
    On Error GoTo GetCaseActivityExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim msgTMobj As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "OPERATION", "GETCASEACTIVITY"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "CASEID", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
        
    Set msgTMobj = GetObjectContext.CreateInstance("MsgTm.MsgTmBO")
    vxmlResponseDoc.loadXML msgTMobj.TmRequest(xmlDoc.xml)
    Set msgTMobj = Nothing
    
    If vxmlResponseDoc.parseError.errorCode <> 0 Then
        xmlParseError vxmlResponseDoc.parseError
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        Err.Raise oeUnspecifiedError, cstrMethodName, vxmlResponseDoc.xml
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/CASEACTIVITY") Is Nothing Then
        Err.Raise oeXMLMissingElement, cstrMethodName, "error retrieving CASEACTIVITY details"
    End If
    
GetCaseActivityExit:
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If

End Sub

Public Function OutstandingTaskExists( _
    ByVal vxmlCaseStage As IXMLDOMNode, _
    ByVal vstrTaskId As String) _
    As Boolean
    
    On Error GoTo OutstandingTaskExistsExit
    
    Const cstrMethodName As String = "OutstandingTaskExists"
    On Error GoTo OutstandingTaskExistsExit
    
    Dim strXpath As String
        
    strXpath = "CASETASK[@TASKID='" & vstrTaskId & "'][@TASKSTATUS='10' or @TASKSTATUS='20']"
    
    If vxmlCaseStage.selectNodes(strXpath).length > 0 Then
        OutstandingTaskExists = True
    End If
    
OutstandingTaskExistsExit:
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If
    
End Function

Public Sub xmlParseError(ByVal objParseError As IXMLDOMParseError)
    
    Dim strErrDesc As String    ' formatted parser error
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str$(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str$(objParseError.Line) & vbCr & _
        "Character: " & Str$(objParseError.linepos) & vbCr & _
        "Source text: " & objParseError.srcText
        
        Err.Raise oeXMLParserError, "xmlParseError", strErrDesc
            
End Sub

Public Function omCrudInterface( _
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
    omCrudInterface = omCRUD.OmRequest(vxmlRequest.xml)
    
omCrudInterfaceExit:

    Set omCRUD = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If

End Function

Public Function GetGlobalParameter(ByVal strParameterName As String) As String

On Error GoTo GetGlobalParameterExit

    Const cstrMethodName As String = "GetGlobalParameter"
    Dim xmlSchema As FreeThreadedDOMDocument40
    Dim XMLRequest As IXMLDOMElement
    Dim xmlGlobalParamater As IXMLDOMElement
    Dim strResponse As String
    Dim XMLResponse As FreeThreadedDOMDocument40
    
    Set xmlSchema = New FreeThreadedDOMDocument40
    xmlSchema.setProperty "NewParser", True
    xmlSchema.async = False

    Set XMLRequest = xmlSchema.createElement("REQUEST")
    Set xmlGlobalParamater = xmlSchema.createElement("GLOBALPARAMETER")
    xmlGlobalParamater.setAttribute "NAME", strParameterName
    
    XMLRequest.appendChild xmlGlobalParamater
    xmlSchema.appendChild XMLRequest
    
    strResponse = omCrudInterface(xmlSchema.documentElement, "GLOBALPARAMETER")
    
    Set XMLResponse = New FreeThreadedDOMDocument40
    XMLResponse.setProperty "NewParser", True
    XMLResponse.async = False
    XMLResponse.loadXML strResponse
       
    If XMLResponse.parseError.errorCode <> 0 Then
        xmlParseError XMLResponse.parseError
    End If
    
    If Not (XMLResponse.selectSingleNode("/RESPONSE[@TYPE=""SUCCESS""]") Is Nothing) Then
        GetGlobalParameter = XMLResponse.selectSingleNode("/RESPONSE/GLOBALPARAMETER/@STRING").Text
    End If
    
GetGlobalParameterExit:

    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
    End If
    
End Function

Public Function CSafeDbl(ByVal vvntExpression As Variant) As Double
' header ----------------------------------------------------------------------------------
' description:
'   creates a double representation of an expression
' pass:
'   Expression  Expression to be converted to a double
' Returns:
'   CSafeDbl   Converted Expression
'------------------------------------------------------------------------------------------
On Error GoTo CSafeDblVbErr

    Const strFunctionName As String = "CSafeDbl"
    
    Dim dblRetValue As Double

    If Len(vvntExpression) > 0 Then
        dblRetValue = CDbl(vvntExpression)
    End If
    
    CSafeDbl = dblRetValue

    Exit Function

CSafeDblVbErr:

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

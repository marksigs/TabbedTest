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
'IK     18/10/2005  created for Project MARS
'IK     22/10/2005  MAR271 return APPLICATIONNUMBER when CREATE
'IK     09/11/2005  MAR498 - re-structuring, call critical data check before omCRUD
'IK     25/11/2005  MAR695 gstrComponentId & gstrComponentResponse now private member variables
'IK     05/12/2005  MAR807 modify MoveToStage, do not move if already at stage
'IK     07/03/2006  EP198, EP199 stored procedures in WebServiceSchema
'IK     09/03/2006  EP2 - add EpsomDecision
'IK     28/03/2006  EP304 - Add TaskName to config for use by KYC task
'                           Credit Check task set complete in omTm
Option Explicit

Private gstrComponentId As String
Private gstrComponentResponse As String


Public Function OmRequest(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "OmRequest"
    On Error GoTo OmRequestVbErr
    
    Dim xmlInDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    
    ' default response - will be overwritten as required
    GetObjectContext.SetComplete
    OmRequest = "<RESPONSE TYPE='SUCCESS'/>"
    
    LoadConfigFile
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    ' Check for Parse errors in XML
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    ' Set context to the REQUEST node
    Set xmlRequestNode = xmlInDoc.selectSingleNode("REQUEST")
    
    ' Return error if REQUEST node not found
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    ' Ensure Mandatory Tags exist
    ValidateRequest xmlRequestNode
    
    ' If a Create CRUD_Op Exists, Create Case and Move to next Stage
    If Not xmlRequestNode.selectSingleNode("APPLICATION[@CRUD_OP='CREATE']") Is Nothing Then
        CreateCaseActivity xmlRequestNode '----->>
        MoveToStage xmlRequestNode, "PreAiP" '----->>
    End If
    
    ' If an ACCEPTQUOTE attribute exists, do Accept Quote routine
    If Not xmlRequestNode.Attributes.getNamedItem("ACCEPTQUOTE") Is Nothing Then
        AcceptQuote xmlRequestNode '----->>
    End If
    
    ' What Operation was specified in the OPERATION Attribute
    Select Case xmlRequestNode.getAttribute("OPERATION")

        Case "SubmitAiP"
            MoveToStage xmlRequestNode, "SubmitAiP" '----->>

        Case "SubmitFMA"
            MoveToStage xmlRequestNode, "SubmitFMA" '----->>

    End Select

    ' Do Post Process Rules and set Return Value
    'IK_17/03/2006_EP2
    Select Case xmlRequestNode.getAttribute("omigaClient")
    
        Case "epsom"

            OmRequest = EpsomPostProcRules(xmlRequestNode) '----->>
            
        Case Else

            OmRequest = PostProcRules(xmlRequestNode) '----->>
    
    End Select
    'IK_09/03/2006_EP2_ends
    
OmRequestVbErr:

    Set xmlInDoc = Nothing
    
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
    
    If vxmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no OPERATION on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERID on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("UNITID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no UNITID on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERAUTHORITYLEVEL on post-processing REQUEST node"
    End If

    If vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION/@APPLICATIONNUMBER on post-processing REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER on post-processing REQUEST node"
    End If
    
ValidateRequestExit:
    
    CheckError cstrMethodName

End Sub

Private Sub CreateCaseActivity(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Const cstrMethodName As String = "CreateCaseActivity"
    On Error GoTo CreateCaseActivityExit
    
    Dim objOmTm As Object
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strComponentResponse As String
    
    strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    strApplicationFactFindNumber = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "OPERATION", "CREATEACTIVITY"
    xmlElem.setAttribute "USERID", vxmlRequestNode.Attributes.getNamedItem("USERID").Text
    xmlElem.setAttribute "UNITID", vxmlRequestNode.Attributes.getNamedItem("UNITID").Text
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "ACTIVITYID", "10"
    xmlElem.setAttribute "SOURCEAPPLICATION", "Omiga"
    xmlElem.setAttribute "CASEID", strApplicationNumber
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
    
    Set objOmTm = GetObjectContext.CreateInstance("omTm.omTmBO")
    strComponentResponse = objOmTm.OmTmRequest(xmlDoc.xml)
    Set objOmTm = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        
        ' create general failure task
        vxmlRequestNode.setAttribute "UPDATEABORT", "true"
        vxmlRequestNode.setAttribute "PostProcTaskId", "GeneralFailure"
        CreateTask vxmlRequestNode
        
        gstrComponentId = "omTm.omTmBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error creating CASEACTIVITY"
    
    End If
    
CreateCaseActivityExit:
    
    Set objOmTm = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub AcceptQuote(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Const cstrMethodName As String = "AcceptQuote"
    On Error GoTo AcceptQuoteExit
    
    Dim objAQ As Object
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strComponentResponse As String
    
    strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    strApplicationFactFindNumber = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "USERID", vxmlRequestNode.Attributes.getNamedItem("USERID").Text
    xmlElem.setAttribute "UNITID", vxmlRequestNode.Attributes.getNamedItem("UNITID").Text
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("QUOTATION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strApplicationNumber
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
    
    Set objAQ = GetObjectContext.CreateInstance("omAQ.ApplicationQuoteBO")
    strComponentResponse = objAQ.AcceptQuotation(xmlDoc.xml)
    Set objAQ = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        
        ' create general failure task
        vxmlRequestNode.setAttribute "UPDATEABORT", "true"
        vxmlRequestNode.setAttribute "PostProcTaskId", "GeneralFailure"
        CreateTask vxmlRequestNode
        
        gstrComponentId = "omAQ.ApplicationQuoteBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error accepting Quotation"
    
    End If
    
AcceptQuoteExit:
    
    Set objAQ = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub MoveToStage(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vstrOperation As String)
    
    Const cstrMethodName As String = "MoveToStage"
    On Error GoTo MoveToStageExit
    
    Dim objOmTm As Object
    
    Dim xmlCaseActivityDoc As DOMDocument40
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strStageId As String, _
        strComponentResponse As String
        
    ' Get the Stage Number for this Operation (from global config)
    Set xmlElem = _
        gxmldocConfig.selectSingleNode( _
            "omIngestion/operation[@name='" & vstrOperation & "'][@stage]")
            
    ' IK_EP2_21/03/2006
    ' If Stage number not found in the config doc, Abort (no error)
    If Not xmlElem Is Nothing Then
    
        ' Get the stage ID from the config doc into a variable
        strStageId = xmlElem.getAttribute("stage")
    
        ' Create a new Case Document
        Set xmlCaseActivityDoc = New DOMDocument40
        xmlCaseActivityDoc.async = False
        xmlCaseActivityDoc.setProperty "NewParser", True
    
        ' Get Case tasks
        GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc '----->>
        
        
        'IK_05/12/2005_MAR807
        ' If the Case stage is not at the required stage, Generate a task to move it to the required stage
        If xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGE[@STAGEID='" & strStageId & "']") Is Nothing Then
        
            strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
            strApplicationFactFindNumber = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
            
            ' Create a new doc
            Set xmlDoc = New DOMDocument40
            xmlDoc.setProperty "NewParser", True
            xmlDoc.async = False
            
            ' Add a REQUEST Element to the new doc and set the OPERATION to MOVETOSTAGE
            Set xmlElem = xmlDoc.createElement("REQUEST")
            xmlElem.setAttribute "OPERATION", "MOVETOSTAGE"
            xmlElem.setAttribute "USERID", vxmlRequestNode.Attributes.getNamedItem("USERID").Text
            xmlElem.setAttribute "UNITID", vxmlRequestNode.Attributes.getNamedItem("UNITID").Text
            xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
            Set xmlNode = xmlDoc.appendChild(xmlElem)
            
            ' Add a CASESTAGE Element to the new doc and set attributes
            Set xmlElem = xmlDoc.createElement("CASESTAGE")
            xmlElem.setAttribute "ACTIVITYID", "10"
            xmlElem.setAttribute "ACTIVITYINSTANCE", "1"
            xmlElem.setAttribute "SOURCEAPPLICATION", "Omiga"
            xmlElem.setAttribute "CASEID", strApplicationNumber
            xmlElem.setAttribute "STAGEID", strStageId
            xmlNode.appendChild xmlElem
    
            ' Add a APPLICATION Element to the new doc and set application attributes
            Set xmlElem = xmlDoc.createElement("APPLICATION")
            xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
            xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
            xmlNode.appendChild xmlElem
            
            ' Instantiate the TaskManager
            Set objOmTm = GetObjectContext.CreateInstance("omTm.omTmNoTxBO")
            
            ' Get TaskManager to execute this task
            strComponentResponse = objOmTm.omTmNoTxRequest(xmlDoc.xml) '---->>
            Set objOmTm = Nothing
            
            ' Put the Task Manager response string into an XML Doc
            xmlDoc.loadXML strComponentResponse
            
            ' If Task Manager didn't return SUCCESS then generate an error
            If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
                
                ' create general failure task
                vxmlRequestNode.setAttribute "UPDATEABORT", "true"
                vxmlRequestNode.setAttribute "PostProcTaskId", "GeneralFailure"
                CreateTask vxmlRequestNode
                
                gstrComponentId = "omTm.omTmNoTxBO"
                gstrComponentResponse = strComponentResponse
                Err.Raise oeUnspecifiedError, cstrMethodName, "error calling MOVETOSTAGE"
            
            End If
        
        End If
        'IK_05/12/2005_MAR807_ends
        
    End If
    
MoveToStageExit:
    
    Set objOmTm = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseActivityDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Function PostProcRules(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "PostProcRules"
    On Error GoTo PostProcRulesExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    ' replaced as required
    PostProcRules = "<RESPONSE TYPE='SUCCESS'/>"
    
    If Not vxmlRequestNode.Attributes.getNamedItem("PostProcTaskId") Is Nothing Then
        CreateTask vxmlRequestNode
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("UPDATEABORT") Is Nothing Then
        
        If Not vxmlRequestNode.selectSingleNode("APPLICATION[@CRUD_OP='CREATE']") Is Nothing Then
    
            Set xmlDoc = New DOMDocument40
            xmlDoc.setProperty "NewParser", True
            xmlDoc.async = False
            
            Set xmlElem = xmlDoc.createElement("RESPONSE")
            xmlElem.setAttribute "TYPE", "SUCCESS"
            Set xmlNode = xmlDoc.appendChild(xmlElem)
            
            Set xmlElem = xmlDoc.createElement("APPLICATION")
            If Not vxmlRequestNode.selectSingleNode("APPLICATION[@APPLICATIONNUMBER]") Is Nothing Then
                xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
            End If
            
            xmlNode.appendChild xmlElem
            
            PostProcRules = xmlDoc.xml
    
        End If
        
        If Not vxmlRequestNode.Attributes.getNamedItem("PostProcTaskId") Is Nothing Then
            PostProcRules = GetDecision(vxmlRequestNode)
        End If
        
    End If
    
PostProcRulesExit:
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function GetDecision(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "GetDecision"
    On Error GoTo GetDecisionExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim strOperation As String, _
        strComponentResponse As String, _
        strDecision As String
    
    Dim lngErr As Long  'MAR1084 GHun
    
    strOperation = vxmlRequestNode.getAttribute("PostProcTaskId")
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    Set xmlNode = xmlDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    xmlNode.setAttribute "OPERATION", strOperation
    xmlNode.setAttribute "NOADDRESSTARGETING", "1" 'MAR1084 GHun
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Dim objTm As Object
    Set objTm = GetObjectContext.CreateInstance("omTm.omTmBO")
    strComponentResponse = objTm.OmTmRequest(xmlDoc.xml)
    Set objTm = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/DECISION") Is Nothing Then
        'MAR1084 GHun If the error is not address targeting then raise it
        lngErr = 0
        Set xmlElem = xmlDoc.selectSingleNode("RESPONSE[@TYPE='APPERR']/ERROR/NUMBER")
        If Not xmlElem Is Nothing Then
            If Len(xmlElem.Text) > 0 Then
                lngErr = errGetOmigaErrorNumber(CLng(xmlElem.Text))
            End If
        End If
        
        If lngErr <> 8540 Then
        'MAR1084 End
            gstrComponentId = "omTm.omTmBO"
            gstrComponentResponse = strComponentResponse
            Err.Raise oeUnspecifiedError, cstrMethodName, "error in credit check operation " & strOperation
        End If
    Else
    
        strDecision = xmlDoc.selectSingleNode("RESPONSE/DECISION").Text
        
        xmlDoc.loadXML ""
        
        Set xmlElem = xmlDoc.createElement("RESPONSE")
        xmlElem.setAttribute "TYPE", "SUCCESS"
        Set xmlNode = xmlDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlDoc.createElement("APPLICATION")
        xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
        xmlElem.setAttribute "UNDERWRITERSDECISION", strDecision
        xmlNode.appendChild xmlElem
        
        If Not CompleteCreditCheckTask(vxmlRequestNode, strOperation) Then
            
            Set xmlElem = xmlDoc.createElement("MESSAGE")
            Set xmlNode = xmlNode.appendChild(xmlElem)
            
            Set xmlElem = xmlDoc.createElement("MESSAGETEXT")
            xmlElem.Text = "unable to update task status for TASKID " & GetTaskId(strOperation)
            xmlNode.appendChild xmlElem
            
            Set xmlElem = xmlDoc.createElement("MESSAGETYPE")
            xmlElem.Text = "WARNING"
            xmlNode.appendChild xmlElem
        
        End If
        
        'MAR1143 GHun
        If strOperation = "RunXMLRescoreCreditCheck" And Not (strDecision = "3" Or strDecision = "4" Or strDecision = "5") Then
            For Each xmlElem In gxmldocConfig.documentElement.selectNodes("operation[@name='" & strOperation & "']/task")
                vxmlRequestNode.setAttribute "PostProcTaskId", xmlElem.getAttribute("id")
                CreateTask vxmlRequestNode
            Next
        End If
        'MAR1143 End
    End If
    
    GetDecision = xmlDoc.xml
    
GetDecisionExit:
    
    Set objTm = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function CompleteCreditCheckTask( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vstrCreditCheckOp As String) _
    As Boolean
    
    Const cstrMethodName As String = "CompleteCreditCheckTask"
    On Error Resume Next
    
    Dim xmlCaseActivityDoc As DOMDocument40
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlCaseTask As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objTm As Object
    
    Dim strTaskId As String
    strTaskId = GetTaskId(vstrCreditCheckOp)
    
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlCaseActivity = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivity.selectSingleNode("CASESTAGE")
    
    If OutstandingTaskExists(xmlCaseStage, strTaskId) Then
    
        Set xmlCaseTask = GetOutstandingTask(xmlCaseStage, strTaskId)
    
        Set xmlRequestDoc = New DOMDocument40
        xmlRequestDoc.async = False
        xmlRequestDoc.setProperty "NewParser", True
        
        Set xmlElem = xmlRequestDoc.createElement("REQUEST")
        xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
        xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
        xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
        xmlElem.setAttribute "OPERATION", "UpdateCaseTask"
        Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("CASETASK")
        xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseActivity.getAttribute("SOURCEAPPLICATION")
        xmlElem.setAttribute "CASEID", xmlCaseActivity.getAttribute("CASEID")
        xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseActivity.getAttribute("ACTIVITYINSTANCE")
        xmlElem.setAttribute "ACTIVITYID", xmlCaseActivity.getAttribute("ACTIVITYID")
        xmlElem.setAttribute "STAGEID", xmlCaseStage.getAttribute("STAGEID")
        xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
        xmlElem.setAttribute "TASKID", xmlCaseTask.getAttribute("TASKID")
        xmlElem.setAttribute "TASKINSTANCE", xmlCaseTask.getAttribute("TASKINSTANCE")
        xmlElem.setAttribute "CASEACTIVITYGUID", xmlCaseTask.getAttribute("CASEACTIVITYGUID")
        xmlElem.setAttribute "TASKSTATUS", "40"
        xmlNode.appendChild xmlElem
    
        Set xmlResponseDoc = New DOMDocument40
        xmlResponseDoc.async = False
        xmlResponseDoc.setProperty "NewParser", True
    
        Set objTm = GetObjectContext.CreateInstance("msgTm.msgTmBO")
        xmlResponseDoc.loadXML objTm.TmRequest(xmlRequestDoc.xml)
        Set objTm = Nothing
        
        If Not xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            CompleteCreditCheckTask = True
        End If
    
        If vstrCreditCheckOp = "RunXMLReprocessCreditCheck" Then
        
            strTaskId = GetTaskId("RunXMLRescoreCreditCheck")
        
            If OutstandingTaskExists(xmlCaseStage, strTaskId) Then
            
                Set xmlCaseTask = GetOutstandingTask(xmlCaseStage, strTaskId)
                
                Set xmlElem = xmlRequestDoc.selectSingleNode("REQUEST/CASETASK")
                xmlElem.setAttribute "TASKID", strTaskId
                xmlElem.setAttribute "TASKINSTANCE", xmlCaseTask.getAttribute("TASKINSTANCE")
                xmlElem.setAttribute "TASKSTATUS", "30"
        
                Set objTm = GetObjectContext.CreateInstance("msgTm.msgTmBO")
                objTm.TmRequest xmlRequestDoc.xml
                Set objTm = Nothing
            
            End If
        
        End If
        
    End If
    
    Set objTm = Nothing
    
    Set xmlCaseActivityDoc = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseTask = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing

End Function

Private Sub CreateTask(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Const cstrMethodName As String = "CreateTask"
    On Error GoTo CreateTaskExit
    
    Dim xmlCaseActivityDoc As DOMDocument40
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlApplicationPriority As IXMLDOMElement
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objTm As Object
    
    Dim strIngestionTaskRef As String
    Dim strTaskId As String

    strIngestionTaskRef = vxmlRequestNode.Attributes.getNamedItem("PostProcTaskId").Text
    
    strTaskId = GetTaskId(strIngestionTaskRef)
    
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlApplicationPriority = xmlCaseActivityDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY")
    Set xmlCaseActivity = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivity.selectSingleNode("CASESTAGE")
    
    If Not OutstandingTaskExists(xmlCaseStage, strTaskId) Then
    
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
        xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseActivity.getAttribute("SOURCEAPPLICATION")
        xmlElem.setAttribute "CASEID", xmlCaseActivity.getAttribute("CASEID")
        xmlElem.setAttribute "ACTIVITYID", xmlCaseActivity.getAttribute("ACTIVITYID")
        xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseActivity.getAttribute("ACTIVITYINSTANCE")
        xmlElem.setAttribute "STAGEID", xmlCaseStage.getAttribute("STAGEID")
        xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
        xmlElem.setAttribute "TASKID", strTaskId
        xmlNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
        xmlElem.setAttribute "APPLICATIONPRIORITY", xmlApplicationPriority.getAttribute("APPLICATIONPRIORITYVALUE")
        xmlNode.appendChild xmlElem
    
        Set xmlResponseDoc = New DOMDocument40
        xmlResponseDoc.async = False
        xmlResponseDoc.setProperty "NewParser", True
    
        Set objTm = GetObjectContext.CreateInstance("omTm.omTmBO")
        xmlResponseDoc.loadXML objTm.OmTmRequest(xmlRequestDoc.xml)
        Set objTm = Nothing
        
        If strIngestionTaskRef <> "GeneralFailure" Then
            
            If xmlResponseDoc.parseError.errorCode <> 0 Then
                xmlParseError xmlResponseDoc.parseError
            End If
            
            If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
                Err.Raise oeUnspecifiedError, "OMIGAERROR", xmlResponseDoc.xml
            End If
        
        End If
    
    End If
    
CreateTaskExit:
        
    Set objTm = Nothing
    
    Set xmlCaseActivityDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlApplicationPriority = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    If strIngestionTaskRef <> "GeneralFailure" Then
        CheckError cstrMethodName
    End If

End Sub

'IK_09/03/2006_EP2

Private Function EpsomPostProcRules(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "EpsomPostProcRules"

    On Error GoTo EpsomPostProcRulesExit
    
    ' replaced as required
    EpsomPostProcRules = "<RESPONSE TYPE='SUCCESS'/>"
    
    If Not vxmlRequestNode.Attributes.getNamedItem("UPDATEABORT") Is Nothing Then
        CreateTask vxmlRequestNode
    Else
        EpsomPostProcRules = EpsomDecision(vxmlRequestNode)
    End If
    
EpsomPostProcRulesExit:
    
    CheckError cstrMethodName

End Function

Private Function EpsomDecision(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "EpsomDecision"
    On Error GoTo EpsomDecisionExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim xmlReturnDoc As DOMDocument40
    Dim xmlReturnElem As IXMLDOMElement
    Dim xmlReturnNode As IXMLDOMElement
    
    Dim xmlCaseActivityDoc As DOMDocument40
    
    Dim strOperation As String
    Dim blnDecisionProcessFailed As Boolean
    
    'IK_EP304_25/03/2006
    Dim strCustName As String
    Dim xmlCustNode As IXMLDOMElement
    
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    strOperation = vxmlRequestNode.getAttribute("PostProcTaskId")
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML EpsomRunTask(vxmlRequestNode, xmlCaseActivityDoc, "RunXMLCreditCheck")
    
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        blnDecisionProcessFailed = True
    End If
    
    For Each xmlElem In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE")
    
        'IK_EP304_25/03/2006
        strCustName = ""
        
        Set xmlCustNode = xmlElem.selectSingleNode("CUSTOMER/CUSTOMERVERSION")
        If Not xmlCustNode Is Nothing Then
            strCustName = _
                xmlCustNode.getAttribute("FIRSTFORENAME") & " " & xmlCustNode.getAttribute("SURNAME")
        End If
        
        xmlDoc.loadXML _
            EpsomRunTask( _
                vxmlRequestNode, _
                xmlCaseActivityDoc, _
                "RunKnowYourCustomer", _
                xmlElem.getAttribute("CUSTOMERNUMBER"), _
                strCustName)
        'IK_EP304_25/03/2006_ends
        
        If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            blnDecisionProcessFailed = True
        End If
            
    Next
    
    xmlDoc.loadXML EpsomRunTask(vxmlRequestNode, xmlCaseActivityDoc, "RunRiskAssess")
    
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        blnDecisionProcessFailed = True
    End If
    
    If blnDecisionProcessFailed Then
        EpsomDecision = xmlDoc.xml
    Else
        'IK_EP304_28/03/2006
        'CompleteCreditCheckTask vxmlRequestNode, strOperation
        EpsomDecision = EpsomResponse(vxmlRequestNode)
    End If
        
EpsomDecisionExit:
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    'IK_EP304_25/03/2006
    Set xmlCustNode = Nothing
    
    CheckError cstrMethodName
    
End Function

Private Function EpsomResponse(ByVal vxmlRequestNode As IXMLDOMElement) As String

    Const cstrMethodName As String = "EpsomResponse"
    On Error GoTo EpsomResponseExit
    
    Dim xmlReturnDoc As DOMDocument40
    Dim xmlReturnElem As IXMLDOMElement
    Dim xmlReturnNode As IXMLDOMElement
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim objCRUD As Object
    
    Dim strApplicationNumber As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
        
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "INGESTIONDECISION"
    xmlElem.setAttribute "SCHEMA_NAME", "WebServiceSchema"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("INGESTIONDECISION")
    xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlNode.appendChild xmlElem
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    EpsomResponse = objCRUD.OmRequest(xmlDoc.xml)
        
EpsomResponseExit:
    
    Set objCRUD = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function
'IK_EP304_25/03/2006
Private Function EpsomRunTask( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlCaseActivityDoc As DOMDocument40, _
    ByVal vstrTaskRef As String, _
    Optional ByVal vstrCustId As String = "", _
    Optional ByVal vstrCustName As String = "") _
    As String
    
    Const cstrMethodName As String = "EpsomRunTask"
    On Error GoTo EpsomRunTaskExit
    
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlApplicationPriority As IXMLDOMElement
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objTm As Object
    
    Dim strTaskId As String
    
    strTaskId = GetTaskId(vstrTaskRef)
    
    Set xmlApplicationPriority = vxmlCaseActivityDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY")
    Set xmlCaseActivity = vxmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivity.selectSingleNode("CASESTAGE")
    
    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.async = False
    xmlRequestDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
    xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
    xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
    xmlElem.setAttribute "OPERATION", "CreateAdhocCaseTask"
    If vstrTaskRef = "RunXMLCreditCheck" Then
        xmlElem.setAttribute "NOADDRESSTARGETING", "1"
    End If
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlRequestDoc.createElement("CASETASK")
    xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseActivity.getAttribute("SOURCEAPPLICATION")
    xmlElem.setAttribute "CASEID", xmlCaseActivity.getAttribute("CASEID")
    xmlElem.setAttribute "ACTIVITYID", xmlCaseActivity.getAttribute("ACTIVITYID")
    xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseActivity.getAttribute("ACTIVITYINSTANCE")
    xmlElem.setAttribute "STAGEID", xmlCaseStage.getAttribute("STAGEID")
    xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
    xmlElem.setAttribute "TASKID", strTaskId
    If Len(vstrCustId) <> 0 Then
        xmlElem.setAttribute "CUSTOMERIDENTIFIER", vstrCustId
    End If
    'IK_EP304_25/03/2006
    If Len(vstrCustName) <> 0 Then
        xmlElem.setAttribute "CASETASKNAME", GetTaskname(vstrTaskRef) & " for " & vstrCustName
    End If
    'IK_EP304_25/03/2006_ends
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONPRIORITY", xmlApplicationPriority.getAttribute("APPLICATIONPRIORITYVALUE")
    xmlNode.appendChild xmlElem

    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True

    Set objTm = GetObjectContext.CreateInstance("omTm.omTmBO")
    EpsomRunTask = objTm.OmTmRequest(xmlRequestDoc.xml)
    
EpsomRunTaskExit:
        
    Set objTm = Nothing
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlApplicationPriority = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    CheckError cstrMethodName

End Function

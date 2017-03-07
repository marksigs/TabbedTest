Attribute VB_Name = "PostProcRules"
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     06/10/2006  EP2_10 - from omIngestionRules.PreProcBO.cls
'IK     18/11/2006  EP2_134 no credit check required for epsom
'IK     18/11/2006  EP2_134 EpsomDecision for SubmitAip only
'IK     25/11/2006  EP2_202 Product Cascade enhancements
'IK     05/12/2006  EP2_287 'Shopping List' enhancements
'IK     02/01/2007  EP2_581 INVALIDQUOTEIND set if product cascade
'IK     04/01/2007  EP2_494 move validation to IngestionManagerBO
'IK     12/01/2007  EP2_791 create quote (as required) before credit check
'IK     13/01/2007  EP2_757 KYC for customers only
'IK     15/01/2007  EP2_861 fix error handling
'IK     16/01/2007  EP2_757 KYC for customers only, test global parameter ExperianApplicantsOnly
'IK     22/01/2007  EP2_920 KYC after credit check, simplify error messages
'IK     15/01/2007  EP2_983 re-fix error handling
'IK     26/01/2007  EP2_461 return address targeting response for ambiguous (AUK1) addresses only
'IK     26/01/2007  EP2_797 call RunIncomeCalcs before case assessment
'IK     29/01/2007  EP2_920 KYC after credit check, no customer data on request
'IK     19/02/2007  EP2_1172, use input TYPEOFVALUATION if present for CreateQuote
'IK     21/03/2007  EP2_1062, create & return application form after SubmitFMA
'IK     04/04/2007  EP2_1041, refine Shopping List test
'-----------------------------------------------------------------------------
Option Explicit

Public Function PostProcRequest(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "PostProcRequest"
    On Error GoTo PostProcRequestVbErr
    
    ' default response - will be overwritten as required
    PostProcRequest = "<RESPONSE TYPE='SUCCESS'/>"
    
    ' EP2_10
    If Not IsAddressTargetResolver(vxmlRequestNode) Then
    
        ' If a Create CRUD_Op Exists, Create Case and Move to next Stage
        If Not vxmlRequestNode.selectSingleNode("APPLICATION[@CRUD_OP='CREATE']") Is Nothing Then
            CreateCaseActivity vxmlRequestNode '----->>
            MoveToStage vxmlRequestNode, "PreAiP" '----->>
        End If
        
        ' If an ACCEPTQUOTE attribute exists, do Accept Quote routine
        If Not vxmlRequestNode.Attributes.getNamedItem("ACCEPTQUOTE") Is Nothing Then
            AcceptQuote vxmlRequestNode '----->>
        End If
    
        ' What Operation was specified in the OPERATION Attribute
        Select Case vxmlRequestNode.getAttribute("OPERATION")
    
            Case "SubmitAiP"
                MoveToStage vxmlRequestNode, "SubmitAiP" '----->>
    
            Case "SubmitFMA"
                MoveToStage vxmlRequestNode, "SubmitFMA" '----->>
    
        End Select

    End If

    ' Do Post Process Rules and set Return Value
    'IK_17/03/2006_EP2
    Select Case vxmlRequestNode.getAttribute("omigaClient")
    
        Case "epsom"

            PostProcRequest = EpsomPostProcRules(vxmlRequestNode) '----->>
            
        Case Else

            PostProcRequest = PostProcRules(vxmlRequestNode) '----->>
    
    End Select
    'IK_09/03/2006_EP2_ends
    
PostProcRequestVbErr:
    
    If Err.Number <> 0 Then
        If Err.Source = "OMIGAERROR" Then
            PostProcRequest = Err.Description
        Else
            If Err.Source <> cstrMethodName Then
                If Err.Source = App.EXEName Then
                    Err.Source = cstrMethodName
                Else
                    Err.Source = cstrMethodName & "." & Err.Source
                End If
            End If
            Err.Source = App.EXEName & ".PostProcRules." & Err.Source
            PostProcRequest = FormatError(vxmlRequestNode.xml, gstrComponentId, gstrComponentResponse)
        End If
    End If

End Function

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
        
        'EP2_920
        LogOmigaError strComponentResponse
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
        
        'EP2_920
        LogOmigaError strComponentResponse
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
                
                'EP2_920
                LogOmigaError strComponentResponse
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
            'EP2_920
            LogOmigaError strComponentResponse
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
    
    Dim strIngestionTaskRef As String

    strIngestionTaskRef = vxmlRequestNode.Attributes.getNamedItem("PostProcTaskId").Text
    CreateTaskByTaskRef vxmlRequestNode, strIngestionTaskRef
    
CreateTaskExit:
    
    If strIngestionTaskRef <> "GeneralFailure" Then
        CheckError cstrMethodName
    End If

End Sub

' EP2_10
Private Sub CreateTaskByTaskRef(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vstrIngestionTaskRef)
    
    Const cstrMethodName As String = "CreateTaskByTaskRef"
    On Error GoTo CreateTaskByTaskRefExit
    
    Dim xmlCaseActivityDoc As DOMDocument40
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlApplicationPriority As IXMLDOMElement
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objTm As Object
    
    Dim strTaskId As String
    
    strTaskId = GetTaskId(vstrIngestionTaskRef)
    
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
        
        If vstrIngestionTaskRef <> "GeneralFailure" Then
            
            If xmlResponseDoc.parseError.errorCode <> 0 Then
                xmlParseError xmlResponseDoc.parseError
            End If
            
            If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
                'EP2_920
                LogOmigaError xmlResponseDoc.xml
                Err.Raise oeUnspecifiedError, cstrMethodName, "error creating task"
            End If
        
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
    Set xmlResponseDoc = Nothing
    
    If vstrIngestionTaskRef <> "GeneralFailure" Then
        CheckError cstrMethodName
    End If


End Sub

Private Function EpsomPostProcRules(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "EpsomPostProcRules"

    On Error GoTo EpsomPostProcRulesExit
    
    ' replaced as required
    EpsomPostProcRules = "<RESPONSE TYPE='SUCCESS'/>"
    
    If Not vxmlRequestNode.Attributes.getNamedItem("UPDATEABORT") Is Nothing Then
        CreateTask vxmlRequestNode
    Else
        ' EP2_134
        Select Case vxmlRequestNode.getAttribute("OPERATION")
            Case "SubmitAiP"
                EpsomPostProcRules = EpsomDecision(vxmlRequestNode)
            ' EP2_1062
            Case "SubmitFMA"
                If GetGlobalString("WebApplicationFormDocumentId", "") <> "" Then
                    EpsomPostProcRules = EpsomGetApplicationForm(vxmlRequestNode)
                End If
        End Select
        ' EP2_134_ends
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
    
    Set xmlCaseActivityDoc = New DOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    strOperation = vxmlRequestNode.getAttribute("PostProcTaskId")
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    ' EP2_10
    If Not IsAddressTargetResolver(vxmlRequestNode) Then
    
        ' EP2_791
        'if product code(s) input, create quote before Risk Assessment
        If vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGEPRODUCTCODE]").length <> 0 Then
            
            xmlDoc.loadXML CreateQuote(vxmlRequestNode)
    
            If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
                EpsomDecision = xmlDoc.xml
                'EP2_920
                LogOmigaError EpsomDecision
                Err.Raise oeUnspecifiedError, cstrMethodName, "error creating quotation details"
            End If
        
        End If

        'EP2_797
        RunIncomeCalcs vxmlRequestNode
        
        ' EP2_10 task Ingest_CreditCheck no longer Automatic on create
        CreateTaskByTaskRef vxmlRequestNode, "RunXMLCreditCheck"
        
    End If
    
    xmlDoc.loadXML RunEpsomCreditCheck(vxmlRequestNode)

    ' 'EP2_461 test for address targeting response
    If IsAmbiguousAddressTargetResponse(xmlDoc) Then
    
        CompleteAddressTargetingResponse vxmlRequestNode, xmlDoc
        
        EpsomDecision = xmlDoc.xml
    
    Else
    
        'EP2_920
        RunKYC vxmlRequestNode, xmlCaseActivityDoc
        
        xmlDoc.loadXML EpsomRunTask(vxmlRequestNode, xmlCaseActivityDoc, "RunRiskAssess")
    
        If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            EpsomDecision = xmlDoc.xml
            'EP2_920
            LogOmigaError EpsomDecision
            Err.Raise oeUnspecifiedError, cstrMethodName, "error runnning risk assessment process"
        Else
            
            xmlDoc.loadXML EpsomResponse(vxmlRequestNode)
            
            'EP2_287
            AddShoppingList vxmlRequestNode, xmlDoc
            
            If IsProductCascadeDecision(xmlDoc) Then
                'EP2_581
                SetInvalidQuoteFlag vxmlRequestNode
                AddProductsToDecison vxmlRequestNode, xmlDoc
            End If
            
            EpsomDecision = xmlDoc.xml
            
            ' EP2_202_ends
        
        End If
    
    End If
        
EpsomDecisionExit:
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName
    
End Function

' EP2_202 EpsomDecision restructuring
Private Sub RunKYC( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlCaseActivityDoc As DOMDocument40)

    Const cstrMethodName As String = "RunKYC"
    On Error GoTo RunKYCExit
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlCustNode As IXMLDOMElement
    Dim xmlCustRoleList As IXMLDOMNodeList
    
    Dim strCustName As String
    
    ' EP2_920 29/01/2007, get customer data
    Dim xmlDoc As DOMDocument40
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    If IsAddressTargetResolver(vxmlRequestNode) Then
    
        ' EP2_920 29/01/2007, get customer data
        GetCustomerDataForKYC vxmlRequestNode, xmlDoc
        
        If GetGlobalBool("ExperianApplicantsOnly", True) = True Then
            Set xmlCustRoleList = xmlDoc.selectNodes("RESPONSE/CUSTOMERROLE[@CUSTOMERROLETYPE='1']")
        Else
            Set xmlCustRoleList = xmlDoc.selectNodes("RESPONSE/CUSTOMERROLE")
        End If
        
    Else
        If GetGlobalBool("ExperianApplicantsOnly", True) = True Then
            Set xmlCustRoleList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE='1']")
        Else
            Set xmlCustRoleList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE")
        End If
        
    End If
    
    For Each xmlElem In xmlCustRoleList
    
        'IK_EP304_25/03/2006
        strCustName = ""
        
        Set xmlCustNode = xmlElem.selectSingleNode("CUSTOMER/CUSTOMERVERSION")
        If Not xmlCustNode Is Nothing Then
            strCustName = _
                xmlCustNode.getAttribute("FIRSTFORENAME") & " " & xmlCustNode.getAttribute("SURNAME")
        End If
        
        EpsomRunTask _
            vxmlRequestNode, _
            vxmlCaseActivityDoc, _
            "RunKnowYourCustomer", _
            xmlElem.getAttribute("CUSTOMERNUMBER"), _
            strCustName
        'IK_EP304_25/03/2006_ends
            
    Next

RunKYCExit:

    Set xmlCustRoleList = Nothing
    Set xmlCustNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub CompleteAddressTargetingResponse( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlDoc As DOMDocument40)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
        
    Set xmlElem = vxmlDoc.createElement("APPLICATION")
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    Set xmlNode = vxmlDoc.documentElement.appendChild(xmlElem)

    Set xmlElem = vxmlDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute _
        "APPLICATIONFACTFINDNUMBER", _
        vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem

    Set xmlElem = Nothing
    Set xmlNode = Nothing

End Sub

'EP2_202 product cascade
Private Function CreateQuote(ByVal vxmlRequestNode As IXMLDOMElement) As String

    Const cstrMethodName As String = "CreateQuote"
    On Error GoTo CreateQuoteExit
    
    Dim xmlDoc As DOMDocument40
    
    Dim objCK As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    BuildCreateQuoteRequest vxmlRequestNode, xmlDoc

    Set objCK = GetObjectContext.CreateInstance("omCK.CreateKFIBO")
    
    CreateQuote = objCK.CalculateKFIFees(xmlDoc.xml)
    
    Set objCK = Nothing

CreateQuoteExit:
    
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Sub BuildCreateQuoteRequest( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlRequestDoc As DOMDocument40)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlReqAttrib As IXMLDOMAttribute
    Dim xmlCustListNode As IXMLDOMElement
    
    Dim strOtherForenames As String
    
    Set xmlElem = vxmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    For Each xmlReqAttrib In xmlElem.Attributes
    
        Select Case xmlReqAttrib.nodeName
        
            Case "USERID", "UNITID", "USERAUTHORITYLEVEL", "PASSWORD", "CHANNELID"
                ' do nothing
                
            Case Else
                xmlElem.removeAttribute xmlReqAttrib.Name
        
        End Select
    
    Next
    
    If xmlElem.Attributes.getNamedItem("USERAUTHORITYLEVEL") Is Nothing Then
        xmlElem.setAttribute "USERAUTHORITYLEVEL", GetWebChannelId
    End If
    xmlElem.setAttribute "NODOC", "true"
    
    Set xmlElem = xmlElem.appendChild(vxmlRequestNode.selectSingleNode("APPLICATION").cloneNode(False))
    xmlElem.setAttribute "DISPOSABLEKFI", "0"
    
    ' EP2_1172, use input TYPEOFVALUATION if present
    If xmlElem.Attributes.getNamedItem("TYPEOFVALUATION") Is Nothing Then
        xmlElem.setAttribute "TYPEOFVALUATION", GetDefaultValuationType
    End If
    
    xmlElem.setAttribute "APPLICATIONDATE", FormatNow
    
    Set xmlElem = xmlElem.appendChild(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND").cloneNode(False))
    
    If xmlElem.Attributes.getNamedItem("LOCATION") Is Nothing Then
        xmlElem.setAttribute "LOCATION", GetDefaultPropertyLocation
    End If
    
    If xmlElem.Attributes.getNamedItem("PURCHASEPRICE") Is Nothing Then
        If Not xmlElem.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            xmlElem.setAttribute "PURCHASEPRICE", xmlElem.getAttribute("PURCHASEPRICEORESTIMATEDVALUE")
        End If
    End If
    
    Set xmlElem = xmlElem.appendChild(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION").cloneNode(False))
    Set xmlElem = xmlElem.appendChild(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE").cloneNode(False))
    Set xmlElem = xmlElem.appendChild(vxmlRequestDoc.createElement("LOANCOMPONENTLIST"))
    For Each xmlNode In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT")
        xmlElem.appendChild xmlNode.cloneNode(True)
    Next
    
    Set xmlCustListNode = vxmlRequestDoc.documentElement.appendChild(vxmlRequestDoc.createElement("CUSTOMERLIST"))
    
    For Each xmlNode In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE='" & GetApplicantCustomerRoleType & "']/CUSTOMER/CUSTOMERVERSION")
    
        Set xmlElem = xmlCustListNode.appendChild(vxmlRequestDoc.createElement("CUSTOMER"))
        
        If Not xmlNode.Attributes.getNamedItem("SURNAME") Is Nothing Then
            xmlElem.setAttribute "SURNAME", xmlNode.Attributes.getNamedItem("SURNAME").Text
        End If
        
        If Not xmlNode.Attributes.getNamedItem("FIRSTFORENAME") Is Nothing Then
            xmlElem.setAttribute "FIRSTNAME", xmlNode.Attributes.getNamedItem("FIRSTFORENAME").Text
        End If
        
        If Not xmlNode.Attributes.getNamedItem("SECONDFORENAME") Is Nothing Then
                        
            strOtherForenames = xmlNode.Attributes.getNamedItem("SECONDFORENAME").Text
        
            If Not xmlNode.Attributes.getNamedItem("OTHERFORENAMES") Is Nothing Then
                strOtherForenames = strOtherForenames & " " & xmlNode.Attributes.getNamedItem("OTHERFORENAMES").Text
            End If
            
            xmlElem.setAttribute "OTHERNAMES", strOtherForenames
        
        End If
        
        If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
            xmlElem.setAttribute "DATEOFBIRTH", xmlNode.Attributes.getNamedItem("DATEOFBIRTH").Text
        End If
        
        If Not xmlNode.Attributes.getNamedItem("TITLE") Is Nothing Then
            xmlElem.setAttribute "TITLE", xmlNode.Attributes.getNamedItem("TITLE").Text
        End If
        
        If Not xmlNode.Attributes.getNamedItem("GENDER") Is Nothing Then
            xmlElem.setAttribute "GENDER", GetGenderValidationType(xmlNode.Attributes.getNamedItem("GENDER").Text)
        End If
    
    Next
    
    Set xmlCustListNode = Nothing
    Set xmlReqAttrib = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    
End Sub

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

Private Function RunEpsomCreditCheck(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "RunEpsomCreditCheck"
    On Error GoTo RunEpsomCreditCheckExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim strTaskRef As String, _
        strOperation As String, _
        strComponentResponse As String, _
        strDecision As String
    
    Dim lngErr As Long  'MAR1084 GHun
    
    strTaskRef = vxmlRequestNode.getAttribute("PostProcTaskId")
    strOperation = GetTaskInterface(strTaskRef)
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    Set xmlNode = xmlDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    xmlNode.setAttribute "OPERATION", strOperation
'    xmlNode.setAttribute "NOADDRESSTARGETING", "1" 'MAR1084 GHun
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    
    ' EP2_10
    xmlNode.appendChild xmlElem
    If IsAddressTargetResolver(vxmlRequestNode) Then
        xmlNode.appendChild vxmlRequestNode.selectSingleNode("TARGETINGDATA").cloneNode(True)
    End If
    
    Dim objTm As Object
    Set objTm = GetObjectContext.CreateInstance("omTm.omTmBO")
    strComponentResponse = objTm.OmTmRequest(xmlDoc.xml)
    Set objTm = Nothing
    
    xmlDoc.loadXML strComponentResponse
    If Not xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
    
        ' EP2_10 test for address targetting response
        If xmlDoc.selectSingleNode("RESPONSE/TARGETINGDATA[ADDRESSTARGETING='YES']") Is Nothing Then
        
            xmlDoc.loadXML ""
            
            Set xmlElem = xmlDoc.createElement("RESPONSE")
            xmlElem.setAttribute "TYPE", "SUCCESS"
            Set xmlNode = xmlDoc.appendChild(xmlElem)
            
            Set xmlElem = xmlDoc.createElement("APPLICATION")
            xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
            xmlNode.appendChild xmlElem
            
            If Not CompleteCreditCheckTask(vxmlRequestNode, strTaskRef) Then
                
                Set xmlElem = xmlDoc.createElement("MESSAGE")
                Set xmlNode = xmlNode.appendChild(xmlElem)
                
                Set xmlElem = xmlDoc.createElement("MESSAGETEXT")
                xmlElem.Text = "unable to update task status for TASKID " & GetTaskId("RunXMLCreditCheck")
                xmlNode.appendChild xmlElem
                
                Set xmlElem = xmlDoc.createElement("MESSAGETYPE")
                xmlElem.Text = "WARNING"
                xmlNode.appendChild xmlElem
            
            End If
        
        ' else EP2_10 return address targetting response
        
        End If
    
    End If
    
    RunEpsomCreditCheck = xmlDoc.xml
    
RunEpsomCreditCheckExit:
    
    Set objTm = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

'EP2_202
Private Function IsProductCascadeDecision(ByVal vxmlDecisionDoc As DOMDocument40) As Boolean
    If Not vxmlDecisionDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/INGESTIONDECISION[@RISKASSESSMENTDECISION='" & GetProductCascadeDecision & "']") Is Nothing Then
        IsProductCascadeDecision = True
    End If
End Function

Private Sub AddProductsToDecison( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlDecisionDoc As DOMDocument40)

    Const cstrMethodName As String = "AddProductsToDecison"
    On Error GoTo AddProductsToDecisonExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim objAQ As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    BuildFindMortgageProductsRequest vxmlRequestNode, xmlDoc

    Set objAQ = GetObjectContext.CreateInstance("omAQ.ApplicationQuoteBO")
    xmlDoc.loadXML objAQ.FindMortgageProducts(xmlDoc.xml)
    Set objAQ = Nothing
    
    If Not xmlDoc.selectSingleNode("RESPONSE/RESPONSE/MORTGAGEPRODUCTLIST") Is Nothing Then
        vxmlDecisionDoc.documentElement.appendChild xmlDoc.selectSingleNode("RESPONSE/RESPONSE/MORTGAGEPRODUCTLIST").cloneNode(True)
    Else
            
        Set xmlElem = vxmlDecisionDoc.createElement("MORTGAGEPRODUCTLIST")
        xmlElem.setAttribute "TOTAL", "0"
        vxmlDecisionDoc.documentElement.appendChild xmlElem
    
    End If
    
    
AddProductsToDecisonExit:
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub BuildFindMortgageProductsRequest( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlRequestDoc As DOMDocument40)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlReqAttrib As IXMLDOMAttribute
    Dim xmlSource As IXMLDOMNode
    
    Set xmlElem = vxmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    For Each xmlReqAttrib In xmlElem.Attributes
    
        Select Case xmlReqAttrib.nodeName
        
            Case "USERID", "UNITID", "USERAUTHORITYLEVEL", "PASSWORD", "CHANNELID"
                ' do nothing
                
            Case Else
                xmlElem.removeAttribute xmlReqAttrib.Name
        
        End Select
    
    Next
    
    xmlElem.setAttribute "CHANNELID", GetWebChannelId
    
    Set xmlNode = xmlElem.appendChild(vxmlRequestDoc.createElement("MORTGAGEPRODUCTREQUEST"))
    
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "SEARCHCONTEXT", "Cost Modelling"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "DISTRIBUTIONCHANNELID", GetWebChannelId
    
    AddMortgageProductRequestElement _
        vxmlRequestDoc, _
        xmlNode, _
        "APPLICATIONNUMBER", _
        vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    
    AddMortgageProductRequestElement _
        vxmlRequestDoc, _
        xmlNode, _
        "APPLICATIONFACTFINDNUMBER", _
        vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
        
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "MORTGAGESUBQUOTENUMBER", "1"
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "TYPEOFAPPLICATION", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@TYPEOFBUYER")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "TYPEOFBUYER", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@PURPOSEOFLOAN")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "PURPOSEOFLOAN", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@TOTALLOANCOMPONENTAMOUNT")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "LOANCOMPONENTAMOUNT", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@TERMINYEARS")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "TERMINYEARS", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@TERMINMONTHS")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "TERMINMONTHS", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "AMOUNTREQUESTED", xmlSource.Text
    End If
    
    Set xmlSource = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@LTV")
    If Not xmlSource Is Nothing Then
        AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "LTV", xmlSource.Text
    End If
    
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "ALLPRODUCTSWITHOUTCHECKS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "ALLPRODUCTSWITHCHECKS", "1"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "PRODUCTSBYGROUP", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "DISCOUNTEDPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "FIXEDRATEPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "STANDARDVARIABLEPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "CAPPEDFLOOREDPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "FLEXIBLENONFLEXIBLEPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "CASHBACKPRODUCTS", "0"
    AddMortgageProductRequestElement vxmlRequestDoc, xmlNode, "IMPAIREDCREDITRATING", "0"
    
    Set xmlSource = Nothing
    Set xmlReqAttrib = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing

End Sub

Private Sub AddMortgageProductRequestElement( _
    ByVal vxmlTargetDoc As DOMDocument40, _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrElementName As String, _
    ByVal vstrElementText As String)
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = vxmlTargetDoc.createElement(vstrElementName)
    xmlElem.Text = vstrElementText
    vxmlParentNode.appendChild xmlElem
    
    Set xmlElem = Nothing

End Sub

Private Function FormatNow() As String

    Dim dtNow As Date
    dtNow = Now()
    
    FormatNow = CStr(Day(dtNow)) & "/" & CStr(Month(dtNow)) & "/" & CStr(Year(dtNow)) & _
                " " & CStr(Hour(dtNow)) & ":" & CStr(Minute(dtNow)) & ":" & CStr(Second(dtNow))

End Function
'EP2_202_ends

'EP2_287
Private Sub AddShoppingList( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlResponseDoc As DOMDocument40)

    Dim strNatureOfLoan As String
    Dim strEmploymentStatus As String
    Dim strIncomeStatus As String
    Dim strPackaged As String
    Dim strValidationType As String
    Dim strProfile As String
    
    Dim blnAddShoppingList As Boolean
    blnAddShoppingList = True
    
    Dim xmlNode As IXMLDOMElement
    
    Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND[@NATUREOFLOAN]")
    If xmlNode Is Nothing Then
        blnAddShoppingList = False
    Else
        strNatureOfLoan = xmlNode.getAttribute("NATUREOFLOAN")
        strNatureOfLoan = GetNatureOfLoanType(strNatureOfLoan)
        If Len(strNatureOfLoan) = 0 Then
            blnAddShoppingList = False
        End If
    End If
    
    If blnAddShoppingList Then
        Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND[@APPLICATIONINCOMESTATUS]")
        If xmlNode Is Nothing Then
            blnAddShoppingList = False
        Else
            strIncomeStatus = xmlNode.getAttribute("APPLICATIONINCOMESTATUS")
            strIncomeStatus = GetIncomeStatusType(strIncomeStatus)
            If Len(strIncomeStatus) = 0 Then
                blnAddShoppingList = False
            End If
        End If
    End If
    
    If blnAddShoppingList Then
        ' EP2_1041 belt & braces test
        Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE='" & GetApplicantCustomerRoleType & "'][@CUSTOMERORDER='1']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='true'][@EMPLOYMENTSTATUS]")
        If xmlNode Is Nothing Then
            Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE='" & GetApplicantCustomerRoleType & "'][@CUSTOMERORDER='1']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS]")
        End If
        If xmlNode Is Nothing Then
            blnAddShoppingList = False
        Else
            strEmploymentStatus = xmlNode.getAttribute("EMPLOYMENTSTATUS")
            strEmploymentStatus = GetEmploymentStatusType(strEmploymentStatus)
            If Len(strEmploymentStatus) = 0 Then
                blnAddShoppingList = False
            End If
        End If
    End If
        
    If blnAddShoppingList Then
        
        strValidationType = _
            strNatureOfLoan & "$" & _
            strEmploymentStatus & "$" & _
            strIncomeStatus & "$"

        If IsPackaged(vxmlRequestNode) Then
            strValidationType = strValidationType & "P"
        Else
            strValidationType = strValidationType & "U"
        End If
        
        Set xmlNode = _
            gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/COMBOGROUP[@GROUPNAME='ShoppingListProfile']/COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" & strValidationType & "']]")
        
        If Not xmlNode Is Nothing Then
            strProfile = xmlNode.Attributes.getNamedItem("VALUENAME").Text
            Set xmlNode = vxmlResponseDoc.createElement("SHOPPINGLISTPROFILE")
            xmlNode.Text = strProfile
            vxmlResponseDoc.documentElement.appendChild xmlNode
        End If
        
    End If
    
    Set xmlNode = Nothing

End Sub

Private Function IsPackaged(ByVal vxmlRequestNode As IXMLDOMElement) As Boolean
    
    Dim xmlDoc As DOMDocument40
    Dim objCRUD As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML _
        "<REQUEST CRUD_OP='READ'><APPLICATIONINTRODUCER APPLICATIONNUMBER='" & _
        vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text & _
        "' APPLICATIONFACTFINDNUMBER='1'><PRINCIPALFIRM/></APPLICATIONINTRODUCER></REQUEST>"
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc.loadXML objCRUD.OmRequest(xmlDoc.xml)
    Set objCRUD = Nothing
    
    If Not xmlDoc.selectSingleNode("RESPONSE/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='1']") Is Nothing Then
        IsPackaged = True
    End If
    
    Set xmlDoc = Nothing

End Function

Private Function GetNatureOfLoanType(ByVal vstrNatureOfLoan As String) As String
    If IsComboValidationType("NatureOfLoan", vstrNatureOfLoan, "RS") Then
        GetNatureOfLoanType = "RS"
    ElseIf IsComboValidationType("NatureOfLoan", vstrNatureOfLoan, "LT") Then
        GetNatureOfLoanType = "LT"
    ElseIf IsComboValidationType("NatureOfLoan", vstrNatureOfLoan, "BI") Then
        GetNatureOfLoanType = "BI"
    ElseIf IsComboValidationType("NatureOfLoan", vstrNatureOfLoan, "BR") Then
        GetNatureOfLoanType = "BR"
    End If
End Function

Private Function GetEmploymentStatusType(ByVal vstrEmploymentStatus As String) As String
    If IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "EMP") Then
        GetEmploymentStatusType = "EMP"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "EMPT") Then
        GetEmploymentStatusType = "EMPT"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "HM") Then
        GetEmploymentStatusType = "HM"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "STU") Then
        GetEmploymentStatusType = "STU"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "R") Then
        GetEmploymentStatusType = "R"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "SELF") Then
        GetEmploymentStatusType = "SELF"
    ElseIf IsComboValidationType("EmploymentStatus", vstrEmploymentStatus, "CON") Then
        GetEmploymentStatusType = "CON"
    End If
End Function

Private Function GetIncomeStatusType(ByVal vstrIncomeStatus As String) As String
    If IsComboValidationType("ApplicationIncomeStatus", vstrIncomeStatus, "FS") Then
        GetIncomeStatusType = "FS"
    ElseIf IsComboValidationType("ApplicationIncomeStatus", vstrIncomeStatus, "SC") Then
        GetIncomeStatusType = "SC"
    End If
End Function
'EP2_287_ends

'EP2_581
Private Sub SetInvalidQuoteFlag(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Dim xmlDoc As DOMDocument40
    Dim xmlNode As IXMLDOMElement
    Dim objCRUD As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlNode = xmlDoc.createElement("REQUEST")
    xmlNode.setAttribute "CRUD_OP", "UPDATE"
    xmlDoc.appendChild xmlNode
    
    Set xmlNode = xmlDoc.createElement("QUOTATION")
    xmlNode.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    xmlNode.setAttribute "QUOTATIONNUMBER", "1"
    xmlNode.setAttribute "INVALIDQUOTEIND", "1"
    xmlDoc.documentElement.appendChild xmlNode
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc.loadXML objCRUD.OmRequest(xmlDoc.xml)
    Set objCRUD = Nothing
    
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

End Sub

'EP2_581_ends

'EP2_461
Private Function IsAmbiguousAddressTargetResponse(ByVal vxmlExperianResponse As DOMDocument40) As Boolean
    If Not vxmlExperianResponse.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/TARGETINGDATA[ADDRESSTARGETING='YES'][ADDRESSTARGET[@BLOCKTYPE='AUK1']]") Is Nothing Then
        IsAmbiguousAddressTargetResponse = True
    End If
End Function
'EP2_461_ends

'EP2_797
Private Sub RunIncomeCalcs(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustListNode As IXMLDOMNode
    Dim xmlCustRoleNode As IXMLDOMNode
    
    Dim objIC As Object
    
    Dim strResponse As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("INCOMECALCULATION")
    xmlElem.setAttribute "CALCULATEMAXBORROWING", "1"
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = "1"
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERLIST")
    Set xmlCustListNode = xmlDoc.documentElement.firstChild.appendChild(xmlElem)
    
    For Each xmlCustRoleNode In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE")
    
        Set xmlElem = xmlDoc.createElement("CUSTOMER")
        Set xmlNode = xmlCustListNode.appendChild(xmlElem)
    
        Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
        xmlElem.Text = xmlCustRoleNode.Attributes.getNamedItem("CUSTOMERNUMBER").Text
        xmlNode.appendChild xmlElem
    
        Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlElem.Text = "1"
        xmlNode.appendChild xmlElem
    
        Set xmlElem = xmlDoc.createElement("CUSTOMERROLETYPE")
        xmlElem.Text = xmlCustRoleNode.Attributes.getNamedItem("CUSTOMERROLETYPE").Text
        xmlNode.appendChild xmlElem
    
        Set xmlElem = xmlDoc.createElement("CUSTOMERORDER")
        xmlElem.Text = xmlCustRoleNode.Attributes.getNamedItem("CUSTOMERORDER").Text
        xmlNode.appendChild xmlElem
    
    Next
    
    Set objIC = GetObjectContext.CreateInstance("omIC.IncomeCalcsBO")
    strResponse = objIC.RunIncomeCalculation(xmlDoc.xml)
    Set objIC = Nothing
    
    Set xmlCustRoleNode = Nothing
    Set xmlCustListNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing

End Sub
'EP2_797_ends

' EP2_920 29/01/2007
Private Sub GetCustomerDataForKYC( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlDoc As DOMDocument40)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlElem = vxmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = vxmlDoc.appendChild(xmlElem)
    Set xmlElem = vxmlDoc.createElement("CUSTOMERROLE")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    xmlNode.appendChild xmlElem
    
    vxmlDoc.loadXML CrudBOCall(vxmlDoc.xml)
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing

End Sub
' EP2_920 29/01/2007

' EP2_1062
Private Function EpsomGetApplicationForm(ByVal vxmlRequestNode As IXMLDOMElement) As String
    
    Const cstrMethodName As String = "EpsomGetApplicationForm"
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objPrint As Object
    
    ' will raise error if fails
    Call LoadAppFormHostTemplateData
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    
    If Not vxmlRequestNode.Attributes.getNamedItem("USERID") Is Nothing Then
        xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
    End If
    
    If Not vxmlRequestNode.Attributes.getNamedItem("UNITID") Is Nothing Then
        xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
    End If
    
    If Not vxmlRequestNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
        xmlElem.setAttribute "CHANNELID", vxmlRequestNode.getAttribute("CHANNELID")
    End If
    
    If Not vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL") Is Nothing Then
        xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
    End If
    
    xmlElem.setAttribute "OPERATION", "PrintDocument"
    xmlElem.setAttribute "ACTION", "PrintDocument"
    xmlElem.setAttribute "PRINTINDICATOR", "1"
    
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    ' PRINTDATA node
    Set xmlElem = xmlDoc.createElement("PRINTDATA")
    
    If Not vxmlRequestNode.selectSingleNode("APPLICATION[@APPLICATIONNUMBER]") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    End If
    
    If Not vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND[@APPLICATIONFACTFINDNUMBER]") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlElem.setAttribute "METHODNAME", gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE/@PDMMETHOD").Text
    
    xmlNode.appendChild xmlElem
    
    ' TEMPLATEDATA node
    Set xmlElem = xmlDoc.createElement("TEMPLATEDATA")
    
    If Not vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@QUOTATIONNUMBER]") Is Nothing Then
        xmlElem.setAttribute "QUOTATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/@QUOTATIONNUMBER").Text
    Else
        xmlElem.setAttribute "QUOTATIONNUMBER", "1"
    End If
    
    If Not vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@MORTGAGESUBQUOTENUMBER]") Is Nothing Then
        xmlElem.setAttribute "MORTGAGESUBQUOTENUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/@MORTGAGESUBQUOTENUMBER").Text
    Else
        xmlElem.setAttribute "MORTGAGESUBQUOTENUMBER", "1"
    End If
    
    xmlNode.appendChild xmlElem
    
    ' CONTROLDATA node
    Set xmlElem = xmlDoc.createElement("CONTROLDATA")
    
    If Not vxmlRequestNode.selectSingleNode("APPLICATION[@APPLICATIONNUMBER]") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    End If
    
    xmlElem.setAttribute "COPIES", "1"
    xmlElem.setAttribute "DESTINATIONTYPE", "W"
    xmlElem.setAttribute "GEMINIPRINTSTATUS", "30"
    
    xmlElem.setAttribute "DELIVERYTYPE", GetComboValueId("DocumentDeliveryType", "pdf", "20")
    
    xmlElem.setAttribute "DOCUMENTID", gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE/@HOSTTEMPLATEID").Text
    xmlElem.setAttribute "DPSDOCUMENTID", gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE/@DPSTEMPLATEID").Text
    
    xmlNode.appendChild xmlElem
    
    Set objPrint = GetObjectContext.CreateInstance("omPrint.omPrintBO")
    xmlDoc.loadXML objPrint.OmRequest(xmlDoc.xml)
    Set objPrint = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
        
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        LogOmigaError xmlDoc.xml
        Set xmlDoc = Nothing
        Err.Raise oeUnspecifiedError, cstrMethodName, "error creating Application Form"
    End If
    
    EpsomGetApplicationForm = xmlDoc.xml
    Set xmlDoc = Nothing

End Function
' EP2_1062_ends

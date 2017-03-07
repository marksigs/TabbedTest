VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MsgTmBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MsgTmBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Msg Generic Task Management
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'APS    12/03/01    SYS1923 : Issue Offer
'APS    21/03/01    SYS2117 : CompleteSimpleCaseTask
'LD     05/07/01    SYS2458 : Modify GetNextStageTaskList
'BG     11/11/01    SYS3458  Need to send APPLICATIONPRIORITY down to GetStageTaskDetailList in
'                               GetNextStageTaskList method.
'DJP    07/05/02    SYS4534 Don't write Now (Date datatype) to an attribute value.
'CL     15/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'SA     22/10/02    BMIDS00672 Changed FindCaseTaskList to use just the dates not datetimes.
'SA     28/10/02    BMIDS00737 New method (GetStageAddtlTaskDetail) added to get the Additional tasks of a stage (DoRequest changed)
'INR    03/04/03    BM0553      Assigned userid to a task not displayed correctly.
'RF     13/05/2003  BM0541     Performance improvements based on MCAP00407
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Baseline History:
'
'Prog   Date        Description
'MV     10/12/2003  CORE19 - Amended DoRequest() ; FindCaseTaskListLite();Added new Method FindUnitTaskSummary()
'MV     23/08/2004  BBG1210 - Added New method FindTaskNoteArchiveList() and amended DoRequest()
'MV     17/01/2005  E2EM00001288 - Amended CreateTaskNote()
'MV     31/01/2005  BBG1908 - Performance enhancements - out of Memory
'MV     31/01/2005  BBG1908 - Typo Error
'MV     01/08/2005  BBG2146 - New method added - CreateTaskNoteArchive()
'SAB    19/09/2005  CORE157 - Added FindTaskDescriptions and GetStageTask
'GHun   17/01/2006  CORE230 - Changed CreateCaseTask and ValidateTaskAuthority
'------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'GHun   09/03/2006  MAR1300 - Added ValidateProcessTaskAuthority
'GHun   05/04/2006  MAR1300 - Added CreateLinkedTasks, changed UpdateCaseTask and minor tidy up
'GHun   10/04/2006  MAR1606 - Changed UpdateCaseTask
'GHun   11/04/2006  MAR1607 - Changed UpdateCaseTask
'PSC    06/06/2006  MAR1855 - Amend GetCurrentStage to use new method
'AW     14/09/2006  EP1103      CC78 Added GetProgressTasks, GetEditableTaskDetailList
'AW     13/03/2007  EP2_1945    Amended TmRequest for CreatTaskNote operation
'------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IMsgTmBaseBO
Private gobjContext As ObjectContext
'Private gobjTxBO As MsgTm.IMsgTmTxBO
Private gobjDO As MsgTm.IMsgTmDO

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
    'Set gobjTxBO = Nothing
    'Set gobjDO = Nothing
End Sub
Public Function TmRequest(ByVal vstrXmlIn As String) As String
    
    On Error GoTo TmRequestVbErr
    Const strFunctionName As String = "TmRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete
    
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlOutDoc.async = False
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        'AW 13/03/07  EP2_1945
        'Task notes contain carriage returns etc. for formatting. When 'NewParser' is True this is lost
        'even with 'preserveWhiteSpace' set to True
        'Re-load original request with 'NewParser' set to False for this specific operation
        If UCase$(xmlRequestNode.Attributes.getNamedItem("OPERATION").Text) = "CREATETASKNOTE" Then
        
            xmlInDoc.setProperty "NewParser", False
            xmlInDoc.loadXML vstrXmlIn
            
            Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
            
        End If
        'AW 13/03/07  EP2_1945  -   End
        
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    TmRequest = xmlOutDoc.xml
    Debug.Print TmRequest
TmRequestExit:
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Exit Function
TmRequestVbErr:
    
    TmRequest = errCreateErrorResponse()
    Debug.Print TmRequest
    GoTo TmRequestExit
End Function
Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo DoRequestExit
        
    Const strFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase$(strOperation)
    
    Select Case strOperation
        ' debug only
        Case "THROWERROR"
            ThrowError vxmlRequestNode  ', vxmlResponseNode
        Case "CREATECASEACTIVITY"
            CreateCaseActivity vxmlRequestNode, vxmlResponseNode
        Case "CREATECASESTAGE"
            CreateCaseStage vxmlRequestNode, vxmlResponseNode
        Case "CREATECASETASK"
            CreateCaseTask vxmlRequestNode, vxmlResponseNode
        Case "CREATETASKNOTE"
            CreateTaskNote vxmlRequestNode, vxmlResponseNode
        Case "UPDATECASEACTIVITY"
            UpdateCaseActivity vxmlRequestNode, vxmlResponseNode
        Case "COMPLETECASEACTIVITY"
            CompleteCaseActivity vxmlRequestNode, vxmlResponseNode
        Case "UPDATECASESTAGE"
            UpdateCaseStage vxmlRequestNode, vxmlResponseNode
        Case "SETCURRENTCASESTAGECOMPLETE"
            SetCurrentCaseStageComplete vxmlRequestNode, vxmlResponseNode
        Case "UPDATECURRENTSTAGETASKDUEDATES"
            UpdateCurrentStageTaskDueDates vxmlRequestNode, vxmlResponseNode
        Case "UPDATECASETASK"
            UpdateCaseTask vxmlRequestNode, vxmlResponseNode
        Case "GETACTIVITYLIST", "FINDACTIVITYLIST"
            GetActivityList vxmlRequestNode, vxmlResponseNode
        Case "GETSTAGELIST", "FINDSTAGELIST"
            GetStageList vxmlRequestNode, vxmlResponseNode
        Case "GETSTAGETASKLIST", "FINDSTAGETASKLIST"
            GetStageTaskList vxmlRequestNode, vxmlResponseNode
        '++ BMIDS00737 SA new method added
        Case "GETSTAGEADDTLTASKDETAIL", "FINDSTAGETASKLIST"
            GetStageAddtlTaskDetail vxmlRequestNode, vxmlResponseNode
        Case "GETSTAGETASKDETAIL"
            GetStageTaskDetail vxmlRequestNode, vxmlResponseNode
        Case "GETSTAGEDETAIL"
            GetStageDetail vxmlRequestNode, vxmlResponseNode
        Case "GETTASKDETAIL", "GETTASKLIST", "FINDTASKLIST"
            GetTaskDetail vxmlRequestNode, vxmlResponseNode
        Case "GETCASEACTIVITY"
            GetCaseActivity vxmlRequestNode, vxmlResponseNode
        Case "GETCASESTAGE"
            GetCaseStage vxmlRequestNode, vxmlResponseNode
        Case "GETCURRENTSTAGE"
            GetCurrentStage vxmlRequestNode, vxmlResponseNode
        Case "GETCURRENTSTAGESTATUS"
            GetCurrentStageStatus vxmlRequestNode, vxmlResponseNode
        Case "GETNEXTSTAGETASKLIST"
            GetNextStageTaskList vxmlRequestNode, vxmlResponseNode
        Case "FINDCASETASKLIST"
            FindCaseTaskList vxmlRequestNode, vxmlResponseNode
        'RF 13/05/2003 BM0541
        Case "FINDCASETASKLISTLITE"
            FindCaseTaskListLite vxmlRequestNode, vxmlResponseNode
        Case "FINDARCHIVESTAGELIST"
            FindArchiveStageList vxmlRequestNode, vxmlResponseNode
        Case "FINDTASKSTATUSLIST"
            FindTaskStatusList vxmlRequestNode, vxmlResponseNode
        Case "FINDTASKOWNERSHIPLIST"
            FindTaskOwnershipList vxmlRequestNode, vxmlResponseNode
        Case "GETCASETASKS"
            GetCaseTasks vxmlRequestNode, vxmlResponseNode
        Case "GETTASKNOTES", "FINDTASKNOTELIST"
            GetTaskNotes vxmlRequestNode, vxmlResponseNode
        Case "VALIDATESTAGEAUTHORITY"       'JLD SYS1788
            ValidateStageAuthority vxmlRequestNode
        Case "COMPLETESIMPLECASETASK"       'APS SYS2117
            CompleteSimpleCaseTask vxmlRequestNode, vxmlResponseNode
        Case "UPDATECASETRACKNOTE"
            UpdateCaseTrackNote vxmlRequestNode, vxmlResponseNode
        Case "CREATECASETRACKNOTE"
            CreateCaseTrackNote vxmlRequestNode, vxmlResponseNode
        Case "GETLATESTCASETRACKNOTE"
            GetLatestCaseTrackNote vxmlRequestNode, vxmlResponseNode
        Case "FINDUNITTASKSUMMARY"
            FindUnitTaskSummary vxmlRequestNode, vxmlResponseNode
        Case "FINDTASKNOTEARCHIVELIST"
            FindTaskNoteArchiveList vxmlRequestNode, vxmlResponseNode
        Case "CREATETASKNOTEARCHIVE"
            CreateTaskNoteArchive vxmlRequestNode, vxmlResponseNode
        'CORE157
        Case "FINDTASKDESCRIPTIONS"
            FindTaskList vxmlRequestNode, vxmlResponseNode
        Case "GETSTAGETASK"
            GetStageTask vxmlRequestNode, vxmlResponseNode
        'CORE157 - End
        'MAR1300 GHun
        Case "VALIDATEPROCESSTASKAUTHORITY"
            ValidateProcessTaskAuthority vxmlRequestNode
        'MAR1300 End
        ' AW 01/09/06 EP1103 - Start
        Case "GETPROGRESSTASKS"
            GetProgressTasks vxmlRequestNode, vxmlResponseNode
        Case "GETEDITABLETASKDETAILLIST"
            GetEditableTaskDetailList vxmlRequestNode, vxmlResponseNode
        ' AW 01/09/06 EP1103 - End
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select

DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub CreateCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseActivityExit
    Const strFunctionName As String = "CreateCaseActivity"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    Set xmlCaseStageNode = vxmlRequestNode.selectSingleNode("CASEACTIVITY/CASESTAGE")
    
    If Not xmlCaseStageNode Is Nothing Then
        PrepareCaseStage vxmlRequestNode, xmlCaseStageNode
        Set xmlCaseTaskList = xmlCaseStageNode.selectNodes("CASETASK")
        For Each xmlCaseTaskNode In xmlCaseTaskList
            PrepareCaseTask vxmlRequestNode, xmlCaseTaskNode
        Next
    End If
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTm.MsgTmTxBO
    End If
    
    objMsgTmTxBO.CreateCaseActivity vxmlRequestNode, vxmlResponseNode
    
CreateCaseActivityExit:
    
    Set xmlCaseStageNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlCaseTaskList = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CreateCaseStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateCaseStageExit
    Const strFunctionName As String = "CreateCaseStage"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    ValidateStageAuthority vxmlRequestNode
    
    Set xmlCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    
    PrepareCaseStage vxmlRequestNode, xmlCaseStageNode
        
    Set xmlCaseTaskList = xmlCaseStageNode.selectNodes("CASETASK")
    For Each xmlCaseTaskNode In xmlCaseTaskList
        PrepareCaseTask vxmlRequestNode, xmlCaseTaskNode
    Next
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.CreateCaseStage vxmlRequestNode, vxmlResponseNode
    
CreateCaseStageExit:

    Set xmlCaseStageNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlCaseTaskList = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CreateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseTaskExit
    Const strFunctionName As String = "CreateCaseTask"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmTxBO  As IMsgTmTxBO
    Dim xmlRequest As IXMLDOMNode 'CORE230 GHun
    
    'ValidateTaskAuthority vxmlRequestNode 'CORE230 GHun moved into the for loop
    
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    
    For Each xmlCaseTaskNode In vxmlRequestNode.selectNodes("CASETASK")
        'CORE230 GHun call ValidateTaskAuthority for each case task
        Set xmlRequest = vxmlRequestNode.cloneNode(False)
        xmlRequest.appendChild xmlCaseTaskNode.cloneNode(True)
        ValidateTaskAuthority xmlRequest
        'CORE230 End
        
        PrepareCaseTask vxmlRequestNode, xmlCaseTaskNode
    Next
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.CreateCaseTask vxmlRequestNode, vxmlResponseNode
    
CreateCaseTaskExit:
    
    Set xmlCaseTaskNode = Nothing
    Set objMsgTmTxBO = Nothing
    Set xmlRequest = Nothing 'CORE230 GHun
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CreateTaskNote(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateTaskNoteExit
    Const strFunctionName As String = "CreateTaskNote"
    
    Dim xmlTaskNode As IXMLDOMNode
    Dim xmlTaskNodeList As IXMLDOMNodeList
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    xmlCheckMandatoryNode vxmlRequestNode, "TASKNOTE"
    
    Set xmlTaskNodeList = vxmlRequestNode.selectNodes("TASKNOTE")
        
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    For Each xmlTaskNode In xmlTaskNodeList
        xmlCopyAttributeValue vxmlRequestNode, xmlTaskNode, "USERID", "NOTEORIGINATINGUSERID"
        
        If xmlTaskNode.Attributes.getNamedItem("NOTEDATEANDTIME") Is Nothing Then
            Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("NOTEDATEANDTIME")
            'SYS4534 Don't write Now (Date datatype) to an attribute value.
            xmlAttrib.Value = CStr(Now)
            xmlTaskNode.Attributes.setNamedItem xmlAttrib
            Set xmlAttrib = Nothing
        End If
    Next
    
    objMsgTmTxBO.CreateTaskNote vxmlRequestNode, vxmlResponseNode
    
CreateTaskNoteExit:
    
    Set xmlTaskNode = Nothing
    Set xmlTaskNodeList = Nothing
    Set xmlAttrib = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateTaskNoteArchive(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateTaskNoteExit
    Const strFunctionName As String = "CreateTaskNoteArchive"
    
    Dim xmlTaskNode As IXMLDOMNode
    Dim xmlTaskNodeList As IXMLDOMNodeList
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    xmlCheckMandatoryNode vxmlRequestNode, "TASKNOTE"
    
    Set xmlTaskNodeList = vxmlRequestNode.selectNodes("TASKNOTE")
        
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    For Each xmlTaskNode In xmlTaskNodeList
        xmlCopyAttributeValue vxmlRequestNode, xmlTaskNode, "USERID", "NOTEORIGINATINGUSERID"
        
        If xmlTaskNode.Attributes.getNamedItem("NOTEDATEANDTIME") Is Nothing Then
            Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("NOTEDATEANDTIME")
            'SYS4534 Don't write Now (Date datatype) to an attribute value.
            xmlAttrib.Value = CStr(Now)
            xmlTaskNode.Attributes.setNamedItem xmlAttrib
            Set xmlAttrib = Nothing
        End If
    Next
    
    objMsgTmTxBO.CreateTaskNoteArchive vxmlRequestNode, vxmlResponseNode
    
CreateTaskNoteExit:
    
    Set xmlTaskNode = Nothing
    Set xmlTaskNodeList = Nothing
    Set xmlAttrib = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub PrepareCaseStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlCaseStageNode As IXMLDOMNode)
    
    Dim xmlAttrib As IXMLDOMAttribute
        
    xmlCopyAttributeValue vxmlRequestNode, vxmlCaseStageNode, "USERID", "STAGEORIGINATINGUSERID"
    xmlCopyAttributeValue vxmlRequestNode, vxmlCaseStageNode, "UNITID", "STAGEORIGINATINGUNITID"
    
    If vxmlCaseStageNode.Attributes.getNamedItem("STAGEORIGINATINGDATETIME") Is Nothing Then
        
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("STAGEORIGINATINGDATETIME")
        'SYS4534 Don't write Now (Date datatype) to an attribute value.
        xmlAttrib.Value = CStr(Now)
        vxmlCaseStageNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
End Sub
Private Sub PrepareCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlCaseTaskNode As IXMLDOMNode)
        
    Const strFunctionName As String = "PrepareCaseTask"
    Dim xmlAttrib As IXMLDOMAttribute
    Dim intCaseTaskNameSize As Integer  ' MAR1541 / UAT2041
    Dim objNode As IXMLDOMNode          ' MAR1541 / UAT2041
    Dim strCaseTaskName As String       ' MAR1541 / UAT2041
    
    If vxmlCaseTaskNode.Attributes.getNamedItem("OWNINGUSERID") Is Nothing Then
        xmlCopyAttributeValue vxmlRequestNode, vxmlCaseTaskNode, "USERID", "OWNINGUSERID"
    End If
    
    If vxmlCaseTaskNode.Attributes.getNamedItem("OWNINGUNITID") Is Nothing Then
        xmlCopyAttributeValue vxmlRequestNode, vxmlCaseTaskNode, "UNITID", "OWNINGUNITID"
    End If
    If vxmlCaseTaskNode.Attributes.getNamedItem("DATEOFOWNERSHIP") Is Nothing Then
        
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("DATEOFOWNERSHIP")
        'SYS4534 Don't write Now (Date datatype) to an attribute value.
        xmlAttrib.Value = CStr(Now)
        vxmlCaseTaskNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
    If vxmlCaseTaskNode.Attributes.getNamedItem("TASKSTATUSSETBYUSERID") Is Nothing Then
        xmlCopyAttributeValue vxmlRequestNode, vxmlCaseTaskNode, "USERID", "TASKSTATUSSETBYUSERID"
    End If
    If vxmlCaseTaskNode.Attributes.getNamedItem("TASKSTATUSSETBYUNITID") Is Nothing Then
        xmlCopyAttributeValue vxmlRequestNode, vxmlCaseTaskNode, "UNITID", "TASKSTATUSSETBYUNITID"
    End If
    If vxmlCaseTaskNode.Attributes.getNamedItem("TASKSTATUSSETDATETIME") Is Nothing Then
        xmlCopyAttributeValue vxmlCaseTaskNode, vxmlCaseTaskNode, "DATEOFOWNERSHIP", "TASKSTATUSSETDATETIME"
    End If
    
    
    ' Peter Edney - 03/04/06
    ' MAR1541 / UAT2041
    ' If the CASETASKNAME text is longer than 100 characters, then truncate it.
    If Not (vxmlCaseTaskNode.Attributes.getNamedItem("CASETASKNAME") Is Nothing) Then
        
        ' Get the defined size of CASETASKNAME
        Set objNode = adoGetSchema("CASETASK")
        If objNode Is Nothing Then
            errThrowError _
                strFunctionName, _
                oeXMLMissingElement
        End If
        intCaseTaskNameSize = CInt(objNode.selectSingleNode("CASETASKNAME/@LENGTH").Text)
    
        strCaseTaskName = vxmlCaseTaskNode.Attributes.getNamedItem("CASETASKNAME").Text
        If Len(strCaseTaskName) > intCaseTaskNameSize Then
            strCaseTaskName = Left(strCaseTaskName, intCaseTaskNameSize - 3) & "..."
            vxmlCaseTaskNode.Attributes.getNamedItem("CASETASKNAME").Text = strCaseTaskName
        End If
    End If
    
End Sub
Private Sub UpdateCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseActivityExit
    Const strFunctionName As String = "UpdateCaseActivity"
    
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.UpdateCaseActivity vxmlRequestNode, vxmlResponseNode
    
UpdateCaseActivityExit:
    
    Set objMsgTmTxBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CompleteCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CompleteCaseActivityExit
    Const strFunctionName As String = "CompleteCaseActivity"
    
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    Set xmlCaseActivityNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "CASEID"
    
    If xmlAttributeValueExists(xmlCaseActivityNode, "ACTIVITYCOMPLETIONDATETIME") = False Then
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("ACTIVITYCOMPLETIONDATETIME")
        'SYS4534 Don't write Now (Date datatype) to an attribute value.
        xmlAttrib.Value = CStr(Now)
        xmlCaseActivityNode.Attributes.setNamedItem xmlAttrib
    End If
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.UpdateCaseActivity vxmlRequestNode, vxmlResponseNode
    
CompleteCaseActivityExit:
    
    Set xmlAttrib = Nothing
    Set xmlCaseActivityNode = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub UpdateCaseStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseStageExit
    Const strFunctionName As String = "UpdateCaseStage"
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.UpdateCaseStage vxmlRequestNode, vxmlResponseNode
    
UpdateCaseStageExit:
    Set objMsgTmTxBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub SetCurrentCaseStageComplete(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo SetCurrentCaseStageCompleteExit
    Const strFunctionName As String = "SetCurrentCaseStageComplete"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    Set xmlCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
        
    xmlCopyAttributeValue vxmlRequestNode, xmlCaseStageNode, "USERID", "STAGECOMPLETIONUSERID"
    xmlCopyAttributeValue vxmlRequestNode, xmlCaseStageNode, "UNITID", "STAGECOMPLETIONUNITID"
    
    If xmlCaseStageNode.Attributes.getNamedItem("STAGECOMPLETIONDATETIME") Is Nothing Then
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("STAGECOMPLETIONDATETIME")
        'SYS4534 Don't write Now (Date datatype) to an attribute value.
        xmlAttrib.Value = CStr(Now)
        xmlCaseStageNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.UpdateCaseStage vxmlRequestNode, vxmlResponseNode
    
SetCurrentCaseStageCompleteExit:

    Set objMsgTmTxBO = Nothing
    Set xmlCaseStageNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub UpdateCurrentStageTaskDueDates(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCurrentStageTaskDueDatesExit
    Const strFunctionName As String = "UpdateCurrentStageTaskDueDates"
    
    Dim xmlDummyRequestDoc As FreeThreadedDOMDocument40
    Dim xmlDummyRequestNode As IXMLDOMNode
    Dim xmlDummyResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDummyResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    'Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim blnOutstandingTask As Boolean
    Dim intOffset As Integer
    Dim dtDueDate As Date
    Dim strDateOffset As String
    Dim objMsgTxBO As IMsgTmTxBO
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlCaseStageNode, "DATEOFFSET"
    strDateOffset = xmlCaseStageNode.Attributes.getNamedItem("DATEOFFSET").Text
    
    If IsNumeric(strDateOffset) = False Then
        errThrowError _
            strFunctionName, _
            oeXMLInvalidAttributeValue, _
            "[CASETASK/@DATEOFFSET]=" & strDateOffset
    End If
    
    intOffset = CInt(xmlCaseStageNode.Attributes.getNamedItem("DATEOFFSET").Text)
    
    If intOffset = 0 Then
        errThrowError _
            strFunctionName, _
            oeXMLInvalidAttributeValue, _
            "[CASETASK/@DATEOFFSET]=" & strDateOffset
    End If
    
    Set xmlDummyResponseDoc = New FreeThreadedDOMDocument40
    xmlDummyResponseDoc.async = False
    Set xmlElem = xmlDummyResponseDoc.createElement("RESPONSE")
    Set xmlDummyResponseNode = xmlDummyResponseDoc.appendChild(xmlElem)
    ' get Task list for current CaseStage
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetCaseTasks xmlCaseStageNode, xmlDummyResponseNode
    
    If xmlDummyResponseNode.hasChildNodes() = True Then
        ' create UpdateCaseTask request
        Set xmlDummyRequestDoc = New FreeThreadedDOMDocument40
        xmlDummyRequestDoc.async = False
        Set xmlDummyRequestNode = _
           xmlDummyRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        For Each xmlNode In xmlDummyResponseNode.childNodes
            blnOutstandingTask = True
            If Not xmlNode.Attributes.getNamedItem("TASKSTATUS") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("TASKSTATUS").Text <> "0" Then
                    blnOutstandingTask = False
                End If
            End If
            If blnOutstandingTask Then
                If Not xmlNode.Attributes.getNamedItem("TASKDUEDATEANDTIME") Is Nothing Then
                    If Len(xmlNode.Attributes.getNamedItem("TASKDUEDATEANDTIME").Text) > "0" Then
                        dtDueDate = xmlNode.Attributes.getNamedItem("TASKDUEDATEANDTIME").Text
                        dtDueDate = DateAdd("y", intOffset, dtDueDate)
                        xmlNode.Attributes.getNamedItem("TASKDUEDATEANDTIME").Text = dtDueDate
                        xmlDummyRequestNode.appendChild xmlNode.cloneNode(False)
                    End If
                End If
            End If
        Next
    End If
    If xmlDummyRequestNode.hasChildNodes() = True Then
        If Not gobjContext Is Nothing Then
            Set objMsgTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
        Else
            Set objMsgTxBO = New MsgTmTxBO
        End If
        
        objMsgTxBO.UpdateCaseTask xmlDummyRequestNode, vxmlResponseNode
    End If
UpdateCurrentStageTaskDueDatesExit:
    Set objMsgTmDO = Nothing
    Set objMsgTxBO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDummyRequestNode = Nothing
    Set xmlDummyResponseNode = Nothing
    Set xmlDummyRequestDoc = Nothing
    Set xmlDummyResponseDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub UpdateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseTaskExit
    Const strFunctionName As String = "UpdateCaseTask"
    
     'MAR1300 GHun
    'ValidateTaskAuthority vxmlRequestNode 'JLD SYS1761 added user authority check
    ValidateProcessTaskAuthority vxmlRequestNode
     'MAR1300 End
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.UpdateCaseTask vxmlRequestNode, vxmlResponseNode
    
    'MAR1606 GHun Create LinkedTasks if updating the status to complete
    Set xmlElem = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    If xmlGetAttributeText(xmlElem, "TASKSTATUS") = "40" Then   'MAR1607 GHun TaskStatus is not mandatory
        'MAR1300 GHun Create LinkedTasks
        CreateLinkedTasks vxmlRequestNode
    End If
    'MAR1606 End
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASEACTIVITY")
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CASEID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "STAGEID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CASESTAGESEQUENCENO"
    xmlThisRequestNode.appendChild xmlElem
    
    GetCurrentStageStatus xmlThisRequestNode, vxmlResponseNode
    
UpdateCaseTaskExit:
    Set objMsgTmTxBO = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub GetActivityList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetActivityListExit
    Const strFunctionName As String = "GetActivityList"
    
    Dim objMsgTmDO As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("ACTIVITY")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
   objMsgTmDO.GetActivityList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetActivityListExit:
    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetStageList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetStageListExit
    Const strFunctionName As String = "GetStageList"
    
    Dim blnStartSeq As Boolean
    Dim blnEndSeq As Boolean
    Dim intStartSeq As Integer
    Dim intEndSeq As Integer
    Dim xmlEntityNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objMsgTmDO  As IMsgTmDO
    
    Set xmlEntityNode = xmlGetMandatoryNode(vxmlRequestNode, "STAGE")
    Set xmlAttrib = xmlEntityNode.Attributes.getNamedItem("STARTSTAGESEQUENCENO")
    If Not xmlAttrib Is Nothing Then
        If IsNumeric(xmlAttrib.Value) Then
            intStartSeq = CInt(xmlAttrib.Value)
            blnStartSeq = True
        End If
    End If
    Set xmlAttrib = xmlEntityNode.Attributes.getNamedItem("ENDSTAGESEQUENCENO")
    If Not xmlAttrib Is Nothing Then
        If IsNumeric(xmlAttrib.Value) Then
            intEndSeq = CInt(xmlAttrib.Value)
            blnEndSeq = True
        End If
    End If
    If (blnStartSeq = True) And (blnEndSeq = True) Then
        If intEndSeq < intStartSeq Then
            errThrowError strFunctionName, oeXMLInvalidRequestNode, vxmlRequestNode.xml
        End If
    End If
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageList xmlEntityNode, vxmlResponseNode
    
    If (blnStartSeq = True) Or (blnEndSeq = True) Then
        For Each xmlEntityNode In vxmlResponseNode.childNodes
            If blnStartSeq = True Then
                If xmlGetAttributeAsInteger(xmlEntityNode, "STAGESEQUENCENO") < intStartSeq Then
                    vxmlResponseNode.removeChild xmlEntityNode
                End If
            End If
            If blnEndSeq = True Then
                If xmlGetAttributeAsInteger(xmlEntityNode, "STAGESEQUENCENO") > intEndSeq Then
                    vxmlResponseNode.removeChild xmlEntityNode
                End If
            End If
        Next
    End If
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
GetStageListExit:
    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

'MAR1300 GHun never called
'Private Sub GetStageTaskDetailList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'
'    On Error GoTo GetStageTaskDetailListExit
'    Const strFunctionName As String = "GetStageTaskDetailList"
'
'    Dim xmlEntityNode As IXMLDOMNode
'    'Dim xmlAttrib As IXMLDOMAttribute
'    Dim objMsgTmDO As IMsgTmDO
'    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("STAGETASK")
'
'    If Not gobjContext Is Nothing Then
'        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
'    Else
'        Set objMsgTmDO = New MsgTmDO
'    End If
'
'    objMsgTmDO.GetStageTaskDetailList xmlEntityNode, vxmlResponseNode
'
'    If vxmlResponseNode.hasChildNodes = False Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'
'GetStageTaskDetailListExit:
'    Set objMsgTmDO = Nothing
'    Set xmlEntityNode = Nothing
'    errCheckError strFunctionName, TypeName(Me)
'End Sub
'MAR1300 End

Private Sub GetStageTaskDetail(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetStageTaskDetailExit
    Const strFunctionName As String = "GetStageTaskDetail"
    
    Dim objMsgTmDO  As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("STAGETASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageTaskDetailList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
GetStageTaskDetailExit:
    
    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetStageTaskList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetStageTaskListExit
    Const strFunctionName As String = "GetStageTaskList"
    
    Dim objMsgTmDO As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("STAGETASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageTaskList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetStageTaskListExit:
    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetTaskDetail(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetTaskDetailExit
    Const strFunctionName As String = "GetTaskDetail"
    
    Dim objMsgTmDO As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("TASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetTaskList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetTaskDetailExit:
    
    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetCaseActivityExit
    Const strFunctionName As String = "GetCaseActivity"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlParentNodeList As IXMLDOMNodeList
    Dim xmlParentNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseStageNode = vxmlRequestNode.selectSingleNode("CASEACTIVITY")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetCaseActivity xmlCaseStageNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetCaseActivityExit:
    
    Set xmlCaseStageNode = Nothing
    Set xmlParentNodeList = Nothing
    Set xmlParentNode = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetCaseStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetCaseStageExit
    Const strFunctionName As String = "GetCaseStage"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlParentNodeList As IXMLDOMNodeList
    Dim xmlParentNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseStageNode = vxmlRequestNode.selectSingleNode("CASESTAGE")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetCaseStage xmlCaseStageNode, vxmlResponseNode
    
    Set xmlParentNodeList = vxmlResponseNode.selectNodes("CASESTAGE")
    
    For Each xmlParentNode In xmlParentNodeList
       objMsgTmDO.GetCaseTasksDetail xmlParentNode, xmlParentNode
    Next
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetCaseStageExit:
    
    Set xmlCaseStageNode = Nothing
    Set xmlParentNodeList = Nothing
    Set xmlParentNode = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub DoStageStatusSummary(ByVal vxmlCaseStageNode As IXMLDOMNode, Optional ByVal vxmlResultsNode As IXMLDOMNode)
    
    On Error GoTo DoStageStatusSummaryExit
    Const strFunctionName As String = "DoStageStatusSummary"
    
    Dim xmlNode As IXMLDOMNode
    Dim intTotalTasks As Integer
    Dim intIncompleteTasks As Integer
    Dim intIncompleteMandatoryTasks As Integer
    Dim intTaskStatus As Integer
    
    intTotalTasks = vxmlCaseStageNode.selectNodes("CASETASK").length
    For Each xmlNode In vxmlCaseStageNode.childNodes
        
        intTaskStatus = xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS")
        
        If IsValidationType("TaskStatus", intTaskStatus, "I") = True Then
            intIncompleteTasks = intIncompleteTasks + 1
            If xmlGetAttributeText(xmlNode, "MANDATORYINDICATOR") = "1" Then
                intIncompleteMandatoryTasks = intIncompleteMandatoryTasks + 1
            End If
        End If
    Next
        
    If Not vxmlResultsNode Is Nothing Then
        xmlSetAttributeValue vxmlResultsNode, "TOTALTASKS", intTotalTasks
        xmlSetAttributeValue vxmlResultsNode, "INCOMPLETETASKS", intIncompleteTasks
        xmlSetAttributeValue vxmlResultsNode, "INCOMPLETEMANDATORYTASKS", intIncompleteMandatoryTasks
    Else
        xmlSetAttributeValue vxmlCaseStageNode, "TOTALTASKS", intTotalTasks
        xmlSetAttributeValue vxmlCaseStageNode, "INCOMPLETETASKS", intIncompleteTasks
        xmlSetAttributeValue vxmlCaseStageNode, "INCOMPLETEMANDATORYTASKS", intIncompleteMandatoryTasks
    End If
    
DoStageStatusSummaryExit:
    
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub GetCurrentStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetCurrentStageExit
    Const strFunctionName As String = "GetCurrentStage"
    
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseActivityNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    ' PSC 06/06/2006  MAR1855
    objMsgTmDO.GetCurrentStage xmlCaseActivityNode, vxmlResponseNode
        
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetCurrentStageExit:
    
    Set xmlCaseActivityNode = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetCurrentStageStatus(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetCurrentStageStatusExit
    Const strFunctionName As String = "GetCurrentStageStatus"
    
    Dim xmlDummyResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDummyResponseNode As IXMLDOMNode
    Dim xmlCaseStageResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    'Dim intTotalTasks As Integer
    'Dim intIncompleteTasks As Integer
    'Dim intIncompleteMandatoryTasks As Integer
    
    Set xmlDummyResponseDoc = New FreeThreadedDOMDocument40
    xmlDummyResponseDoc.async = False
    Set xmlElem = xmlDummyResponseDoc.createElement("RESPONSE")
    Set xmlDummyResponseNode = xmlDummyResponseDoc.appendChild(xmlElem)
    
    GetCurrentStage vxmlRequestNode, xmlDummyResponseNode
    
    If Not xmlDummyResponseNode.selectSingleNode("CASESTAGE") Is Nothing Then
        Set xmlCaseStageResponseNode = _
            vxmlResponseNode.appendChild( _
                xmlDummyResponseNode.selectSingleNode("CASESTAGE").cloneNode(False))
        DoStageStatusSummary _
            xmlDummyResponseNode.selectSingleNode("CASESTAGE"), _
            xmlCaseStageResponseNode
    End If
GetCurrentStageStatusExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlDummyResponseNode = Nothing
    Set xmlCaseStageResponseNode = Nothing
    Set xmlDummyResponseDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub GetStageDetail(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetStageDetailExit
    Const strFunctionName As String = "GetStageDetail"
    
    Dim objMsgTmDO  As IMsgTmDO
    Dim xmlDummyRequestDoc As FreeThreadedDOMDocument40
    Dim xmlDummyRequestNode As IXMLDOMNode
    Dim xmlStageRequestNode As IXMLDOMNode
    Dim xmlStageNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Set xmlDummyRequestDoc = New FreeThreadedDOMDocument40
    xmlDummyRequestDoc.async = False
    Set xmlElem = xmlDummyRequestDoc.createElement("STAGE")
    Set xmlDummyRequestNode = xmlDummyRequestDoc.appendChild(xmlElem)
    Set xmlStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "STAGE")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageList xmlStageRequestNode, vxmlResponseNode
    
    For Each xmlStageNode In vxmlResponseNode.childNodes
        xmlCopyAttribute xmlStageNode, xmlDummyRequestNode, "STAGEID"
        xmlCopyAttribute xmlStageRequestNode, xmlDummyRequestNode, "CASEPRIORITY"
       objMsgTmDO.GetStageTaskDetailList xmlDummyRequestNode, xmlStageNode
    Next
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetStageDetailExit:
    
    Set xmlElem = Nothing
    Set xmlDummyRequestNode = Nothing
    Set xmlStageRequestNode = Nothing
    Set xmlStageNode = Nothing
    Set xmlDummyRequestDoc = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub GetNextStageTaskList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateActivityExit
    Const strFunctionName As String = "GetNextStageTaskList"
    
    Dim xmlDummyRequestDoc As FreeThreadedDOMDocument40
    Dim xmlDummyRequestNode As IXMLDOMNode
    Dim xmlDummyResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDummyResponseNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strThisStageId As String
    Dim strPattern As String
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    Set xmlDummyRequestDoc = New FreeThreadedDOMDocument40
    xmlDummyRequestDoc.async = False
    Set xmlDummyResponseDoc = New FreeThreadedDOMDocument40
    xmlDummyResponseDoc.async = False
    
    Set xmlElem = xmlDummyRequestDoc.createElement("REQUEST")
    Set xmlDummyRequestNode = xmlDummyRequestDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDummyResponseDoc.createElement("RESPONSE")
    Set xmlDummyResponseNode = xmlDummyResponseDoc.appendChild(xmlElem)
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    ' create GetStageList request for current Activity
    Set xmlElem = xmlDummyRequestDoc.createElement("STAGE")
    xmlCopyAttribute vxmlRequestNode.selectSingleNode("CASESTAGE"), xmlElem, "ACTIVITYID"
    Set xmlDummyRequestNode = xmlDummyRequestNode.appendChild(xmlElem)
    objMsgTmDO.GetStageList xmlDummyRequestNode, xmlDummyResponseNode
        
    If xmlAttributeValueExists(xmlCaseStageNode, "STAGEID") = True Then
        
        ' find current stage in Stage List
        strThisStageId = xmlCaseStageNode.Attributes.getNamedItem("STAGEID").Text
        strPattern = "STAGE[@STAGEID='" & strThisStageId & "']"
        Set xmlNode = xmlDummyResponseNode.selectSingleNode(strPattern)
        If Not xmlNode Is Nothing Then
            Do While xmlNextStageNode Is Nothing
                ' next stage is nextSibling of current stage
                If xmlNode.nextSibling Is Nothing Then
                    Exit Do
                End If
                If xmlGetAttributeAsBoolean(xmlNode.nextSibling, "EXCEPTIONSTAGEINDICATOR") = False Then
                    ' add next Stage to response
                    Set xmlNextStageNode = vxmlResponseNode.appendChild(xmlNode.nextSibling.cloneNode(False))
                Else
                    Set xmlNode = xmlNode.nextSibling
                        
                End If
            Loop
        End If
    Else
        
        ' no current stage, use first stage with a non-null STAGESEQUENCENO (i.e. not cancelled/declined stages)
        For Each xmlNode In xmlDummyResponseNode.childNodes
            If xmlAttributeValueExists(xmlNode, "STAGESEQUENCENO") Then
                Set xmlNextStageNode = vxmlResponseNode.appendChild(xmlNode.cloneNode(False))
                Exit For
            End If
        Next
    End If
    If Not xmlNextStageNode Is Nothing Then
        ' create GetStageTaskDetailList request for next stage
        Set xmlElem = xmlDummyRequestDoc.createElement("STAGETASK")
        xmlCopyAttribute xmlNextStageNode, xmlElem, "STAGEID"
        Set xmlDummyRequestNode = xmlDummyRequestNode.appendChild(xmlElem)
        'BG 11/11/01 SYS3458  Need to send APPLICATIONPRIORITY down to GetStageTaskDetailList
        Dim xmlApplicationPriorityNode As IXMLDOMNode
        Dim xmlCasePriorityNode As IXMLDOMNode
        Set xmlApplicationPriorityNode = vxmlRequestNode.selectSingleNode(".//CASESTAGE/@APPLICATIONPRIORITY")
        If Not xmlApplicationPriorityNode Is Nothing Then
            Set xmlCasePriorityNode = xmlDummyRequestNode.ownerDocument.createAttribute("CASEPRIORITY")
            xmlDummyRequestNode.Attributes.setNamedItem xmlCasePriorityNode
            xmlCasePriorityNode.Text = xmlApplicationPriorityNode.Text
        End If
        ' append Stage Task (Detail) list to response
       objMsgTmDO.GetStageTaskDetailList xmlDummyRequestNode, xmlNextStageNode
        ' drop non 'Default' tasks
        For Each xmlNode In xmlNextStageNode.childNodes
            If xmlGetAttributeAsBoolean(xmlNode, "DEFAULTFLAG") = False Then
                xmlNextStageNode.removeChild xmlNode
            End If
        Next
    End If
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
CreateActivityExit:
    
    Set xmlDummyRequestNode = Nothing
    Set xmlDummyResponseNode = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlNextStageNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDummyRequestDoc = Nothing
    Set xmlDummyResponseDoc = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub FindArchiveStageList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindArchiveStageListExit
    Const strFunctionName As String = "FindArchiveStageList"
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    Set xmlThisRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    
    xmlCheckMandatoryAttribute xmlThisRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYINSTANCE"
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetArchiveStageList xmlThisRequestNode, vxmlResponseNode
    
    For Each xmlThisRequestNode In vxmlResponseNode.childNodes
        objMsgTmDO.GetArchiveStageTaskList xmlThisRequestNode, xmlThisRequestNode
    Next
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindArchiveStageListExit:
    
    Set objMsgTmDO = Nothing
    Set xmlThisRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindTaskNoteArchiveList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindTaskNoteArchiveListExit
    
    Const strFunctionName As String = "FindTaskNoteArchiveList"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("TASKNOTE")
    
    If objMsgTmDO Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.FindTaskNoteArchiveList xmlCaseTaskNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindTaskNoteArchiveListExit:

    Set objMsgTmDO = Nothing
    Set xmlCaseTaskNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindTaskStatusList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindTaskStatusListExit
    Const strFunctionName As String = "FindTaskStatusList"
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlThisRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "TASKSTATUS")
    xmlCheckMandatoryAttribute xmlThisRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYINSTANCE"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "CASESTAGESEQUENCENO"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "TASKID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "TASKINSTANCE"
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetTaskStatusHistoryList xmlThisRequestNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindTaskStatusListExit:
    
    Set objMsgTmDO = Nothing
    Set xmlThisRequestNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub FindTaskOwnershipList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindTaskOwnershipListExit
    Const strFunctionName As String = "FindTaskOwnershipList"
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlThisRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "TASKOWNERSHIPHISTORY")
    xmlCheckMandatoryAttribute xmlThisRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "ACTIVITYINSTANCE"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "CASESTAGESEQUENCENO"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "TASKID"
    xmlCheckMandatoryAttribute xmlThisRequestNode, "TASKINSTANCE"
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetTaskOwnershipHistoryList xmlThisRequestNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindTaskOwnershipListExit:
    Set objMsgTmDO = Nothing
    Set xmlThisRequestNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
' CORE157 - Applied BMIDS performance enhancement
Private Sub FindCaseTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindCaseTaskListExit
    
    Const strFunctionName As String = "FindCaseTaskList"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim strAttribValue As String
    Dim dtLimit As Date, _
        dtTest As Date
    
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")

    If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME") Then
        strAttribValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME")
        If Not IsDate(strAttribValue) Then
            errThrowError _
                strFunctionName, _
                oeXMLInvalidAttributeValue, _
                "TASKDUEDATEANDTIME = " & strAttribValue & " (not a valid date)"
         End If
    End If
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    Set xmlThisRequestNode = vxmlRequestNode.cloneNode(False)
    
    ' drop nodes with TASKDUEDATEANDTIME after TASKDUEDATEANDTIME in request
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"
    
    'SG 28/05/02 SYS4767 START
    'SG 27/02/02 SYS4183
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKTYPE"
    'SG 28/05/02 SYS4767 END
    
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "OWNINGUNITID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "OWNINGUSERID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKSTATUS"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CUSTOMERIDENTIFIER"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, gstrOrderById
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "_COMBOLOOKUP_"
    Set xmlThisCaseTaskNode = xmlThisRequestNode.appendChild(xmlElem)
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    End If
    
    If xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKSTATUS") = "I" Then
        xmlThisCaseTaskNode.Attributes.removeNamedItem ("TASKSTATUS")
        gobjDO.GetOutstandingCaseTasks xmlThisCaseTaskNode, vxmlResponseNode
    Else
        gobjDO.FindCaseTaskList xmlThisCaseTaskNode, vxmlResponseNode
    End If
    
    If vxmlResponseNode.hasChildNodes = True Then
        ' apply any filters
        If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME") Then
            ' drop nodes with TASKDUEDATEANDTIME after TASKDUEDATEANDTIME in request
            dtLimit = _
                xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME")
                
            ' if time not input & due date is today
            ' set time = now
            If DatePart("h", dtLimit) = 0 And _
                DatePart("n", dtLimit) = 0 And _
                DatePart("s", dtLimit) = 0 _
            Then
                If DatePart("yyyy", dtLimit) = DatePart("yyyy", Now()) And _
                    DatePart("m", dtLimit) = DatePart("m", Now()) And _
                    DatePart("d", dtLimit) = DatePart("d", Now()) _
                Then
                    '--BMIDS00672 We need the date - not the datetime.
                    'dtLimit = Now()
                    dtLimit = Format$(Now(), "dd/mm/yyyy")
                End If
            End If
                
            For Each xmlThisCaseTaskNode In vxmlResponseNode.childNodes
            
                If xmlAttributeValueExists( _
                    xmlThisCaseTaskNode, "TASKDUEDATEANDTIME") _
                Then
                
                    dtTest = _
                        xmlGetAttributeText(xmlThisCaseTaskNode, "TASKDUEDATEANDTIME")
                    'BMIDS00672 SA We are not taking time into account so reformat date which comes back.
                    dtTest = Format$(dtTest, "dd/mm/yyyy")
                    If dtTest > dtLimit Then
                        vxmlResponseNode.removeChild xmlThisCaseTaskNode
                    End If
                
                End If
            
            Next
            
        End If
    
    End If
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindCaseTaskListExit:

    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing

    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
' CORE157 - End
Private Sub FindCaseTaskListLite(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindCaseTaskListLiteExit
    
    Const strFunctionName As String = "FindCaseTaskListLite"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objMsgTmDO  As IMsgTmDO
    
    Dim strAttribValue As String
    
    Dim dtLimit As Date
    Dim dtTest As Date
    
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME") Then
        strAttribValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME")
        If Not IsDate(strAttribValue) Then
            errThrowError _
                strFunctionName, _
                oeXMLInvalidAttributeValue, _
                "TASKDUEDATEANDTIME = " & strAttribValue & " (not a valid date)"
         End If
    End If
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlThisRequestNode = vxmlRequestNode.cloneNode(False)
    
    ' drop nodes with TASKDUEDATEANDTIME after TASKDUEDATEANDTIME in request
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "OWNINGUNITID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "OWNINGUSERID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKSTATUS"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CUSTOMERIDENTIFIER"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKTYPE"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, gstrOrderById
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "_COMBOLOOKUP_"
    Set xmlThisCaseTaskNode = xmlThisRequestNode.appendChild(xmlElem)
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    xmlThisCaseTaskNode.Attributes.removeNamedItem ("TASKSTATUS")

    objMsgTmDO.FindCaseTaskListLite xmlThisCaseTaskNode, vxmlResponseNode

    If vxmlResponseNode.hasChildNodes = True Then
        ' apply any filters
        If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME") Then
            ' drop nodes with TASKDUEDATEANDTIME after TASKDUEDATEANDTIME in request
            dtLimit = _
                xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUEDATEANDTIME")
            ' if time not input & due date is today
            ' set time = now
            If DatePart("h", dtLimit) = 0 And _
                DatePart("n", dtLimit) = 0 And _
                DatePart("s", dtLimit) = 0 _
            Then
                If DatePart("yyyy", dtLimit) = DatePart("yyyy", Now()) And _
                    DatePart("m", dtLimit) = DatePart("m", Now()) And _
                    DatePart("d", dtLimit) = DatePart("d", Now()) _
                Then
                    dtLimit = Now()
                End If
            End If
                
            For Each xmlThisCaseTaskNode In vxmlResponseNode.childNodes
                If xmlAttributeValueExists( _
                    xmlThisCaseTaskNode, "TASKDUEDATEANDTIME") _
                Then
                    dtTest = _
                        xmlGetAttributeText(xmlThisCaseTaskNode, "TASKDUEDATEANDTIME")
                    If dtTest > dtLimit Then
                        vxmlResponseNode.removeChild xmlThisCaseTaskNode
                    End If
                End If
            Next
        End If
    End If
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
FindCaseTaskListLiteExit:
    Set objMsgTmDO = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub FindUnitTaskSummary(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName  As String = "FindUnitTaskSummary"
    
    Dim xmlUnit As IXMLDOMNode
    Dim objMsgTmDO   As IMsgTmDO
    
    On Error GoTo FindUnitTaskSummaryExit
    
    Set xmlUnit = xmlGetMandatoryNode(vxmlRequestNode, "UNIT")
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.FindUnitTaskSummary xmlUnit, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindUnitTaskSummaryExit:
    
    Set objMsgTmDO = Nothing
    Set xmlUnit = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
Private Sub GetCaseTasks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetCaseTasksExit
    Const strFunctionName As String = "GetCaseTasks"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetCaseTasksDetail xmlCaseTaskNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetCaseTasksExit:
    
    Set xmlCaseTaskNode = Nothing
    Set objMsgTmDO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetTaskNotes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetTaskNotesExit
    Const strFunctionName As String = "GetTaskNotes"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("TASKNOTE")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetTaskNotes xmlCaseTaskNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetTaskNotesExit:

    Set objMsgTmDO = Nothing
    Set xmlCaseTaskNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub ValidateStageAuthority(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo ValidateStageAuthorityExit
    Const strFunctionName As String = "ValidateStageAuthority"
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlStageRequestNode As IXMLDOMNode
    Dim xmlStageResponseNode As IXMLDOMNode
    Dim objMsgTmDO  As IMsgTmDO
    
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    
    Set xmlStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    xmlCheckMandatoryAttribute xmlStageRequestNode, "STAGEID"
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisResponseDoc.createElement("RESPONSE")
    Set xmlNode = xmlThisResponseDoc.appendChild(xmlElem)
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStage xmlStageRequestNode, xmlNode
    
    Set xmlStageResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/STAGE")
    
    If xmlStageResponseNode Is Nothing Then
        errThrowError strFunctionName, oeTmNoStageDetail
    End If
    
    If xmlGetAttributeAsInteger(vxmlRequestNode, "USERAUTHORITYLEVEL") < _
        xmlGetAttributeAsInteger(xmlStageResponseNode, "USERAUTHORITYLEVEL") _
    Then
        errThrowError strFunctionName, oeTmNoStageAuthority
    End If
        
ValidateStageAuthorityExit:
    Set objMsgTmDO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlStageResponseNode = Nothing
    Set xmlThisResponseDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub ValidateTaskAuthority(ByVal vxmlRequestNode As IXMLDOMNode)
    
    On Error GoTo ValidateTaskAuthorityExit
    Const strFunctionName As String = "ValidateTaskAuthority"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTaskRequestNode As IXMLDOMNode
    Dim xmlTaskResponseNode As IXMLDOMNode
    
    'Dim strResult As String
    'Dim strStageId As String
    Dim objMsgTmDO   As IMsgTmDO
    
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    Set xmlTaskRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlTaskRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlTaskRequestNode, "TASKID"
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisRequestDoc.createElement("STAGETASK")
    xmlCopyAttribute xmlTaskRequestNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlTaskRequestNode, xmlElem, "TASKID"
    ' if ORIGINATINGSTAGEID not equal current STAGEID
    ' use ORIGINATINGSTAGEID for test
    If Not xmlTaskRequestNode.Attributes.getNamedItem("ORIGINATINGSTAGEID") Is Nothing Then
        If xmlTaskRequestNode.Attributes.getNamedItem("ORIGINATINGSTAGEID").Text <> _
            xmlTaskRequestNode.Attributes.getNamedItem("STAGEID").Text _
        Then
            xmlElem.setAttribute _
                "STAGEID", _
                xmlTaskRequestNode.Attributes.getNamedItem("ORIGINATINGSTAGEID").Text
        End If
    End If
    xmlThisRequestDoc.appendChild xmlElem
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisResponseDoc.createElement("RESPONSE")
    Set xmlNode = xmlThisResponseDoc.appendChild(xmlElem)
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageTask xmlThisRequestDoc.firstChild, xmlNode
    
    Set xmlTaskResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/STAGETASK")
    
    If xmlTaskResponseNode Is Nothing Then
        'CORE230 GHun improve error message
        errThrowError strFunctionName, oeTmNoStageTaskAuthority, "TASKID=" & xmlGetAttributeText(xmlTaskRequestNode, "TASKID") & ", STAGEID=" & xmlGetAttributeText(xmlTaskRequestNode, "STAGEID")
    End If
    
    If xmlGetAttributeAsInteger(vxmlRequestNode, "USERAUTHORITYLEVEL") < _
        xmlGetAttributeAsInteger(xmlTaskResponseNode, "USERAUTHORITYLEVEL") _
    Then
        'CORE230 GHun improve error message
        errThrowError strFunctionName, oeTmNoTaskAuthority, "TASKID=" & xmlGetAttributeText(xmlTaskRequestNode, "TASKID") & ", STAGEID=" & xmlGetAttributeText(xmlTaskRequestNode, "STAGEID")
    End If
        
ValidateTaskAuthorityExit:
    Set objMsgTmDO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTaskResponseNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Function IMsgTmBaseBO_TmBaseRequest(ByVal vstrXmlIn As String) As String
    On Error GoTo TmBaseRequestVbErr
    Const strFunctionName As String = "TmBaseRequest"
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim strOperation As String
    'Set gobjTxBO = Nothing
    'Set gobjDO = Nothing
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlInDoc.async = False
    xmlOutDoc.async = False
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    Debug.Print vstrXmlIn
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXmlIn
    Set xmlRequestNode = xmlInDoc.selectSingleNode("REQUEST")
    strOperation = xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    strOperation = UCase$(strOperation)
    Select Case strOperation
        Case "CREATE"
            BaseCreate xmlRequestNode, xmlResponseElem
        Case "UPDATE"
            BaseUpdate xmlRequestNode, xmlResponseElem
        Case "GETLIST"
            BaseGetList xmlRequestNode, xmlResponseElem
        Case Else
            errThrowError _
                "TmBaseRequest", _
                oeNotImplemented, _
                xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    End Select
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    IMsgTmBaseBO_TmBaseRequest = xmlOutDoc.xml
    Debug.Print IMsgTmBaseBO_TmBaseRequest
TmBaseRequestExit:
    'Set gobjTxBO = Nothing
    'Set gobjDO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Function
TmBaseRequestVbErr:
    
    IMsgTmBaseBO_TmBaseRequest = errCreateErrorResponse()
    Debug.Print IMsgTmBaseBO_TmBaseRequest
    GoTo TmBaseRequestExit
End Function
Private Sub BaseCreate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo BaseCreateExit
    Const strFunctionName As String = "BaseCreate"
    
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTm.MsgTmTxBO
    End If
    
    objMsgTmTxBO.BaseCreate vxmlRequestNode, vxmlResponseNode
    
BaseCreateExit:

    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub BaseUpdate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo BaseUpdateExit
    Const strFunctionName As String = "BaseUpdate"
    
    Dim objMsgTmTxBO  As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.BaseUpdate vxmlRequestNode, vxmlResponseNode
    
BaseUpdateExit:
    
    Set objMsgTmTxBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub BaseGetList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo BaseGetListExit
    Const strFunctionName As String = "BaseGetList"
    
    Dim objMsgTmDO As IMsgTmDO
    
    Dim xmlEntityNode As IXMLDOMNode
    Set xmlEntityNode = vxmlRequestNode.firstChild
    
    If xmlEntityNode Is Nothing Then
        errThrowError strFunctionName, oeXMLMissingElement, "No Entity Element specified"
    End If
        
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
        
    Select Case xmlEntityNode.nodeName
        Case "ACTIVITY"
            objMsgTmDO.GetActivityList xmlEntityNode, vxmlResponseNode
        Case "STAGE"
            objMsgTmDO.GetStageList xmlEntityNode, vxmlResponseNode
        Case "STAGETASK"
            objMsgTmDO.GetStageTaskList xmlEntityNode, vxmlResponseNode
        Case "TASK"
            objMsgTmDO.GetTaskList xmlEntityNode, vxmlResponseNode
    End Select
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
BaseGetListExit:

    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub ThrowError(ByVal vxmlRequestNode As IXMLDOMNode) ', ByVal vxmlResponseNode As IXMLDOMNode)
    Dim intErrNum As Integer
    intErrNum = CLng(vxmlRequestNode.Attributes.getNamedItem("ERRNUM").Text)
    errThrowError "bogus", intErrNum
End Sub
Private Sub CompleteSimpleCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CompleteSimpleCaseTaskExit
    Const strFunctionName As String = "CompleteSimpleCaseTask"
    
    Dim lngErrorNo As Long
    
    On Error Resume Next
    
    ' find all tasks matching request xml for the current stage
    Dim xmlDomDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskResponseNode As IXMLDOMNode
        
    Set xmlDomDoc = New FreeThreadedDOMDocument40
    Set xmlCaseTaskResponseNode = xmlDomDoc.createElement("RESPONSE")
    
    FindCaseTaskList vxmlRequestNode, xmlCaseTaskResponseNode
        
    Set xmlDomDoc = Nothing
        
    On Error GoTo CompleteSimpleCaseTaskExit
    lngErrorNo = Err.Number
    If (lngErrorNo <> 0) Then
        ' if we get a record not found error then we need to handle this error otherwise we need to re-raise it
        ' we do however now loose the extra message information which might have been added to the error!!
        If errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
            errThrowError strFunctionName, lngErrorNo
        End If
    Else
        
        Dim xmlNodeList As IXMLDOMNodeList
        Dim xmlCaseTaskNode As IXMLDOMNode
        Dim xmlNode As IXMLDOMNode
        ' get the list of tasks
        Set xmlNodeList = xmlCaseTaskResponseNode.selectNodes("CASETASK")
                
        Dim lngTaskStatus As Long
        Dim intLoop As Integer
        Dim blnFound As Boolean
        blnFound = False
                        
        While blnFound = False And intLoop < xmlNodeList.length
            
            Set xmlNode = xmlNodeList.Item(intLoop)
            If Not xmlNode Is Nothing Then
                
                lngTaskStatus = xmlGetAttributeAsLong(xmlNode, "TASKSTATUS")
                ' find the first one that in incomplete
                If (IsValidationType("TaskStatus", lngTaskStatus, "I")) Then
                                                            
                    blnFound = True
                End If
            End If
            intLoop = intLoop + 1
        Wend
        
        ' if we have found a task which is incomplete then we will update it to complete
        If (blnFound = True And Not xmlNode Is Nothing) Then
            Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("CASETASK")
            xmlCopyAttribute xmlNode, xmlCaseTaskNode, "TASKINSTANCE"
            xmlCopyAttribute xmlNode, xmlCaseTaskNode, "CASESTAGESEQUENCENO"
            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "40"
                        
            UpdateCaseTask vxmlRequestNode, vxmlResponseNode
        End If
                
    End If
    
CompleteSimpleCaseTaskExit:

    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CreateCaseTrackNote(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseTrackNoteExit
    
    Const strFunctionName As String = "CreateCaseTrackNote"
        
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    Set xmlElem = xmlThisRequestDoc.createElement("CASETRACKNOTE")
    'xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "SOURCEAPPLICATION"
    'xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CASEID"
    'xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYID"
    'xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "TEXTVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "TEXT"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CREATEDBYUSERID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CREATEDDATE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VERIFIEDBYUSERID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VERIFIEDDATE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VALIDTODATE"
       
    xmlThisRequestNode.appendChild xmlElem
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    objMsgTmTxBO.CreateCaseTrackNote vxmlRequestNode, vxmlResponseNode
    
CreateCaseTrackNoteExit:
    
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set objMsgTmTxBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub GetLatestCaseTrackNote(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetLatestCaseTrackNoteExit
    Const strFunctionName As String = "GetLatestCaseTrackNote"
    
    'Dim lngErrorNo As Long
    Dim objMsgTmDO As IMsgTmDO
    
    On Error Resume Next
    
    ' find the latest TrackNote
    Dim xmlDomDoc As FreeThreadedDOMDocument40
    Dim xmlGetLatestCaseTrackNoteResponseNode As IXMLDOMNode
        
    Set xmlDomDoc = New FreeThreadedDOMDocument40
    Set xmlGetLatestCaseTrackNoteResponseNode = xmlDomDoc.createElement("RESPONSE")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetLatestCaseTrackNote vxmlRequestNode, vxmlResponseNode
        
    Set xmlDomDoc = Nothing
                
GetLatestCaseTrackNoteExit:
    
    Set objMsgTmDO = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub UpdateCaseTrackNote(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseTrackNoteExit
    Const strFunctionName As String = "UpdateCaseTrackNote"
        
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objMsgTmTxBO As IMsgTmTxBO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmTxBO = gobjContext.CreateInstance("MsgTm.MsgTmTxBO")
    Else
        Set objMsgTmTxBO = New MsgTmTxBO
    End If
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASETRACKNOTE")
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CASEID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "TEXTVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "TEXT"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CREATEDBYUSERID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CREATEDDATE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VERIFIEDBYUSERID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VERIFIEDDATE"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "VALIDTODATE"
       
    xmlThisRequestNode.appendChild xmlElem
    
    objMsgTmTxBO.UpdateCaseTrackNote vxmlRequestNode, vxmlResponseNode
    
UpdateCaseTrackNoteExit:
    Set objMsgTmTxBO = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub DoClientRequest(ByVal strOperation As String, _
                                    ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"
    
    Select Case strOperation
    
    'Enter Customer specific case details here
    Case Else
        errThrowError strFunctionName, oeNotImplemented, strOperation
    End Select
    
DoClientRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub GetStageAddtlTaskDetail(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetStageAddtlTaskDetailExit
    Const strFunctionName As String = "GetStageAddtlTaskDetail"
        
    Dim objMsgTmDO  As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("STAGETASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageAddtlTaskDetailList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetStageAddtlTaskDetailExit:

    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

' CORE157  New Methods
Private Sub FindTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName  As String = "FindTaskList"
    Dim xmlTaskList         As IXMLDOMNode
    
On Error GoTo FindTaskListExit
    
    Set xmlTaskList = xmlGetMandatoryNode(vxmlRequestNode, "TASKLIST")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    End If
    
    gobjDO.FindTaskList xmlTaskList, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindTaskListExit:

    Set xmlTaskList = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
Private Sub GetStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetStageTaskExit
    
    Const strFunctionName As String = "GetStageTask"
    Dim xmlEntityNode As IXMLDOMNode
    
    Set xmlEntityNode = xmlGetMandatoryNode(vxmlRequestNode, "STAGETASK")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    End If
    
    gobjDO.GetStageTask xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetStageTaskExit:
    Set xmlEntityNode = Nothing
    errCheckError strFunctionName, TypeName(Me)

End Sub
' CORE157 - End

'MAR1300 GHun
Private Sub ValidateProcessTaskAuthority(ByVal vxmlRequestNode As IXMLDOMNode)
    
    Const strFunctionName As String = "ValidateProcessTaskAuthority"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTaskRequestNode As IXMLDOMNode
    Dim xmlTaskResponseNode As IXMLDOMNode
    
    Dim objMsgTmDO   As IMsgTmDO
    
On Error GoTo ValidateProcessTaskAuthorityExit
    
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    Set xmlTaskRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlTaskRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlTaskRequestNode, "TASKID"
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisRequestDoc.createElement("STAGETASK")
    xmlCopyAttribute xmlTaskRequestNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlTaskRequestNode, xmlElem, "TASKID"

    xmlThisRequestDoc.appendChild xmlElem
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisResponseDoc.createElement("RESPONSE")
    Set xmlNode = xmlThisResponseDoc.appendChild(xmlElem)
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetStageTask xmlThisRequestDoc.firstChild, xmlNode
    
    Set xmlTaskResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/STAGETASK")
    
    If xmlTaskResponseNode Is Nothing Then
        errThrowError strFunctionName, oeTmNoStageTaskAuthority, "TASKID=" & xmlGetAttributeText(xmlTaskRequestNode, "TASKID") & ", STAGEID=" & xmlGetAttributeText(xmlTaskRequestNode, "STAGEID")
    End If
    
    If xmlGetAttributeAsInteger(vxmlRequestNode, "USERAUTHORITYLEVEL") < _
        xmlGetAttributeAsInteger(xmlTaskResponseNode, "TASKPROCESSAUTHORITYLEVEL") _
    Then
        errThrowError strFunctionName, oeTmNoTaskProcessAuthority, "TASKID=" & xmlGetAttributeText(xmlTaskRequestNode, "TASKID") & ", STAGEID=" & xmlGetAttributeText(xmlTaskRequestNode, "STAGEID") & ", required TASKPROCESSAUTHORITYLEVEL=" & xmlGetAttributeText(xmlTaskResponseNode, "TASKPROCESSAUTHORITYLEVEL")
    End If
    
ValidateProcessTaskAuthorityExit:
    Set objMsgTmDO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTaskResponseNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
'MAR1300 End

'MAR1300 GHun
Private Sub CreateLinkedTasks(ByVal vxmlRequestNode As IXMLDOMNode)
    Const cstrFunctionName  As String = "CreateLinkedTasks"
    Dim objMsgTmDO          As IMsgTmDO
    Dim objTmBO             As Object
    Dim objAppBO            As Object
    Dim xmlResponse         As IXMLDOMNode
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim strResponse         As String
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlElem             As IXMLDOMElement
    Dim xmlTemp             As IXMLDOMElement
    Dim xmlOldCaseTask      As IXMLDOMElement
    Dim xmlNewCaseTask      As IXMLDOMElement
    Dim xmlCaseStage        As IXMLDOMElement

On Error GoTo CreateLinkedTasksExit

    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    Set xmlOldCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")

    'Set xmlRequest = vxmlRequestNode.cloneNode(False)
    Set xmlElem = xmlDoc.createElement("CURRENTCASESTAGE")
    xmlCopyMandatoryAttribute xmlOldCaseTask, xmlElem, "CASEACTIVITYGUID"
    'xmlRequest.appendChild xmlElem

    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
        
    objMsgTmDO.GetCurrentCaseStage xmlElem, xmlResponse
    Set xmlCaseStage = xmlGetMandatoryNode(xmlResponse, "CASESTAGE")
    
    'Set xmlRequest = vxmlRequestNode.cloneNode(False)
    Set xmlElem = xmlDoc.createElement("LINKEDTASK")
    xmlCopyMandatoryAttribute xmlOldCaseTask, xmlElem, "TASKID"
    'xmlRequest.appendChild xmlElem
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    objMsgTmDO.GetLinkedTasks xmlElem, xmlResponse
    
    If Not xmlResponse.selectSingleNode("LINKEDTASK") Is Nothing Then
        Set xmlNewCaseTask = xmlDoc.createElement("CASETASK")
        
        xmlCopyMandatoryAttribute xmlOldCaseTask, xmlNewCaseTask, "SOURCEAPPLICATION"
        xmlCopyMandatoryAttribute xmlOldCaseTask, xmlNewCaseTask, "CASEID"
        xmlCopyMandatoryAttribute xmlOldCaseTask, xmlNewCaseTask, "ACTIVITYID"
        xmlCopyMandatoryAttribute xmlOldCaseTask, xmlNewCaseTask, "ACTIVITYINSTANCE"
        xmlCopyMandatoryAttribute xmlCaseStage, xmlNewCaseTask, "CASESTAGESEQUENCENO"
        xmlCopyMandatoryAttribute xmlCaseStage, xmlNewCaseTask, "STAGEID"
               
        Set xmlRequest = vxmlRequestNode.cloneNode(False)
        Set xmlTemp = xmlDoc.createElement("APPLICATIONPRIORITY")
        xmlRequest.appendChild xmlTemp
        Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlElem.Text = xmlOldCaseTask.getAttribute("CASEID")
        xmlTemp.appendChild xmlElem
        'AddApplicationPriority xmlTemp
        
        If Not gobjContext Is Nothing Then
            Set objTmBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
            Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Else
            Set objTmBO = CreateObject(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
            Set objAppBO = CreateObject(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        End If
        
        strResponse = objAppBO.GetApplicationPriority(xmlRequest.xml)
        xmlDoc.loadXML strResponse
        errCheckXMLResponseNode xmlDoc.documentElement, , True
        
        Set xmlRequest = vxmlRequestNode.cloneNode(False)
        xmlRequest.setAttribute "OPERATION", "CreateAdhocCaseTask"
        
        Set xmlTemp = xmlDoc.createElement("APPLICATION")
        xmlTemp.setAttribute "APPLICATIONPRIORITY", xmlGetMandatoryNodeText(xmlDoc, "/RESPONSE/APPLICATIONPRIORITY/APPLICATIONPRIORITYVALUE")
        xmlRequest.appendChild xmlTemp
        
        For Each xmlElem In xmlResponse.selectNodes("LINKEDTASK")
            Set xmlTemp = xmlNewCaseTask.cloneNode(True)
            xmlTemp.setAttribute "TASKID", xmlElem.getAttribute("LINKEDTASKID")
            xmlRequest.appendChild xmlTemp
        Next
        
        strResponse = objTmBO.omTmRequest(xmlRequest.xml)
        errCheckXMLResponse strResponse, True
    End If
    
CreateLinkedTasksExit:
    Set objMsgTmDO = Nothing
    Set objTmBO = Nothing
    Set xmlResponse = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set objAppBO = Nothing
    Set xmlTemp = Nothing
    Set xmlOldCaseTask = Nothing
    Set xmlNewCaseTask = Nothing
    Set xmlCaseStage = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'MAR1300 End
' AW 01/09/06 EP1103 - Start
Private Sub GetProgressTasks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetProgressTasksExit
    Const strFunctionName As String = "GetProgressTasks"
    
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim objMsgTmDO As IMsgTmDO
    
    Set xmlCaseActivityNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetProgressTasks xmlCaseActivityNode, vxmlResponseNode
        
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetProgressTasksExit:
    
    Set xmlCaseActivityNode = Nothing
    Set objMsgTmDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetEditableTaskDetailList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetEditableTaskDetailListExit
    Const strFunctionName As String = "GetEditableTaskDetailList"
        
    Dim objMsgTmDO  As IMsgTmDO
    Dim xmlEntityNode As IXMLDOMNode
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("STAGETASK")
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    objMsgTmDO.GetEditableTaskDetailList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetEditableTaskDetailListExit:

    Set objMsgTmDO = Nothing
    Set xmlEntityNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
' AW 01/09/06 EP1103 - End

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MsgTmDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'Workfile:      MsgTmDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Msg Generic Task Management
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'
'History:
'
'Prog   Date        Description
'IK     12/10/2001  SYS2803 Work Arounds for MSXML3 IXMLDOMNodeList bug
'-------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'SA     28/10/02    BMIDS00737 New method added - GetStageAddtlTaskDetailList()
'INR    06/05/03    BM0533 GetTaskList using stored procedure returning XML, rather
'                   than using adoGetAsXML.
'RF     13/05/2003  BM0541     Performance improvements based on MCAP00407
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'MV     10/12/2003  CORE20 - Performance improvements based on MCAP00433 for TM010
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        Description
'MV     16/08/2004  BBG1210 - Amended IMsgTmDO_GetTaskNotes()
'MV     23/08/2004  BBG1210 - Amended IMsgTmDO_GetTaskNotes() ;Created New Method  IMsgTmDO_FindTaskNoteArchiveList()
'MV     01/08/2005  BBG2146 - New method added - IMsgTmDO_CreateTaskNoteArchive()
'SAB    19/09/2005  CORE157 - New methods added - FindTaskList() and GetCaseTasksDetailEx()
'                             updated GetTaskList() to return ALWAYSAUTOMATICONCREATION
'PSC    27/09/2005  CORE197 - Update CreateCaseTask to return Request with generated keys
'PSC    02/03/2006  CORE251 - Amend GetTaskList to use an external xslt file
'------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IMsgTmDO
Private gobjContext As ObjectContext

Private Sub IMsgTmDO_CreateTaskNoteArchive(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const strFunctionName As String = "IMsgTmDO_CreateTaskNoteArchive"
    
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    
    adoCreateFromNode vxmlRequestNode, "TASKNOTEARCHIVE"
    
CreateTaskNoteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IMsgTmDO_FindTaskNoteArchiveList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindTaskNoteArchiveListExit
    
    Const strFunctionName As String = "IMsgTmDO_FindTaskNoteArchiveList"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TASKNOTEARCHIVE"
    
FindTaskNoteArchiveListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IMsgTmDO_FindUnitTaskSummary( _
        ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName  As String = "IMsgTmDO_FindUnitTaskSummary"
    Dim adoConnection       As ADODB.Connection
    Dim adoCommand          As ADODB.Command
    Dim adoStream           As ADODB.Stream
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim strUnitID           As String
    On Error GoTo FindUnitTaskSummaryExit
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
            
    strUnitID = xmlGetMandatoryAttributeText(vxmlRequestNode, "UNITID")
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .CursorLocation = adUseClient
        .open
    End With
    With adoCommand
        .CommandText = "USP_TASKSUMMARYTOTALS"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("UNITID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
    End With
    Set adoStream = New ADODB.Stream
    adoStream.open
        
    With adoCommand
        .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "TASKSUMMARY"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    If xmlDoc.loadXML(adoStream.ReadText) Then
        If xmlDoc.selectSingleNode("TASKSUMMARY/OMIGAUSER") Is Nothing Then
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
    Else
        ' raise application error to be interpreted by calling object
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    adoConnection.Close
    Set adoConnection = Nothing
    vxmlResponseNode.appendChild xmlDoc.documentElement
FindUnitTaskSummaryExit:
    
    gobjContext.SetComplete
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlDoc = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Private Sub IMsgTmDO_CreateActivity( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateActivityExit
    Const strFunctionName As String = "CreateActivity"
    adoCreateFromNode vxmlRequestNode, "ACTIVITY"
CreateActivityExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateActivity( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateActivityExit
    Const strFunctionName As String = "UpdateActivity"
    adoUpdateFromNode vxmlRequestNode, "ACTIVITY"
UpdateActivityExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetActivityList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetActivityListExit
    Const strFunctionName As String = "GetActivityList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVEACTIVITY"
GetActivityListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Public Sub IMsgTmDO_CreateStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateStageExit
    Const strFunctionName As String = "CreateStage"
    adoCreateFromNode vxmlRequestNode, "STAGE"
CreateStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateStageExit
    Const strFunctionName As String = "UpdateStage"
    adoUpdateFromNode vxmlRequestNode, "STAGE"
UpdateStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetStageExit
    Const strFunctionName As String = "GetStage"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "STAGE"
GetStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetStageList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetStageListExit
    Const strFunctionName As String = "GetStage"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVESTAGE", , "STAGESEQUENCENO"
GetStageListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetArchiveStageList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetArchiveStageListExit
    Const strFunctionName As String = "GetArchiveStageList"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ARCHIVECASESTAGE", , "STAGESEQUENCENO"
GetArchiveStageListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateStageTaskExit
    Const strFunctionName As String = "CreateStageTask"
    adoCreateFromNode vxmlRequestNode, "STAGETASK"
CreateStageTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateStageTaskExit
    Const strFunctionName As String = "UpdateStageTask"
    adoUpdateFromNode vxmlRequestNode, "STAGETASK"
UpdateStageTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetStageTaskExit
    
    Const strFunctionName As String = "GetStageTask"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "STAGETASK"
GetStageTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetStageTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetStageTaskListExit
    
    Const strFunctionName As String = "GetStageTaskList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVESTAGETASK"
GetStageTaskListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetStageTaskDetailList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetStageTaskDetailListExit
    Const strFunctionName As String = "GetStageTaskDetailList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVESTAGETASKDETAIL"
GetStageTaskDetailListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetArchiveStageTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetArchiveStageTaskListExit
    Const strFunctionName As String = "GetArchiveStageTaskList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ARCHIVECASETASK"
GetArchiveStageTaskListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateTaskExit
    Const strFunctionName As String = "CreateTask"
    adoCreateFromNode vxmlRequestNode, "TASK"
CreateTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateTaskExit
    Const strFunctionName As String = "UpdateTask"
    adoUpdateFromNode vxmlRequestNode, "TASK"
UpdateTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetTaskStatusHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetTaskStatusHistoryListExit
    Const strFunctionName As String = "GetTaskStatusHistoryList"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TASKSTATUSHISTORY"
GetTaskStatusHistoryListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetTaskOwnershipHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetTaskOwnershipHistoryListExit
    Const strFunctionName As String = "GetTaskOwnershipHistoryList"
   ' Burst CASEACTIVITYGUID into logical key SOURCEAPPLICATION + CASEID + ACTIVITYID + ACTIVITYINSTANCE +
   '                                         + STAGEID + CASESEQUENCENO
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TASKOWNERSHIPHISTORY"
GetTaskOwnershipHistoryListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
'INR   BM0533 GetTaskList now using stored procedure returning XML, rather
'                   than using adoGetAsXML.
Private Sub IMsgTmDO_GetTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetTaskListExit
    Const strFunctionName As String = "GetTaskList"
    ' PSC 02/03/2006 CORE251 - Start
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
    ' PSC 02/03/2006 CORE251 - End

    Dim adoConn As ADODB.Connection
    Dim adoCmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim adoParam As ADODB.Parameter
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xslDoc As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    Set xslDoc = New FreeThreadedDOMDocument40
    Set xmlTransformedDoc = New FreeThreadedDOMDocument40
    Set adoConn = New ADODB.Connection
    adoConn.ConnectionString = adoGetDbConnectString
    Set adoCmd = New ADODB.Command
    With adoCmd
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETTASKPRIORITYDETAIL"
    End With
    Set adoParam = New ADODB.Parameter
    With adoParam
        .Name = "pTaskId"
        .Type = adVarChar
        .Direction = adParamInput
        .Size = 20
        .Value = xmlGetAttributeText(vxmlRequestNode, "TASKID")
    End With
    adoCmd.Parameters.Append adoParam
    Set adoParam = New ADODB.Parameter
    With adoParam
        .Name = "PCASEPRIORITY"
        .Type = adInteger
        .Direction = adParamInput
        .Value = xmlGetAttributeAsInteger(vxmlRequestNode, "CASEPRIORITY")
    End With
    adoCmd.Parameters.Append adoParam
    Set adoParam = New ADODB.Parameter
    With adoParam
        .Name = "pStageId"
        .Type = adVarChar
        .Direction = adParamInput
        .Size = 20
        .Value = xmlGetAttributeText(vxmlRequestNode, "STAGEID")
    End With
    adoCmd.Parameters.Append adoParam
    Set adoStr = New ADODB.Stream
    adoConn.open
    adoStr.open
    
    adoCmd.ActiveConnection = adoConn
    adoCmd.Properties("Output Stream") = adoStr
    adoCmd.Execute , , adExecuteStream
        
    ' PSC 02/03/2006 CORE251 - Start
    xslDoc.setProperty "NewParser", True
    xslDoc.validateOnParse = False
    strXslPath = App.Path & "\GetTaskList.xslt"
    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
    
    blnLoadedOk = xslDoc.Load(strXslPath)
    
    If Not blnLoadedOk Then
        errThrowError strFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
    End If
    ' PSC 02/03/2006 CORE251 - End

    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.validateOnParse = False
    
    If xmlTempResponseDoc.loadXML(adoStr.ReadText()) Then
        xmlTempResponseDoc.transformNodeToObject xslDoc.documentElement, _
                                xmlTransformedDoc
        vxmlResponseNode.appendChild xmlTransformedDoc.documentElement
    End If
'BM0553
'    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVETASK"
GetTaskListExit:
    gobjContext.SetComplete
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
            adoConn.Close
        End If
    End If
    Set adoConn = Nothing
    Set adoCmd = Nothing
    Set adoStr = Nothing
    Set adoParam = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateCaseActivity( _
    ByVal vobjCaseActivityNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseActivityExit
    Const strFunctionName As String = "CreateCaseActivity"
    adoCreateFromNode vobjCaseActivityNode, "CASEACTIVITY"
CreateCaseActivityExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateCaseStage( _
    ByVal vobjCaseStageNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseStageExit
    Const strFunctionName As String = "CreateCaseActivity"
    If vobjCaseStageNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vobjCaseStageNode
    End If
    adoCreateFromNode vobjCaseStageNode, "CASESTAGE"
CreateCaseStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateCaseTask( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseTaskExit
    Const strFunctionName As String = "CreateCaseTask"
    AddCaseActivityGuids vxmlRequestNodeList
    adoCreateFromNodeList vxmlRequestNodeList, "CASETASK"
    
    ' PSC 27/09/2005 CORE197 - Start
    Dim xmlCaseTask As IXMLDOMNode
    
    For Each xmlCaseTask In vxmlRequestNodeList
        vxmlResponseNode.appendChild xmlCaseTask.cloneNode(True)
    Next
    ' PSC 27/09/2005 CORE197 - End
    
CreateCaseTaskExit:
    
    Set xmlCaseTask = Nothing
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_CreateTaskNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateTaskNoteExit
    Const strFunctionName As String = "CreateTaskNote"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoCreateFromNode vxmlRequestNode, "TASKNOTE"
CreateTaskNoteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetTaskNotes( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetTaskNotesExit
    
    Const strFunctionName As String = "GetTaskNotes"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TASKNOTE"
GetTaskNotesExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_FindCaseTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo FindCaseTaskListExit
    Const strFunctionName As String = "FindCaseTaskList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASKDETAIL"
FindCaseTaskListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_FindCaseTaskListLite( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'RF     13/05/2003  BM0541     Performance improvements based on MCAP00407
'------------------------------------------------------------------------------
On Error GoTo FindCaseTaskListLiteExit
    
    Const strFunctionName As String = "IMsgTmDO_FindCaseTaskListLite"
    Dim cintValueIds As Collection
    Dim strFilter As String, _
        strInClause As String
    Dim intIndex As Integer
    Set cintValueIds = New Collection
    GetValueIdsForValidationType "TaskStatus", "I", cintValueIds
    For intIndex = 1 To cintValueIds.Count
        If Len(strInClause) > 0 Then
            strInClause = strInClause & ","
        End If
        strInClause = strInClause & CStr(cintValueIds.Item(intIndex))
    Next
    strFilter = "TASKSTATUS IN (" & strInClause & ")"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASKDETAILLITE", strFilter, "TASKDUEDATEANDTIME"
FindCaseTaskListLiteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetCaseTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetCaseTasksExit
    Const strFunctionName As String = "GetCaseTasks"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASK"
GetCaseTasksExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetOutstandingCaseTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetOutstandingCaseTasksExit
    Const strFunctionName As String = "GetOutstandingCaseTasks"
    Dim cintValueIds As Collection
    Dim strFilter As String, _
        strInClause As String
    Dim intIndex As Integer
    Set cintValueIds = New Collection
    GetValueIdsForValidationType "TaskStatus", "I", cintValueIds
    For intIndex = 1 To cintValueIds.Count
        If Len(strInClause) > 0 Then
            strInClause = strInClause & ","
        End If
        strInClause = strInClause & CStr(cintValueIds.Item(intIndex))
    Next
    strFilter = "TASKSTATUS IN (" & strInClause & ")"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASKDETAIL", strFilter
GetOutstandingCaseTasksExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetCaseTasksDetail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetCaseTasksDetailExit
    
    Const strFunctionName As String = "GetCaseTasksDetail"
    If Not vxmlRequestNode.Attributes.getNamedItem("CASEID") Is Nothing Then
        If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
            AddCaseActivityGuid vxmlRequestNode
        End If
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASKDETAIL"
GetCaseTasksDetailExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetCurrentCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetCurrentCaseStageExit
    Const strFunctionName As String = "GetCurrentCaseStage"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CURRENTCASESTAGE"
GetCurrentCaseStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetCaseActivity( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetCaseActivityExit
    Const strFunctionName As String = "GetCaseActivity"
    If vxmlRequestNode.Attributes.getNamedItem("ACTIVITYINSTANCE") Is Nothing Then
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CURRENTCASEACTIVITY"
    Else
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASEACTIVITY"
    End If
GetCaseActivityExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetCaseStageExit
    Const strFunctionName As String = "GetCaseStage"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASESTAGE"
GetCaseStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateCaseActivity( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateCaseActivityExit
    Const strFunctionName As String = "UpdateCaseActivity"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoUpdateFromNode vxmlRequestNode, "CASEACTIVITY"
UpdateCaseActivityExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateCaseStageExit
    Const strFunctionName As String = "UpdateCaseStage"
    If vxmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
        AddCaseActivityGuid vxmlRequestNode
    End If
    adoUpdateFromNode vxmlRequestNode, "CASESTAGE"
UpdateCaseStageExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateCaseTask( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateCaseTaskExit
    Const strFunctionName As String = "UpdateCaseTask"
    AddCaseActivityGuids vxmlRequestNodeList
    adoUpdateFromNodeList vxmlRequestNodeList, "CASETASK"
UpdateCaseTaskExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub AddCaseActivityGuids(ByVal vxmlRequestNodeList As IXMLDOMNodeList)
    On Error GoTo AddCaseActivityGuidsExit
    Const strFunctionName As String = "AddCaseActivityGuids"
    Dim xmlRequestNode As IXMLDOMNode
' IK AQR SYS2803
' fix for MSXML bug
    Dim intIndex As Integer
    For intIndex = 1 To vxmlRequestNodeList.length
        Set xmlRequestNode = vxmlRequestNodeList.Item(intIndex - 1)
'    For Each xmlRequestNode In vxmlRequestNodeList
        If xmlRequestNode.Attributes.getNamedItem("CASEACTIVITYGUID") Is Nothing Then
            AddCaseActivityGuid xmlRequestNode
        End If
    Next
AddCaseActivityGuidsExit:
    
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub AddCaseActivityGuid(ByVal vxmlRequestNode As IXMLDOMNode)
    On Error GoTo AddCaseActivityGuidExit
    Const strFunctionName As String = "AddCaseActivityGuid"
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlResponseDoc.appendChild(xmlElem)
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema("CASEACTIVITY")
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlElem = xmlThisRequestDoc.createElement("REQUEST")
    For Each xmlNode In xmlSchemaNode.childNodes
        If xmlGetAttributeText(xmlNode, "KEYTYPE") = "SECONDARY" Then
            xmlCopyAttribute vxmlRequestNode, xmlElem, xmlNode.nodeName
        End If
    Next
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(xmlElem)
    adoGetAsXML xmlThisRequestNode, xmlResponseNode, "CURRENTCASEACTIVITY"
    If Not xmlResponseNode.selectSingleNode("CASEACTIVITY") Is Nothing Then
        xmlCopyAttribute _
            xmlResponseNode.selectSingleNode("CASEACTIVITY"), _
            vxmlRequestNode, _
            "CASEACTIVITYGUID"
    End If
AddCaseActivityGuidExit:
    
    Set xmlSchemaNode = Nothing
    Set xmlElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlResponseDoc = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
Public Sub IMsgTmDO_CreateCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseTrackNoteExit
    Const strFunctionName As String = "CreateCaseTrackNote"
    adoCreateFromNode vxmlRequestNode, "CASETRACKNOTE"
CreateCaseTrackNoteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_UpdateCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateCaseTrackNoteExit
    Const strFunctionName As String = "UpdateCaseTrackNote"
    adoUpdateFromNode vxmlRequestNode, "CASETRACKNOTE"
UpdateCaseTrackNoteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmDO_GetLatestCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetLatestCaseTrackNoteExit
    Const strFunctionName As String = "GetLatestCaseTrackNote"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETRACKNOTE"
GetLatestCaseTrackNoteExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub
'++ BMIDS00737 SA 28/10/2002 New method added to retrieve all additional tasks on a stage
Private Sub IMsgTmDO_GetStageAddtlTaskDetailList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetStageAddtlTaskDetailListExit
    Const strFunctionName As String = "GetStageAddtlTaskDetailList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACTIVESTAGEADDTLTASKDETAIL"
GetStageAddtlTaskDetailListExit:
    
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)
End Sub

'CORE157  Functions added
Private Sub IMsgTmDO_FindTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindTaskListExit
    Const strFunctionName As String = "IMsgTmDO_FindTaskList"
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TASK", "DELETEFLAG <> 1" 'KRW BMIDS890 14/10/04
    
FindTaskListExit:
    gobjContext.SetComplete
    errCheckError strFunctionName, TypeName(Me)

End Sub
'CORE157 End

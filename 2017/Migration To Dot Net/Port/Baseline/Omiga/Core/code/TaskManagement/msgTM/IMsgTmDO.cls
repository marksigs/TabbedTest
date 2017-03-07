VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IMsgTmDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'SA     28/10/02    BMIDS00737 New method added - GetStageAddtlTaskDetailList()
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Baseline History:
'
'Prog   Date        Description
'MV     10/12/2003  CORE20 - New method added - FindUnitTaskSummary
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Baseline History:
'
'Prog   Date        Description
'MV     23/08/2004  BBG1210 - New method added - FindTaskNoteArchiveList()
'MV     01/08/2005  BBG2146 - New method added - CreateTaskNoteArchive()
'SAB    19/09/2005  CORE157 - New method added - FindTaskList() and GetCaseTasksDetailEx()
'------------------------------------------------------------------------------
Option Explicit
Public Sub CreateActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateStageTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub FindUnitTaskSummary(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetActivityList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetStageList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetArchiveStageList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetStageTask( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetStageTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetArchiveStageTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub FindTaskNoteArchiveList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetStageTaskDetailList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetTaskStatusHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetTaskOwnershipHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseActivity(ByVal vobjCaseActivityNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseStage(ByVal vobjCaseStageNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseTask(ByVal vxmlRequestNodeList As IXMLDOMNodeList, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateTaskNote(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateTaskNoteArchive(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseTask(ByVal vxmlRequestNodeList As IXMLDOMNodeList, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetCaseActivity( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetCurrentCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub FindCaseTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'RF 13/05/2003 BM0541
Public Sub FindCaseTaskListLite( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetCaseTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetOutstandingCaseTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetCaseTasksDetail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetTaskNotes( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub GetLatestCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseTrackNote( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'++ BMIDS00737 SA 28/10/02
'++ New method added to get all the additional tasks for a stage
Public Sub GetStageAddtlTaskDetailList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'CORE157
Public Sub FindTaskList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'CORE157 End

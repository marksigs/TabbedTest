VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "MsgTmTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        Description
'MV     31/01/2005  BBG1908 - Performance enhancements - out of Memory
'MV     31/01/2005  BBG1908 - Typo Error
'MV     01/08/2005  BBG2146 - New method added - IMsgTmTxBO_CreateTaskNoteArchive()
'------------------------------------------------------------------------------
Option Explicit
Implements IMsgTmTxBO
Implements ObjectControl
Private gobjContext As ObjectContext

Private Sub IMsgTmTxBO_CreateCaseTrackNote(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseTrackNoteVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateCaseTrackNote"
    
    Dim xmlCaseTaskList As IXMLDOMNodeList
    
    Dim objMsgTmDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objMsgTmDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objMsgTmDO = New MsgTmDO
    End If
    
    Set xmlCaseTaskList = vxmlRequestParentNode.selectNodes("CASETASK")
        
    If xmlCaseTaskList.length > 0 Then
        objMsgTmDO.CreateCaseTrackNote xmlCaseTaskList, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreateCaseTrackNoteExit:
    
    Set objMsgTmDO = Nothing
    Set xmlCaseTaskList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCaseTrackNoteVbErr:
    
    gobjContext.SetAbort
    GoTo CreateCaseTrackNoteExit
    
End Sub

Private Sub IMsgTmTxBO_CreateTaskNoteArchive(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateTaskNoteVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateTaskNoteArchive"
    
    Dim xmlTaskList As IXMLDOMNodeList
    Dim xmlTaskNode As IXMLDOMNode
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlTaskList = vxmlRequestParentNode.selectNodes("TASKNOTE")
    
    For Each xmlTaskNode In xmlTaskList
        objDO.CreateTaskNoteArchive xmlTaskNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
CreateTaskNoteExit:
    
    Set objDO = Nothing
    Set xmlTaskList = Nothing
    Set xmlTaskNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateTaskNoteVbErr:
    
    gobjContext.SetAbort
    GoTo CreateTaskNoteExit
    
End Sub

Private Sub IMsgTmTxBO_UpdateCaseTrackNote(ByVal vxmlRequestParentNode As MSXML2.IXMLDOMNode, ByVal vxmlResponseNode As MSXML2.IXMLDOMNode)
   
    On Error GoTo UpdateCaseTrackNoteVbErr
    Const strFunctionName As String = "IMsgTmTxBO_UpdateCaseTrackNote"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskNodeList As IXMLDOMNodeList
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
        
    Set xmlCaseTaskNodeList = xmlCaseStageNode.selectNodes("CASETASK")
    objDO.UpdateCaseTrackNote xmlCaseTaskNodeList, vxmlResponseNode
    gobjContext.SetComplete
    
UpdateCaseTrackNoteExit:
    
    Set objDO = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlCaseTaskNodeList = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
UpdateCaseTrackNoteVbErr:
    
    gobjContext.SetAbort
    GoTo UpdateCaseTrackNoteExit
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
Private Sub IMsgTmTxBO_CreateCaseActivity(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseActivityVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateCaseActivity"
    
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlCaseTaskNode As IXMLDOMNode
    
    Dim objDO As IMsgTmDO
   If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseActivityNode = vxmlRequestParentNode.selectSingleNode("CASEACTIVITY")
    objDO.CreateCaseActivity xmlCaseActivityNode, vxmlResponseNode
    
    Set xmlCaseStageNode = xmlCaseActivityNode.selectSingleNode("CASESTAGE")
    If Not xmlCaseStageNode Is Nothing Then
        
        adoPopulateChildKeys xmlCaseActivityNode, xmlCaseStageNode
        objDO.CreateCaseStage xmlCaseStageNode, vxmlResponseNode
        Set xmlCaseTaskList = xmlCaseStageNode.selectNodes("CASETASK")
        If xmlCaseTaskList.length > 0 Then
            
            For Each xmlCaseTaskNode In xmlCaseTaskList
                adoPopulateChildKeys xmlCaseStageNode, xmlCaseTaskNode
            Next
            objDO.CreateCaseTask xmlCaseTaskList, vxmlResponseNode
        End If
    End If
    
    gobjContext.SetComplete
    
CreateCaseActivityExit:
    Set objDO = Nothing
    Set xmlCaseActivityNode = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlCaseTaskList = Nothing
    Set xmlCaseTaskNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
CreateCaseActivityVbErr:
    
    gobjContext.SetAbort
    GoTo CreateCaseActivityExit
    
End Sub
Private Sub IMsgTmTxBO_CreateCaseStage(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseStageVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateCaseStage"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseStageNode = vxmlRequestParentNode.selectSingleNode("CASESTAGE")
    objDO.CreateCaseStage xmlCaseStageNode, vxmlResponseNode
    
    Set xmlCaseTaskList = xmlCaseStageNode.selectNodes("CASETASK")
        
    If xmlCaseTaskList.length > 0 Then
        For Each xmlCaseTaskNode In xmlCaseTaskList
            adoPopulateChildKeys xmlCaseStageNode, xmlCaseTaskNode
        Next
        objDO.CreateCaseTask xmlCaseTaskList, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreateCaseStageExit:
    
    Set objDO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
CreateCaseStageVbErr:

    gobjContext.SetAbort
    GoTo CreateCaseStageExit
    
End Sub
Private Sub IMsgTmTxBO_CreateCaseTask(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCaseTaskVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateCaseTask"
    
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseTaskList = vxmlRequestParentNode.selectNodes("CASETASK")
        
    If xmlCaseTaskList.length > 0 Then
        objDO.CreateCaseTask xmlCaseTaskList, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreateCaseTaskExit:
    
    Set objDO = Nothing
    Set xmlCaseTaskList = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
CreateCaseTaskVbErr:
    
    gobjContext.SetAbort
    GoTo CreateCaseTaskExit
    
End Sub
Private Sub IMsgTmTxBO_CreateTaskNote(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateTaskNoteVbErr
    Const strFunctionName As String = "IMsgTmTxBO_CreateTaskNote"
    
    Dim xmlTaskList As IXMLDOMNodeList
    Dim xmlTaskNode As IXMLDOMNode
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlTaskList = vxmlRequestParentNode.selectNodes("TASKNOTE")
    
    For Each xmlTaskNode In xmlTaskList
        objDO.CreateTaskNote xmlTaskNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
CreateTaskNoteExit:
    
    Set objDO = Nothing
    Set xmlTaskList = Nothing
    Set xmlTaskNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
CreateTaskNoteVbErr:
    
    gobjContext.SetAbort
    GoTo CreateTaskNoteExit
    
End Sub
Private Sub IMsgTmTxBO_UpdateCaseActivity(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseActivityVbErr
    Const strFunctionName As String = "IMsgTmTxBO_UpdateCaseActivity"
    
    Dim xmlCaseActivityList As IXMLDOMNodeList
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseActivityList = vxmlRequestParentNode.selectNodes("CASEACTIVITY")
    
    For Each xmlCaseActivityNode In xmlCaseActivityList
        objDO.UpdateCaseActivity xmlCaseActivityNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateCaseActivityExit:
    
    Set objDO = Nothing
    Set xmlCaseActivityList = Nothing
    Set xmlCaseActivityNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateCaseActivityVbErr:
    
    gobjContext.SetAbort
    GoTo UpdateCaseActivityExit
    
End Sub
Private Sub IMsgTmTxBO_UpdateCaseStage(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseStageVbErr
    Const strFunctionName As String = "IMsgTmTxBO_UpdateCaseStage"
    
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlCaseTaskNodeList As IXMLDOMNodeList
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseStageNode = vxmlRequestParentNode.selectSingleNode("CASESTAGE")
        
    objDO.UpdateCaseStage xmlCaseStageNode, vxmlResponseNode
    Set xmlCaseTaskNodeList = xmlCaseStageNode.selectNodes("CASETASK")
    objDO.UpdateCaseTask xmlCaseTaskNodeList, vxmlResponseNode
    gobjContext.SetComplete
    
UpdateCaseStageExit:
    
    Set objDO = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlCaseTaskNodeList = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateCaseStageVbErr:
    
    gobjContext.SetAbort
    GoTo UpdateCaseStageExit
    
End Sub
Private Sub IMsgTmTxBO_UpdateCaseTask(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateCaseTaskVbErr
    Const strFunctionName As String = "IMsgTmTxBO_UpdateCaseTask"
    
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    
    Set xmlCaseTaskList = vxmlRequestParentNode.selectNodes("CASETASK")
    xmlSetAttributeValue xmlCaseTaskList.Item(0), "TASKSTATUSSETDATETIME", CStr(Now())
    objDO.UpdateCaseTask xmlCaseTaskList, vxmlResponseNode
    gobjContext.SetComplete
    
UpdateCaseTaskExit:
    
    Set objDO = Nothing
    Set xmlCaseTaskList = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateCaseTaskVbErr:
    
    gobjContext.SetAbort
    GoTo UpdateCaseTaskExit
    
End Sub
Private Sub IMsgTmTxBO_BaseCreate(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo BaseCreateVbErr
    Const strFunctionName As String = "IMsgTmTxBO_BaseCreate"
    
    Dim objDO As IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
    RecursiveCreate objDO, vxmlRequestParentNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
BaseCreateExit:
    
    Set objDO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
BaseCreateVbErr:
    
    gobjContext.SetAbort
    GoTo BaseCreateExit
    
End Sub
Private Sub RecursiveCreate( _
    ByVal vobjDO As MsgTm.IMsgTmDO, _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo RecursiveCreateExit
    Const strFunctionName As String = "RecursiveCreate"
    Dim xmlRequestNode As IXMLDOMNode
    For Each xmlRequestNode In vxmlRequestParentNode.childNodes
        If vxmlRequestParentNode.nodeName <> "REQUEST" Then
            adoPopulateChildKeys vxmlRequestParentNode, xmlRequestNode
        End If
        Select Case xmlRequestNode.nodeName
            Case "ACTIVITY"
                vobjDO.CreateActivity xmlRequestNode, vxmlResponseNode
                RecursiveCreate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "STAGE"
                vobjDO.CreateStage xmlRequestNode, vxmlResponseNode
                RecursiveCreate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "STAGETASK"
                vobjDO.CreateStageTask xmlRequestNode, vxmlResponseNode
                RecursiveCreate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "TASK"
                vobjDO.CreateTask xmlRequestNode, vxmlResponseNode
                RecursiveCreate vobjDO, xmlRequestNode, vxmlResponseNode
        End Select
    Next
RecursiveCreateExit:
    
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
Private Sub IMsgTmTxBO_BaseUpdate(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo BaseUpdateVbErr
    Const strFunctionName As String = "IMsgTmTxBO_BaseUpdate"
    
    Dim objDO As MsgTm.IMsgTmDO
    
    If Not gobjContext Is Nothing Then
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
    Else
        Set objDO = New MsgTmDO
    End If
        
    RecursiveUpdate objDO, vxmlRequestParentNode, vxmlResponseNode
    gobjContext.SetComplete
    
BaseUpdateExit:
    
    Set objDO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
BaseUpdateVbErr:
    
    gobjContext.SetAbort
    GoTo BaseUpdateExit
    
End Sub
Private Sub RecursiveUpdate( _
    ByVal vobjDO As MsgTm.IMsgTmDO, _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo RecursiveUpdateExit
    Const strFunctionName As String = "RecursiveUpdate"
    Dim xmlRequestNode As IXMLDOMNode
    For Each xmlRequestNode In vxmlRequestParentNode.childNodes
        If vxmlRequestParentNode.nodeName <> "REQUEST" Then
            adoPopulateChildKeys vxmlRequestParentNode, xmlRequestNode
        End If
        Select Case xmlRequestNode.nodeName
            Case "ACTIVITY"
                vobjDO.UpdateActivity xmlRequestNode, vxmlResponseNode
                RecursiveUpdate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "STAGE"
                vobjDO.UpdateStage xmlRequestNode, vxmlResponseNode
                RecursiveUpdate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "STAGETASK"
                vobjDO.UpdateStageTask xmlRequestNode, vxmlResponseNode
                RecursiveUpdate vobjDO, xmlRequestNode, vxmlResponseNode
            Case "TASK"
                vobjDO.UpdateTask xmlRequestNode, vxmlResponseNode
                RecursiveUpdate vobjDO, xmlRequestNode, vxmlResponseNode
        End Select
    Next
RecursiveUpdateExit:
    
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

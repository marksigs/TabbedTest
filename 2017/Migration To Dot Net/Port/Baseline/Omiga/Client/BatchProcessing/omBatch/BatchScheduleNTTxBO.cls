VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "BatchScheduleNTTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BatchScheduleNTTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: Starts a New Transaction
'Dependencies:
'
'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'INR    13/10/03    BMIDS628 Created

'------------------------------------------------------------------------------------------

Option Explicit

Implements IBatchScheduleNTTXBO

Implements ObjectControl
Private gobjContext As ObjectContext

Private Sub IBatchScheduleNTTXBO_IncrementBatchScheduleFailure(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "IBatchScheduleNTTXBO_IncrementBatchScheduleFailure"
    Dim objBatchScheduleDO As IBatchScheduleDO

On Error GoTo IncrementBatchScheduleFailureVbErr

    Set objBatchScheduleDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    
    objBatchScheduleDO.IncrementBatchScheduleFailure vxmlRequestParentNode, vxmlResponseNode
    gobjContext.SetComplete
    
IncrementBatchScheduleFailureExit:
    Set objBatchScheduleDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

IncrementBatchScheduleFailureVbErr:
    gobjContext.SetAbort
    GoTo IncrementBatchScheduleFailureExit
End Sub


Private Sub IBatchScheduleNTTXBO_IncrementBatchScheduleSuccess(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "IBatchScheduleNTTXBO_IncrementBatchScheduleSuccess"
    Dim objBatchScheduleDO As IBatchScheduleDO

On Error GoTo IncrementBatchScheduleSuccessVbErr

    Set objBatchScheduleDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    
    objBatchScheduleDO.IncrementBatchScheduleSuccess vxmlRequestParentNode, vxmlResponseNode
    gobjContext.SetComplete
    
IncrementBatchScheduleSuccessExit:
    Set objBatchScheduleDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

IncrementBatchScheduleSuccessVbErr:
    gobjContext.SetAbort
    GoTo IncrementBatchScheduleSuccessExit
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


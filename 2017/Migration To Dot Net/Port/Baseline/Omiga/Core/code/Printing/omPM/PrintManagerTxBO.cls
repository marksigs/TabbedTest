VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PrintManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrintManagerTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Created:       26/06/2001
'Author:        Mike Kennedy
'Description:   Transaction BO for the Print Manager component
'Dependencies:  IPrintManagerTxBO.cls
'               omDPS - Document Production System
'               omFVS - File Versioning System
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MK     26/06/01    Created
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
' BMIDS specific History:
' Prog  Date        Description
' IK    17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
Implements IPrintManagerTxBO
Implements ObjectControl
'=============================================
'Variable Declaration Section
'=============================================
    
Private gobjContext As ObjectContext
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Private Sub IPrintManagerTxBO_GetDocumentHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

Const strFunctionName As String = "GetDocumentHistoryList"

Dim objDO As IPrintManagerDO

    On Error GoTo GetDocumentHistoryList_VbErr
    ' ik_20030211
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerDO")
    End If
    objDO.GetDocumentHistoryList vxmlRequestNode, vxmlResponseNode
    ' ik_20030211
    ' gobjContext.SetComplete
GetDocumentHistoryList_Exit:
    
    Set objDO = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
GetDocumentHistoryList_VbErr:
    gobjContext.SetAbort
    GoTo GetDocumentHistoryList_Exit
End Sub
Private Sub IPrintManagerTxBO_GetDocumentEventList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

Const strFunctionName As String = "GetDocumentEventList"

Dim objDO As IPrintManagerDO

    On Error GoTo GetDocumentEventList_VbErr
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerDO")
    End If
    objDO.GetDocumentEventList vxmlRequestNode, vxmlResponseNode

GetDocumentEventList_Exit:
    
    Set objDO = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
GetDocumentEventList_VbErr:
    
    gobjContext.SetAbort
    GoTo GetDocumentEventList_Exit
End Sub
Public Sub IPrintManagerTxBO_PutAuditTrail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
Const strFunctionName As String = "PutAuditTrail"

Dim objDO As IPrintManagerDO

    On Error GoTo CreateAuditTrail_VbErr
    
    ' ik_20030211
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerDO")
    End If
    objDO.PutAuditTrail vxmlRequestNode, vxmlResponseNode
    ' ik_20030211
    ' gobjContext.SetComplete

CreateAuditTrail_Exit:
    Set objDO = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub

CreateAuditTrail_VbErr:
    gobjContext.SetAbort
    GoTo CreateAuditTrail_Exit
End Sub

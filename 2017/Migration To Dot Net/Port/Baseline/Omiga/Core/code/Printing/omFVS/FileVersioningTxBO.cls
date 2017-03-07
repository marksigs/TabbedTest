VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FileVersioningTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'header ----------------------------------------------------------------------------------
'Workfile:      FileVersioningTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Helper functions for error handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'LD    26/01/01     Initial creation
'IK    17/02/2003   BM0200 traceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IFileVersioningTxBO
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
' IFileVersioningTxBO_CallTxBO header ---------------------------------------------------------
' description:
'   Calls Operation in the DO
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
Public Sub IFileVersioningTxBO_CallTxBO( _
    ByVal strOperation As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CallTxBOVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CallTxBO"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim objDO As IFileVersioningDO
    Set objDO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningDO")
    Select Case strOperation
        Case "CHECKOUTFILE"
            objDO.CheckoutFile vxmlRequestNode, vxmlResponseNode
        Case "CHECKINFILE"
            objDO.CheckinFile vxmlRequestNode, vxmlResponseNode
        Case "CREATEFILE"
            objDO.CreateFile vxmlRequestNode, vxmlResponseNode
        Case "CREATEFOLDER"
            objDO.CreateFolder vxmlRequestNode, vxmlResponseNode
        Case "UNLOCKFILE"
            objDO.UnlockFile vxmlRequestNode, vxmlResponseNode
        Case "SHAREFILE"
            objDO.ShareFile vxmlRequestNode, vxmlResponseNode
        Case "GETFILE"
            objDO.GetFile vxmlRequestNode, vxmlResponseNode
        Case Else
            errThrowError _
                "FileVersioningRequest", _
                oeNotImplemented, _
                strOperation
    End Select
        
    ' ik_BM0261
    ' gobjContext.SetComplete
CallTxBOExit:
    
    Set objDO = Nothing
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CallTxBOVbErr:
    ' ik_BM0261
    ' ik_20030211
    gobjContext.SetAbort
    GoTo CallTxBOExit
End Sub

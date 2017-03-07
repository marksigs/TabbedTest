VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "AuditNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AuditNTxBO.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:   New Transaction Business objects class for omAU
'
'Dependencies:  List any other dependent components
'               e.g. AuditDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresNewTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date         Description
'AS     11/04/05    Created
'------------------------------------------------------------------------------------------
Option Explicit
Implements IAuditNTxBO
Implements ObjectControl

Private m_objContext As ObjectContext

Private Function IAuditNTxBO_CreateAccessAudit(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo CreateAccessAuditExit
    Dim objIAuditTxBO As IAuditTxBO
    If m_objContext Is Nothing Then
        Set objIAuditTxBO = New AuditTxBO
    Else
        Set objIAuditTxBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditTxBO")
    End If
    ' Delegate to TxBO.
    Set IAuditNTxBO_CreateAccessAudit = objIAuditTxBO.CreateAccessAudit(vxmlRequest)
CreateAccessAuditExit:
    Set objIAuditTxBO = Nothing
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

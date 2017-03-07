VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "LockManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LockManagerTxBO.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:   A transactional manager object that co-ordinates the calls to create and
'               release locks for the customer and application.
'               Note: XML is not used internally within this module, but only when interfacing
'               to other components that have an XML based interface.
'Dependencies:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AS     05/04/2005  Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private m_objContext As ObjectContext

Public Function omRequest(ByVal vstrXmlIn As String) As String
On Error GoTo omRequestExit
    Dim objLockManagerBO As LockManagerBO
    If m_objContext Is Nothing Then
        Set objLockManagerBO = New LockManagerBO
    Else
        Set objLockManagerBO = m_objContext.CreateInstance(gstrLOCKMANAGER_COMPONENT & ".LockManagerBO")
    End If
    ' Delegate to BO.
    omRequest = objLockManagerBO.omRequest(vstrXmlIn)
omRequestExit:
    Set objLockManagerBO = Nothing
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

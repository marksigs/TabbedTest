VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAuditDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Workfile:      AuditDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Data objects class for omAU
'
'Dependencies:  ADOAssist
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date         Description
'MV     06/03/01    SYS2001: changed CreateAccessAudit Function to CreateAccessaudit Procedure (SUB)
'------------------------------------------------------------------------------------------
Public Sub CreateAccessAudit(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub AddDerivedData(ByVal vxmlData As IXMLDOMNode)
End Sub
Public Function GetNumberOfFailedAttempts(ByVal vxmlTableElement As IXMLDOMElement) As Long
End Function
Public Function IsPasswordChange(ByVal vstrAuditRecType As String) As Boolean
End Function
Public Function IsApplicationAccess(ByVal vstrAuditRecType As String) As Boolean
End Function
Public Function IsApplicationRelease(ByVal vstrAuditRecType As String) As Boolean
End Function
Public Function IsLogon(ByVal vstrAuditRecType As String) As Boolean
End Function
Public Function GetApplicationLockValueId() As String
End Function
Public Function GetApplicationReleaseValueId() As String
End Function
Public Function GetChangePasswordValueId() As String
End Function
Public Function GetLogonValueId() As String
End Function

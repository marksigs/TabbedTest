VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityManagerCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : SecurityManagerCS
' Description   : Contains client specific code for handling Security
' Change history
' Prog      Date        Description
' DJP       19/02/02    SYS4149 Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddSecurity
' Description   : Adds security for client variants. Client specific version
'                 of Core method.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddSecurity(iManagement As ISecurityManager) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    AddSecurity = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ICustomerNTxBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Code template for Customer Business Object interface which requires new transaction
'
'Dependencies:  List any other dependent components
'               none

'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: not MTS (interface only)
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' IK    20/06/02    SYS2372 CREATED
' IK    20/06/02    SYS2372 do GetNextCustomerNumber in separate transaction
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Public Function GetNextCustomerNumber() As Long
End Function


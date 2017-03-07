VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "CustomerNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerNTxBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Code template for Customer Business Object which requires new transaction
'
'Dependencies:  List any other dependent components
'               none

'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: Requires New
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' IK    20/06/02    SYS2372 CREATED
' IK    20/06/02    SYS2372 do GetNextCustomerNumber in separate transaction
' DM    20/11/2002  BMIDS01022 Changed transaction mode to 1 - NoTransactions
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerNTxBO

Private Function ICustomerNTxBO_GetNextCustomerNumber() As Long
    
    Dim objCustomerDO As ICustomerDO
    Set objCustomerDO = GetObjectContext.CreateInstance(App.Title & ".CustomerDO")
    
    ICustomerNTxBO_GetNextCustomerNumber = objCustomerDO.GetNextCustomerNumber
    
    Set objCustomerDO = Nothing
    
End Function

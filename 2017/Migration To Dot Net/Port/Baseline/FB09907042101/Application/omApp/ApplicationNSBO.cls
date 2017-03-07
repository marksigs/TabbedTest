VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationNSBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationNSBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Code template for Application Business Object which does not support transactions
'
'Dependencies:  List any other dependent components
'               none

'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: No transactions
'------------------------------------------------------------------------------------------
' BMids History:
' Prog  Date        Description
' GHun  25/11/2002  BMIDS01043 Moved GetNextApplicationNumber from ApplicationNTxBO, as it should not support transactions after BMIDS01022 change
' SR    08/01/2003  BM0209 - Changed MTSTransactionMode property of the class From 'NoTransactions'
'                   to 'NotAnMTSObject'. Though it was not found to be 100% OK for 10 users in load test, this
'                   was found to be giving better results across multiple runs.
'-----------------------------------------------------------------------------------------------------------------
Option Explicit

Implements IApplicationNSBO

Private Function IApplicationNSBO_GetNextApplicationNumber(ByVal vstrChannelID As String) As String

    Dim objDO As IApplicationDO
    Set objDO = GetObjectContext.CreateInstance(App.Title & ".ApplicationDO")
    IApplicationNSBO_GetNextApplicationNumber = objDO.GetNextApplicationNumber(vstrChannelID)
    Set objDO = Nothing

End Function


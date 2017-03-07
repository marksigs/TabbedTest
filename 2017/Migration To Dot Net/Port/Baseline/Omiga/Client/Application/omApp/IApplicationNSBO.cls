VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationNSBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationNSBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Code template for Application Business Object interface which does not support transactions
'
'Dependencies:  List any other dependent components
'               none

'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: not MTS (interface only)
'------------------------------------------------------------------------------------------
'History:
'
' BMids History:
' Prog  Date        Description
' GHun  25/11/2002  BMIDS01043 Created
'                              Moved GetNextApplicationNumber from IApplicationNTxBO, as it should not support transactions after BMIDS01022 change
'----------------------------------------------------------------------------------------------------------------

Option Explicit

'BMIDS01043
Public Function GetNextApplicationNumber(ByVal vstrChannelID As String) As String
End Function
'BMIDS01043 End

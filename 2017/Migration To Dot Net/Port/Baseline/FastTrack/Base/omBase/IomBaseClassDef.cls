VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomBaseClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IomBaseClassDef.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'Description:   Interfaces to ClassDef for omCm.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MS    09/08/2000    Created for performance rework.
' PF
'------------------------------------------------------------------------------------------
Public Function LoadMessageData() As FreeThreadedDOMDocument40
End Function
Public Function LoadCurrencyData() As FreeThreadedDOMDocument40
End Function

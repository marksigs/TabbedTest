VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataCacheItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Workfile:      IDataCache.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   A way of caching data - recordsets, xml, othre.
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     23/02/00    Created
'------------------------------------------------------------------------------------------

Public Property Get Name() As String
End Property

Public Property Get Count() As Long
End Property

Public Property Get HasData() As Boolean
End Property

Public Property Get AtEndOfData() As Boolean
End Property

Public Property Get XML() As String
End Property

Public Sub GotoFirstEntry()
End Sub

Public Sub GotoNextEntry()
End Sub

Public Function Entry(ByVal vlngIndex As Long) As IDataCacheItem
End Function

Public Function GetValue(ByVal vstrKeyName As String) As Variant
End Function

Public Function Clone(ByVal vstrNewName As String) As IDataCacheItem
End Function

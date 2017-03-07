VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This class is a very simple class to allow the use of counters within rules.
Public Name As String
Private m_lngCount As Long

Public Property Get Count() As Long
    Count = m_lngCount
End Property

Public Function NextValue() As Long
    m_lngCount = m_lngCount + 1
    NextValue = m_lngCount
End Function



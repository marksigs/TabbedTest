VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SwapExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sValueID As String
Private m_sStartDate As String
Private m_sAmount As String
Private m_colExtra As New Collection
Public Function GetValueID() As String
    GetValueID = m_sValueID
End Function
Public Sub SetValueID(sValueID As String)
    m_sValueID = sValueID
End Sub
Public Function GetStartDate() As String
    GetStartDate = m_sStartDate
End Function
Public Sub SetStartDate(sStartDate As String)
    m_sStartDate = sStartDate
End Sub
Public Sub SetExtra(colExtra As Collection)
    Set m_colExtra = colExtra
End Sub
Public Function GetExtra() As Collection
    Set GetExtra = m_colExtra
End Function
Public Sub SetAmount(sAmount As String)
    m_sAmount = sAmount
End Sub
Public Function GetAmount() As String
    GetAmount = m_sAmount
End Function


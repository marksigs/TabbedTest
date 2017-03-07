VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopulateDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_colMatchValues As Collection
Private m_sObjectDescription As String
Private m_sExtraValue As String
Public Sub SetKeyMatchValues(colMatchValues As Collection)
    Set m_colMatchValues = colMatchValues
End Sub
Public Sub SetObjectDescription(sObjectDescription As String)
    m_sObjectDescription = sObjectDescription
End Sub
Public Function GetKeyMatchValues() As Collection
    Set GetKeyMatchValues = m_colMatchValues
End Function
Public Function GetObjectDescription() As String
    GetObjectDescription = m_sObjectDescription
End Function
Public Sub SetExtra(sExtra As String)
    m_sExtraValue = sExtra
End Sub
Public Function GetExtra() As String
    GetExtra = m_sExtraValue
End Function

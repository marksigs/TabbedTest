VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type properties
    bCopied As Boolean
    originalItem As ListItem
    nIndex As Integer
End Type

Private m_listProperties As properties

Public Function IsCopied() As Boolean
    IsCopied = m_listProperties.bCopied
End Function

Public Sub SetCopied(Optional bTrue As Boolean = True)
    m_listProperties.bCopied = bTrue
End Sub

Public Sub SetOriginal(originalItem As ListItem)
    Set m_listProperties.originalItem = originalItem
End Sub

Public Function GetOriginal() As ListItem
    Set GetOriginal = m_listProperties.originalItem
End Function

Public Function SetIndex(nIndex As Integer)
    m_listProperties.nIndex = nIndex
End Function

Public Function GetIndex() As Integer
    GetIndex = m_listProperties.nIndex
End Function


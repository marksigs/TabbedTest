VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TreeItem
' Description   : Stores information about the tree node, and is stored in the tag of a tree node
' Change history
' Prog      Date        Description
' AA        10/04/01    Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type TreeItem
    objExtra As Object
    sType As String
    sTag As String
    bSystemNode As Boolean
    bHeader As Boolean
End Type
Dim m_TreeItem As TreeItem

Public Sub SetExtra(objExtra As Object)
    m_TreeItem.objExtra = objExtra
End Sub
Public Function GetExtra() As Object
    Set GetExtra = m_TreeItem.objExtra
End Function

Public Sub SetType(sType As String)
    m_TreeItem.sType = sType
End Sub
Public Function GetType() As String
    GetType = m_TreeItem.sType
End Function

Public Sub SetTag(sTag As String)
    m_TreeItem.sTag = sTag
End Sub
Public Function GetTag() As String
    GetTag = m_TreeItem.sTag
End Function
Public Sub SetIsSupervisorSystemNode(Optional sIsSupervisorNode As Boolean = False)
    m_TreeItem.bSystemNode = sIsSupervisorNode
End Sub
Public Function GetIsSupervisorSystemNode() As Boolean
    GetIsSupervisorSystemNode = m_TreeItem.bSystemNode
End Function
Public Function GetIsHeader() As Boolean
    GetIsHeader = m_TreeItem.bHeader
End Function
Public Sub SetIsHeader(bHeader As Boolean)
    m_TreeItem.bHeader = bHeader
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TreeAccess
' Description   : Class used to maintain details about a TreeView node.
' Change history
' Prog      Date        Description
' DJP       17/12/01    SYS2831 Client variants.
' JR        27/05/02    SYS4537 Added sKey variable and SetName sub to allow Clients to re-label TreeView
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private Type data
Private Type treeData
    sName As String
    columnHeadings As New Collection
    columnWidth As New Collection
    rightClick As New Collection
    nRelationship As Integer
    nParent As Integer
    nItemPosition As Integer
    sKey As String 'JR SYS4537
End Type

' Private data
Private m_treeData As treeData
''''''''''''''''''''''''''''''' Access Routines for data ''''''''''''''''''''''''''
Public Sub AddName(sName As String, sKey As String, Optional nRelationship As Integer = -1, Optional nParent = -1, Optional nPosition = -1)
    On Error GoTo Failed
    
    m_treeData.sName = sName
    m_treeData.sKey = sKey 'JR SYS4537
    m_treeData.nParent = nParent
    m_treeData.nRelationship = nRelationship
    m_treeData.nItemPosition = nPosition

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub AddHeading(sName As String, nWidth As Long)
    On Error GoTo Failed
    Dim sKey As String
    
    sKey = GetKeyFromString(sName)
    m_treeData.columnHeadings.Add sName, sKey
    m_treeData.columnWidth.Add nWidth, sKey
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetItemPosition() As Integer
    GetItemPosition = m_treeData.nItemPosition
End Function
Public Function GetRelationship() As Integer
    GetRelationship = m_treeData.nRelationship
End Function
Public Function GetParent() As Integer
    GetParent = m_treeData.nParent
End Function
Public Function GetName() As String
    GetName = m_treeData.sName
End Function
Public Function GetColumnHeadings() As Collection
    Set GetColumnHeadings = m_treeData.columnHeadings
End Function
Public Function GetColumnWidths() As Collection
    Set GetColumnWidths = m_treeData.columnWidth
End Function

'JR SYS4537
Public Function GetKey() As String
    GetKey = m_treeData.sKey
End Function

'JR SYS4537, Sets Text of TreeItem
Public Function SetName(sName As String)
    m_treeData.sName = sName
End Function

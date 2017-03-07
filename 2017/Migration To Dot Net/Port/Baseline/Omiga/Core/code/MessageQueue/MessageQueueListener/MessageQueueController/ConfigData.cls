VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ConfigData.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     10/01/01    Created
'------------------------------------------------------------------------------------------

Option Explicit

Public m_SnapInData As SnapInData

' Computers that this snap-in is connected to
Private ArrayComputers() As String ' may contain invalid computer where computers have been removed
Private ArrayComputersCount As Integer
Private Const INVALIDCOMPUTER = ""

Public m_datappgES As New datappgES

Public Sub AddComputer(ByVal strComputerName As String)
    Dim bAdded As Boolean
    bAdded = False
    
    ' try to find an empty slot
    Dim nIndex As Integer
    For nIndex = 1 To ArrayComputersCount
        If (ArrayComputers(nIndex) = INVALIDCOMPUTER) Then
            ArrayComputers(nIndex) = strComputerName
            bAdded = True
            Exit For
        End If
    Next
    
    ' add to end of array is empty slot is not found
    If bAdded = False Then
        ArrayComputersCount = ArrayComputersCount + 1
        ReDim Preserve ArrayComputers(ArrayComputersCount)
        ArrayComputers(ArrayComputersCount) = strComputerName
    End If
End Sub

Public Sub RemoveComputer(ByVal strComputerName As String)
    Dim nIndex As Integer
    For nIndex = 1 To ArrayComputersCount
        If ArrayComputers(nIndex) = strComputerName Then
            ArrayComputers(nIndex) = INVALIDCOMPUTER
            Exit For
        End If
    Next
End Sub
Public Sub RestoreComputers(ByRef nodeParent As ScopeNode)
    Dim nIndex As Integer
    For nIndex = 1 To ArrayComputersCount
        If ArrayComputers(nIndex) <> INVALIDCOMPUTER Then
            m_SnapInData.Computer.AddComputerToScope nodeParent, ArrayComputers(nIndex)
        End If
    Next
End Sub

Public Sub WriteProperties(ByVal PropertyBag As SnapInLib.PropertyBag)

    ' Computers
    PropertyBag.WriteProperty "ArrayComputersCount", ArrayComputersCount
    Dim nIndex As Integer
    For nIndex = 1 To ArrayComputersCount
        PropertyBag.WriteProperty "ArrayComputers_" + CStr(nIndex), ArrayComputers(nIndex)
    Next

    ' EnterpriseServer properties
    m_datappgES.WriteProperties PropertyBag

End Sub

Public Sub ReadProperties(ByVal PropertyBag As SnapInLib.PropertyBag)

    ' Computers
    ArrayComputersCount = PropertyBag.ReadProperty("ArrayComputersCount", 0)
    ReDim Preserve ArrayComputers(ArrayComputersCount)
    Dim nIndex As Integer
    For nIndex = 1 To ArrayComputersCount
        ArrayComputers(nIndex) = PropertyBag.ReadProperty("ArrayComputers_" + CStr(nIndex), INVALIDCOMPUTER)
    Next

    ' EnterpriseServer properties
    m_datappgES.ReadProperties PropertyBag

End Sub




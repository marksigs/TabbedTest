VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Computer.cls
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

Public Sub AddComputersToScope(ByRef nodeParent As ScopeNode)

    On Error GoTo ErrTrap_AddComputersToScope
    
    m_SnapInData.ConfigData.RestoreComputers nodeParent

    Exit Sub

' Error Handler for this method
ErrTrap_AddComputersToScope:
    m_SnapInData.Error.DisplayError "AddComputersToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddComputerToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String)
    
    On Error GoTo ErrTrap_AddComputerToScope
    
    Dim ComputerScopeItem As ScopeItem
    Set ComputerScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_COMPUTER, _
        GetComputerKey(strComputerName), _
        nodeParent, _
        siParent)
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddComputerToScope:
    m_SnapInData.Error.DisplayError "AddComputerToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeComputerScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeComputerScope
    
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(ScopeItem.Key)
    
    ScopeItem.ScopeNode.DisplayName = strComputerName
    'ScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_COMPUTERNAME, COLUMNDISPLAY_COMPUTERNAME
    ScopeItem.ColumnHeaders.Add 2, COLUMNKEY_COMPUTERSTATUS, COLUMNDISPLAY_COMPUTERSTATUS
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_COMPUTERNAME, strComputerName
    ScopeItem.ListSubItems.Add 2, COLUMNKEY_COMPUTERSTATUS, GetComputerStatus(strComputerName)

    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeComputerScope:
    m_SnapInData.Error.DisplayError "InitializeComputerScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function GetComputerStatus(ByVal strComputerName As String) As String
    Select Case GetComputerMQLServiceStatus(strComputerName)
        Case SERVICESTATUS_STARTED
            GetComputerStatus = "MQL Service Running"
        Case SERVICESTATUS_STOPPED
            GetComputerStatus = "MQL Service Stopped"
        Case SERVICESTATUS_NOTFOUND
            GetComputerStatus = "MQL Service Not Found"
        Case SERVICESTATUS_UNKNOWN
            GetComputerStatus = "MQL Service State Not Known"
        Case Else
            Debug.Assert 0
    End Select
End Function

Public Function GetComputerMQLServiceStatus(ByVal strComputerName As String) As SERVICESTATUS
    GetComputerMQLServiceStatus = m_SnapInData.Service.GetServiceStatus(strComputerName, MQLSERVICENAME)
End Function

Public Sub PutComputerMQLServiceStatus(ByVal ScopeItem As SnapInLib.IScopeItem, ByVal RequestedStatus As SERVICESTATUS)
    m_SnapInData.Service.PutServiceStatus GetComputerFromKey(ScopeItem.Key), MQLSERVICENAME, RequestedStatus
    ScopeItem.ListSubItems(2) = GetComputerStatus(GetComputerFromKey(ScopeItem.Key))
End Sub

Public Sub ConnectComputer(ByVal Selection As SnapInLib.IMMCClipboard)
    Dim Shell As Shell32.Shell
    Set Shell = New Shell32.Shell
    Dim Folder As Shell32.Folder
    Const BF_BROWSEFORCOMPUTER = 4096
    Set Folder = Shell.BrowseForFolder(0, "Select a computer on which MessageQueueListener service is installed", BF_BROWSEFORCOMPUTER, ssfNETWORK)
    If Not Folder Is Nothing Then
        AddComputerToScope Selection.ScopeItems(1).ScopeNode, Folder.Title
        m_SnapInData.ConfigData.AddComputer (Folder.Title)
    End If
    Set Shell = Nothing
End Sub

Public Function DeleteComputer(ByVal strComputerName As String) As Boolean
    m_SnapInData.ConfigData.RemoveComputer (strComputerName)
    DeleteComputer = True
End Function


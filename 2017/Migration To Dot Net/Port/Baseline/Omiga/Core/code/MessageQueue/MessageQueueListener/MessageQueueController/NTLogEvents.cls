VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NTLogEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      NTLogEvents.cls
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

Public Sub AddNTLogEventsToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String)
    
    On Error GoTo ErrTrap_AddNTLogEventsToScope
    
    Dim NTLogEventScopeItem As ScopeItem
    Set NTLogEventScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_NTLOGEVENTS, _
        GetNTLogEventsKey(strComputerName), _
        nodeParent, _
        siParent)
    
    NTLogEventScopeItem.SlowRetrieval = True
    
    Exit Sub

' Error Handler for this method
ErrTrap_AddNTLogEventsToScope:
    m_SnapInData.Error.DisplayError "AddNTLogEventsToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeNTLogEventsScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeNTLogEventsScope
    
    ScopeItem.ScopeNode.DisplayName = "NT Event Log (filtered)"
    'ScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_GROUPCOMPUTER, COLUMNDISPLAY_GROUPCOMPUTER
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_GROUPCOMPUTER, "NT Event Log (filtered)"

    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeNTLogEventsScope:
    m_SnapInData.Error.DisplayError "InitializeNTLogEventsScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddNTLogEventsToView(ByRef ResultView As SnapInLib.ResultView)
    
    On Error GoTo ErrTrap_AddNTLogEventsToView
    
    Dim ScopeItem As ScopeItem
    Set ScopeItem = ResultView.ScopePaneItem.ScopeItem
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(ScopeItem.Key)
    
    ResultView.ListView.ListItems.Clear
    
    Dim strNTLogEventFile As String
    Dim strNTLogEventRecord As String
    Dim strNTLogEventType As String
    Dim strNTLogEventDateTime As String
    Dim strNTLogEventMessage As String
    
    Dim WbemServices As SWbemServices
    Set WbemServices = m_SnapInData.WMI.WMIGetWbemServices(strComputerName, "root/cimv2")
    Dim WbemObjectSet As SWbemObjectSet
    Set WbemObjectSet = WbemServices.ExecQuery("SELECT Logfile, RecordNumber, Type, TimeGenerated, Message FROM Win32_NTLogEvent WHERE SourceName = '" + MQLSERVICENAME + "'")
    
    Dim WbemObject As SWbemObject
    Dim WbemProperty As SWbemProperty
    Dim strTimeGenerated As String
    For Each WbemObject In WbemObjectSet
        For Each WbemProperty In WbemObject.Properties_
            Select Case WbemProperty.Name
                Case "Logfile"
                    strNTLogEventFile = WbemProperty.Value
                Case "RecordNumber"
                    strNTLogEventRecord = WbemProperty.Value
                Case "Type"
                    strNTLogEventType = WbemProperty.Value
                Case "TimeGenerated"
                    strTimeGenerated = WbemProperty.Value
                    strNTLogEventDateTime = _
                        Mid(strTimeGenerated, 7, 2) + "/" + _
                        Mid(strTimeGenerated, 5, 2) + "/" + _
                        Mid(strTimeGenerated, 1, 4) + " " + _
                        Mid(strTimeGenerated, 9, 2) + ":" + _
                        Mid(strTimeGenerated, 11, 2) + ":" + _
                        Mid(strTimeGenerated, 13, 2)
                Case "Message"
                    strNTLogEventMessage = Left(WbemProperty.Value, Len(WbemProperty.Value) - 2)
            End Select
        Next
        AddNTLogEventToView ResultView.ListView.ListItems, strComputerName, strNTLogEventFile, strNTLogEventRecord, strNTLogEventType, strNTLogEventDateTime, strNTLogEventMessage
    Next
        
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddNTLogEventsToView:
    m_SnapInData.Error.DisplayError "AddNTLogEventsToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddNTLogEventToView(ByRef ListItems As MMCListItems, ByVal strComputerName As String, ByVal strNTLogEventFile As String, ByVal strNTLogEventRecord As String, ByVal strNTLogEventType As String, ByVal strNTLogEventDateTime As String, ByVal strNTLogEventMessage As String)
    
    On Error GoTo ErrTrap_AddNTLogEventToView
    
    Dim ListItem As MMCListItem
    Set ListItem = ListItems.Add(, GetNTLogEventKey(strComputerName, strNTLogEventFile, strNTLogEventRecord), strNTLogEventRecord)
    ListItem.ListSubItems.Add 1, COLUMNKEY_NTLOGEVENTSTYPE, strNTLogEventType
    ListItem.ListSubItems.Add 2, COLUMNKEY_NTLOGEVENTSDATETIME, strNTLogEventDateTime
    ListItem.ListSubItems.Add 3, COLUMNKEY_NTLOGEVENTSMESSAGE, strNTLogEventMessage
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddNTLogEventToView:
    m_SnapInData.Error.DisplayError "AddNTLogEventToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueueEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QueueEvents.cls
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

Public Sub AddQueueEventsToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String, ByVal strQueueName As String)
    
    On Error GoTo ErrTrap_AddQueueEventsToScope

    Dim QueueEventScopeItem As ScopeItem
    Set QueueEventScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_QUEUEEVENTS, _
        GetQueueEventsKey(strComputerName, strQueueName), _
        nodeParent, _
        siParent)
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddQueueEventsToScope:
    m_SnapInData.Error.DisplayError "AddQueueEventsToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeQueueEventsScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeQueueEventsScope
    
    ScopeItem.ScopeNode.DisplayName = "Events"
    'QueueEventScopeItem.Tag = // extra data - not currently required
       
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_GROUPQUEUE, COLUMNDISPLAY_GROUPQUEUE
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_GROUPQUEUE, "Events"
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeQueueEventsScope:
    m_SnapInData.Error.DisplayError "InitializeQueueEventsScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddQueueEventsToView(ByRef ResultView As SnapInLib.ResultView)
    
    On Error GoTo ErrTrap_AddQueueEventsToView
    
    Dim ScopeItem As ScopeItem
    Set ScopeItem = ResultView.ScopePaneItem.ScopeItem
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(ScopeItem.Key)
    Dim strQueueName As String
    strQueueName = GetQueueFromKey(ScopeItem.Key)
    
    ResultView.ListView.ListItems.Clear
    
    Dim strXMLOut As String
    strXMLOut = m_SnapInData.MQL.MQLConfigureGetQueueInfo(strComputerName, strQueueName)
    
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNodeList As IXMLDOMNodeList
    Set XMLNodeList = XMLDocument.documentElement.selectNodes(".//QUEUELIST/QUEUE/QUEUEEVENTLIST/QUEUEEVENT")
    
    Dim XMLNode As IXMLDOMNode
    Dim strKey As String
    Dim strDayTime As String
    Dim strName As String
    For Each XMLNode In XMLNodeList
        strKey = XMLNode.selectSingleNode(".//KEY").Text
        strDayTime = XMLNode.selectSingleNode(".//DAYTIME").Text
        strName = XMLNode.selectSingleNode(".//NAME").Text
        AddQueueEventToView ResultView.ListView.ListItems, strComputerName, strQueueName, strKey, strDayTime, strName
    Next
    
    Set XMLNode = Nothing
    Set XMLNodeList = Nothing
    Set XMLDocument = Nothing
        
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddQueueEventsToView:
    m_SnapInData.Error.DisplayError "AddQueueEventsToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddQueueEventToView(ByRef ListItems As MMCListItems, ByVal strComputerName As String, ByVal strQueueName As String, ByVal strKey As String, ByVal strDayTime As String, ByVal strName As String)
    
    On Error GoTo ErrTrap_AddQueueEventToView
    
    Dim ListItem As MMCListItem
    Set ListItem = ListItems.Add(, GetQueueEventKey(strComputerName, strQueueName, strKey), strKey)
    ListItem.ListSubItems.Add 1, COLUMNKEY_QUEUEEVENTSDAYTIME, strDayTime
    ListItem.ListSubItems.Add 2, COLUMNKEY_QUEUEEVENTSNAME, strName
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddQueueEventToView:
    m_SnapInData.Error.DisplayError "AddQueueEventToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddQueueEventToView2(datawpgNewQEvent As datawpgNewQEvent)
    
    On Error GoTo ErrTrap_AddQueueEvent
    
    Dim strDayTime As String
    strDayTime = datawpgNewQEvent.m_strQueueEventDay + " " + Right("0" + CStr(datawpgNewQEvent.m_nQueueEventHour), 2) + ":" + Right("0" + CStr(datawpgNewQEvent.m_nQueueEventMinute), 2) + ":" + Right("0" + CStr(datawpgNewQEvent.m_nQueueEventSecond), 2)
    Dim strName As String
    Select Case datawpgNewQEvent.m_strQueueEventType
        Case QUEUEEVENTTYPE_STARTQUEUE
            strName = QUEUEEVENTTYPE_STARTQUEUE
        Case QUEUEEVENTTYPE_STOPQUEUE
            strName = QUEUEEVENTTYPE_STOPQUEUE
        Case QUEUEEVENTTYPE_SETTHREADS
            strName = "Set Threads to " + CStr(datawpgNewQEvent.m_nQueueEventThreads)
        Case Else
            Debug.Assert 0
    End Select
    
    Dim ResultView As ResultView
    Set ResultView = m_SnapInData.SnapIn.CurrentScopePaneItem.ResultView
    If IsSameQueueEventsKey(ResultView.ScopePaneItem.Key, datawpgNewQEvent.m_ScopeItem.Key) Then
        AddQueueEventToView ResultView.ListView.ListItems, datawpgNewQEvent.m_strComputerName, datawpgNewQEvent.m_strQueueName, datawpgNewQEvent.m_strKey, strDayTime, strName
    End If
    
    Exit Sub
' Error Handler for this method
ErrTrap_AddQueueEvent:
    m_SnapInData.Error.DisplayError "AddQueueEvent"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub NewQueueEvent(ByVal ScopeItem As SnapInLib.IScopeItem)

    On Error GoTo ErrTrap_NewQueueEvent
    
    Dim PropertySheetProvider As MMCPropertySheetProvider
    Set PropertySheetProvider = m_SnapInData.SnapIn.Views.CurrentView.PropertySheetProvider
    ScopeItem.Key = ScopeItem.Key
    PropertySheetProvider.CreatePropertySheet "New Queue Event", siWizard, ScopeItem
    PropertySheetProvider.AddPrimaryPages False
    PropertySheetProvider.Show 1

    Exit Sub
    
' Error Handler for this method
ErrTrap_NewQueueEvent:
    m_SnapInData.Error.DisplayError "NewQueueEvent"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function DeleteQueueEvent(ByVal strKey As String) As Boolean
    DeleteQueueEvent = True
End Function


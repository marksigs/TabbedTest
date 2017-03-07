VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StallComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      StalledComponents.cls
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

Public Sub AddStalledComponentsToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String, ByVal strQueueName As String)
    
    On Error GoTo ErrTrap_AddStalledComponentsToScope
    
'TODO - determine computers dynamically
    Dim StalledComponentScopeItem As ScopeItem
    Set StalledComponentScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_STALLEDCOMPONENTS, _
        GetStalledComponentsKey(strComputerName, strQueueName), _
        nodeParent, _
        siParent)
   
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddStalledComponentsToScope:
    m_SnapInData.Error.DisplayError "AddStalledComponentsToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeStalledComponentsScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeStalledComponentsScope
    
    ScopeItem.ScopeNode.DisplayName = "Stalled Components"
    'StalledComponentScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_GROUPQUEUE, COLUMNDISPLAY_GROUPQUEUE
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_GROUPQUEUE, "Stalled Components"
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeStalledComponentsScope:
    m_SnapInData.Error.DisplayError "InitializeStalledComponentsScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddStalledComponentsToView(ByRef ResultView As SnapInLib.ResultView)
    
    On Error GoTo ErrTrap_AddStalledComponentsToView
    
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
    Set XMLNodeList = XMLDocument.documentElement.selectNodes(".//QUEUELIST/QUEUE/COMPONENTLIST/COMPONENT")
    
    Dim XMLNode As IXMLDOMNode
    Dim strComponent As String
    For Each XMLNode In XMLNodeList
        strComponent = XMLNode.Text
        AddStalledComponentToView ResultView.ListView.ListItems, strComputerName, strQueueName, strComponent
    Next
    
    Set XMLNode = Nothing
    Set XMLNodeList = Nothing
    Set XMLDocument = Nothing
        
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddStalledComponentsToView:
    m_SnapInData.Error.DisplayError "AddStalledComponentsToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddStalledComponentToView(ByRef ListItems As MMCListItems, ByVal strComputerName As String, ByVal strQueueName As String, ByVal strProgID As String)
    
    On Error GoTo ErrTrap_AddStalledComponentToView
    
    Dim ListItem As MMCListItem
    Set ListItem = ListItems.Add(, GetStalledComponentKey(strComputerName, strQueueName, strProgID), strProgID)
    ListItem.ListSubItems.Add 1, COLUMNKEY_STALLEDCOMPONENTSPROGID, strProgID
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddStalledComponentToView:
    m_SnapInData.Error.DisplayError "AddStalledComponentToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddStalledComponentToView2(ByVal datawpgStallComp As datawpgStallComp)
    
    On Error GoTo ErrTrap_AddStalledComponentToView
    
    Dim ResultView As ResultView
    Set ResultView = m_SnapInData.SnapIn.CurrentScopePaneItem.ResultView
    If IsSameStalledComponentsKey(ResultView.ScopePaneItem.Key, datawpgStallComp.m_ScopeItem.Key) Then
        AddStalledComponentToView ResultView.ListView.ListItems, datawpgStallComp.m_strComputerName, datawpgStallComp.m_strQueueName, datawpgStallComp.m_strProgID
    End If
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddStalledComponentToView:
    m_SnapInData.Error.DisplayError "AddStalledComponentToView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub RemoveStalledComponentFromView(ByVal ListItem As SnapInLib.IMMCListItem)
    
    On Error GoTo ErrTrap_RemoveStalledComponentFromView
    
    Dim ResultView As ResultView
    Set ResultView = m_SnapInData.SnapIn.CurrentScopePaneItem.ResultView
    If IsChildStalledComponentKey(ResultView.ScopePaneItem.Key, ListItem.Key) Then
        ResultView.ListView.ListItems.Remove ListItem.Key
    End If
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_RemoveStalledComponentFromView:
    m_SnapInData.Error.DisplayError "RemoveStalledComponentFromView"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function StallComponent(ByVal ScopeItem As SnapInLib.IScopeItem) As Boolean

    On Error GoTo ErrTrap_StallComponent
    
    Dim PropertySheetProvider As MMCPropertySheetProvider
    Set PropertySheetProvider = m_SnapInData.SnapIn.Views.CurrentView.PropertySheetProvider
    ScopeItem.Key = ScopeItem.Key
    PropertySheetProvider.CreatePropertySheet "New Stalled Component", siWizard, ScopeItem
    PropertySheetProvider.AddPrimaryPages False
    PropertySheetProvider.Show 1

    Exit Function

' Error Handler for this method
ErrTrap_StallComponent:
    m_SnapInData.Error.DisplayError "StallComponent"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub StartComponent(ByVal ListItem As SnapInLib.IMMCListItem)
    
    On Error GoTo ErrTrap_StartComponent
    
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(ListItem.Key)
    Dim strQueueName As String
    strQueueName = GetQueueFromKey(ListItem.Key)
    Dim strProgID As String
    strProgID = GetProgIDFromKey(ListItem.Key)
        
    ' get the type from the queue
    
        Dim strXMLOut As String
    strXMLOut = m_SnapInData.MQL.MQLConfigureGetQueueInfo(strComputerName, strQueueName)
    
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    Set XMLNode = XMLDocument.documentElement.selectSingleNode(".//QUEUELIST/QUEUE")
    
    Dim strQueueType As String
    strQueueType = XMLNode.selectSingleNode("TYPE").Text
    
    If m_SnapInData.MQL.MQLStartComponent(strComputerName, strQueueName, strProgID, strQueueType) Then
        RemoveStalledComponentFromView ListItem
    End If

    Exit Sub
    
' Error Handler for this method
ErrTrap_StartComponent:
    m_SnapInData.Error.DisplayError "StartComponent"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

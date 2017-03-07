VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Queue.cls
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

Public Enum QUEUESTATUS
    QUEUESTATUS_STARTED
    QUEUESTATUS_STALLED
    QUEUESTATUS_UNKNOWN
End Enum

Public Sub AddQueuesToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String)
    
    On Error GoTo ErrTrap_AddQueuesToScope
    
    Dim strXMLOut As String
    strXMLOut = m_SnapInData.MQL.MQLConfigureGetAllQueuesInfo(strComputerName)
    
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNodeList As IXMLDOMNodeList
    Set XMLNodeList = XMLDocument.documentElement.selectNodes(".//QUEUELIST/QUEUE")
    
    Dim XMLNode As IXMLDOMNode
    Dim strQueueName As String
    For Each XMLNode In XMLNodeList
        strQueueName = XMLNode.selectSingleNode("NAME").Text
        AddQueueToScope nodeParent, strComputerName, strQueueName
    Next
    
    Set XMLNode = Nothing
    Set XMLNodeList = Nothing
    Set XMLDocument = Nothing
    
    Exit Sub

' Error Handler for this method
ErrTrap_AddQueuesToScope:
    m_SnapInData.Error.DisplayError "AddQueuesToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub AddQueueToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String, ByVal strQueueName As String)
    
    On Error GoTo ErrTrap_AddQueueToScope
    
    Dim QueueScopeItem As ScopeItem
    Set QueueScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_QUEUE, _
        GetQueueKey(strComputerName, strQueueName), _
        nodeParent, _
        siParent)
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddQueueToScope:
    m_SnapInData.Error.DisplayError "AddQueueToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeQueueScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeQueueScope
    
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(ScopeItem.Key)
    Dim strQueueName As String
    strQueueName = GetQueueFromKey(ScopeItem.Key)
    
    ScopeItem.ScopeNode.DisplayName = strQueueName
    'ScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_QUEUENAME, COLUMNDISPLAY_QUEUENAME
    ScopeItem.ColumnHeaders.Add 2, COLUMNKEY_QUEUETYPE, COLUMNDISPLAY_QUEUETYPE
    ScopeItem.ColumnHeaders.Add 3, COLUMNKEY_QUEUESTATUS, COLUMNDISPLAY_QUEUESTATUS
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_QUEUENAME, strQueueName
    
    Dim strXMLOut As String
    strXMLOut = m_SnapInData.MQL.MQLConfigureGetQueueInfo(strComputerName, strQueueName)
    
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    Set XMLNode = XMLDocument.documentElement.selectSingleNode(".//QUEUELIST/QUEUE")
    
    Dim strQueueType As String
    Dim strQueueStatus As String
    strQueueType = XMLNode.selectSingleNode("TYPE").Text
    strQueueStatus = XMLNode.selectSingleNode("STATUS").Text
    ScopeItem.ListSubItems.Add 2, COLUMNKEY_QUEUETYPE, strQueueType
    ScopeItem.ListSubItems.Add 3, COLUMNKEY_QUEUESTATUS, strQueueStatus
    
    Set XMLNode = Nothing
    Set XMLDocument = Nothing
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeQueueScope:
    m_SnapInData.Error.DisplayError "InitializeQueueScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function DeleteQueue(strComputerName As String, strQueueName As String) As Boolean
    m_SnapInData.SnapIn.ConsoleMsgBox "delete queue " + strComputerName + "," + strQueueName
    DeleteQueue = True
End Function

Public Function GetComputerQueueStatus(strKey As String) As QUEUESTATUS
    
    On Error GoTo ErrTrap_GetComputerQueueStatus
    GetComputerQueueStatus = QUEUESTATUS_UNKNOWN
    
    Dim strComputerName As String
    strComputerName = GetComputerFromKey(strKey)
    Dim strQueueName As String
    strQueueName = GetQueueFromKey(strKey)

    Dim strXMLOut As String
    strXMLOut = m_SnapInData.MQL.MQLConfigureGetQueueInfo(strComputerName, strQueueName)
    
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    Set XMLNode = XMLDocument.documentElement.selectSingleNode(".//QUEUELIST/QUEUE")
    
    Dim strQueueType As String
    Dim strQueueStatus As String
    strQueueStatus = XMLNode.selectSingleNode("STATUS").Text
    Select Case strQueueStatus
        Case "running"
            GetComputerQueueStatus = QUEUESTATUS_STARTED
        Case "stalled"
            GetComputerQueueStatus = QUEUESTATUS_STALLED
        Case Else
            Debug.Assert 0
    End Select
    
    Set XMLNode = Nothing
    Set XMLDocument = Nothing
    
    GetComputerQueueStatus = QUEUESTATUS_STARTED
    
    Exit Function

' Error Handler for this method
ErrTrap_GetComputerQueueStatus:
    m_SnapInData.Error.DisplayError "GetComputerQueueStatus"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub NewMessageQueue(ByVal ScopeItem As SnapInLib.IScopeItem)

    On Error GoTo ErrTrap_NewMessageQueue
    
    Dim PropertySheetProvider As MMCPropertySheetProvider
    Set PropertySheetProvider = m_SnapInData.SnapIn.Views.CurrentView.PropertySheetProvider
    ScopeItem.Key = ScopeItem.Key
    PropertySheetProvider.CreatePropertySheet "New Message Queue", siWizard, ScopeItem
    PropertySheetProvider.AddPrimaryPages False
    PropertySheetProvider.Show 1

    Exit Sub
    
' Error Handler for this method
ErrTrap_NewMessageQueue:
    m_SnapInData.Error.DisplayError "NewMessageQueue"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

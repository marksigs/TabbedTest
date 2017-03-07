VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueuesFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QueuesFolder.cls
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

Public Sub AddQueuesFolderToScope(ByRef nodeParent As ScopeNode, ByVal strComputerName As String)
    
    On Error GoTo ErrTrap_AddQueueToScope
    
    Dim QueueScopeItem As ScopeItem
    Set QueueScopeItem = m_SnapInData.ScopeItems.AddPreDefined( _
        NODETYPENAME_QUEUESFOLDER, _
        GetQueuesFolderKey(strComputerName), _
        nodeParent, _
        siParent)
    
    Exit Sub
    
' Error Handler for this method
ErrTrap_AddQueueToScope:
    m_SnapInData.Error.DisplayError "AddQueueToScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub InitializeQueuesFolderScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeQueuesFolderScope
    
    ScopeItem.ScopeNode.DisplayName = "Queues"
    'ScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_GROUPCOMPUTER, COLUMNDISPLAY_GROUPCOMPUTER
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_GROUPCOMPUTER, "Queues"
   
    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeQueuesFolderScope:
    m_SnapInData.Error.DisplayError "InitializeQueuesFolderScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



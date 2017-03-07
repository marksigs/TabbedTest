VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewItemsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TreeViewItemsCS
' Description   : Provides client specific functionality for the TreeView to the Supervisor
'                 Main form.
' Change history
' Prog      Date        Description
' DJP       17/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements TreeViewItems

' Private data
Private m_clsTreeViewItems As TreeViewItems
Private Sub Class_Initialize()
    Set m_clsTreeViewItems = New TreeViewItems
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   TreeViewItems_SetupTreeValues
' Description   :   Sets up the treeview with all required nodes that need to
'                   be displayed. Delegate to the Core method, then do any client specific
'                   processing as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TreeViewItems_SetupTreeValues()
    On Error GoTo Failed
    
    ' Create Core nodes
    m_clsTreeViewItems.SetupTreeValues
        
    ' Any Client specific Nodes
    CreateClientNodes
    
    ' Create the nodes
    m_clsTreeViewItems.AddTreeNodes
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub TreeViewItems_AddTreeNodes()
    m_clsTreeViewItems.AddTreeNodes
End Sub
Private Sub TreeViewItems_ClearCurrentTreeitem()
    m_clsTreeViewItems.ClearCurrentTreeitem
End Sub
Private Function TreeViewItems_GetSelectedTreeKey() As Variant
    TreeViewItems_GetSelectedTreeKey = m_clsTreeViewItems.GetSelectedTreeKey
End Function
Private Function TreeViewItems_GetTreeData() As Collection
    Set TreeViewItems_GetTreeData = m_clsTreeViewItems.GetTreeData
End Function
Private Function TreeViewItems_GetTreeItem(sTag As String) As TreeAccess
    Set TreeViewItems_GetTreeItem = m_clsTreeViewItems.GetTreeItem(sTag)
End Function
Private Sub TreeViewItems_SetTreeStatus(Optional node As MSComctlLib.INode = Nothing)
    m_clsTreeViewItems.SetTreeStatus node
End Sub
''''''''''''''''''''''''''''''''' Client Specific methods ''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CreateClientNodes
' Description   :   Sets up the treeview with all required nodes that need to
'                   be displayed. These are the Core treeview items that need to be displayed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateClientNodes()
    On Error GoTo Failed

    ' Example of how to add a new node to the Supervisor TreeView. This example would add a new
    ' node to the SYSTEM_PARAMETERS node of the TreeView.

'    Dim globalParamItem As TreeAccess
'    Dim tmpTreeItem As TreeAccess
'    Dim colTreeData As Collection
'    Dim nGlobalParam As Integer
'
'    Set colTreeData = m_clsTreeViewItems.GetTreeData
'
'
'    Set globalParamItem = m_clsTreeViewItems.GetTreeItem(SYSTEM_PARAMETERS)
'    nGlobalParam = globalParamItem.GetItemPosition
'
'    'Global Parameters -> Fixed
'    Set tmpTreeItem = New TreeAccess
'    tmpTreeItem.AddName "Test", tvwChild, nGlobalParam
'    tmpTreeItem.AddHeading "Name", 25
'    tmpTreeItem.AddHeading "Description", 40
'    tmpTreeItem.AddHeading "Start Date", 15
'    tmpTreeItem.AddHeading "Amount", 12
'    tmpTreeItem.AddHeading "Percentage", 13
'    tmpTreeItem.AddHeading "Maximum", 10
'    tmpTreeItem.AddHeading "Boolean", 10
'    tmpTreeItem.AddHeading "String", 15
'    colTreeData.Add tmpTreeItem, "Test"
'
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

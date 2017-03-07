VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TreeViewCS
' Description   : Contains client specific code for handling Add New functionality.
' Change history
' Prog      Date        Description
' DJP       19/02/02    SYS4142 Created
' JR        27/05/02    SYS4537 Added UpdateTreeNodeText
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupTreeValues
' Description   :   Client specific version of the Core method. Creates client entries on the
'                   TreeView.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetupTreeValues(ByRef colTreeData As Collection)
    On Error GoTo Failed
    
    UpdateTreeNodeText colTreeData 'JR SYS4537
    CreateClientNodes colTreeData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CreateClientNodes
' Description   :   Sets up the treeview with all required nodes that need to
'                   be displayed. These are the Core treeview items that need to be displayed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateClientNodes(ByRef colTreeData As Collection)
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   HandleTreeFind
' Description   :   Given the key passed in, perform a Find on the associated data. This is the
'                   client specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleTreeFind(vKey As Variant, frmFind As Form, clsTableAccess As TableAccess)
     On Error GoTo Failed
     
     Exit Sub
Failed:
     g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   HandleTreeNew
' Description   :   Given the key passed in, set the form to be used by an edit on the listview.
'                   This is the client specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleTreeNew(ByVal vKey As Variant, ByRef frmNew As Form, ByRef enumReturn As MSGReturnCode)
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   HandleTreeViewPopup
' Description   :   Given the key passed in, set the popup for a right click on the TreeView.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleTreeViewPopup(sKey As String)
    On Error GoTo Failed
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'JR SYS4537
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   UpdateTreeNodeText
' Description   :   Allows Clients to re-label Tree Node Text whilst maintaining Key Value
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateTreeNodeText(colTreeData As Collection)
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'JR End

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnterpriseServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      EnterpriseServer.cls
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


Public Sub InitializeEnterpriseServerScope(ByVal ScopeItem As SnapInLib.IScopeItem)
    
    On Error GoTo ErrTrap_InitializeEnterpriseServerScope
    
    ScopeItem.ScopeNode.DisplayName = "Enterprise Server"
    'ScopeItem.Tag = // extra data - not currently required
    
    ' NOTE: These ColumnHeaders MUST match the ColumnHeaders defined
    ' in the ListView, otherwise they will not display.
    ScopeItem.ColumnHeaders.Add 1, COLUMNKEY_ENTERPRISESERVERNAME, COLUMNDISPLAY_ENTERPRISESERVERNAME
    ScopeItem.ListSubItems.Add 1, COLUMNKEY_ENTERPRISESERVERNAME, "Enterprise Server"

    Exit Sub
    
' Error Handler for this method
ErrTrap_InitializeEnterpriseServerScope:
    m_SnapInData.Error.DisplayError "InitializeEnterpriseServerScope"
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


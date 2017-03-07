VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      UICommand.cls
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

Public Enum UICOMMANDREQUEST
    UICOMMANDREQUEST_NONE
    UICOMMANDREQUEST_CONNECTCOMPUTER
    UICOMMANDREQUEST_NEWMESSAGEQUEUE
    UICOMMANDREQUEST_STALLCOMPONENT
    UICOMMANDREQUEST_STARTCOMPONENT
    UICOMMANDREQUEST_STARTSERVICE
    UICOMMANDREQUEST_STOPSERVICE
    UICOMMANDREQUEST_STARTQUEUE
    UICOMMANDREQUEST_STALLQUEUE
    UICOMMANDREQUEST_NEWQUEUEEVENT
End Enum


' =============================================================================
' Method:       ProcessUICommand
' Type:         Subroutine
' Description:  Processes user input from toolbar, menu button, and context menus.
'
' Parameters:   Selection   MMCClipboard object containing the current selected
'                           items. The selection may be in either the scope pane
'                           or in the result pane.
'               Command     String containing key of command to execute
'
' Output:       None
' =============================================================================
'
Public Sub ProcessUICommand(ByVal Selection As SnapInLib.IMMCClipboard, _
                            ByVal Command As UICOMMANDREQUEST)

    On Error GoTo ErrTrap_ProcessUICommand
  
    Dim SelectionType As SnapInSelectionTypeConstants
    SelectionType = Selection.SelectionType

    ' Add UICommand to Key
    Select Case SelectionType
        Case siSingleScopeItem
            Selection.ScopeItems(1).Key = AddUICommandToKey(Selection.ScopeItems(1).Key, CStr(Command))
        Case siSingleListItem
            Selection.ListItems(1).Key = AddUICommandToKey(Selection.ListItems(1).Key, CStr(Command))
    End Select

    Select Case Command

        Case UICOMMANDREQUEST_CONNECTCOMPUTER
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.Computer.ConnectComputer Selection
            End If
        Case UICOMMANDREQUEST_NEWMESSAGEQUEUE
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.Queue.NewMessageQueue Selection.ScopeItems(1)
            End If
        Case UICOMMANDREQUEST_STALLCOMPONENT
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.StalledComponents.StallComponent Selection.ScopeItems(1)
            End If
        Case UICOMMANDREQUEST_STARTCOMPONENT
            If SelectionType = siSingleListItem Then
                m_SnapInData.StalledComponents.StartComponent Selection.ListItems(1)
            End If
        Case UICOMMANDREQUEST_STARTSERVICE
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.Computer.PutComputerMQLServiceStatus Selection.ScopeItems(1), SERVICESTATUS_STARTED
            End If
        Case UICOMMANDREQUEST_STOPSERVICE
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.Computer.PutComputerMQLServiceStatus Selection.ScopeItems(1), SERVICESTATUS_STOPPED
            End If
        Case UICOMMANDREQUEST_STARTQUEUE
            If SelectionType = siSingleListItem Then
                m_SnapInData.SnapIn.ConsoleMsgBox "Start Queue Clicked " + Selection.ListItems(1).Key
            End If
        Case UICOMMANDREQUEST_STALLQUEUE
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.SnapIn.ConsoleMsgBox "Stall Queue Clicked " + Selection.ScopeItems(1).Key
            End If
        Case UICOMMANDREQUEST_NEWQUEUEEVENT
            If SelectionType = siSingleScopeItem Then
                m_SnapInData.QueueEvents.NewQueueEvent Selection.ScopeItems(1)
            End If
        Case Else
            ' This should not occur. If it does, it is because the
            ' toolbar, menu button, or context menus have been modified and
            ' this select statement has not been updated to deal with it.
            ' This will catch the developer's attention without causing problems
            ' for any unwary end users.
            Debug.Assert False
    End Select

Exit_ProcessUICommand:
    
    ' Remove UICommand to Key to indicate that a UI command is no longer being processed
    Select Case SelectionType
        Case siSingleScopeItem
            Selection.ScopeItems(1).Key = RemoveUICommandFromKey(Selection.ScopeItems(1).Key)
        Case siSingleListItem
            Selection.ListItems(1).Key = RemoveUICommandFromKey(Selection.ListItems(1).Key)
    End Select
    
    Exit Sub

' Error Handler for this method
ErrTrap_ProcessUICommand:
    
    m_SnapInData.Error.DisplayError "ProcessUICommand"
    Resume Exit_ProcessUICommand ' Ensure UI Command is removed from the key
End Sub




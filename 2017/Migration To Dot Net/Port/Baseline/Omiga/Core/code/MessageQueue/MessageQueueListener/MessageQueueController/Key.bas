Attribute VB_Name = "Key"
'Workfile:      Key.bas
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

Const DELIMITER = ","
Const DELIMITERQUEUE = DELIMITER + "Queue="
Const DELIMITERCOMPUTER = DELIMITER + "Computer="
Const DELIMITERPROGID = DELIMITER + "ProgID="
Const DELIMITEREVENTKEY = DELIMITER + "EventKey="
Const DELIMITERUICOMMAND = DELIMITER + "UICommand="
Const DELIMITERNTLOGEVENTFILE = DELIMITER + "NTLogEventFile="
Const DELIMITERNTLOGEVENTRECORD = DELIMITER + "NTLogEventRecord="
Const DELIMITEREND = DELIMITER

' Auto-create nodes
Const KEY_STATICNODE = "Static Node"
Const KEY_ENTERPRISESERVER = "EnterpriseServer"

' Other nodes
Const PREFIX_KEY_COMPUTER = "PrefixKeyComputer" + DELIMITER
Const PREFIX_KEY_QUEUESFOLDER = "PrefixKeyQueuesFolder" + DELIMITER
Const PREFIX_KEY_QUEUE = "PrefixKeyQueue" + DELIMITER
Const PREFIX_KEY_STALLEDCOMPONENTS = "PrefixKeyStalledComponents" + DELIMITER
Const PREFIX_KEY_STALLEDCOMPONENT = "PrefixKeyStalledComponent" + DELIMITER
Const PREFIX_KEY_QUEUEEVENTS = "PrefixKeyQueueEvents" + DELIMITER
Const PREFIX_KEY_QUEUEEVENT = "PrefixKeyQueueEvent" + DELIMITER
Const PREFIX_KEY_NTLOGEVENTS = "PrefixKeyNTLogEvents" + DELIMITER
Const PREFIX_KEY_NTLOGEVENT = "PrefixKeyNTLogEvent" + DELIMITER

Public Function IsStaticNodeKey(ByVal strKey As String) As Boolean
    If strKey = KEY_STATICNODE Then
        IsStaticNodeKey = True
    Else
        IsStaticNodeKey = False
    End If
End Function
 
Public Function GetEnterpriseServerKey() As String
    GetEnterpriseServerKey = KEY_ENTERPRISESERVER
End Function

Public Function IsEnterpriseServerKey(ByVal strKey As String) As Boolean
    If strKey = KEY_ENTERPRISESERVER Then
        IsEnterpriseServerKey = True
    Else
        IsEnterpriseServerKey = False
    End If
End Function

Public Function GetComputerKey(ByVal strComputerName As String) As String
    GetComputerKey = PREFIX_KEY_COMPUTER + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function

Public Function IsComputerKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_COMPUTER) = 1 Then
        IsComputerKey = True
    Else
        IsComputerKey = False
    End If
End Function

Public Function GetComputerFromKey(ByVal strKey As String) As String
    Dim nIndexDelimiterComputer
    nIndexDelimiterComputer = InStr(strKey, DELIMITERCOMPUTER)
    
    If nIndexDelimiterComputer > 0 Then
        Dim nIndexDelimiterNext
        nIndexDelimiterNext = InStr(nIndexDelimiterComputer + 1, strKey, DELIMITER)
        GetComputerFromKey = Mid(strKey, nIndexDelimiterComputer + Len(DELIMITERCOMPUTER), nIndexDelimiterNext - nIndexDelimiterComputer - Len(DELIMITERCOMPUTER))
    Else
        Err.Raise 0, "GetComputerFromKey", "Failed"
    End If
End Function

Public Function GetQueuesFolderKey(ByVal strComputerName As String) As String
    GetQueuesFolderKey = PREFIX_KEY_QUEUESFOLDER + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function

Public Function IsQueuesFolderKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_QUEUESFOLDER) = 1 Then
        IsQueuesFolderKey = True
    Else
        IsQueuesFolderKey = False
    End If
End Function
Public Function GetQueueKey(ByVal strComputerName As String, ByVal strQueueName As String) As String
    GetQueueKey = PREFIX_KEY_QUEUE + DELIMITERQUEUE + strQueueName + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function

Public Function IsQueueKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_QUEUE) = 1 Then
        IsQueueKey = True
    Else
        IsQueueKey = False
    End If
End Function

Public Function GetQueueFromKey(ByVal strKey As String) As String
    Dim nIndexDelimiterQueue
    nIndexDelimiterQueue = InStr(strKey, DELIMITERQUEUE)
    
    If nIndexDelimiterQueue > 0 Then
        Dim nIndexDelimiterNext
        nIndexDelimiterNext = InStr(nIndexDelimiterQueue + 1, strKey, DELIMITER)
        GetQueueFromKey = Mid(strKey, nIndexDelimiterQueue + Len(DELIMITERQUEUE), nIndexDelimiterNext - nIndexDelimiterQueue - Len(DELIMITERQUEUE))
    Else
        Err.Raise 0, "GetQueueFromKey", "Failed"
    End If
End Function

Public Function GetStalledComponentsKey(ByVal strComputerName As String, ByVal strQueueName As String) As String
    GetStalledComponentsKey = PREFIX_KEY_STALLEDCOMPONENTS + DELIMITERQUEUE + strQueueName + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function

Public Function IsStalledComponentsKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_STALLEDCOMPONENTS) = 1 Then
        IsStalledComponentsKey = True
    Else
        IsStalledComponentsKey = False
    End If
End Function

Public Function GetStalledComponentKey(ByVal strComputerName As String, ByVal strQueueName As String, ByVal strProgID As String) As String
    GetStalledComponentKey = PREFIX_KEY_STALLEDCOMPONENT + DELIMITERQUEUE + strQueueName + DELIMITERCOMPUTER + strComputerName + DELIMITERPROGID + strProgID + DELIMITEREND
End Function

Public Function IsStalledComponentKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_STALLEDCOMPONENT) = 1 Then
        IsStalledComponentKey = True
    Else
        IsStalledComponentKey = False
    End If
End Function

Public Function IsSameStalledComponentsKey(ByVal strKey1 As String, ByVal strKey2 As String) As Boolean
    If IsStalledComponentsKey(strKey1) And _
        IsStalledComponentsKey(strKey2) And _
        GetComputerFromKey(strKey1) = GetComputerFromKey(strKey2) And _
        GetQueueFromKey(strKey1) = GetQueueFromKey(strKey2) Then
        IsSameStalledComponentsKey = True
    Else
        IsSameStalledComponentsKey = False
    End If
End Function

Public Function IsChildStalledComponentKey(ByVal strKeyStalledComponents As String, ByVal strKeyStalledComponent As String) As Boolean
    If IsStalledComponentsKey(strKeyStalledComponents) And _
        IsStalledComponentKey(strKeyStalledComponent) And _
        GetComputerFromKey(strKeyStalledComponents) = GetComputerFromKey(strKeyStalledComponent) And _
        GetQueueFromKey(strKeyStalledComponents) = GetQueueFromKey(strKeyStalledComponent) Then
        IsChildStalledComponentKey = True
    Else
        IsChildStalledComponentKey = False
    End If
End Function

Public Function GetProgIDFromKey(ByVal strKey As String) As String
    Dim nIndexDelimiterProgID
    nIndexDelimiterProgID = InStr(strKey, DELIMITERPROGID)
    
    If nIndexDelimiterProgID > 0 Then
        Dim nIndexDelimiterNext
        nIndexDelimiterNext = InStr(nIndexDelimiterProgID + 1, strKey, DELIMITER)
        GetProgIDFromKey = Mid(strKey, nIndexDelimiterProgID + Len(DELIMITERPROGID), nIndexDelimiterNext - nIndexDelimiterProgID - Len(DELIMITERPROGID))
    Else
        Err.Raise 0, "GetProgIDFromKey", "Failed"
    End If
End Function

Public Function GetQueueEventsKey(ByVal strComputerName As String, ByVal strQueueName As String) As String
    GetQueueEventsKey = PREFIX_KEY_QUEUEEVENTS + DELIMITERQUEUE + strQueueName + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function

Public Function IsQueueEventsKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_QUEUEEVENTS) = 1 Then
        IsQueueEventsKey = True
    Else
        IsQueueEventsKey = False
    End If
End Function

Public Function IsSameQueueEventsKey(ByVal strKey1 As String, ByVal strKey2 As String) As Boolean
    If IsQueueEventsKey(strKey1) And _
        IsQueueEventsKey(strKey2) And _
        GetComputerFromKey(strKey1) = GetComputerFromKey(strKey2) And _
        GetQueueFromKey(strKey1) = GetQueueFromKey(strKey2) Then
        IsSameQueueEventsKey = True
    Else
        IsSameQueueEventsKey = False
    End If
End Function

Public Function GetQueueEventKey(ByVal strComputerName As String, ByVal strQueueName As String, ByVal strKey As String) As String
    GetQueueEventKey = PREFIX_KEY_QUEUEEVENT + DELIMITERQUEUE + strQueueName + DELIMITERCOMPUTER + strComputerName + DELIMITEREVENTKEY + strKey + DELIMITEREND
End Function

Public Function IsQueueEventKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_QUEUEEVENT) = 1 Then
        IsQueueEventKey = True
    Else
        IsQueueEventKey = False
    End If
End Function

Public Function AddUICommandToKey(ByRef strKeyIn As String, ByVal strUICommand As String) As String
     AddUICommandToKey = strKeyIn + DELIMITERUICOMMAND + strUICommand + DELIMITEREND
End Function

Public Function RemoveUICommandFromKey(strKeyIn As String) As String
    Dim nIndexDelimiterUICommand
    nIndexDelimiterUICommand = InStr(strKeyIn, DELIMITERUICOMMAND)
    
    If nIndexDelimiterUICommand > 0 Then
        RemoveUICommandFromKey = Left(strKeyIn, nIndexDelimiterUICommand - 1)
    Else
        RemoveUICommandFromKey = strKeyIn
    End If
End Function

Public Function KeyHasUICommand(strKey As String) As Boolean
    If InStr(strKey, DELIMITERUICOMMAND) = 1 Then
        KeyHasUICommand = True
    Else
        KeyHasUICommand = False
    End If
End Function

Public Function GetUICommandFromKey(strKey As String) As UICOMMANDREQUEST
    Dim nIndexDelimiterUICommand
    nIndexDelimiterUICommand = InStr(strKey, DELIMITERUICOMMAND)
    
    If nIndexDelimiterUICommand > 0 Then
        Dim nIndexDelimiterNext
        nIndexDelimiterNext = InStr(nIndexDelimiterUICommand + 1, strKey, DELIMITER)
        GetUICommandFromKey = Mid(strKey, nIndexDelimiterUICommand + Len(DELIMITERUICOMMAND), nIndexDelimiterNext - nIndexDelimiterUICommand - Len(DELIMITERUICOMMAND))
    Else
        GetUICommandFromKey = UICOMMANDREQUEST_NONE
    End If
End Function

Public Function GetNTLogEventsKey(ByVal strComputerName As String) As String
    GetNTLogEventsKey = PREFIX_KEY_NTLOGEVENTS + DELIMITERCOMPUTER + strComputerName + DELIMITEREND
End Function
Public Function IsNTLogEventsKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_NTLOGEVENTS) = 1 Then
        IsNTLogEventsKey = True
    Else
        IsNTLogEventsKey = False
    End If
End Function

Public Function GetNTLogEventKey(ByVal strComputerName As String, strNTLogFile As String, strNTLogEventRecord As String) As String
    GetNTLogEventKey = PREFIX_KEY_NTLOGEVENT + DELIMITERCOMPUTER + strComputerName + DELIMITERNTLOGEVENTFILE + strNTLogFile + DELIMITERNTLOGEVENTRECORD + strNTLogEventRecord + DELIMITEREND
End Function

Public Function IsNTLogEventKey(ByVal strKey As String) As Boolean
    If InStr(strKey, PREFIX_KEY_NTLOGEVENT) = 1 Then
        IsNTLogEventKey = True
    Else
        IsNTLogEventKey = False
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datawpgNewQEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      datawpgNewQueueEvent.cls
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
Public m_ScopeItem As SnapInLib.IScopeItem
Public m_strComputerName As String
Public m_strQueueName As String

Public m_strQueueEventType As String
Public m_nQueueEventThreads As Integer
Public m_strQueueEventDay As String
Public m_nQueueEventHour As Integer
Public m_nQueueEventMinute As Integer
Public m_nQueueEventSecond As Integer

Public m_strKey As String

' same order of pages as added to the property sheet (starting with 1)
Public Enum PAGENEWQUEUEEVENT
    PAGENEWQUEUEEVENT_ETYPE = 1
    PAGENEWQUEUEEVENT_ETHREADS = 2
    PAGENEWQUEUEEVENT_EDAYTIME = 3
    PAGENEWQUEUEEVENT_EFINISH = 4
End Enum


Public Function GetNextPage(CurrentPage As PAGENEWQUEUEEVENT) As PAGENEWQUEUEEVENT
    
    GetNextPage = CurrentPage
    
    Select Case CurrentPage
        Case PAGENEWQUEUEEVENT_ETYPE
            Select Case m_strQueueEventType
                Case QUEUEEVENTTYPE_SETTHREADS
                    GetNextPage = PAGENEWQUEUEEVENT_ETHREADS
                Case Else
                    GetNextPage = PAGENEWQUEUEEVENT_EDAYTIME
            End Select
        Case PAGENEWQUEUEEVENT_ETHREADS
            GetNextPage = PAGENEWQUEUEEVENT_EDAYTIME
        Case PAGENEWQUEUEEVENT_EDAYTIME
            GetNextPage = PAGENEWQUEUEEVENT_EFINISH
        Case PAGENEWQUEUEEVENT_EFINISH
        Case Else
            Debug.Assert 0
    End Select
End Function


Public Function GetPreviousPage(CurrentPage As PAGENEWQUEUEEVENT) As PAGENEWQUEUEEVENT
    
    GetPreviousPage = CurrentPage
    
    Select Case CurrentPage
        Case PAGENEWQUEUEEVENT_ETYPE
            GetPreviousPage = PAGENEWQUEUEEVENT_ETYPE
        Case PAGENEWQUEUEEVENT_ETHREADS
            GetPreviousPage = PAGENEWQUEUEEVENT_ETYPE
        Case PAGENEWQUEUEEVENT_EDAYTIME
            Select Case m_strQueueEventType
                Case QUEUEEVENTTYPE_SETTHREADS
                    GetPreviousPage = PAGENEWQUEUEEVENT_ETHREADS
                Case Else
                    GetPreviousPage = PAGENEWQUEUEEVENT_ETYPE
            End Select
        Case PAGENEWQUEUEEVENT_EFINISH
            GetPreviousPage = PAGENEWQUEUEEVENT_EDAYTIME
        Case Else
            Debug.Assert 0
    End Select
End Function





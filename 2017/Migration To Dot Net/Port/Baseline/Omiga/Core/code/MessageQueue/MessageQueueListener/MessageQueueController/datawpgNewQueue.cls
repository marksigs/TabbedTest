VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datawpgNewQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      datawpgNewQueue.cls
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

Public m_strQueueType As String
Public m_strQueueName As String
Public m_nQueueThreads As Integer
Public m_strConnectionString As String
Public m_nPollingInterval As Integer

' same order of pages as added to the property sheet (starting with 1)
Public Enum PAGENEWQUEUE
    PAGENEWQUEUE_QTYPE = 1
    PAGENEWQUEUE_QNAME = 2
    PAGENEWQUEUE_QTHREADS = 3
    PAGENEWQUEUE_QCONNECT = 4
    PAGENEWQUEUE_QPOLLING = 5
    PAGENEWQUEUE_QFINISH = 6
End Enum

Public Function GetNextPage(CurrentPage As PAGENEWQUEUE) As PAGENEWQUEUE
    
    GetNextPage = CurrentPage
    
    Select Case CurrentPage
        Case PAGENEWQUEUE_QTYPE
            GetNextPage = PAGENEWQUEUE_QNAME
        Case PAGENEWQUEUE_QNAME
            GetNextPage = PAGENEWQUEUE_QTHREADS
        Case PAGENEWQUEUE_QTHREADS
            Select Case m_strQueueType
                Case "MSMQ1"
                    GetNextPage = PAGENEWQUEUE_QFINISH
                Case "OMMQ1"
                    GetNextPage = PAGENEWQUEUE_QCONNECT
                Case Else
                    Debug.Assert 0
            End Select
        Case PAGENEWQUEUE_QCONNECT
            Select Case m_strQueueType
                Case "MSMQ1"
                Case "OMMQ1"
                    GetNextPage = PAGENEWQUEUE_QPOLLING
                Case Else
                    Debug.Assert 0
            End Select
        Case PAGENEWQUEUE_QPOLLING
            Select Case m_strQueueType
                Case "MSMQ1"
                Case "OMMQ1"
                    GetNextPage = PAGENEWQUEUE_QFINISH
                Case Else
                    Debug.Assert 0
            End Select
        Case PAGENEWQUEUE_QFINISH
            Debug.Assert 0
        Case Else
            Debug.Assert 0
    End Select
End Function


Public Function GetPreviousPage(CurrentPage As PAGENEWQUEUE) As PAGENEWQUEUE
    
    GetPreviousPage = CurrentPage
    
    Select Case CurrentPage
        Case PAGENEWQUEUE_QTYPE
        Case PAGENEWQUEUE_QNAME
            GetPreviousPage = PAGENEWQUEUE_QTYPE
        Case PAGENEWQUEUE_QTHREADS
            GetPreviousPage = PAGENEWQUEUE_QNAME
        Case PAGENEWQUEUE_QCONNECT
            Select Case m_strQueueType
                Case "MSMQ1"
                Case "OMMQ1"
                    GetPreviousPage = PAGENEWQUEUE_QTHREADS
                Case Else
                    Debug.Assert 0
            End Select
        Case PAGENEWQUEUE_QPOLLING
            Select Case m_strQueueType
                Case "MSMQ1"
                Case "OMMQ1"
                    GetPreviousPage = PAGENEWQUEUE_QCONNECT
                Case Else
                    Debug.Assert 0
            End Select
        Case PAGENEWQUEUE_QFINISH
            Select Case m_strQueueType
                Case "MSMQ1"
                    GetPreviousPage = PAGENEWQUEUE_QTHREADS
                Case "OMMQ1"
                    GetPreviousPage = PAGENEWQUEUE_QPOLLING
                Case Else
                    Debug.Assert 0
            End Select
        Case Else
            Debug.Assert 0
    End Select
End Function




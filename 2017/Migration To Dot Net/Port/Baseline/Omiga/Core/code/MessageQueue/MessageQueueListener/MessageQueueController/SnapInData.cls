VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnapInData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SnapInData.cls
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

' snap in back references (instance specific)
Private m_Controller1 As StaticNode
Private m_SnapIn As ISnapIn
Private m_ScopeItems As IScopeItems

' snap in data (instance specific)
Private m_Computer As Computer
Private m_ConfigData As ConfigData
Private m_datappgES As datappgES
Private m_datawpgNewQueue As datawpgNewQueue
Private m_datawpgNewQEvent As datawpgNewQEvent
Private m_datawpgStallComp As datawpgStallComp
Private m_EnterpriseServer As EnterpriseServer
Private m_Error As Error
Private m_MQL As MQL
Private m_NTLogEvents As NTLogEvents
Private m_Queue As Queue
Private m_QueueEvents As QueueEvents
Private m_QueuesFolder As QueuesFolder
Private m_Service As Service
Private m_StalledComponents As StallComp
Private m_UICommand As UICommand
Private m_WMI As WMI
    
Public Function Controller1() As StaticNode
    Set Controller1 = m_Controller1
End Function

Public Function SnapIn() As ISnapIn
    Set SnapIn = m_SnapIn
End Function
    
Public Function ScopeItems() As IScopeItems
    Set ScopeItems = m_ScopeItems
End Function
    
Public Function Computer() As Computer
    Set Computer = m_Computer
End Function

Public Function ConfigData() As ConfigData
    Set ConfigData = m_ConfigData
End Function
    
Public Function datappgES() As datappgES
    Set datappgES = m_datappgES
End Function
    
Public Function datawpgNewQueue() As datawpgNewQueue
    Set datawpgNewQueue = m_datawpgNewQueue
End Function
    
Public Function datawpgNewQEvent() As datawpgNewQEvent
    Set datawpgNewQEvent = m_datawpgNewQEvent
End Function
    
Public Function datawpgStallComp() As datawpgStallComp
    Set datawpgStallComp = m_datawpgStallComp
End Function

Public Function EnterpriseServer() As EnterpriseServer
    Set EnterpriseServer = m_EnterpriseServer
End Function

Public Function Error() As Error
    Set Error = m_Error
End Function

Public Function MQL() As MQL
    Set MQL = m_MQL
End Function
    
Public Function NTLogEvents() As NTLogEvents
    Set NTLogEvents = m_NTLogEvents
End Function
    
Public Function Queue() As Queue
    Set Queue = m_Queue
End Function
    
Public Function QueueEvents() As QueueEvents
    Set QueueEvents = m_QueueEvents
End Function

Public Function QueuesFolder() As QueuesFolder
    Set QueuesFolder = m_QueuesFolder
End Function
    
Public Function Service() As Service
    Set Service = m_Service
End Function

Public Function StalledComponents() As StallComp
    Set StalledComponents = m_StalledComponents
End Function

Public Function UICommand() As UICommand
    Set UICommand = m_UICommand
End Function

Public Function WMI() As WMI
    Set WMI = m_WMI
End Function

Public Sub Load(ByVal SnapIn As StaticNode, ByVal ScopeItems As ScopeItems)

    Set m_Controller1 = SnapIn
    Set m_SnapIn = SnapIn
    Set m_ScopeItems = ScopeItems
   
    Set m_Computer = New Computer
    Set m_ConfigData = New ConfigData
    Set m_datappgES = New datappgES
    Set m_datawpgNewQueue = New datawpgNewQueue
    Set m_datawpgNewQEvent = New datawpgNewQEvent
    Set m_datawpgStallComp = New datawpgStallComp
    Set m_EnterpriseServer = New EnterpriseServer
    Set m_Error = New Error
    Set m_MQL = New MQL
    Set m_NTLogEvents = New NTLogEvents
    Set m_Queue = New Queue
    Set m_QueueEvents = New QueueEvents
    Set m_QueuesFolder = New QueuesFolder
    Set m_Service = New Service
    Set m_StalledComponents = New StallComp
    Set m_UICommand = New UICommand
    Set m_WMI = New WMI
    
    Set m_Computer.m_SnapInData = Me
    Set m_ConfigData.m_SnapInData = Me
    Set m_datappgES.m_SnapInData = Me
    Set m_datawpgNewQueue.m_SnapInData = Me
    Set m_datawpgNewQEvent.m_SnapInData = Me
    Set m_datawpgStallComp.m_SnapInData = Me
    Set m_EnterpriseServer.m_SnapInData = Me
    Set m_Error.m_SnapInData = Me
    Set m_MQL.m_SnapInData = Me
    Set m_NTLogEvents.m_SnapInData = Me
    Set m_Queue.m_SnapInData = Me
    Set m_QueueEvents.m_SnapInData = Me
    Set m_QueuesFolder.m_SnapInData = Me
    Set m_Service.m_SnapInData = Me
    Set m_StalledComponents.m_SnapInData = Me
    Set m_UICommand.m_SnapInData = Me
    Set m_WMI.m_SnapInData = Me
    
End Sub

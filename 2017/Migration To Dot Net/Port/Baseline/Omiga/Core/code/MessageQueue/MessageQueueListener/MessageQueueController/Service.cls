VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Service.cls
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

Public Enum SERVICESTATUS
    SERVICESTATUS_STARTED
    SERVICESTATUS_STOPPED
    SERVICESTATUS_NOTFOUND
    SERVICESTATUS_UNKNOWN
End Enum

Const MQLSERVICENAME = "MessageQueueListener"

Function GetServiceStatus(ByVal strComputerName As String, ByVal strServiceName As String) As SERVICESTATUS
    
    On Error GoTo ErrTrap_GetServiceStatus
    
    GetServiceStatus = SERVICESTATUS_UNKNOWN
    
    Dim WbemServices As SWbemServices
    Set WbemServices = m_SnapInData.WMI.WMIGetWbemServices(strComputerName, "root/cimv2")
    Dim WbemObjectSet As SWbemObjectSet
    Set WbemObjectSet = WbemServices.ExecQuery("SELECT * FROM Win32_Service WHERE Name = '" + strServiceName + "'")
    
    Dim WbemObject As SWbemObject
    If WbemObjectSet.Count = 0 Then
        GetServiceStatus = SERVICESTATUS_NOTFOUND
    Else
        For Each WbemObject In WbemObjectSet
            If WbemObject.Started = True Then
                GetServiceStatus = SERVICESTATUS_STARTED
            Else
                GetServiceStatus = SERVICESTATUS_STOPPED
            End If
        Next
    End If

Exit_GetServiceStatus:
    Set WbemObjectSet = Nothing
    Set WbemServices = Nothing
    
    Exit Function

' Error Handler for this method
ErrTrap_GetServiceStatus:
    GoTo Exit_GetServiceStatus
End Function

Public Sub PutServiceStatus(ByVal strComputerName As String, ByVal strServiceName As String, ByVal RequestedStatus As SERVICESTATUS)
    
    On Error GoTo ErrTrap_PutServiceStatus
    
    Dim WbemServices As SWbemServices
    Set WbemServices = m_SnapInData.WMI.WMIGetWbemServices(strComputerName, "root/cimv2")
    Dim WbemObjectSet As SWbemObjectSet
    Set WbemObjectSet = WbemServices.ExecQuery("SELECT * FROM Win32_Service WHERE Name = '" + strServiceName + "'")
    
    Dim WbemObject As SWbemObject
    Dim outParam As Long

    For Each WbemObject In WbemObjectSet
        Select Case RequestedStatus
            Case SERVICESTATUS_STARTED
                outParam = WbemObject.StartService
                Select Case outParam
                    Case 0
                        ' success
                    Case 1
                        StopServiceError strComputerName, strServiceName, "The request is not supported."
                    Case Else
                        StopServiceError strComputerName, strServiceName, "Code (" + CStr(outParam) + "). Failed"
                End Select
            Case SERVICESTATUS_STOPPED
                outParam = WbemObject.StopService
                Select Case outParam
                    Case 0
                        ' success
                    Case 1
                        StartServiceError strComputerName, strServiceName, "The request is not supported."
                    Case 2
                        StartServiceError strComputerName, strServiceName, "The user did not have the necessary access."
                    Case 3
                        StartServiceError strComputerName, strServiceName, "The service cannot be stopped because other services that are running are dependent on it."
                    Case 4
                        StartServiceError strComputerName, strServiceName, "The requested control code is not valid, or it is unacceptable to the service."
                    Case 5
                        StartServiceError strComputerName, strServiceName, "The requested control code cannot be sent to the service because the state of the service (Win32_BaseService:State) is equal to 0, 1, or 2."
                    Case 6
                        StartServiceError strComputerName, strServiceName, "The service has not been started."
                    Case 7
                        StartServiceError strComputerName, strServiceName, "The service did not respond to the start request in a timely fashion."
                    Case 8
                        StartServiceError strComputerName, strServiceName, "Unknown failure when starting the service."
                    Case 9
                        StartServiceError strComputerName, strServiceName, "The directory path to the service executable was not found."
                    Case 10
                        StartServiceError strComputerName, strServiceName, "The service is already running."
                    Case 11
                        StartServiceError strComputerName, strServiceName, "The database to add a new service is locked."
                    Case 12
                        StartServiceError strComputerName, strServiceName, "A dependency for which this service relies on has been removed from the system."
                    Case 13
                        StartServiceError strComputerName, strServiceName, "The service failed to find the service needed from a dependent service."
                    Case 14
                        StartServiceError strComputerName, strServiceName, "The service has been disabled from the system."
                    Case 15
                        StartServiceError strComputerName, strServiceName, "The service does not have the correct authentication to run on the system."
                    Case 16
                        StartServiceError strComputerName, strServiceName, "This service is being removed from the system."
                    Case 17
                        StartServiceError strComputerName, strServiceName, "There is no execution thread for the service."
                    Case 18
                        StartServiceError strComputerName, strServiceName, "There are circular dependencies when starting the service."
                    Case 19
                        StartServiceError strComputerName, strServiceName, "There is a service running under the same name."
                    Case 20
                        StartServiceError strComputerName, strServiceName, "There are invalid characters in the name of the service."
                    Case 21
                        StartServiceError strComputerName, strServiceName, "Invalid parameters have been passed to the service."
                    Case 22
                        StartServiceError strComputerName, strServiceName, "The account which this service is to run under is either invalid or lacks the permissions to run the service."
                    Case 23
                        StartServiceError strComputerName, strServiceName, "The service exists in the database of services available from the system."
                    Case 24
                        StartServiceError strComputerName, strServiceName, "The service is currently paused in the system."
                    Case Else
                        StartServiceError strComputerName, strServiceName, "Code (" + CStr(outParam) + "). Failed"
                End Select
        End Select
    Next

Exit_PutServiceStatus:
    Set WbemObjectSet = Nothing
    Set WbemServices = Nothing
    
    Exit Sub

' Error Handler for this method
ErrTrap_PutServiceStatus:
    GoTo Exit_PutServiceStatus
End Sub

Private Sub StartServiceError(ByVal strComputerName As String, ByVal strServiceName As String, ByVal strMessage As String)
    m_SnapInData.SnapIn.ConsoleMsgBox "Error starting " + strServiceName + " Service on computer " + strComputerName + ".  " + strMessage
End Sub

Private Sub StopServiceError(ByVal strComputerName As String, ByVal strServiceName As String, ByVal strMessage As String)
    m_SnapInData.SnapIn.ConsoleMsgBox "Error stopping " + strServiceName + " Service on computer " + strComputerName + ".  " + strMessage
End Sub



Attribute VB_Name = "Logging"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        : Logging.bas
' Description   : Contains code to log events in Supervisor
' Change history
' Prog      Date        Description
' DJP       10/03/03    Created for BM0316.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Data
Private m_fso As FileSystemObject
Private m_txtStream As TextStream
Private m_sContext As String

' Constants
Public Const PROMOTION_LOG_FILE As String = "Promotion Log.txt"
Public Function OpenLogging(Optional tMode As IOMode = ForAppending, Optional sContext As String) As Boolean
#If LOGGING_ENABLED Then
    On Error GoTo Failed
    Dim sPath As String
    Dim bCreate As Boolean
    Dim bSuccess As Boolean
    
    bCreate = True
    bSuccess = True
    
    If tMode = ForReading Then
        bCreate = False
    End If
    
    Set m_fso = New FileSystemObject
    sPath = m_fso.BuildPath(App.Path, PROMOTION_LOG_FILE)
    Set m_txtStream = m_fso.OpenTextFile(sPath, tMode, bCreate)

    If Len(sContext) > 0 Then
        m_sContext = sContext
    End If

    OpenLogging = bSuccess

    Exit Function
Failed:
    If tMode = ForReading Then
        Err.Raise Err.Number, Err.DESCRIPTION
    Else
        App.LogEvent "OpenLogging: " & Err.DESCRIPTION, vbLogEventTypeWarning
    End If
#End If
End Function
Public Sub CloseLogging()
#If LOGGING_ENABLED Then
    On Error GoTo Failed

    If Not m_txtStream Is Nothing Then
        m_txtStream.Close
    End If
    
    Set m_fso = Nothing
    Set m_txtStream = Nothing

    Exit Sub
Failed:
#End If
End Sub
Public Sub WriteLogEvent(sData As String)
#If LOGGING_ENABLED Then
    On Error GoTo Failed
    Dim sTxt As String
    
    sTxt = m_sContext & " " & Now & ": " & sData
    
    If m_txtStream Is Nothing Then
        OpenLogging
    End If
    
    If Not m_txtStream Is Nothing Then
        m_txtStream.Write sTxt & vbNewLine
    End If
    
    Exit Sub
Failed:
    App.LogEvent "WriteLogEvent: " & Err.DESCRIPTION, vbLogEventTypeWarning
#End If
End Sub
Public Sub WriteLogKeys(clsToLog As TableAccess)
#If LOGGING_ENABLED Then
    On Error GoTo Failed
    Dim sLog As String
    Dim vFld As Variant
    Dim sKey As String
    Dim nIndex As Integer
    Dim colMatchFields As Collection
    Dim colMatchValues As Collection
    
    Set colMatchFields = clsToLog.GetKeyMatchFields
    Set colMatchValues = clsToLog.GetKeyMatchValues
    
    If Not colMatchValues Is Nothing And Not colMatchFields Is Nothing Then
        If colMatchValues.Count = colMatchFields.Count Then
            For nIndex = 1 To colMatchValues.Count
                vFld = colMatchValues(nIndex)
                If TypeName(vFld) = "Byte()" Then
                    vFld = g_clsSQLAssistSP.GuidToString(CStr(vFld))
                End If
                sLog = sLog & colMatchFields(nIndex) & " = '" & vFld & "' "
            Next
        End If
    End If
    
    WriteLogEvent sLog
    
    Exit Sub
Failed:
#End If
End Sub
Public Sub SetLoggingContext(sContext As String)
#If LOGGING_ENABLED Then
    m_sContext = sContext
#End If
End Sub
Public Function GetStream() As TextStream
    Set GetStream = m_txtStream
End Function

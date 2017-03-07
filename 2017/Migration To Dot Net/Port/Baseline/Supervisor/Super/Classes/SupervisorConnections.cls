VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupervisorConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : SupervisorConnection
' Description   : Class that contains connection details for a particular database connection.
'                 Will be stored internally by Supervisor as part of a collection of database
'                 connections for easy retrieval.
' Change history
' Prog      Date        Description
' DJP       7/05/01     Added Database Server
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Change History
' RF        29/11/05    MAR734 Supervisor should cope with integrated security
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private m_sConnectionString As String
Private m_Conn As ADODB.Connection
Private m_Cmd As ADODB.Command
Private m_sDatabaseName As String
Private m_sAppServer As String
Private m_sUserID As String
Private m_sPassword As String
Private m_bTransactionActive As Boolean
Private m_sProvider As String
Private m_bIsActive As Boolean
Private m_sDatabaseServer As String
Private m_sDatabaseType As String
Private Sub Class_Initialize()
    Set m_Conn = New ADODB.Connection
    Set m_Cmd = New ADODB.Command
    m_bTransactionActive = False
End Sub
Public Function GetDatabaseType() As String
    GetDatabaseType = m_sDatabaseType
End Function
Public Function GetDatabaseServer() As String
    GetDatabaseServer = m_sDatabaseServer
End Function
Public Function GetAppServer() As String
    GetAppServer = m_sAppServer
End Function
Public Function GetTransactionActive() As Boolean
    GetTransactionActive = m_bTransactionActive
End Function
Public Sub SetTransactionActive(Optional bActive As Boolean = True)
    m_bTransactionActive = bActive
End Sub
Public Function GetDatabaseKey() As String
    GetDatabaseKey = GetConnectionKey(m_sDatabaseName, m_sUserID)
End Function
Public Function GetIsActive() As Boolean
    GetIsActive = m_bIsActive
End Function
Public Function GetDatabaseName() As String
    On Error GoTo Failed
    
    If Len(m_sDatabaseName) > 0 Then
        GetDatabaseName = m_sDatabaseName
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetDataBaseName: Database Name is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
Public Function GetProvider() As String
    On Error GoTo Failed
    If Len(m_sProvider) > 0 Then
        GetProvider = m_sProvider
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetProvider: Provider is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetPassword() As String
    On Error GoTo Failed
    ' RF 29/11/05 MAR734 Start
    'If Len(m_sPassword) > 0 Then
    '    GetPassword = m_sPassword
    'Else
    '    g_clsErrorHandling.RaiseError errGeneralError, "GetPassword: Password is empty"
    'End If
    GetPassword = m_sPassword
    ' RF 29/11/05 MAR734 End
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetUserID() As String
    On Error GoTo Failed
    ' RF 29/11/05 MAR734 Start
    'If Len(m_sUserID) > 0 Then
    '    GetUserID = m_sUserID
    'Else
    '    g_clsErrorHandling.RaiseError errGeneralError, "GetUserID: UserID is empty"
    'End If
    GetUserID = m_sUserID
    ' RF 29/11/05 MAR734 End

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub SetProvider(sProvider As String)
    If Len(sProvider) > 0 Then
        m_sProvider = sProvider
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetProvider: Provider is empty"
    End If
End Sub
Public Sub SetDatabaseType(sType As String)
    m_sDatabaseType = sType
End Sub
Public Sub SetIsActive(bIsActive As Boolean)
    m_bIsActive = bIsActive
End Sub
Public Sub SetDatabaseName(sName As String)
    m_sDatabaseName = sName
End Sub
Public Sub SetUserID(sUserID As String)
    m_sUserID = sUserID
End Sub
Public Sub SetDatabaseServer(sDatbaseServer As String)
    m_sDatabaseServer = sDatbaseServer
End Sub
Public Sub SetAppServer(sAppServer As String)
    m_sAppServer = sAppServer
End Sub
Public Sub SetPassword(sPassword As String)
    On Error GoTo Failed
    ' RF 29/11/05 MAR734 Start
    'If Len(sPassword) > 0 Then
    '    m_sPassword = sPassword
    'Else
    '    g_clsErrorHandling.RaiseError errGeneralError, "SetPassword - password is empty"
    'End If
    m_sPassword = sPassword
    ' RF 29/11/05 MAR734 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetConnectionState() As ObjectStateEnum
    GetConnectionState = m_Conn.State
End Function
Public Function GetConnection() As ADODB.Connection
    On Error GoTo Failed
    
    If Not m_Conn Is Nothing Then
        Set GetConnection = m_Conn
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetConnection: Connection is empty"
    End If

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetCommand() As ADODB.Command
    On Error GoTo Failed
    Dim Param As ADODB.Parameter
    
    Set Param = New ADODB.Parameter
    
    If Not m_Cmd Is Nothing Then
        ' Default command execution parameters
        m_Cmd.CommandType = adCmdUnknown
        m_Cmd.CommandText = ""
        
        Do While m_Cmd.Parameters.Count <> 0
            m_Cmd.Parameters.Delete m_Cmd.Parameters.Count - 1
        Loop
        
        Set GetCommand = m_Cmd
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetCommand: Command is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub SetConnectionString(sConnectionstring)
    m_sConnectionString = sConnectionstring
End Sub
Public Function GetConnectionString() As String
    On Error GoTo Failed
    
    If Len(m_sConnectionString) > 0 Then
        GetConnectionString = m_sConnectionString
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetConnectionString: Connection string is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : DatabaseOptions
' Description   : Support class for the Database Options screen.
' Change history
' Prog      Date        Description
' DJP       08/06/2001  Created for SQL Server port
' DJP       13/02/2002  SYS4052 Remove message informing user promotions won't be recorded if user
'                       has not logged in.
' RF        28/11/05    MAR353 Windows authentication fails when app server not set
' RF        29/11/05    MAR734 Supervisor should cope with integrated security
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateConnections
' Description   : Reads the list of connections from the registry and stores them internally for
'                 future use.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateConnections() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bFound As Boolean
    Dim bActiveConnection As Boolean
    Dim bActiveConnectionIsValid As Boolean
    Dim nThisEntry As Integer
    Dim sKey As String
    Dim sActive As String
    Dim sUserID As String
    Dim sPassword As String
    Dim sProvider As String
    Dim sAppServer As String
    Dim sConnection As String
    Dim sServiceName As String
    Dim sDatabaseType As String
    Dim sDatabaseServer As String
    Dim clsConnection As SupervisorConnection
    
    bRet = True
    bFound = True
    nThisEntry = 1
    BeginWaitCursor
    
    bActiveConnection = False
    bActiveConnectionIsValid = False
    Do
        sServiceName = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_SERVICE_NAME)
        
        If Len(sServiceName) > 0 Then
            Set clsConnection = New SupervisorConnection
            
            sActive = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_ACTIVE)
            sUserID = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_USER_ID)
            sConnection = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_CONNECTION_STRING)
            sServiceName = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_SERVICE_NAME)
            sPassword = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_PASSWORD)
            sAppServer = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_APP_SERVER)
            sProvider = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_PROVIDER)
            
            sDatabaseType = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_DATABASE_TYPE)
                        
            If Len(sDatabaseType) = 0 Then
                ' If no database type, default to Oracle
                sDatabaseType = DATABASE_ORACLE
            End If

            sDatabaseServer = QueryValue(HKEY_LOCAL_MACHINE, REG_CONNECTION & nThisEntry, REG_DATABASE_SERVER)
            
            clsConnection.SetConnectionString sConnection
            clsConnection.SetDatabaseName sServiceName
            clsConnection.SetUserID sUserID
            clsConnection.SetProvider sProvider
            clsConnection.SetPassword sPassword
            clsConnection.SetAppServer sAppServer
            clsConnection.SetIsActive CBool(sActive = "*")
            clsConnection.SetDatabaseServer sDatabaseServer
            clsConnection.SetDatabaseType sDatabaseType
            
            sKey = GetConnectionKey(sServiceName, sUserID)
            
            g_clsDataAccess.AddConnection clsConnection, sKey
            
            If Len(sActive) > 0 Then
                ' This is the active connection
                g_clsDataAccess.SetActiveConnection sKey
                
                g_clsMainSupport.SetDatabaseCaption
                bActiveConnection = True
                
                ' RF 28/11/05 MAR353 Start
                ' RF 29/11/05 MAR734 Start
                'If Len(sUserID) > 0 And _
                '    Len(sConnection) > 0 And _
                '    Len(sServiceName) > 0 And _
                '    Len(sPassword) > 0 And _
                '    Len(sAppServer) > 0 And _
                '    Len(sProvider) > 0 Then
                '    bActiveConnectionIsValid = True
                'End If
                If Len(sConnection) > 0 And _
                    Len(sServiceName) > 0 And _
                    Len(sAppServer) > 0 And _
                    Len(sProvider) > 0 Then
                    bActiveConnectionIsValid = True
                End If
                ' RF 28/11/05 MAR353 End
                ' RF 29/11/05 MAR734 End
            End If
            
            nThisEntry = nThisEntry + 1
        
        Else
            bFound = False
        End If
    
    Loop While bFound
    
    Dim nRet As Integer
    EndWaitCursor
    
    If nThisEntry = 1 Then
        nRet = MsgBox("No database connections exist. Would you like to enter them now?", vbYesNo + vbQuestion)
    ElseIf bActiveConnection = False Then
        nRet = MsgBox("No active database connections exist. Would you like to set one active now?", vbYesNo + vbQuestion)
    ' RF 28/11/05 MAR353 Start
    ElseIf bActiveConnectionIsValid = False Then
        nRet = MsgBox("No valid active database connections exist. Would you like to create one now?", vbYesNo + vbQuestion)
    ' RF 28/11/05 MAR353 End
    End If
    
    If nRet = vbYes Then
        Dim nReturn As MSGReturnCode
        frmDatabaseOptions.Show vbModal, frmMain
        
        nReturn = frmDatabaseOptions.GetReturnCode()
        Unload frmDatabaseOptions
        
        If nReturn = MSGSuccess Then
            g_bFirstTimeExecuted = True
        Else
            bRet = False
        End If
    ElseIf bActiveConnection = False Then
        bRet = False
        
    End If
    
    CreateConnections = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveConnection
' Description   : Saves the connection identified by nConnection to the registry, and to the dataaccess
'                 Supervisorc class for use during this Supervisor session.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveConnection(colLine As Collection, nActiveConnection As Integer, nConnection As Integer, bActive As Boolean)
    On Error GoTo Failed
    Dim sUserID As String
    Dim sActive As String
    Dim sService As String
    Dim sPassword As String
    Dim sProvider As String
    Dim sAppServer As String
    Dim sConnection As String
    Dim sDatabaseType As String
    Dim sConnectionKey As String
    Dim sDatabaseServer As String
    'Dim colLine As Collection
    Dim connDetails As ConnectionDetails
    Dim clsConnection As SupervisorConnection
    
    sActive = ""
    
    'Set colLine = lvConnections.GetLine(nConnection)
    
    If Not colLine Is Nothing Then
        'If colLine.Count = PROVIDER + 2 Then ' + 2 because the text box starts at o, plus the Active field
            
        ' Get connection details
        sService = colLine(COL_SERVICE_NAME)
        sUserID = colLine(COL_USER_ID)
        sPassword = colLine(COL_PASSWORD)
        sProvider = colLine(COL_PROVIDER)
        sAppServer = colLine(COL_APP_SERVER)
        sDatabaseType = colLine(COL_DATABASE_TYPE)
        sDatabaseServer = colLine(COL_DATABASE_SERVER)
        
        ' If it's the active one, store the fact
        If bActive Then
            sActive = "*"
        End If
        
        ' Save the connection info into the registry
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_SERVICE_NAME, sService, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_USER_ID, sUserID, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_PASSWORD, sPassword, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_PROVIDER, sProvider, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_APP_SERVER, sAppServer, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_ACTIVE, sActive, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_DATABASE_TYPE, sDatabaseType, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_DATABASE_SERVER, sDatabaseServer, REG_SZ
                
        ' Need to save this connection to the DataAccess class for future use in this Supervisor session
        GetConnectionDetails colLine, connDetails
        sConnection = BuildConnectionString(connDetails)

        Set clsConnection = New SupervisorConnection
        clsConnection.SetConnectionString sConnection
        clsConnection.SetUserID sUserID
        clsConnection.SetDatabaseName sService
        clsConnection.SetProvider sProvider
        clsConnection.SetPassword sPassword
        clsConnection.SetAppServer sAppServer
        clsConnection.SetDatabaseServer sDatabaseServer
        clsConnection.SetDatabaseType sDatabaseType
        
        sConnectionKey = GetConnectionKey(sService, sUserID)
        
        SetKeyValue HKEY_LOCAL_MACHINE, REG_CONNECTION & nConnection, REG_CONNECTION_STRING, sConnection, REG_SZ
        
        ' Add it to the DataAccess class
        g_clsDataAccess.AddConnection clsConnection, sConnectionKey
    
        If Len(sActive) > 0 Then
            ' Update the main view to reflect the new connection.
            g_clsDataAccess.SetActiveConnection sConnectionKey
            g_clsMainSupport.SetDatabaseCaption
        End If
        'End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetConnectionDetails
' Description   : Populates the class connDetails with the connection details of the current
'                 active connection if nIndex is missing, or the row identified by nIndex if
'                 it isn't.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetConnectionDetails(colLine As Collection, connDetails As ConnectionDetails)
    On Error GoTo Failed
    Dim nItemIndex As Integer
    Dim lstItem As ListItem
    
'   ' If nIndex isn't there, just get the active connection
'    If IsMissing(nIndex) Then
'        Set lstItem = lvConnections.SelectedItem
'
'        If Not lstItem Is Nothing Then
'            nItemIndex = lstItem.Index
'        Else
'            g_clsErrorHandling.RaiseError errGeneralError, "GetConnectionDetails: No selected items"
'        End If
'    Else
'        nItemIndex = nIndex
'    End If
        
    'Set colLine = lvConnections.GetLine(nItemIndex)

    If Not colLine Is Nothing Then
        connDetails.sService = colLine(COL_SERVICE_NAME)
        connDetails.sUserID = colLine(COL_USER_ID)
        connDetails.sPassword = colLine(COL_PASSWORD)
        connDetails.sProvider = colLine(COL_PROVIDER)
        connDetails.sAppServer = colLine(COL_APP_SERVER)
        connDetails.sDatabaseServer = colLine(COL_DATABASE_SERVER)
        connDetails.sDatabaseType = colLine(COL_DATABASE_TYPE)

    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetConnectionDetails: Row from listbox is empty"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : BuildConnectionString
' Description   : Builds the provider specific connection string to be used to open a connection
'                 to the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function BuildConnectionString(connDetails As ConnectionDetails)
    On Error GoTo Failed
    Dim sUID As String
    Dim sServer As String
    Dim sPassword As String
    Dim sDataSource As String
    Dim sConnection As String
    
    sConnection = "Provider=" & connDetails.sProvider
    
    Select Case connDetails.sDatabaseType
    
    Case DATABASE_SQL_SERVER
        sUID = "UID="
        sPassword = "PWD="
        sDataSource = "DATABASE="
        sServer = "SERVER="
        
    ' default to Oracle
    Case Else
        sUID = "USER ID="
        sPassword = "PASSWORD="
        sDataSource = "DATA SOURCE="
    End Select
    
    ' RF 29/11/05 MAR734 Start - enable integrated security
    With connDetails
        If (Len(.sUserID) = 0 And Len(.sPassword) = 0) Then
        
            ' use integrated security
            sConnection = sConnection & "; " & sDataSource & .sService & ";" & _
                    "Integrated Security=SSPI"
        
        Else
        
            ' use native SQL server connection
            sConnection = sConnection & "; " & sDataSource & .sService & ";" & _
                    sUID & .sUserID & "; " & sPassword & .sPassword
        
        End If
        
        If Len(sServer) > 0 Then
            sConnection = sConnection & "; " & sServer & connDetails.sDatabaseServer
        End If
    End With
    ' RF 29/11/05 MAR734 End

    BuildConnectionString = sConnection
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function



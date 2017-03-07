VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : DataAccess
' Description   : Contains functions associated with database access. This class
'                 maintains a list of all currently defined databases and manages
'                 switches between them
' Who?  When?       What?
' DJP   29/11/00    Phase 2 Task Management
' DJP   09/06/01    SQL Server port
' DJP   21/06/01    Moved GenerateSearchCriteria from form processing to here
' DJP   27/06/01    SQL Server port for Intermediaries.
' DJP   04/08/01    Added logging to aid debugging of errors
' DJP   03/10/01    GetAllTablesName and GetAllTablesFieldName need the database name to be passed
'                   through
' DJP   22/10/01    SYS2831 Fix bug where sometimes a string would not get quotes in WHERE clause
' DJP   20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Added NextNumber
' STB   10/01/02    SYS3225 Products could not be keyed with starttime.
' STB   08/03/02    SYS4244 Similar to above, but both Date and String data-types now covered.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMids specific history
' Who?  When?       What?
' DJP   24/02/03    BM0318 DoesFieldExist doesn't need to bring back any rows to work, so force a zero
'                   row return (to aid performance)
' DJP   05/03/03    BM0423 Add ValidateProduct stored procedure
' GHun  28/03/2003  BM0425 Avoid creating unnecessary locks when accessing reading only data
' PJO   28/11/2005  MAR81 Added stored procedure USP_GetNextSequenceNumber to list
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS specific history
' RF        28/11/05    MAR353 Windows authentication fails when app server not set
' PSC       06/03/2006  MAR1298 Add GetTablePrimaryKeys
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom specific history
' PB        23/10/2006  EP2_13  Have extensively modified GetTableData() so that it
'                       (1) will advise if more than 1000 records returned
'                       (2) now accepts optional parameter enumClassOption so that in cases
'                           where a table represents multiple entity types, the required type
'                           can be specified.
' TW        17/10/2006  EP2_15 Processing for Association/Club
' TW        08/12/2006  EP2_154 The change made by EP2_13 for advising on more than 1000 records introduced errors in some circumstances
' TW        12/12/2006  EP2_20 Further corrections for advising on more than 1000 records introduced errors in some circumstances
' TW        21/03/2007  EP2_1224 - Run-time error if combo valuename entered with apostrophe
' TW        04/04/2007  EP2_2263 - Getting error message when trying to edit/view a mortgage product.
' TW        07/04/2007  EP2_2307 - Unable to edit Mortgage Products
' TW        25/04/2007  EP2_2559 - Count problems when editing a Stage Task
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_colConnections As Collection
Private m_colTableExists As Collection
Private m_sActiveConnectionKey As String
Private m_sActiveSQL As String
Private m_bSaveSQL As Boolean

' Enums
Public Enum StoredProcNames
    spGetPanelDetails = 1
    spGetUsers
    spGetComboWithValidation
    spGetNextConditionsRef
    spGetNextQuestionRef
    spGetSelectedStages
    spGetNextProcurationRef
    spGetNextTargetID
    spGetNextSplitID
    spValidateProduct
    spGetNextSequenceNumber     'PJO 28/11/2005 MAR81
    spGetTablePrimaryKeys       ' PSC 06/03/2006 MAR1298
End Enum
Private Function CheckRecordCount(ByRef sDatabaseName As String, ByRef sSearch As String) As Boolean
Dim arrFrom

Dim intOrderByPos As Integer
Dim lngRecordsAffected As Long
Dim C As Integer

Dim rs As ADODB.Recordset
Dim strSQL As String
Dim strWork As String

    On Error GoTo Failed:
    
' 1 If the SQL contains UNION or DISTINCT, don't try and sort out the count
' TW 25/04/2007 EP2_2559
'    If InStr(1, UCase$(sSearch), " UNION ") > 0 Then
    If InStr(1, UCase$(sSearch), " UNION ") > 0 Or InStr(1, UCase$(sSearch), " DISTINCT ") > 0 Then
' TW 25/04/2007 EP2_2559 End
        CheckRecordCount = True
        Exit Function
    End If
    
' 2 Ensure all separators have spaces around them
    strWork = Replace(sSearch, "(", " ( ")
    strWork = Replace(strWork, ")", " ) ")
    strWork = Replace(strWork, "=", " = ")
' TW 07/04/2007 EP2_2307 Correction for *= and =* constructs
    strWork = Replace(strWork, "* =", " *=")
    strWork = Replace(strWork, "= *", "=* ")
' TW 07/04/2007 EP2_2307 End
    strWork = Trim$(strWork)
    
' 3 Split the SQL into FROM clauses
    arrFrom = Split(strWork, " FROM ", -1, vbTextCompare)
    
' 4 Sort out which FROM clause (if more than one) is the Primary
    For C = UBound(arrFrom) To 0 Step -1
        arrFrom(C) = " FROM " & arrFrom(C)
        If C > 0 Then
            If InStr(1, UCase$(arrFrom(C - 1)), " SELECT ") = 0 Then
                Exit For
            End If
        End If
    Next C

' 5 Eliminate the field names
    For C = C - 1 To 0 Step -1
        arrFrom(C) = ""
    Next C
    strWork = Join(arrFrom, " ")

' 6 Reconstruct the SQL as a SELECT COUNT
    strSQL = "SELECT COUNT(1) " & strWork
        
' 7 Check for ORDER BY clause and remove if present
    intOrderByPos = InStr(UCase(strSQL), " ORDER BY")
    If intOrderByPos > 0 Then
        strSQL = Left(strSQL, intOrderByPos - 1)
    End If
    
    Set rs = g_clsDataAccess.ExecuteCommand(strSQL)
    
    lngRecordsAffected = CLng(rs(0).Value)
    rs.Close
    
    If lngRecordsAffected > 1000 Then ' Should really be a constant
        If MsgBox("A large number of records was found (" & lngRecordsAffected & ")." & vbCrLf & _
                "Do you want to list these?" & _
                "(Click 'No' to cancel the Find operation)", vbQuestion + vbYesNo, _
                "Additional filtering required") = vbNo Then
            g_blnDoNotValidateRS = True
            Exit Function
        End If
    End If
    CheckRecordCount = True
    Exit Function
Failed:
    CheckRecordCount = False
    g_clsErrorHandling.RaiseError errGeneralError, "CheckRecordCount: SQL : " & sSearch & vbCrLf & "Error :" + Err.DESCRIPTION

End Function

Private Function DealWithApostrophesInSQL(sCommand As String) As String
' TW 21/03/2007 EP2_1224 - New function to deal with apostrophes within data fields in SQL (eg XXX = 'Acct's')
Dim strTemp As String
Dim strWork As String
Dim X As Integer
Dim arrReplace
Dim arrReplaceWith

    DealWithApostrophesInSQL = sCommand ' Default is to return the input string unmodified

' 1 Check the SQL to see if there are any apostrophes
    If InStr(1, sCommand, "'") = 0 Then
        Exit Function   ' No apostrophes
    End If

' TW 04/04/2007 EP2_2263
    strWork = Replace(sCommand, "'""", "' """) ' Ensure there is a space between "'
' TW 04/04/2007 EP2_2263 End

' 2 Set up arrays of replacements to effect the handling of embedded apostrophes in field data
' TW 02/04/2007 EP2_1224
'    arrReplace = Array("'AND", "'OR", " '", ",'", "('", "' ", "',", "')", "'", "`")
'    arrReplaceWith = Array("' AND", "' OR", " `", ",`", "(`", "` ", "`,", "`)", "''", "'")
' TW 09/04/2007 EP2_2318
'    arrReplace = Array("'AND", "'OR", " '", ",'", "('", "' ", "',", "')", "''", "'", "`")
'    arrReplaceWith = Array("' AND", "' OR", " `", ",`", "(`", "` ", "`,", "`)", "``", "''", "'")
    arrReplace = Array("'WHERE", "'LEFT", "'RIGHT", "'OUTER", "'INNER", "'AND", "'OR", " '", ",'", "('", "' ", "',", "')", "''", "'", "`")
    arrReplaceWith = Array("' WHERE", "' LEFT", "' RIGHT", "' OUTER", "' INNER", "' AND", "' OR", " `", ",`", "(`", "` ", "`,", "`)", "``", "''", "'")
' TW 09/04/2007 EP2_2318 End


' 3 Do the replace
'    strWork = Trim$(sCommand) & " " 'Ensure there is a terminating space to satisfy all "arrReplace" values
' TW 04/04/2007 EP2_2263
'    strWork = Trim$(Replace(sCommand, "=", " = ")) & " " 'Ensure there is a terminating space to satisfy all "arrReplace" values
' TW 07/04/2007 EP2_2307 Correction for spaces around =
' This was necessary because of the old *= and =* constructions which
' must not have a sace between the * and = characters
'    strWork = Trim$(Replace(strWork, "=", " = ")) & " " 'Ensure there is a terminating space to satisfy all "arrReplace" values
    X = 1
    Do
        X = InStr(X, strWork, "=")
        If X = 0 Then
            Exit Do
        End If
        strTemp = Mid$(strWork, X - 1, 1)
        Select Case strTemp
            Case " ", "*"
            Case Else
                strWork = Mid$(strWork, 1, X - 1) & " " & Mid$(strWork, X)
                X = X + 1
        End Select
    
        strTemp = Mid$(strWork, X + 1, 1)
        Select Case strTemp
            Case " ", "*"
            Case Else
                strWork = Mid$(strWork, 1, X) & " " & Mid$(strWork, X + 1)
        End Select
        X = X + 1
    Loop
    strWork = Trim$(strWork) & " "
    DealWithApostrophesInSQL = strWork
' TW 07/04/2007 EP2_2307 End
' TW 04/04/2007 EP2_2263 End
' TW 02/04/2007 EP2_1224 End
    
    For X = 0 To UBound(arrReplace, 1)
        strWork = Replace(strWork, arrReplace(X), arrReplaceWith(X), 1, -1, vbTextCompare)
' TW 09/04/2007 EP2_2318
'        If X = 7 Then
        If X = 12 Then
' TW 09/04/2007 EP2_2318 End
            If InStr(1, strWork, "'") = 0 Then
                Exit Function   ' No embedded apostrophes
            End If
        End If
    Next X

    DealWithApostrophesInSQL = strWork
    Exit Function

End Function

' Return a list of the currently defined databases
Public Function GetConnectionList() As Collection
    On Error GoTo Failed
    
    If Not m_colConnections Is Nothing Then
        If m_colConnections.Count > 0 Then
            Set GetConnectionList = m_colConnections
        End If
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' Returns the state of the current database - open, closed etc.
Public Function GetConnectionState(Optional sDatabaseName As String) As ObjectStateEnum
    Dim clsConn As SupervisorConnection
    GetSupervisorConnection clsConn, sDatabaseName
    
    GetConnectionState = clsConn.GetConnectionState()
End Function
' Returns the name of the current database
Public Function GetDatabaseName(Optional sKey As String)
    On Error GoTo Failed
    Dim clsConnection As SupervisorConnection
    
    GetSupervisorConnection clsConnection, sKey
    
    GetDatabaseName = clsConnection.GetDatabaseName()

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetDatabaseUserID(Optional sKey As String)
    On Error GoTo Failed
    Dim clsConnection As SupervisorConnection
    
    GetSupervisorConnection clsConnection, sKey
    
    GetDatabaseUserID = clsConnection.GetUserID()

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetConnectionKey() As String
    On Error GoTo Failed
    
    If Len(m_sActiveConnectionKey) > 0 Then
        GetConnectionKey = m_sActiveConnectionKey
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetActiveConnection: Key is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Sub AddConnection(superConn As SupervisorConnection, sConnectionKey As String)
' RF 28/11/05 MAR353 Start - Ensure connections are updated
'    On Error Resume Next
'
'    ' Don't care if it's there or not
'    On Error Resume Next
'    m_colConnections.Add superConn, sConnectionKey
On Error GoTo Failed
    
    If Not m_colConnections Is Nothing Then
        If m_colConnections.Count <> 0 Then
            Dim clsConnection As SupervisorConnection
            ' search will error if item not found
            On Error Resume Next
            Set clsConnection = m_colConnections(sConnectionKey)
            On Error GoTo Failed
            If Not clsConnection Is Nothing Then
                m_colConnections.Remove sConnectionKey
            End If
        End If
    End If
    
    m_colConnections.Add superConn, sConnectionKey
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, "AddConnection: " & Err.DESCRIPTION
' RF 28/11/05 MAR353 End
    
    Exit Sub
End Sub
Public Function GetActiveDatabaseName() As String
    On Error GoTo Failed
    GetActiveDatabaseName = GetConnectionKey()
    Exit Function
Failed:

End Function
Public Sub GetSupervisorConnection(clsConnection As SupervisorConnection, Optional sKey As String)
    On Error GoTo Failed
    Dim sConnectionKey As String
        
    If Len(sKey) > 0 Then
        sConnectionKey = sKey
    Else
        sConnectionKey = GetConnectionKey()
    End If
    
    On Error Resume Next
    Set clsConnection = m_colConnections(sConnectionKey)
    
    If Err.Number <> 0 Or clsConnection Is Nothing Then
        On Error GoTo Failed
        g_clsErrorHandling.RaiseError errGeneralError, "Supervisor Connection Class is empty"
    End If
    
    Err.Clear
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "GetSupervisorConnection: " & Err.DESCRIPTION
End Sub
Private Function GetCommand(Optional sDatabaseName As String) As ADODB.Command
    On Error GoTo Failed
    Dim clsConnection As SupervisorConnection
    Dim cmd As ADODB.Command
    Dim sKey As String
    
    If Len(sDatabaseName) > 0 Then
        sKey = sDatabaseName
    Else
        sKey = GetConnectionKey()
    End If
        
    Set clsConnection = m_colConnections(sKey)

    If Not clsConnection Is Nothing Then
        Set cmd = clsConnection.GetCommand()
    
        If Not cmd Is Nothing Then
            Set GetCommand = cmd
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Command returned is empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Connection class returned is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError errGeneralError, " Unable to get command: " & Err.DESCRIPTION
End Function
Public Function GetActiveCommand() As ADODB.Command
    On Error GoTo Failed
    
    Set GetActiveCommand = GetCommand()
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError errGeneralError, " Unable to get command: " & Err.DESCRIPTION
End Function
Private Function GetConnection(Optional sDatabaseKey As String) As ADODB.Connection
    On Error GoTo Failed
    
    Dim clsConnection As SupervisorConnection
    Dim conn As ADODB.Connection
    Dim sKey As String
    
    If Len(sDatabaseKey) > 0 Then
        sKey = sDatabaseKey
    Else
        sKey = GetConnectionKey()
    End If
    
    On Error Resume Next
    Set clsConnection = m_colConnections(sKey)

    If Err.Number = 0 Then
        On Error GoTo Failed
        If Not clsConnection Is Nothing Then
            Set conn = clsConnection.GetConnection()
        
            If Not conn Is Nothing Then
                Set GetConnection = conn
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "Connection returned is empty"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Connection Class returned is empty"
        End If
    Else
        On Error GoTo Failed
        Set conn = Nothing
    End If
    
    Set GetConnection = conn
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError errGeneralError, " Unable to add connection: " & Err.DESCRIPTION

End Function
Public Function GetActiveConnection() As ADODB.Connection
    On Error GoTo Failed
    
    Set GetActiveConnection = GetConnection()
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError errGeneralError, " Unable to add connection: " & Err.DESCRIPTION
End Function
Friend Function SetActiveConnection(sConnectionKey As String) As Boolean
    On Error GoTo Failed
    Dim clsSupervisorConnection As SupervisorConnection
    Dim bActive As Boolean
    
    bActive = False
    
    If Len(sConnectionKey) > 0 Then
        On Error Resume Next
        Set clsSupervisorConnection = m_colConnections(m_sActiveConnectionKey)
        
        If Err.Number = 0 Then
            If Not clsSupervisorConnection Is Nothing Then
                clsSupervisorConnection.SetIsActive False
            End If
        End If
        
        Err.Clear
        
        Set clsSupervisorConnection = m_colConnections(sConnectionKey)
        
        If Err.Number = 0 Then
            m_sActiveConnectionKey = sConnectionKey
            clsSupervisorConnection.SetIsActive True
            'g_clsVersion.CheckVersioning
            bActive = True
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to set connection - connection is empty"
    End If
    
    SetActiveConnection = bActive
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Sub ClearConnections()
    Set m_colConnections = New Collection
End Sub
Friend Sub BeginTrans(Optional sDatabaseKey As String)
    On Error GoTo Failed
    Dim adoConn As ADODB.Connection
    Dim clsSuperConn As SupervisorConnection
    Dim bTransactionActive As Boolean
    
    GetSupervisorConnection clsSuperConn, sDatabaseKey
    
    If Not clsSuperConn Is Nothing Then
        Set adoConn = clsSuperConn.GetConnection
        
        If Not adoConn Is Nothing Then
            If adoConn.State = adStateClosed Then
                OpenDatabase sDatabaseKey
            End If

            bTransactionActive = clsSuperConn.GetTransactionActive()
            
            If Not bTransactionActive Then
                adoConn.BeginTrans
                clsSuperConn.SetTransactionActive
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
Friend Sub CommitAll(Optional bCommit As Boolean = True)
    On Error GoTo Failed
    Dim sDatabase As String
    Dim enumState As ObjectStateEnum
    Dim adoConn As ADODB.Connection
    Dim superConn As SupervisorConnection
    Dim bTransactionActive As Boolean
    
    For Each superConn In m_colConnections
        
        Set adoConn = superConn.GetConnection()
        sDatabase = superConn.GetDatabaseKey()
        
        If Not adoConn Is Nothing Then
            enumState = adoConn.State
            bTransactionActive = superConn.GetTransactionActive()
            
            If enumState = adStateOpen And bTransactionActive Then
                If bCommit Then
                    CommitTrans sDatabase
                Else
                    RollbackTrans sDatabase
                End If
            End If
        End If
    Next

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub CommitTrans(Optional sDatabaseKey As String, Optional bClose As Boolean = True)
    On Error GoTo Failed
    Dim adoConn As ADODB.Connection
    Dim clsSuperConn As SupervisorConnection
    
    GetSupervisorConnection clsSuperConn, sDatabaseKey
    
    If Not clsSuperConn Is Nothing Then
        clsSuperConn.SetTransactionActive False
        
        Set adoConn = clsSuperConn.GetConnection()
        
        If Not adoConn Is Nothing Then
            If adoConn.State = adStateOpen Then
                adoConn.CommitTrans
                
                If bClose Then
                    adoConn.Close
                End If
            End If
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub RollbackTrans(Optional sDatabaseKey As String)
    On Error GoTo Failed
    Dim adoConn As ADODB.Connection
    Dim clsSuperConn As SupervisorConnection
    
    GetSupervisorConnection clsSuperConn, sDatabaseKey
    
    If Not clsSuperConn Is Nothing Then
        clsSuperConn.SetTransactionActive False
        
        Set adoConn = clsSuperConn.GetConnection

        If Not adoConn Is Nothing Then
            If adoConn.State = adStateOpen Then

                adoConn.RollbackTrans
                         
                adoConn.Close
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub CloseConnection(Optional sDatabaseKey As String)
    On Error GoTo Failed
    Dim adoConn As ADODB.Connection
    Dim clsSuperConn As SupervisorConnection
    
    On Error Resume Next
    GetSupervisorConnection clsSuperConn, sDatabaseKey
    
    If Err.Number = 0 Then
        On Error GoTo Failed
        If Not clsSuperConn Is Nothing Then
            Set adoConn = clsSuperConn.GetConnection()
            
            ' Don't allow close if we're in a transaction
            If clsSuperConn.GetTransactionActive() = False Then
                clsSuperConn.SetTransactionActive False
                If Not adoConn Is Nothing Then
                    If adoConn.State = adStateOpen Then
                        adoConn.Close
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub OpenDatabase(Optional sDatabaseKey As String)
    On Error GoTo Failed
    Dim adoConn As ADODB.Connection
    Dim adoCmd As ADODB.Command
    Dim clsConnection As SupervisorConnection
    
    GetSupervisorConnection clsConnection, sDatabaseKey
    
    Set adoConn = clsConnection.GetConnection()
    Set adoCmd = clsConnection.GetCommand()
    
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateClosed Then
            Dim sConn As String
    
            sConn = clsConnection.GetConnectionString()
            
            adoConn.Open sConn
            adoConn.CursorLocation = adUseClient
            
            Set adoCmd.ActiveConnection = adoConn
            clsConnection.SetTransactionActive False
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "Unable to open Connection: " + Err.DESCRIPTION
End Sub
Friend Function DoesFieldExist(sTable As String, sField As String, Optional sDatabaseKey As String, Optional bLog As Boolean = True) As Boolean
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sTableField As String
    Dim nThisField As Integer
    Dim sKey As String
    Dim bExists As Boolean
    Dim adoConn As ADODB.Connection
    Dim adoCmd As ADODB.Command
    On Error Resume Next
    
    If Len(sDatabaseKey) > 0 Then
        sKey = sDatabaseKey
    Else
        sKey = GetConnectionKey()
    End If
    
    sKey = sKey + sTable + sField
    bExists = m_colTableExists(sKey)

    If Err.Number <> 0 Then
        On Error GoTo Failed
    
        If DoesTableExist(sTable) Then
            Set rs = New ADODB.Recordset
            Set adoConn = GetActiveConnection()
            Set adoCmd = GetActiveCommand()
            If adoConn.State = adStateClosed Then
                OpenDatabase
            End If
            
            ' DJP BM0318 - force no rows to be returned
            adoCmd.CommandText = "SELECT * FROM " & sTable & " WHERE 1=2"
            
            rs.CursorLocation = adUseClient
            ' DJP BM0318 - only has to be readonly
            rs.Open adoCmd, , adOpenStatic, adLockReadOnly
            
            For nThisField = 0 To rs.fields.Count - 1
                sTableField = UCase(rs.fields(nThisField).Name)
                If sTableField = UCase(sField) Then
                    bExists = True
                End If
            Next
            rs.Close
        End If

        SetExistsState sKey, bExists
        
    End If

    If bExists = False And bLog Then
        If g_clsMainSupport.GetLogging Then
            App.LogEvent "Field " & sField & " not found in table " & sTable, vbLogEventTypeWarning
        End If
    End If
    
    DoesFieldExist = bExists
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function DoesTableExist(sTable As String, Optional sDatabaseKey As String, Optional bWriteToLog As Boolean = True)
    Dim bExists As Boolean
    Dim sKey As String
    Dim sCommand As String
    Dim sFieldName As String
    Dim sAllTables As String
    Dim sSynonym As String
    Dim rs As ADODB.Recordset
        
    On Error Resume Next
    
    If Len(sDatabaseKey) > 0 Then
        sKey = sDatabaseKey
    Else
        sKey = GetConnectionKey()
    End If
    
    sKey = sKey + sTable
    bExists = m_colTableExists(sKey)
        
    If Err.Number <> 0 Then
        On Error GoTo Failed
        ' DJP SQL Server port
        sAllTables = g_clsSQLAssistSP.GetAllTablesName(sDatabaseKey)
        sFieldName = g_clsSQLAssistSP.GetAllTablesFieldName(sDatabaseKey)
        sCommand = "SELECT COUNT(*) ""EXISTS"" FROM " & sAllTables & " WHERE " & sFieldName & " = " + g_clsSQLAssistSP.FormatString(UCase(sTable))
    
        Set rs = ExecuteCommand(sCommand, sDatabaseKey)
    
        If rs.RecordCount > 0 Then
            If CLng(rs("EXISTS")) > 0 Then
                bExists = True
            Else
                ' DJP Phase 2 Task Management - check for synonyms too...
                sSynonym = g_clsSQLAssistSP.GetAllSynonymsName(sDatabaseKey)
                
                sCommand = "SELECT COUNT(*) ""EXISTS"" FROM " & sSynonym & " WHERE " & sFieldName & " = " + g_clsSQLAssistSP.FormatString(UCase(sTable))
                
                Set rs = ExecuteCommand(sCommand, sDatabaseKey)
            
                If rs.RecordCount > 0 Then
                    If CLng(rs("EXISTS")) > 0 Then
                        bExists = True
                    Else
                        bExists = False
                    End If
                End If
            End If
        End If
        
        SetExistsState sKey, bExists
        rs.Close
    
    End If
    
    If bExists = False And bWriteToLog Then
        If g_clsMainSupport.GetLogging Then
            App.LogEvent "Table " & sTable & " not found", vbLogEventTypeWarning
        End If
    End If
    
    DoesTableExist = bExists
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'BM0425 Added optional parameter bReadOnly (default = false)
Friend Function GetTableData(Optional ByVal sTable As String = "", Optional ByVal sSearch As String = "", Optional ByVal sDatabaseName As String, Optional ByVal bUseExistingSQL As Boolean = False, Optional ByVal bReadOnly = False, Optional enumClassOption As ClassOption = OPTION_NONE) As ADODB.Recordset
    '
    ' Notes on the enumClassOption optional parameter
    '==========================================================================
    ' This is used because some tables represent more than one type of entity.
    ' If this is the case, specific SQL can be added to the usual query so that
    ' only the correct data is returned.
    '
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTmp As String
    Dim adoCmd As ADODB.Command
    
    'PB 19/10/2006 EP2_13
' TW 08/12/2006 EP2_154
'    Dim lngRecordsAffected As Long
'    Dim strSQL As String
'    Dim intOrderByPos As Integer
' TW 08/12/2006 EP2_154 End
    Dim intWherePos As Integer
    Dim strEntitySpecificSQL As String
    
    bRet = True
    On Error GoTo Failed
    
    Set adoCmd = GetCommand(sDatabaseName)
    
    Set rs = New ADODB.Recordset
    'BM0425 LockType already specified below
    'rs.LockType = adLockBatchOptimistic

    If bUseExistingSQL Then
        sSearch = m_sActiveSQL
    End If
    
    'PB
    If Len(sSearch) = 0 Then
        sSearch = "SELECT * FROM " & sTable
    End If
    
    ' Determine any entity specific SQL
    Select Case enumClassOption
        Case OPTION_NONE
            ' No additional SQL required
        Case OPTION_PRINCIPALFIRM_FIRMBROKER
            strEntitySpecificSQL = "(PACKAGERINDICATOR IS NULL OR PACKAGERINDICATOR=0)"
        Case OPTION_PRINCIPALFIRM_FIRMPACKAGER
            strEntitySpecificSQL = "(PACKAGERINDICATOR=1)"
        Case OPTION_ARFIRM_INDIVIDUALBROKER
            strEntitySpecificSQL = "(PACKAGERINDICATOR IS NULL OR PACKAGERINDICATOR=0)"
        Case OPTION_ARFIRM_INDIVIDUALPACKAGER
            strEntitySpecificSQL = "(PACKAGERINDICATOR=1)"
' TW 17/10/2006 EP2_15
        Case OPTION_CLUB_ASSOCIATION_CLUB
            strEntitySpecificSQL = "(PACKAGERINDICATOR=0)"
        Case OPTION_CLUB_ASSOCIATION_ASSOCIATION
            strEntitySpecificSQL = "(PACKAGERINDICATOR=1)"
' TW 17/10/2006 EP2_15 End
            
            
    End Select
    ' Just a table name, or an SQL statement?
    If InStr(sSearch, " ") = 0 Then
        ' Just the table
        sSearch = "SELECT * FROM " & sTable
        'PB 23/10/2006 EP2_13 Begin
        If strEntitySpecificSQL <> "" Then
            sSearch = sSearch & " WHERE " & strEntitySpecificSQL
        End If
    Else
        ' sSearch contains an SQL statement
        'PB 23/10/2006
        If strEntitySpecificSQL <> "" Then
            intWherePos = InStr(sSearch, " WHERE ")
            If intWherePos > 0 Then
                'There is already a WHERE clause so we must be careful to maintain SQL syntax
                sSearch = Left(sSearch, intWherePos + 6) & strEntitySpecificSQL & _
                    " AND " & Mid(sSearch, intWherePos + 6, Len(sSearch))
                '
                ' An example of how this should work is as follows:
                '
                ' Before:
                ' SELECT col1, col2, col3 FROM table
                ' WHERE col2=0 AND col5=99 ORDER BY col1
                '
                ' After:
                ' SELECT col1, col2, col3 FROM table
                ' WHERE [EntitySpecificSQL]
                ' AND col2=0 AND col5=99 ORDER BY col1
                '
            Else
                'No WHERE clause, so we can append our bit to the end.
                sSearch = sSearch & " WHERE " & strEntitySpecificSQL
            End If
        End If
    End If
    ' TW 08/12/2006 EP2_154
    '
    ' Count how many records would be returned
    '
'    If GetConnectionState() = adStateClosed Then
'        g_clsDataAccess.OpenDatabase sDatabaseName
'    End If
    '
    
'    strSQL = "SELECT COUNT(1)" & Mid(sSearch, InStr(UCase(sSearch), " FROM "), Len(sSearch))
'    ' Check for ORDER BY clause and remove if present (as this might break!)
'    intOrderByPos = InStr(UCase(strSQL), " ORDER BY")
'    If intOrderByPos > 0 Then
'        strSQL = Left(strSQL, intOrderByPos - 1)
'    End If
    
'    adoCmd.CommandText = strSQL
'    '
'    Set rs = adoCmd.Execute
'    lngRecordsAffected = CLng(rs(0).Value)
    
'    If lngRecordsAffected > 1000 Then ' Should really be a constant
'        ' Too many - cancel query
'        If MsgBox("A large number of records was found (" & lngRecordsAffected & ")." & vbCrLf & _
'                "Do you want to list these?" & _
'                "(Click 'No' to cancel the Find operation)", vbQuestion + vbYesNo, _
'                "Additional filtering required") = vbNo Then
'            rs.Close
'            g_blnDoNotValidateRS = True
'            GoTo Failed
'        End If
'    End If
    ' rs.Close
    'PB 19/10/2006 EP2_13 End
    
    If GetConnectionState() = adStateClosed Then
        g_clsDataAccess.OpenDatabase sDatabaseName
    End If
    
' TW 02/04/2007 EP2_1224
    sSearch = DealWithApostrophesInSQL(sSearch)
' TW 02/04/2007 EP2_1224 End
    
    If CheckRecordCount(sDatabaseName, sSearch) Then
    ' TW 08/12/2006 EP2_154 End
        
        adoCmd.CommandText = sSearch
    
        rs.CursorLocation = adUseClient
        'BM0425
        If bReadOnly Then
            rs.Open adoCmd, , adOpenForwardOnly, adLockReadOnly
        Else
            rs.Open adoCmd, , adOpenStatic, adLockBatchOptimistic
        End If
        'BM0425 End
        
        Set GetTableData = rs
        
        If m_bSaveSQL Then
            m_sActiveSQL = sSearch
            m_bSaveSQL = False
        End If
    ' TW 08/12/2006 EP2_154
    End If
    ' TW 08/12/2006 EP2_154 End
    
Failed:
    Set rs = Nothing
    Set adoCmd = Nothing

    If Err.Number <> 0 Then
        Set GetTableData = Nothing
        EndWaitCursor
        g_clsErrorHandling.RaiseError errGeneralError, "GetTableData: Unable to get data for table: " + sTable + " " + Err.DESCRIPTION
    End If
End Function

Public Sub SaveSearcSQL(bSaveSQL As Boolean)
    m_bSaveSQL = bSaveSQL
    m_sActiveSQL = ""
End Sub
' DJP SQL Server Port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ExecuteStoredProcRS
' Description   : Executes the stored procedure in sText and returns a recordset with the results.
'                 The collection col contains a list of ADO parameters to be passed to the stored
'                 procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExecuteStoredProcRS(nName As StoredProcNames, Optional col As Collection = Nothing) As ADODB.Recordset
    On Error GoTo Failed
    Dim sText As String
    Dim nParamCnt As Integer
    Dim rs As ADODB.Recordset
    Dim Param As ADODB.Parameter
    Dim adoCmd As ADODB.Command
    
    ' Get a connection
    If GetConnectionState = adStateClosed Then
        g_clsDataAccess.OpenDatabase
    End If
    
    sText = GetStoredProcName(nName)
    
    ' And command
    Set adoCmd = GetCommand()
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = sText
    
    'Set the input parameters
    If Not col Is Nothing Then
        For nParamCnt = 1 To col.Count
            Set Param = col(nParamCnt)
            adoCmd.Parameters.Append Param
        Next nParamCnt
    End If
    
    Set rs = adoCmd.Execute

    ' Rest the default
    adoCmd.CommandType = adCmdTable
    Set ExecuteStoredProcRS = rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' DJP SQL Server port, pass in enum of names not hard coded proc name
' DJP BM0423 add optional database connection as a parameter
Public Function ExecuteStoredProcedure(nName As StoredProcNames, col As Collection, Optional sDatabase As String) As Collection
    On Error GoTo Failed
    Dim sProcName As String
    Dim nParamCnt As Integer
    Dim Param As ADODB.Parameter
    Dim adoCmd As ADODB.Command
    
    ' If a database is passed in, close the current connection
    If Len(sDatabase) > 0 Then
        g_clsDataAccess.CloseConnection
    End If
    
    ' Open the database - if the database is closed, this will open it using the database passed in
    ' or the current one if not
    g_clsDataAccess.OpenDatabase sDatabase
    
    sProcName = GetStoredProcName(nName)
    
    Set adoCmd = GetCommand()
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = sProcName
    
    'Set the input parameters
    If Not col Is Nothing Then
        For nParamCnt = 1 To col.Count
            Set Param = col(nParamCnt)
            adoCmd.Parameters.Append Param
        Next nParamCnt
    End If
    
    adoCmd.Execute

    For Each Param In adoCmd.Parameters
        If Param.Direction = adParamOutput Then
            'Return the values if output types
            col.Add Param.Value
        End If
    Next
    
    ' Close the current connection
    g_clsDataAccess.CloseConnection sDatabase
    
    Set ExecuteStoredProcedure = col
    adoCmd.CommandType = adCmdTable
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function ExecuteCommand(sCommand As String, Optional sDatabaseName As String) As ADODB.Recordset
    On Error GoTo Failed
    Dim adoCmd As ADODB.Command
    
    If GetConnectionState(sDatabaseName) = adStateClosed Then
        g_clsDataAccess.OpenDatabase sDatabaseName
    End If
    
    Set adoCmd = GetCommand(sDatabaseName)
' TW 21/03/2007 EP2_1224
'    adoCmd.CommandText = sCommand
    adoCmd.CommandText = DealWithApostrophesInSQL(sCommand)
' TW 21/03/2007 EP2_1224 End
    Set ExecuteCommand = adoCmd.Execute
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SetExistsState(sKey As String, bExists As Boolean)
    m_colTableExists.Add bExists, sKey
End Sub

Private Sub Class_Initialize()
    m_sActiveSQL = ""
    Set m_colTableExists = New Collection
    Set m_colConnections = New Collection
End Sub
Private Sub Class_Terminate()
    On Error GoTo Failed
    CloseConnection
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
' DJP SQL Server Port
Public Function GetDatabaseType(Optional sDatabaseKey As String) As enumDatabaseTypes
    On Error GoTo Failed
    Dim clsSuperConn As SupervisorConnection
    Dim nType As enumDatabaseTypes
    Dim sType As String
    GetSupervisorConnection clsSuperConn, sDatabaseKey

    sType = clsSuperConn.GetDatabaseType()
    Select Case sType
    Case DATABASE_SQL_SERVER
        nType = INDEX_SQL_SERVER
    
    Case DATABASE_ORACLE
        nType = INDEX_ORACLE
    
    Case Else
        g_clsErrorHandling.RaiseError errDatabaseNotSupported
    End Select
    
    GetDatabaseType = nType
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' DJP SQL Server Port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetStoredProcName
' Description   : Returns the name of the stored procedure, which may or may not be different
'                 between SQL Server and Oracle. The enum nName is the entry point to get the
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetStoredProcName(nName As StoredProcNames) As String
    On Error GoTo Failed
    Dim sName As String
    Dim nType As enumDatabaseTypes
    
    nType = g_clsDataAccess.GetDatabaseType
    
    Select Case nName
    Case spGetPanelDetails
        If nType = INDEX_ORACLE Then
            sName = "SP_SUPERVISOR.GETPANELDETAILS"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETPANELDETAILS"
        End If
    
    Case spGetUsers
        If nType = INDEX_ORACLE Then
            sName = "SP_SUPERVISOR.GETUSERS"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETUSERS"
        End If
    Case spGetComboWithValidation
        If nType = INDEX_ORACLE Then
            sName = "SP_SUPERVISOR.GETCOMBOWITHVALIDATION"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETCOMBOWITHVALIDATION"
        End If
    Case spGetNextConditionsRef
        If nType = INDEX_ORACLE Then
            sName = "SP_GETNEXTQUESTIONREF"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTCONDITIONSREFNO"
        End If
    Case spGetNextQuestionRef
        If nType = INDEX_ORACLE Then
            sName = "SP_GETNEXTQUESTIONREF"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTQUESTIONSREFNO"
        End If
    Case spGetSelectedStages
        If nType = INDEX_ORACLE Then
            sName = "SP_SUPERVISOR.USP_GETSELECTEDSTAGES"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETSELECTEDSTAGES"
        End If
    Case spGetNextProcurationRef
        If nType = INDEX_ORACLE Then
            sName = "USP_GETNEXTPROCURATIONREF"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTPROCURATIONREF"
        End If
    Case spGetNextTargetID
        If nType = INDEX_ORACLE Then
            sName = "USP_GETNEXTTARGETID"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTTARGETID"
        End If
    Case spGetNextSplitID
        If nType = INDEX_ORACLE Then
            sName = "USP_GETNEXTFEESPLITREF"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTFEESPLITREF"
        End If
    
    Case spValidateProduct
        If nType = INDEX_ORACLE Then
            sName = "USP_VALIDATEPRODUCT"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_VALIDATEPRODUCT"
        End If
    
    Case spGetNextSequenceNumber
        If nType = INDEX_ORACLE Then
            sName = "USP_GETNEXTSEQUENCENUMBER"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETNEXTSEQUENCENUMBER"
        End If
    ' PSC 06/03/2006 MAR1298 - Start
    Case spGetTablePrimaryKeys
        If nType = INDEX_ORACLE Then
            sName = "USP_GETTABLEPRIMARYKEYS"
        ElseIf nType = INDEX_SQL_SERVER Then
            sName = "USP_GETTABLEPRIMARYKEYS"
        End If
    End Select
    ' PSC 06/03/2006 MAR1298 - End
    
    If Len(sName) = 0 Then
        g_clsErrorHandling.RaiseError errDatabaseNotSupported, " nType"
    End If
    
    GetStoredProcName = sName
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function ConvertDatabaseItem(vValue As Variant) As Variant
    On Error GoTo Failed


    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GenerateSearchCriteria
' Description   : Given the list of fields and field values passed in, generate a where clause
'                 to be used in a SQL statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GenerateSearchCriteria(colFields As Collection, colValues As Collection, Optional sDatabaseKey As String)
    On Error GoTo Failed
    Dim nTotal As Integer
    Dim nThisItem As Integer
    Dim sAnd As String
    Dim sType As String
    Dim sValue As Variant
    Dim sEqual As String
    Dim sQuote As String
    Dim sField As String
    Dim sSearch As String
    
    nTotal = colFields.Count
    
    sEqual = " = "
    sQuote = "'"
    sAnd = ""
    
    For nThisItem = 1 To nTotal
        sField = colFields(nThisItem)
        sValue = colValues(nThisItem)
        
        ' Only do this if there's a value to search against.
        If Len(sValue) > 0 Then
            sSearch = sSearch + sAnd
            
            sType = TypeName(colValues(nThisItem))
            
            Select Case sType
                Case "Date"
                    'STB: SYS3225 - This is stopping products from loading.
                    'sValue = g_clsValidation.GetDate(CStr(sValue))
                    
                    ' DJP SQL Server port
                    sValue = g_clsSQLAssistSP.FormatDate(sValue, dtfDate, sDatabaseKey)
                    
                    'STB: SYS3225 - This is stopping products from loading.
                    ' DJP SQL Server port
                    'sField = g_clsSQLAssistSP.TruncateDateColumn(sField, sDatabaseKey)
                    
                    sQuote = ""
                Case "String"
    
                    If IsDate(sValue) Then
                        ' Looks something like a date, although vb considers "001a" to be a date
                        If InStr(sValue, "/") Or InStr(sValue, "-") Or InStr(sValue, ".") Then
                            'STB: SYS4244 - This is stopping products from saving (Emply Eligibility is key'd with dates).
                            'sValue = g_clsValidation.GetDate(CStr(sValue))
                            
                            ' DJP SQL Server port
                            sValue = g_clsSQLAssistSP.FormatDate(sValue, dtfDate, sDatabaseKey) '"to_date('" + sValue + "','dd/mm/yyyy')"
                            
                            'STB: SYS4244 - This is stopping products from saving (Emply Eligibility is key'd with dates).
                            ' DJP SQL Server port
                            'sField = g_clsSQLAssistSP.TruncateDateColumn(sField, sDatabaseKey)
                            
                            sQuote = ""
                        End If
                    ElseIf sValue = "null" Then
                        sQuote = ""
                        sEqual = " is "
                    End If
                
                Case "Byte()" ' GUID
                    sValue = g_clsSQLAssistSP.GuidToString(CStr(sValue), , , sDatabaseKey)
                    sValue = g_clsSQLAssistSP.FormatGuid(CStr(sValue), , sDatabaseKey)
                
                    sQuote = ""
                
                Case "Integer", "Double", "Long", "Single", "Decimal", "Boolean"
                    sQuote = ""
                Case Else
            
            End Select

            If Len(sQuote) > 0 Then
                sSearch = sSearch & sField & sEqual & g_clsSQLAssistSP.FormatString(sValue) & " "
            Else
                sSearch = sSearch & sField & sEqual & sValue & " "
            End If
            
            ' Reset quote and AND variables
            sQuote = "'"
            sAnd = " AND "
            sEqual = "="
        End If
    Next

    GenerateSearchCriteria = sSearch
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ClearTableHistory
' Description   : The m_colTableExists collection contains a list of whether or not each table
'                 that has been attempted to be accessed actually exists or not. If we change database,
'                 we need to reinitialise it, so this method will clear the list.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearTableHistory()
    On Error GoTo Failed
    
    Set m_colTableExists = New Collection
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : NextNumber
' Description   : Returns the next number in the sequence, using sTable and sSetField.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextNumber(sTable As String, sSetField As String) As Variant
    Dim vFeeSet As Variant
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    vFeeSet = 1
    sSearch = "SELECT MAX(" + sSetField + ") " + sSetField + " FROM " + sTable
   
    Set rs = g_clsDataAccess.GetTableData(sTable, sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            vFeeSet = rs(sSetField)
                    
            If Not IsNull(vFeeSet) Then
                vFeeSet = vFeeSet + 1
            Else
                vFeeSet = 1
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "NextNumber - No records returned"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "NextNumber: Recordset returned is NULL"
    End If
        
    NextNumber = vFeeSet
    
    Exit Function
Failed:

    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

' PSC 06/03/2006 MAR1298 - Start
Friend Function GetTablePrimaryKeys(ByVal strTableName As String) As Collection
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim colParams As Collection
    Dim adoParam As ADODB.Parameter
    Dim colKeys As Collection

    On Error Resume Next
    
    Set colKeys = New Collection
    Set colParams = New Collection
    Set adoParam = New ADODB.Parameter

    With adoParam
        .Name = "@pTableName"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = strTableName
        .Size = Len(.Value)
    End With

    colParams.Add adoParam
 
    Set rs = ExecuteStoredProcRS(spGetTablePrimaryKeys, colParams)
    
    If rs Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "GetTablePrimaryKeys: Recordset returned is NULL"
    End If
    
    If rs.BOF Or rs.EOF Then
        g_clsErrorHandling.RaiseError errGeneralError, "GetTablePrimaryKeys - No records returned"
    End If
    
    rs.MoveFirst
    
    While Not rs.EOF
         colKeys.Add rs("COLUMN_NAME").Value
        rs.MoveNext
    Wend
    
    Set GetTablePrimaryKeys = colKeys
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' PSC 06/03/2006 MAR1298 - End



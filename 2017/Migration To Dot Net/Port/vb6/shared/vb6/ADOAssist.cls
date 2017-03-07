VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ADOAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Workfile:     ADOAssist.cls
' Copyright:    Copyright © 1999 Marlborough Stirling
'
' Description:  ADO helper object
' Dependencies:
' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     29/06/99    Created
'RF     28/9/99     Added GetNumberOfRecords
'SR     18/11/99    Added methods - CheckRecordExists, GetValueFromTable
'SR     01/12/99    Modified method - GetValueFromTable
'PSC    20/12/99    Do not use default interface
'PSC    01/02/00    Add adExecuteNoRecords to ExecuteSQLCommand
'RF     25/02/00    Enhanced error messaging to help fix AQR SYS0321.
'MC     07/08/00    SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
'PSC    11/08/00    SYS1430 Back ouut SYS1409
'MC     31/08/00    Updated GetConnectionStringFromSPM to correct load testing error.
'APS    07/09/00    Corrected the check for oeChildRecordsFound in ExecuteSQLCommand
'DJP    02/10/00    P514 - get connection string from location if passed in.
'LD     07/11/00    Explicity close database connections
'LD     07/11/00    Explicity destroy command objects
'APS    27/02/01    SYS1986
'DM     22/05/01    SYS2321 Common handling of error numbers
'LD     12/06/01    SYS2368 Modify code to get the connection string for Oracle and SQLServer
'AS     13/06/01    CC012 Added GetValidRecordset
'LD     29/06/01    Modify GetProviderSpecificConnStr
'DC     27/09/01    SYS2748 Modify ExecuteSQLCommand to handle SQL Server error
'DC     02/10/01    SYS2695 GetValuefromTable - check for adbinary added
'PSC    17/10/01    SYS2815 Enable SQL Server Integrated Security
'AS     13/11/03    CORE1 Removed GENERIC_SQL.
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
' Database connection options are soft coded via registry setttings:
'
' Omiga4/Database Connection/Provider
'           - e.g. "MSDAORA" for Oracle or "SQLOLEDB" for SQL Server
'
' Omiga4/Database Connection/Server
'           - database server, e.g. "MSGCH5209" or "OMDBSRV"
'
' Omiga4/Database Connection/Database Name
'           - e.g. "Omiga_IV_POC2_Test"
'
' Omiga4/Database Connection/User ID
'           - e.g. "sa" or "production"
'
' Omiga4/Database Connection/Password
'           - e.g. "" or "production"
'------------------------------------------------------------------------------------------
Option Explicit
Implements IADOAssist
'TK 22/11/2004 BBG1821
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
'TK 22/11/2004 BBG1821 End
'Public Enum DBENGINETYPE
'    Undefined
'    SQLServer
'    Oracle
'End Enum
'
'Enum enumLockType
'    lckShare
'    lckUpdate
'    lckExclusive
'End Enum
' database engine
' fixme: this should be set from SPM
Private eDBEngine As DBENGINETYPE
' application name
Private Const strAppName = "Omiga4"
' registry section for database connection info
Private Const strREGISTRY_SECTION = "Database Connection"
' registry keys for database connection info
Private Const strPROVIDER_KEY = "Provider"
Private Const strSERVER_KEY = "Server"
Private Const strDATABASE_KEY = "Database Name"
Private Const strUID_KEY = "User ID"
Private Const strPASSWORD_KEY = "Password"
Private Const strDATA_SOURCE_KEY = "Data Source"
' registry keys for other database info
Private Const strRETRIES_KEY = "Retries"
Private Const lngUNIQUE_KEY_CONSTRAINT = -2147217900
'DM SYS2321
Private Const lngSQL_SERVER_UNIQUE_KEY_CONSTRAINT = -2147217873 'OLE DB error number
'Public Sub SetConnectionOptions(ByRef refConn As ADODB.Connection)
'    Dim strProvider As String
'
'    refConn.CursorLocation = adUseClient
'
'    strProvider = GetConnectionItem(strPROVIDER_KEY)
'    refConn.Provider = strProvider
'
'    refConn.ConnectionString = _
'            GetProviderSpecificConnStr(strProvider)
'
'End Sub
' DJP    02/10/00    P514, Added Location
Public Function GetConnStr(Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database connection string
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "GetConnStr"
    Dim strConn As String
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    ' DJP    02/10/00    P514 - get connection string from location, if present
    If Not m_objContext Is Nothing Then
        strConn = GetConnectionStringFromSPM(strLocation)
    Else
        strConn = GetConnectionStringFromRegistry(strLocation)
    End If
    GetConnStr = strConn
End Function
Public Function IsUniqueKeyConstraint() As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "IsUniqueKeyConstraint"
    'DM SYS2321 Added processing for SQL Server error number
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine()
    Case Oracle
        If Err.Number = lngUNIQUE_KEY_CONSTRAINT Then
            IsUniqueKeyConstraint = True
        Else
            IsUniqueKeyConstraint = False
        End If
    Case SQLServer
        If Err.Number = lngSQL_SERVER_UNIQUE_KEY_CONSTRAINT Then
            IsUniqueKeyConstraint = True
        Else
            IsUniqueKeyConstraint = False
        End If
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
'    Set objErrAssist = Nothing
End Function
Public Function IsThisUniqueKeyConstraint(intErrNumber As Integer) As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "IsThisUniqueKeyConstraint"
    ' DM SYS2321 Added processing for SQL Server unique constraint checking.
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine()
    Case Oracle
        If intErrNumber = lngUNIQUE_KEY_CONSTRAINT Then
            IsThisUniqueKeyConstraint = True
        Else
            IsThisUniqueKeyConstraint = False
        End If
    Case SQLServer
        If intErrNumber = lngSQL_SERVER_UNIQUE_KEY_CONSTRAINT Then
            IsThisUniqueKeyConstraint = True
        Else
            IsThisUniqueKeyConstraint = False
        End If
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
'    Set objErrAssist = Nothing
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function GetRetriesFromRegistry(Optional sLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database retries string from the registry
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetRetriesFromRegistryVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetRetriesFromRegistry"
'    Dim objErrAssist As New ErrAssist
    Dim strRetries As String
    ' DJP    02/10/00    P514 - get connection string from location, if present
    strRetries = GetConnectionItem(strRETRIES_KEY, sLocation)
    GetRetriesFromRegistry = strRetries
'    Set objErrAssist = Nothing
    Exit Function
GetRetriesFromRegistryVbErr:
    m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    GetRetriesFromRegistry = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function GetConnectionStringFromRegistry(Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database connection string from the registry
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetConnectionStringFromRegistryVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetConnectionStringFromRegistry"
'    Dim objErrAssist As New ErrAssist
    Dim strConnection As String
    Dim strProvider As String
                    
    ' DJP    02/10/00    P514 - get connection string from location, if present
    strProvider = GetConnectionItem(strPROVIDER_KEY, strLocation)
    strConnection = _
            "Provider=" & strProvider & ";" & _
            GetProviderSpecificConnStr(strProvider, strLocation)
'    Set objErrAssist = Nothing
    GetConnectionStringFromRegistry = strConnection
    Exit Function
GetConnectionStringFromRegistryVbErr:
    m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    GetConnectionStringFromRegistry = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function GetProviderSpecificConnStr(ByVal strProvider As String, Optional sLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   build the connection string relevant to provider <strProvider>
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetProviderSpecificConnStrVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetProviderSpecificConnStr"
'    Dim objErrAssist As New ErrAssist
    Dim strConnection As String
    Dim strUserId As String
                        
    ' DJP    02/10/00    P514 - get connection string from location, if present
    ' Default to the default if location not passed in
    If Len(sLocation) = 0 Then
        sLocation = strAppName
    End If
                        
    If Len(strProvider) = 0 Then
        ' "Invalid parameter value: empty string received"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err106, "Provider name must be supplied"
    End If
    If strProvider = "SQLOLEDB" Then
        ' PSC 17/10/01 SYS2815 - Start
        strConnection = _
                "Server=" & GetConnectionItem(strSERVER_KEY, sLocation) & ";" & _
                "database=" & GetConnectionItem(strDATABASE_KEY, sLocation) & ";"
        strUserId = GetConnectionItem(strUID_KEY, sLocation)
        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(strUserId) > 0 Then
            strConnection = strConnection & "UID=" & strUserId & ";" & _
                "pwd=" & GetConnectionItem(strPASSWORD_KEY, sLocation) & ";"
        Else
            strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
        End If
        ' PSC 17/10/01 SYS2815 - End
    Else
        strConnection = _
            "Data Source=" & GetConnectionItem(strDATA_SOURCE_KEY, sLocation) & ";" & _
            "User ID=" & GetConnectionItem(strUID_KEY, sLocation) & ";" & _
            "Password=" & GetConnectionItem(strPASSWORD_KEY, sLocation) & ";"
    End If
        
'    Set objErrAssist = Nothing
    GetProviderSpecificConnStr = strConnection
    Exit Function
GetProviderSpecificConnStrVbErr:
    
    m_objErrAssist.AddToErrSource strFunctionName
    m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    Set objErrAssist = Nothing
    GetProviderSpecificConnStr = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
'Sub Main()
'    Call InitialiseConnectionString
'
'    eDBEngine = GetDBEngine()
'End Sub
' DJP    02/10/00    P514 - get connection string from location, if present
Public Function GetDBEngine(Optional strLocation As String = "") As DBENGINETYPE
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo GetDBEngineVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetDBEngine"
    If eDBEngine = Undefined Then
        InitialiseDBEngine strLocation
    End If
    GetDBEngine = eDBEngine
    Exit Function
        
GetDBEngineVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function InitialiseDBEngine(Optional strLocation As String = "")
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "InitialiseDBEngine"
    Dim strProvider As String
        
    strProvider = GetConnectionItem(strPROVIDER_KEY, strLocation)
    If strProvider = "SQLOLEDB" Then
        eDBEngine = SQLServer
    Else
        eDBEngine = Oracle
    End If
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function GetConnectionItem(ByVal strKey As String, Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get connection string item
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo errhandler
    Dim strFunctionName As String
    strFunctionName = "GetConnectionItem"
'    Dim objErrAssist As New ErrAssist
    Dim strItem As String
    Dim strRegDir As String
                
    'strItem = GetSetting( _
    '    appname:=strAppName, _
    '    Section:=strREGISTRY_SECTION, _
    '    Key:=strKey, _
    '    Default:="")
    If Len(strLocation) = 0 Then
        strLocation = strAppName
    End If
    strRegDir = "SOFTWARE\" + strLocation + "\" + strREGISTRY_SECTION
    strItem = QueryValue(HKEY_LOCAL_MACHINE, strRegDir, strKey)
        
'    Set objErrAssist = Nothing
    GetConnectionItem = strItem
    Exit Function
errhandler:
    
    m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    GetConnectionItem = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function GetConnectionStringFromSPM(Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database connection string from SharedPropertyGroupManager
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetConnectionStringFromSPMVbErr
    Dim strFunctionName As String
    strFunctionName = "GetConnectionStringFromSPM"
    Dim strConn As String
'    Dim objContext As ObjectContext
    Dim strSPMLocation As String
'    Set objContext = GetObjectContext()
    If Len(strLocation) = 0 Then
        strLocation = strAppName
    End If
    If Not m_objContext Is Nothing Then
        Dim spmMgr As SharedPropertyGroupManager
        Dim spmGroup As SharedPropertyGroup
        Dim spmPropConnStr As SharedProperty
        Dim blnExists As Boolean
        ' Create the SharedPropertyGroupManager,
        ' SharedPropertyGroup, and SharedProperty.
        Set spmMgr = CreateObject _
            (gstrMSGSPM)
        'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
        'PSC 11/08/2000. Amend isolation back to LockSetGet
        ' DJP    02/10/00    P514 - If a location is passed in we need to ensure a unique connection
        ' string is saved into SPM.
       strSPMLocation = strLocation + ".Connection String"
       Set spmGroup = spmMgr.CreatePropertyGroup _
            ("Database Connection Details", LockSetGet, Process, blnExists)
        Set spmPropConnStr = _
            spmGroup.CreateProperty(strSPMLocation, blnExists)
        If blnExists Then
            If spmPropConnStr.Value = "" Or spmPropConnStr.Value = "0" Then
                ' initialise the SPM from the registry
                strConn = GetConnectionStringFromRegistry(strLocation)
                spmPropConnStr.Value = strConn
            Else
                ' set the connection string from the SPM
                strConn = spmPropConnStr.Value
            End If
        Else
            'Initialise SPM
            spmPropConnStr.Value = ""
            ' ...and get the correct value for SPM from the registry
            strConn = GetConnectionStringFromRegistry(strLocation)
            spmPropConnStr.Value = strConn
        End If
    Else
        ' MTS not running, so get the connection string value from
        ' the registry
        strConn = GetConnectionStringFromRegistry(strLocation)
    End If
    GetConnectionStringFromSPM = strConn
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmPropConnStr = Nothing
    Exit Function
GetConnectionStringFromSPMVbErr:
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmPropConnStr = Nothing
    GetConnectionStringFromSPM = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetRetriesFromSPM() As Integer
' header ----------------------------------------------------------------------------------
' description:
'   get number of retry attempts from SharedPropertyGroupManager
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetRetriesFromSPMVbErr
    Dim strFunctionName As String
    strFunctionName = "GetRetriesFromSPM"
    Dim strRetries As String
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    If Not m_objContext Is Nothing Then
        Dim spmMgr As SharedPropertyGroupManager
        Dim spmGroup As SharedPropertyGroup
        Dim spmPropRetries As SharedProperty
        Dim blnExists As Boolean
        ' Create the SharedPropertyGroupManager,
        ' SharedPropertyGroup, and SharedProperty.
        Set spmMgr = CreateObject _
            (gstrMSGSPM)
        'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
        'PSC 11/08/2000. Amend isolation back to LockSetGet
        Set spmGroup = spmMgr.CreatePropertyGroup _
            ("Database Connection Details", LockSetGet, Process, blnExists)
        Set spmPropRetries = _
            spmGroup.CreateProperty("Retries", blnExists)
        If blnExists Then
            If spmPropRetries.Value = "" Then
                ' initialise the SPM from the registry
                strRetries = GetRetriesFromRegistry
                spmPropRetries.Value = strRetries
            Else
                ' set the connection string from the SPM
                strRetries = spmPropRetries.Value
            End If
        Else
            ' initialise the SPM from the registry
            strRetries = GetRetriesFromRegistry
            spmPropRetries.Value = strRetries
        End If
    Else
        ' MTS not running, so get the retries value from
        ' the registry
        strRetries = GetRetriesFromRegistry
    End If
    If Len(strRetries) = 0 Then
        ' FIXME err.Raise(
    End If
    GetRetriesFromSPM = CInt(strRetries)
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmPropRetries = Nothing
    Exit Function
GetRetriesFromSPMVbErr:
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmPropRetries = Nothing
    GetRetriesFromSPM = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetErrorNumber(ByVal strErrDesc) As Long
' header ----------------------------------------------------------------------------------
' description:
'   get an error number from an error description
' pass:
'   strErrDesc      Error description. For Oracle, this should be in the format
'                   ' FIXME
' return:
'   GetErrorNumber  Error number stripped out of strErrDesc
' Raise Errors:
'   omiga4Err105
'   omiga4err107
'------------------------------------------------------------------------------------------
On Error GoTo GetErrorNumberVbErr
    Dim strFunctionName As String
    strFunctionName = "GetErrorNumber"
    Dim strErr As String
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine()
    Case Oracle
        ' oracle errors have format "ORA-nnnnn"
        If Len(strErrDesc) > 10 Then
            strErr = Mid(strErrDesc, 5, 5)
        Else
            ' "Cannot interpret database error"
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err107
        End If
    Case SQLServer
        ' FIXME
        ' Cannot get the error number from here, instead use the NativeError attribute
        ' from the ActiveConnection Errors collection. This method should never be
        ' called when SQL Server is the database type.
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
'    Set objErrAssist = Nothing
    'APS SYS1986 27/02/01
    If IsNumeric(strErr) Then
        GetErrorNumber = CLng(strErr)
    Else
        GetErrorNumber = 0
    End If
    Exit Function
GetErrorNumberVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetOmigaErrorNumber(ByVal lngErrNo As Long) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Convert an error number generated by the database engine into an Omiga
'   error number
' pass:
'   lngErrNo    Database engine-specific error number
' return:       Omiga error number
'------------------------------------------------------------------------------------------
On Error GoTo GetOmigaErrorNumberVbErr
    Dim strFunctionName As String
    strFunctionName = "GetOmigaErrorNumber"
'    Dim objErrAssist As New ErrAssist
    Dim lngOmigaErrNo As Long
    Select Case GetDBEngine()
    Case Oracle
        If lngErrNo = 1 Then
            lngOmigaErrNo = oeDuplicateKey
        ElseIf lngErrNo = 2292 Then
            lngOmigaErrNo = oeChildRecordsFound
        End If
        If lngErrNo = 1205 Then
            ' etc
        End If
    Case SQLServer
        ' DM SYS2321
        ' Should be fixed, not a good idea to have hard coded error numbers in here.
        If lngErrNo = 2627 Then
            lngOmigaErrNo = oeDuplicateKey
        ElseIf lngErrNo = 547 Then
            lngOmigaErrNo = oeChildRecordsFound
        End If
        ' No more errors checked for in here.
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4Err105
    End Select
'    Set objErrAssist = Nothing
    GetOmigaErrorNumber = lngOmigaErrNo
    Exit Function
GetOmigaErrorNumberVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IADOAssist_GetRetries() As Integer
' header ----------------------------------------------------------------------------------
' description:
'   get number of database retry attempts to make
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetRetriesVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetRetries"
'    Dim objErrAssist As New ErrAssist
    Dim intRetries As Integer
    ' get retries from SPM where possible
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    If Not m_objContext Is Nothing Then
        intRetries = GetRetriesFromSPM()
    Else
        intRetries = GetRetriesFromRegistry()
    End If
    If intRetries < 0 Then
        Dim strErrDesc As String
        strErrDesc = "Database retries value invalid: " & intRetries
        ' "Invalid parameter value"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, strErrDesc
    End If
    IADOAssist_GetRetries = intRetries
'    Set objErrAssist = Nothing
    Exit Function
GetRetriesVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ExecuteSQLCommand( _
    ByVal vstrSQL As String, _
    Optional ByVal vblnRetry As Boolean = True) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Execute the SQL string <vstrSQL>. If <vblnRetry> is True, do retries where
'   necessary.
' pass:
'   vstrSQL             SQL command to execute
'   vblnRetry           If true, use the global retry value to determine how many
'                       execution attempts to make. If false, do not retry if the
'                       command fails at first attempt.
' return:
'   ExecuteSQLCommand   Number of records affected by the command. This applies
'                       only for action queries or stored procedures. It does not
'                       return the number of records returned by a result-returning
'                       query or stored procedure.
' Raise Errors:
'   omiga4CommandFailed
'   omiga4err106
'------------------------------------------------------------------------------------------
On Error GoTo ExecuteSQLCommandVbErr
    
    Dim strFunctionName As String
    strFunctionName = "ExecuteSQLCommand"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
    Dim strErrDesc As String
    If Len(vstrSQL) < 1 Then
        ' "Invalid parameter value: empty string received"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err106
    End If
    Dim cmd As ADODB.Command
    Dim blnExecutedOk As Boolean
    Dim blnAbandonAttempts As Boolean
    Dim intAttempt As Integer
    Dim intMaxAttempts As Integer
    Dim intRetries As Integer
    If vblnRetry Then
        intRetries = IADOAssist_GetRetries()        ' CL 23/10/00 SYS****
    Else
        intRetries = 0
    End If
        
    intMaxAttempts = 1 + intRetries
    blnExecutedOk = False
    blnAbandonAttempts = False
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = vstrSQL
On Error Resume Next
    
    intAttempt = 1
    While (Not blnExecutedOk) And _
        (intAttempt <= intMaxAttempts) And _
        (Not blnAbandonAttempts)
        cmd.ActiveConnection.ERRORS.Clear
        Dim lngRecordsAffected As Integer
        cmd.Execute lngRecordsAffected, , adExecuteNoRecords
        If cmd.ActiveConnection.ERRORS.Count = 0 Then
            blnExecutedOk = True
        Else
            Dim lngErrNo As Long
            'DM SYS2321 The Error description from SQL Server cannot be used to generate an error number
            ' however the NativeError holds the error number so use that for SQLServer.
            Select Case GetDBEngine
            Case Oracle
                lngErrNo = GetErrorNumber( _
                                cmd.ActiveConnection.ERRORS.Item(0).Description)
            Case SQLServer
                lngErrNo = cmd.ActiveConnection.ERRORS.Item(0).NativeError
            Case Else
                ' "Invalid database engine type"
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
            End Select
            Dim lngOmigaErrType As Long
            lngOmigaErrType = GetOmigaErrorNumber(lngErrNo)
            If lngOmigaErrType = oeDuplicateKey Or _
                lngOmigaErrType = oeChildRecordsFound Then '10/02/2000 MCS make sure we try again
                blnAbandonAttempts = True
            End If
        End If
            
        intAttempt = intAttempt + 1
    Wend
On Error GoTo ExecuteSQLCommandVbErr
    
    If Not blnExecutedOk Then
        
        ' RF 25/02/00 Start - Enhanced error messaging to help fix AQR SYS0321.
        'If cmd.ActiveConnection.Errors.Count > 0 Then
            ' strErrDesc = cmd.ActiveConnection.Errors.Item(0).Description
        'End If
        Dim intItem As Integer
        For intItem = 0 To (cmd.ActiveConnection.ERRORS.Count() - 1)
            If intItem > 0 Then
                strErrDesc = strErrDesc & " ... "
            End If
            strErrDesc = strErrDesc & cmd.ActiveConnection.ERRORS.Item(intItem).Description
        Next
        ' RF 25/02/00 End
        'Commented the following If block and included Case block by SR
        Select Case lngOmigaErrType
            Case oeDuplicateKey
                ' "Database operation failed: Duplicate key"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeDuplicateKey, strErrDesc
            Case oeChildRecordsFound
                ' "Database operation failed: Child Records Found "
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeChildRecordsFound, strErrDesc
            Case Else
                ' "Database operation failed"
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    oeCommandFailed, strErrDesc, Err.HelpFile, Err.HelpContext
        End Select
    End If
    Set cmd = Nothing
'    Set objErrAssist = Nothing
    If lngRecordsAffected = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoRowsAffected
    End If
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ExecuteSQLCommand = lngRecordsAffected
    Exit Function
ExecuteSQLCommandVbErr:
    Set cmd = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    ExecuteSQLCommand = 0
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Sub LockTable(strTableName As String, LockType As enumLockType)
' header ----------------------------------------------------------------------------------
' description:
'   See CSQLConnection::LockTable in SQLX.CPP of Eserver for original code
' pass:
'   strTableName
'   LockType
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo LockTableVbErr
    Dim strFunctionName As String
    strFunctionName = "LockTable"
    Dim strSQL As String
    Dim strLockType As String
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine
    Case SQLServer
        
        Select Case LockType
        Case lckShare
            ' allow others to read table, but prevent them from updating it;
            ' lock held until end of transaction
            strLockType = "tablock holdlock"
        Case lckUpdate
             ' allow transaction to read data (without blocking other readers) and update it
             ' later with the assurance that the data will not have changed since it was read
             ' lock held until end of transaction
             strLockType = "updlock holdlock"
        Case lckExclusive
            ' prevent others from reading or updating the table;
            ' lock held until end of transaction
            strLockType = "tablockx holdlock"
        Case Else
            ' "Invalid lock type"
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err103
        End Select
        ' dummy select on table to lock it (with no overhead of reading records from table)
        strSQL = "select count(*) from " + strTableName + " (" + strLockType + ")"
    Case Oracle
        Select Case LockType
        Case lckShare
            ' allow others to read table, but prevent them from updating it;
            ' lock held until end of transaction
            strLockType = "share"
        Case lckUpdate
            ' allow others to query, but not update the table;
            ' lock held until end of transaction
            strLockType = "share row exclusive"
        Case lckExclusive
            ' prevent others from locking the table
            ' lock held until end of transaction
            strLockType = "exclusive"
        Case Else
            ' "Invalid lock type"
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err103
        End Select
        ' explicit lock statement
        strSQL = "lock table " + strTableName + " in " + strLockType + " mode"
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
    ExecuteSQLCommand (strSQL)
    Exit Sub
LockTableVbErr:
    If Err.Number <> omiga4NoRowsAffected Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub
Public Sub CheckSingleRecordExists(strTableName As String, strKeys As String)
' header ----------------------------------------------------------------------------------
' description:
'   check that one and only one record exists in table <strTableName> for
'   key values <strKeys>.
' pass:
'   strTableName
'   strKeys         Format is: "KeyFieldA = ValueX and KeyFieldB = ValueY
'                   and KeyFieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
' return:           n/a
'------------------------------------------------------------------------------------------
On Error GoTo CheckSingleRecordExistsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "CheckSingleRecordExists"
'    Dim objErrAssist As New ErrAssist
    Dim strSQL As String
    strSQL = "select count(*) Count from " & strTableName & " where " & strKeys
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    If rst!Count <> 1 Then
        ' FIXME: find a better error number to raise
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    rst.Close
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Exit Sub
CheckSingleRecordExistsVbErr:
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
        Set rst = Nothing
        Set cmd = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetFieldTypes(vstrTableName As String) As ADODB.Recordset
' header ----------------------------------------------------------------------------------
' description:
'   Return a Recordset consisting of FIELDNAME and DATATYPE for all the columns in the
'   specified table.
' pass:
'   vstrTableName
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetFieldTypesVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetFieldTypes"
'    Dim objErrAssist As New ErrAssist
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim strSQL As String
    Dim eDbEngineType As DBENGINETYPE
    eDbEngineType = GetDBEngine()
    Select Case eDbEngineType
    Case Oracle
        ' rename the columns to match those used by sp_help in SQLServer, for greater
        ' compatibility
        strSQL = _
            "select COLUMN_NAME Column_Name, DATA_TYPE Type from USER_TAB_COLUMNS " & _
            "where TABLE_NAME = '" & vstrTableName & "'"
    Case SQLServer
        ' not implemented
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        '------------------------------------------------------------------------------------------
        ' hint for SQL Server implementation:
        ' use the following query -
        '       select syscolumns.name 'Column', systypes.name 'Type'
        '       From syscolumns, systypes, sysobjects
        '       where syscolumns.xtype =  systypes.xtype and
        '       syscolumns.id = sysobjects.id and
        '       sysobjects.name = 'tablename'
        ' or the following stored procedure -
        '       sp_help tablename
        '------------------------------------------------------------------------------------------
    Case Else
        ' not implemented
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    End Select
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    Set GetFieldTypes = rst
    Set rst = Nothing
    Set cmd = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetFieldTypesVbErr:
    Set rst = Nothing
    Set cmd = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetNumberOfRecords( _
    ByVal vstrTableName As String, ByVal vstrCriteria As String) As Integer
' header ----------------------------------------------------------------------------------
' description:
'   Returns the number of records based on the table name and criteria passed in.
' pass:
'   vstrTableName
'   vstrCriteria    Format is: "FieldA = ValueX and FieldB = ValueY
'                   and FieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
' return:           Number of records
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfRecordsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetNumberOfRecords"
'    Dim objErrAssist As New ErrAssist
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim strSQL As String
    strSQL = "select count(*) Count from " & vstrTableName & " where " & vstrCriteria
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    GetNumberOfRecords = rst!Count
        rst.Close
            
'    Set objErrAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
GetNumberOfRecordsVbErr:
    
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    GetNumberOfRecords = 0
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' header ----------------------------------------------------------------------------------
' description:
'   Uses the SPM to enable logging of SQL. Anyone who then calls writeSQLtofile will
'   output the SQL
'
'------------------------------------------------------------------------------------------
Public Sub StartSQLLogging()
    On Error GoTo StartSQLLoggingErr
    Const strMethodName As String = "StartSQLLogging"
    SetSQLLoggingFlag True
    Exit Sub
StartSQLLoggingErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
' header ----------------------------------------------------------------------------------
' description:
'   Uses the SPM to disable logging of SQL. Anyone who then calls writeSQLtofile will
'   output the SQL
'
'------------------------------------------------------------------------------------------
Public Sub StopSQLLogging()
    On Error GoTo StopSQLLoggingErr
    Const strMethodName As String = "StopSQLLogging"
        
    SetSQLLoggingFlag False
    Exit Sub
StopSQLLoggingErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub WriteSQLToFile(ByVal vstrSQL As String, _
                           ByVal vstrFunctionName As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Writes the passed SQL to a file on the C drive.
'
'   Note the conditional compilation constants around the code. These are present so that the
'   file is only created when specifically required.
'
' Pass:
'   vstrSQL
'       The SQL to write to the file.
' Returns:
'       N/A
'------------------------------------------------------------------------------------------
    Dim hFile As Integer
    On Error Resume Next
    If IsSQLLoggingOn Then
        hFile = FreeFile
        Open "C:\OutputSQL.log" For Append Access Write As hFile
        Print #hFile, vbCrLf & vbCrLf
        Print #hFile, "********************************************************************"
        Print #hFile, Format$(Now, "DD/MM/YYYY HH:MM:SS") & "     (" & vstrFunctionName & ")" & vbCrLf
        Print #hFile, vstrSQL
        Print #hFile, "********************************************************************"
        Close hFile
    End If
End Sub
Public Function CheckRecordExists(ByVal strTableName As String, ByVal strKeys As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   check that a record exists in table <strTableName> for key values <strKeys>.
' pass:
'   strTableName
'   strKeys         Format is: "KeyFieldA = ValueX and KeyFieldB = ValueY
'                   and KeyFieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
' return:           True If record exist else False
'------------------------------------------------------------------------------------------
On Error GoTo CheckRecordExistsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "CheckRecordExists"
'    Dim objErrAssist As New ErrAssist
    Dim strSQL As String
    strSQL = "select 'x' from " & strTableName & " where " & strKeys
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    If rst.BOF And rst.EOF Then
        CheckRecordExists = False
    Else
        CheckRecordExists = True
    End If
        rst.Close
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Exit Function
CheckRecordExistsVbErr:
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
        Set rst = Nothing
        Set cmd = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetValueFromTable(ByVal strTableName As String, _
                                  ByVal strCondition As String, _
                                  ByVal strColumn As String, _
                                  Optional ByRef varReturnValue As Variant, _
                                  Optional ByRef blnRecordFound As Boolean = True) As Variant
' header ----------------------------------------------------------------------------------
' description:
'   check that a record exists in table <strTableName> for the condition <strCondition> and
'   gets the value of column <strColumn>. Even if multiple records exist, the value in the
'   first record will be passed back.
' pass:
'   strTableName
'   strCondition    Format is: "KeyFieldA = ValueX and KeyFieldB = ValueY
'                   and KeyFieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
'   strColumn     : Name of the column (of which the value is returned)
'   varReturnValue: Value returned - Sent ByRef (RAW value is converted to string)
'   blnRecordFound: Whether a recrd is found in the table that satifies the condition passed
' return:           Value of the column
' Raise Errors:     oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetValueFromTableVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetValueFromTable"
'    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim varTemp As Variant
    Dim strSQL As String
    strSQL = "select " & strColumn & " from " & strTableName & " where " & strCondition
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    If rst.BOF And rst.EOF Then
        blnRecordFound = False
        GetValueFromTable = Null
        Exit Function
    Else
        blnRecordFound = True
    End If
   If Not IsNull(rst.Fields(0).Value) Then
          'SYS2695 - DRC added the check for adBinary
        If rst.Fields(0).Type = adVarBinary Or rst.Fields(0).Type = adBinary Then  'If the value is GUID
            varTemp = objSQLAssist.GuidToString(rst.Fields(0).Value)
        Else
            varTemp = rst.Fields(0).Value
        End If
        If Not IsMissing(varReturnValue) Then 'If varReturnValue was passed to the function
            varReturnValue = varTemp
        End If
        GetValueFromTable = varTemp
    Else
        GetValueFromTable = rst.Fields(0).Value
    End If
        rst.Close
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetValueFromTableVbErr:
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IADOAssist_CheckRecordExists(ByVal strTableName As String, ByVal strKeys As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   check that a record exists in table <strTableName> for key values <strKeys>.
' pass:
'   strTableName
'   strKeys         Format is: "KeyFieldA = ValueX and KeyFieldB = ValueY
'                   and KeyFieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
' return:           True If record exist else False
'------------------------------------------------------------------------------------------
On Error GoTo CheckRecordExistsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "CheckRecordExists"
'    Dim objErrAssist As New ErrAssist
    Dim strSQL As String
    strSQL = "select 'x' from " & strTableName & " where " & strKeys
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    If rst.BOF And rst.EOF Then
        IADOAssist_CheckRecordExists = False
    Else
        IADOAssist_CheckRecordExists = True
    End If
        rst.Close
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Exit Function
CheckRecordExistsVbErr:
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IADOAssist_ExecuteSQLCommand(ByVal vstrSQL As String, Optional ByVal vblnRetry As Boolean = True) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Execute the SQL string <vstrSQL>. If <vblnRetry> is True, do retries where
'   necessary.
' pass:
'   vstrSQL             SQL command to execute
'   vblnRetry           If true, use the global retry value to determine how many
'                       execution attempts to make. If false, do not retry if the
'                       command fails at first attempt.
' return:
'                       Number of records affected by the command. This applies
'                       only for action queries or stored procedures. It does not
'                       return the number of records returned by a result-returning
'                       query or stored procedure.
'------------------------------------------------------------------------------------------
On Error GoTo ExecuteSQLCommandVbErr
    
    Dim strFunctionName As String
    strFunctionName = "ExecuteSQLCommand"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
    Dim strErrDesc As String
    If Len(vstrSQL) < 1 Then
        ' "Invalid parameter value: empty string received"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 106
    End If
    Dim cmd As ADODB.Command
    Dim blnExecutedOk As Boolean
    Dim blnAbandonAttempts As Boolean
    Dim intAttempt As Integer
    Dim intMaxAttempts As Integer
    Dim intRetries As Integer
    If vblnRetry Then
        intRetries = IADOAssist_GetRetries()        ' CL 23/10/00 SYS****
    Else
        intRetries = 0
    End If
        
    intMaxAttempts = 1 + intRetries
    blnExecutedOk = False
    blnAbandonAttempts = False
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = vstrSQL
On Error Resume Next
    
    intAttempt = 1
    While (Not blnExecutedOk) And _
        (intAttempt <= intMaxAttempts) And _
        (Not blnAbandonAttempts)
        cmd.ActiveConnection.ERRORS.Clear
        Dim lngRecordsAffected As Integer
        cmd.Execute lngRecordsAffected, , adExecuteNoRecords
        If cmd.ActiveConnection.ERRORS.Count = 0 Then
            blnExecutedOk = True
        Else
            Dim lngErrNo As Long
            'DRC SYS2748 - cloned from adoAssist.ExecuteSQLCommand
            'DM SYS2321 The Error description from SQL Server cannot be used to generate an error number
            ' however the NativeError holds the error number so use that for SQLServer.
            Select Case GetDBEngine
            Case Oracle
                lngErrNo = GetErrorNumber( _
                                cmd.ActiveConnection.ERRORS.Item(0).Description)
            Case SQLServer
                lngErrNo = cmd.ActiveConnection.ERRORS.Item(0).NativeError
            Case Else
                ' "Invalid database engine type"
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
            End Select
            'End DRC SYS2748
            Dim lngOmigaErrType As Long
            lngOmigaErrType = GetOmigaErrorNumber(lngErrNo)
            ' APS 07/09/00: Added to the oeChildRecordsFound check to
            ' make the check work
            If lngOmigaErrType = oeDuplicateKey Or _
               lngOmigaErrType = oeChildRecordsFound Then
                blnAbandonAttempts = True
            End If
        End If
            
        intAttempt = intAttempt + 1
    Wend
On Error GoTo ExecuteSQLCommandVbErr
    
    If Not blnExecutedOk Then
        
        ' RF 25/02/00 Start - Enhanced error messaging to help fix AQR SYS0321.
        'If cmd.ActiveConnection.Errors.Count > 0 Then
            ' strErrDesc = cmd.ActiveConnection.Errors.Item(0).Description
        'End If
        Dim intItem As Integer
        For intItem = 0 To (cmd.ActiveConnection.ERRORS.Count() - 1)
            If intItem > 0 Then
                strErrDesc = strErrDesc & " ... "
            End If
            strErrDesc = strErrDesc & cmd.ActiveConnection.ERRORS.Item(intItem).Description
        Next
        ' RF 25/02/00 End
        'Commented the following If block and included Case block by SR
        Select Case lngOmigaErrType
            Case oeDuplicateKey
                ' "Database operation failed: Duplicate key"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeDuplicateKey, strErrDesc
            Case oeChildRecordsFound
                ' "Database operation failed: Child Records Found "
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeChildRecordsFound, strErrDesc
            Case Else
                ' "Database operation failed"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeCommandFailed, strErrDesc
        End Select
    End If
    Set cmd = Nothing
'    Set objErrAssist = Nothing
    If lngRecordsAffected = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNoRowsAffected
    End If
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    IADOAssist_ExecuteSQLCommand = lngRecordsAffected
    Exit Function
ExecuteSQLCommandVbErr:
    Set cmd = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    IADOAssist_ExecuteSQLCommand = 0
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Err.Raise Err.Number, Err.Source, Err.Description & vbCrLf & "SQL: " & vstrSQL, Err.HelpFile, Err.HelpContext
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function IADOAssist_GetConnStr(Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database connection string
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "GetConnStr"
    Dim strConn As String
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    If Not m_objContext Is Nothing Then
        strConn = GetConnectionStringFromSPM(strLocation)
    Else
        strConn = GetConnectionStringFromRegistry(strLocation)
    End If
    IADOAssist_GetConnStr = strConn
End Function
' DJP    02/10/00    P514 - get connection string from location, if present
Private Function IADOAssist_GetDBEngine(Optional strLocation As String = "") As DBENGINETYPE
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo GetDBEngineVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetDBEngine"
    If eDBEngine = Undefined Then
        InitialiseDBEngine strLocation
    End If
    IADOAssist_GetDBEngine = eDBEngine
    Exit Function
        
GetDBEngineVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IADOAssist_GetFieldTypes(vstrTableName As String) As ADODB.Recordset
' header ----------------------------------------------------------------------------------
' description:
'   Return a Recordset consisting of FIELDNAME and DATATYPE for all the columns in the
'   specified table.
' pass:
'   vstrTableName
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetFieldTypesVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetFieldTypes"
'    Dim objErrAssist As New ErrAssist
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim strSQL As String
    Dim eDbEngineType As DBENGINETYPE
    eDbEngineType = GetDBEngine()
    Select Case eDbEngineType
    Case Oracle
        ' rename the columns to match those used by sp_help in SQLServer, for greater
        ' compatibility
        strSQL = _
            "select COLUMN_NAME Column_Name, DATA_TYPE Type from USER_TAB_COLUMNS " & _
            "where TABLE_NAME = '" & vstrTableName & "'"
    Case SQLServer
        ' not implemented
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
        '------------------------------------------------------------------------------------------
        ' hint for SQL Server implementation:
        ' use the following query -
        '       select syscolumns.name 'Column', systypes.name 'Type'
        '       From syscolumns, systypes, sysobjects
        '       where syscolumns.xtype =  systypes.xtype and
        '       syscolumns.id = sysobjects.id and
        '       sysobjects.name = 'tablename'
        ' or the following stored procedure -
        '       sp_help tablename
        '------------------------------------------------------------------------------------------
    Case Else
        ' not implemented
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
    End Select
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    Set IADOAssist_GetFieldTypes = rst
    Set rst = Nothing
    Set cmd = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetFieldTypesVbErr:
    Set rst = Nothing
    Set cmd = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IADOAssist_GetNumberOfRecords(ByVal vstrTableName As String, ByVal vstrCriteria As String) As Integer
' header ----------------------------------------------------------------------------------
' description:
'   Returns the number of records based on the table name and criteria passed in.
' pass:
'   vstrTableName
'   vstrCriteria    Format is: "FieldA = ValueX and FieldB = ValueY
'                   and FieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
' return:           Number of records
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfRecordsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetNumberOfRecords"
'    Dim objErrAssist As New ErrAssist
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim strSQL As String
    strSQL = "select count(*) Count from " & vstrTableName & " where " & vstrCriteria
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    IADOAssist_GetNumberOfRecords = rst!Count
        rst.Close
                    
'    Set objErrAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
GetNumberOfRecordsVbErr:
    
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set cmd = Nothing
    IADOAssist_GetNumberOfRecords = 0
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IADOAssist_GetValueFromTable(ByVal strTableName As String, ByVal strCondition As String, ByVal strColumn As String, Optional varReturnValue As Variant, Optional blnRecordFound As Boolean = True) As Variant
' header ----------------------------------------------------------------------------------
' description:
'   check that a record exists in table <strTableName> for the condition <strCondition> and
'   gets the value of column <strColumn>. Even if multiple records exist, the value in the
'   first record will be passed back.
' pass:
'   strTableName
'   strCondition    Format is: "KeyFieldA = ValueX and KeyFieldB = ValueY
'                   and KeyFieldC = ValueZ". The field values are assumed to have
'                   been formatted appropriately for their associated datatype.
'   strColumn     : Name of the column (of which the value is returned)
'   varReturnValue: Value returned - Sent ByRef (RAW value is converted to string)
'   blnRecordFound: Whether a recrd is found in the table that satifies the condition passed
' return:           Value of the column
' Raise Errors:     oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetValueFromTableVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetValueFromTable"
'    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim varTemp As Variant
    Dim strSQL As String
    strSQL = "select " & strColumn & " from " & strTableName & " where " & strCondition
    Dim rst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)
    If rst.BOF And rst.EOF Then
        blnRecordFound = False
        IADOAssist_GetValueFromTable = Null
        Exit Function
    Else
        blnRecordFound = True
    End If
   If Not IsNull(rst.Fields(0).Value) Then
        'SYS2695 - DRC added the check for adBinary
        If rst.Fields(0).Type = adVarBinary Or rst.Fields(0).Type = adBinary Then  'If the value is GUID
            varTemp = objSQLAssist.GuidToString(rst.Fields(0).Value)
        Else
            varTemp = rst.Fields(0).Value
        End If
        If Not IsMissing(varReturnValue) Then 'If varReturnValue was passed to the function
            varReturnValue = varTemp
        End If
        IADOAssist_GetValueFromTable = varTemp
    Else
        IADOAssist_GetValueFromTable = rst.Fields(0).Value
    End If
    rst.Close
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetValueFromTableVbErr:
    
    ' Close the recordset if still open
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        Set rst = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IADOAssist_IsDBConnectionError(ByVal vstrErrDescription As Variant) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Determines if an error is caused by failure to acquire a database connection
' pass:
' return:
'   Whether error is a database connection error or not
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsDBConnectionErrorVbErr
    Const strFunctionName As String = "IADOAssist_IsDBConnectionError"
    Dim lngErrNo As Long
    Dim blnConnError As Boolean
'    Dim objErrAssist As ErrAssist
'    Set objErrAssist = New ErrAssist
    blnConnError = False
     
    Select Case GetDBEngine()
    Case Oracle
        'DM SYS2321
        ' We can use the GetErrorNumber function for Oracle but not for SQLServer
        lngErrNo = GetErrorNumber(vstrErrDescription)
          
        'ORA-00018   maximum number of sessions exceeded
        'ORA-00019   maximum number of session licenses exceeded
        'ORA-00020   maximum number of processes num exceeded
        If lngErrNo >= 18 And lngErrNo <= 20 Then
            blnConnError = True
        End If
    Case SQLServer
        'DM SYS2321
        ' Need to determin if this is a connection error from the description returned from
        ' SQL Server Needs to be fixed.
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
        If True Then
            blnConnError = True
        End If
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInternalError
    End Select
'    Set objErrAssist = Nothing
    IADOAssist_IsDBConnectionError = blnConnError
    Exit Function
IsDBConnectionErrorVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IADOAssist_IsThisUniqueKeyConstraint(intErrNumber As Integer) As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "IsThisUniqueKeyConstraint"
    'DM SYS2321 Handle the SQL Server unique constraint error.
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine()
    Case Oracle
        If intErrNumber = lngUNIQUE_KEY_CONSTRAINT Then
            IADOAssist_IsThisUniqueKeyConstraint = True
        Else
            IADOAssist_IsThisUniqueKeyConstraint = False
        End If
    Case SQLServer
        If intErrNumber = lngSQL_SERVER_UNIQUE_KEY_CONSTRAINT Then
            IADOAssist_IsThisUniqueKeyConstraint = True
        Else
            IADOAssist_IsThisUniqueKeyConstraint = False
        End If
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
'    Set objErrAssist = Nothing
End Function
Private Function IADOAssist_IsUniqueKeyConstraint() As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "IsUniqueKeyConstraint"
    'DM SYS2321 Handle the SQL Server unique constraint error.
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine()
    Case Oracle
        If Err.Number = lngUNIQUE_KEY_CONSTRAINT Then
            IADOAssist_IsUniqueKeyConstraint = True
        Else
            IADOAssist_IsUniqueKeyConstraint = False
        End If
    Case SQLServer
        If Err.Number = lngSQL_SERVER_UNIQUE_KEY_CONSTRAINT Then
            IADOAssist_IsUniqueKeyConstraint = True
        Else
            IADOAssist_IsUniqueKeyConstraint = False
        End If
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4Err105
    End Select
'    Set objErrAssist = Nothing
End Function
Private Sub IADOAssist_LockTable(strTableName As String, LockType As enumLockType)
' header ----------------------------------------------------------------------------------
' description:
'   See CSQLConnection::LockTable in SQLX.CPP of Eserver for original code
' pass:
'   strTableName
'   LockType
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo LockTableVbErr
    Dim strFunctionName As String
    strFunctionName = "LockTable"
    Dim strSQL As String
    Dim strLockType As String
'    Dim objErrAssist As New ErrAssist
    Select Case GetDBEngine
    Case SQLServer
        
        Select Case LockType
        Case lckShare
            ' allow others to read table, but prevent them from updating it;
            ' lock held until end of transaction
            strLockType = "tablock holdlock"
        Case lckUpdate
             ' allow transaction to read data (without blocking other readers) and update it
             ' later with the assurance that the data will not have changed since it was read
             ' lock held until end of transaction
             strLockType = "updlock holdlock"
        Case lckExclusive
            ' prevent others from reading or updating the table;
            ' lock held until end of transaction
            strLockType = "tablockx holdlock"
        Case Else
            ' "Invalid lock type"
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 103
        End Select
        ' dummy select on table to lock it (with no overhead of reading records from table)
        strSQL = "select count(*) from " + strTableName + " (" + strLockType + ")"
    Case Oracle
        Select Case LockType
        Case lckShare
            ' allow others to read table, but prevent them from updating it;
            ' lock held until end of transaction
            strLockType = "share"
        Case lckUpdate
            ' allow others to query, but not update the table;
            ' lock held until end of transaction
            strLockType = "share row exclusive"
        Case lckExclusive
            ' prevent others from locking the table
            ' lock held until end of transaction
            strLockType = "exclusive"
        Case Else
            ' "Invalid lock type"
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 103
        End Select
        ' explicit lock statement
        strSQL = "lock table " + strTableName + " in " + strLockType + " mode"
    Case Else
        ' "Invalid database engine type"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 105
    End Select
    ExecuteSQLCommand (strSQL)
    Exit Sub
LockTableVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeNoRowsAffected Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub
Private Sub SetSQLLoggingFlag(vblnFlag As Boolean)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnExists As Boolean
        
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
    'PSC 11/08/2000. Amend isolation back to LockSetGet
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("SQLLogOptions", LockSetGet, Process, blnExists)
    Set spmProperty = _
        spmGroup.CreateProperty("Logging", blnExists)
    If vblnFlag = True Then
        spmProperty.Value = "ON"
    Else
        spmProperty.Value = "OFF"
    End If
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing
End Sub
Private Function IsSQLLoggingOn() As Boolean
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnGroupExists As Boolean
    Dim blnPropertyExists As Boolean
    IsSQLLoggingOn = False
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
            
    'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
    'PSC 11/08/2000. Amend isolation back to LockSetGet
    Set spmGroup = spmMgr.CreatePropertyGroup _
                ("SQLLogOptions", LockSetGet, Process, blnGroupExists)
    Set spmProperty = _
                spmGroup.CreateProperty("Logging", blnPropertyExists)
    IsSQLLoggingOn = (spmProperty.Value = "ON")
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing
End Function
' AS 13/06/01 CC012 Start
' -----------------------------------------------------------------------------------------
' Description:  Gets a valid recordset returned by executing a stored procedure.
'               Useful when a stored procedure may return multiple recordsets,
'               only one of which is the one you want. This is particularly
'               true of SQL Server, where inserts, creates etc generated
'               recordsets.
' Pass:
' vrstRecordSet The initial recordset returned by the stored procedure. Note: This is
'               passed by reference, which is OK as we're not marshalling across
'               process boundaries.
' nRecordSet    The n'th valid recordset to return. Valid means the recordset is open
'               and contains records. Default to 1 to return first recordset.
' Return:       True if successful.
' AS            31/05/01    First version
'------------------------------------------------------------------------------------------
Public Function GetValidRecordset(ByRef vrstRecordSet As ADODB.Recordset, Optional ByVal nRecordSet As Integer = 1) As Boolean
On Error GoTo GetValidRecordsetErr
    Const cstrFunctionName = "GetValidRecordset"
    Dim bSuccess As Boolean
    bSuccess = False
    If GetDBEngine = SQLServer Then
        Do Until vrstRecordSet Is Nothing
            If Not vrstRecordSet.State = adStateClosed Then
                If Not (vrstRecordSet.BOF And vrstRecordSet.EOF) Then
                    If nRecordSet > 0 Then
                        nRecordSet = nRecordSet - 1
                    End If
                    bSuccess = True
                    If nRecordSet = 0 Then
                        Exit Do
                    End If
                End If
            End If
            Set vrstRecordSet = vrstRecordSet.NextRecordset
        Loop
    ElseIf Not vrstRecordSet Is Nothing And Not vrstRecordSet.State = adStateClosed And Not (vrstRecordSet.BOF And vrstRecordSet.EOF) Then
        ' Found an open recordset with records.
        bSuccess = True
    End If
    GetValidRecordset = bSuccess
    Exit Function
GetValidRecordsetErr:
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Function
Public Function IADOAssist_GetValidRecordset(ByRef vrstRecordSet As ADODB.Recordset, Optional ByVal nRecordSet As Integer = 1) As Boolean
    IADOAssist_GetValidRecordset = GetValidRecordset(vrstRecordSet, nRecordSet)
End Function
' AS 13/06/01 CC012 End
'TK 22/11/2004 BBG1821
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
End Sub
'TK 22/11/2004 BBG1821 End

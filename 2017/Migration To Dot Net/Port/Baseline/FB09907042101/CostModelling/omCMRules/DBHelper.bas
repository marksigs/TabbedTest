Attribute VB_Name = "DBHelper"
Option Explicit

Public Enum DBPROVIDER
    omiga4DBPROVIDERUnknown
    omiga4DBPROVIDEROracle
    omiga4DBPROVIDERSQLServer
End Enum

Private Const m_cstrAppName = "Omiga4"
Private Const m_cstrREGISTRY_SECTION = "Database Connection"
Private Const m_cstrPROVIDER_KEY = "Provider"
Private Const m_cstrSERVER_KEY = "Server"
Private Const m_cstrDATABASE_KEY = "Database Name"
Private Const m_cstrUID_KEY = "User ID"
Private Const m_cstrPASSWORD_KEY = "Password"
Private Const m_cstrDATA_SOURCE_KEY = "Data Source"
Private Const m_cstrRETRIES_KEY = "Retries"

Private Const m_lngENTRYNOTFOUND As Long = -2147024894
Private m_xmldocSchemas As FreeThreadedDOMDocument40
Private m_eDbProvider As DBPROVIDER
Private m_intDBRetries As Integer
Private m_strDbConnectionString As String
'
Public Sub DBAssistBuildDbConnectionString()
On Error GoTo DBAssistBuildDbConnectionStringVbErr

    Dim objWshShell As Object

    Dim strConnection As String, _
        strProvider As String, _
        strRegSection As String, _
        strRetries As String

    Dim strUserId As String
    Dim strPassword As String

    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String

    Set objWshShell = CreateObject("WScript.Shell")

    strRegSection = "HKLM\SOFTWARE\" & m_cstrAppName & "\" & App.Title & "\" & m_cstrREGISTRY_SECTION & "\"

On Error Resume Next

    strProvider = objWshShell.RegRead(strRegSection & m_cstrPROVIDER_KEY)

On Error GoTo DBAssistBuildDbConnectionStringVbErr

    If Len(strProvider) = 0 Then
        strRegSection = "HKLM\SOFTWARE\" & m_cstrAppName & "\" & m_cstrREGISTRY_SECTION & "\"
        strProvider = objWshShell.RegRead(strRegSection & m_cstrPROVIDER_KEY)
    End If

    m_eDbProvider = omiga4DBPROVIDERUnknown

    If strProvider = "MSDAORA" Then
        m_eDbProvider = omiga4DBPROVIDEROracle
        strConnection = _
            "Provider=MSDAORA;Data Source=" & objWshShell.RegRead(strRegSection & m_cstrDATA_SOURCE_KEY) & ";" & _
            "User ID=" & objWshShell.RegRead(strRegSection & m_cstrUID_KEY) & ";" & _
            "Password=" & objWshShell.RegRead(strRegSection & m_cstrPASSWORD_KEY) & ";"
    ElseIf strProvider = "SQLOLEDB" Then
        m_eDbProvider = omiga4DBPROVIDERSQLServer

        ' PSC 17/10/01 SYS2815 - Start
        strUserId = ""
        strPassword = ""

        On Error Resume Next

        strUserId = objWshShell.RegRead(strRegSection & m_cstrUID_KEY)

        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description

        On Error GoTo DBAssistBuildDbConnectionStringVbErr

        If Err.Number <> m_lngENTRYNOTFOUND And Err.Number <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If

        On Error Resume Next

        strPassword = objWshShell.RegRead(strRegSection & m_cstrPASSWORD_KEY)

        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description

        On Error GoTo DBAssistBuildDbConnectionStringVbErr

        If Err.Number <> m_lngENTRYNOTFOUND And Err.Number <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If

        strConnection = _
            "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & m_cstrSERVER_KEY) & ";" & _
            "database=" & objWshShell.RegRead(strRegSection & m_cstrDATABASE_KEY) & ";"

        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(strUserId) > 0 Then
            strConnection = strConnection & "UID=" & strUserId & ";" & _
                "pwd=" & strPassword & ";"
        Else
            strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
        End If
        ' PSC 17/10/01 SYS2815 - End
    End If

    m_strDbConnectionString = strConnection

    strRetries = objWshShell.RegRead(strRegSection & m_cstrRETRIES_KEY)

    If Len(strRetries) > 0 Then
        m_intDBRetries = CInt(strRetries)
    End If

    Set objWshShell = Nothing

    Debug.Print strConnection

    Exit Sub

DBAssistBuildDbConnectionStringVbErr:

    Set objWshShell = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Omiga4Support"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : Omiga4Support
' Description   : Helper functions for access to the Omiga 4 Application Server
' Change history
' Prog      Date        Description
' DJP       03/09/01    Updated for XML Parser 3.0
' SA        21/02/02    SYS4111 Trap error from server object.
' STB       20/05/02    SYS4661 Dump failed ASP request to Supervisor.dmp also,
'                       made provision to supply NT credentials to IIS server (if supplied).
' GHun      03/11/2005  MAR411  Changed RunASP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public Function RunASP(sRequest As String, sServerAsp As String) As String
    On Error GoTo Failed

    Dim sServer As String
    Dim sAppServer As String
    Dim sNTUsername As String
    Dim sNTPassword As String
    Dim objHTTP As ServerXMLHTTP
    Dim xmlDOMResp As FreeThreadedDOMDocument
    
    Set xmlDOMResp = New FreeThreadedDOMDocument
    Set objHTTP = New ServerXMLHTTP
    
    sAppServer = GetApplicationServer()
    
    If Len(sAppServer) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Application Server not defined - See Tools->Options for setup"
    End If
    
    If Len(sAppServer) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Application Server directory not defined - See Tools->Options for setup"
    End If
    
    If sAppServer Like "[A-Za-z]*://*" Then
        sServer = sAppServer & "/xml/" & sServerAsp
    Else
        sServer = "http://" & sAppServer & "/xml/" & sServerAsp
    End If

    'SYS4661 - Allow override of credentials.
    sNTUsername = QueryValue(HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, "NTUsername")
    sNTPassword = QueryValue(HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, "NTPassword")
    
    'MAR411 GHun HTTP verb "POST" must be in upper case
    objHTTP.Open "POST", sServer, False, sNTUsername, sNTPassword
    objHTTP.send sRequest
    
    'SYS4111 If the send method failed, a trappable error is raised
    If objHTTP.Status >= 400 And objHTTP.Status <= 599 Then
        'SYS4661 - Dump the result to a physical disk file.
        DumpResponse objHTTP.responseText
        
        RunASP = objHTTP.statusText
    Else
        RunASP = objHTTP.responseText
    End If
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub DumpResponse(ByVal sMessage As String)

    Dim oFileSystem As Object
    Dim oFile As Object
    
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFileSystem.OpenTextFile(App.EXEName & ".dmp", 2, True)
    
    oFile.Write sMessage
    oFile.Close
    
    Set oFile = Nothing
    Set oFileSystem = Nothing
    
End Sub

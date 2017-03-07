VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      WMI.cls
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

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_COMPUTER_LENGTH = 31

Public m_SnapInData As SnapInData

Public Function WMIGetWbemServices(ByVal strComputerName, ByVal strNameSpace As String) As ISWbemServices
    
    On Error GoTo ErrTrap_WMIGetWbemServices
    
    Dim WbemLocator As SWbemLocator
    Dim strThisComputerName As String
    
    Dim lLength As Long
    lLength = MAX_COMPUTER_LENGTH + 1
    strThisComputerName = String(lLength, vbNullChar)
    
    If GetComputerName(strThisComputerName, lLength) <> 0 Then
        strThisComputerName = Left(strThisComputerName, lLength)
        ' can only use username/password on remote computers
        If UCase(strComputerName) <> UCase(strThisComputerName) Then
            Select Case m_SnapInData.ConfigData.m_datappgES.m_eAccount
                Case Account_CurrentUser
                    Set WMIGetWbemServices = GetObject("winmgmts:{impersonationLevel=impersonate}//" + strComputerName + "/" + strNameSpace)
                Case Account_ConnectAsUser
                    Dim strDomainUserName As String
                    strDomainUserName = m_SnapInData.ConfigData.m_datappgES.m_strDomain
                    If strDomainUserName = "" Then
                        strDomainUserName = strDomainUserName + "\" + m_SnapInData.ConfigData.m_datappgES.m_strUserName
                    Else
                        strDomainUserName = m_SnapInData.ConfigData.m_datappgES.m_strUserName
                    End If
                    Set WbemLocator = New SWbemLocator
                    Set WMIGetWbemServices = WbemLocator.ConnectServer(strComputerName, strNameSpace, strDomainUserName, m_SnapInData.ConfigData.m_datappgES.m_strPassword)
                    Set WbemLocator = Nothing
                Case Else
            End Select
        Else
            Set WbemLocator = New SWbemLocator
            Set WMIGetWbemServices = WbemLocator.ConnectServer(strComputerName, strNameSpace)
            Set WbemLocator = Nothing
        End If
    Else
        Err.Raise 0, "GetComputerName", "Failed"
    End If
    
    Exit Function
    
' Error Handler for this method
ErrTrap_WMIGetWbemServices:
    m_SnapInData.Error.DisplayError "WMIGetWbemServices"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



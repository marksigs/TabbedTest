Attribute VB_Name = "dbxmlAssist"
'Workfile:      dbxmlAssist.bas
'Copyright:     Copyright © 2005 Marlborough Stirling
'
'Description:   Provides routines to return data selected from gthe database as xml
'
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date            Description
' TW    13/05/2005      Initial Creation
'------------------------------------------------------------------------------------------
Option Explicit
Dim strDBConnectionString As String

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
        Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, lpData As _
        Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As Long, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

Private Function GetDataFromDatabase(strSQL As String) As IXMLDOMNode
' Get the requested data as an xml node

Const cstrFunctionName As String = "GetDataFromDatabase"
Dim adoConn As ADODB.Connection
Dim adoCommand As ADODB.Command
Dim adoStream As ADODB.Stream

Dim intError As Integer

Dim strData As String

Dim xmlElem As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40

    On Error GoTo GetDataFromDatabaseVbErr:
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    
    With adoConn
        .ConnectionString = GetConnectionStringFromRegistry()
        .open
    End With
    
    adoStream.open
    With adoCommand
        .CommandText = strSQL
        .CommandType = adCmdText
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Execute , , adExecuteStream
    End With
    strData = adoStream.ReadText()
    
    If Len(strData) = 0 Then
        Set GetDataFromDatabase = Nothing
    Else
    
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.setProperty "NewParser", True
        xmlDoc.async = True
        xmlDoc.loadXML strData
        
        If xmlDoc.parseError = 0 Then
            intError = 0
        Else
            intError = 502
            Set xmlElem = xmlDoc.createElement("XML_ERROR")
            xmlElem.setAttribute "ERRORCODE", xmlDoc.parseError.errorCode
            xmlElem.setAttribute "FILEPOS", xmlDoc.parseError.filepos
            xmlElem.setAttribute "LINE", xmlDoc.parseError.Line
            xmlElem.setAttribute "LINEPOS", xmlDoc.parseError.linepos
            xmlElem.setAttribute "REASON", xmlDoc.parseError.reason
            xmlElem.Text = xmlDoc.parseError.srcText
            xmlDoc.appendChild xmlElem
        End If
        
        If Not xmlDoc.documentElement Is Nothing Then
            Set GetDataFromDatabase = xmlDoc.documentElement
        End If
    End If
    
TidyUp:
    
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    
    ' Close the stream if still open
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
            adoStream.Close
        End If
    End If
        
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
            adoConn.Close
        End If
    End If
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    
    If intError <> 0 Then
        Err.Raise intError, Err.Source, Err.Description
    End If
    
    Exit Function

GetDataFromDatabaseVbErr:
    intError = Err.Number
    Resume TidyUp:
End Function

Public Function dbxmlGetCurrentParameterXML(ByVal strParameterName As String) As IXMLDOMNode
' Get the requested Global Parameter as an xml node

Const cstrFunctionName As String = "dbxmlGetCurrentParameterXML"

Dim strSQL As String

    On Error GoTo GetCurrentParameterXMLVbErr:
    
    strSQL = "SELECT NAME, CONVERT (varchar , GLOBALPARAMETERSTARTDATE , 103) AS GLOBALPARAMETERSTARTDATE, DESCRIPTION, AMOUNT, MAXIMUMAMOUNT, PERCENTAGE, BOOLEAN, STRING FROM GLOBALPARAMETER WHERE NAME = '" & strParameterName & "' FOR XML AUTO, ELEMENTS"
    
    Set dbxmlGetCurrentParameterXML = GetDataFromDatabase(strSQL)

GetCurrentParameterXMLVbErr:

End Function

Public Function dbxmlGetTemplateXML(ByVal strTemplateId As String) As IXMLDOMNode
' Get the requested Template as an xml node

Const cstrFunctionName As String = "dbxmlGetTemplateXML"

Dim strSQL As String

    On Error GoTo GetTemplateXMLVbErr
    
    strSQL = "SELECT * FROM TEMPLATE WHERE TEMPLATEID = " & strTemplateId & " FOR XML AUTO, ELEMENTS"
    
    Set dbxmlGetTemplateXML = GetDataFromDatabase(strSQL)

GetTemplateXMLVbErr:

End Function


Private Function GetConnectionStringFromRegistry(Optional strLocation As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   get database connection string from the registry
' pass:
' return:
'------------------------------------------------------------------------------------------
Dim strFunctionName As String
Dim strUserId As String

    On Error GoTo GetConnectionStringFromRegistryVbErr
    strFunctionName = "GetConnectionStringFromRegistry"
    If Len(strDBConnectionString) = 0 Then
        strDBConnectionString = _
                "Provider=" & GetConnectionItem("Provider") & ";" & _
                "Server=" & GetConnectionItem("Server") & ";" & _
                "database=" & GetConnectionItem("Database Name") & ";"
        strUserId = GetConnectionItem("User ID")
        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(strUserId) > 0 Then
            strDBConnectionString = strDBConnectionString & "UID=" & strUserId & ";" & _
                "pwd=" & GetConnectionItem("Password") & ";"
        Else
            strDBConnectionString = strDBConnectionString & "Integrated Security=SSPI;Persist Security Info=False"
        End If
    End If
    GetConnectionStringFromRegistry = strDBConnectionString
    Exit Function
GetConnectionStringFromRegistryVbErr:
    GetConnectionStringFromRegistry = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetConnectionItem(ByVal strKey As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   get connection string item
' pass:
' return:
'------------------------------------------------------------------------------------------
Const HKEY_LOCAL_MACHINE = &H80000002
Dim strFunctionName As String
Dim strItem As String
                
    On Error GoTo errhandler:
    strFunctionName = "GetConnectionItem"
    strItem = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Omiga4\Database Connection", strKey)
        
    GetConnectionItem = strItem
    Exit Function
errhandler:
    
    GetConnectionItem = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Variant
Const KEY_READ_ACCESS = &H19
Dim lRetVal As Long         'result of the API functions
Dim hKey As Long            'handle of opened key
Dim vValue As Variant       'setting of queried value
    
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_READ_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    QueryValue = vValue
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4

Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

    On Error GoTo QueryValueExError:
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> 0 Then
        Error 5
    End If
    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = 0 Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = 0 Then
                vValue = lValue
            End If
        Case Else
            'all other data types not supported
            lrc = -1
    End Select
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function
    



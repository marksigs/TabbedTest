VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "VersionAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      VersionAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: Uses transaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' RF    06/12/99 Created for AQR SC011: GetVersionList
' CL    26/03/01 SYS2042    Modification for SC035 table structure
' CL    30/03/01 SYS2042    Modifications to allow for non existant general assist
' CL    17/04/01 SYS2255    Removed unwanted comments
' AS    04/08/04 BBG1116    Fixed "Missing Version Information". GetVersionList no longer
'                           calls method on each component, but reads version information directly.
'------------------------------------------------------------------------------------------
Option Explicit

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const MAX_PATH = 260
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type
                                                                                                        
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Const INVALID_HANDLE_VALUE = -1

Public Function GetVersionList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the version info for all components specified in OmigaComponents combo.
' pass:
'   vstrXMLRequest
'       Not currently used - added only for standardisation.
'       Format:
'           <REQUEST></REQUEST>
' return:
'   Format:
'       <RESPONSE>
'           <COMPONENTLIST>
'               <COMPONENT>
'                   <DISPLAYNAME></DISPLAYNAME>
'                   <COMPONENTNAME></COMPONENTNAME>
'                   <VERSION></VERSION>
'                   <BUILDDATE></BUILDDATE>
'               </COMPONENT>
'           </COMPONENTLIST>
'       </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetVersionListVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetVersionList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlComponentNode As IXMLDOMNode
    Dim objComboDO As ComboDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(App.Title & ".ComboDO")
    End If
    
    Dim strComboName As String
    
    strComboName = "OmigaComponents"
    
    '------------------------------------------------------------------------------------------
    ' get a list of components which should be installed
    '------------------------------------------------------------------------------------------
    
    Dim strResponse, strXmlRequest As String
    
    strXmlRequest = "<LIST><LISTNAME>" & strComboName & "</LISTNAME></LIST>"
    
    strResponse = objComboDO.GetComboList(strXmlRequest)
    
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

    '------------------------------------------------------------------------------------------
    ' get version of each component
    '------------------------------------------------------------------------------------------
    
    Set objElem = xmlOut.createElement("COMPONENTLIST")
    Set xmlTempNode = xmlDataNode.appendChild(objElem)
    
    Set xmlNodeList = xmlDoc.getElementsByTagName("LISTENTRY")
    
    Dim intNode As Integer
    Dim strValueId, strVALUENAME, strComponent, strObject As String
    Dim strModuleFileName As String
    Dim strComponentName As String
    Dim nPos As Integer
    Dim strVersion As String
    Dim strBuildDate As String
      
    For intNode = 0 To xmlNodeList.Length - 1
    
        strValueId = objXmlAssist.GetTagValue(xmlNodeList.Item(intNode), "VALUEID")
        strVALUENAME = objXmlAssist.GetTagValue(xmlNodeList.Item(intNode), "VALUENAME")
        strComponent = objComboDO.GetFirstComboValidation(strComboName, strValueId)
        
        If Len(strComponent) = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInternalError, _
                "Component name not found for VALUEID " & strValueId
        End If
               
        ' Assume component is in the same directory as this component.
        strModuleFileName = App.Path & "\" & strComponent
        
        nPos = InStrRev(strComponent, ".", , vbTextCompare)
        If nPos > 0 Then
            ' File extension in component name on database.
            strComponentName = Left(strComponent, nPos - 1)
        Else
            ' No file extension given, so assume component is a DLL.
            strModuleFileName = strModuleFileName & ".dll"
            strComponentName = strComponent
        End If
        
        GetFileVersionDetails strModuleFileName, strVersion, strBuildDate
               
        Set objElem = xmlOut.createElement("COMPONENT")
        Set xmlComponentNode = xmlTempNode.appendChild(objElem)
        
        Set objElem = xmlOut.createElement("DISPLAYNAME")
        objElem.Text = strVALUENAME
        xmlComponentNode.appendChild objElem

        Set objElem = xmlOut.createElement("COMPONENTNAME")
        objElem.Text = strComponentName
        xmlComponentNode.appendChild objElem
           
        Set objElem = xmlOut.createElement("VERSION")
        objElem.Text = strVersion
        xmlComponentNode.appendChild objElem
        
        Set objElem = xmlOut.createElement("BUILDDATE")
        objElem.Text = strBuildDate
        xmlComponentNode.appendChild objElem
               
    Next intNode

    GetVersionList = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetVersionListExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objComboDO = Nothing

    Exit Function

GetVersionListVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetVersionList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume GetVersionListExit

End Function

Private Function GetFileVersionDetails(ByVal strFileName As String, ByRef strVersion As String, ByRef strBuildDate As String) As Boolean
    On Error GoTo GetFileVersionDetailsExit
    
    Dim bSuccess As Boolean
    Dim lFileVersionInfoSize As Long
    Dim lpdwHandle As Long

    bSuccess = False
    strVersion = "Missing Version Information"
    strBuildDate = strVersion
    
    lFileVersionInfoSize = GetFileVersionInfoSize(strFileName, lpdwHandle)
    If lFileVersionInfoSize Then
        Dim FileVersionInfo() As Byte
        ReDim FileVersionInfo(0 To lFileVersionInfoSize - 1) As Byte
        If GetFileVersionInfo(strFileName, 0, lFileVersionInfoSize, FileVersionInfo(0)) Then
            ' Get the code page and language.
            Dim pVerInfo As Long
            Dim lLength As Long
            If VerQueryValue(FileVersionInfo(0), "\VarFileInfo\Translation", pVerInfo, lLength) Then
                
                Dim byCodePageLanguage(0 To 3) As Byte  ' Buffer for code page & language.
                Dim strCodePageLanguage As String       ' 8-digit hex string of cpl.
                Dim strComments As String
                
                CopyMemory byCodePageLanguage(0), ByVal pVerInfo, 4
                ' Swap first two bytes, and last two bytes.
                SwapByte byCodePageLanguage(0), byCodePageLanguage(1)
                SwapByte byCodePageLanguage(2), byCodePageLanguage(3)
                ' Convert four bytes into a 8-digit hexadecimal string.
                strCodePageLanguage = FixedHex(byCodePageLanguage(0), 2) & FixedHex(byCodePageLanguage(1), 2) & FixedHex(byCodePageLanguage(2), 2) & FixedHex(byCodePageLanguage(3), 2)
                If VerQueryValue(FileVersionInfo(0), "\StringFileInfo\" & strCodePageLanguage & "\Comments", pVerInfo, lLength) Then
                    strComments = Space$(lLength)
                    lstrcpy strComments, pVerInfo
                    ' lLength is in bytes; allow 2 bytes per character for Unicode strings.
                    ' Remove trailing NULL byte left over from lstrcpy.
                    strComments = Left(strComments, (lLength / 2) - 1)
                End If
                
                If Len(strComments) >= 36 Then
                    If Left(strComments, 6) = "Build:" Then
                        ' Comments contain version and date information, i.e., a VB component compiled by the
                        ' Omiga 4 build process.
                        strVersion = Mid(strComments, 8, 15)
                        If (Len(strComments) >= 37) Then
                            strBuildDate = Mid(strComments, 26, 11)
                        Else
                            strBuildDate = Mid(strComments, 26, 10)
                        End If
                        bSuccess = True
                    End If
                End If
                
                If Not bSuccess Then
                    ' An external component - use File version.
                    Dim vffi As VS_FIXEDFILEINFO
                    If VerQueryValue(FileVersionInfo(0), "\", pVerInfo, lLength) Then
                        CopyMemory vffi, ByVal pVerInfo, lLength
                        strVersion = _
                            Trim(Str(HIWORD(vffi.dwFileVersionMS))) & "." & _
                            Trim(Str(LOWORD(vffi.dwFileVersionMS))) & "." & _
                            Trim(Str(HIWORD(vffi.dwFileVersionLS))) & "." & _
                            Trim(Str(LOWORD(vffi.dwFileVersionLS)))
                            
                        ' Do not use file creation/modified/accessed date as these do not reflect
                        ' build date.
                        'strBuildDate = GetFileCreationDate(strFileName)
                        strBuildDate = "Not applicable"
                        
                        bSuccess = True
                    End If
                End If
            End If
        End If
    End If
   
GetFileVersionDetailsExit:
    GetFileVersionDetails = bSuccess
End Function

' Swaps two byte values without needing a buffer variable.
' This works as long as the two values to be swapped are  numeric and of the same data type.
Private Sub SwapByte(byte1 As Byte, byte2 As Byte)
    byte1 = byte1 Xor byte2
    byte2 = byte1 Xor byte2
    byte1 = byte1 Xor byte2
End Sub

' Create a hexadecimal string to represent a number, but it outputs a string of a fixed
' number of digits.  Extra zeros are added to make the string the proper length.
' The "&H" prefix is not put into the string.
Private Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
    FixedHex = Right("00000000" & Hex(hexval), nDigits)
End Function

' HIWORD and LOWORD are API macros defined below.
Private Function HIWORD(ByVal dwValue As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(dwValue), 8)
    HIWORD = CLng("&H" & Left(hexstr, 4))
End Function

Private Function LOWORD(ByVal dwValue As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(dwValue), 8)
    LOWORD = CLng("&H" & Right(hexstr, 4))
End Function

Private Function GetFileCreationDate(strFileName As String) As String
    On Error GoTo GetFileCreationDateExit
    
    Dim strBuildDate As String
    
    strBuildDate = ""
    
    ' Get the file creation date.
    Dim hFind As Long
    Dim win32FindData As WIN32_FIND_DATA
    hFind = FindFirstFile(strFileName, win32FindData)
    If Not hFind = INVALID_HANDLE_VALUE Then
        Dim LocalTime As FILETIME
        Dim SysTime As SYSTEMTIME
        If Not FileTimeToLocalFileTime(win32FindData.ftCreationTime, LocalTime) = 0 Then
            If Not FileTimeToSystemTime(LocalTime, SysTime) = 0 Then
                Dim strDate As String
                strDate = Trim(Str(SysTime.wDay)) & "/" & Trim(Str(SysTime.wMonth)) & "/" & Trim(Str(SysTime.wYear))
                If IsDate(strDate) Then
                    Dim dtBuildDate As Date
                    dtBuildDate = CDate(strDate)
                    strBuildDate = Format(dtBuildDate, "DD MMM YYYY")
                End If
            End If
        End If
        FindClose hFind
    End If
    
GetFileCreationDateExit:
    GetFileCreationDate = strBuildDate
End Function

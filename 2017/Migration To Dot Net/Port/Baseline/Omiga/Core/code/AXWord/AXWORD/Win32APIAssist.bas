Attribute VB_Name = "Win32APIAssist"
'Workfile:      Win32APIAssist.bas
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DR     07/03/02 Created
'DJB    01/06/02 Many tweaks and bug fixes.
'DJB    09/08/02 Converted to process Word files passed in in bin.base64
'DJB    02/10/02 Changed to late binding of MSXML2 and Scripting to help with deployment issues.
'LD     15/10/02 Word 97 compatible
'DJB,CD 14/02/03 Create GUID functions added.
'DJB    21/03/03 Changed word activation techniques to try and counteract known MS bug Q188546
'TW     16/06/04 New function (saveBinBase64AsPDFFile) to support viewing pdf files
'AS     25/04/06 CORE266 Axword: add support for viewing AFP files
'AS     31/08/06 CORE297 Axword: Reviewing command bar is not hidden
'------------------------------------------------------------------------------------------
Option Explicit

' Global variables.
Global hMsgHook As Long
Global gblnError As Boolean
Global gblnReadOnly As Boolean
Global gFileSize As Long
Global gFileSaved As Boolean
Global gintWidth As Integer
Global gintHeight As Integer
Global gintCurrentWidth As Integer
Global gintCurrentHeight As Integer
Global gintCurrentTop As Integer
Global gintCurrentLeft As Integer
Global gblnEdit As Boolean
Global gblnHTMLGenerated As Boolean
'TW 16/6/2004
Global gblnPDFGenerated As Boolean
'End TW 16/6/2004
Global gblnWindowShown As Boolean
Global gblnInError As Boolean
Global gblnTrackedChanges As Boolean
Global gblnShowRevisions As Boolean
Global gblnTextSearchEnabled As Boolean
Global gblnResizeableFrame As Boolean
Global gblnPersistState As Boolean
Global gintFileRetries As Integer
Global gblnViewAsWord As Boolean
Global gblnViewAsPDF As Boolean
Global gblnSpellCheckOnSave As Boolean
Global gblnSpellCheckWhileEditing As Boolean
Global gintPageFit As Integer
Global gblnShowFindFreeText As Boolean
Global gblnShowCommandBars As Boolean
Global gblnCommandBarsOn As Boolean
Global gblnShowPrint As Boolean
Global gblnShowPrintDialog As Boolean
Global gblnShowProgressBar As Boolean
Global gblnShowTrackedChanges As Boolean
Global gobjWordApplication As Object
Global glWordVersion As Long
Global gblnDocumentEdited As Boolean
Global gblnDocumentPrinted As Boolean
Global gblnDisablePrintOut As Boolean
Global gblnModeless As Boolean

Public Type LAST_ERR
    Number As Long
    Source As String
    Description As String
End Type
Global gLastErrObject As LAST_ERR

Public Type FILE_CONTENTS
    strBinBase64 As String
    strCompressionMethod As String
    bCompressed As Boolean
    bCompressedArray As Boolean
End Type

Public Type REQUEST_DATA
    FileContents As FILE_CONTENTS
    strOperation As String
    strFileExtension As String
    strCompressionMethod As String
    strDocumentID As String
    strDocumentTitle As String
    strPrinter As String
    nFirstPagePrinterTray As Integer
    nOtherPagesPrinterTray As Integer
    nCopies As Integer
    nDeliveryType As Integer
    bUseDifferentTrayForOtherPages As Boolean
End Type

Global gRequestData As REQUEST_DATA

Private Const gstrObjectName = "Win32APIAssist.bas"

Public Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
 
Public Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long
 
Public Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadID As Long) As Long
     
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long           ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long            ' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long     ' This process's parent process
   pcPriClassBase As Long          ' Base priority of process threads
   dwFlags As Long
   szExeFile As String * 260       ' MAX_PATH
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95.
                                  '2 = Windows NT
   szCSDVersion As String * 128
End Type

Public gDefaultWindowProc As Long
Public minX As Long
Public minY As Long
Public maxX As Long
Public maxY As Long

Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24

Public Type POINTAPI
    X As Long
    y As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Public Const HC_ACTION = 0
Public Const WH_GETMESSAGE = 3
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_CONTROL = &H8
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MOUSEMOVE = &H200
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const MF_BYCOMMAND As Long = &H0
Public Const MF_BYPOSITION As Long = &H400
Public Const MF_REMOVE As Long = &H1000
Public Const SC_SIZE As Long = &HF000
Public Const SC_MOVE As Long = &HF010
Public Const SC_MINIMIZE As Long = &HF020
Public Const SC_MAXIMIZE As Long = &HF030
Public Const SC_NEXTWINDOW As Long = &HF040
Public Const SC_PREVWINDOW As Long = &HF050
Public Const SC_CLOSE As Long = &HF060

Global Const SWP_NOSIZE As Long = &H1
Global Const SWP_NOMOVE As Long = &H2
Global Const SWP_NOZORDER As Long = &H4
Global Const SWP_NOREDRAW As Long = &H8
Global Const SWP_NOACTIVATE As Long = &H10
Global Const SWP_FRAMECHANGED As Long = &H20
Global Const SWP_SHOWWINDOW As Long = &H40
Global Const SWP_HIDEWINDOW As Long = &H80
Global Const SWP_NOCOPYBITS As Long = &H100
Global Const SWP_NOOWNERZORDER As Long = &H200
Global Const SWP_NOSENDCHANGING As Long = &H400

Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global Const ERROR_CALL_NOT_IMPLEMENTED = 120
Global Const STG_E_UNIMPLEMENTEDFUNCTION = &H800300FE

Public Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type WindowPos
    hwnd As Long
    hWndInsertAfter As Long
    X As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Private Type Guid
    D1       As Long
    D2       As Integer
    D3       As Integer
    D4(8)    As Byte
End Type

Private Declare Function WinCoCreateGuid Lib "ole32.dll" Alias "CoCreateGuid" (guidNewGuid As Guid) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_FORCEMINIMIZE = 11
Private Const SW_MAX = 11

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESIZE = &H2
Private Const STARTF_USEPOSITION = &H4
Private Const STARTF_USECOUNTCHARS = &H8
Private Const STARTF_USEFILLATTRIBUTE = &H10
Private Const STARTF_RUNFULLSCREEN = &H20
Private Const STARTF_FORCEONFEEDBACK = &H40
Private Const STARTF_FORCEOFFFEEDBACK = &H80
Private Const STARTF_USESTDHANDLES = &H100

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, _
    ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
Private Declare Function ShellExecuteEx Lib "Shell32" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function FindExecutable Lib "Shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private g_hwndFound As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Boolean

' ----------------------
' Methods and functions.
' ----------------------

' Do not use Win API ShellExecuteEx as it returns a process handle, not the process id.
' The process id is required to activate the process window in AppActivate.
Public Function ShellExecute(ByVal strFileName As String) As Long
    Dim lResult As Long
    Dim strFolder As String
    Dim strExecutable255 As String * 255
    Dim strExecutable As String
    Dim strCommandLine As String
    Dim lProcessId As Long
    lProcessId = 0
    
    ' Find the executable for this document type.
    If Len(strFileName) > 4 And Right$(LCase$(strFileName), 4) = ".afp" Then
        ' AFPs (IBM print file format) should be viewed in Internet Explorer using IBM's browser plug-in.
        Dim strHtmlFileName As String
        strHtmlFileName = Left$(strFileName, Len(strFileName) - 4) & ".html"
        Call FileCopy(strFileName, strHtmlFileName)
        lResult = FindExecutable(strHtmlFileName, strFolder, strExecutable255)
        Kill strHtmlFileName
    Else
        lResult = FindExecutable(strFileName, strFolder, strExecutable255)
    End If
    
    If lResult > 32 And Len(strExecutable255) > 0 Then
        ' Truncate executable name.
        Dim lPos As Long
        lPos = InStr(strExecutable255, Chr(0))
        If lPos > 0 Then
            strExecutable = Left$(strExecutable255, lPos - 1)
        Else
            strExecutable = strExecutable255
        End If
                
        ' Start the shelled application.
        strCommandLine = """" & strExecutable & """ """ & strFileName & """"
        lProcessId = Shell(strCommandLine, vbNormalFocus)
    End If
    
    ShellExecute = lProcessId
    
End Function

' Alternative to VB function AppActivate.
Public Sub ActivateApp(ByVal lProcessId As Long)
    g_hwndFound = 0
    Call EnumWindows(AddressOf EnumWindowProc, lProcessId)
    If g_hwndFound <> 0 Then
        Call SetForegroundWindow(g_hwndFound)
        'Call SetWindowPos(g_hwndFound, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub

Public Function FindProcessWindow(ByVal lProcessId As Long) As Long
    g_hwndFound = 0
    Call EnumWindows(AddressOf EnumWindowProc, lProcessId)
    If g_hwndFound <> 0 Then
        FindProcessWindow = g_hwndFound
    End If
End Function

Private Function EnumWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim hProcID As Long
    
    ' Eliminate windows that are not top-level.
    If GetParent(hwnd) = 0& And IsWindowVisible(hwnd) Then
        Call GetWindowThreadProcessId(hwnd, hProcID)
        If hProcID = lParam Then
            ' Found process.
            g_hwndFound = hwnd
            EnumWindowProc = 0 'stop
            Exit Function
        End If
    End If
    
    'To continue enumeration, return True. To stop enumeration return False (0).
    'When 1 is returned, enumeration continues until there are no more windows left.
    EnumWindowProc = 1
End Function


Public Function CreateGUID() As String
' -----------------------------------------------------------------------------------------
' description:  Creates a Global Unique Identifier
' return:
'------------------------------------------------------------------------------------------

    Dim guidNewGuid  As Guid
    Dim strBuffer    As String

    Call WinCoCreateGuid(guidNewGuid)

    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D1), 8)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D2), 4)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D3), 4)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(0)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(1)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(2)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(3)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(4)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(5)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(6)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(7)), 2)

    CreateGUID = strBuffer

End Function

Private Function PadRight0(ByVal vstrBuffer As String, _
                           ByVal vstrBit As String, _
                           ByVal intLenRequired As Integer, _
                           Optional bHyp As Boolean _
                         ) As String
' -----------------------------------------------------------------------------------------
' description:
' return:
'------------------------------------------------------------------------------------------

    PadRight0 = vstrBuffer & _
                vstrBit & _
                String$(Abs(intLenRequired - Len(vstrBit)), "0")

End Function

#If Not MIN_BUILD Then
Public Function MsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next

    'Tmp holder of a message
    Dim msgTemp As Msg

    If nCode >= 0 Then
        If nCode = HC_ACTION Then
            
            'Make a copy of the message
            Call CopyMemory(msgTemp, ByVal lParam, Len(msgTemp))
            
            Dim szName As String * 255
            
            'Check to see if this is an IE object (I tried intercepting just
            'the IE object but it cant return the hwnd for some reason?!?)
            GetClassName msgTemp.hwnd, szName, Len(szName)

            If Left(szName, Len("Internet Explorer")) = "Internet Explorer" Then
                'Right, what sort of message have we intercepted?
                Select Case msgTemp.Message
                Case WM_KEYDOWN, WM_CONTEXTMENU, WM_RBUTTONDBLCLK, _
                     WM_LBUTTONDBLCLK, WM_RBUTTONDOWN
                
                    'Reset the message ...
                    msgTemp.Message = 0
                    '... and copy it back into memory
                    Call CopyMemory(ByVal lParam, msgTemp, Len(msgTemp))
                    
                Case Else
                    'Do nothing
                        
                End Select
            End If
        End If
    End If
    
    'Process the next message
    MsgProc = CallNextHookEx(hMsgHook, nCode, wParam, lParam)
    
End Function

Public Sub SubClass(hwnd As Long)
    ' Assign our own window message procedure.
    On Error Resume Next
    gDefaultWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub


Public Sub UnSubClass(hwnd As Long)
    ' Restore the default message handling.
    If gDefaultWindowProc Then
        SetWindowLong hwnd, GWL_WNDPROC, gDefaultWindowProc
        gDefaultWindowProc = 0
    End If
End Sub

Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    'window message procedure
    On Error Resume Next
    
    WindowProc = 1
    
    If hwnd = frmaxword.hwnd Then
        'The handle returned is to our form, do perform form-specific message handling
        'to deal with the notifications.
        On Error Resume Next
          
        ' Form-specific handler.
        Select Case uMsg
        Case WM_WINDOWPOSCHANGING
            If Not gblnResizeableFrame And Not IsIconic(hwnd) Then
                Dim WinPos As WindowPos
                CopyMemory WinPos, ByVal lParam, LenB(WinPos)
                'Prevent moving of window.
                'Don't use SWP_NOSIZE as this prevent minimising window when axword is in
                'read only mode; therefore we trap WM_GETMINMAXINFO message instead to
                'prevent resizing.
                WinPos.flags = WinPos.flags Or SWP_NOMOVE
                CopyMemory ByVal lParam, WinPos, LenB(WinPos)
            End If
        Case WM_GETMINMAXINFO
            If Not gblnResizeableFrame And Not IsIconic(hwnd) Then
                Dim MMI As MINMAXINFO
                CopyMemory MMI, ByVal lParam, LenB(MMI)
                With MMI
                    .ptMinTrackSize.X = gintWidth
                    .ptMinTrackSize.y = gintHeight
                    .ptMaxTrackSize.X = gintWidth
                    .ptMaxTrackSize.y = gintHeight
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                'Don't call default WinProc.
                WindowProc = 0
            End If
        End Select
    End If
    
    If WindowProc = 1 Then
        WindowProc = CallWindowProc(gDefaultWindowProc, hwnd, uMsg, wParam, lParam)
    End If
    
End Function
#End If

Public Function getWordVersion(ByRef pobjWordApp As Object) As Long
    ' Find out what version of Word we have.
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "getWordVersion"
    
    Dim dotPos As Integer
    
    dotPos = InStr(pobjWordApp.Version, ".")
    
    If dotPos > 0 Then
        getWordVersion = CInt(Left(pobjWordApp.Version, dotPos))
    Else
        getWordVersion = CInt(Left(pobjWordApp.Version, 1))
    End If
    
    Debug.Print "Word version: " & getWordVersion
    
ExitPoint:

    Handle_Error Err, gstrObjectName, strFunctionName

End Function

#If Not MIN_BUILD Then
Public Function CheckWord() As Boolean
    Dim bFound As Boolean
    Dim cb As Long
    Dim cbNeeded As Long
    Dim NumElements As Long
    Dim ProcessIDs() As Long
    Dim cbNeeded2 As Long
    Dim NumElements2 As Long
    Dim Modules(1 To 200) As Long
    Dim lRet As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    Dim i As Long
    
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
       cb = cb * 2
       ReDim ProcessIDs(cb / 4) As Long
       lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
    NumElements = cbNeeded / 4
    
    For i = 1 To NumElements
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
        
        'Got a Process handle
        If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
           lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
           
           'If the Module Array is retrieved, Get the ModuleFileName
           If lRet <> 0 Then
              ModuleName = Space(MAX_PATH)
              nSize = 500
              lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
              
              If InStr(LCase(ModuleName), "winword.exe") > 0 Then
                   bFound = True
              End If
           End If
        End If
        
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
    
    CheckWord = bFound
End Function
#End If

Public Function saveBinBase64AsWordDoc( _
    ByRef FileContents As FILE_CONTENTS, _
    ByVal nDeliveryType As Integer, _
    ByVal strPath As String, _
    Optional ByVal blnViewAsWord As Boolean = True)
    On Error GoTo ExitPoint

    'Variables.
    Const strFunctionName As String = "saveBinBase64AsWordDoc"

    Dim bSuccess As Boolean
       
    Dim strFullPath As String
    strFullPath = strPath & gRequestData.strFileExtension
    
    bSuccess = saveBinBase64ToFile(FileContents, nDeliveryType, strFullPath)
    
    If bSuccess Then
        Dim blnCreatedApp As Boolean
        
        'AS 31/08/2006 CORE297 Reviewing command bar is not hidden
        'Always call CreateWordApplication as this has the side effect of setting glWordVersion, which is used
        'by ShowCommandBars() to determine if the Reviewing toolbar should be turned off.
        'Create instance of the word app and document
        Set gobjWordApplication = CreateWordApplication(gobjWordApplication, blnCreatedApp)
        
        If Not blnViewAsWord Then
            Dim objWordDocOut As Object
            
            ' Save document as HTML.
    
            ' Open the file in Word.
            If glWordVersion <= 8 Then
                ' Word 97 does not support Visible parameter.
                Set objWordDocOut = _
                    gobjWordApplication.Documents.Open( _
                        FileName:=strFullPath, _
                        ConfirmConversions:=False, _
                        ReadOnly:=True, _
                        AddToRecentFiles:=False)
            Else
                Set objWordDocOut = _
                    gobjWordApplication.Documents.Open( _
                        FileName:=strFullPath, _
                        ConfirmConversions:=False, _
                        ReadOnly:=True, _
                        AddToRecentFiles:=False, _
                        Visible:=False)
            End If
            
            objWordDocOut.Activate
    
            'Stop complaining about the template
            gobjWordApplication.NormalTemplate.Saved = True
        
            ' Check if tracked changes are enabled.
            If objWordDocOut.TrackRevisions Then
                gblnTrackedChanges = True
                ' Turn on and off visible revisions.
                objWordDocOut.ShowRevisions = gblnShowRevisions
            Else
                gblnTrackedChanges = False
            End If
            
            Dim wordFC As Object
            Dim nSaveFormat As Long
            nSaveFormat = wdFormatHTML
            For Each wordFC In gobjWordApplication.FileConverters
                If wordFC.FormatName = "HTML Document" Then
                    nSaveFormat = wordFC.SaveFormat
                    Exit For
                End If
            Next
    
            'Now convert to a HTML file (Word 2000 / Word 97)
            objWordDocOut.SaveAs _
                FileName:=strPath & ".htm", FileFormat:=nSaveFormat, LockComments:=False, _
                password:="", AddToRecentFiles:=False, WritePassword:="", _
                ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
    
            ' Flag the success.
            gblnHTMLGenerated = True
            bSuccess = True
        End If
    End If
    
ExitPoint:
    ' Clean up.
    Close #1
    
    ' Cache error.
    Dim lngErrNo As Long
    Dim strErrDesc As String, strErrSrc As String
    
    lngErrNo = Err.Number
    strErrDesc = Err.Description
    strErrSrc = Err.Source
    
    ' Clean up Word objects.
    Close_Word_Doc objWordDocOut, gstrObjectName
    If blnCreatedApp Then
        Close_Word_App gobjWordApplication, gstrObjectName
    End If
   
    ' Put error back.
    Err.Number = lngErrNo
    Err.Description = strErrDesc
    Err.Source = strErrSrc
    
    saveBinBase64AsWordDoc = bSuccess
    
    Handle_Error Err, gstrObjectName, strFunctionName
End Function

Public Function saveBinBase64ToFile(ByRef FileContents As FILE_CONTENTS, ByVal nDeliveryType As Integer, ByVal strPath As String) As Boolean
    On Error GoTo ExitPoint
    
    Dim bSuccess As Boolean
    bSuccess = False

    'Variables.
    Const strFunctionName As String = "saveBinBase64ToFile"
                  
    If Len(FileContents.strBinBase64) = 0 Then
        Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Error loading XML, zero length XML encountered."
    Else
        
        Dim arrFile() As Byte
        ' Convert the binary Base 64 to a byte array.
        arrFile = ConvertBase64ToBin(gstrObjectName, FileContents, bDecompress:=True)
                  
        ' Save the byte array to disk to produce a document.
        On Error Resume Next
        Dim intTries As Integer
        For intTries = 1 To gintFileRetries
            Open strPath For Binary Access Write Lock Read Write As #1
            If Err.Number <> 0 Then
                ' File is locked, sleep and try again.
                Sleep 500
            Else
                Exit For
            End If
        Next
         
        ' Check for a persistant error.
        If Err.Number <> 0 Then
            GoTo ExitPoint
        End If
        
        On Error GoTo ExitPoint
        
        ' Write the file.
        If nDeliveryType = DELIVERYTYPE_RTF And Not FileContents.bCompressedArray Then
            ' File contents is uncompressed RTF, which need to be converted to a string before saving.
            Dim strFile As String
            strFile = arrFile
            Put #1, , strFile
        Else
            Put #1, , arrFile
        End If
        Close #1
        
        bSuccess = True
        
    End If
    
ExitPoint:
    ' Clean up.
    Close #1
    Dim objErr As ErrObject
    Set objErr = Err
       
    saveBinBase64ToFile = bSuccess
    
    Handle_Error objErr, gstrObjectName, strFunctionName

End Function

Public Function CreateWordApplication(ByRef objWordApplication As Object, ByRef blnCreated As Boolean) As Object
    blnCreated = False
    
    If objWordApplication Is Nothing Then
        Dim objTempWordApplication As Object
        ' Create word application, note two word instances are created and one immediately closed,
        ' due to known Microsoft bug Q188546.
        Set objTempWordApplication = Create_Word_App(gstrObjectName)
        Set objWordApplication = Create_Word_App(gstrObjectName)
        ' See MSDN article Q259971.
        objWordApplication.DisplayAlerts = wdAlertsNone
        blnCreated = True
        Close_Word_App objTempWordApplication, gstrObjectName
    End If
    
    Set CreateWordApplication = objWordApplication
End Function


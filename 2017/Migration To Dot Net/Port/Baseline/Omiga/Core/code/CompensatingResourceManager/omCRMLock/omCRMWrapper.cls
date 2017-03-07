VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "omCRMWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCRMWrapper.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Created:       18/12/2003
'Author:        Ian Ross
'Description:   Wrapper component for the CRM.
'
'This component "wraps" the methods of the CRM Component.
'
'
'Dependencies:  omCRM --- The CRM is OK mixing CRM access with non CRM access to
'               the text file, but fails sometimes if we create multiple instances
'               of the worker/compensator. The Lock file in omCRMWrapper stops this
'               from happening.
'
'Issues:
'           Uses a mixture of FSO and Win32API calls for file work. My testing showed
'           no speed advantage of the Win32API for checking if files and folders exist
'           and the FSO was easier to implement and follow for this. The Win32API was
'           however, much faster for writing to the files and was required anyway to
'           implement the locking.
'
'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'INR    18/12/03    Created
'------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        AQR     Description
'MV     04/06/2004  BBG48 - Created
'MSla   23/06/2004  BBG48   Ammend AddText() - Replace CreateFolder() with call to
'                           EnsurePathExistsAndCreateFolder() and remove CreateFolder()as function is redundant.
'MV     06/08/2004  E2EM00000582    Amended LockFile()
'------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl

'=============================================
'API/DLL Procedure Declaration Section
'=============================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'=============================================
'Variable Declaration Section
'=============================================

'The following error code results when one tries to execute a method call
'on a COM+ component that has a transaction and has already aborted or
'in the process of aborting that transaction.  See ErrorHandler for details
Const CONTEXT_E_ABORTING = &H8004E003
Private gobjContext As ObjectContext

'Public Function AddText(ByVal strText As String, ByVal strPathAndFilename As String, _
'               ByVal nMaxRetries As Integer, ByVal nSleepMilliSecs As Integer) As String
Public Function AddText(ByVal vstrXmlIn As String, ByVal strText As String) As String

    Const cstrFunctionName As String = "AddText"
    Debug.Print cstrFunctionName
    
    On Error GoTo AddTexterr
    
    Dim strPathAndFilename As String
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    Dim bSuccess As Boolean
    Dim strPath As String
    Dim strFilename As String
    Dim strWriteMode As String
    Dim bAppend As Boolean
    Dim strLockPathAndFilename As String
    Dim hLockFile As Long
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTextNode As IXMLDOMNode
    Dim xmlInDoc As FreeThreadedDOMDocument
    Dim xmlOutDoc As FreeThreadedDOMDocument
    Dim objWorker As omCRM.omCRMWorker
    
    Set xmlInDoc = New FreeThreadedDOMDocument
    Set xmlOutDoc = New FreeThreadedDOMDocument
    
    xmlInDoc.loadXML vstrXmlIn
    Set xmlTextNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST/TEXTOUT")
    
    'Get strText, strPathAndFilename nMaxRetries and nSleepMilliSecs
    strPathAndFilename = xmlGetMandatoryAttributeText(xmlTextNode, "PATHANDFILENAME")
    nMaxRetries = xmlGetMandatoryAttributeText(xmlTextNode, "MAXRETRIES")
    nSleepMilliSecs = xmlGetMandatoryAttributeText(xmlTextNode, "SLEEPMILLISECS")
    strWriteMode = xmlGetMandatoryAttributeText(xmlTextNode, "WRITEMODE")
    
    If (StrComp(strWriteMode, "OverWrite") = 0) Then
        bAppend = False
    Else
        bAppend = True
    End If
        
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem
    
    'MSla   23/06/2004  BBG48
    'Dim fsoFind As FileSystemObject
    'Set fsoFind = New FileSystemObject
    'MSla   23/06/2004  BBG48 - End

    Dim iPos As Integer
    iPos = 0

    bSuccess = False
        
    'Path and text file we are writing to.
    iPos = InStrRev(strPathAndFilename, "\", , vbTextCompare)
    If iPos > 0 Then
        strPath = Left$(strPathAndFilename, iPos)
        strFilename = Right$(strPathAndFilename, Len(strPathAndFilename) - iPos)
    End If

    'MSla   23/06/2004  BBG48
    EnsurePathExistsAndCreateFolder strPath
    'If fsoFind.FolderExists(strPath) = False Then
    '    bSuccess = CreateFolder(strPath)
    'End If
    'MSla   23/06/2004  BBG48 - End

    strLockPathAndFilename = strPath & "Lock" & strFilename

    bSuccess = LockFile(strLockPathAndFilename, nMaxRetries, nSleepMilliSecs, hLockFile)
        
    If bSuccess Then
    
        'Also get an instance of the XMLWorker component we are wrapping:
        Set objWorker = gobjContext.CreateInstance("omCRM.omCRMWorker")
        objWorker.AddTextToFile strText, strPathAndFilename, bAppend, nMaxRetries, nSleepMilliSecs
    Else
        'Unable to open interface lock file
        errThrowError cstrFunctionName, oeCRMUnableToOpenFile, strLockPathAndFilename
         
    End If
    
    'Try & Unlock the file here (i.e. regardless of bSuccess) just in case,
    'really don't want to leave it locked
    UnLockFile hLockFile
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    AddText = xmlOutDoc.xml
   
    gobjContext.SetComplete
    
    'MSla   23/06/2004  BBG48
    'Set fsoFind = Nothing
    'MSla   23/06/2004  BBG48 - End
    
    Set objWorker = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTextNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing

    Exit Function

AddTexterr:
    
    AddText = errCreateErrorResponse()
    
    'Try & Unlock the file here (i.e. regardless of bSuccess) just in case,
    'really don't want to leave it locked
    UnLockFile hLockFile
    gobjContext.SetAbort
    
    'MSla   23/06/2004  BBG48
    'Set fsoFind = Nothing
    'MSla   23/06/2004  BBG48 - End
    
    Set objWorker = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTextNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
  
End Function

'Function to lock a file to prevent concurrent access.
Private Function LockFile(ByVal sFileName As String, ByVal nMaxRetries As Integer, _
                    ByVal nSleepMilliSecs As Integer, ByRef hLockFile As Long) As Boolean

    Const cstrFunctionName As String = "LockFile"
    Debug.Print cstrFunctionName
    
    On Error GoTo LockFileErr

    Dim bSuccess As Boolean
    bSuccess = False
    
    Dim nRetries As Integer
    nRetries = 0
    hLockFile = INVALID_HANDLE_VALUE
    
    Do While ((nRetries < nMaxRetries) And (hLockFile = INVALID_HANDLE_VALUE))
        'CreateFile call also locks the file OPEN_ALWAYS should create
        'the file if does not exist
        hLockFile = CreateFile(sFileName, _
                    GENERIC_WRITE, 0, 0, _
                    OPEN_ALWAYS, _
                    FILE_ATTRIBUTE_NORMAL, 0)
'                    OPEN_EXISTING,

        'Give someone else time to finish using the file before trying again
        'if we weren't able to lock the file
        If (hLockFile = INVALID_HANDLE_VALUE) Then
            Sleep (nSleepMilliSecs)
        End If

TryToLockAgain:
        nRetries = nRetries + 1
    Loop
    
    If ((nRetries < nMaxRetries) And (hLockFile <> INVALID_HANDLE_VALUE)) Then
        bSuccess = True
    Else
        'Unable to open interface lock file
        errThrowError cstrFunctionName, oeCRMUnableToOpenFile, sFileName & ": " & Err.Description
    End If

    If ((hLockFile <> INVALID_HANDLE_VALUE) And (bSuccess = False)) Then
        
        CloseHandle hLockFile
        hLockFile = INVALID_HANDLE_VALUE

    End If
        
    LockFile = bSuccess
    Exit Function

LockFileErr:
   
    If ((Err.Number = 70) And (nRetries < nMaxRetries)) Then
        'Permission denied, Go back and try at least nMaxRetries
        Resume TryToLockAgain
    End If
    
    UnLockFile hLockFile
   
    LockFile = False
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function UnLockFile(ByRef hLockFile As Long) As Boolean

    Const cstrFunctionName As String = "UnLockFile"
    Debug.Print cstrFunctionName
    
    On Error GoTo UnLockFileErr
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    bSuccess = CloseHandle(hLockFile)

    UnLockFile = bSuccess

    Exit Function

UnLockFileErr:
    
    UnLockFile = False
    errCheckError cstrFunctionName, TypeName(Me)

End Function


Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    adoBuildDbConnectionString
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub


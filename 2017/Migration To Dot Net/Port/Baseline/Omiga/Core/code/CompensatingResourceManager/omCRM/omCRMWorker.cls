VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "omCRMWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCRMWorker.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Created:       18/12/2003
'Author:        Ian Ross
'Description:   Worker component for the CRM.
'
'This component exposes the interface that a client sees and interacts with.
'It is responsible for associating itself with a compensator component that
'understands the log it writes, and can commit/abort the operations that it
'performs
'
'Dependencies:  omCRMCompensator
'
'Issues:
'           Uses a mixture of FSO and Win32API calls for file work. My testing showed
'           no speed advantage of the Win32API for checking if files and folders exist
'           and the FSO was easier to implement and follow for this. The Win32API was
'           however, much faster for writing to the files and was required anyway to
'           implement the locking. The CRM is OK mixing CRM access with non CRM access
'           to the text file, but fails sometimes if we create multiple instances
'           of the worker/compensator. The Lock file in omCRMWrapper stops this
'           from happening.
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
'MV     09/06/2004  BBG48 - Amended AddTextToFile() to include EnsurePathExistsAndCreateFolder()
'MSla   23/06/2004  BBG48   Removed CreateFolder() as function is redundant.
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
Private gobjContext As ObjectContext

'The worker component needs an instance of the CRM clerk, to associate itself
'with a compensator and to write its operations to the log:
Private Clerk As CRMClerk
Private CrmLogControl As ICrmLogControl

'The progID of the compensator this worker component is associated with:
Private Const COMPENSATOR_ProgID = "omCRM.omCRMCompensator"
Private Const COMPENSATOR_DESCRIPTION = "VB Compensator component"

'Indicates if the compensator has been registered:
Private bIsCompensatorRegistered As Boolean

'A Variant array that is used to write to the log:
Private vntLogFileEntries(4) As Variant

Private Sub Class_Initialize()
    Set gobjContext = GetObjectContext()
 
   bIsCompensatorRegistered = 0
End Sub

'This routine is called by the Initialize method. ; It
'obtains an instance of the Clerk for writing to the log.

Private Sub ObtainCRMClerk()
    Const cstrFunctionName As String = "ObtainCRMClerk"
    Debug.Print cstrFunctionName
    
    On Error GoTo ObtainCRMClerkErr
    Set Clerk = New CRMClerk
    Set CrmLogControl = Clerk
    Exit Sub

ObtainCRMClerkErr:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'This procedure associates this worker component with
'the compensator:
Private Sub RegisterCompensator()
    Const cstrFunctionName As String = "RegisterCompensator"
    Debug.Print cstrFunctionName

    'If we have already registered the compensator then
     'exit the subroutine:
     If bIsCompensatorRegistered = True Then
             Exit Sub
     End If
     'Register the compensator with the DTC.
     'There is the possibility that the Compensator is in the
     'the process of recovering from a previous shutdown, hence
     'the error checking loop:
     On Error Resume Next
     
     Do
          CrmLogControl.RegisterCompensator COMPENSATOR_ProgID, COMPENSATOR_DESCRIPTION, CRMREGFLAG_ALLPHASES
          DoEvents
     Loop Until Err.Number <> XACT_E_RECOVERYINPROGRESS

     'Was there an error registering the compensator?
     If Err.Number <> 0 Then GoTo RegisterCompensatorErr

     'Indicate that the compensator has been registered, so
     'subsequent calls to register the compensator abort:
     bIsCompensatorRegistered = True

     Exit Sub

RegisterCompensatorErr:
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Public Sub AddTextToFile(ByVal strText As String, _
        ByVal strPathAndFilename As String, ByVal bAppend As Boolean, _
        ByVal nMaxRetries As Integer, ByVal nSleepMilliSecs As Integer)
                                                
    Const cstrFunctionName As String = "AddTextToFile"
    Debug.Print cstrFunctionName
    
    On Error GoTo AddTextToFileErr
    
    'MSla 23/06/2004 BBG48
    'Dim fsoFind As FileSystemObject
    'Set fsoFind = New FileSystemObject
    'MSla 23/06/2004 BBG48 - End
    
    
    Dim strFileName As String
    Dim strPath As String
    Dim bSuccess As Boolean
    Dim hLockFile As Long
    Dim iPos As Integer
    iPos = 0
        
    'Path and text file we are writing to.
    iPos = InStrRev(strPathAndFilename, "\", , vbTextCompare)
    If iPos > 0 Then
        strPath = Left$(strPathAndFilename, iPos)
        strFileName = Right$(strPathAndFilename, Len(strPathAndFilename) - iPos)
    End If
    
    Initialize
    
    bSuccess = LockFile(strPathAndFilename, nMaxRetries, _
                            nSleepMilliSecs, hLockFile)

    vntLogFileEntries(0) = ADD_TEXT_COMMAND
    vntLogFileEntries(1) = strText
    vntLogFileEntries(2) = strPathAndFilename
    vntLogFileEntries(3) = hLockFile
    vntLogFileEntries(4) = bAppend
    
    'Write to the log file, and ensure it is durable:
    CrmLogControl.WriteLogRecordVariants vntLogFileEntries
    CrmLogControl.ForceLog
              
    'MSla 23/06/2004 BBG48
    'Redundant function call
    'If fsoFind.FolderExists(strPath) = False Then
        'CreateFolder strPath
    'End If
    
    'New function to create path is non-existant
    EnsurePathExistsAndCreateFolder strPath
    'MSla 23/06/2004 BBG48 - End
    
    gobjContext.SetComplete
    
    'MSla 23/06/2004 BBG48
    'Set fsoFind = Nothing
    'MSla 23/06/2004 BBG48 - End
    
    Exit Sub

AddTextToFileErr:

    'MSla 23/06/2004 BBG48
    'Set fsoFind = Nothing
    'MSla 23/06/2004 BBG48 - End
    
    gobjContext.SetAbort
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub Initialize()
    Const cstrFunctionName As String = "Initialize"

    Debug.Print cstrFunctionName

   'Have we obtained a ICrmLogControl interface yet?
   'If not obtain it from the CRMClerk:
    If CrmLogControl Is Nothing Then
       ObtainCRMClerk
    End If

   'Register the compensator with the DTC:
   Call RegisterCompensator

End Sub


'Function to lock a file to prevent concurrent access.
Private Function LockFile(ByVal strFileName As String, ByVal nMaxRetries As Integer, _
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
        'CreateFile call also locks the file OPEN_ALWAYS will create
        'the file if does not exist
        hLockFile = CreateFile(strFileName, _
                    GENERIC_WRITE, 0, 0, _
                    OPEN_ALWAYS, _
                    FILE_ATTRIBUTE_NORMAL, 0)

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
        errThrowError cstrFunctionName, oeCRMUnableToOpenFile, strFileName & ": " & Err.Description
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
    
    CloseHandle hLockFile
          
    LockFile = False
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


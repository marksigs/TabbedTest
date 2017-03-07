Attribute VB_Name = "omCRMGlobal"
'------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        Description
'MV     04/06/2004  BBG48 - Created
'MV     09/06/2004  BBG48 - Added New Method EnsurePathExistsAndCreateFolder to create the
'                   child directories recursively
'MV     09/06/2004  BBG48 - Amended EnsurePathExistsAndCreateFolder
'------------------------------------------------------------------------------------------

Option Explicit

'Constants shared by both the compensator and worker components:

Public Enum CRMERROR
    oeCRMUnableToCreateFile = 7015
    oeCRMUnableToOpenFile = 7016
    oeCRMUnableToCreateFolder = 7017
End Enum

'Error Code indicating that the compensator is in recovery mode:
Public Const XACT_E_RECOVERYINPROGRESS = &H8004D082

'Constants used when writing the log file:
Public Const ADD_TEXT_COMMAND = 1
Public Const REMOVE_TEXT_COMMAND = 2

'Constants used when deciding which Mutex to set up:
Public Const TT_MUTEX = 1
Public Const BACS_MUTEX = 2
Public Const HUNTER_MUTEX = 3

'Constants used when writing to text file with win32API
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const INVALID_HANDLE_VALUE = -1&
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100


'Used by SHFileOperation, allows copying of a locked file
Public Type SHFILEOPSTRUCT
  hwnd As Long
  'FO_COPY - Copies the files specified in the pFrom member
  'to the location specified in the pTo member
  wFunc As Long
  pFrom As String
  pTo As String
  'FOF_NOCONFIRMMKDIR - Does not confirm the creation of a
  'new directory if the operation requires one to be created.
  'FOF_NOCONFIRMATION - Respond with Yes to All for any dialog box that is displayed.
  'FOF_NOERRORUI - No user interface will be displayed if an error occurs.
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Public Const FO_COPY = &H2
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_NOERRORUI = &H400


Public Sub EnsurePathExistsAndCreateFolder(ByVal strPath As String)
'to create child directories recursively

On Error Resume Next

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    With fso
        
        If .FolderExists(strPath) = False Then
            
            If .GetParentFolderName(strPath) <> "" Then
                EnsurePathExistsAndCreateFolder .GetParentFolderName(strPath)
            End If
            
            .CreateFolder strPath
        
        End If
        
    End With
    
    Set fso = Nothing
    
End Sub





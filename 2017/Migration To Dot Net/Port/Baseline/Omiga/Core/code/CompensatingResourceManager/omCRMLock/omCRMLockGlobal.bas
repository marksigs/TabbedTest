Attribute VB_Name = "omCRMLockGlobal"
'------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        AQR     Description
'MV     04/06/2004  BBG48 - Created
'MSla   23/06/2004  BBG48   New function EnsurePathExistsAndCreateFolder()
'MV     12/07/2004  BBG981  Amended EnsurePathExistsAndCreateFolder() - Removed the msgbox
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
Public Const FILE_END = 2
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const INVALID_HANDLE_VALUE = -1&
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

'MSla   23/06/2004  BBG48
'--------------------------------------------------------------------------
' In Parameter: strPath as String
'
' Check to see if file path strPath exists
' If not, creates path
'--------------------------------------------------------------------------
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
'MSla   23/06/2004  BBG48 - End


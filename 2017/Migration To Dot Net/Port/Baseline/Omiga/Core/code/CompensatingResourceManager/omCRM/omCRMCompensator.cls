VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omCRMCompensator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCRMCompensator.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Created:       18/12/2003
'Author:        Ian Ross
'Description:   Compensator component for the CRM.
'
'This component is called when a transaction involving the omCRMWorker component
'aborts or commits.  Either one of two interfaces MUST be implemented by all
'compensator components:
'ICrmCompensator         - for unstructured storage in the log
'                           (C++ components)
'ICrmCompensatorVariants - for structured storage in the log
'                           (VB/Java):
'
'Dependencies:
'
'Issues:
'
'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'INR    18/12/03    Created
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        Description
'MV     04/06/2004  BBG48 - Created
'------------------------------------------------------------------------------------------
Option Explicit

'The worker component this compensator is associated with writes
'variants to the log, so this component implements the ICrmCompensatorVariants
'interface:

Implements ICrmCompensatorVariants

'=============================================
'API/DLL Procedure Declaration Section
'=============================================
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

'=============================================
'Variable Declaration Section
'=============================================

'The DTC gives the ICrmLogControl interface to the compensator
'in the SetLogControlVariants method so the compensator can access
'the log:

Private CrmLogControl As ICrmLogControl

'Variables to read from the log file:
Private vntCommand, vntOrigFilename, vntTempFilename, vntText, vnthLockFile

'AbortRecordVariants is called after BeginAbortVariants.
'This method delivers log records that were written by the
'worker.  The compensator must "undo" the changes indicated
'by such records.

Private Function ICrmCompensatorVariants_AbortRecordVariants(pLogRecord As Variant) As Boolean
'     MsgBox "In AbortRecordVariants"
   
    Const cstrFunctionName As String = "ICrmCompensatorVariants_AbortRecordVariants"
    On Error GoTo AbortRecordVariantProblem
    
    Dim iNode As Integer
    Dim hLockFile As Long
  
    'Obtain information from the logfile that indicates what was done:
    vntCommand = pLogRecord(0)
    vntText = pLogRecord(1)
    vntOrigFilename = pLogRecord(2)
    hLockFile = pLogRecord(3)
    
    'Just need to unlock the file
    UnLockFile hLockFile
    
    ICrmCompensatorVariants_AbortRecordVariants = True
    Exit Function

AbortRecordVariantProblem:
    UnLockFile hLockFile
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'BeginAbortVariants is called to let the compensator know that it
'must abort the current transaction. A call to this method is followed
'by a call to AbortRecordVariants, where the compensator recieves
'log records of the operations it must undo.  In this phase
'the compensator will do any preparatory work it has to, to undo these
'changes (this compensator does not have any such preparatory work)

Private Sub ICrmCompensatorVariants_BeginAbortVariants(ByVal bRecovery As Boolean)
'   MsgBox "In beginAbortVariants"

End Sub

'BeginCommitVariants is called to let the compensator know that the
'second phase of the two phase commit has been reached.  A call to this
'method is followed by a call to CommitRecordVariants, where the compensator
'receives log records of the operations is must commit.  In this phase
'the compensator will do any preparatory work it has to, to commit these changes
'(this compensator does not have any such preparatory work)

Private Sub ICrmCompensatorVariants_BeginCommitVariants(ByVal bRecovery As Boolean)
'   MsgBox "In BeginCommitVariants"
   

End Sub

'BeginPrepareVariants is called to let the compensator know that the
'first phase of the two phase commit has been reached.  A call to this
'method is followed by a call to PrepareRecordVariants, where the compensator
'receives log records of the operations is must prepare.  In this phase
'the compensator will do any preparatory work it has to, to prepare these changes
'(again, this compensator does not have any such preparatory work)

Private Sub ICrmCompensatorVariants_BeginPrepareVariants()
'   MsgBox "In BeginPrepareVariants"

End Sub

'CommitRecordVariants is called after BeginCommitVariants.
'This method delivers log records that were written by the
'worker.  The compensator must do whatever it has to do
'to make the changes specified in the log records permanent.

Private Function ICrmCompensatorVariants_CommitRecordVariants(pLogRecord As Variant) As Boolean
'    MsgBox "In CommitRecordVariants"
    Const cstrFunctionName As String = "ICrmCompensatorVariants_CommitRecordVariants"
    On Error GoTo CommitRecordVariantsProblem
    
    Dim hLockFile As Long
    Dim bSuccess As Boolean
    Dim lBytesWritten As Long
    Dim lPos As Long
    Dim strText As String
    Dim bAppend As Boolean
        
    'Obtain information from the logfile that indicates what needs to be done:
    vntCommand = pLogRecord(0)
    strText = pLogRecord(1)
    vntOrigFilename = pLogRecord(2)
    hLockFile = pLogRecord(3)
    bAppend = pLogRecord(4)
    
    'Determine what operation is to be performed and do it:
    If vntCommand = ADD_TEXT_COMMAND Then
        
        'Write our text to the file
        If bAppend = True Then
            lPos = SetFilePointer(hLockFile, 0, 0, FILE_END)
        Else
            lPos = SetFilePointer(hLockFile, 0, 0, FILE_BEGIN)
        End If
        
        bSuccess = WriteFile(hLockFile, strText, _
                             Len(strText), lBytesWritten, 0)
        'Check to see if you were successful writing the data
        If bSuccess <> 0 Then
           'Flush the file buffers to force writing of the data.
           bSuccess = FlushFileBuffers(hLockFile)
        End If
       
    End If
    
    ICrmCompensatorVariants_CommitRecordVariants = True
    UnLockFile hLockFile
   
    Exit Function

CommitRecordVariantsProblem:
    UnLockFile hLockFile
    ICrmCompensatorVariants_CommitRecordVariants = False
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'EndAbortVariants is called at the end of the abort phase, after
'AbortRecordVariants.This is the last method called during the
'abort phase.  This is where the compensator
'would do any cleanup operations associated with its entire abort phase.

Private Sub ICrmCompensatorVariants_EndAbortVariants()
'   MsgBox "In EndAbortVariants"

End Sub

'EndCommitVariants is called at the end of the commit phase, after
'ComitRecordVariants. This is the last method called during the
'commit phase.  This is where the compensator would do any cleanup
'operations associated with its entire commit phase.

Private Sub ICrmCompensatorVariants_EndCommitVariants()
'   MsgBox "In EndCommitVariants"
End Sub

'EndPrepareVariants is called at the end of the prepare phase, after
'PrepareRecordVariants. This is the last method called during the
'prepare phase.  This is where the compensator would do any cleanup
'operations associated with its entire prepare phase. In addition,
'this compensator returns a boolean in this method: if the prepare
' phase proceeded smoothly, the compensator indicates it is ready for
'the commit phase by returning true. If there was a problem during the
'prepare phase, the compensator returns false, at which point the
'transaction is aborted.

Private Function ICrmCompensatorVariants_EndPrepareVariants() As Boolean
'   MsgBox "In EndPrepareVariants"
   ICrmCompensatorVariants_EndPrepareVariants = True
End Function

'PrepareRecordVariants is called after BeginPrepareVariants.
'This method delivers log records that were written by the
'worker.  The compensator must do whatever it has to do
'to prepare the records for the commit phase of the transaction.

Private Function ICrmCompensatorVariants_PrepareRecordVariants(pLogRecord As Variant) As Boolean

'    MsgBox "In PrepareRecordVariants"

End Function

'SetLogControlVariants is called by the DTC to give the compensator
'an instance of ILogControl so it can access the log:

Private Sub ICrmCompensatorVariants_SetLogControlVariants(ByVal pLogControl As COMSVCSLib.ICrmLogControl)
'   MsgBox "In SetLogControlVariants"
    Set CrmLogControl = pLogControl
End Sub

Private Function UnLockFile(ByRef hLockFile As Long) As Boolean

Const cstrFunctionName As String = "UnLockFile"
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const START_ERROR_COUNT = vbObjectError + 100

Private m_nIcon As Integer
Private m_nErrorType As ErrorType

Public Enum ErrorType
    ErrorUser
    ErrorSystem
End Enum

Public Enum SupervisorErrorNumbers
    ' Global Parameter Errors
    errGeneralError = START_ERROR_COUNT
    ErrGlobalParamNotFound
    
    ' Password Errors
    ErrPasswordExists
    ErrPasswordIncorrect
    ' RecordSet Errors
    errRecordSetEmpty
    errRecordNotFound
    ErrKeyViolation
    ErrUserIDIsEmpty
    errUserNotSupervisor
    errInvalidPassword
    errPasswordNotFound
End Enum
Public Sub RaiseError(nErrNum As SupervisorErrorNumbers, Optional sDescription As String = "", Optional nErrorType As ErrorType = ErrorUser, Optional nIcon As Integer = vbCritical)
    m_nIcon = nIcon
    m_nErrorType = nErrorType
    
    Err.Raise nErrNum, , sDescription
End Sub
Public Sub DisplayError()
    Dim sError As String
    Dim nIcon As Integer
    
    If m_nErrorType = ErrorSystem Or Err.Number < START_ERROR_COUNT Then
        sError = Err.Source + ": Error is - " + Err.DESCRIPTION
    ElseIf Err.Number >= START_ERROR_COUNT Then
        
        If Len(Err.DESCRIPTION) > 0 Then
            sError = Err.DESCRIPTION + " - "
        End If
        
        Select Case Err.Number
        Case errRecordNotFound
            sError = sError + " Record not found"
        
        Case errRecordSetEmpty
            sError = sError + " Empty recordset returned"
        
        Case ErrKeyViolation
            sError = sError + " Record already exists - please enter a unique combination"
        
        Case ErrUserIDIsEmpty
            sError = sError + " User ID is empty"
        
        Case errUserNotSupervisor
            sError = sError + " User ID does not have suitable rights for Supervisor functions - check User Role and Active From/To dates"
        
        Case errInvalidPassword
            sError = sError + " Invalid password"
            
        Case errPasswordNotFound
            sError = sError + " No password records found for this user"
        
        Case ErrPasswordIncorrect
            sError = sError + " Password Incorrect "
        
        Case Else
            sError = Err.DESCRIPTION
        
        End Select

    End If

    MsgBox sError, m_nIcon
    EndWaitCursor
End Sub

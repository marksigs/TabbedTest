VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ErrorHandling
' Description   : Class that handles all Supervisor error handling
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port changes
' DJP       19/10/01    SYS2825 Some errors don't get displayed correctly in DisplayError
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Added
'                       errRateSetEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Constants
Private Const START_ERROR_COUNT = vbObjectError + 100

' Private data
Private m_frmError As Form
Private m_nIcon As Integer
Private m_nErrorType As ErrorType
Private m_nErrorNum As Long
Private m_sErrorDescription As String

' Enums
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
    errMandatoryFieldsRequired
    errKeysEmpty
    errIncorrectType
    errDontDisplayError
    errNotImplemented
    errPostCodeInvalid
    errCanNotUpdateBatchJob
    errClientSpecificNotFound
    errDatabaseNotSupported
    errEmptyLenderCode
    errPromoteFieldsDifferent
    errInvalidParameter
    errMissingPrimaryTag
    errOrganisationIDMissing
    errRateSetEmpty
End Enum
Public Sub SaveError(Optional nErrNum As Variant, Optional sDescription As Variant)
    Dim nErr As Long
    Dim sError As String
    
    If IsMissing(nErrNum) Then
        If Err.Number <> 0 Then
            nErr = Err.Number
            sError = Err.Description
        End If
    Else
        nErr = nErrNum
        If Not IsMissing(sDescription) Then
            sError = sDescription
        End If
    End If
    
    m_nErrorNum = nErr
    m_sErrorDescription = sError
    
End Sub
Public Sub SetErrorForm(frmError As Form)
    Set m_frmError = frmError
End Sub
Public Sub RaiseError(Optional nErrNum As SupervisorErrorNumbers, Optional sDescription As String = "", Optional nErrorType As ErrorType = ErrorUser, Optional nIcon As Integer = vbCritical)
    m_nIcon = nIcon

    m_nErrorType = nErrorType
    
    If nErrNum = 0 Then
        nErrNum = m_nErrorNum
    End If
    
    m_nErrorNum = 0
    
    If Len(sDescription) = 0 Then
        sDescription = m_sErrorDescription
    End If
    
    SaveError nErrNum, sDescription
    
    Err.Raise nErrNum, , sDescription
End Sub
Public Sub DisplayError(Optional sMessage As String)
    Dim sError As String
    Dim nIcon As Integer
    Dim bDisplay As Boolean
    
    bDisplay = True
    
    If m_nErrorNum <> 0 Then
        Err.Number = m_nErrorNum
        Err.Description = m_sErrorDescription
    End If
    
    If m_nErrorType = ErrorSystem Or Err.Number < START_ERROR_COUNT Then
        sError = Err.Source + ":  - " + Err.Description
    ElseIf Err.Number >= START_ERROR_COUNT Then
        
        ' DJP SYS2825
        If (Len(Err.Description) > 0) Then
            sError = Err.Description + " "
        End If
        
        Select Case Err.Number
        Case errOrganisationIDMissing
            sError = sError + "Organisation ID is missing"
        
        Case errRateSetEmpty
            sError = sError + "RateSet is missing"
        
        Case errRecordNotFound
            sError = sError + "Record not found"
        
        Case errRecordSetEmpty
            sError = sError + "Empty recordset returned"
        
        Case ErrKeyViolation
            sError = sError + "Record already exists - please enter a unique combination"
        
        Case ErrUserIDIsEmpty
            sError = sError + "User ID is empty"
        
        Case errUserNotSupervisor
            sError = sError + "User ID does not exist, or does not have suitable rights for Supervisor functions"
        
        Case errInvalidPassword
            sError = sError + "Invalid password"
            
        Case errPasswordNotFound
            sError = sError + "No password records found for this user"
        
        Case ErrPasswordIncorrect
            sError = sError + "Password Incorrect "
        
        Case errMandatoryFieldsRequired
            sError = sError + "Mandatory fields required"
        
        Case errKeysEmpty
            sError = sError + "Keys are empty"
        
        Case errIncorrectType
            sError = sError + " Data type is incorrect"
        
        Case errDontDisplayError
            bDisplay = False
        
        Case errNotImplemented
            sError = sError + " Not implemented"
        
        Case errPostCodeInvalid
            sError = sError + " Please enter a valid postcode"
        
        Case errCanNotUpdateBatchJob
            sError = sError + " Cannot Cancel the Batch Job"
        
        Case errClientSpecificNotFound
            sError = sError + " Client specific form class not set"
        
        Case errDatabaseNotSupported
            sError = sError + " Database type not supported"
        
        Case errEmptyLenderCode
            sError = sError + " Lender code is empty"
        
        Case errPromoteFieldsDifferent
            sError = sError + ". Database field count is different - unable to perform promotion"
        Case errInvalidParameter
            sError = sError + ". Invalid Parameter"
        
        Case errMissingPrimaryTag
            sError = sError + ". Primary tag is missing"
        Case Else
        End Select

    End If

    If bDisplay Then
        If Len(sMessage) = 0 Then
            sMessage = sError
            If Not m_frmError Is Nothing Then
                m_frmError.SetExtra ""
            End If
        Else
            If Not m_frmError Is Nothing Then
                m_frmError.SetExtra sError
            End If
        End If
        
        If Not m_frmError Is Nothing Then
            m_frmError.txtMessage = sMessage
            SetTopMostWindow m_frmError.hWnd, True
            m_frmError.Show vbModal

            Unload m_frmError
        Else
            If Len(sError) > 0 Then
                sMessage = sError
            End If
            
            'MsgBox sMessage, m_nIcon
        End If
    End If
    
    SaveError 0, ""
    
    EndWaitCursor
End Sub

Private Sub Class_Initialize()
    ' If the icon isn't set, use the default vbcritial
    m_nIcon = vbCritical
End Sub


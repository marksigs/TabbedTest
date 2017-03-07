VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditLifeCover 
   Caption         =   "Life Cover Rates Add/Edit"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "frmEditLifeCover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   6540
      TabIndex        =   10
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   1980
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   7
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   9
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   1
      Left            =   6300
      TabIndex        =   0
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   10
   End
   Begin MSGOCX.MSGComboBox cboCoverType 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   480
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Mandatory       =   -1  'True
      Text            =   ""
   End
   Begin MSGOCX.MSGComboBox cboApplicantGender 
      Height          =   315
      Left            =   6300
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Mandatory       =   -1  'True
      Text            =   ""
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   2
      Left            =   2100
      TabIndex        =   3
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   6
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MaxValue        =   ""
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   5
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   3
      Left            =   6300
      TabIndex        =   4
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   6
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "1"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   5
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   4
      Left            =   2100
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "0"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   9
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   5
      Left            =   6300
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "0"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   9
   End
   Begin MSGOCX.MSGEditBox txtLifeCover 
      Height          =   315
      Index           =   6
      Left            =   2100
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "0"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   9
   End
   Begin VB.Label Label3 
      Caption         =   "Maximum Applicant Age"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Poor Health Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Additional Smokers Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Annual Rate"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Max Term of Loan (mths)"
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Applicants Gender"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Cover Type"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Life Cover Rate Number"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditLifeCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditLifeCover
' Description   :   Form which allows the user to edit and add LifeCover rates.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Changed call
'                       to NextNumber as it has been moved from FormProcessing to DataAccess.
' STB       06/12/01    SYS2522 - Clicking on the Another button now generates a new Life Cover
'                       Rate Number.
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Control Indexes.
Private Const LIFE_COVER_NUMBER As Long = 0
Private Const START_DATE        As Long = 1
Private Const MAX_APPLICANT_AGE As Long = 2
Private Const MAX_TERM_OF_LOAN  As Long = 3
Private Const SMOKERS_RATE      As Long = 4
Private Const ANNUAL_RATE       As Long = 5
Private Const HEALTH_RATE       As Long = 6

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The underlying table object.
Private m_clsLifeCover As New LifeCoverRatesTable

'The Life Cover Rate number.
Private m_vNextNumber As Variant

'A status indicator to the forms caller.
Private m_ReturnCode As MSGReturnCode

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Set the edit/add state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the current state of the form (add or edit).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboApplicantGender_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboApplicantGender_Validate(Cancel As Boolean)
    Cancel = Not cboApplicantGender.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboCoverType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCoverType_Validate(Cancel As Boolean)
    Cancel = Not cboCoverType.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form, the return status will indicate failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate data and attempt to save. If successful, close the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    'Call into the shared routine to save the record(s).
    bRet = DoOKProcessing()

    'If successful, hide this form and return processing to the opener.
    If bRet = True Then
        Hide
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Populate combos, prepare the underlying objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Default return status is failure.
    SetReturnCode MSGFailure
    
    'Create an underlying table object.
    Set m_clsLifeCover = New LifeCoverRatesTable
    
    ' Setup combos
    g_clsFormProcessing.PopulateCombo "LifeCoverType", cboCoverType
    g_clsFormProcessing.PopulateCombo "LifeCoverGender", cboApplicantGender
    
    'Setup the controls according to the forms state.
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Load the desired record into the table object and populate the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    'Cast a generic table interface onto the underlying table object.
    Set clsTableAccess = m_clsLifeCover
    
    'Set the keys collection to the one passed in from MainSuper.
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    'Load the record and obtain a reference to the ADO recordset.
    Set rs = clsTableAccess.GetTableData()
    
    'Another is disabled (we're editing).
    cmdAnother.Enabled = False
    
    'Ensure a record was loaded.
    ValidateRecordset rs, "Life Cover Rates "
    
    'Double ensure? - the populate the screen with the data.
    If rs.RecordCount = 1 Then
        'We can't edit the primary key if we're in an edit state.
        Me.txtLifeCover(LIFE_COVER_NUMBER).Enabled = False
        
        PopulateScreenFields
    Else
        g_clsErrorHandling.RaiseError errRecordNotFound
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populate each screen control from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PopulateScreenFields() As Boolean
    
    Dim vTmp As Variant
    
    On Error GoTo Failed

    'Product Number
    m_vNextNumber = m_clsLifeCover.GetLifeCoverRateNumber()
        
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        ' Life cover rate number
        txtLifeCover(LIFE_COVER_NUMBER).Text = m_vNextNumber
        
        ' Start Date
        vTmp = m_clsLifeCover.GetStartDate()
        g_clsFormProcessing.HandleDate txtLifeCover(START_DATE), vTmp, SET_CONTROL_VALUE
        
        ' Cover Type
        vTmp = m_clsLifeCover.GetCoverType()
        g_clsFormProcessing.HandleComboExtra cboCoverType, vTmp, SET_CONTROL_VALUE
        
        ' Gender
        vTmp = m_clsLifeCover.GetApplicantGender()
        g_clsFormProcessing.HandleComboExtra Me.cboApplicantGender, vTmp, SET_CONTROL_VALUE
        
        ' Max Applicant Age
        txtLifeCover(MAX_APPLICANT_AGE).Text = m_clsLifeCover.GetApplicantAgeMax()
        
        ' Max Applicant Age
        txtLifeCover(MAX_TERM_OF_LOAN).Text = m_clsLifeCover.GetTermMax()
        
        ' Smoker rate
        txtLifeCover(SMOKERS_RATE).Text = m_clsLifeCover.GetSmokerRate()
        
        ' Annual Rate
        txtLifeCover(ANNUAL_RATE).Text = m_clsLifeCover.GetAnnualRate()
        
        ' Poor Health Rate
        txtLifeCover(HEALTH_RATE).Text = m_clsLifeCover.GetNotGoodHealthRate()
    End If
    
    PopulateScreenFields = True
    
    Exit Function
    
Failed:
    PopulateScreenFields = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Copy all the control values into the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsLifeCover
    
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        'Life Cover Rate Number
        m_clsLifeCover.SetLifeCoverRateNumber m_vNextNumber
        
        ' Start Date
        g_clsFormProcessing.HandleDate txtLifeCover(START_DATE), vTmp, GET_CONTROL_VALUE
        m_clsLifeCover.SetStartDate vTmp
        
        ' Cover Type
        g_clsFormProcessing.HandleComboExtra cboCoverType, vTmp, GET_CONTROL_VALUE
        m_clsLifeCover.SetCoverType vTmp
        
        ' Gender
        g_clsFormProcessing.HandleComboExtra Me.cboApplicantGender, vTmp, GET_CONTROL_VALUE
        m_clsLifeCover.SetApplicantGender vTmp
        
        ' Max Applicant Age
        m_clsLifeCover.SetApplicantAgeMax txtLifeCover(MAX_APPLICANT_AGE).Text
        
        ' Max Applicant Age
        m_clsLifeCover.SetTermMax txtLifeCover(MAX_TERM_OF_LOAN).Text
        
        ' Smoker rate
        m_clsLifeCover.SetSmokerRate txtLifeCover(SMOKERS_RATE).Text
        
        ' Annual rate
        m_clsLifeCover.SetAnnualRate txtLifeCover(ANNUAL_RATE).Text
        
        ' Poor Health Rate
        m_clsLifeCover.SetNotGoodHealthRate txtLifeCover(HEALTH_RATE).Text
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "LifeCoverRates - Next number is empty"
    End If
    
    clsTableAccess.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Save the current record and re-initialise the form to add another record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Call into the shared routine to save the record(s).
    bRet = DoOKProcessing()
        
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        'If the record was saved, clear the screen fields ready for another record.
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If
    
    If bRet = True Then
        'Ensure the focus is back in the first field.
        Me.txtLifeCover(START_DATE).SetFocus
        
        'Re-initialise the form state, this will generate a life rate number.
        SetAddState
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateNewRecord
' Description   : Creates a blank record in the underlying table object and generates a new
'                 Life Rate number (the primary key).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateNewRecord()
    
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed

    'Cast a generic interface onto the underlying table object.
    Set clsTableAccess = m_clsLifeCover
    
    'Create a new record.
    g_clsFormProcessing.CreateNewRecord m_clsLifeCover
    
    'Obtain a new Life Cover Rate number.
    m_vNextNumber = m_clsLifeCover.NextCoverRate
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Validate and save the current record then route into the promotion code. This
'                 routine returns true if successful. Note: This routine is present on forms
'                 with an 'Another' button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    'Ensure all mandatory fields have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    'If they have then proceed.
    If bRet = True Then
        'Save the data.
        SaveScreenData
        
        'Ensure the record is flagged for promotion.
        SaveChangeRequest
        
        'Set the return status to success.
        SetReturnCode
    End If

    'Return True (success) or False (failure) to the caller.
    DoOKProcessing = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Ensures the collection of primary key values is set against the underlying
'                 table object and then passes it into the promotion code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
        
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    'Create a new collection.
    Set colMatchValues = New Collection
            
    'Add the Life Cover Rate number (primary key) into it.
    colMatchValues.Add txtLifeCover(LIFE_COVER_NUMBER).Text
    
    'Cast a generic interface onto the table object.
    Set clsTableAccess = m_clsLifeCover
    
    'Set the keys collection.
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    'Call into the promotion code to ensure the update marked for promotion.
    g_clsHandleUpdates.SaveChangeRequest m_clsLifeCover
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Creates a blank record, assigns a new Life rate number and populates the
'                 screen control with its value.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    
    On Error GoTo Failed
    
    'Ensure the life rate number is cleared.
    m_vNextNumber = Null
    
    'Create a new record and generate a new Life Rate Number.
    CreateNewRecord
    
    'Set the Life Cover Number.
    Me.txtLifeCover(LIFE_COVER_NUMBER).Text = CStr(m_vNextNumber)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtLifeCover_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLifeCover_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtLifeCover(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return status of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Returns the sucess/fail status to the form's caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

VERSION 5.00
Object = "{5F540CC8-EA22-4F95-9EFE-BDB4E09F976D}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditMPMigRates 
   Caption         =   "Add / Edit HLC Rate Sets"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5700
      TabIndex        =   4
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6990
      TabIndex        =   5
      Top             =   5340
      Width           =   1215
   End
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Set Bands"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgMIGRates 
         Height          =   3855
         Left            =   420
         TabIndex        =   3
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSGOCX.MSGEditBox txtMIGRates 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
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
   End
   Begin MSGOCX.MSGEditBox txtMIGRates 
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      Mandatory       =   -1  'True
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin VB.Label lblBaseRates 
      Caption         =   "Description"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   7
      Top             =   160
      Width           =   915
   End
   Begin VB.Label lblBaseRates 
      Caption         =   "Set Code"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmEditMPMigRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditMPMIGRates
' Description   : Form that controls editing of a mig rate set - that is, a Set of Base Rates as
'                 defined by their rate number
'
' Change history
' Prog      Date        Description
' AW       16/05/02    BM087 - Created
' AW        27/05/02    Added SaveChangeRequest()
' GD        02/07/2002  BMIDS00080 Changes error message in ValidateScreenData
' JD        30/03/2005  BMIDS982 Changed screen text from MIG to HLC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Text indexes
Private Const BANDED_FEE_SET = 0
Private Const BANDED_FEE_DESC = 1

' Private data
Private m_sOriginalSet As String
Private m_sOriginalDesc As String
Private m_bIsEdit As Boolean

Private m_clsMIGBandTable As MPMigRateTable
Private m_clsMIGSetTable As MPMigRateSetTable
Private m_clsTableAccess As TableAccess

Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection

Private Const EFFECT_DATE_COL As Integer = 1
Private Const END_DATE_COL As Integer = 2
Private Const HIGH_LTV_COL As Integer = 4
Private Const LOW_LTV_COL As Integer = 3
Private Const HIGHER_LOAN_COL As Integer = 6
Private Const LOWER_LOAN_COL As Integer = 5

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Form_Initialize()
    Set m_colKeys = New Collection
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Initialisation to this screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    Set m_clsTableAccess = New MPMigRateTable
    Set m_clsMIGSetTable = New MPMigRateSetTable
    dgMIGRates.Enabled = False
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    m_sOriginalSet = txtMIGRates(BANDED_FEE_SET).Text
    m_sOriginalDesc = txtMIGRates(BANDED_FEE_DESC).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Checks to see if a record exists already with the same
'                   keys. Returns true if it exists, false if not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim sStartDate As String
    
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    Dim clsMigFeeSet As MPMigRateSetTable
    
    sSet = txtMIGRates(BANDED_FEE_SET).Text
    sStartDate = txtMIGRates(BANDED_FEE_DESC).Text
    
    If Len(sSet) > 0 And Len(sStartDate) > 0 Then
    
        Set clsTableAccess = m_clsMIGSetTable
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        bRet = clsTableAccess.DoesRecordExist(col)
        
        If bRet = True Then
            MsgBox "Set already exist - please enter a unique combination", vbCritical
            txtMIGRates(BANDED_FEE_SET).SetFocus
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Set must be valid", vbCritical
    End If
    
    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user pressed the OK button. Performs necessary
'                   validation and saves any data that needs to be saved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing()
    
    If bRet = True Then
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sSet As String
    Dim sDate As String
    
    Set colMatchValues = New Collection
    sSet = Me.txtMIGRates(BANDED_FEE_SET).Text
    'sDate = Me.txtAdminFees(BANDED_DATE).Text
    colMatchValues.Add sSet
    'colMatchValues.Add sDate
    
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok or
'                   presses Another
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()
    
    If bRet = True Then
        'g_clsDataAccess.BeginTrans
        ' Do the banded updates, if there are any
        'DoUpdates
        ' Update the set table if needed
        DoSetUpdates
        ' Update the AdminFeeSet table.
        m_clsTableAccess.Update
        
        SaveChangeRequest
        SetReturnCode
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    'g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoUpdates
' Description   :   Updates the banded MigRate. The key is the start date
'                   and set number, so if they have changed we need to update
'                   all records with the new values.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoSetUpdates
' Description   :   Need to see if the SetID the user has typed in exists already
'                   in the MigRateSet table. If it does, do nothing, because there will
'                   already be entries for it in the MigRate table. If it doesn't exist, we
'                   need to put an entry into the MigRateSet table first.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoSetUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String, sDesc As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim clsMigFeeSet As MPMigRateSetTable
    
    Dim col As New Collection
    
    sSet = txtMIGRates(BANDED_FEE_SET).Text
    sDesc = txtMIGRates(BANDED_FEE_DESC).Text

    If Len(sSet) > 0 Then

        Set clsTableAccess = m_clsMIGSetTable
        
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        
        bRet = clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess
        
            m_clsMIGSetTable.SetMPMigFeeSet sSet
            m_clsMIGSetTable.SetMPMigDescription sDesc
            clsTableAccess.Update
        Else
            m_clsMIGSetTable.SetMPMigFeeSet sSet
            m_clsMIGSetTable.SetMPMigDescription sDesc
            
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   dgMIGRates_BeforeAdd
' Description   :   Called before the Add button is pressed on the datagrid
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgMIGRates_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgMIGRates.ValidateRow(nCurrentRow)
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupDatagrid
' Description   :   Setsup the datagrid by populating it and setting the fields
'                   to be used
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupDataGrid()
    On Error GoTo Failed
    Dim colTables As New Collection
    Dim colDataControls As New Collection
    
    colTables.Add m_clsTableAccess
    colDataControls.Add dgMIGRates
    
    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If

    If m_bIsEdit = True Then
        If m_clsTableAccess.RecordCount > 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate current record"
        End If
    End If

    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new Admin Fee Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    
    txtMIGRates(BANDED_FEE_DESC).Enabled = True
    txtMIGRates(BANDED_FEE_SET).Enabled = True
    
    SetupDataGrid
    
    'txtMIGRates(BANDED_FEE_SET).SetFocus
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing an Admin Fee Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    On Error GoTo Failed
    
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDataGrid
    
    txtMIGRates(BANDED_FEE_SET).Enabled = False
    txtMIGRates(BANDED_FEE_DESC).Enabled = True
    
    dgMIGRates.Enabled = True
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetGridFields
' Description   :   Sets the field names for the grid, specifies which
'                   are mandatory and which are banded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    Dim clsMigRateTable As MPMigRateTable
    Dim sVersionField As String
    Set clsMigRateTable = m_clsTableAccess
    
    fields.sField = "MPMIGRateSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtMIGRates(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    fields.sField = "EffectiveDate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sError = ""
    fields.sTitle = "Effective Date"
    fields.bDateField = True
    colFields.Add fields
    
    fields.sField = "EndDate"
    fields.bRequired = False
    fields.bVisible = True
    fields.sError = ""
    fields.sTitle = "End Date"
    fields.bDateField = True

    colFields.Add fields
    
    fields.sField = "LowerLTV"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Lower LTV"
    fields.bDateField = False
    
    colFields.Add fields
    
    fields.sField = "HigherLTV"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Higher LTV must be entered"
    fields.sTitle = "Higher LTV"
    fields.bDateField = False
    
    colFields.Add fields
    
    
    fields.sField = "LowerLoanAmount"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Lower Loan Amount"
    fields.bDateField = False
    colFields.Add fields
    
    fields.sField = "HigherLoanAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Higher Loan Amount"
    fields.bDateField = False
    fields.sError = "Higher Loan Amount must be entered"
    colFields.Add fields
    
    fields.sField = "MinimumPremium"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Minimum Premium"
    fields.bDateField = False
    colFields.Add fields

    fields.sField = "IndemnityRate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Indemnity Rate"
    fields.bDateField = False
    fields.sError = "Indemnity Rate must be entered"
    colFields.Add fields
    
    If g_clsVersion.DoesVersioningExist() Then
        ' Versioning...
        fields.sField = clsMigRateTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        
        colFields.Add fields
    End If
    
    dgMIGRates.SetColumns colFields, "EditMPMigFees", "HLC Fees"
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets the date and set number
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String, sKey As String
    
    Dim clsMigSetTable As MPMigRateSetTable
    Set clsMigSetTable = m_clsMIGSetTable
    
    Dim clsMigBandTable As MPMigRateTable
    Set clsMigBandTable = m_clsTableAccess

    sKey = clsMigBandTable.GetMPMigRateSet()
    
    Set rs = clsMigSetTable.GetMIGRateSetDesc(sKey)

    If rs.RecordCount = 1 Then
        txtMIGRates(BANDED_FEE_SET).Text = rs!MPMIGRateSet
        txtMIGRates(BANDED_FEE_DESC).Text = rs!MigrateSetDescription
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    ' If we're adding, check that this record doesn't exist already
    If bRet And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet Then
        bRet = dgMIGRates.ValidateRows()
    End If

    If bRet Then
        bRet = ValidateGridFields
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            ' GD        02/07/2002  BMIDS00080 Changes error message in ValidateScreenData
            g_clsErrorHandling.RaiseError errGeneralError, "Effective Date, Higher LTV and Higher Loan Amount must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   txtMIGRates_Validate
' Description   :   Called when the text boxes lose focus. Checks for mandatory
'                   requirments and for the entry to be valid
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtMIGRates_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtMIGRates(Index).ValidateData()

    If Cancel = False Then
        SetDataGridState
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetDataGridState
' Description   :   Enables or disables the datagrid based on other fields
'                   that need data to be entered.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDataGridState()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    ' Make sure all data that needs to be entered, has been
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgMIGRates.Enabled
        dgMIGRates.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtMIGRates(BANDED_FEE_SET).Text
                m_sOriginalDesc = txtMIGRates(BANDED_FEE_DESC).Text

                dgMIGRates.AddRow
                dgMIGRates.SetFocus
            End If
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   AW      16/05/02    BM023
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateGridFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    Dim nRowIndex As Integer
    Dim strEffectDate As String, strEndDate As String, strHighLTV As String
    Dim strLowLTV As String, strHighLoan As String, strLowLoan As String

    ValidateGridFields = False
    
    'Iterate through all the rows of the grid control.
    For nRowIndex = 0 To (dgMIGRates.Rows - 1)

        strEffectDate = dgMIGRates.GetAtRowCol(nRowIndex, EFFECT_DATE_COL)
        strEndDate = dgMIGRates.GetAtRowCol(nRowIndex, END_DATE_COL)
        strHighLTV = dgMIGRates.GetAtRowCol(nRowIndex, HIGH_LTV_COL)
        strLowLTV = dgMIGRates.GetAtRowCol(nRowIndex, LOW_LTV_COL)
        strHighLoan = dgMIGRates.GetAtRowCol(nRowIndex, HIGHER_LOAN_COL)
        strLowLoan = dgMIGRates.GetAtRowCol(nRowIndex, LOWER_LOAN_COL)

        If Len(strEndDate) > 0 Then
            If ValidateDate(strEffectDate, strEndDate) = False Then
            
                'Place the cursor in the invalid field.
                dgMIGRates.col = EFFECT_DATE_COL
                dgMIGRates.Row = nRowIndex
                dgMIGRates.SetGridFocus
    
                g_clsErrorHandling.RaiseError errGeneralError, "The Effective Date cannot exceed the End Date."
                
            End If
        End If

        If Val(strHighLTV) < Val(strLowLTV) Then

            'Place the cursor in the invalid field.
            dgMIGRates.col = LOWER_LOAN_COL
            dgMIGRates.Row = nRowIndex
            dgMIGRates.SetGridFocus

            g_clsErrorHandling.RaiseError errGeneralError, "The Lower LTV cannot exceed the Higher LTV."
        End If
        
        If Val(strHighLoan) < Val(strLowLoan) Then

            'Place the cursor in the invalid field.
            dgMIGRates.col = HIGHER_LOAN_COL
            dgMIGRates.Row = nRowIndex
            dgMIGRates.SetGridFocus

            g_clsErrorHandling.RaiseError errGeneralError, "The Lower Loan Amount cannot exceed the Higher Loan Amount."
        End If
    
    Next nRowIndex
       
    ValidateGridFields = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

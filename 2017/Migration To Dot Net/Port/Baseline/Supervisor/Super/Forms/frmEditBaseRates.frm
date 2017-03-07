VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBaseRates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Rate Set Add/Edit"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmEditBaseRates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGEditBox txtDescription 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
   End
   Begin MSGOCX.MSGDataCombo cboBaseRate 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   -1  'True
   End
   Begin MSGOCX.MSGEditBox txtIntRate 
      Height          =   315
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
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
   Begin VB.CommandButton cmdAnother 
      Caption         =   "A&nother"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6060
      TabIndex        =   7
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   6
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   5580
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtBaseRates 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TextType        =   6
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
   End
   Begin MSGOCX.MSGEditBox txtBaseRates 
      Height          =   315
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   60
      Width           =   1090
      _ExtentX        =   1931
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
   Begin MSGOCX.MSGDataGrid dgBaseRates 
      Height          =   3855
      Left            =   360
      TabIndex        =   4
      Top             =   1560
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
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblIntRate 
      Caption         =   "Interest Rate"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblBaseRate 
      Caption         =   "Base Rate"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblBaseRates 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblBaseRates 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditBaseRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditBaseRAtes
' Description   :   Form which allows the user to edit Base Rates. This is a banded
'                   form, so has a datagrid to enter the details.
' Change history
' Prog      Date        Description
' STB       06/12/01    SYS1942 - Another button commits current transaction.
' SA        17/12/01    SYS3504 Various changes to select a base rate against Base Rate Set
' DJP       28/01/02    SYS3933 Base rates cannot be created when no base rates exist already.
' STB       08/03/02    SYS4238 Removed the Rate column as it has been replaced by Rate Difference.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        02/07/2002  BMIDS00081 - application of core AQR SYS4935

Option Explicit

' Local text indexes
Private Const BANDED_FEE_SET = 0
Private Const BANDED_DATE = 1

' Module variables
Private m_sOriginalSet As String
Private m_sOriginalStartDate As String

Private Const FEE_SET_COL = 0
Private Const START_DATE_COL = 1
Private Const MAX_LOAN_AMOUNT = 2
Private Const MAX_LTV = 3

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_colKeys As Collection
'SYS3504
Private m_clsRateTable As RateTable
Private m_rsBaseRates As ADODB.Recordset
Private m_clsBaseRateSet As BaseRateSetTable

    
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
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok or
'                   presses Another
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()
    
    If bRet = True Then
        DoUpdates
        DoSetUpdates
        m_clsTableAccess.Update
        SaveChangeRequest
        SetReturnCode
    End If
    DoOKProcessing = bRet
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sSet As String
    Dim sDate As String
    
    Set colMatchValues = New Collection
    sSet = txtBaseRates(BANDED_FEE_SET).Text
    sDate = txtBaseRates(BANDED_DATE).Text
    
    colMatchValues.Add sSet
    colMatchValues.Add sDate
    
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'SYS3504 Move the recordset on manually to get Rate to display
Private Sub cboBaseRate_Click(Area As Integer)
    If cboBaseRate.SelectedItem <> "" Then
        m_rsBaseRates.Move cboBaseRate.SelectedItem - 1, 1
        txtIntRate.Text = m_rsBaseRates("BaseInterestRate").Value
    End If
End Sub
Private Sub cboBaseRate_Validate(Cancel As Boolean)
    Cancel = Not cboBaseRate.ValidateData()
End Sub

''SYS3504
Private Sub txtdescription_Validate(Cancel As Boolean)
    Cancel = Not txtDescription.ValidateData()
    If Cancel = False Then
        SetDataGridState
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user presses the Another button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing()
        
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
    
        g_clsFormProcessing.ClearScreenFields Me
        Form_Load
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
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
    Dim bRet As Boolean
    Dim sSet As String
    Dim sStartDate As String
    
    Dim col As New Collection
    
    sSet = txtBaseRates(BANDED_FEE_SET).Text
    sStartDate = txtBaseRates(BANDED_DATE).Text
    
    col.Add sSet
    col.Add sStartDate
    
    If g_clsVersion.DoesVersioningExist() Then
        col.Add g_sVersionNumber
    End If
    
    bRet = m_clsTableAccess.DoesRecordExist(col)
    
    If bRet = True Then
        MsgBox "Name and Start Date already exist - please enter a unique combination", vbCritical
        txtBaseRates(BANDED_FEE_SET).SetFocus
    End If
    DoesRecordExist = bRet
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoUpdates
' Description   :   Updates the banded Base Rates. The key is the start date
'                   and set number, so if they have changed we need to update
'                   all records with the new values.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    Dim colValues As New Collection
    Dim sSetID As String
    Dim sStartDate As String
    
    Set clsBandedTable = m_clsTableAccess
    
    sStartDate = txtBaseRates(BANDED_DATE).Text
    sSetID = txtBaseRates(BANDED_FEE_SET).Text
    
    If Len(sStartDate) > 0 And Len(sSetID) > 0 Then
        If sStartDate <> m_sOriginalStartDate Or sSetID <> m_sOriginalSet Then
            colValues.Add sSetID
            colValues.Add sStartDate
        
            clsBandedTable.SetUpdateValues colValues
        
            ' We only want to check that the record exists if the name or start date has changed
            clsBandedTable.SetUpdateSets
            clsBandedTable.DoUpdateSets
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoSetUpdates
' Description   :   Need to see if the SetID the user has typed in exists
'                   already in the BaseRateSet table. If it does, do nothing, because
'                   there will already be entries for it in the BaseRateBand table.
'                   If it doesn't exist, we need to put an entry into the BaseRateBand table first.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoSetUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    Dim col As New Collection
    
    sSet = txtBaseRates(BANDED_FEE_SET).Text

    If Len(sSet) > 0 Then
        Set clsTableAccess = m_clsBaseRateSet
        
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        
        bRet = clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess
            m_clsBaseRateSet.SetFeeSet sSet
            'SYS3504 Store rateid and description
            m_clsBaseRateSet.SetRateID cboBaseRate.SelText
            m_clsBaseRateSet.SetDescription txtDescription.Text
            
            clsTableAccess.Update
        'SYS3504 Record exists, so need to update it.
        Else
            clsTableAccess.GetTableData POPULATE_KEYS
            m_clsBaseRateSet.SetDescription txtDescription.Text
            m_clsBaseRateSet.SetRateID cboBaseRate.SelText
            m_clsBaseRateSet.SetRateStartDate txtBaseRates(BANDED_DATE).Text
            
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub dgBaseRates_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgBaseRates.ValidateRow(nCurrentRow)
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Initialisation to this screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New BaseRateTable
    
    'SYS3504
    Set m_clsRateTable = New RateTable
    Set m_clsBaseRateSet = New BaseRateSetTable
    
    dgBaseRates.Enabled = False
      
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
   
    m_sOriginalSet = txtBaseRates(BANDED_FEE_SET).Text
    m_sOriginalStartDate = txtBaseRates(BANDED_DATE).Text
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupDataGrid
' Description   : Bind the BaseRateBand table to the data grid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupDataGrid()
    '' GD        02/07/2002  BMIDS00081 - application of core AQR SYS4935
    Dim clsBandTable As BaseRateTable
    
    On Error GoTo Failed
        
    Set clsBandTable = m_clsTableAccess
    
    'Populate the BaseRateBand table with bands for this set (colKeys is already set).
    If m_bIsEdit Then
        'SYS4935 Send in start date as well.
        clsBandTable.GetBandsForSet m_colKeys(1), m_colKeys(2)
    Else
        clsBandTable.GetBandsForSet
    End If
    
    'Bind the grid to the underlying recordset.
    Set dgBaseRates.DataSource = m_clsTableAccess.GetRecordSet
        
    If m_bIsEdit = True Then
        PopulateScreenFields
    End If
    
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new Base Rate Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed
    
    SetupDataGrid
    
    If txtBaseRates(BANDED_FEE_SET).Visible = True Then
        txtBaseRates(BANDED_FEE_SET).SetFocus
    End If
           
    'SYS3504
    PopulateBaseRates
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' SYS3504
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateBaseRates
' Description   : Populates the Base Rate Combo with a list of Lenders. Need to read the Rate Id's
'                 from the BaseRate table class. The Combo is a data combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateBaseRates()
    On Error GoTo Failed
    Dim sRateID As String
        
    Set m_rsBaseRates = m_clsRateTable.GetBaseRateData
    
    If Not m_rsBaseRates Is Nothing Then
        Set cboBaseRate.RowSource = m_rsBaseRates
        sRateID = m_clsRateTable.GetRateIdField()
        cboBaseRate.ListField = sRateID
    End If
               
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing an Base Rate Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
    'SYS3504 Populate from BaseRateSet table
    Dim colBaseRateSets As New Collection
    colBaseRateSets.Add m_colKeys.Item(1)
    TableAccess(m_clsBaseRateSet).SetKeyMatchValues colBaseRateSets
    TableAccess(m_clsBaseRateSet).GetTableData POPULATE_KEYS
      
    txtDescription.Text = m_clsBaseRateSet.GetDescription
    m_clsTableAccess.SetKeyMatchValues m_colKeys
        
    SetupDataGrid
    
    ' SYS3504 Now get the relevant base rate
    'cboBaseRate.SelText = m_clsBaseRateSet.GetRateID
    cboBaseRate.Text = m_clsBaseRateSet.GetRateID
    If cboBaseRate.SelectedItem <> "" Then
        m_rsBaseRates.MoveFirst
        m_rsBaseRates.Move cboBaseRate.SelectedItem - 1
        txtIntRate.Text = m_rsBaseRates("BaseInterestRate").Value
    End If
    
    txtBaseRates(BANDED_FEE_SET).Enabled = False
    'SYS3504 All fields should be editable if start date hasn't past yet.
    If CDate(txtBaseRates(BANDED_DATE).Text) < Now Then
        txtBaseRates(BANDED_DATE).Enabled = False
        txtDescription.Enabled = False
        cboBaseRate.Enabled = False
        txtIntRate.Enabled = False
        dgBaseRates.Enabled = False
        cmdOK.Enabled = False
    Else
        dgBaseRates.Enabled = True
    End If
    cmdAnother.Enabled = False
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetGridFields
' Description   :   Sets the field names for the grid, specifies which
'                   are mandatory and which are banded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsBaseRateTable As BaseRateTable
    
    Set clsBaseRateTable = New BaseRateTable
    
    ' First, Valuation FeeSet. Not visible
    fields.sField = "BaseRateSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtBaseRates(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = "BaseRateBandStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtBaseRates(BANDED_DATE).Text
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' Next, TypeOfApplication has to be entered
    fields.sField = "MaximumLoanAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Loan Amount must be entered"
    fields.sTitle = "Maximum Loan Amount"

    colFields.Add fields
    
    fields.sField = "MaximumLTV"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum LTV must be entered"
    fields.sTitle = "Maximum LTV"

    colFields.Add fields
        
    ' SYS3504 New field on table
    fields.sField = "RateDifference"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Rate Difference must be entered"
    fields.sTitle = "Rate Difference"
    colFields.Add fields
    
    ' Versioning
    If g_clsVersion.DoesVersioningExist() Then
        fields.sField = clsBaseRateTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        colFields.Add fields
    End If
    
    Me.dgBaseRates.SetColumns colFields, "EditBaseRates", "Base Rates"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets the date and set number
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim clsBaseRateSet As BaseRateTable
    
    Set clsBaseRateSet = m_clsTableAccess
    txtBaseRates(BANDED_FEE_SET).Text = clsBaseRateSet.GetFeeSet()
    txtBaseRates(BANDED_DATE).Text = clsBaseRateSet.GetStartDate()
    
    'SYS3504 Populate the combo
    PopulateBaseRates
        
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
    
    If bRet = True Then
        bRet = dgBaseRates.ValidateRows()
    End If
    
    ' If we're adding, check that this record doesn't exist already
    If bRet = True And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Loan Amount and Maximum LTV must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetDataGridState
' Description   :   Enables or disables the datagrid based on other fields
'                   that need data to be entered.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgBaseRates.Enabled
        dgBaseRates.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtBaseRates(BANDED_FEE_SET).Text
                m_sOriginalStartDate = txtBaseRates(BANDED_DATE).Text
                    
                dgBaseRates.AddRow
                dgBaseRates.SetFocus
            End If
        End If
    End If
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


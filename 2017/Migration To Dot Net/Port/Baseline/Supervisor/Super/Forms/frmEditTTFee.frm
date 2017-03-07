VERSION 5.00
Object = "{1EB14087-1826-44AD-A7E5-2292F4882971}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditTTFee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[Add/Edit] TT Fee Set"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmEditTTFee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAction 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4395
      TabIndex        =   3
      Top             =   5145
      Width           =   1215
   End
   Begin VB.CommandButton CmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5670
      TabIndex        =   4
      Top             =   5145
      Width           =   1215
   End
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Fee Set Bands"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   180
      TabIndex        =   6
      Top             =   570
      Width           =   7965
      Begin MSGOCX.MSGDataGrid dgTTFee 
         Height          =   4065
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   7170
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
   Begin VB.CommandButton CmdAction 
      Caption         =   "A&nother"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6930
      TabIndex        =   5
      Top             =   5145
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtSetNumber 
      Height          =   330
      Left            =   1275
      TabIndex        =   0
      Top             =   180
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   582
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
   Begin MSGOCX.MSGEditBox txtStartDate 
      Height          =   330
      Left            =   6885
      TabIndex        =   1
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
   Begin VB.Label lblForm 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   5760
      TabIndex        =   8
      Top             =   225
      Width           =   810
   End
   Begin VB.Label lblForm 
      AutoSize        =   -1  'True
      Caption         =   "Set Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   225
      Width           =   990
   End
End
Attribute VB_Name = "frmEditTTFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
' Form          :   frmEditProductSwitchFee
' Description   :   Allow to Add/Edit Product Switch Fee
' Change history
' Initial      Prog      Date        Description
'
'   MC      BMIDS763    01/06/2004  This form allows to edit/add product switch fee.
'
'**************************************************************************************
Option Explicit

' Private data
Private m_sOriginalSet          As String
Private m_sOriginalStartDate    As String
Private m_bIsEdit               As Boolean
Private m_clsTableAccess        As TableAccess
Private m_ReturnCode            As MSGReturnCode
Private m_colKeys               As Collection

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


Private Sub CmdAction_Click(Index As Integer)
    Dim bRet As Boolean
    On Error GoTo ErrorHandler
    Select Case Index
        Case 0
            bRet = DoOKProcessing()
            If bRet = True Then
                Hide
            End If
        Case 1
            'Unload Me
            Hide
        Case 2
            bRet = DoOKProcessing()
            If bRet = True Then
                'If the record was saved, commit the transaction and begin another.
                g_clsDataAccess.CommitTrans
                g_clsDataAccess.BeginTrans
                g_clsFormProcessing.ClearScreenFields Me
                Form_Load
            End If
    End Select
ExitHandler:
    '*=[MC]Cleanup code here
    
    Exit Sub
    
ErrorHandler:
    '*=Log Error and Display
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Initialize()
    Set m_colKeys = New Collection
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    '*=[MC]Initialize form
    Call SetReturnCode(MSGFailure)
    
    Set m_clsTableAccess = New TTFeeBand
    dgTTFee.Enabled = False
    If m_bIsEdit = True Then
        Call SetEditState
    Else
        Call SetAddState
    End If
    '*=Persist SetNumber and StartDate
    m_sOriginalSet = txtSetNumber.Text
    m_sOriginalStartDate = txtStartDate.Text

ExitHandler:
    '*=Cleanup
    
    Exit Sub
ErrorHandler:
    'Log and show error
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub


Private Function DoesRecordExist() As Boolean
    
    Dim bRet        As Boolean
    Dim sSet        As String
    Dim sStartDate  As String
    Dim col         As Collection
    
    On Error GoTo ErrorHandler
    
    Set col = New Collection
    '*=Get SetNumber and Start Dt from input fields
    sSet = txtSetNumber.Text
    sStartDate = txtStartDate.Text
    
    '*=If Set Number and Startdate are blank throw an error
    If Len(sSet) > 0 And Len(sStartDate) > 0 Then
        col.Add sSet
        col.Add sStartDate
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        '*=[MC]Check the data, If record already exist
        bRet = m_clsTableAccess.DoesRecordExist(col)
        If bRet = True Then
            MsgBox "A Fee Set with this Set Number and Start Date already exist!", vbCritical
            '*=Set focus back to Set number field
            txtSetNumber.SetFocus
        End If
    Else
        '*=Raise an error.// Set Number or Start Date can not be blank
        g_clsErrorHandling.RaiseError errGeneralError, "Set Number and Start date can not be blank or Invalid", vbCritical
    End If
    
    DoesRecordExist = bRet

ExitHandler:
    '*=[MC]Cleanup Code here
    Set col = Nothing
    Exit Function
ErrorHandler:
    'Log  and throw Errors
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Private Sub SaveChangeRequest()
    
    Dim colMatchValues  As Collection
    Dim sSet            As String
    Dim sDate           As String
    
    On Error GoTo ErrorHandler
    
    Set colMatchValues = New Collection
    '*=Add values to the colleciton
    sSet = txtSetNumber.Text
    sDate = txtStartDate.Text
    colMatchValues.Add sSet
    colMatchValues.Add sDate
    '*=Set values collection to the TableAccess Database Object
    Call m_clsTableAccess.SetKeyMatchValues(colMatchValues)
    '*=Process SaveChanges Request
    Call g_clsHandleUpdates.SaveChangeRequest(m_clsTableAccess)

ExitHandler:
    '*=[MC]Cleanup code here
    Set colMatchValues = Nothing
    Exit Sub
ErrorHandler:
    'Log and Throw error
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoOKProcessing() As Boolean
    
    Dim bRet As Boolean
    On Error GoTo ErrorHandler
    
    '*=Validate Screen Data
    bRet = ValidateScreenData()
    '*=If valid continue process
    If bRet = True Then
        '*=Update ProductSwitchFeeBand records, If they are any
        DoUpdates
        '*[MC]Update ProductSwitchFeeSet
        DoSetUpdates
        '*=Update the ProductSwitchFeeSet Table
        m_clsTableAccess.Update
        '*=Save changes
        Call SaveChangeRequest
        Call SetReturnCode(MSGSuccess)
    End If
    
    DoOKProcessing = bRet

ExitHandler:
    Exit Function
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    Exit Function
    Resume
End Function

Private Sub DoUpdates()

    Dim clsBandedTable  As BandedTable
    Dim colValues       As Collection
    Dim sSetID          As String
    Dim sStartDate      As String
    Dim bRet            As Boolean

    On Error GoTo ErrorHandler
    
    Set colValues = New Collection
    '*=Use the banded interface
    Set clsBandedTable = m_clsTableAccess
    
    bRet = False
    
    sStartDate = txtStartDate.Text
    sSetID = txtSetNumber.Text
    
    If Len(sStartDate) > 0 And Len(sSetID) > 0 Then
        If sStartDate <> m_sOriginalStartDate Or sSetID <> m_sOriginalSet Then
            Call colValues.Add(sSetID)
            Call colValues.Add(sStartDate)
        
            Call clsBandedTable.SetUpdateValues(colValues)
            Call clsBandedTable.SetUpdateSets
            '*=Do the set updates with the new values
            Call clsBandedTable.DoUpdateSets
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "AdminFees - StartDate and SetID must be valid"
    End If
    
ExitHandler:
    '*=Cleanup here
        Set colValues = Nothing
        Set clsBandedTable = Nothing
    Exit Sub
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub DoSetUpdates()
    
    Dim bRet                As Boolean
    Dim sSet                As String
    Dim rs                  As Recordset
    Dim clsTableAccess      As TableAccess
    Dim clsProductFeeSet    As TTFeeSetTable
    Dim col                 As Collection
    
    On Error GoTo ErrorHandler

    Set col = New Collection
    
    sSet = txtSetNumber.Text
    If Len(sSet) > 0 Then
        Set clsProductFeeSet = New TTFeeSetTable
        Set clsTableAccess = clsProductFeeSet
        'Add field values to the collection
        Call col.Add(sSet)
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        '*=Set Value collection to the tableAccess object
        clsTableAccess.SetKeyMatchValues col
        '*=Does RecordExist?
        bRet = clsTableAccess.DoesRecordExist(col)
        '*=If Doesnot exist, Then Add
        If bRet = False Then
            '*=Add New Record
            g_clsFormProcessing.CreateNewRecord clsTableAccess
            '*=Update the changes
            Call clsProductFeeSet.SetFeeSet(sSet)
            Call clsTableAccess.Update
        End If
    End If
ExitHandler:
    '*=[MC]Cleanup
    Set col = Nothing
    Set clsTableAccess = Nothing
    Set clsProductFeeSet = Nothing
    Set rs = Nothing
    
    Exit Sub
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub dgTTFee_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    
    If nCurrentRow >= 0 Then
        bCancel = Not dgTTFee.ValidateRow(nCurrentRow)
    End If
End Sub

Private Sub SetupDataGrid()
    
    Dim colTables       As Collection
    Dim colDataControls As Collection
    
    On Error GoTo ErrorHandler
    
    Set colTables = New Collection
    Set colDataControls = New Collection
    '*=Add Table to collection
    Call colTables.Add(m_clsTableAccess)
    Call colDataControls.Add(dgTTFee)
    
    '*=Populate date if edit mode
    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If

    If m_bIsEdit = True Then
        If m_clsTableAccess.RecordCount > 0 Then
            Call PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate current record"
        End If
    End If

    Call SetGridFields
ExitHandler:
    '*=[MC]Cleanup Code here
    Set colTables = Nothing
    Set colDataControls = Nothing
    
    Exit Sub
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAddState()
    '*=Setup ProductSwitchFee DataGrid
    Call SetupDataGrid
    
    If txtSetNumber.Visible = True Then
        txtSetNumber.SetFocus
    End If
    Me.Caption = "Add TT Fee Set"
End Sub

Public Sub SetEditState()
    
    On Error GoTo ErrorHandler
    
    Call m_clsTableAccess.SetKeyMatchValues(m_colKeys)
    '*=Setup DataGrid
    Call SetupDataGrid
    '*=Set control state
    txtSetNumber.Enabled = False
    txtStartDate.Enabled = False
    dgTTFee.Enabled = True
    CmdAction(2).Enabled = False
    Me.Caption = "Edit TT Fee Set"
ExitHandler:
    Exit Sub
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetGridFields()
    
    Dim fields                  As FieldData
    Dim colFields               As Collection
    Dim colComboValues          As Collection
    Dim colComboIDS             As Collection
    Dim clsCombo                As ComboUtils
    Dim clsAdminFeeTable        As TTFeeBand
    Dim sVersionField           As String
    
    On Error GoTo ErrorHandler
    
    Set colFields = New Collection
    Set colComboValues = New Collection
    Set colComboIDS = New Collection
    Set clsCombo = New ComboUtils
    
    Set clsAdminFeeTable = m_clsTableAccess
    
    '*=[MC]PRODUCTSWITCHFEESET (First, Fee Set. Not visible, but needs the Fee Set copied in)
    fields.sField = "TTFeeSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtSetNumber.Text
    fields.sError = ""
    fields.sTitle = "Fee Set"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    Call colFields.Add(fields)
    
    '*=StartDate
    fields.sField = clsAdminFeeTable.GetStartDateField()
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtStartDate.Text
    fields.sError = ""
    fields.sTitle = "Start Date"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    Call colFields.Add(fields)
    
    ' TypeOfApplicationText
    fields.sField = "TypeOfApplicationText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Type of Application must be entered"
    fields.sTitle = "Type of Application"
    fields.sOtherField = "TypeOfApplication"
    
    clsCombo.FindComboGroup "TypeOfMortgage", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    Call colFields.Add(fields)
        
    ' Next, TypeOfApplication has to be entered
    fields.sField = "TypeOfApplication"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    Call colFields.Add(fields)
    
    ' Location
    fields.sField = "Location"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    Call colFields.Add(fields)
    
    ' Location Text - a combo
    fields.sField = "LocationText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Location must be entered"
    fields.sTitle = "Location"
    fields.sOtherField = "Location"
    
    clsCombo.FindComboGroup "PropertyLocation", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    Call colFields.Add(fields)
    
    ' Amount
    fields.sField = "Amount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Amount must be entered"
    fields.sTitle = "Amount"

    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    Call colFields.Add(fields)
    
    If g_clsVersion.DoesVersioningExist() Then
        ' Versioning...
        fields.sField = clsAdminFeeTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        Call colFields.Add(fields)
    End If
    '*=[mc]SET DATAGRID COLUMNS
    Call dgTTFee.SetColumns(colFields, "EditProductSwitchFees", "Product Switch Fees")
ExitHandler:
    '*=Cleanup here
    Set colFields = Nothing
    Set colComboValues = Nothing
    Set colComboIDS = Nothing
    Set clsCombo = Nothing
    Set clsAdminFeeTable = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PopulateScreenFields()
    
    On Error GoTo ErrorHandler
    
    Dim clsAdminFees As TTFeeBand
    Set clsAdminFees = m_clsTableAccess
    
    txtSetNumber.Text = clsAdminFees.GetFeeSet()
    txtStartDate.Text = clsAdminFees.GetStartDate()

ExitHandler:
    Exit Sub
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidateScreenData() As Boolean
    
    Dim bRet        As Boolean
    Dim bShowError  As Boolean
    
    On Error GoTo ErrorHandler
    
    bShowError = True
    '*=Validate
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    '*=In Add operation, Validate first if records exists with same data
    If bRet And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet Then
        bRet = dgTTFee.ValidateRows()
    End If

    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Type of Application and Location must be unique"
        End If
    End If
    
    'If bRet Then
    '    bRet = ValidateSpecialRules()
    'End If
    
ExitHandler:
    ValidateScreenData = bRet
    Exit Function

ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'*******************************************************************************
'   FUNCTION    : ValidateSpecialRules()
'   Date        : 07/06/2004
'   Author      : Mallesh Cheekoti
'   Desc        : This function validates client specific rules
'
'*******************************************************************************
Private Function ValidateSpecialRules()
    
    Dim bRet        As Boolean
    Dim sSetNumber  As String
    Dim sStartDate  As String
    Dim sLocation   As String
    Dim sTypeofApp  As String
    Dim colFields   As Collection
    Dim colValues   As Collection
    Dim lLoop       As Long
    Dim lTotalCount As Long
    
    On Error GoTo ErrorHandler
    bRet = True
    Set colFields = New Collection
    Set colValues = New Collection
    'Add columns to the collection
    colFields.Add "TTFeeSet"
    colFields.Add "STARTDATE"
    colFields.Add "TypeOfApplication"
    colFields.Add "Location"
    
    '*=Get StartDt, TypeofApp and Location from screen fields
    sSetNumber = txtSetNumber.Text
    sStartDate = txtStartDate.Text
    lTotalCount = dgTTFee.Rows
    '*=[MC]Iterate through all records in the datagrid and validate the data
    For lLoop = 0 To lTotalCount - 1
        Set colValues = New Collection
        sTypeofApp = dgTTFee.GetAtRowCol(CInt(lLoop), 3)
        sLocation = dgTTFee.GetAtRowCol(CInt(lLoop), 4)
        'Add field Values
        colValues.Add sSetNumber
        colValues.Add sStartDate
        colValues.Add sTypeofApp
        colValues.Add sLocation
        '*=If record exist return false
        If m_clsTableAccess.DoesRecordExist(colValues, colFields) Then
            bRet = False
            Exit For
        End If
        Set colValues = Nothing
    Next lLoop
    '*=If records exist and returns false show message
    If lTotalCount > 0 And Not bRet Then
        MsgBox "Fee Set Already Exist. Please enter unique StartDate, Type Of Application and Location", vbInformation + vbOKOnly, "Supervisor"
    End If
    
ExitHandler:
    
    ValidateSpecialRules = bRet
    'Cleanup
    Set colValues = Nothing
    Set colFields = Nothing
    
    Exit Function
ErrorHandler:
    bRet = False
    Resume ExitHandler
    
End Function


Private Sub txtSetNumber_Validate(Cancel As Boolean)
    
    Cancel = Not txtSetNumber.ValidateData()
    
    If Cancel = False Then
        Call SetDataGridState
    End If
    
End Sub

Private Sub txtStartDate_Validate(Cancel As Boolean)
    
    Cancel = Not txtStartDate.ValidateData()
    
    If Cancel = False Then
        SetDataGridState
    End If
    
End Sub


Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim bEnabled As Boolean
    
    On Error GoTo ErrorHandler
    
    bShowError = False
    
    '*=Validate Form before commit changes
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        
        bEnabled = dgTTFee.Enabled
        dgTTFee.Enabled = True
        
        Call SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtSetNumber.Text
                m_sOriginalStartDate = txtStartDate.Text
                'Add row and set focus
                dgTTFee.AddRow
                dgTTFee.SetFocus
            End If
        End If
    End If
    
ExitHandler:
    'Cleanup
    
    Exit Sub
    
ErrorHandler:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function




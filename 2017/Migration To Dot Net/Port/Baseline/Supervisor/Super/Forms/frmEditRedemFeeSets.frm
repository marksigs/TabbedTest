VERSION 5.00
Object = "{5F540CC8-EA22-4F95-9EFE-BDB4E09F976D}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditRedemFeeSets 
   Caption         =   "Add / Edit ERC Fees"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Set Bands"
      Height          =   4215
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   7755
      Begin MSGOCX.MSGDataGrid dgRedemFees 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6870
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5580
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtRedemFees 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
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
   Begin MSGOCX.MSGEditBox txtRedemFees 
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   780
      Width           =   6795
      _ExtentX        =   9869
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
      MaxLength       =   50
   End
   Begin VB.Label lblBaseRates 
      AutoSize        =   -1  'True
      Caption         =   "Set Code"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   660
   End
   Begin VB.Label lblBaseRates 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   825
      Width           =   795
   End
End
Attribute VB_Name = "frmEditRedemFeeSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditRedemFeeSets
' Description   : Form that controls editing of a redemption fee set - that is, a Set of Base Rates as
'                 defined by their rate number
'
' Change history
' Prog      Date        Description
' AW       16/05/02    BM017 - Created
' AW       27/05/02    Added SaveChangeRequest()
' SR       20/07/04    BMIDS801
' HMA      03/11/04    BMIDS940  Changed headings.
' JD        06/04/2005 BMIDS982 Changed 'Redemption Fee' text to 'ERC'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Text indexes
Private Const BANDED_FEE_SET = 0
Private Const BANDED_FEE_DESC = 1

' Private data
Private m_sOriginalSet As String
Private m_sOriginalDesc As String
Private m_bIsEdit As Boolean

Private m_clsRedemSetTable As RedemptionFeeSetTable
Private m_clsTableAccess As TableAccess

Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection

'Private Const PERIOD_COL As Integer = 3
'Private Const PERIOD_END_DATE_COL As Integer = 4
'Private Const FEE_PERCENT_COL As Integer = 5
'Private Const FEE_MONTHS_COL As Integer = 6

Private Const PERIOD_COL As Integer = 2
Private Const PERIOD_END_DATE_COL As Integer = 3
Private Const FEE_PERCENT_COL As Integer = 4
Private Const FEE_MONTHS_COL As Integer = 5

Public Enum SearchDirection
    SEARCH_UP = 1
    SEARCH_DOWN = 2
    SEARCH_NONE = 0
    
End Enum

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

Private Sub dgRedemFees_AddButtonClicked(bCancel As Boolean, nCurrentRow As Integer)
    
    'dgRedemFees.Rows(dgRedemFees.Row).Text = dgRedemFees.Rows(dgRedemFees.Row - 1).Text
    Call dgRedemFees.SetAtRowCol(nCurrentRow, 1, dgRedemFees.GetAtRowCol(nCurrentRow - 1, 1) + 1)
    
    '*=Set Added row as current row
    dgRedemFees.Row = nCurrentRow
    
    'SR 20/07/2004 : BMIDS801
    If dgRedemFees.Rows > 1 Then
        dgRedemFees.AllowDelete = True
    Else
        dgRedemFees.AllowDelete = False
    End If
    'SR 20/07/2004 : BMIDS801 - End
    
End Sub

Private Function GetDataIFExists(ByVal v_eDirection As SearchDirection, ByVal currentRow As Long, ByVal columnIndex As Long) As String
    
    Dim lLoop As Long
    Dim sReturn As String
    'Get the available Redemptionfee end date from the given direction
    Select Case v_eDirection
        Case SEARCH_DOWN
            For lLoop = currentRow + 1 To dgRedemFees.Rows - 1
                sReturn = CStr(dgRedemFees.GetAtRowCol(CLng(lLoop), CLng(columnIndex)))
                If sReturn <> "" Then
                    Exit For
                End If
            Next lLoop
                
        Case SEARCH_UP
            For lLoop = currentRow - 1 To 0 Step -1
                sReturn = CStr(dgRedemFees.GetAtRowCol(CLng(lLoop), CLng(columnIndex)))
                If sReturn <> "" Then
                    Exit For
                End If
            Next lLoop
        Case SEARCH_NONE
            sReturn = ""
    End Select
    
    dgRedemFees.Row = currentRow
    
    GetDataIFExists = sReturn

End Function
'SR 20/07/2004 : BMIDS801
Private Sub dgRedemFees_BeforeDelete(Cancel As Integer)
    If dgRedemFees.Rows = 1 Then
        MsgBox "Last Row cannot be deleted"
        Cancel = True
    Else
        If dgRedemFees.Rows > 2 Then
            dgRedemFees.AllowDelete = True
        Else
            dgRedemFees.AllowDelete = False
        End If
    End If
End Sub
'SR 20/07/2004 : BMIDS801 - End

Private Sub dgRedemFees_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
'''    Dim sDtUP As String
'''    Dim sDtDown As String
'''    Dim sCurrentRowValue As String
'''
'''
'''    Select Case LastCol
'''        Case 3  'Period End Date
'''            If LastRow < dgRedemFees.Rows Then
'''
'''
'''            'dgRedemFees.Row = dgRedemFees.Rows - 1
'''            sCurrentRowValue = dgRedemFees.GetAtRowCol(CLng(LastRow), CLng(LastCol))
'''            If sCurrentRowValue <> "" Then
'''            'MsgBox sCurrentRowValue
'''            'dgRedemFees.Row = dgRedemFees.Row
'''            sDtUP = GetDataIFExists(SEARCH_UP, CLng(LastRow), LastCol)
'''
'''            sDtDown = GetDataIFExists(SEARCH_DOWN, CLng(LastRow), LastCol)
'''
'''            'dgRedemFees.Row = LastRow
'''            Call dgRedemFees.SetAtRowCol(CLng(LastRow), CLng(LastCol), CVar(sCurrentRowValue))
'''            If Len(sDtUP) > 0 And Len(sDtDown) > 0 Then
'''                If Not ((CDate(sCurrentRowValue) > CDate(sDtUP)) And (CDate(sCurrentRowValue) < CDate(sDtDown))) Then
'''                    MsgBox "Date must be with in " & sDtUP & " ---  " & sDtDown
'''                    dgRedemFees.col = LastCol
'''                    dgRedemFees.Row = LastRow
'''                    dgRedemFees.SetGridFocus
'''
'''                End If
'''            Else
'''                If Len(sDtUP) > 0 Then
'''                    If Not (CDate(sCurrentRowValue) > CDate(sDtUP)) Then
'''                        MsgBox "Date must be later than " & sDtUP
'''                        dgRedemFees.col = LastCol
'''                        dgRedemFees.Row = LastRow
'''                        dgRedemFees.SetGridFocus
'''                    End If
'''                ElseIf Len(sDtDown) > 0 Then
'''                    If Not CDate(sCurrentRowValue) < CDate(sDtDown) Then
'''                        MsgBox "date must be less than " & sDtDown
'''                        dgRedemFees.col = LastCol
'''                        dgRedemFees.Row = LastRow
'''                        dgRedemFees.SetGridFocus
'''                    End If
'''                End If
'''            End If
'''
'''        End If
'''        End If
'''    End Select
'''
'''
    
End Sub

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
    Set m_clsTableAccess = New RedemptionFeeTable
    Set m_clsRedemSetTable = New RedemptionFeeSetTable
    dgRedemFees.Enabled = False
    
    If m_bIsEdit = True Then
        SetEditState
        'SR 20/07/2004 : BMIDS801
        If dgRedemFees.Rows > 1 Then
            dgRedemFees.AllowDelete = True
        Else
            dgRedemFees.AllowDelete = False
        End If
        'SR 20/07/2004 : BMIDS801 - End
    Else
        SetAddState
        dgRedemFees.AllowDelete = False  'SR 20/07/2004 : BMIDS801 - End
    End If
    
    m_sOriginalSet = txtRedemFees(BANDED_FEE_SET).Text
    m_sOriginalDesc = txtRedemFees(BANDED_FEE_DESC).Text
    '*=Lock the Set No. Column, It seems Start Date column is hiding
    dgRedemFees.Columns.Item(0).Locked = True 'start Date
    dgRedemFees.Columns.Item(1).Locked = True ' Set No.
    
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
    Dim sDesc As String
    
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    
    sSet = txtRedemFees(BANDED_FEE_SET).Text
    sDesc = txtRedemFees(BANDED_FEE_DESC).Text
    
    If Len(sSet) > 0 And Len(sDesc) > 0 Then
    
        Set clsTableAccess = m_clsRedemSetTable
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        bRet = clsTableAccess.DoesRecordExist(col)
        
        If bRet = True Then
            MsgBox "Set already exist - please enter a unique combination", vbCritical
            txtRedemFees(BANDED_FEE_SET).SetFocus
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
    Dim sDesc As String
    
    Set colMatchValues = New Collection
    sSet = Me.txtRedemFees(BANDED_FEE_SET).Text
    'sDesc = Me.txtAdminFees(BANDED_DESC).Text
    colMatchValues.Add sSet
    'colMatchValues.Add sDesc
    
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
    
    Dim col As New Collection
    
    sSet = txtRedemFees(BANDED_FEE_SET).Text
    sDesc = txtRedemFees(BANDED_FEE_DESC).Text

    If Len(sSet) > 0 Then

        Set clsTableAccess = m_clsRedemSetTable
        
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        
        bRet = clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess
        
            m_clsRedemSetTable.SetTheRedemptionFeeSet sSet
            m_clsRedemSetTable.SetDescription sDesc
            clsTableAccess.Update
        Else
            m_clsRedemSetTable.SetTheRedemptionFeeSet sSet
            m_clsRedemSetTable.SetDescription sDesc
            
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   dgRedemFees_BeforeAdd
' Description   :   Called before the Add button is pressed on the datagrid
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgRedemFees_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgRedemFees.ValidateRow(nCurrentRow)
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
    colDataControls.Add dgRedemFees
    Call m_clsTableAccess.SetOrderCriteria("RedemptionFeeStepNumber")
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
    
    txtRedemFees(BANDED_FEE_DESC).Enabled = True
    txtRedemFees(BANDED_FEE_SET).Enabled = True
    
    SetupDataGrid
    
    'txtRedemFees(BANDED_FEE_SET).SetFocus
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing an Admin Fee Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    On Error GoTo Failed
    
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDataGrid
    
    txtRedemFees(BANDED_FEE_SET).Enabled = False
    txtRedemFees(BANDED_FEE_DESC).Enabled = True
    
    dgRedemFees.Enabled = True
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
    Dim clsRedemFeeTable As RedemptionFeeTable
    
    Set clsRedemFeeTable = m_clsTableAccess
    
    fields.sField = "RedemptionFeeSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtRedemFees(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = ""
    fields.bDateField = False
    colFields.Add fields
    
    fields.sField = "RedemptionFeeStepNumber"
    fields.bRequired = True
    fields.bVisible = True
    fields.sError = "Step Number must be entered"
    fields.sDefault = ""
    fields.sTitle = "Step No."
    fields.bDateField = False
    colFields.Add fields
    
'    fields.sField = "RedemptionFeeStartDate"
'    fields.bRequired = True
'    fields.bVisible = True
'    fields.sDefault = ""
'    fields.sError = "Start Date must be entered"
'    fields.sTitle = "Start Date"
'    fields.bDateField = True
'
'    colFields.Add fields
    
    fields.sField = "Period"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Period"
    fields.bDateField = False
    
    colFields.Add fields
    
    fields.sField = "PeriodEndDate"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Period End Date must be entered"
    fields.sTitle = "Period End Date"
    fields.bDateField = True
    
    colFields.Add fields
    
    fields.sField = "FeePercentage"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Fee Percentage"
    fields.bDateField = False
    colFields.Add fields
    
    fields.sField = "FeeMonthsInterest"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sTitle = "Fee Months Interest"
    fields.bDateField = False

    colFields.Add fields
    
    fields.sField = "LastChangeDate"
    fields.bRequired = False
    fields.bVisible = False
    fields.sDefault = ""
    fields.bDateField = True
    colFields.Add fields

  
    If g_clsVersion.DoesVersioningExist() Then
        ' Versioning...
        fields.sField = clsRedemFeeTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        
        colFields.Add fields
    End If
    
    dgRedemFees.SetColumns colFields, "EditRedemptionFees", "ERC"  'JD BMIDS982
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
    
    Dim clsRedemSetTable As RedemptionFeeSetTable
    Set clsRedemSetTable = m_clsRedemSetTable
    
    Dim clsRedemBandTable As RedemptionFeeTable
    Set clsRedemBandTable = m_clsTableAccess

    sKey = clsRedemBandTable.GetRedemFeeSet()
    
    Set rs = clsRedemSetTable.GetRedemptionSetDesc(sKey)

    If rs.RecordCount = 1 Then
        txtRedemFees(BANDED_FEE_SET).Text = rs!REDEMPTIONFEESET
        txtRedemFees(BANDED_FEE_DESC).Text = rs!REDEMPTIONFEESETDESCRIPTION
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
        bRet = dgRedemFees.ValidateRows()
    End If

    If bRet Then
        bRet = ValidateGridFields
    End If
    
    If bRet Then
        bRet = ValidateCustomRules()
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Start Date and Step No. must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    Exit Function
    Resume
End Function

Private Function ValidateCustomRules()
    
    Dim lLoop As Long
    Dim lTotalRows As Long
    Dim sDtUP As String
    Dim sDtDown As String
    Dim bResult As Boolean
    Dim sCurrentRowValue As String
    
    bResult = True
    
    lTotalRows = dgRedemFees.Rows
    
    'Validate the dates in order
    For lLoop = lTotalRows - 1 To 0 Step -1
        sCurrentRowValue = dgRedemFees.GetAtRowCol(CInt(lLoop), CInt(PERIOD_END_DATE_COL))
        If sCurrentRowValue <> "" Then
            sDtUP = GetDataIFExists(SEARCH_UP, lLoop, PERIOD_END_DATE_COL)
            'sDtDown = GetDataIFExists(SEARCH_DOWN, lLoop, PERIOD_END_DATE_COL)  ' SR 20/07/2004 : BMIDS800
        ' SR 20/07/2004 : BMIDS800
            If Len(sDtUP) > 0 Then
                If Not (CDate(sCurrentRowValue) > CDate(sDtUP)) Then
                    MsgBox "[Period End Date] must be later than " & sDtUP, vbInformation + vbOKOnly, "Omiga4 Supervisor"
                    dgRedemFees.col = PERIOD_END_DATE_COL
                    dgRedemFees.Row = lLoop
                    dgRedemFees.SetGridFocus
                    bResult = False
                    Exit For
                End If
            End If
        End If
        ' SR 20/07/2004 : BMIDS800 - End
    Next lLoop
    
    ValidateCustomRules = bResult
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   txtRedemFees_Validate
' Description   :   Called when the text boxes lose focus. Checks for mandatory
'                   requirments and for the entry to be valid
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRedemFees_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtRedemFees(Index).ValidateData()

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
        
        bEnabled = dgRedemFees.Enabled
        dgRedemFees.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtRedemFees(BANDED_FEE_SET).Text
                m_sOriginalDesc = txtRedemFees(BANDED_FEE_DESC).Text

                dgRedemFees.AddRow
                '*=MC SetNo. Auto Increment, First Row must have a number
                If dgRedemFees.Rows = 1 Then
                    dgRedemFees.Columns(1).Text = 1
                End If
                
                dgRedemFees.SetFocus
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
'   AW      16/05/02    BM017
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateGridFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    Dim nRowIndex As Integer
    Dim strPeriod As String, strPeriodEndDate As String, strFeePercent As String
    Dim strFeeMonths As String

    ValidateGridFields = False
    
    'Iterate through all the rows of the grid control.
    For nRowIndex = 0 To (dgRedemFees.Rows - 1)

        strPeriod = dgRedemFees.GetAtRowCol(nRowIndex, PERIOD_COL)
        strPeriodEndDate = dgRedemFees.GetAtRowCol(nRowIndex, PERIOD_END_DATE_COL)
        strFeePercent = dgRedemFees.GetAtRowCol(nRowIndex, FEE_PERCENT_COL)
        strFeeMonths = dgRedemFees.GetAtRowCol(nRowIndex, FEE_MONTHS_COL)

        If (Len(strPeriod) = 0 And Len(strPeriodEndDate) = 0) Or _
            (Len(strPeriod) > 0 And Len(strPeriodEndDate) > 0) Then
                'Place the cursor in the invalid field.
                dgRedemFees.col = PERIOD_COL
                dgRedemFees.Row = nRowIndex
                dgRedemFees.SetGridFocus
    
                g_clsErrorHandling.RaiseError errGeneralError, "Either Period or Period End Date must be entered."
        End If

        If (Len(strFeePercent) = 0 And Len(strFeeMonths) = 0) Or _
            (Len(strFeePercent) > 0 And Len(strFeeMonths) > 0) Then
            
                'Place the cursor in the invalid field.
                dgRedemFees.col = FEE_PERCENT_COL
                dgRedemFees.Row = nRowIndex
                dgRedemFees.SetGridFocus
    
                g_clsErrorHandling.RaiseError errGeneralError, "Either a Fee Percentage or Fee Months Interest must be entered."
        End If
        
    Next nRowIndex
       
    ValidateGridFields = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


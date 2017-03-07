VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmLegalFeeSets 
   Caption         =   "Legal Fees"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "frmLegalFeeSets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboApplicationType 
      Height          =   315
      Left            =   6180
      TabIndex        =   2
      Top             =   900
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
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
      ListText        =   "New Loan|Remortgage|Further Advance|Transfer of Equity"
      Text            =   ""
   End
   Begin MSGOCX.MSGComboBox cboFeeType 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   900
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
      ListText        =   "Redeeming|Dormant"
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7380
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Fee Set Bands"
      Height          =   4215
      Left            =   300
      TabIndex        =   11
      Top             =   1320
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgLegalFees 
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
   Begin MSGOCX.MSGEditBox txtLegalFeeSet 
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
   Begin MSGOCX.MSGEditBox txtLegalFeeSet 
      Height          =   315
      Index           =   1
      Left            =   2100
      TabIndex        =   0
      Top             =   600
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
   Begin VB.Label lblLegalFeeSet 
      Caption         =   "Application Type"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblLegalFeeSet 
      Caption         =   "Fee Type"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblLegalFeeSet 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLegalFeeSet 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmLegalFeeSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditLegalFeeSets
' Description   : Form that controls editing of a Legal Fee set - that is, a Set of Legal Fees as
'                 defined by their Set number, date, app type and fee type
'
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Textedit keys
Private Const FEE_SET_COL = 0
Private Const START_DATE_COL = 1
Private Const FEE_TYPE_COL = 2
Private Const MAXIMUM_VALUE_COL = 3
Private Const TYPE_OF_APPLICATION_COL = 4
Private Const AMOUNT_COL = 5

' Private data
Private m_clsLegalFees As LegalFeesTable
Private m_colMatchValues As Collection
Private m_bIsEdit As Boolean
Private m_vFeeSet As Variant
Private m_ReturnCode As MSGReturnCode
Private m_OriginalValues As ScreenValues

' Types
Private Type ScreenValues
    sSet As String
    sStartDate As String
    sFeeType As String
    sAppType As String
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    Dim bRet As Boolean
    
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    Initialise
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Performs all initialisation for this form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Initialise()
    On Error GoTo Failed
    Dim bEdit As Boolean
    
    bEdit = True
    Set m_clsLegalFees = New LegalFeesTable
    
    g_clsFormProcessing.PopulateCombo "LegalFeeType", cboFeeType
    g_clsFormProcessing.PopulateCombo "TypeOfMortgage", cboApplicationType
    
    If Not m_colMatchValues Is Nothing Then
        TableAccess(m_clsLegalFees).SetKeyMatchValues m_colMatchValues
        TableAccess(m_clsLegalFees).GetTableData
    Else
        bEdit = False
        TableAccess(m_clsLegalFees).GetTableData POPULATE_EMPTY
    End If
    
    TableAccess(m_clsLegalFees).ValidateData
    
    SetupDataGrid
    
    If bEdit Then
        PopulateScreenFields
    End If
    
    txtLegalFeeSet(FEE_SET_COL).Text = m_vFeeSet
    
    dgLegalFees.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupDatagrid
' Description   :   Sets the DataSource of the datagrid (m_clsMIGRates will contain the records)
'                   and call SetGridFields to set the column headers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupDataGrid()
    On Error GoTo Failed
    
    Set dgLegalFees.DataSource = TableAccess(m_clsLegalFees).GetRecordSet
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetGridFields
' Description   :   Sets the column header names and behaviour of the datagrid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim vTmp As Variant
    
    ' First, OrganisationID. Not visible, but needs the lender code copied in
    fields.sField = "FeeSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtLegalFeeSet(FEE_SET_COL).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' StartDate has to be entered
    fields.sField = "LegalFeeStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtLegalFeeSet(START_DATE_COL).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Fee Type
    g_clsFormProcessing.HandleComboExtra cboFeeType, vTmp, GET_CONTROL_VALUE
    
    If IsNull(vTmp) Then
        vTmp = ""
    End If
    
    fields.sField = "FeeType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = CStr(vTmp)
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' Type of Application
    g_clsFormProcessing.HandleComboExtra Me.cboApplicationType, vTmp, GET_CONTROL_VALUE
    
    If IsNull(vTmp) Then
        vTmp = ""
    End If
    
    fields.sField = "TypeOfApplication"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = CStr(vTmp)
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' Maximum Value
    fields.sField = "MaximumValue"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Value must be entered"
    fields.sTitle = "Maximum Value"

    colFields.Add fields
    
    ' Amount
    fields.sField = "AMOUNT"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Amount must be entered"
    fields.sTitle = "Amount"
    
    colFields.Add fields

    dgLegalFees.SetColumns colFields, "EditLegalFees", "Legal Fees"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Returns true if a record exists for the keys in col.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoesRecordExist(col As Collection)
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = Not TableAccess(m_clsLegalFees).DoesRecordExist(col)

    If bRet = False Then
        g_clsFormProcessing.SetControlFocus txtLegalFeeSet(FEE_SET_COL)
        g_clsErrorHandling.RaiseError errGeneralError, "Record already exists - please enter a unique combination", vbCritical
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoUpdates
' Description   :   Updates the LegalFees table (as populated by the datagrid) with the keys
'                   the user has entered (and may have changed since they entered the screen,
'                   or they may be creating a new feeset)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim colValues As New Collection
    Dim currentValues As ScreenValues
    
    ' Get current values
    currentValues.sSet = txtLegalFeeSet(FEE_SET_COL).Text
    currentValues.sStartDate = txtLegalFeeSet(START_DATE_COL).Text

    g_clsFormProcessing.HandleComboExtra cboFeeType, currentValues.sFeeType, GET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboApplicationType, currentValues.sAppType, GET_CONTROL_VALUE

    colValues.Add currentValues.sSet
    colValues.Add currentValues.sStartDate
    colValues.Add currentValues.sFeeType
    colValues.Add currentValues.sAppType
    
    ' If any of these values have changed, make sure a record doesn't already exist
    If m_OriginalValues.sSet <> currentValues.sSet Or m_OriginalValues.sStartDate <> currentValues.sStartDate Or _
       m_OriginalValues.sFeeType <> currentValues.sFeeType Or m_OriginalValues.sAppType <> currentValues.sAppType Then
        
        DoesRecordExist colValues
    End If

    ' Do the updates
    BandedTable(m_clsLegalFees).SetUpdateValues colValues
    BandedTable(m_clsLegalFees).SetUpdateSets
    BandedTable(m_clsLegalFees).DoUpdateSets
    
    TableAccess(m_clsLegalFees).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user presses OK - validate all screen data then save it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = ValidateScreenData()
    If bRet = True Then
        DoUpdates
        SetReturnCode
        
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets all screen fields from data retrieved from the database, in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    
    ' Fee set
    m_vFeeSet = m_clsLegalFees.GetFeeSet
    txtLegalFeeSet(START_DATE_COL).Text = m_clsLegalFees.GetStartDate()
    
    g_clsFormProcessing.HandleComboExtra cboApplicationType, m_clsLegalFees.GetTypeOfApplication(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboFeeType, m_clsLegalFees.GetFeeType(), SET_CONTROL_VALUE
        
    m_OriginalValues.sSet = m_vFeeSet
    m_OriginalValues.sStartDate = txtLegalFeeSet(START_DATE_COL).Text
    
    g_clsFormProcessing.HandleComboExtra cboFeeType, m_OriginalValues.sFeeType, GET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboApplicationType, m_OriginalValues.sAppType, GET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateScreenData
' Description   :   Validates all mandatory data has been entered, and that at least one legal fee
'                   row has been created.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim nCount As Long
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    nCount = dgLegalFees.Rows()
            
    If nCount > 0 Then
        bRet = dgLegalFees.ValidateRows()
    Else
        MsgBox "At least one row must be entered", vbCritical
        dgLegalFees.SetGridFocus
        bRet = False
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetKeyMatchValues
' Description   :   Sets the keys to be used to identify a MIG Rate set, if known.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeyMatchValues(colMatchValues As Collection)
    On Error GoTo Failed
    
    If Not colMatchValues Is Nothing Then
        If colMatchValues.Count > 0 Then
            Set m_colMatchValues = colMatchValues
        Else
            g_clsErrorHandling.RaiseError errRateSetEmpty
        End If
    Else
        g_clsErrorHandling.RaiseError errRateSetEmpty
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetReturnCode
' Description   :   Returns if the operation on this form was successful, or not. Will be called
'                   externally to ascertain the status of this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    On Error GoTo Failed
    If m_ReturnCode = MSGSuccess Then
        If TableAccess(m_clsLegalFees).GetUpdated = False Then
            m_ReturnCode = MSGFailure
        End If
    End If
    
    GetReturnCode = m_ReturnCode
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set m_colMatchValues = Nothing
End Sub
Private Sub txtLegalFeeSet_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtLegalFeeSet(Index).ValidateData()
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Private Sub cboApplicationType_Validate(Cancel As Boolean)
    cboApplicationType.ValidateData
End Sub
Public Sub SetSetNumber(sSet As String)
    m_vFeeSet = sSet
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub dgLegalFees_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgLegalFees.ValidateRow(nCurrentRow)
    End If
End Sub


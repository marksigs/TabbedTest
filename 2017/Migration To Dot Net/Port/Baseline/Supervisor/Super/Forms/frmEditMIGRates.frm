VERSION 5.00
Object = "{5F540CC8-EA22-4F95-9EFE-BDB4E09F976D}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditMIGRates 
   Caption         =   "HLC Rate Set Bands"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "frmEditMIGRates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Set Bands"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   780
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgMIGRates 
         Height          =   3855
         Left            =   420
         TabIndex        =   1
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7110
      TabIndex        =   3
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   5460
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtMIGRates 
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   4
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
      Left            =   4860
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
   Begin VB.Label lblBaseRates 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   7
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblBaseRates 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditMIGRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditMIGRates
' Description   : Form that controls editing of a baserate set - that is, a Set of Base Rates as
'                 defined by their rate number and Start Date
'
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' JD	    29/03/2005	BMIDS982 Change all screen text MIG to HLC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Textedit constants
Private Const SET_NUMBER = 0
Private Const START_DATE = 1
Private Const RATE_SET_COL = 0
Private Const START_DATE_COL = 1

' Private data
Private m_vRateSet As Variant
Private m_sStartDate As String
Private m_clsMIGRates As MIGRateTable
Private m_colMatchValues As Collection
Private m_ReturnCode As MSGReturnCode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    SetReturnCode MSGFailure
    Set m_clsMIGRates = New MIGRateTable
    
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
    
    If Not m_colMatchValues Is Nothing Then
        TableAccess(m_clsMIGRates).SetKeyMatchValues m_colMatchValues
        TableAccess(m_clsMIGRates).GetTableData
    Else
        bEdit = False
        TableAccess(m_clsMIGRates).GetTableData POPULATE_EMPTY
    End If
    
    TableAccess(m_clsMIGRates).ValidateData
    
    SetupDataGrid
    
    If bEdit Then
        PopulateScreenFields
    End If
    
    dgMIGRates.Enabled = True
    txtMIGRates(SET_NUMBER).Text = CStr(m_vRateSet)
    
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
    
    Set dgMIGRates.DataSource = TableAccess(m_clsMIGRates).GetRecordSet
    
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
    
    ' First, RateSet. Not visible
    fields.sField = "RateSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtMIGRates(RATE_SET_COL).Text
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = "MIGRateStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtMIGRates(START_DATE_COL).Text
    fields.sError = ""
    fields.sTitle = ""

    colFields.Add fields
    
    ' Next, TypeOfApplication has to be entered
    fields.sField = "MaximumLTV"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum LTV must be entered"
    fields.sTitle = "Maximum LTV"

    colFields.Add fields
    
    fields.sField = "MaximumLoanAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Loan Amount must be entered"
    fields.sTitle = "Maximum Loan Amount"

    colFields.Add fields
    
    
    fields.sField = "IndemnityRate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Indemnity Rate must be entered"
    fields.sTitle = "Indemnity Rate"
    
    colFields.Add fields
    
    ' Versioning
    If g_clsVersion.DoesVersioningExist() Then
        fields.sField = "MIGRATESETVERSIONNUMBER"
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        colFields.Add fields
    End If

    dgMIGRates.SetColumns colFields, "EditMIGRates", "HLC Rates"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets any screen fields to the values retrieved from the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    
    ' Save the start date for furture reference
    m_sStartDate = m_clsMIGRates.GetStartDate()
    txtMIGRates(START_DATE).Text = m_sStartDate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateRows
' Description   :   Validates the rows on the datagrid. Returns TRUE if successful, FALSE if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateRows() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    
    bRet = True
    
    nCount = dgMIGRates.Rows()
            
    If nCount > 0 Then
        bRet = dgMIGRates.ValidateRows()
    Else
        MsgBox "At least one row must be entered", vbCritical
        dgMIGRates.SetGridFocus
        bRet = False
    End If
    
    ValidateRows = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Checks to see if a record exists already with the same
'                   keys. Returns true if it exists, false if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoesRecordExist(col As Collection)
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = TableAccess(m_clsMIGRates).DoesRecordExist(col)

    If bRet = True Then
        g_clsFormProcessing.SetControlFocus txtMIGRates(SET_NUMBER)
        g_clsErrorHandling.RaiseError errGeneralError, "Record already exists for this Start Date"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoUpdates
' Description   :   Updates the banded MIG Rates. The key is the start date
'                   and set number, so if they have changed we need to update all records with
'                   the new values.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim sSet As String
    Dim colValues As New Collection
    Dim sStartDate As String

    ' Get the current Set and Start Date
    sSet = txtMIGRates(SET_NUMBER).Text
    sStartDate = txtMIGRates(START_DATE).Text
    
    ' Add them to the colValues collection so it can be used in DoesRecordExist and for the updates,
    ' if required.
    colValues.Add sSet
    colValues.Add sStartDate
    
    ' If it's a new start date, check that one doesn't exist alreaedy
    If m_sStartDate <> sStartDate Then
        DoesRecordExist colValues
    End If
    
    ' Set the values to be used on the update
    BandedTable(m_clsMIGRates).SetUpdateValues colValues

    ' Do the Updates
    BandedTable(m_clsMIGRates).SetUpdateSets
    BandedTable(m_clsMIGRates).DoUpdateSets
    TableAccess(m_clsMIGRates).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user presses OK and performs all validation and updates.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bFailed As Boolean
    Dim bShowError As Boolean
    
    bFailed = False
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        bRet = ValidateRows()
        
        If bRet = True Then
            DoUpdates
            SetReturnCode
            Hide
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetKeyMatchValues
' Description   :   Sets the keys to be used to identify a MIG Rate set, if known.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeyMatchValues(colMatchValues As Collection)
    On Error GoTo Failed
    
    If Not colMatchValues Is Nothing Then
        If colMatchValues.Count > 0 Then
            Set m_colMatchValues = colMatchValues
        
            ' Save the rateset too
            m_vRateSet = m_colMatchValues(RATESET_KEY)
                    
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
' Function      :   SetRateSet
' Description   :   Sets the RateSet for a MIG Rate set. If a Rate Set doesn't exist already,
'                   this will be used to form part of the key.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetRateSet(vRateSet As Variant)
    On Error GoTo Failed
    
    m_vRateSet = vRateSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set m_colMatchValues = Nothing
End Sub
Private Sub txtMIGRates_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtMIGRates(Index).ValidateData()
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetReturnCode
' Description   :   Returns if the operation on this form was successful, or not. Will be called
'                   externally to ascertain the status of this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    On Error GoTo Failed
    If m_ReturnCode = MSGSuccess Then
        If TableAccess(m_clsMIGRates).GetUpdated = False Then
            m_ReturnCode = MSGFailure
        End If
    End If
    
    GetReturnCode = m_ReturnCode
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user presses Cancel. Hide the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub dgMIGRates_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgMIGRates.ValidateRow(nCurrentRow)
    End If
End Sub


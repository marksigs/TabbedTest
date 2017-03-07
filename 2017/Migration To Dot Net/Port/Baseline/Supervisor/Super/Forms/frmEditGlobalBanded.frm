VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditGlobalBanded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banded Global Add/Edit"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmEditGlobalBanded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGDataGrid dgBandedParameter 
      Height          =   2535
      Left            =   180
      TabIndex        =   3
      Top             =   1740
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4471
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
   Begin MSGOCX.MSGEditBox txtEditBanded 
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      TextType        =   4
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
      MaxLength       =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4530
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5850
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtEditBanded 
      Height          =   315
      Index           =   2
      Left            =   1260
      TabIndex        =   1
      Top             =   540
      Width           =   1095
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
   Begin MSGOCX.MSGTextMulti txtDescription 
      Height          =   735
      Left            =   1260
      TabIndex        =   2
      Top             =   900
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      MaxLength       =   255
   End
   Begin VB.Label lblEditBanded 
      Caption         =   "Description"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblEditBanded 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblEditBanded 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmEditGlobalBanded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BANDED_NAME = 0
Private Const BANDED_DATE = 2

Private Const NAME_POS = 0
Private Const START_DATE_POS = 1
Private Const MAX_VALUE_POS = 2
Private Const DESCRIPTION_POS = 3

Private Const AMOUNT_POS = 4
Private Const PERCENTAGE_POS = 5
Private Const MAXIMUM_POS = 6
Private Const BOOLEAN_POS = 7
Private Const STRING_POS = 8


Private m_sOriginalName As String
Private m_sOriginalDescription As String
Private m_sOriginalStartDate As String
Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean

Private m_clsTableAccess As TableAccess
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
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
'Public Sub SetTableClass(clsTable As BandedGlobalParametersTable)
'    Set m_clsTableAccess = clsTable
'End Sub
Private Sub CheckRecordsExist()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sStartDate As String
    Dim sName As String
    Dim col As New Collection
    
    sStartDate = txtEditBanded(BANDED_DATE).Text
    sName = txtEditBanded(BANDED_NAME).Text
    
    col.Add sName
    col.Add sStartDate
    
    bRet = m_clsTableAccess.DoesRecordExist(col)
    
    If bRet = True Then
        g_clsFormProcessing.SetControlFocus txtEditBanded(BANDED_NAME)
        g_clsErrorHandling.RaiseError errGeneralError, "Name and Start Date already exist - please enter a unique combination"

    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sName As String
    Dim sStartDate As String
    Dim sDescription As String
    Dim clsBandedTable As BandedTable
    
    ' If the Name, Description, or Date has changed then we need to update ALL rows
    ' with the new Values.
    sName = txtEditBanded(BANDED_NAME).Text
    sStartDate = txtEditBanded(BANDED_DATE).Text
    sDescription = txtDescription.Text 'txtEditBanded(BANDED_DESCRIPTION).Text
    
    If m_sOriginalName <> sName Or m_sOriginalDescription <> sDescription Or _
        m_sOriginalStartDate <> sStartDate Then
        Dim colValues As New Collection
        ' Point at a BandedGlobalPArameters object to set the update values
        Set clsBandedTable = m_clsTableAccess
        colValues.Add sName
        colValues.Add sStartDate
        colValues.Add sDescription
        
        clsBandedTable.SetUpdateValues colValues
    
        ' We only want to check that the record exists if the name or start date has changed
        If m_sOriginalName <> sName Or m_sOriginalStartDate <> sStartDate Then
            CheckRecordsExist
        End If
        
        clsBandedTable.SetUpdateSets
        clsBandedTable.DoUpdateSets

    End If
    m_clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()
    
    If bRet = True Then
        DoUpdates
        SaveChangeRequest
        SetReturnCode
        Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub dgBandedParameter_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgBandedParameter.ValidateRow(nCurrentRow)
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New BandedGlobalParametersTable
    
    If m_bIsEdit = True Then
        PopulateScreen
        SetEditState
    Else
        SetAddState
    End If
    
    m_sOriginalName = txtEditBanded(BANDED_NAME).Text
    m_sOriginalDescription = txtDescription.Text 'txtEditBanded(BANDED_DESCRIPTION).Text
    m_sOriginalStartDate = txtEditBanded(BANDED_DATE).Text
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub SetupDBControl()
    On Error GoTo Failed
    Dim colTables As New Collection
    Dim colDataControls As New Collection
    
    colTables.Add m_clsTableAccess
    colDataControls.Add dgBandedParameter

    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If
    
    If m_bIsEdit = True Then
       PopulateScreenFields
    End If

    SetGridFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' Called from FormProcessing
Public Sub SetAddState()
    SetupDBControl
End Sub

' Called from FormProcessing
Public Sub SetEditState()
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDBControl
    dgBandedParameter.Enabled = True
    txtEditBanded(BANDED_NAME).Enabled = False
    txtEditBanded(BANDED_DATE).Enabled = False


End Sub
Public Sub PopulateScreen()
    On Error GoTo Failed
    
    SetupDBControl
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetGridFields()
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    
    ' First, Banded Name. Not visible, but needs the Fee Set copied in
    fields.sField = "Name"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtEditBanded(BANDED_NAME).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' HighBand
    fields.sField = "HighBand"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "High Band must be entered"
    fields.sTitle = "High Band"
    colFields.Add fields
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = "GBParamStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtEditBanded(BANDED_DATE).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Next, Description
    fields.sField = "Description"
    fields.bRequired = False
    fields.bVisible = False
    fields.sDefault = txtDescription.Text 'txtEditBanded(BANDED_DESCRIPTION).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Amount
    fields.sField = "Amount"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Amount"
    colFields.Add fields
    
    ' Percentage
    fields.sField = "Percentage"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Percentage"
    colFields.Add fields
    
    ' Maximum
    fields.sField = "Maximum"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Value must be entered"
    fields.sTitle = "Maximum"
    colFields.Add fields
    
    ' Boolean
    fields.sField = "Boolean"
    fields.bRequired = False
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Boolean Text
    fields.sField = "BooleanText"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Boolean"
    fields.sOtherField = "Boolean"
    g_clsCombo.FindComboGroup "Boolean", colComboValues, colComboIDS
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    colFields.Add fields
    
    ' String
    fields.sField = "String"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "String"
    fields.sOtherField = "Boolean"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    
    colFields.Add fields
    
    dgBandedParameter.SetColumns colFields, "GlobalBandedParameter", "Global Parameters"
End Sub
Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim clsBandedParameter As BandedGlobalParametersTable
    
    Set clsBandedParameter = m_clsTableAccess
    txtEditBanded(BANDED_NAME).Text = clsBandedParameter.GetName()
    txtDescription.Text = clsBandedParameter.GetDescription()
    txtEditBanded(BANDED_DATE).Text = clsBandedParameter.GetStartDate()
    
    PopulateScreenFields = True
    Exit Function
Failed:
    MsgBox "PopulateScreenFields: Error is, " + Err.DESCRIPTION
    PopulateScreenFields = False
End Function
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
       bRet = dgBandedParameter.ValidateRows()
    
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "High Band must be unique"
        End If
    
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub txtEditBanded_LostFocus(Index As Integer)
    SetDataGridState
End Sub
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgBandedParameter.Enabled
        dgBandedParameter.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                dgBandedParameter.AddRow False
            End If
        End If
    End If
End Sub
Private Sub txtEditBanded_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtEditBanded(Index).ValidateData()
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sSetNumber As String
    Dim sStartDate As String
    
    sSetNumber = Me.txtEditBanded(BANDED_NAME).Text
    sStartDate = Me.txtEditBanded(BANDED_DATE).Text
    
    Set colMatchValues = New Collection
    colMatchValues.Add sSetNumber
    colMatchValues.Add sStartDate
    m_clsTableAccess.SetKeyMatchValues colMatchValues

    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

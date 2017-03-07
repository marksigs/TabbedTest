VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditDistChannel 
   Caption         =   "Add/Edit Distribution Channel"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   Icon            =   "frmEditDistChannel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maintain Non-Working Days"
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   8115
      Begin MSGOCX.MSGDataGrid dgDistChannel 
         Height          =   2775
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
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
      Begin MSComctlLib.TabStrip tbsDistributionChannel 
         Height          =   3375
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5953
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Weekday"
               Key             =   "Weekday"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Other"
               Key             =   "Other"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1455
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtDistChannel 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   180
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtDistChannel 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   6195
      _ExtentX        =   10927
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
      MaxLength       =   30
   End
   Begin MSGOCX.MSGComboBox cboCountry 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1020
      Width           =   6195
      _ExtentX        =   10927
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
   Begin MSGOCX.MSGEditBox txtDistChannel 
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
      MaxLength       =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Country"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Application No. Prefix"
      Height          =   360
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Channel Name"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Channel ID"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "frmEditDistChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmDistChannel
' Description   :   Form which allows the user to edit and add distribution channel details
' Change history
' Prog      Date        Description
' AA        06/12/00    Added Non workingday datagrid
' DJP       21/06/01    Make Cancel button default for Cancel
' STB       27/11/01    SYS3074 - restricted length of application prefix to 2 chars.
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const CHANNEL_ID = 0
Private Const CHANNEL_NAME = 1
Private Const CHANNEL_APP_PREFIX = 2
Private Const OTHER_TYPE_KEY = "Other"

' private enum's
Private Enum WeekDayType
    WorkingDay = 1
    nonWorkingDay = 2
End Enum

' Private data
Private m_bIsEdit As Boolean
Private m_clsDistChannel As New DistributionChannelTable
Private m_clsNonWorkingDay As New NonWorkingDayTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_sDistributionChannel As String
Private m_sDistributionChannelName As String
Private m_WeekDayType As WeekDayType
Private m_rsWeekday As ADODB.Recordset
Private m_rsNonWeekday As ADODB.Recordset

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cboCountry_Validate(Cancel As Boolean)
    Cancel = Not cboCountry.ValidateData()

    If Cancel = False Then
        SetDataGridState
    End If
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function DoesRecordExist() As Boolean
    Dim bRet As Boolean
    Dim sChannelID As String
    Dim sDistributionChannelName As String
    Dim clsTableAccess As TableAccess
    Dim col As New Collection

    On Error GoTo Failed
    bRet = False
    sChannelID = txtDistChannel(CHANNEL_ID).Text

    If Len(sChannelID) > 0 Then
        col.Add sChannelID
    
        Set clsTableAccess = m_clsDistChannel
        bRet = Not clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            g_clsErrorHandling.DisplayError "Distribution Channel already exists - please enter a unique combination"
            txtDistChannel(CHANNEL_ID).SetFocus
        End If
        
        sDistributionChannelName = txtDistChannel(CHANNEL_NAME).Text
        
        If bRet = True And m_sDistributionChannelName <> sDistributionChannelName Then
            bRet = CheckForDuplicateValues
        End If
    End If

    DoesRecordExist = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean


    bRet = DoOKProcessing()


    If bRet = True Then
         Hide
     End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Private Function DoOKProcessing() As Boolean
    Dim bSuccess As Boolean
    Dim bShowError As Boolean
    
    On Error GoTo Failed
    bShowError = True

    bSuccess = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
        
    If bSuccess Then
        bSuccess = ValidateScreenData()
        
        If bSuccess Then
            DoGridUpdates
            SaveScreenData
            DoUpdates
            SaveChangeRequest
            SetReturnCode
            SaveNonWorkingDays
        End If
    End If
    
    DoOKProcessing = bSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsDistChannel
    
    Set colValues = New Collection
    colValues.Add Me.txtDistChannel(CHANNEL_ID).Text
    
    clsTableAccess.SetKeyMatchValues colValues
    
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsDistChannel
    clsTableAccess.Update
    Set dgDistChannel.DataSource = Nothing
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoGridUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    Dim colValues As New Collection
    Dim sChannelID As String

    Set clsBandedTable = m_clsNonWorkingDay

    sChannelID = txtDistChannel(CHANNEL_ID).Text

    If Len(sChannelID) > 0 Then
        colValues.Add sChannelID

        clsBandedTable.SetUpdateValues colValues

        ' We only want to check that the record exists if the name or start date has changed
        clsBandedTable.SetUpdateSets
        clsBandedTable.DoUpdateSets
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub dgDistChannel_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgDistChannel.ValidateRow(nCurrentRow)
    End If
End Sub
Private Sub PopulateCountry()
    On Error GoTo Failed
    Dim colFields As New Collection
    Dim clsCountry As New CountryTable
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    ' Populate the Counry combo
    
    Set clsTableAccess = clsCountry
    
    Set rs = clsTableAccess.GetTableData(POPULATE_ALL)
    
    If Not rs Is Nothing Then
        If clsTableAccess.RecordCount() > 0 Then
            Set colFields = clsCountry.GetComboFields()
            g_clsFormProcessing.PopulateComboFromTable cboCountry, clsCountry, colFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "No Country entries found"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Countries recordset is empty"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function GetTabKey(enumKey As WeekDayType) As String
    GetTabKey = "Key" & enumKey
End Function

Private Sub Form_Load()
    On Error GoTo Failed
    
    SetReturnCode MSGFailure
    
    
    Set m_clsDistChannel = New DistributionChannelTable
    
    ' Set the the datagrid
    tbsDistributionChannel.Tabs(WorkingDay).Key = GetTabKey(WorkingDay)
    tbsDistributionChannel.Tabs(nonWorkingDay).Key = GetTabKey(nonWorkingDay)
    m_WeekDayType = WorkingDay
    
    ' Populate country table
    PopulateCountry
    PopulateDistributionChannel
    PopulateNonWorkingDay
    

    SetGridDataSource
    SetGridFields
    ' Do add/edit specific processing
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    m_sDistributionChannel = txtDistChannel(CHANNEL_ID).Text
    m_sDistributionChannelName = txtDistChannel(CHANNEL_NAME).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

' Called from FormProcessing
Public Sub SetAddState()
    On Error GoTo Failed
    
    CmdAnother.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' Called from FormProcessing
Public Sub SetEditState()
    On Error GoTo Failed
        
    txtDistChannel(CHANNEL_ID).Enabled = False
    CmdAnother.Enabled = False
    PopulateScreenFields
    
    SetDataGridState
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    Dim clsCombo As New ComboUtils
    Dim sCombo As String
    
    ' Next, Non Working TypeID
    fields.sField = "NonWorkingType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Non Working Type Text
    fields.sField = "NonWorkingTypeText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Non Working Type must be entered"
    fields.sTitle = "Non Working Type"
    fields.sOtherField = "NonWorkingType"
    
    sCombo = "NonWorkingDayType"
    
    If m_WeekDayType = nonWorkingDay Then
        clsCombo.FindComboGroup sCombo, colComboValues, colComboIDS, RECURRING_NONWORKING_DAY, False, , True
    Else
        clsCombo.FindComboGroup sCombo, colComboValues, colComboIDS, RECURRING_NONWORKING_DAY, True, , True
    End If
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    colFields.Add fields

    ' First, Channel ID
    fields.sField = "ChannelID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtDistChannel(CHANNEL_ID).Text
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
   ' Day of Week Text

    fields.sField = "WeekdayText"
    fields.bRequired = False
    fields.bDisabled = True
    fields.sDefault = ""
    fields.sError = "Day of Week must be entered"
    fields.bDateField = False
    fields.sTitle = "Weekday"
    'fields.bDisabled = True
    fields.sOtherField = ""
        
    If m_WeekDayType = nonWorkingDay Then
        fields.bVisible = False
    Else
        fields.bVisible = True
    End If
    colFields.Add fields
    
    'Date/first occurence
    fields.sField = "NONWORKINGDAYDATE"
    fields.bRequired = True
    fields.bDisabled = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.bDateField = True
    fields.sError = "Date of First Occurence must be entered"
    If m_WeekDayType = nonWorkingDay Then
        fields.sTitle = "Date"
    Else
        fields.sTitle = "Date of First Occurence"
    End If
    
    fields.sOtherField = "WeekDay"
    Set fields.colComboValues = Nothing
    colFields.Add fields

    
    ' Next, Day of Week
    fields.sField = "Weekday"
    fields.bRequired = True
    fields.bDisabled = False
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    fields.bDateField = False
    g_clsCombo.FindComboGroup "DaysOfWeek", colComboValues, colComboIDS, , , "ID"
        
    If colComboValues.Count > 0 And colComboIDS.Count > 0 Then
        Set fields.colComboValues = colComboValues
        Set fields.colComboIDS = colComboIDS
        colFields.Add fields
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate combo group" + sCombo
    End If
    
    If m_WeekDayType = nonWorkingDay Then
        fields.sOtherField = ""
    Else
        fields.sOtherField = "WeekDayText"
    End If
    
    colFields.Add fields

    
    'Description
    fields.sField = "NONWORKINGDAYDESCRIPTION"
    fields.bRequired = False
    fields.bVisible = True
    fields.bDateField = False
    fields.sDefault = ""
    fields.sTitle = "Description"
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    colFields.Add fields

    dgDistChannel.SetColumns colFields, "Weekday" & m_WeekDayType, "Distribution Channels"
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    ' If adding, create a new record
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsDistChannel
    End If
        
    m_clsDistChannel.SetChannelID txtDistChannel(CHANNEL_ID).Text
    m_clsDistChannel.SetChannelName txtDistChannel(CHANNEL_NAME).Text

    g_clsFormProcessing.HandleComboExtra Me.cboCountry, vTmp, GET_CONTROL_VALUE
    m_clsDistChannel.SetCountry CStr(vTmp)

    m_clsDistChannel.SetAppPrefix txtDistChannel(CHANNEL_APP_PREFIX).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim vTmp As Variant
    txtDistChannel(CHANNEL_ID).Text = m_clsDistChannel.GetChannelID()
    txtDistChannel(CHANNEL_NAME).Text = m_clsDistChannel.GetChannelName()

    vTmp = m_clsDistChannel.GetCountry()
    g_clsFormProcessing.HandleComboExtra Me.cboCountry, vTmp, SET_CONTROL_VALUE
    txtDistChannel(CHANNEL_APP_PREFIX).Text = m_clsDistChannel.GetAppPrefix()

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    Dim newWeekdayType As WeekDayType
    On Error GoTo Failed
    bRet = False

    bRet = ValidateNonWorkingDays

    If bRet Then
        If m_WeekDayType = nonWorkingDay Then
            newWeekdayType = WorkingDay
        Else
            newWeekdayType = nonWorkingDay
        End If
        
        tbsDistributionChannel.Tabs(newWeekdayType).Selected = True
        
        If bRet = True Then
            bRet = ValidateNonWorkingDays
        End If
        
        If bRet = True And m_bIsEdit = False Then
            bRet = DoesRecordExist()
        End If
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub tbsDistributionChannel_Click()
    
    Dim bRet As Boolean
    On Error GoTo Failed

    If GetTabKey(m_WeekDayType) <> tbsDistributionChannel.SelectedItem.Key Then
        If m_WeekDayType = nonWorkingDay Then
            m_WeekDayType = WorkingDay
        Else
            m_WeekDayType = nonWorkingDay
        End If
        
        SetGridDataSource
        SetGridFields
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub txtDistChannel_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtDistChannel(Index).ValidateData()
    
    If Cancel = False Then
'        SetDataGridState
    End If
End Sub
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)

    If bRet = True Then
        Dim bEnabled As Boolean

        bEnabled = dgDistChannel.Enabled
        dgDistChannel.Enabled = True

        SetGridFields

        If m_bIsEdit = False Then
            If bEnabled = False Then
                If dgDistChannel.Visible = True And dgDistChannel.Enabled = True Then
                    dgDistChannel.SetFocus
                End If
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
Private Sub cmdAnother_Click()
On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
    
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
        PopulateNonWorkingDay
        SetGridDataSource
        SetGridFields
        
        tbsDistributionChannel.Tabs(WorkingDay).Selected = True
        
        txtDistChannel(CHANNEL_ID).SetFocus
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveNonWorkingDays
' Description   : Updates the nonworkingday table with all nonworkingday records (using function (UpdateRecordset)
'                 for the current distribution channel and Non working type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveNonWorkingDays()
    
    On Error GoTo Failed
    Dim colUpdateValues As Collection
    Dim clsBandedTable As BandedTable
  
    m_sDistributionChannel = txtDistChannel(CHANNEL_ID).Text
    
    ' DJP first the non weekday
    UpdateRecordset m_rsNonWeekday
    TableAccess(m_clsNonWorkingDay).Update
    'Now the weekday
    UpdateRecordset m_rsWeekday
    TableAccess(m_clsNonWorkingDay).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckForDuplicateValues
' Description   : Checks for duplicate values such as Distribution Channel Name. The value
'                 to check for duplicates to is specified in the colvalues collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CheckForDuplicateValues() As Boolean
    Dim bRet As Boolean
    Dim colFields As Collection
    Dim colValues As Collection
    Dim sDistributionName As String
    Dim clsTableAccess  As TableAccess
    
    On Error GoTo Failed
    
    Set clsTableAccess = New DistributionChannelTable
    Set colFields = New Collection
    Set colValues = New Collection
    
    sDistributionName = txtDistChannel(CHANNEL_NAME).Text
    
    bRet = False
    colValues.Add sDistributionName
    colFields.Add m_clsDistChannel.GetNameField

    bRet = Not clsTableAccess.DoesRecordExist(colValues, colFields)
    
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Distribution Channel already exists - please enter a unique name"
        txtDistChannel(CHANNEL_NAME).SetFocus
    End If
    
    CheckForDuplicateValues = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : UpdateRecordset
' Description   : Commits the data in the datgrid record source to the database. The recordset
'                 to be update is passed to the function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UpdateRecordset(rs As ADODB.Recordset)
    
    Dim colUpdateValues As Collection
    Dim clsBandedTable As BandedTable
    
    On Error GoTo Failed
    Set colUpdateValues = New Collection
    
    colUpdateValues.Add m_sDistributionChannel
    Set clsBandedTable = m_clsNonWorkingDay
    
    TableAccess(m_clsNonWorkingDay).SetRecordSet rs
    clsBandedTable.SetUpdateValues colUpdateValues
    clsBandedTable.SetUpdateSets
    clsBandedTable.DoUpdateSets
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridDataSource
' Description   : Sets the datagrid datasource to the correct recordset (set in PopulateNonWorkingDay)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridDataSource(Optional bEmpty As Boolean = False)
    On Error GoTo Failed
    
    If m_WeekDayType = nonWorkingDay Then
        If Not m_rsNonWeekday Is Nothing Then
            Set dgDistChannel.DataSource = m_rsNonWeekday
        Else
            g_clsErrorHandling.RaiseError errRecordSetEmpty
        End If
    ElseIf m_WeekDayType = WorkingDay Then
        If Not m_rsNonWeekday Is Nothing Then
            Set dgDistChannel.DataSource = m_rsWeekday
        Else
            g_clsErrorHandling.RaiseError errRecordSetEmpty
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unknown working day type: " & m_WeekDayType
    End If
    
    TableAccess(m_clsNonWorkingDay).SetRecordSet dgDistChannel.DataSource
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateNonWorkingDay
' Description   : Populates the datagrid with the relevant data (ie Weekday tab with weekday data)
'                 Populates an empty recordset if the form is in Add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateNonWorkingDay()
    Dim sChannel As String
    
    If m_bIsEdit = True Then
        sChannel = m_clsDistChannel.GetChannelID()
        
        m_clsNonWorkingDay.GetNonWorkingDays sChannel, RECURRING_NONWORKING_DAY
        Set m_rsWeekday = TableAccess(m_clsNonWorkingDay).GetRecordSet()
    
        m_clsNonWorkingDay.GetNonWorkingDays sChannel
        Set m_rsNonWeekday = TableAccess(m_clsNonWorkingDay).GetRecordSet()
    
    Else
        Dim colValues As New Collection
        Dim clsTableAccess As TableAccess
        
        TableAccess(m_clsNonWorkingDay).GetTableData POPULATE_EMPTY
        'Set m_rsNonWeekday = TableAccess(m_clsNonWorkingDay).GetTableData(POPULATE_EMPTY)
        Set m_rsWeekday = TableAccess(m_clsNonWorkingDay).GetRecordSet().Clone()
    
        TableAccess(m_clsNonWorkingDay).GetTableData POPULATE_EMPTY
        'Set m_rsNonWeekday = TableAccess(m_clsNonWorkingDay).GetTableData(POPULATE_EMPTY)
        Set m_rsNonWeekday = TableAccess(m_clsNonWorkingDay).GetRecordSet()
    
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateDistributionChannel
' Description   : Populates the form data(ie the channel name textbox etc), with the relevant
'                 Data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub PopulateDistributionChannel()
    On Error GoTo Failed
    Dim sChannel As String
    
    If m_bIsEdit = True Then
        TableAccess(m_clsDistChannel).SetKeyMatchValues m_colKeys
        TableAccess(m_clsDistChannel).GetTableData
        
    Else
        TableAccess(m_clsDistChannel).GetTableData POPULATE_EMPTY
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateNonWorkingDays()
' Description   : Called once for each tabstrip. Validates the data entered in the datagrid
'                 and returns a boolean indicating whether the data is valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateNonWorkingDays() As Boolean

    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = True

    ' Validate grid, whichever one it is
    bRet = dgDistChannel.ValidateRows()
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsNonWorkingDay)
        
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Date of First Occurence and Non Working Type must be unique"
        End If
    
    End If
    
    ValidateNonWorkingDays = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditWorkingHoursType 
   Caption         =   "Add/Edit Working Hours Type"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "frmEditWorkingHoursType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkBankHoliday 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Holiday Indicator"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7350
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6060
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Working Hours"
      Height          =   3795
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   8295
      Begin MSGOCX.MSGDataGrid dgWorkingHours 
         Height          =   3375
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5953
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
   Begin MSGOCX.MSGEditBox txtWorkingHours 
      Height          =   315
      Index           =   0
      Left            =   2460
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
   Begin MSGOCX.MSGEditBox txtWorkingHours 
      Height          =   315
      Index           =   1
      Left            =   2460
      TabIndex        =   1
      Top             =   420
      Width           =   5655
      _ExtentX        =   9975
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
      MaxLength       =   50
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Working Hours Type"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "frmEditWorkingHoursType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditWorkingHoursType
' Description   :
'
' Change history
' Prog      Date        Description
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Working Hours Text Fields
Private Const WORKING_HOURS_TYPE = 0
Private Const DESCRIPTION = 1

Private m_bIsEdit As Boolean
Private m_clsWorkingHours As New WorkingHoursTable
Private m_clsWorkingHourType As WorkingHourTypeTable
Private m_ReturnCode As MSGReturnCode
Private colComboValues As New Collection
Private colComboIDS As New Collection
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsWorkingHourType = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()
        
        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
            
        End If
    End If
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim nCount As Integer
    Dim bRet As Boolean

    nCount = 0
    bRet = True

    bRet = dgWorkingHours.ValidateRows()

    If bRet = True Then
        If m_bIsEdit = False Then
            bRet = Not DoesRecordExist()

            If bRet = False Then
                txtWorkingHours(WORKING_HOURS_TYPE).SetFocus
                g_clsErrorHandling.RaiseError errGeneralError, "Working Hour Type already exists"
            End If
        End If
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsWorkingHours)
        
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Day of week must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sType As String
    Dim clsTableAccess As TableAccess
    Dim col As New Collection

    bRet = False

    sType = Me.txtWorkingHours(WORKING_HOURS_TYPE).Text

    If Len(sType) > 0 Then
        col.Add sType

        Set clsTableAccess = m_clsWorkingHourType
        bRet = clsTableAccess.DoesRecordExist(col)
    End If

    DoesRecordExist = bRet
End Function
Private Sub dgWorkingHours_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgWorkingHours.ValidateRow(nCurrentRow)
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    Set m_clsWorkingHourType = New WorkingHourTypeTable
    g_clsCombo.FindComboGroup "DaysOfWeek", colComboValues, colComboIDS
    
    If m_bIsEdit = True Then
        SetEditState
        PopulateScreen
        SetDataGridState
    Else
        SetAddState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Function PopulateWorkingHours(sType As String) As Boolean
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    Dim rs As ADODB.Recordset

    bRet = False
    Set clsTableAccess = m_clsWorkingHours
    If Len(sType) > 0 Then
        colValues.Add sType
    
        clsTableAccess.SetKeyMatchValues colValues
    
        Set rs = clsTableAccess.GetTableData()
    
        If Not rs Is Nothing Then
            Set dgWorkingHours.DataSource = rs
            bRet = True
        End If
    End If
    
    PopulateWorkingHours = bRet
End Function
Public Sub SetAddState()
    PopulateWorkingHours "null"

    SetGridFields
End Sub
Public Sub SetEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sType As String
    Dim colValues As New Collection

    Set clsTableAccess = m_clsWorkingHourType
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    txtWorkingHours(WORKING_HOURS_TYPE).Enabled = False
    ' First, the Working Hour Type table
    
    Set rs = clsTableAccess.GetTableData()

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            ' Working Hours
            sType = m_clsWorkingHourType.GetWorkingHourType()
            
            If Len(sType) > 0 Then
                PopulateWorkingHours sType
            End If
        End If
    End If

End Sub
Public Sub PopulateScreen()
'    Dim rs As adodb.Recordset
    Dim bRet As Boolean
    Dim colTables As New Collection
    On Error GoTo Failed

    bRet = PopulateScreenFields()

    Exit Sub
Failed:
    MsgBox "PopulateScreen: Error is, " + Err.DESCRIPTION
End Sub
Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    txtWorkingHours(WORKING_HOURS_TYPE).Text = m_clsWorkingHourType.GetWorkingHourType()
    txtWorkingHours(DESCRIPTION).Text = m_clsWorkingHourType.GetDescription()

    vTmp = m_clsWorkingHourType.GetBankHolidayIndicator()
    g_clsFormProcessing.HandleCheckBox chkBankHoliday, vTmp, SET_CONTROL_VALUE

    PopulateScreenFields = True
    Exit Function
Failed:
    PopulateScreenFields = False
End Function
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsWorkingHourType
    End If

    m_clsWorkingHourType.SetWorkingHourType txtWorkingHours(WORKING_HOURS_TYPE).Text
    m_clsWorkingHourType.SetDescription txtWorkingHours(DESCRIPTION).Text

    g_clsFormProcessing.HandleCheckBox chkBankHoliday, vTmp, GET_CONTROL_VALUE
    m_clsWorkingHourType.SetBankHolidayIndicator CStr(vTmp)

    ' First, working hour type
    Set clsTableAccess = m_clsWorkingHourType
    clsTableAccess.Update

    ' Working Hours
    Set clsTableAccess = m_clsWorkingHours
    clsTableAccess.Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub txtWorkingHours_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtWorkingHours(Index).ValidateData()

    If Cancel = False Then
        SetDataGridState
    End If

End Sub
Private Sub SetGridFields()
    Dim bRet As Boolean
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsCombo As New ComboUtils
    Dim sCombo As String
    bRet = True

    ' Day of the Week Text
    fields.sField = "DayOfTheWeekText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Day of the Week must be entered"
    fields.sTitle = "Day of the Week"
    fields.sOtherField = "DayOfTheWeek"

    If colComboValues.Count > 0 And colComboIDS.Count > 0 Then
        Set fields.colComboValues = colComboValues
        Set fields.colComboIDS = colComboIDS
        colFields.Add fields
    Else
        MsgBox "Unable to locate combo group" + sCombo
    End If
    
    ' Working Hour Type
    fields.sField = "WorkingHourType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtWorkingHours(WORKING_HOURS_TYPE).Text
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Day of the Week
    fields.sField = "DayOfTheWeek"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' From Hour
    fields.sField = "FromHour"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "From Hour must be entered"
    fields.sTitle = "From Hour"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = "0"
    fields.sMaxValue = "23"
    colFields.Add fields

    ' From Minute
    fields.sField = "FromMinute"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "From Minute must be entered"
    fields.sTitle = "From Minute"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = "0"
    fields.sMaxValue = "59"
    colFields.Add fields

    ' To Hour
    fields.sField = "ToHour"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "To Hour must be entered"
    fields.sTitle = "To Hour"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = "0"
    fields.sMaxValue = "23"
    colFields.Add fields

    ' To Minute
    fields.sField = "ToMinute"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "To Minute must be entered"
    fields.sTitle = "To Minute"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = "0"
    fields.sMaxValue = "59"
    colFields.Add fields

    Me.dgWorkingHours.SetColumns colFields, "EditWorkingHours", "Working Hours"
End Sub
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False

    Dim bEnabled As Boolean

    ' Unit ID has to be entered

    If Len(Me.txtWorkingHours(WORKING_HOURS_TYPE).Text) > 0 Then
        bEnabled = dgWorkingHours.Enabled
        dgWorkingHours.Enabled = True

        SetGridFields

        If m_bIsEdit = False Then
            If bEnabled = False Then
                'dgWorkingHours.AddRow
                'dgWorkingHours.SetFocus
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
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add txtWorkingHours(WORKING_HOURS_TYPE).Text
    Set clsTableAccess = m_clsWorkingHours
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsWorkingHours
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

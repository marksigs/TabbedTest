VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditRate 
   Caption         =   " Add/Edit Base Rate"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmEditRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraRateChange 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin MSMask.MaskEdBox txtStartTime 
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   750
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSGOCX.MSGEditBox txtAppliedDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Mask            =   "##/##/####"
         TextType        =   1
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Enabled         =   0   'False
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
      Begin MSGOCX.MSGEditBox txtStartDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Mask            =   "##/##/####"
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
      Begin MSGOCX.MSGComboBox cboType 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   300
         Width           =   1995
         _ExtentX        =   3519
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
         Text            =   ""
      End
      Begin MSGOCX.MSGEditBox txtRateID 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSGOCX.MSGEditBox txtRate 
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         TextType        =   2
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         MinValue        =   "-1"
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
      Begin MSGOCX.MSGEditBox txtDescription 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1260
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
         MaxLength       =   40
      End
      Begin VB.Label Label1 
         Caption         =   "Start Time"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAppliedDate 
         Caption         =   "Applied Date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblStartDate 
         Caption         =   "Start Date"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         Height          =   195
         Left            =   4440
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblRate 
         Caption         =   "Rate"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblRateID 
         Caption         =   "Rate ID"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmEditRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditRate
' Description   : Form which allows the adding/editing of Base Rates
'
' Change history
' Prog  Date      Description
' AA    13/02/01  Added Form
' DJP   27/06/01  SQL Server port
' DJP   14/11/01  SYS2996 Change <None> to <Select> in combos, but put into
'                 constants.bas
' SA    17/12/01  SYS3504 Remove BaseRateSet combo and Base Rate Bands list
'                 view.
' STB   04/02/02  SYS3857 - Base Rates can now be promoted.
' STB   08/03/02  SYS4251 - Applied Date set in Add mode and tidied-up form.
' STB   20/03/02  SYS4251 - Applied date only set when Rate ID is first
'                 entered.
' STB   20/03/02  SYS4300 - StartDate and Rate fields enabled when required.
' GHun  15/03/02  SYS4394 - Base Rate pop up screen does not reflect the Applied date.
' JD    28/07/2004 BMIDS827 check time is entered. Default to 00:00:00 if it is not set
' JD    28/07/04   BMIDS827 validate the time
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        18/05/2007  VR262 Enable the Description field on Add/Edit Base Rate in Supervisor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Name of Rate Type combo group.
Private Const RATETYPE_COMBOGROUP As String = "RateType"

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'Underlying table object(s).
Private m_clsRate As RateTable

'A collection of primary keys to identify the current record.
Private m_colKeys  As Collection

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'SYS3504 When checking for Dup records, only check in edit mode if the date has changed!
Private m_sOriginalStartDate As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return code for the form for any calling method to
'                 check.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(ByVal uReturn As MSGReturnCode)
    m_uReturnCode = uReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the success code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the record, closing the form if everything
'                 is okay.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    On Error GoTo Failed
    
    'Delegate this to another routine.
    DoOKProcessing
    
    'Hide the form and return to the caller.
    Hide
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create the underlying table objects and setup the controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    On Error GoTo Failed
    
    'Create the underlying table object.
    Set m_clsRate = New RateTable
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
        
    PopulateScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Load the relevant record into the underlying table object and
'                 setup any controls for edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    'Load the Base Rate record.
    TableAccess(m_clsRate).SetKeyMatchValues m_colKeys
    TableAccess(m_clsRate).GetTableData POPULATE_KEYS
    
    'Ensure a record was found.
    If TableAccess(m_clsRate).RecordCount = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Rate ID not Found"
    End If
    
    'SYS3504 Rate Id cannot be changed once added.
    txtRateID.Enabled = False
        
    'SYS3504 If the start date has past, don't allow editing of any kind!
    If CDate(m_clsRate.GetStartDate) < Now Then
        txtStartDate.Enabled = False
        txtStartTime.Enabled = False  'JD BMIDS749
        txtAppliedDate.Enabled = False
' TW 18/05/2007 VR262
'        txtDescription.Enabled = False
' TW 18/05/2007 VR262 End
        cboType.Enabled = False
        txtRate.Enabled = False
' TW 18/05/2007 VR262
'        cmdOK.Enabled = False
' TW 18/05/2007 VR262 End
    End If
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Create a blank record and setup any controls for add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    
    'Create a blank record.
    g_clsFormProcessing.CreateNewRecord m_clsRate
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populate the screen from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    
    Dim vTmp As Variant
    Dim clsComboUtils As ComboUtils
    
    On Error GoTo Failed
    
    'Create a combo helper class.
    Set clsComboUtils = New ComboUtils
    
    'Populate Rate Type Combo
    g_clsFormProcessing.PopulateCombo RATETYPE_COMBOGROUP, cboType
    
    'Rate ID.
    txtRateID.Text = m_clsRate.GetRateID
    
    'Interest Rate.
    txtRate.Text = m_clsRate.GetInterestRate
    
    'Start Date.
    txtStartDate.Text = Format(m_clsRate.GetStartDate, "dd/mm/yyyy") 'JD BMIDS749
    If Not m_clsRate.GetStartDate = "" Then
        txtStartTime.Text = Format(m_clsRate.GetStartDate, "hh:mm:ss")
    End If
    
    'SYS3504 - Store the start date.
    'm_sOriginalStartDate = m_clsRate.GetStartDate
    m_sOriginalStartDate = txtStartDate.Text + " " + txtStartTime.Text  ' JD BMIDS749
    
    'Description.
    txtDescription.Text = m_clsRate.GetDescription
    
    'Get the RateType ValueID.
    vTmp = m_clsRate.GetRateType
    
    'Get the RateType ValidationID.
    If Len(vTmp) > 0 Then
        vTmp = clsComboUtils.GetValueNameFromValueID(CStr(vTmp), RATETYPE_COMBOGROUP)
    End If
    
    'RateType.
    g_clsFormProcessing.HandleComboText cboType, vTmp, SET_CONTROL_VALUE
    
    'Applied Date.
    'SYS4394 Format Applied Date to conform to expected format and remove any time information
    txtAppliedDate.Text = Format(m_clsRate.GetRateAppliedDate, "dd/mm/yyyy")
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Copy the screen fields to the table object and perform an
'                 update.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    
    Dim vval As Variant
    
    'First RateID
    m_clsRate.SetRateID txtRateID.Text
    
    'Next the Rate
    m_clsRate.SetInterestRate txtRate.Text
    
    'Start Date
    g_clsFormProcessing.HandleDate txtStartDate, vval, GET_CONTROL_VALUE
    m_clsRate.SetStartDate CDate(vval + " " + txtStartTime.Text) ' JD BMIDS749 added time
        
    'Applied Date - Only save if we're adding.
    If m_bIsEdit = False Then
        'Only set the Applied Date if this RateID is unique and we're adding.
        If m_clsRate.IsRateIDUnique(txtRateID.Text) Then
            txtAppliedDate.Text = Format(txtStartDate.Text, "dd/mm/yyyy")
        
            g_clsFormProcessing.HandleDateEx txtAppliedDate, vval, GET_CONTROL_VALUE
            m_clsRate.SetRateAppliedDate vval
        End If
    End If
    
    'Description
    m_clsRate.SetDecsription txtDescription.Text
    
    'Rate Type.
    g_clsFormProcessing.HandleComboExtra cboType, vval, GET_CONTROL_VALUE
    m_clsRate.SetRateType vval
             
    'Now update the underlying table object.
    TableAccess(m_clsRate).Update
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Validate and save the data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoOKProcessing()
    
    Dim bRet As Boolean
    
    'Validate the screen data.
    bRet = Not ValidateScreenData
    
    If bRet Then
        'Save the data to the database.
        SaveScreenData
        
        'Mark the record for promotion.
        SaveChangeRequest
        
        'Set the return code to success.
        SetReturnCode MSGSuccess
    End If
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Ensure the contents of the screen fields are valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    On Error GoTo Failed

    Dim bRet As Boolean
    Dim clsRate As RateTable
    
    Set clsRate = New RateTable
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, True)

    If bRet Then
        'SYS3504 if edit mode, only check record exists if date has changed
        'JD BMIDS827 check time is entered. Default to 00:00:00 if it is not set
        If txtStartTime.Text = "__:__:__" Then
            txtStartTime.Text = "00:00:00"
        End If
        'JD BMIDS827 validate time
        bRet = g_clsValidation.ValidateTime(txtStartTime.ClipText)
    
        If Not bRet Then
            txtStartTime.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Time must be between 00:00:00 and 23:59:59"
        Else
            Dim sStartTimeNow As String
            sStartTimeNow = txtStartDate.Text + " " + txtStartTime.Text  'JD BMIDS749 added time check
            If (m_bIsEdit And sStartTimeNow <> m_sOriginalStartDate) Or Not m_bIsEdit Then
                bRet = DoesRecordExist
            Else
                bRet = False
            End If
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesRecordExist
' Description   : Check for the existance of a record with the RateID and Start
'                 Date currently shown on the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist() As Boolean
    
    Dim bRet As Boolean
    Dim colValues As Collection
    Dim sStartDateTime As String
    
    If Len(txtRateID.Text) > 0 Then
        'Create a keys collection based on the screen values.
        sStartDateTime = txtStartDate.Text + " " + txtStartTime.Text  'JD BMIDS749 added time
        Set colValues = New Collection
        colValues.Add txtRateID.Text
        'colValues.Add txtStartDate.Text
        colValues.Add sStartDateTime
        
        'See if the record exists.
        bRet = TableAccess(m_clsRate).DoesRecordExist(colValues)
    
        'If the records exist raise an error.
        If bRet = True Then
            g_clsErrorHandling.RaiseError errGeneralError, "RateID and Start Date must be unique"
            txtRateID.SetFocus
        End If
    End If

    DoesRecordExist = bRet
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Mark the record for promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()

    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    Dim sStartDateTime As String

    'Build a keys collection.
    Set colMatchValues = New Collection
    sStartDateTime = txtStartDate.Text + " " + txtStartTime.Text  'JD BMIDS749
    colMatchValues.Add txtRateID.Text
    'colMatchValues.Add txtStartDate.Text
    colMatchValues.Add sStartDateTime

    'Ensure the keys collection is set against the underlying table.
    Set clsTableAccess = m_clsRate
    clsTableAccess.SetKeyMatchValues colMatchValues

    'Mark this record for promotion.
    g_clsHandleUpdates.SaveChangeRequest m_clsRate

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtRateID_LostFocus
' Description   : If we're adding a new Rate and the ID exists, copy populate
'                 description, type and rate from any existing rate with the
'                 same id (use the latest record).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRateID_LostFocus()

    Dim vTmp  As Variant
    Dim clsRateTable As RateTable
    Dim clsComboUtils As ComboUtils
    
    'If we're editing.
    If m_bIsEdit = False Then
        Set clsRateTable = New RateTable
        
        'See if there is an existing rate with this RateID.
        If clsRateTable.GetLatestRate(txtRateID.Text) Then
            'Create a combo helper class.
            Set clsComboUtils = New ComboUtils
                        
            'Interest Rate.
            txtRate.Text = clsRateTable.GetInterestRate
            
            'Description.
            txtDescription.Text = clsRateTable.GetDescription
            
            'Get the RateType ValueID.
            vTmp = clsRateTable.GetRateType
            
            'Get the RateType ValidationID.
            If Len(vTmp) > 0 Then
                vTmp = clsComboUtils.GetValueNameFromValueID(CStr(vTmp), RATETYPE_COMBOGROUP)
            End If
            
            'RateType.
            g_clsFormProcessing.HandleComboText cboType, vTmp, SET_CONTROL_VALUE
            
            'Set the focus to the rate field.
            txtRate.SetFocus
            txtRate.SelStart = 1
            txtRate.SelLength = Len(txtRate.Text)
        End If
    End If
    
End Sub



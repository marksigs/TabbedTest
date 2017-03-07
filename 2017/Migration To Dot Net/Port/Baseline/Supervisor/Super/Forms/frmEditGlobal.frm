VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditGlobalFixed 
   Caption         =   "Fixed Global Parameters Add/Edit"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmEditGlobal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGlobalString 
      Height          =   1155
      Left            =   1440
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   4155
   End
   Begin MSGOCX.MSGComboBox cboBoolean 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
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
      Text            =   ""
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3075
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4395
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   1155
      Left            =   1440
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   4155
   End
   Begin MSGOCX.MSGEditBox txtEditGlobal 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      _ExtentX        =   2037
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
   Begin MSGOCX.MSGEditBox txtEditGlobal 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   540
      Width           =   1035
      _ExtentX        =   7117
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
   Begin MSGOCX.MSGEditBox txtEditGlobal 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      TextType        =   2
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
   Begin MSGOCX.MSGEditBox txtEditGlobal 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   2700
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MaxValue        =   "999"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   9
   End
   Begin MSGOCX.MSGEditBox txtEditGlobal 
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      MaxValue        =   "100"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   5
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "String"
      Height          =   255
      Index           =   6
      Left            =   300
      TabIndex        =   16
      Top             =   4020
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Boolean"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   15
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Maximum"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   14
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Percentage"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   13
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Amount"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   12
      Top             =   2340
      Width           =   1035
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   11
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   10
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   9
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmEditGlobalFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditGlobalFixed
' Description   : Generic parameter edit form.
'
' Change History
' Prog      Date        Description
' STB       05/12/01    SYS1664 - Altered length of percentage field (to cator for
'                       more decimal places) and did a general tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Control indexes.
Private Const NAME_VAL          As Long = 0
Private Const START_DATE_VAL    As Long = 1
Private Const AMOUNT_VAL        As Long = 2
Private Const PERCENTAGE_VAL    As Long = 3
Private Const MAXIMUM_VAL       As Long = 4
Private Const STRING_VAL        As Long = 6

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'Collection containing the primary-key values for the current record.
Private m_colKeys As Collection

'A status indicator to the forms caller.
Private m_uReturnCode As MSGReturnCode

'Underlying table object for the user record.
Private m_clsTableAccess As TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Set the edit/add state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the current state of the form (add or edit).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form, the return status will indicate failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate data and attempt to save. If successful, close the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Ensure all mandatory fields have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    'If the data was valid.
    If bRet = True Then
        'Validate all captured input.
        ValidateScreenData
        
        'Save the data.
        SaveScreenData
        
        'Ensure the record is flagged for promotion.
        SaveChangeRequest
        
        'Set the return status to success.
        SetReturnCode
        
        'Hide the form so control returns to its opener.
        Hide
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates any control values relevant to this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateScreenData()
    
    ' Only one of Amount, Percentage, Boolean and string can be entered.
    Dim nCount As Integer
    
    On Error GoTo Failed
    
    nCount = 0
    
    If Len(txtEditGlobal(PERCENTAGE_VAL).Text) > 0 Then
        nCount = nCount + 1
    End If

    If Len(txtEditGlobal(AMOUNT_VAL).Text) > 0 Then
        nCount = nCount + 1
    End If

    If Len(cboBoolean.Text) > 0 Then
        nCount = nCount + 1
    End If
    'BMIDS01033 Changed control to ordinary text control
    'If Len(txtEditGlobal(STRING_VAL).Text) > 0 Then
    '    nCount = nCount + 1
    'End If
    If Len(txtGlobalString.Text) > 0 Then
        nCount = nCount + 1
    End If
    'BMIDS01033 }
    
    If nCount <> 1 Then
        txtEditGlobal(AMOUNT_VAL).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "One of Amount, Percentage, Boolean and String can be entered"
    End If
    
    'If we're adding, ensure the record doesn't already exist.
    If m_bIsEdit = False Then
        ' Does the record exist already?
        
        If DoesRecordExist() Then
            txtEditGlobal(NAME_VAL).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Parameter exists - Please enter a unique combination of Name and Start Date"
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesRecordExist
' Description   : Ensures a duplicate record does not already exist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist()
    
    Dim bRet As Boolean
    Dim col As New Collection
    Dim clsGlobal As New FixedParametersTable
    
    'Add the current key values into a collection.
    col.Add txtEditGlobal(NAME_VAL).Text
    col.Add txtEditGlobal(START_DATE_VAL).Text
    
    'Check to see if this record already exists.
    bRet = TableAccess(clsGlobal).DoesRecordExist(col)
    
    'Return true to the caller if it does.
    DoesRecordExist = bRet
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create underlying objects and populate/initialse screen controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Assume the form fails until we successfully save some data.
    SetReturnCode MSGFailure
    
    'Cast a generic table interface onto a new instance of a parameter table object.
    Set m_clsTableAccess = New FixedParametersTable
    
    'Populate the boolean combo from the database.
    g_clsFormProcessing.PopulateCombo "Boolean", cboBoolean
    
    'If we're editing, intialise controls and populate them.
    If m_bIsEdit = True Then
        SetEditState
        PopulateScreen
    'Otherwise, we're adding, just initalise controls for the add state.
    Else
        SetAddState
    End If
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Configure the screen controls to work in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    
    On Error GoTo Failed
    
    'Add a record to the underlying table object.
    m_clsTableAccess.AddRow
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Configure the screen controls to work in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
    'Ensure the keys collection is set against the table class.
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    
    'Disable the parameter name control as it's part of the key. Note: The start
    'date/time control is still enabled, presumably to allow the parameter to be
    'copied into a new period?
    txtEditGlobal(NAME_VAL).Enabled = False

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtDescription_GotFocus
' Description   : When the forcus enters the description input control, ensure the entire
'                 contents are selected.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_GotFocus()
    
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription)

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetPercentageMax
' Description   : If a percentage value is being entered, ensure the Maximum control is
'                 enabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPercentageMax(ByVal Index As Integer)
    
    If Index = PERCENTAGE_VAL Then
        If Len(txtEditGlobal(Index).Text) > 0 Then
            txtEditGlobal(MAXIMUM_VAL).Enabled = True
            lblEditGlobal(MAXIMUM_VAL).Enabled = True
        Else
            txtEditGlobal(MAXIMUM_VAL).Text = ""
            txtEditGlobal(MAXIMUM_VAL).Enabled = False
            lblEditGlobal(MAXIMUM_VAL).Enabled = False
        End If
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtEditGlobal_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtEditGlobal_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not txtEditGlobal(Index).ValidateData()

    'If the data is valid, ensure the 'maximum' control is enabled/disabled.
    If Not Cancel Then
        SetPercentageMax Index
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreen
' Description   : *shrug*
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreen()
    
    Dim colTables As New Collection
    
    On Error GoTo Failed
    
    colTables.Add m_clsTableAccess
    
    g_clsFormProcessing.PopulateDBScreen colTables
    
    PopulateScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populate the screen fields from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    
    Dim clsFixedParameters  As FixedParametersTable
    
    On Error GoTo Failed
    
    'Get an interface on the underlying table object.
    Set clsFixedParameters = m_clsTableAccess

    'Update the screen controls from the underlying table object.
    txtEditGlobal(NAME_VAL).Text = clsFixedParameters.GetName()
    txtEditGlobal(START_DATE_VAL).Text = clsFixedParameters.GetStartDate()
    txtDescription.Text = clsFixedParameters.GetDescription()
    txtEditGlobal(AMOUNT_VAL).Text = clsFixedParameters.GetAmount()
    
    txtEditGlobal(PERCENTAGE_VAL).Text = clsFixedParameters.GetPercentage()
    txtEditGlobal(MAXIMUM_VAL).Text = clsFixedParameters.GetMaximum()
    g_clsFormProcessing.HandleComboExtra cboBoolean, clsFixedParameters.GetBoolean(), SET_CONTROL_VALUE
    'BMIDS01033 Changed to text control
    'txtEditGlobal(STRING_VAL).Text = clsFixedParameters.GetString()
    txtGlobalString.Text = clsFixedParameters.GetString()
    'BMIDS01033 }
    'If the underlying data is NOT a percentage, then disable the 'Maximum' control.
    SetPercentageMax PERCENTAGE_VAL
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Update the underlying table object from the screen values.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    
    Dim vTmp As Variant
    Dim clsParam As FixedParametersTable
    
    On Error GoTo Failed
    
    Set clsParam = m_clsTableAccess
        
    clsParam.SetName txtEditGlobal(NAME_VAL).Text
    clsParam.SetStartDate txtEditGlobal(START_DATE_VAL).Text
    clsParam.SetDescription txtDescription.Text
    clsParam.SetAmount txtEditGlobal(AMOUNT_VAL).Text
    clsParam.SetPercentage txtEditGlobal(PERCENTAGE_VAL).Text
    clsParam.SetMaximum txtEditGlobal(MAXIMUM_VAL).Text
    g_clsFormProcessing.HandleComboExtra cboBoolean, vTmp, GET_CONTROL_VALUE
    clsParam.SetBoolean CStr(vTmp)
    'BMIDS01033 SA Changed to text control {
    'clsParam.SetString txtEditGlobal(STRING_VAL).Text
    clsParam.SetString txtGlobalString.Text
    'BMIDS01033 }
    m_clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Set the return code of the form to success/failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Returns the success code of the form to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Ensures the keys collection is set against the underlying table and then
'                 calls into the promotion code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    
    Dim colMatchValues As Collection
    
    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add txtEditGlobal(NAME_VAL).Text
    colMatchValues.Add txtEditGlobal(START_DATE_VAL).Text
    
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

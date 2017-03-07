VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditConditions 
   Caption         =   "Add/Edit Conditions"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmEditConditions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9015
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   5700
      TabIndex        =   32
      Top             =   4680
      Width           =   1575
      Begin VB.OptionButton optConditional 
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   900
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optConditional 
         Caption         =   "Yes"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   7500
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Condition Settings"
      Height          =   2415
      Left            =   60
      TabIndex        =   21
      Top             =   3420
      Width           =   8835
      Begin MSGOCX.MSGComboBox cboDistChannel 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1920
         Width           =   2475
         _ExtentX        =   4366
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
      Begin VB.CheckBox chkDeleteFlag 
         Height          =   195
         Left            =   5700
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1560
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   -60
            TabIndex        =   16
            Top             =   -60
            Width           =   1575
            Begin VB.OptionButton optFreeFormat 
               Caption         =   "No"
               Height          =   255
               Index           =   0
               Left            =   900
               TabIndex        =   7
               Top             =   180
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optFreeFormat 
               Caption         =   "Yes"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   6
               Top             =   180
               Width           =   675
            End
         End
         Begin VB.OptionButton optDetailsRequired 
            Caption         =   "No"
            Height          =   255
            Index           =   3
            Left            =   900
            TabIndex        =   31
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton optDetailsRequired 
            Caption         =   "Yes"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   30
            Top             =   180
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.Frame fraDeleteFlag 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1500
         TabIndex        =   17
         Top             =   780
         Width           =   1575
         Begin VB.OptionButton optEditable 
            Caption         =   "Yes"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   4
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton optEditable 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin MSGOCX.MSGTextMulti txtRuleReference 
         Height          =   315
         Left            =   5700
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
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
         Text            =   ""
         MaxLength       =   30
      End
      Begin MSGOCX.MSGComboBox cboConditionType 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   2475
         _ExtentX        =   4366
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
      Begin VB.Label lblChannelID 
         Caption         =   "Channel Id"
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblConditional 
         Caption         =   "Conditional"
         Height          =   315
         Left            =   4380
         TabIndex        =   27
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblFreeFormat 
         Caption         =   "Free Format"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDeleteFlag 
         Caption         =   "Delete Flag"
         Height          =   315
         Left            =   4380
         TabIndex        =   25
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblEditable 
         Caption         =   "Editable"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblRuleReference 
         Caption         =   "Rule Reference"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblType 
         Caption         =   "Condition Type"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.Frame fraGeneral 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8835
      Begin MSGOCX.MSGEditBox txtName 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   900
         Width           =   7155
         _ExtentX        =   12621
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
      Begin VB.TextBox txtConditionRef 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   360
         Width           =   2475
      End
      Begin MSGOCX.MSGTextMulti txtDescription 
         Height          =   1515
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   2672
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
         MaxLength       =   5000
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label lblConditionReference 
         Caption         =   "Condition Reference"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEditConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditConditions
' Description   : Form which allows the adding/editing of Conditions
'
' Change history
' Prog      Date        Description
' AA        29/01/00    Added Form
' STB       06/12/01    SYS1942 - Another button commits current transaction.
' STB       04/03/02    SYS4156 - Increased max length of description control.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        09/09/02    BMIDS00313 BM004 - embed database values within conditions text.
'                       Changed DoOKProcessing Function
' GD        06/11/2002  BMIDS00858 - Make Description have max size of 5000
' BS        25/03/03     BM0282 Reset form return code on form load

Option Explicit

Private m_ReturnCode As MSGReturnCode
'm_bScreenUpdated is set to true when any field is changed!
Private m_bScreenUpdated As Boolean
Private m_clsConditions As ConditionsTable
Private m_colKeys  As Collection
Private m_bIsEdit As Boolean
Private m_sConditionReference As String

Private Const CHECKBOX_NO = 0
Private Const CHECKBOX_YES = 1

Private Const m_sComboGroup As String = "ConditionType"


Private Sub cboConditionType_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cboDistChannel_Click()
    m_bScreenUpdated = True
End Sub

Private Sub chkDeleteFlag_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cmdAnother_Click()

    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
    
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
        txtName.SetFocus
        chkDeleteFlag.Value = 0
        optFreeFormat(CHECKBOX_NO).Value = True
        optConditional(CHECKBOX_NO).Value = True
        Me.optEditable(CHECKBOX_NO).Value = True
    End If

    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo Failed
    
    Dim bRet As Boolean
    bRet = DoOKProcessing
    
    If bRet Then
        SetReturnCode
        Hide
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub Form_Load()
On Error GoTo Failed
    
    'BS BM0282 25/03/03
    'The default return code is failure (until data is successfully saved).
    SetReturnCode MSGFailure
    
    Set m_clsConditions = New ConditionsTable
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetFormEditState
    Else
        SetFormAddState
    End If
    
    
    
    m_bScreenUpdated = False
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


Private Sub PopulateScreenControls()
    On Error GoTo Failed
    
    'Populate Distribution Combo
    g_clsFormProcessing.PopulateChannel cboDistChannel
    
    'Poulate Condition Type Combo
    g_clsFormProcessing.PopulateCombo m_sComboGroup, cboConditionType
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenData()
On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetFormAddState()
On Error GoTo Failed
    
    chkDeleteFlag.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
'BMIDS00313
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim vSaveChanges As Variant
    Dim vSaveFreeFormat As Variant
    Dim bValidTextStructure As Boolean
    Dim bFreeFormat As Boolean
    'Dim bValidDataItems As Boolean
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    vSaveFreeFormat = vbYes
    'BMIDS00313
    'If the Condition is Non-FreeFormat, then set FreeFormat flag to false, else set it to true.
    If optFreeFormat(CHECKBOX_NO).Value = True Then
        bFreeFormat = False
    Else 'Free format, so the text structure will always be valid.
        bFreeFormat = True
    End If
    'If the FreeFormat flag is set then we validate the contents of the Description text box.
    If Not bFreeFormat Then
        bValidTextStructure = ValidateTextStructureOfDescription()
    Else
        bValidTextStructure = True
    End If
    
    If bValidTextStructure Then
        If m_bIsEdit And m_bScreenUpdated And bRet Then
            'Make sure that the user is aware that any embedded (<<...>>) values will be ignored (treated as normal text)
            '..if the Condition is a FreeFormat one.
            If bFreeFormat Then
                vSaveFreeFormat = MsgBox("The Description text will be saved as free format text, ie. Any data items you have included will NOT be resolved from database values. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
            End If
            If vSaveFreeFormat = vbYes Then
                vSaveChanges = MsgBox("You are about to update the selected condition. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
                 Select Case vSaveChanges
                    Case vbYes
                        If bRet Then
                            SaveScreenData
                            SaveChangeRequest
                        End If
                    Case vbNo
                        bRet = False
                End Select
            End If
        ElseIf Not m_bIsEdit Then
            If bFreeFormat Then
                vSaveFreeFormat = MsgBox("The Description text will be saved as free format text, ie. Any data items you have included will NOT be resolved from database values. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
            End If
            If vSaveFreeFormat = vbYes Then
                SaveScreenData
                SaveChangeRequest
                bRet = True
            End If
            
        End If
    Else
        bRet = False
    End If
    DoOKProcessing = bRet
        
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Private Sub SetFormEditState()

    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sDepartmentID As String
    Dim colValues As New Collection

    Set clsTableAccess = m_clsConditions
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()
    cmdAnother.Enabled = False

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.DisplayError "Additional Questions - no records to edit"
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub optConditional_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub optDetailsRequired_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub optEditable_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub optFreeFormat_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub txtDescription_Change()
    m_bScreenUpdated = True
End Sub

Private Sub txtName_Change()
    m_bScreenUpdated = True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub txtRuleReference_Change()
    m_bScreenUpdated = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsConditions
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsConditions
    End If
    
    'Condition Reference
    If m_bIsEdit Then
        m_clsConditions.SetConditionReferenceID txtConditionRef
    Else
        'This is a new record, so get a new Unique reference for the record
        m_sConditionReference = m_clsConditions.GetNextConditionRef
        m_clsConditions.SetConditionReferenceID m_sConditionReference
    End If
    
    'Condition Name
    m_clsConditions.SetConditionName txtName.Text
    
    'Condition Description
    m_clsConditions.SetConditionDescription txtDescription.Text
    
    'Condition Type
    g_clsFormProcessing.HandleComboExtra cboConditionType, vTmp, GET_CONTROL_VALUE
    m_clsConditions.SetConditionType CStr(vTmp)
    
    'Editable Indicator
    If optEditable(CHECKBOX_YES).Value = True Then
        m_clsConditions.SetConditionEditInd GetNumberFromBoolean(True)
    Else
        m_clsConditions.SetConditionEditInd GetNumberFromBoolean(False)
    End If

    'Set the FreeFormat indicator to True/False
    If optFreeFormat(CHECKBOX_YES).Value = True Then
        m_clsConditions.SetConditionFormatInd GetNumberFromBoolean(True)
    Else
        m_clsConditions.SetConditionFormatInd GetNumberFromBoolean(False)
    End If
    
    'Get ChannelID
    g_clsFormProcessing.HandleComboExtra cboDistChannel, vTmp, GET_CONTROL_VALUE
    'Set ChannelID
    m_clsConditions.SetConditionChannelID CStr(vTmp)
    
    'Set the Rule reference
    m_clsConditions.SetConditionRuleReference txtRuleReference.Text
    
    'Delete Flag
    If chkDeleteFlag.Value = 0 Then
        m_clsConditions.SetConditionDeleteFlag GetNumberFromBoolean(False)
    Else
        m_clsConditions.SetConditionDeleteFlag GetNumberFromBoolean(True)
    End If
    
    'Conditional Indicator
    If optConditional(CHECKBOX_YES).Value = True Then
        m_clsConditions.SetConditionalIndicator GetNumberFromBoolean(True)
    Else
        m_clsConditions.SetConditionalIndicator GetNumberFromBoolean(False)
    End If
    
    Set clsTableAccess = m_clsConditions
    clsTableAccess.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub PopulateScreenFields()
    On Error GoTo Failed
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    
    txtConditionRef.Text = m_clsConditions.GetConditionReference
    m_sConditionReference = txtConditionRef.Text
    'Condition Name
    txtName.Text = m_clsConditions.GetConditionName
    'Description
    txtDescription.Text = m_clsConditions.GetConditionDescription
    'Get the Condition Type, and select it from the combo
    vTmp = m_clsConditions.GetConditionType()
    g_clsFormProcessing.HandleComboExtra cboConditionType, vTmp, SET_CONTROL_VALUE
    
    'Is the Editable Flag Set?
    bRet = GetBooleanFromNumberAsBoolean(m_clsConditions.GetConditionEditFlag)
    If bRet Then
        optEditable(CHECKBOX_YES).Value = True
    Else
        optEditable(CHECKBOX_NO).Value = True
    End If
    
    'Is the FreeFormat Flag Set?
    bRet = GetBooleanFromNumberAsBoolean(m_clsConditions.GetConditionFreeFormatFlag)
    If bRet Then
        optFreeFormat(CHECKBOX_YES).Value = True
    Else
        optFreeFormat(CHECKBOX_NO).Value = True
    End If
    
    'Get the Distribution Channel, and select it from the combo
    vTmp = m_clsConditions.GetConditionChannelID
    g_clsFormProcessing.HandleComboExtra cboDistChannel, vTmp, SET_CONTROL_VALUE
    
    'Rule Reference
    txtRuleReference.Text = m_clsConditions.GetConditionRuleRef
    
    'Delete Flag
    bRet = GetBooleanFromNumberAsBoolean(m_clsConditions.GetConditionDeleteFlag)
    If bRet Then
        chkDeleteFlag.Value = 1
    Else
        chkDeleteFlag.Value = 0
    End If
    
    'Conditional Indicator
    bRet = GetBooleanFromNumberAsBoolean(m_clsConditions.GetConditionalIndicator)
    If bRet Then
        optConditional(CHECKBOX_YES).Value = True
    Else
        optConditional(CHECKBOX_NO).Value = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Common Function used to setup a promotion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sDesc = txtConditionRef.Text

    colMatchValues.Add m_sConditionReference
    Set clsTableAccess = m_clsConditions
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsConditions, sDesc
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateTextStructureOfDescription() As Boolean


    Dim blnResult As Boolean 'result of function
    Dim strParseString As String 'string in txtDescription
    Dim intLength As Integer 'length of string in txtDescription
    Dim intIndex As Integer 'loop counter
    Dim blnOpenExpected As Boolean
    Dim strErrorMsg As String
    Dim colDBRefs As Collection
    Dim intLeftCount As Integer
    Dim intRightCount As Integer
    Dim strChar As String
    Dim intBegin As Integer
    Dim intEnd As Integer
    Const strLeftBrace = "<<"
    Const strRightBrace = ">>"
    
    intLeftCount = 0
    intRightCount = 0
    blnResult = True
    strParseString = txtDescription.Text
    intLength = Len(strParseString)
    Set colDBRefs = New Collection

    intIndex = 1
    intLeftCount = 0
    intRightCount = 0
    blnOpenExpected = True
    intBegin = 1
    intEnd = 1
    
    strErrorMsg = ""
    While (blnResult And (intIndex <= intLength))
        strChar = Mid(strParseString, intIndex, 2)
        Debug.Print "#" & strChar & "#"
        If blnOpenExpected Then
            If strChar = strLeftBrace Then 'found a '<<' now expecting a '>>'
                blnOpenExpected = False
                intLeftCount = intLeftCount + 1
                intBegin = intIndex
                
            End If
            
            If strChar = strRightBrace Then
                strErrorMsg = "Expecting a '" & strLeftBrace & "', but found a '" & strRightBrace & "'."
                blnResult = False
            End If
            
        Else 'Close Expected
                If strChar = strRightBrace Then 'found a '>>' now expecting a '<<'
                    blnOpenExpected = True
                    intRightCount = intRightCount + 1
                    intEnd = intIndex
                    Debug.Print "BRACE CONTENT : " & Mid(strParseString, intBegin + 2, intEnd - intBegin - 2)
                    colDBRefs.Add Mid(strParseString, intBegin + 2, intEnd - intBegin - 2)
                End If
                
                If strChar = strLeftBrace Then
                    strErrorMsg = "Expecting a '" & strRightBrace & "', but found a '" & strLeftBrace & "'."
                    blnResult = False
                End If
        
            'End If
        End If

        intIndex = intIndex + 1
        
    Wend
    
    If intLeftCount <> intRightCount Then
        blnResult = False
        If Len(strErrorMsg) <> 0 Then
            strErrorMsg = "Unmatching number of '" & strLeftBrace & "' and '" & strRightBrace & "'." & strErrorMsg
        Else
            strErrorMsg = "Unmatching number of '" & strLeftBrace & "' and '" & strRightBrace & "'."
        End If
    End If
    If blnResult Then
        'Output Collection - testing
    
        For intIndex = 1 To colDBRefs.Count
            Debug.Print colDBRefs.Item(intIndex)
        Next
    Else
        Debug.Print strErrorMsg
        MsgBox "The contents of the Description box are not in the correct format. You will need to amend the text before you can save the data, Details : " & strErrorMsg, vbOKOnly
    End If
    
    
    Debug.Print "error is " & strErrorMsg
    ValidateTextStructureOfDescription = blnResult

End Function

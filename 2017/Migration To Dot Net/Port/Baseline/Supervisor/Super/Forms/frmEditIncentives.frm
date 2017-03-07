VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIncentives 
   Caption         =   "Add/Edit Incentives"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmEditIncentives.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboIncentiveBenefitType 
      Height          =   315
      Left            =   1740
      TabIndex        =   14
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSGOCX.MSGTextMulti txtDescription 
      Height          =   1515
      Left            =   1740
      TabIndex        =   4
      Top             =   2400
      Width           =   3795
      _ExtentX        =   6694
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
      MaxLength       =   50
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   4365
      TabIndex        =   7
      Top             =   4155
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtIncentives 
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   1
      Top             =   1050
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
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
   Begin MSGOCX.MSGComboBox cboType 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
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
      ListText        =   "Inclusive|Exclusive"
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3015
      TabIndex        =   6
      Top             =   4155
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1665
      TabIndex        =   5
      Top             =   4155
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtIncentives 
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Top             =   1485
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
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
   Begin MSGOCX.MSGEditBox txtIncentives 
      Height          =   315
      Index           =   3
      Left            =   1740
      TabIndex        =   3
      Top             =   1920
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
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
   Begin VB.Label Label5 
      Caption         =   "Benefit Type"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblMaintIncentives 
      Caption         =   "Set Type"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Percentage Max"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Percentage"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1140
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditIncentives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditIncentives
' Description   : To Add and Edit Incentives.
'
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' STB       06/12/01    SYS1942 - 'Another' button / transactions.
' SDS       29/01/02    SYS3320 -  Replacement of "Benefit Type" textbox with Combobox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Text control Constants
Private Const INCENTIVE_BENEFIT_TYPE = 0
Private Const AMOUNT_VAL = 1
Private Const PERCENTAGE_VAL = 2
Private Const MAXIMUM_VAL = 3

Private m_bIsEdit As Boolean

'Used to indicate if cmdAnother has been used, this will store the number of
'records added in one 'session'.
Private m_iRecordsAdded As Integer

' Private data
' Table classes
Private m_colIncentiveKey As Collection
Private m_clsIncentive As MortProdIncentiveTable
Private m_clsInclusiveExlcusive As MortProdIncIncentiveTable

' The type of incentive
Private m_sTypeText As String
Private m_ReturnCode As MSGReturnCode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Set the number of records modified to one.
    m_iRecordsAdded = 1
    
    ' Default return code
    SetReturnCode MSGFailure
    
    ' Create table classes
    Set m_clsIncentive = New MortProdIncentiveTable
    Set m_clsInclusiveExlcusive = New MortProdIncIncentiveTable
        
    g_clsFormProcessing.PopulateCombo "IncentiveBenefitType", cboIncentiveBenefitType, True
    
    If m_bIsEdit = False Then
        SetAddState
    Else
        SetEditState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Called when this form is unloaded. Free any objects used by this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsIncentive = Nothing
    Set m_clsInclusiveExlcusive = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Sets the status of the Incentives screen when adding an incentive.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    If m_clsInclusiveExlcusive Is Nothing Then
        Set m_clsInclusiveExlcusive = New MortProdIncIncentiveTable
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Sets the status of the Incentives form when editing an Incentive. Need to
'                 read the relevant records and set the relevant fields.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    On Error GoTo Failed
    
    TableAccess(m_clsIncentive).SetKeyMatchValues m_colIncentiveKey
    TableAccess(m_clsIncentive).GetTableData
    
    PopulateScreenFields
    
    cboType.Enabled = False
    cmdAnother.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populates all screen fields for the Incentives.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    
    Dim vVal As Variant

    If Len(m_sTypeText) > 0 Then
        g_clsFormProcessing.HandleComboText cboType, m_sTypeText, SET_CONTROL_VALUE
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "EditIncentives - Type is empty"
    End If
    
    g_clsFormProcessing.HandleComboExtra cboIncentiveBenefitType, m_clsIncentive.GetIncentiveBenefitType(), SET_CONTROL_VALUE
    
    txtDescription.Text = m_clsIncentive.GetDescription()
'    txtIncentives(INCENTIVE_BENEFIT_TYPE).Text = m_clsIncentive.GetIncentiveBenefitType()
    txtIncentives(AMOUNT_VAL).Text = m_clsIncentive.GetAmount()
    txtIncentives(PERCENTAGE_VAL).Text = m_clsIncentive.GetPercentage()
    txtIncentives(MAXIMUM_VAL).Text = m_clsIncentive.GetPercentageMax()

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveIncentives
' Description   : Creates and saves the Incentive records - the Incentive table and either the InclusiveIncentive
'                 table or the ExlusiveIncentive table. The IncentiveGUID will be created and written
'                 too.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveIncentives()
    On Error GoTo Failed
    Dim sType As String
    Dim vIncentiveGUID As Variant
    
    sType = cboType.SelText
    
    If Len(sType) > 0 Then
        Select Case sType
            Case "Exclusive"
                m_clsInclusiveExlcusive.SetType Exclusive
            Case "Inclusive"
                m_clsInclusiveExlcusive.SetType Inclusive
            Case Else
                g_clsErrorHandling.RaiseError errGeneralError, "SaveIncentives - Invalid Incentive"
        End Select
    
        g_clsFormProcessing.CreateNewRecord m_clsInclusiveExlcusive
        vIncentiveGUID = m_clsInclusiveExlcusive.SetIncentiveGUID()
        
        g_clsFormProcessing.CreateNewRecord m_clsIncentive
        m_clsIncentive.SetIncentiveGUID vIncentiveGUID
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SaveIncentives: Type is empty"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all screen fields
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim clsIncentives As MortProdIncentiveTable
    Dim vVal As Variant
    Set clsIncentives = m_clsIncentive
    
    ' If in Add mode, create the records.
    If m_bIsEdit = False Then
        SaveIncentives
    End If
    
    clsIncentives.SetAmount txtIncentives(AMOUNT_VAL).Text
    clsIncentives.SetDescription txtDescription.Text
    'clsIncentives.SetIncentiveBenefitType Me.txtIncentives(INCENTIVE_BENEFIT_TYPE).Text
'    clsIncentives.SetIncentiveBenefitType Me.txtIncentives(INCENTIVE_BENEFIT_TYPE).Text
    g_clsFormProcessing.HandleComboExtra cboIncentiveBenefitType, vVal, GET_CONTROL_VALUE
    clsIncentives.SetIncentiveBenefitType CStr(vVal)
    
    clsIncentives.SetPercentage txtIncentives(PERCENTAGE_VAL).Text
    clsIncentives.SetPercentageMax txtIncentives(MAXIMUM_VAL).Text



    DoUpdates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Called to update all tables used by this tab, once all data has been written
'                 to the tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsIncentive).Update
    
    If m_bIsEdit = False Then
        TableAccess(m_clsInclusiveExlcusive).Update
        TableAccess(m_clsInclusiveExlcusive).CloseRecordSet
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIncentive
' Description   : When editing, we will need to know the Incentive key. This method will be called
'                 externally to set the keys.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIncentive(colIncentive As Collection, sType As String)
    On Error GoTo Failed
    
    ' The type is Inclusive or Exclusive.
    m_sTypeText = sType
    Set m_colIncentiveKey = colIncentive
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user presses the OK button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing()

    If bRet = True Then
        Hide
        SetReturnCode
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Common method called when the user presses OK or Another. Validate screen data,
'                 then Save it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        ValidateScreenData
        SaveScreenData
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    bRet = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called to validate all data entered on the screen. Returns true if ok, false if
'                 not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateScreenData()
    On Error GoTo Failed
    Dim nCount As Integer

    nCount = 0

    If Len(txtIncentives(PERCENTAGE_VAL).Text) > 0 Then
        nCount = nCount + 1
    End If

    If Len(txtIncentives(AMOUNT_VAL).Text) > 0 Then
        nCount = nCount + 1
    End If

    If nCount > 1 Or nCount = 0 Then
        g_clsFormProcessing.SetControlFocus txtIncentives(AMOUNT_VAL)
        g_clsErrorHandling.RaiseError errGeneralError, "One of Amount and Percentage must be entered"
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Called when the user clicks the Another button. Need to perform all validation,
'                 save screen data and clear the screen ready for the next record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    On Error GoTo Failed
    
    Dim bRet As Boolean
    
    'Trip this flag to true.
    m_iRecordsAdded = m_iRecordsAdded + 1
    
    bRet = DoOKProcessing

    If bRet Then
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    
        If bRet = True Then
            g_clsFormProcessing.SetControlFocus cboType
        End If
    
        ' Last thing to do is to clear the current type
        m_sTypeText = ""
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIncentives
' Description   : Returns the Incentive table class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIncentives() As MortProdIncentiveTable
    On Error GoTo Failed
        
    Set GetIncentives = m_clsIncentive
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub txtIncentives_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtIncentives(Index).ValidateData()
End Sub
Private Sub cboType_Validate(Cancel As Boolean)
    Cancel = Not cboType.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : If more than one record has been added, only proceeds if the user allows it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()

    Dim uResponse As VbMsgBoxResult
    
    On Error GoTo Failed

    'If the another button has been used, we should warn that more than one record
    'is about to be lost. Normally this would only rollback the last record added,
    'but this form is a popup which runs in a wider transaction.
    If m_iRecordsAdded > 1 Then
        uResponse = MsgBox("You have added more than one record. Proceeding will cancel all records added. Do you wish to proceed?", vbYesNo Or vbExclamation, "Loose Changes?")
    End If
    
    'If editing, only one record has been added or the user is willing to loose
    'multiple changes, then allow the cancel to proceed,
    If (m_iRecordsAdded < 2) Or (uResponse = vbYes) Then
        Hide
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


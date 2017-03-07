VERSION 5.00
Object = "{DECD0E42-82AF-11D5-8328-000102A316E5}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPackagingProcFee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Packaging Fee Details"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmEditPackagingProcFee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFees 
      Caption         =   "Fee Details"
      Height          =   1395
      Left            =   120
      TabIndex        =   18
      Top             =   1140
      Width           =   5895
      Begin VB.OptionButton OptPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   435
         Left            =   4440
         TabIndex        =   13
         Top             =   660
         Width           =   555
      End
      Begin VB.OptionButton OptPounds 
         Alignment       =   1  'Right Justify
         Caption         =   "£"
         Height          =   435
         Left            =   3840
         TabIndex        =   6
         Top             =   660
         Value           =   -1  'True
         Width           =   495
      End
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   3
         Left            =   1260
         TabIndex        =   3
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         TextType        =   2
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
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   4
         Left            =   1260
         TabIndex        =   5
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         TextType        =   2
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
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   5
         Left            =   4140
         TabIndex        =   4
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
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
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Fee"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Upper Limit"
         Height          =   255
         Index           =   4
         Left            =   3180
         TabIndex        =   20
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Lower Limit"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Fee Split"
      Height          =   1275
      Left            =   120
      TabIndex        =   14
      Top             =   2700
      Width           =   5895
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   6
         Left            =   1380
         TabIndex        =   7
         Top             =   300
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
      End
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   7
         Left            =   1380
         TabIndex        =   9
         Top             =   720
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
      End
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   8
         Left            =   4080
         TabIndex        =   8
         Top             =   300
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
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Individual"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Company"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Lead Agent"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   4260
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1470
      TabIndex        =   11
      Top             =   4260
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4260
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
   End
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
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
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   2
      Left            =   4740
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
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
   Begin VB.Label lblProdProcDetails 
      Caption         =   "Packaging Product"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   24
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblProdProcDetails 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label lblProdProcDetails 
      Caption         =   "End Date"
      Height          =   255
      Index           =   2
      Left            =   3300
      TabIndex        =   22
      Top             =   660
      Width           =   1395
   End
End
Attribute VB_Name = "frmEditPackagingProcFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form : frmEditPackagingProcFee
' Description   : Form which allows the user add and edit Intermediary Proc Fee details
'
' Change history
' Prog      Date        Description
' AA        22/06/2001  added class
' DJP       27/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Screen Constants
Private Const START_DATE = 1
Private Const END_DATE = 2
Private Const PRODUCT_TYPE = 0
Private Const LOWER_LIMIT = 3
Private Const UPPER_LIMIT = 5
Private Const PROC_FEE = 4
Private Const TXT_LEAD_AGENT = 6
Private Const TXT_COMPANY = 7
Private Const TXT_INDIVIDUAL = 8

Private m_sSplitSequence As String
Private m_nFeeType As Integer
Private m_sProduct As String
Private m_bIsEdit As Boolean
Private m_clsProcFeeForInt As IntProcFeeSplitForIntTable
Private m_clsProcFee As IntProcFeeSplitTable
Private m_vIntermediaryGUID As Variant

Public Sub SetTableClass(clsProcFee As IntProcFeeSplitTable)
    Set m_clsProcFee = clsProcFee
    m_vIntermediaryGUID = clsProcFee.GetIntermediaryGUID
End Sub
Public Sub SetIsEdit(bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOkProcessing
    
    If bRet Then
        g_clsFormProcessing.ClearScreenFields Me
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOkProcessing
    
    If bRet Then
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
        
    Set m_clsProcFeeForInt = New IntProcFeeSplitForIntTable
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub SetAddState()
    On Error GoTo Failed

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetEditState()
    On Error GoTo Failed
    
        Dim sGuid As String
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    m_clsProcFee.GetAllFeeSplitsForGuid sGuid, CStr(m_nFeeType), m_sSplitSequence
    
    If TableAccess(m_clsProcFee).RecordCount > 0 Then
        SetScreenFields
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : The active from/to and product details must be populated from the parent form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    On Error GoTo Failed

    Dim clsProcFeeType As IntermediaryProcFeeTable
    Dim sIntermediaryGUID As String
    Dim sType As String
    Dim sSequence As String
    
    Set clsProcFeeType = New IntermediaryProcFeeTable
    
    'Get Values from parent form
    sType = m_clsProcFee.GetFeeType
    sSequence = m_clsProcFee.GetTypeSequenceNumber
    ' DJP SQL Server port
    sIntermediaryGUID = g_clsSQLAssistSP.GuidToString(m_clsProcFee.GetIntermediaryGUID)
    
    'Find the data relating to current record
    clsProcFeeType.GetProcFee sIntermediaryGUID, sType, sSequence

    txtProdProcFees(START_DATE).Text = clsProcFeeType.GetFeeActiveFrom
    txtProdProcFees(END_DATE).Text = clsProcFeeType.GetFeeActiveTo
    txtProdProcFees(PRODUCT_TYPE).Text = m_sProduct
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductName
' Description   : Sets the name of the product to be edited
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductName(sName As String)
    On Error GoTo Failed

    m_sProduct = sName

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Sets the proc fee type to that of the valueid from the combogroup
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeType(nType As Long)
    On Error GoTo Failed
    m_nFeeType = nType
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves the screen data to the table Class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    
    
    If Not m_bIsEdit Then
        m_clsProcFee.SetIntermediaryGUID m_vIntermediaryGUID
        m_sSplitSequence = m_clsProcFee.GetNextSplitSequenceNumber
        m_clsProcFee.SetSplitSequenceNumber m_sSplitSequence
    End If
    
    
    m_clsProcFee.SetLowBand txtProdProcFees(LOWER_LIMIT).Text
    m_clsProcFee.SetHighBand txtProdProcFees(UPPER_LIMIT).Text
    m_clsProcFee.SetTotalAmount txtProdProcFees(PROC_FEE).Text
    
    m_clsProcFee.SetFeeType m_nFeeType

    SaveSplitDetails INTERMEDIARIES_LEADAGENT, txtProdProcFees(TXT_LEAD_AGENT)
    SaveSplitDetails INTERMEDIARIES_COMPANY, txtProdProcFees(TXT_COMPANY)
    SaveSplitDetails INTERMEDIARIES_INDIVIDUAL, txtProdProcFees(TXT_INDIVIDUAL)
    
    DoUpdates
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveSplitDetails
' Description   : Saves the intermediay splits for the proc fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveSplitDetails(sType As String, txt As MSGEditBox)
    On Error GoTo Failed

    Dim nValueID As Integer
    Dim clsIntermediary As Intermediary
    Dim sGuid As String
    Set clsIntermediary = New Intermediary

    If Len(sType) > 0 Then
        nValueID = clsIntermediary.GetValueIDForIntType(sType)
        If Not m_bIsEdit Then
            g_clsFormProcessing.CreateNewRecord m_clsProcFeeForInt
            m_clsProcFeeForInt.SetSplitSequenceNumber m_sSplitSequence
            m_clsProcFeeForInt.SetTypeSequenceNumber m_clsProcFee.GetTypeSequenceNumber
            m_clsProcFeeForInt.SetIntermediaryGUID m_vIntermediaryGUID
            m_clsProcFeeForInt.SetFeeType m_nFeeType
        Else
            sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
            m_clsProcFeeForInt.GetSplitDetails sGuid, CStr(nValueID), m_sSplitSequence
        End If
        
        m_clsProcFeeForInt.SetIntermediaryType nValueID
    
        m_clsProcFeeForInt.SetPercentage txt.Text
        If m_bIsEdit Then
            TableAccess(m_clsProcFeeForInt).Update
        End If
    End If
      
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Called onOK_Click
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData
    
    If bRet Then
        SaveScreenData
            
    End If
    
    DoOkProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates the screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
        bRet = ValidateLimits
    End If
    If bRet Then
        bRet = ValidatePercentage
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates the recordset data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsProcFee).Update
    TableAccess(m_clsProcFeeForInt).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateLimits
' Description   : Checks the upper lower limit is valid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateLimits() As Boolean
    On Error GoTo Failed

    Dim nLower As Long
    Dim nUpper As Long
    Dim nFee As Long
    Dim bRet As Boolean
    
    nLower = CLng(txtProdProcFees(LOWER_LIMIT).Text)
    nUpper = CLng(txtProdProcFees(UPPER_LIMIT).Text)
    nFee = CLng(txtProdProcFees(PROC_FEE).Text)

    bRet = True
    If nLower > nUpper Then
        g_clsErrorHandling.DisplayError "Upper Limit must be greater than the Lower Limit"
        txtProdProcFees(UPPER_LIMIT).SetFocus
        bRet = False
    End If
    
    If bRet Then
        If nFee > nUpper Then
            g_clsErrorHandling.DisplayError "Procuration Fee must be less than or equal to the Upper Limit"
            txtProdProcFees(PROC_FEE).SetFocus
            bRet = False
        End If
    End If
    
    If bRet Then
        If nFee < nLower Then
            g_clsErrorHandling.DisplayError "Procuration Fee must be greater than the Lower Limit"
            txtProdProcFees(PROC_FEE).SetFocus
            bRet = False
        End If
    End If

    ValidateLimits = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSplitSequence
' Description   : If an edit the sequence number must be set
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSplitSequence(sSequence As String)
    On Error GoTo Failed
    
    m_sSplitSequence = sSequence
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates the screen in edit mode with the values from the DB
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
'
    Dim clsProcFee As IntermediaryProcFeeTable
    Dim sGuid As String
    Dim sTypeSequence As String
    Dim sProduct As String
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    Set clsProcFee = New IntermediaryProcFeeTable
    
    'Fee Dates
    clsProcFee.GetFeeTypesForIntermediary sGuid, CStr(m_nFeeType), sTypeSequence
    sTypeSequence = clsProcFee.GetTypeSequence
    m_sSplitSequence = m_clsProcFee.GetSplitSequenceNumber
    
    txtProdProcFees(START_DATE).Text = clsProcFee.GetFeeActiveFrom
    txtProdProcFees(END_DATE).Text = clsProcFee.GetFeeActiveTo

    'Fee split details
    txtProdProcFees(UPPER_LIMIT).Text = m_clsProcFee.GetHighBand
    txtProdProcFees(LOWER_LIMIT).Text = m_clsProcFee.GetLowBand
    txtProdProcFees(PROC_FEE).Text = m_clsProcFee.GetTotalAmount
    
    SetFeeSplits INTERMEDIARIES_LEADAGENT, txtProdProcFees(TXT_LEAD_AGENT)
    SetFeeSplits INTERMEDIARIES_INDIVIDUAL, txtProdProcFees(TXT_INDIVIDUAL)
    SetFeeSplits INTERMEDIARIES_COMPANY, txtProdProcFees(TXT_COMPANY)

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : validatePercentage
' Description   : Checks that all the percentages = 100
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidatePercentage() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nLeadAgent As Integer
    Dim nCompany As Integer
    Dim nIndividual As Integer
    Dim nTotalPercent As Integer
    
    bRet = True
    If Len(txtProdProcFees(TXT_LEAD_AGENT).Text) > 0 Then
        nLeadAgent = txtProdProcFees(TXT_LEAD_AGENT).Text
    End If
    
    If Len(txtProdProcFees(TXT_COMPANY).Text) > 0 Then
        nCompany = txtProdProcFees(TXT_COMPANY).Text
    End If
    
    If Len(txtProdProcFees(TXT_INDIVIDUAL).Text) > 0 Then
        nIndividual = txtProdProcFees(TXT_INDIVIDUAL).Text
    End If
    
    
    nTotalPercent = nCompany + nLeadAgent + nIndividual
    
    If nTotalPercent <> 100 Then
        g_clsErrorHandling.DisplayError "The Fee split percentage must equal 100"
        txtProdProcFees(TXT_LEAD_AGENT).SetFocus
        bRet = False
    End If
    
    ValidatePercentage = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetIntermediaryGUID(vGuid As Variant)
    m_vIntermediaryGUID = vGuid
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeSplits
' Description   : Populates the split details
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFeeSplits(sType As String, txt As MSGEditBox)
    On Error GoTo Failed

    Dim sGuid As String
    Dim clsSplit As IntProcFeeSplitForIntTable
    
    Set clsSplit = New IntProcFeeSplitForIntTable
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    clsSplit.GetSplitAmounts sGuid, sType, m_sSplitSequence
    
    If TableAccess(clsSplit).RecordCount > 0 Then
        txt.Text = clsSplit.GetSplitPercentage
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

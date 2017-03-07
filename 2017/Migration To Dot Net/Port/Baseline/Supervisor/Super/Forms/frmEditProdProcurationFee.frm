VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditProductProcFee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Product Procuration Fee Details"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmEditProdProcurationFee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   4260
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1530
      TabIndex        =   11
      Top             =   4260
      Width           =   1230
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4860
      TabIndex        =   12
      Top             =   4260
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VB.Frame fraSplit 
      Caption         =   "Fee Split"
      Height          =   1275
      Left            =   180
      TabIndex        =   15
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
         MaxLength       =   10
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
         MaxLength       =   10
      End
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   8
         Left            =   3960
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
         MaxLength       =   10
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Individual"
         Height          =   255
         Index           =   8
         Left            =   2940
         TabIndex        =   24
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Company"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label lblProdProcDetails 
         Caption         =   "Lead Agent"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame fraFees 
      Caption         =   "Fee Details"
      Height          =   1395
      Left            =   180
      TabIndex        =   14
      Top             =   1140
      Width           =   5895
      Begin VB.OptionButton OptPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   435
         Left            =   4440
         TabIndex        =   13
         Top             =   720
         Width           =   555
      End
      Begin VB.OptionButton OptPounds 
         Alignment       =   1  'Right Justify
         Caption         =   "£"
         Height          =   435
         Left            =   3780
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   555
      End
      Begin MSGOCX.MSGEditBox txtProdProcFees 
         Height          =   315
         Index           =   3
         Left            =   1260
         TabIndex        =   0
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
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   1
      Left            =   1860
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
   Begin MSGOCX.MSGEditBox txtProdProcFees 
      Height          =   315
      Index           =   2
      Left            =   4860
      TabIndex        =   3
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
      Caption         =   "End Date"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   18
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label lblProdProcDetails 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label lblProdProcDetails 
      Caption         =   "Mortgage Product"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmEditProductProcFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditProductProcFee
' Description   : Form which allows the user add and edit Intermediary Proc Fee
'                 details for any fee type (i.e. generic). The controls which
'                 aren't relevant for the specified fee-type will be disabled.
'
' Change history
' Prog      Date        Description
' AA        22/06/2001  added form
' DJP       27/06/01    SQL Server port, replace SQLAssist with SQLAssistSP
' STB       17/12/01    SYS2550 - SQL Server support.
' STB       22/04/02    SYS4398 - Fee should not be validated against upper or
'                       lower limit bands.
' STB       22/04/02    SYS4400 - Fee percent and amount option controls
'                       enabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Control Indexes.
Private Const START_DATE     As Long = 1
Private Const END_DATE       As Long = 2
Private Const PRODUCT_TYPE   As Long = 0
Private Const LOWER_LIMIT    As Long = 3
Private Const UPPER_LIMIT    As Long = 5
Private Const PROC_FEE       As Long = 4
Private Const TXT_LEAD_AGENT As Long = 6
Private Const TXT_COMPANY    As Long = 7
Private Const TXT_INDIVIDUAL As Long = 8

'The key-value collection for this record.
Private m_colKeys As Collection

'The ID of the associated mortgage product. Used to populate the screen caption.
Private m_sProductID As String

'The procuration fee-type this band/split will be used for.
Private m_uFeeType As ProcFeeTypeEnum

'The parent period's sequence number.
Private m_sTypeSequence As String

'The split sequence for this record (generated if adding).
Private m_sSplitSequence As String

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'These are the underlying table objects required by this form. They are created
'and deleted by the main intermediaries form but the data is controlled by this
'form and the procuration fees form.
Private m_clsProcFeeSplitTable As IntProcFeeSplitTable
Private m_clsProcFeeTypeTable As IntermediaryProcFeeTable
Private m_clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable

'Note: The pounds/percent option buttons have been hidden as they served no purpose.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplitTable
' Description   : Associcate the Fee Split table with this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplitTable(ByRef clsProcFeeSplit As IntProcFeeSplitTable)
    Set m_clsProcFeeSplitTable = clsProcFeeSplit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Mimick the OK button and then prepare the form to add a new record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Validate and save the current record.
    bRet = DoOkProcessing
    
    If bRet Then
        'Clear the fields and reset state for a new record.
        g_clsFormProcessing.ClearScreenFields Me
        
        'TODO: SYS1942 needs to be catored for here. Until then, the 'Another'
        'button is hidden.
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the opener.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    'TODO: SYS1942 needs to be catored for here.
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the current record. Close the form and return contro to
'                 the opener if the record is valid and saved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Validate and save the record.
    bRet = DoOkProcessing
    
    'Close the form if the record saved okay.
    If bRet Then
        Hide
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Setup the form according to the add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Setup common table state.
    SetCommonState
    
    'Setup the tables appropriately.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Populate static data and combo lists.
    PopulateScreenControls

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup tables with new records and GUIDs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    
    'Note: This form is a bit messy because the 3 split records are created on-
    'save which does not conform to the standard.
    
    'Create a fee-band record.
    g_clsFormProcessing.CreateNewRecord TableAccess(m_clsProcFeeSplitTable)
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup tables and load data. Do this by filtering all other records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    Dim sFilter As String
    
    'Filter the band table (GUID, Type, TypeSequence, SplitSequnce).
    sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(m_uFeeType)
    sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sTypeSequence)
    sFilter = sFilter & " AND SplitSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sSplitSequence)
    
    'Apply the filter.
    TableAccess(m_clsProcFeeSplitTable).ApplyFilter sFilter
    
    'Ensure we have a single record.
    If TableAccess(m_clsProcFeeSplitTable).RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Incorrect number of band records returned (" & TableAccess(m_clsProcFeeSplitTable).RecordCount & " records returned, 1 was expected)"
    End If
    
    'Populate the screen from the underlying data.
    SetScreenFields
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetCommonState
' Description   : Setup some table data which is required for both Add and Edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCommonState()

    Dim sFilter As String
    
    'Filter the period table (GUID, Type, TypeSequence).
    sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(m_uFeeType)
    sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sTypeSequence)
    
    'Apply the filter.
    TableAccess(m_clsProcFeeTypeTable).ApplyFilter sFilter
    
    'Ensure we have a single record.
    If TableAccess(m_clsProcFeeTypeTable).RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Incorrect number of period records returned (" & TableAccess(m_clsProcFeeTypeTable).RecordCount & " records returned, 1 was expected)"
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate static data and combo lists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
            
    Dim clsProductTable As MortgageProductTable
    
    'Alter the form's caption to match the type of fee.
    Select Case m_uFeeType
        Case NonSpecificFee, MortgageFee
            Me.Caption = "Add/Edit Product Procuration Fee Details"
        
        Case InsuranceFee
            Me.Caption = "Add/Edit Insurance Product Fee Details"
            
        Case PackagingFee
            Me.Caption = "Add/Edit Packaging Fee Details"
    End Select
    
    'If the fee type is a mortgage fee. Find out the mortgage product's name.
    If m_uFeeType = MortgageFee Then
        'Create a mortgage product table.
        Set clsProductTable = New MortgageProductTable
    
        'Load the specific product record.
        clsProductTable.GetProductByID m_sProductID
        
        'Populate the mortgage name.
        txtProdProcFees(PRODUCT_TYPE).Text = clsProductTable.GetProductName()
    End If
    
    'Disable the product label if this isn't a mortgage product.
    lblProdProcDetails(0).Enabled = (m_uFeeType = MortgageFee)
    
    'Start/from date.
    txtProdProcFees(START_DATE).Text = m_clsProcFeeTypeTable.GetFeeActiveFrom
    
    'End/to date.
    txtProdProcFees(END_DATE).Text = m_clsProcFeeTypeTable.GetFeeActiveTo
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductID
' Description   : Sets the ID of the product related to this period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductID(ByVal sProductID As String)
    m_sProductID = sProductID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Sets the proc fee type to that of the valueid from the combogroup.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeType(ByVal uFeeType As ProcFeeTypeEnum)
    m_uFeeType = uFeeType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves the screen data to the table Class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
        
    On Error GoTo Failed
    
    'Set the primary key fields if we're adding.
    If Not m_bIsEdit Then
        'Intermediary GUID.
        m_clsProcFeeSplitTable.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
        
        'Type Sequence number.
        m_clsProcFeeSplitTable.SetTypeSequenceNumber m_sTypeSequence
        
        'Generate a split sequence number.
        m_sSplitSequence = m_clsProcFeeSplitTable.GetNextSplitSequenceNumber
        
        'SplitSequence.
        m_clsProcFeeSplitTable.SetSplitSequenceNumber m_sSplitSequence
    End If
        
    'Low band.
    m_clsProcFeeSplitTable.SetLowBand txtProdProcFees(LOWER_LIMIT).Text
    
    'High band.
    m_clsProcFeeSplitTable.SetHighBand txtProdProcFees(UPPER_LIMIT).Text
    
    'Either save the fee as a percentage or an amount.
    If OptPercent.Value = True Then
        m_clsProcFeeSplitTable.SetTotalAmount ""
        m_clsProcFeeSplitTable.SetTotalPercent txtProdProcFees(PROC_FEE).Text
    Else
        m_clsProcFeeSplitTable.SetTotalAmount txtProdProcFees(PROC_FEE).Text
        m_clsProcFeeSplitTable.SetTotalPercent ""
    End If
    
    'Procuration Fee Type.
    m_clsProcFeeSplitTable.SetFeeType m_uFeeType
        
    'Save each split record.
    SaveSplitDetails LeadAgentType, txtProdProcFees(TXT_LEAD_AGENT).Text
    SaveSplitDetails CompanyType, txtProdProcFees(TXT_COMPANY).Text
    SaveSplitDetails IndividualType, txtProdProcFees(TXT_INDIVIDUAL).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveSplitDetails
' Description   : Saves the split fee information for the specified intermediary type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveSplitDetails(ByVal uIntermediaryType As IntermediaryTypeEnum, ByVal sPercentage As String)
    
    Dim sFilter As String
        
    If m_bIsEdit Then
        'If we're editing a split, then filter the split table to the existing
        'record (filter on GUID, Type, TypeSequence, SplitSequnce and, IntType).
        sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
        sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(m_uFeeType)
        sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sTypeSequence)
        sFilter = sFilter & " AND SplitSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sSplitSequence)
        sFilter = sFilter & " AND SplitForIntermediaryType = " & g_clsSQLAssistSP.FormatString(uIntermediaryType)
        
        'Apply the filter on the table object.
        TableAccess(m_clsProcFeeSplitByIntTable).ApplyFilter sFilter
        
        'Ensure only one record is available.
        If TableAccess(m_clsProcFeeSplitByIntTable).RecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, "Incorrect number of split records returned (" & TableAccess(m_clsProcFeeSplitByIntTable).RecordCount & " records returned, 1 was expected)"
        End If
    Else
        'If we're adding a split, create the record.
        g_clsFormProcessing.CreateNewRecord m_clsProcFeeSplitByIntTable
        
        'Set all the primary key fields.
        m_clsProcFeeSplitByIntTable.SetSplitSequenceNumber m_sSplitSequence
        m_clsProcFeeSplitByIntTable.SetTypeSequenceNumber m_sTypeSequence
        m_clsProcFeeSplitByIntTable.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
        m_clsProcFeeSplitByIntTable.SetFeeType m_uFeeType
    End If
    
    'The Intermediary Type.
    m_clsProcFeeSplitByIntTable.SetIntermediaryType uIntermediaryType
    
    'The split percentage.
    m_clsProcFeeSplitByIntTable.SetPercentage sPercentage

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Validate and save the current record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Validate the screen data.
    bRet = ValidateScreenData
    
    If bRet Then
        'If valid, attempt to save.
        SaveScreenData
    End If
    
    DoOkProcessing = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates the screen and returns true if valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Ensure all mandatory information has been specified.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
        'Validate the upper/lower bands.
        bRet = ValidateLimits
    End If
    
    If bRet Then
        'Ensure all splits total 100%.
        bRet = ValidatePercentage
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateLimits
' Description   : Checks the upper lower limit is valid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateLimits() As Boolean

    Dim nFee As Long
    Dim nLower As Long
    Dim nUpper As Long
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    nLower = CLng(txtProdProcFees(LOWER_LIMIT).Text)
    nUpper = CLng(txtProdProcFees(UPPER_LIMIT).Text)
    nFee = CLng(txtProdProcFees(PROC_FEE).Text)

    bRet = True
    
    If nLower > nUpper Then
        g_clsErrorHandling.DisplayError "Upper Limit must be greater than the Lower Limit"
        txtProdProcFees(UPPER_LIMIT).SetFocus
        bRet = False
    End If
    
    'SYS4398 - Fee should not be validated against the upper or lower loan amount bands.
'    If bRet Then
'        If nFee > nUpper Then
'            g_clsErrorHandling.DisplayError "Procuration Fee must be less than or equal to the Upper Limit"
'            txtProdProcFees(PROC_FEE).SetFocus
'            bRet = False
'        End If
'    End If
'
'    If bRet Then
'        If nFee < nLower Then
'            g_clsErrorHandling.DisplayError "Procuration Fee must be greater than the Lower Limit"
'            txtProdProcFees(PROC_FEE).SetFocus
'            bRet = False
'        End If
'    End If
    'SYS4398 - End.

    ValidateLimits = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : validatePercentage
' Description   : Checks that all the percentages = 100
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidatePercentage() As Boolean
    
    Dim bRet As Boolean
    Dim nLeadAgent As Integer
    Dim nCompany As Integer
    Dim nIndividual As Integer
    Dim nTotalPercent As Integer
    
    On Error GoTo Failed
    
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates the screen fields with the data from the DB
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    
    'Upper limit.
    txtProdProcFees(UPPER_LIMIT).Text = m_clsProcFeeSplitTable.GetHighBand
    
    'Lower limit.
    txtProdProcFees(LOWER_LIMIT).Text = m_clsProcFeeSplitTable.GetLowBand
    
    'Either get the total amount or total percentage.
    If m_clsProcFeeSplitTable.GetTotalAmount = "" Then
        txtProdProcFees(PROC_FEE).Text = m_clsProcFeeSplitTable.GetTotalPercent
        OptPercent.Value = True
    Else
        txtProdProcFees(PROC_FEE).Text = m_clsProcFeeSplitTable.GetTotalAmount
        OptPounds.Value = True
    End If
    
    'Set/populate each intermediary split record.
    SetFeeSplits LeadAgentType, txtProdProcFees(TXT_LEAD_AGENT)
    SetFeeSplits CompanyType, txtProdProcFees(TXT_COMPANY)
    SetFeeSplits IndividualType, txtProdProcFees(TXT_INDIVIDUAL)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeSplits
' Description   : Populates the split details.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFeeSplits(ByVal uIntermediaryType As IntermediaryTypeEnum, ByRef txtSplit As MSGEditBox)

    Dim sFilter As String
    
    'Filter on GUID, Type, TypeSequence, SplitSequnce and, IntType.
    sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(m_uFeeType)
    sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sTypeSequence)
    sFilter = sFilter & " AND SplitSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sSplitSequence)
    sFilter = sFilter & " AND SplitForIntermediaryType = " & g_clsSQLAssistSP.FormatString(uIntermediaryType)

    'Apply the filter to the Split table.
    TableAccess(m_clsProcFeeSplitByIntTable).ApplyFilter sFilter
    
    'Ensure there is only one record loaded.
    If TableAccess(m_clsProcFeeSplitByIntTable).RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Incorrect number of split records returned (" & TableAccess(m_clsProcFeeSplitByIntTable).RecordCount & " records returned, 1 was expected)"
    End If

    'Split percentage.
    txtSplit.Text = m_clsProcFeeSplitByIntTable.GetSplitPercentage

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSplitSequence
' Description   : Store the record's sequence number (if we're editing).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSplitSequence(ByVal sSequence As String)
    m_sSplitSequence = sSequence
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Tidy-up and release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    
    Set m_colKeys = Nothing
    Set m_clsProcFeeTypeTable = Nothing
    Set m_clsProcFeeSplitTable = Nothing
    Set m_clsProcFeeSplitByIntTable = Nothing
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeType
' Description   : Set the procuration fee-type for this fee-band/split.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFeeType(ByVal uFeeType As ProcFeeTypeEnum)
    m_uFeeType = uFeeType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTypeSequence
' Description   : Set the period sequence number for this fee-band/split.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTypeSequence(ByVal sTypeSequence As String)
    m_sTypeSequence = sTypeSequence
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeTypeTable
' Description   : Associate the Proc Fee Type (period) table with this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeTypeTable(ByRef clsProcFeeTypeTable As IntermediaryProcFeeTable)
    Set m_clsProcFeeTypeTable = clsProcFeeTypeTable
End Sub
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplitByIntTable
' Description   : Associate the Split By Int (split) table with this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplitByIntTable(ByRef clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable)
    Set m_clsProcFeeSplitByIntTable = clsProcFeeSplitByIntTable
End Sub
 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Associate the collection of primary key values with this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

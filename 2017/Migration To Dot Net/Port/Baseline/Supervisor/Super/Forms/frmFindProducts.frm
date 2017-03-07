VERSION 5.00
Object = "{C23ED101-AD8A-42B6-82B4-D802F130BEF9}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindProducts 
   Caption         =   "Find Products"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "frmFindProducts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Lender"
      Height          =   1035
      Left            =   180
      TabIndex        =   16
      Top             =   180
      Width           =   4515
      Begin VB.OptionButton optLender 
         Caption         =   "All Lenders"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   660
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optLender 
         Caption         =   "Search by Lender"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   1755
      End
      Begin MSGOCX.MSGDataCombo cboLender 
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
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
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other search criteria"
      Height          =   2685
      Left            =   180
      TabIndex        =   15
      Top             =   1320
      Width           =   4515
      Begin VB.TextBox txtProductCode 
         Height          =   345
         Left            =   2490
         TabIndex        =   9
         Top             =   2130
         Width           =   1710
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "By End Date"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1500
         Width           =   1515
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "All Products"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1755
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "Active Products"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "Withdrawn Products"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1140
         Width           =   2235
      End
      Begin MSGOCX.MSGEditBox txtEndDate 
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSGOCX.MSGEditBox txtName 
         Height          =   315
         Left            =   315
         TabIndex        =   8
         Top             =   2130
         Width           =   1785
         _ExtentX        =   3149
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "or"
         Height          =   195
         Index           =   2
         Left            =   2205
         TabIndex        =   14
         Top             =   2190
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Product Code"
         Height          =   195
         Index           =   1
         Left            =   2475
         TabIndex        =   13
         Top             =   1905
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   12
         Top             =   1905
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmFindProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindProducts
' Description   : Allows the user to narrow the search of mortgage products.
'
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
' BS       25/03/2003  BM0282 Move listview lbltitle.Caption refresh
' MC       24/06/2004  BMIDS783 GetProducts() METHOD extended to give ProductCode for additional
'                      Search criteria.
' MC       13/07/2004  Product Search Enhancements for tabbing order.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const SEARCH_LENDER = 0
Private Const SEARCH_ALL_LENDERS = 1

Private Const SEARCH_ACTIVE_PRODUCTS = 0
Private Const SEARCH_WITHDRAWN_PRODUCTS = 1
Private Const SEARCH_END_DATE = 2

Private m_ReturnCode As MSGReturnCode
Private m_clsMortgageProduct As MortgageProductTable
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsMortgageProduct = clsTableAccess
End Sub
Private Sub cboLender_Validate(Cancel As Boolean)
    Cancel = Not cboLender.ValidateData()
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    Dim enumSearch As SearchType
    Dim sEndDate As String
    enumSearch = SearchNone
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
        If optSearchType(SEARCH_ACTIVE_PRODUCTS).Value = True Then
            enumSearch = SearchActive
        
        ElseIf optSearchType(SEARCH_WITHDRAWN_PRODUCTS).Value = True Then
            enumSearch = SearchWithdrawn
        
        ElseIf optSearchType(SEARCH_LENDER).Value = True Then
            DoLenderSearch
        ElseIf optSearchType(SEARCH_END_DATE).Value = True Then
            sEndDate = txtEndDate.Text
        End If
        '*=[MC]Search Criteria extended "Product Code" option added
        m_clsMortgageProduct.GetProducts enumSearch, cboLender.SelText, sEndDate, Me.txtName.Text, Me.txtProductCode.Text
        'section end
        Set clsTableAccess = m_clsMortgageProduct
        clsTableAccess.SetUpdated
        Hide
                
        'BS BM0282 25/03/03
        'Reset title as listview is being refreshed
        frmMain.lblTitle(constListViewLabel).Caption = frmMain.tvwDB.SelectedItem()
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
Private Sub DoLenderSearch()

End Sub
Private Sub DoWithdrawnSearch()
    On Error GoTo Failed
    
    m_clsMortgageProduct.GetProducts SearchWithdrawn
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoEndDateSearch()
    On Error GoTo Failed
    
    m_clsMortgageProduct.GetProducts , , Me.txtEndDate.Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoActiveSearch()
    On Error GoTo Failed
    Dim sDepartment As String
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    m_clsMortgageProduct.GetProducts SearchActive
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    EndWaitCursor
    SetReturnCode MSGFailure
    
    Me.optLender(SEARCH_ALL_LENDERS).Value = True
    optSearchType(SEARCH_ACTIVE_PRODUCTS).Value = True
    cboLender.Enabled = False
    
    PopulateLender
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub PopulateLender()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sLenderField As String
    Dim clsLenderDetails As MortgageLendersTable
    Set clsLenderDetails = New MortgageLendersTable
    
    Set rs = clsLenderDetails.GetLenderCodes()
    
    If Not rs Is Nothing Then
        Set cboLender.RowSource = rs
        sLenderField = clsLenderDetails.GetLenderCodeField()
        
        cboLender.ListField = sLenderField
    End If

    If rs.RecordCount = SINGLE_LENDER Then
        ' Single Lender
        cboLender.SelText = rs(sLenderField)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub MSGEditBox1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub optLender_Click(Index As Integer)
    Dim bEnable As Boolean
    
    Select Case Index
        Case SEARCH_LENDER
            cboLender.Enabled = True
            bEnable = True
        Case SEARCH_ALL_LENDERS
            bEnable = False
            cboLender.SelText = ""
    End Select
    cboLender.Mandatory = bEnable
    cboLender.Enabled = bEnable
End Sub

Private Sub optSearchType_Click(Index As Integer)
    Dim bEnableDate As Boolean
    Dim bClear    As Boolean
    
    bEnableDate = False
    bClear = True
    
    Select Case Index
    Case SEARCH_ACTIVE_PRODUCTS
    
    Case SEARCH_WITHDRAWN_PRODUCTS

    Case SEARCH_END_DATE
        bClear = False
        bEnableDate = True

    End Select
            
    txtEndDate.Enabled = bEnableDate
    txtEndDate.Mandatory = bEnableDate
    
    If bEnableDate Then
        Dim sDate As String
        sDate = g_clsValidation.GetFullDate(Format(Now, "Short date"))
        txtEndDate.Text = sDate
        txtEndDate.SetFocus
    Else
        txtEndDate.Text = ""
    End If

End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub txtEndDate_Validate(Cancel As Boolean)
    Cancel = Not txtEndDate.ValidateData
End Sub

Private Sub txtName_Change()
    If Len(txtName.Text) <= 0 Then
        txtProductCode.Enabled = True
    Else
        txtProductCode.Enabled = False
        txtProductCode.Text = ""
    End If
End Sub

Private Sub txtProductCode_Change()
    If Len(txtProductCode.Text) <= 0 Then
        txtName.Enabled = True
    Else
        txtName.Enabled = False
        txtName.Text = ""
    End If
End Sub

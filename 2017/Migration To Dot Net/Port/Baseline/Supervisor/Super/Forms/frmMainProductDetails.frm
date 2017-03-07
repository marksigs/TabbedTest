VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"

Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMainProductDetails 
   Caption         =   "Main Product Details"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmMainProductDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   315
      Left            =   4380
      TabIndex        =   3
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin Supervisor.MSGTextMulti txtProductName 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   540
      Width           =   3495
      _ExtentX        =   6165
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
      Mandatory       =   -1  'True
      MaxLength       =   80
   End
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
      MaxLength       =   6
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4395
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3075
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   1380
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
   Begin MSGOCX.MSGDataCombo cboLenderCode 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Width           =   3435
      _ExtentX        =   6059
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
      Mandatory       =   -1  'True
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Start Time"
      Height          =   255
      Index           =   5
      Left            =   3300
      TabIndex        =   15
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label lblProductDetails 
      Caption         =   "Lender Code"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   14
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Withdrawn Date"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   13
      Top             =   1860
      Width           =   1335
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Mortgage Product Code"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   12
      Top             =   180
      Width           =   1755
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Mortgage Product Name"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   11
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   10
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "End Date"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmMainProductDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constants
Private Const MORTGAGE_PRODUCT_CODE = 0
Private Const START_DATE = 2
Private Const END_DATE = 3
Private Const WITHDRAWN_DATE = 4

Private Const SINGLE_LENDER = 1

Private m_bIsEdit As Boolean
Private m_clsMortProdLanguage As MortProdLanguageTable
Private m_clsMortgageProduct As MortgageProductTable
Private m_clsMortProdBand As MortgageProductBandsTable
Private m_clsMortProdAdditional As MortProdParamsTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Private Sub cboLenderCode_Validate(Cancel As Boolean)
    Cancel = Not cboLenderCode.ValidateData()
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
'Public Sub SetTableClass(clsTable As TableAccess)
'    Set m_clsMortgageProduct = clsTable
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        bRet = ValidateScreenData()
        
        If bRet = True Then
            SaveScreenData
            LoadProductDetails
            'SetReturnCode
            Hide
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub LoadProductDetails()
    On Error GoTo Failed
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess
    Dim enumReturn As MSGReturnCode
    Dim sTime As String
    Dim vStartDate As Variant
    
    colMatchValues.Add txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    vStartDate = txtProductDetails(START_DATE).Text

    ' Use the time also, if it's there
    sTime = txtStartTime.ClipText
    
    If Len(sTime) > 0 Then
        vStartDate = vStartDate + " " + txtStartTime.Text
    End If
    
    colMatchValues.Add vStartDate
                    
    If g_clsVersion.DoesVersioningExist() Then
        colMatchValues.Add g_sVersionNumber
    End If

    Set clsTableAccess = m_clsMortgageProduct
    clsTableAccess.SetKeyMatchValues colMatchValues

    ' Now load the mortgage product details screen
    frmProductDetails.SetKeys colMatchValues
    
    ' Always edititing
    frmProductDetails.SetIsEdit
    frmProductDetails.Show vbModal, frmMain
    
    enumReturn = frmProductDetails.GetReturnCode()
    Unload frmProductDetails
    
    SetReturnCode enumReturn

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bExists As Boolean
    Dim bTimeValid  As Boolean
    Dim sProductCode As String
    Dim sDate As String
    
    sProductCode = txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    sDate = txtProductDetails(START_DATE).Text
    
    g_clsValidation.ValidateProductDates txtProductDetails(START_DATE), txtProductDetails(END_DATE), txtProductDetails(WITHDRAWN_DATE)
        
    bExists = m_clsMortgageProduct.DoesProductExist(sProductCode, sDate)

    If bExists Then
        txtProductDetails(MORTGAGE_PRODUCT_CODE).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Product exists - please enter a unique combination of Product Code and Start Date"
    End If
    
    bTimeValid = g_clsValidation.ValidateTime(txtStartTime.ClipText)
    
    If Not bTimeValid Then
        txtStartTime.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Time must be between 00:00:00 and 23:59:59"
    End If
    
    ValidateScreenData = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub Form_Load()
    Dim bRet As Boolean
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsMortProdBand = New MortgageProductBandsTable
    Set m_clsMortProdAdditional = New MortProdParamsTable
    Set m_clsMortgageProduct = New MortgageProductTable
    Set m_clsMortProdLanguage = New MortProdLanguageTable
    
    PopulateLenderCodes
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Public Sub SetAddState()

End Sub
Public Sub SetEditState()

End Sub
Private Sub SaveProductLanguage(sCode As String, vStartDate As Variant)
    On Error GoTo Failed
    Dim sLanguage As String
    Dim sLanguageID As String
    Dim sLanguageCombo As String
    Dim colValues As New Collection
    Dim colIDS As New Collection
    
    sLanguageCombo = "MortgageProductLanguage"
    
    g_clsCombo.FindComboGroup sLanguageCombo, colValues, colIDS
    
    ' Mortgage Product Language
    If Not IsEmpty(vStartDate) And Not IsNull(vStartDate) Then
        If Len(vStartDate) > 0 And Len(sCode) > 0 Then
            m_clsMortProdLanguage.SetProductCode sCode
            m_clsMortProdLanguage.SetStartDate vStartDate
            m_clsMortProdLanguage.SetProductName txtProductName.Text
             
            If colValues.Count > 0 Then
                sLanguage = colValues(1) ' First one
                sLanguageID = colIDS(1)
                
                If Len(sLanguage) > 0 And Len(sLanguageID) > 0 Then
                    m_clsMortProdLanguage.SetLanguage sLanguageID
                End If
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveProductLanguage - Start Date and Product Code must be valid"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SaveProductLanguage - Start Date must be valid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim sCode As String
    Dim vStartDate As Variant
    Dim sTime As String
    
    CreateTables
    
    sCode = txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    
    g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), vStartDate, GET_CONTROL_VALUE
    sTime = txtStartTime.ClipText
    
    If Len(sTime) > 0 Then
        vStartDate = vStartDate + " " + txtStartTime.Text
    End If
    
    ' First the Mortgage Product
    m_clsMortgageProduct.SetMortgageProductCode sCode
    
    m_clsMortgageProduct.SetStartDate vStartDate
    
    g_clsFormProcessing.HandleDate txtProductDetails(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProduct.SetEndDate vTmp
    
    g_clsFormProcessing.HandleDate txtProductDetails(WITHDRAWN_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProduct.SetWithdrawnDate vTmp
    
    g_clsFormProcessing.HandleComboText cboLenderCode, vTmp, GET_CONTROL_VALUE
    
    GetOrgID vTmp
    
    m_clsMortgageProduct.SetOrganisationID vTmp
    
    SaveProductLanguage sCode, vStartDate

    ' Now Mortgage Product Parameters
 '   g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortProdAdditional.SetStartDate vStartDate

    m_clsMortProdAdditional.SetMortgageProductCode txtProductDetails(MORTGAGE_PRODUCT_CODE).Text

    ' Now Mortgage Product Bands
    'g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortProdBand.SetStartDate vStartDate

    m_clsMortProdBand.SetMortgageProductCode sCode

    DoUpdates
    
    'g_clsDataAccess.CommitTrans
    
    Exit Sub
Failed:
    'g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CreateTables()
    On Error GoTo Failed
    g_clsFormProcessing.CreateNewRecord m_clsMortgageProduct
    g_clsFormProcessing.CreateNewRecord m_clsMortProdBand
    g_clsFormProcessing.CreateNewRecord m_clsMortProdAdditional
    g_clsFormProcessing.CreateNewRecord m_clsMortProdLanguage
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess

    Set clsTableAccess = m_clsMortgageProduct
    clsTableAccess.Update

    Set clsTableAccess = m_clsMortProdAdditional
    clsTableAccess.Update
        
    Set clsTableAccess = m_clsMortProdBand
    clsTableAccess.Update
        
    Set clsTableAccess = m_clsMortProdLanguage
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetOrgID(vOrgID As Variant)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sLenderCode As String
    
    sLenderCode = Me.cboLenderCode.SelText

    m_clsMortgageProduct.GetOrgIDFromLenderCode sLenderCode, vOrgID

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub txtProductDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtProductDetails(Index).ValidateData()
End Sub
Private Sub PopulateLenderCodes()
    Dim rs As ADODB.Recordset
    Dim sLenderField As String
    
    Set rs = m_clsMortgageProduct.GetLenderCodes()
    
    If Not rs Is Nothing Then
        Set cboLenderCode.RowSource = rs
        sLenderField = m_clsMortgageProduct.GetLenderCodeField()
        
        cboLenderCode.ListField = sLenderField
    End If

    If rs.RecordCount = SINGLE_LENDER Then
        ' Single Lender
        cboLenderCode.SelText = rs(sLenderField)
    End If
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub txtProductName_Validate(Cancel As Boolean)
    Cancel = Not txtProductName.ValidateData()
End Sub
Private Sub txtStartTime_GotFocus()
    On Error GoTo Failed
    
    ' Make sure the cursor is at the start
    txtStartTime.Text = txtStartTime.Text

    ' Highlight the text, if there is any.
    If Len(txtStartTime.ClipText) > 0 Then
        txtStartTime.SelStart = 0
        txtStartTime.SelLength = Len(txtStartTime.Text)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

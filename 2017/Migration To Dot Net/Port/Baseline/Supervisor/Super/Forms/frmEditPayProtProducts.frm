VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPayProtProducts 
   Caption         =   "Add Edit Payment Protection Products"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   Icon            =   "frmEditPayProtProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3315
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   8
      Top             =   3315
      Width           =   1230
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   2820
      TabIndex        =   9
      Top             =   3315
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
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
      MaxLength       =   6
   End
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      HelpContextID   =   2
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   480
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
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      HelpContextID   =   2
      Index           =   2
      Left            =   2100
      TabIndex        =   2
      Top             =   900
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
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      Index           =   3
      Left            =   2100
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      Index           =   4
      Left            =   2100
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      Index           =   5
      Left            =   2100
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSGOCX.MSGDataCombo cboPayProt 
      Height          =   315
      Left            =   2100
      TabIndex        =   6
      Top             =   2580
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
      Mandatory       =   -1  'True
   End
   Begin VB.Label lblBuildingContents 
      Caption         =   "Display Order"
      Height          =   255
      Index           =   5
      Left            =   135
      TabIndex        =   16
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Rate Number"
      Height          =   435
      Left            =   135
      TabIndex        =   15
      Top             =   2640
      Width           =   1755
   End
   Begin VB.Label Label4 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "End Date"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblBuildingContents 
      Caption         =   "Max Applicant Age"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditPayProtProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPayProtProducts
' Description   :
' Change history
' Prog      Date        Description
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const PRODUCT_CODE = 0
Private Const START_DATE = 1
Private Const END_DATE = 2
Private Const PRODUCT_NAME = 3
Private Const MAX_APPLICANT_AGE = 4
Private Const DISPLAY_ORDER = 5

Private m_bIsEdit As Boolean
Private m_clsPayProt As PayProtProductsTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsPayProt = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cboPayProt_Validate(Cancel As Boolean)
    Cancel = Not cboPayProt.ValidateData()
End Sub
Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If

    If bRet = True Then
        txtPayProt(PRODUCT_CODE).SetFocus
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        ValidateScreenData
        SaveScreenData
        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
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
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sSet As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sSet = Me.txtPayProt(PRODUCT_CODE).Text
    colMatchValues.Add sSet
    Set clsTableAccess = m_clsPayProt
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsPayProt
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ValidateScreenData()
    On Error GoTo Failed

    If m_bIsEdit = False Then
        DoesRecordExist
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoesRecordExist()
    On Error GoTo Failed
    
    Dim bRet As Boolean
    Dim sCode As String
    Dim clsPayProt As New PayProtProductsTable
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    
    sCode = txtPayProt(PRODUCT_CODE).Text
    col.Add sCode
    
    Set clsTableAccess = clsPayProt
    bRet = clsTableAccess.DoesRecordExist(col)
    
    If bRet = True Then
        txtPayProt(PRODUCT_CODE).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Record already exists - please enter a unique Product Code"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Load()
    On Error GoTo Failed

    Set m_clsPayProt = New PayProtProductsTable
    ' First the Pay Prot Rates number
    PopulateRatesCombo

    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub PopulateRatesCombo()
    On Error GoTo Failed
    Dim clsPayProtRates As New PayProtRatesTable
    Dim clsTableAccess As TableAccess
    
    clsPayProtRates.GetRates
    
    g_clsFormProcessing.PopulateDataCombo Me.cboPayProt, clsPayProtRates
    Set clsTableAccess = clsPayProtRates
    
    If clsTableAccess.RecordCount = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "No Payment Protection Rates exist"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetAddState()
End Sub
Public Sub SetEditState()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGUID As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection

    Set clsTableAccess = m_clsPayProt
    clsTableAccess.SetKeyMatchValues m_colKeys
    Set rs = clsTableAccess.GetTableData()

    cmdAnother.Enabled = False
    ValidateRecordset rs, "PaymentProtectionProducts"
    
    If rs.RecordCount >= 0 Then
        PopulateScreenFields
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Payment Protection Product record"
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub txtBuildingContents_InvalidData(Index As Integer)

End Sub

Private Sub txtPayProt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtPayProt(Index).ValidateData()
End Sub
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim vTmp As Variant

    ' Product Code
    txtPayProt(PRODUCT_CODE).Text = m_clsPayProt.GetProductCode()

    ' Start Date
    vTmp = m_clsPayProt.GetStartDate()
    g_clsFormProcessing.HandleDate txtPayProt(START_DATE), vTmp, SET_CONTROL_VALUE

    ' End Date
    vTmp = m_clsPayProt.GetEndDate()
    g_clsFormProcessing.HandleDate txtPayProt(END_DATE), vTmp, SET_CONTROL_VALUE

    ' Product Name
    txtPayProt(PRODUCT_NAME).Text = m_clsPayProt.GetProductName()
    
    ' MAX applicant age
    txtPayProt(MAX_APPLICANT_AGE).Text = m_clsPayProt.GetMaxApplicantAge()
    
    ' Display order
    txtPayProt(DISPLAY_ORDER).Text = m_clsPayProt.GetDisplayOrder()

    ' Payment Protection Order
    vTmp = m_clsPayProt.GetRateNumber()
    g_clsFormProcessing.HandleComboText cboPayProt, vTmp, SET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess

    Set clsTableAccess = m_clsPayProt

    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsPayProt
    End If

    ' Product Code
    m_clsPayProt.SetProductCode txtPayProt(PRODUCT_CODE).Text

    ' Start Date
    g_clsFormProcessing.HandleDate txtPayProt(START_DATE), vTmp, GET_CONTROL_VALUE
    m_clsPayProt.SetStartDate vTmp

    ' End Date
    g_clsFormProcessing.HandleDate txtPayProt(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsPayProt.SetEndDate vTmp

    ' Product Name
    m_clsPayProt.SetProductName txtPayProt(PRODUCT_NAME).Text
    
    ' MAX applicant age
    m_clsPayProt.SetMaxApplicantAge txtPayProt(MAX_APPLICANT_AGE).Text
    
    ' Display order
    m_clsPayProt.SetDisplayOrder txtPayProt(DISPLAY_ORDER).Text

    ' Payment Protection number
    g_clsFormProcessing.HandleComboText cboPayProt, vTmp, GET_CONTROL_VALUE
    m_clsPayProt.SetRateNumber vTmp
        
    clsTableAccess.Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

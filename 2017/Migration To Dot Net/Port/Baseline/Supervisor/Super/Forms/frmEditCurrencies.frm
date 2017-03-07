VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditCurrencies 
   Caption         =   "Add/Edit Currencies"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmEditCurrencies.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   5100
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame fraCurrency 
      Height          =   3255
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   6135
      Begin MSGOCX.MSGComboBox cboRoundingDirection 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.CheckBox chkBase 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Currency"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
      Begin MSGOCX.MSGEditBox txtCurrency 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
      Begin MSGOCX.MSGEditBox txtCurrency 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
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
         MaxLength       =   30
      End
      Begin MSGOCX.MSGEditBox txtCurrency 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
      Begin MSGOCX.MSGEditBox txtCurrency 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         MaxLength       =   1
      End
      Begin VB.Label lblCurrencies 
         Caption         =   "Rounding Direction"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   14
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label lblCurrencies 
         Caption         =   "Rounding Precision"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   1860
         Width           =   1635
      End
      Begin VB.Label lblCurrencies 
         Caption         =   "Conversion Rate"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblCurrencies 
         Caption         =   "Currency Name"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblCurrencies 
         Caption         =   "Currency Code"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEditCurrencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditCurrencies
' Description   :   Form which allows the user to edit and add Currency details
' Change history
' Prog      Date        Description
' CL        24/04/01    Initial creation of screen
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_ReturnCode As MSGReturnCode
Private m_clsCurrencies As CurrencyTable
Private m_colKeys  As Collection
Private m_bIsEdit As Boolean
Private m_sCurrencyCode As String
Private Const DEFAULT_PRECISION = 4
Private Const CURRENCY_CODE = 0
Private Const CURRENCY_NAME = 1
Private Const CONVERSION_RATE = 2
Private Const ROUNDING_PRECISION = 3


Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Private Sub cmdAnother_Click()
 On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOkProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        SetFormAddState
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If

    If bRet = True Then
        If m_bIsEdit Then
            txtCurrency(CURRENCY_NAME).SetFocus
        Else
            txtCurrency(CURRENCY_CODE).SetFocus
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub Form_Load()
On Error GoTo Failed
    
    Set m_clsCurrencies = New CurrencyTable
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetFormEditState
        txtCurrency(CURRENCY_CODE).Enabled = False
        cmdAnother.Enabled = False
    Else
        SetFormAddState
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Private Sub PopulateScreenControls()
On Error GoTo Failed
    
    g_clsFormProcessing.PopulateCombo "RoundingDirection", Me.cboRoundingDirection
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetFormEditState()
On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Set clsTableAccess = m_clsCurrencies
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    clsTableAccess.GetTableData POPULATE_KEYS
    
    If clsTableAccess.RecordCount > 0 Then
        PopulateScreenFields
    Else
        g_clsErrorHandling.RaiseError , "Empty Record Set Returned"
    End If
       
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateScreenFields()
On Error GoTo Failed
      
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    'CurrencyCode
    txtCurrency(CURRENCY_CODE).Text = m_clsCurrencies.GetCurrencyCode()
    txtCurrency(CURRENCY_NAME).Text = m_clsCurrencies.GetCurrencyName()
    txtCurrency(CONVERSION_RATE).Text = m_clsCurrencies.GetConversionRate()
    txtCurrency(ROUNDING_PRECISION).Text = m_clsCurrencies.GetRoundingPrecision()
    
    vTmp = m_clsCurrencies.GetRoundingDirection()
    g_clsFormProcessing.HandleComboExtra Me.cboRoundingDirection, vTmp, SET_CONTROL_VALUE
    
    vTmp = m_clsCurrencies.GetBaseCurrencyInd
    g_clsFormProcessing.HandleCheckBox chkBase, vTmp, SET_CONTROL_VALUE
    
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetFormAddState()
On Error GoTo Failed
    
    txtCurrency(ROUNDING_PRECISION).Text = DEFAULT_PRECISION
    g_clsFormProcessing.CreateNewRecord m_clsCurrencies
              
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
               
    'CurrencyCode
    m_clsCurrencies.SetCurrencyCode txtCurrency(CURRENCY_CODE).Text
    'Currency Name
    m_clsCurrencies.SetCurrencyName txtCurrency(CURRENCY_NAME).Text
    'Conversion Rate
    m_clsCurrencies.SetConversionRate txtCurrency(CONVERSION_RATE).Text
    'RoundingPrecision
    m_clsCurrencies.SetRoundingPrecision txtCurrency(ROUNDING_PRECISION).Text
        
    'RoundingDirection
    g_clsFormProcessing.HandleComboExtra Me.cboRoundingDirection, vTmp, GET_CONTROL_VALUE
    m_clsCurrencies.SetRoundingDirection CStr(vTmp)
    
    
    'BaseCurrency
    g_clsFormProcessing.HandleCheckBox Me.chkBase, vTmp, GET_CONTROL_VALUE
    m_clsCurrencies.SetBaseCurrency CStr(vTmp)
        
    Set clsTableAccess = m_clsCurrencies
    TableAccess(m_clsCurrencies).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function DoOkProcessing() As Boolean
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

    DoOkProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOkProcessing = False
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOkProcessing()

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
    Dim sDesc As String
    Dim colMatchValues As Collection

    Set colMatchValues = New Collection

    sDesc = txtCurrency(CURRENCY_NAME).Text

    colMatchValues.Add txtCurrency(CURRENCY_CODE).Text
    
    TableAccess(m_clsCurrencies).SetKeyMatchValues colMatchValues

    g_clsHandleUpdates.SaveChangeRequest m_clsCurrencies, sDesc

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim nCount As Integer
    Dim bRet As Boolean
    Dim bBase As Boolean
    
    bRet = True
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    m_sCurrencyCode = txtCurrency(CURRENCY_CODE).Text
    
    bBase = chkBase.Value
    If bRet And bBase Then
        bRet = m_clsCurrencies.CheckBaseCurrency(m_sCurrencyCode)
    End If
    
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number
End Function
Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sCurrencyCode As String
    Dim colValues As New Collection

    sCurrencyCode = Me.txtCurrency(CURRENCY_CODE).Text

    If Len(sCurrencyCode) > 0 Then
        colValues.Add sCurrencyCode

        bRet = TableAccess(m_clsCurrencies).DoesRecordExist(colValues)

        If bRet = True Then
           g_clsErrorHandling.DisplayError "Currency code must be unique"
           txtCurrency(CURRENCY_CODE).SetFocus
        End If
    End If
    DoesRecordExist = bRet
End Function

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


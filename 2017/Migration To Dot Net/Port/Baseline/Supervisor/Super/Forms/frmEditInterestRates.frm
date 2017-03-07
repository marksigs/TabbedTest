VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditInterestRates 
   Caption         =   "Add/Edit Interest Rates"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmEditInterestRates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGDataCombo cboBaseRateSet 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4080
      Width           =   1230
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   4080
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      _ExtentX        =   3413
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
      MaxLength       =   5
   End
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   507
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      TextType        =   6
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "-1"
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
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   954
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
   Begin MSGOCX.MSGComboBox cboType 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1395
      Width           =   2835
      _ExtentX        =   5001
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
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   1848
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   ""
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
      MaxLength       =   12
   End
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   2295
      Width           =   1395
      _ExtentX        =   2461
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
      MaxLength       =   12
   End
   Begin MSGOCX.MSGEditBox txtInterestRates 
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   2742
      Width           =   1395
      _ExtentX        =   2461
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
      MaxLength       =   12
   End
   Begin VB.Label lblBaseRateSet 
      Caption         =   "Base Rate Set"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Months"
      Height          =   315
      Left            =   4200
      TabIndex        =   13
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Ceiling Rate"
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Floor Rate"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Interest Rate"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Rate Type"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Rate End Date"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Rate Period"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblInterestRates 
      Caption         =   "Sequence Number"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditInterestRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditInterestRates
' Description   : To Add and Edit Interest Rates
'
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' STB       06/12/01    SYS1942 - 'Another' button / transactions.
' DJP       15/12/01    SYS2831 Support client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009. Amended :None
'                                   Added   :DataCombo for BaseRateSet.
' MO        11/06/2002  BMIDS00050 : TabIndex re-arrange
' AW        11/07/02    BMIDS00177   Removed redundant controls
'

Option Explicit

' Private data
Private m_clsInterestRates As EditInterestRate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetInterestRateClass
' Description   : Sets the Processing class to be used by this form for all events - i.e., this
'                 form will call into m_clsInterestRates for all business processing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetInterestRateClass(clsInterestRates As EditInterestRate)
    On Error GoTo Failed
    
    If clsInterestRates Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "InterestRates table class emtpy"
    End If
    
    Set m_clsInterestRates = clsInterestRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, TypeName(Me) & ".SetInterestRateClass " & Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Called when the user clicks the Another button. Need to perform all validation,
'                 save screen data and clear the screen ready for the next record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    On Error GoTo Failed
    
    m_clsInterestRates.Another
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user presses the OK button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    m_clsInterestRates.OK
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : If more than one record has been added, only proceeds if the user allows it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    
    m_clsInterestRates.Cancel
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cboType_Click()
    On Error GoTo Failed
    
    m_clsInterestRates.HandleType
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cboType_Validate(Cancel As Boolean)
    On Error GoTo Failed
    
    Cancel = Not cboType.ValidateData()
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub txtInterestRates_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo Failed
    
    Cancel = Not txtInterestRates(Index).ValidateData()
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsInterestRates = Nothing
End Sub


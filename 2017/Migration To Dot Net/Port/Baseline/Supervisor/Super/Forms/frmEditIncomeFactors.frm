VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIncomeFactors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Income Factors"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmEditIncomeFactors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Income Factors"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4815
      Begin MSGOCX.MSGComboBox cboEmploymentStatus 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGComboBox cboLenderName 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGEditBox txtFactor 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSGOCX.MSGComboBox cboIncomeGroup 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGComboBox cboIncomeType 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label lblFactor 
         Caption         =   "Factor"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblIncomeType 
         Caption         =   "Income Type"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblIncomeGroup 
         Caption         =   "Income Group"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblEmploymentStatus 
         Caption         =   "Employment Status"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblLenderName 
         Caption         =   "Lender Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmEditIncomeFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditIncomeFactors
' Description   : This form allows the setting-up of income factor bands for
'                 each MortgageLender in the database. All method calls and
'                 event code is delegated to a customer-specific support class.
' Change history
' Prog  Date        Description
' STB   13-May-2002 SYS4417 Added AllowableIncomeFactors.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Client-specific support class.
Private m_clsIncomeFactor As IncomeFactor


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Form_Initialize
' Description : Create the support class and associate this form instance with
'               it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    
    On Error GoTo Err_Handler
    
    'Create a support class (client variant).
    Set m_clsIncomeFactor = New IncomeFactorCS
    
    'Associate this form with the support class.
    m_clsIncomeFactor.SetForm Me
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Form_Load
' Description : Setup the screen and load the desired record (if editing).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Err_Handler
    
    m_clsIncomeFactor.Initialise
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Form_Terminate
' Description : Release the support class and clean-up cyclic reference.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Terminate()
    
    On Error GoTo Err_Handler
    
    'Release cyclic reference.
    m_clsIncomeFactor.Terminate
    
    'Release reference to support class.
    Set m_clsIncomeFactor = Nothing

Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_clsIncomeFactor.SetIsEdit bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    m_clsIncomeFactor.SetKeys colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the success code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_clsIncomeFactor.GetReturnCode()
End Function

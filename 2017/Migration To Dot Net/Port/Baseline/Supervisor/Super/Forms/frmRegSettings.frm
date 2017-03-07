VERSION 5.00
Object = "{71BF2CB7-F08B-11D4-8271-0050DA3B9D05}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "frmRegSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetLive 
      Caption         =   "Set Live"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin MSGOCX.MSGEditBox txtSettings 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   540
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      TextType        =   6
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
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
   Begin MSGOCX.MSGEditBox txtSettings 
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   540
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      TextType        =   6
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      BackColor       =   16777215
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
   Begin MSGOCX.MSGDataCombo cboVersion 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1635
      _ExtentX        =   2884
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
   End
   Begin VB.Label lblExistingVersion 
      Caption         =   "Existing Versions"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label lblCurrentVersion 
      Caption         =   "Live version"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblVersion 
      Caption         =   "Supervisor version"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const APPLICATION_SERVER As Integer = 0
Private Const VERSION As Integer = 1
Private Const CURRENT_VERSION = 2
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub cboVersion_Click(Area As Integer)
    If cboVersion.SelectedItem > 0 Then
        cmdSetLive.Enabled = True
        Me.txtSettings(VERSION).Text = cboVersion.SelText
    Else
        cmdSetLive.Enabled = False
    End If
End Sub
Private Sub cmdCancel_Click()
    g_clsDataAccess.RollbackTrans
    Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bSuccess As Boolean
    bSuccess = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bSuccess Then
        SaveSettings
        SetReturnCode
        g_clsDataAccess.CommitTrans
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsDataAccess.RollbackTrans
End Sub

Private Sub cmdSetLive_Click()
    Dim sVersion As String
    
    sVersion = Me.cboVersion.SelText
    
    If Len(sVersion) > 0 Then
        g_clsVersion.SetActiveVersion sVersion
        Me.txtSettings(CURRENT_VERSION).Text = sVersion
        SetCombo
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    g_clsDataAccess.BeginTrans
    SetReturnCode MSGFailure
    
    cmdSetLive.Enabled = False
    
    LoadSettings
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub LoadSettings()
    Dim sCurrentVersion As String
    If g_clsVersion.DoesVersioningExist() Then
        txtSettings(VERSION).Text = CStr(g_sVersionNumber)

        sCurrentVersion = g_clsVersion.GetCurrentVersion()
        txtSettings(CURRENT_VERSION).Text = sCurrentVersion
        SetCombo
    Else
        txtSettings(VERSION).Enabled = False
        lblVersion.Enabled = False
        cboVersion.Enabled = False
        txtSettings(CURRENT_VERSION).Enabled = False
        lblCurrentVersion.Enabled = False
        lblExistingVersion.Enabled = False
    End If
    
    'txtSettings(APPLICATION_SERVER).Text = GetApplicationServer()
    'txtSettings(APPLICATION_SERVER_DIRECTORY).Text = GetApplicationServerDirectory()
End Sub
Private Sub SetCombo()
    On Error GoTo Failed
    Dim sField As String
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    g_clsVersion.GetVersions
    Set clsTableAccess = g_clsVersion
    sField = g_clsVersion.GetVersionField()
    
    Set rs = clsTableAccess.GetRecordSet()
    Set cboVersion.RowSource = rs
    cboVersion.ListField = sField
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveSettings()
    Dim sAppServer As String
    Dim sAppServerDir As String
    Dim sVersion As String
    
    If g_clsVersion.DoesVersioningExist() Then
        sVersion = txtSettings(VERSION).Text
        
        If Len(sVersion) > 0 Then
            g_clsVersion.CreateVersion sVersion
            g_sVersionNumber = CLng(sVersion)
        End If
    End If
    
    'AA 13/03/01 - Commented out, Application Server is now input in frmEditDataBaseOptions
'    sAppServer = txtSettings(APPLICATION_SERVER).Text
'    sAppServer = txtSettings(APPLICATION_SERVER).Text
'    SetKeyValue HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, APP_SERVER, sAppServer, REG_SZ

'    sAppServer = txtSettings(APPLICATION_SERVER_DIRECTORY).Text
'    SetKeyValue HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, APP_SERVER_DIR, sAppServer, REG_SZ



End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmPromotionUser 
   Caption         =   "Promotion User"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "frmPromotionUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGDataCombo cboUser 
      Height          =   345
      Left            =   1830
      TabIndex        =   0
      Top             =   1170
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   609
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
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   1770
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1230
      TabIndex        =   2
      Top             =   1770
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   $"frmPromotionUser.frx":0442
      Height          =   825
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "Available Users"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   1365
   End
End
Attribute VB_Name = "frmPromotionUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmPromotionUser
' Description   : Displays a list of users on the current database and allows the user to select
'                 one.
'
' Change history
' Prog      Date        Description
' DJP       04/03/03    Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_ReturnCode  As MSGReturnCode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Sets up and sets all screen controls to their default values
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim sField As String
    Dim clsUser As OmigaUserTable
    
    Set clsUser = New OmigaUserTable
    
    ' Retrieve a list of users on the database
    clsUser.GetUsers

    ' Set the recordset of the rows returned to the data combo
    Set cboUser.RowSource = TableAccess(clsUser).GetRecordSet

    ' Set the field used to display the users
    sField = clsUser.GetComboField()
    cboUser.ListField = sField
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when the form is loaded - performs all form initialisation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Default to failure
    SetReturnCode MSGFailure
    
    PopulateScreenFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the OK button is pressed - performs all mandatory data processing,
'                 validation, and returns control to the caller if all is successful.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    ' Ensure all data has been entered
    g_clsFormProcessing.DoMandatoryProcessing Me
    
    MsgBox "The selected User ID will be used for the selected target database only", vbInformation
    
    SetReturnCode
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Public Function GetUserID() As String
    On Error GoTo Failed
    Dim sUserID As String
    
    sUserID = cboUser.SelText
    
    GetUserID = sUserID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set cboUser.DataSource = Nothing
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLegalFees 
   Caption         =   "View Legal Fees"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3480
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5820
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtLegalFees 
      Height          =   330
      Left            =   1590
      TabIndex        =   1
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   4
      Mask            =   "1234"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   2520
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4445
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Start Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "End Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "1st Type of Application"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "From £"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "To £"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "1st Legal Fee Amount £"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label lblLenderCode 
      Caption         =   "Lender Code"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "frmLegalFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_frmEdit As Form
Private Sub cmdAdd_Click()
    frmEditLegalFees.Show vbModal
End Sub
Private Sub cmdEdit_Click()
    frmEditLegalFees.Show vbModal
End Sub
Private Sub cmdOK_Click()
    Hide
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub Form_Load()
    Dim line As Collection
    
    Set line = New Collection
    lvListView.View = lvwReport
    
    line.Add ("1234")
    line.Add ("01/01/1999")
    line.Add ("01/12/2000")
    line.Add ("Some application type")
    line.Add ("30000")
    line.Add ("100000")
    line.Add ("300")
    
    g_clsFormProcessing.AddLine lvListView, line
    
End Sub
Public Sub SetEditForm(frmEdit As Form)
    Set m_frmEdit = frmEdit
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

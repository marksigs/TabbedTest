VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Axword Test Harness"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAutoSize 
      Caption         =   "Auto Size"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "600"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtWidth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "800"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdViewXML 
      Caption         =   "Show Output XML"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtFeedback 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3000
      Width           =   6495
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Axword"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtXMLIn 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Height"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   4680
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Width"
      Height          =   195
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Output window"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Input XML"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gstrOutXML As String

Private Sub chkAutoSize_Click()
    If (chkAutoSize.Value = 0) Then
        txtHeight.Enabled = True
        txtWidth.Enabled = True
    Else
        txtHeight.Enabled = False
        txtWidth.Enabled = False
    End If
End Sub

Private Sub cmdRun_Click()
    ' Run axword viewer.
    Dim objAxword As New axword.axwordclass
    
    ' Get settings.
    objAxword.ReadOnly = chkReadOnly.Value
    
    If (chkAutoSize.Value = 0) Then
        objAxword.FormWidth = CInt(txtWidth.Text)
        objAxword.FormHeight = CInt(txtHeight.Text)
    End If
    
    objAxword.XML = txtXMLIn.Text
    
    ' Output feedback.
    If (objAxword.FileSaved) Then
        txtFeedback.Text = "File was saved."
    Else
        txtFeedback.Text = "File not saved."
    End If
    
    txtFeedback.Text = txtFeedback.Text & vbCrLf & "Filesize = " & objAxword.FileSize
    
    gstrOutXML = objAxword.XML
    
    ' Cleanup.
    Set objAxword = Nothing
End Sub

Private Sub cmdViewXML_Click()
    ' Show output XML.
    txtFeedback.Text = gstrOutXML
End Sub

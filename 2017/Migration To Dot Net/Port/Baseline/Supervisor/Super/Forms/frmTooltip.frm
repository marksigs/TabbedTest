VERSION 5.00
Begin VB.Form frmTooltip 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblText 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    lblText.Left = Me.Left
    lblText.Top = Me.Top
    lblText.Width = Me.Width
    lblText.Height = Me.Height
End Sub

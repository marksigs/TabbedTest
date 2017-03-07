VERSION 5.00
Begin VB.UserControl MSGPasswordEditBox 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ScaleHeight     =   315
   ScaleWidth      =   1260
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "MSGPasswordEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmMain
' Description   :   ActiveX control to add functionality to the Password edit
'                   control, such as mandatory processing etc.
' Change history
'
' Prog      Date        Description

' TW        28/11/2007  VR973 - Improved Mandatory checks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Const m_def_Mandatory = 0
Dim m_Mandatory As Boolean
Const m_def_Text = ""
Dim m_Text As String

' TW 28/11/2007 VR973
Private m_clsErrorHandling As ErrorHandling
' TW 28/11/2007 VR973 End


Event MSGGotFocus()
Event Validate(Cancel As Boolean) 'MappingInfo=txtPassword,txtPassword,-1,Validate
'Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtPassword,txtPassword,-1,KeyDown
Public Sub HighLight(Optional bFocus As Boolean = False)
    On Error GoTo Error
    Dim nLen As String
    Dim sTmp As String
    
    
    sTmp = txtPassword.Text
    
    nLen = Len(sTmp)
    
    If (nLen > 0) Then
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
    
    If (bFocus) Then
        txtPassword.SetFocus
    End If

Error:
End Sub

Private Sub txtPassword_GotFocus()
    HighLight
    
    RaiseEvent MSGGotFocus
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
' TW 28/11/2007 VR973
    On Error GoTo Failed
    
    Cancel = Not ValidateData()
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
' TW 28/11/2007 VR973 End
End Sub


'Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub

Private Sub UserControl_Initialize()
    txtPassword.PasswordChar = "*"
End Sub
Private Sub UserControl_Resize()
    UserControl.txtPassword.Width = UserControl.Width
    UserControl.txtPassword.Height = UserControl.Height
End Sub
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_Mandatory = m_def_Mandatory
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim nProp As Integer
    Dim sProp As String
    
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    txtPassword.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    nProp = PropBag.ReadProperty("FontSize", 0)
    
    If (nProp > 0) Then
        UserControl.FontSize = nProp
    End If
    
    sProp = PropBag.ReadProperty("FontName", "")
    
    If (Len(sProp) > 0) Then
        UserControl.FontName = sProp
    End If
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Enabled", txtPassword.Enabled, True)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
End Sub
Public Function ValidateData(Optional bReportError As Boolean = True) As Boolean
    Dim bRet As Boolean
' TW 28/11/2007 VR973
'    bRet = True
'    CheckMandatory
    bRet = CheckMandatory
' TW 28/11/2007 VR973 End
    
    ValidateData = bRet
End Function
Public Function CheckMandatory(Optional bHighLight As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim sVal As String
    
    bRet = True

    sVal = txtPassword.Text
    
    If m_Mandatory = True Then
        If Len(sVal) > 0 Then
            txtPassword.BackColor = vbWhite
            bRet = True
        Else
            bRet = False
            If bHighLight = True Then
                txtPassword.BackColor = vbRed '&H80FF&
            End If
        End If
    End If

    CheckMandatory = bRet
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property
Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
End Property

Public Property Get Enabled() As Boolean
    Enabled = txtPassword.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtPassword.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Text() As String
    Text = txtPassword.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtPassword.Text = New_Text
    PropertyChanged "Text"
End Property


VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmAddPackMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Pack Member"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmAddPackMember.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame fraPackMember 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin MSGOCX.MSGDataCombo cboTemplate 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   300
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label lblSubOrdPack 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblTemplate 
         AutoSize        =   -1  'True
         Caption         =   "Template"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmAddPackMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmAddPackMember
' Description   :   Form which allows the user add a member to a printing pack
' Change history
' Prog      Date        AQR    Description
' RF        06/12/2005  MAR202 Created based on code begun by GHun
' RF        14/12/2005  MAR867 Complete pack handling changes -
'                              simplified screen as just returning simple data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'EPSOM HISTORY
'
' Prog      Date        Description
' TW        19/03/2007  EP2_1276 - Inactive template should not be available for selection in a pack.
'                       Ensured that only active templates which were not already in the pack were
'                       in the list. Also ordered by name.
' TW        02/04/2007  EP2_2137 - When trying to test EP2_1276 I get a runtime error when trying to add a Pack
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_ReturnCode As MSGReturnCode

' Variable in which data captured in the screen is held
Private m_colFormData As Collection

' TW 19/03/2007 EP2_1276
Private m_colKeys As Collection
' TW 19/03/2007 EP2_1276 End

' TW 19/03/2007 EP2_1276
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
' TW 19/03/2007 EP2_1276 End

Private Sub PopulateTemplateCombo()
' TW 19/03/2007 EP2_1276 - Rewritten

Dim rs As ADODB.Recordset
    On Error GoTo Failed

' TW 02/04/2007 EP2_2137
'    Set rs = g_clsDataAccess.GetActiveConnection.Execute("SELECT HOSTTEMPLATEID, HOSTTEMPLATENAME FROM HOSTTEMPLATE WHERE ISNULL(INACTIVEINDICATOR, 0) = 0 ORDER BY HOSTTEMPLATENAME")
    If m_colKeys Is Nothing Then
        Set rs = g_clsDataAccess.GetActiveConnection.Execute("SELECT HOSTTEMPLATEID, HOSTTEMPLATENAME FROM HOSTTEMPLATE WHERE ISNULL(INACTIVEINDICATOR, 0) = 0 ORDER BY HOSTTEMPLATENAME")
    Else
        Set rs = g_clsDataAccess.GetActiveConnection.Execute("SELECT HOSTTEMPLATEID, HOSTTEMPLATENAME FROM HOSTTEMPLATE WHERE ISNULL(INACTIVEINDICATOR, 0) = 0 AND HOSTTEMPLATEID NOT IN (SELECT HOSTTEMPLATEID FROM PACKMEMBER WHERE PACKCONTROLNUMBER = '" & m_colKeys(1) & "') ORDER BY HOSTTEMPLATENAME")
    End If
' TW 02/04/2007 EP2_2137 End
    
    ' Set the datacombo to the recordset returned.
    If Not rs Is Nothing Then
        Set cboTemplate.RowSource = rs
        cboTemplate.ComboDataField "HOSTTEMPLATEID"
        cboTemplate.ListField = "HOSTTEMPLATENAME"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " PopulateTemplateCombo: " & Err.DESCRIPTION
End Sub


Private Sub cmdCancel_Click()
    'Set m_colKeys = Nothing
    Set m_colFormData = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = SetScreenData()
    
    If bRet = False Then
        cboTemplate.SetFocus
    Else
        SetReturnCode
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " cmdOK_Click: " & Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub Form_Load()
    Set m_colFormData = New Collection
    PopulateTemplateCombo
End Sub

Public Function GetScreenData() As Collection
On Error GoTo Failed
    Set GetScreenData = m_colFormData
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError " GetScreenData: " & Err.DESCRIPTION
End Function

Private Function SetScreenData() As Boolean
On Error GoTo Failed
    Dim intSelection As Integer
    Dim intCnt As Integer
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    
    ' update form data
    With cboTemplate
        intSelection = CSafeInt(.SelectedItem)
        If intSelection > 0 Then
            .RowSource.MoveFirst
            .RowSource.Move (intSelection - 1)
            m_colFormData.Add .RowSource.fields("HOSTTEMPLATEID").Value, _
                "HOSTTEMPLATEID"
            blnSuccess = True
        Else
            g_clsErrorHandling.DisplayError "No template selected"
        End If
    End With
    
    SetScreenData = blnSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError " SetScreenData: " & Err.DESCRIPTION
    
    SetScreenData = False
End Function




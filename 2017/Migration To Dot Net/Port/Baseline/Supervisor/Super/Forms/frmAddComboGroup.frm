VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmAddComboGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Combo Group"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmAddComboGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin MSGOCX.MSGTextMulti TxtNotes 
         Height          =   975
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         MaxLength       =   255
      End
      Begin MSGOCX.MSGEditBox TxtGroupName 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
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
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         Caption         =   "Notes"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblGroupName 
         AutoSize        =   -1  'True
         Caption         =   "Group Name"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmAddComboGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmAddComboGroup
' Description   :   Form which allows the user add Combo group
' Change history
' Prog      Date        AQR     Description
' MV        15/01/2003  BM0085  Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsComboGroup As ComboValueGroupTable
Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private Function DoesGroupExist() As Boolean
    
    On Error GoTo Failed
        
    Dim bExists As Boolean
    Dim sGroup As String
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    
    sGroup = TxtGroupName.Text
    Set clsTableAccess = m_clsTableAccess
    
    If Len(sGroup) > 0 Then
        col.Add sGroup
        clsTableAccess.SetKeyMatchValues col
        
        bExists = Not clsTableAccess.DoesRecordExist(col)
                
        If bExists = False Then
            Me.TxtGroupName.SetFocus
            g_clsErrorHandling.DisplayError "Entry for " + sGroup + " already exists"
            DoesGroupExist = True
        End If
    End If
    

    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub DoGroupUpdate()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sGroup As String
    Dim sNotes As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim bUpdate As Boolean
    Dim col As New Collection
        
    On Error GoTo Failed
    sGroup = TxtGroupName.Text
    sNotes = TxtNotes.Text
    bUpdate = False
    
    Set clsTableAccess = m_clsComboGroup
    
    If Len(sGroup) > 0 Then
        DoesGroupExist
        bUpdate = True
    End If
    
    If bUpdate = True Then
        g_clsFormProcessing.CreateNewRecord clsTableAccess
         m_clsComboGroup.SetGroupName sGroup
        m_clsComboGroup.SetGroupNote sNotes
        clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim bRet As Boolean
    
    bRet = SaveScreenData()
    
    If bRet = False Then
        TxtGroupName.SetFocus
    Else
        SetReturnCode
        Hide
    End If
    
    
End Sub

Private Sub Form_Load()
    Set m_clsTableAccess = New ComboValueGroupTable
    SetAddState
End Sub


Public Sub SetAddState()
    
    On Error GoTo Failed
    
    'Add a record to the underlying table object.
    m_clsTableAccess.AddRow
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Function SaveScreenData() As Boolean
    
    Dim vTmp As Variant
    Dim clsParam As ComboValueGroupTable
    Dim sGroupName As String
    Dim sGroupNote As String
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    sGroupName = TxtGroupName.Text
    sGroupNote = TxtNotes.Text
    
    If sGroupName <> "" Then
        
        bRet = DoesGroupExist
        
        If bRet = False Then
            Set clsParam = m_clsTableAccess
            clsParam.SetGroupName sGroupName
            clsParam.SetGroupNote sGroupNote
            m_clsTableAccess.Update
            SaveScreenData = True
        Else
            SaveScreenData = False
        End If
    
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

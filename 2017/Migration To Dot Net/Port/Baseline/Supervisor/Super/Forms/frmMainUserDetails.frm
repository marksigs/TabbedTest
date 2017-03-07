VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMainUserDetails 
   Caption         =   "Main User Details"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   Icon            =   "frmMainUserDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   2
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   3
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGComboBox cboAccessType 
      Height          =   315
      Left            =   5940
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   3
      Left            =   1740
      TabIndex        =   5
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
      MaxLength       =   3
   End
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   4
      Left            =   1740
      TabIndex        =   6
      Top             =   1560
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
      MaxLength       =   30
   End
   Begin MSGOCX.MSGEditBox txtUserDetails 
      Height          =   315
      Index           =   5
      Left            =   1740
      TabIndex        =   7
      Top             =   1920
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
      MaxLength       =   35
   End
   Begin MSGOCX.MSGComboBox cboTitle 
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   840
      Width           =   1755
      _ExtentX        =   3096
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
   Begin VB.Label lblLenderDetails 
      Caption         =   "User Title"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "User Initials"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "User First Forename"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "User Surname"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "User ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Access Type"
      Height          =   255
      Index           =   7
      Left            =   4620
      TabIndex        =   12
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Active To"
      Height          =   255
      Index           =   8
      Left            =   4620
      TabIndex        =   11
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Active From"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "frmMainUserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const USER_ID = 0
Private Const USER_ACTIVE_FROM = 1
Private Const USER_ACTIVE_TO = 2
Private Const USER_INITIALS = 3
Private Const USER_FIRST_FORENAME = 4
Private Const USER_SURNAME = 5
Private m_clsUserTable As OmigaUserTable
Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsORGUSER As OrgUser

'Public Sub SetTableClass(clsTable As TableAccess)
'    Set m_clsUserTable = clsTable
'End Sub
Private Sub cboAccessType_Validate(Cancel As Boolean)
    Cancel = Not cboAccessType.ValidateData()
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    ' Always used for add - this function is just here for compatibility
    m_bIsEdit = False
End Sub
Private Sub SaveUserDetails()
    Dim vTmp As Variant
    Dim sUserID As String
    Dim clsTableAccess As TableAccess
    
    sUserID = txtUserDetails(USER_ID).Text
    
    If Len(sUserID) > 0 Then
        
        g_clsFormProcessing.CreateNewRecord m_clsUserTable
        
        ' User ID
        m_clsUserTable.SetUserID sUserID
        
        ' Access Type
        g_clsFormProcessing.HandleComboExtra cboAccessType, vTmp, GET_CONTROL_VALUE
         m_clsUserTable.SetAccessType CStr(vTmp)
        
        ' Active From
        g_clsFormProcessing.HandleDate txtUserDetails(USER_ACTIVE_FROM), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetActiveFrom vTmp
        
        ' Active To
        g_clsFormProcessing.HandleDate txtUserDetails(USER_ACTIVE_TO), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetActiveTo vTmp
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SaveUserDetails - User is empty"
    End If
    
    Set clsTableAccess = m_clsUserTable
    clsTableAccess.Update
                
    ' Need to create an OrganisationUser entry, too
    Dim clsOrgUser As New OrgUserTable
    Set clsTableAccess = clsOrgUser
    g_clsFormProcessing.CreateNewRecord clsOrgUser
    
    clsOrgUser.SetUserID sUserID
    ' User Initials
    clsOrgUser.SetInitials txtUserDetails(USER_INITIALS).Text
    
    ' User Forename
    clsOrgUser.SetForeName txtUserDetails(USER_FIRST_FORENAME).Text
    
    ' User Surname
    clsOrgUser.SetUserSurname txtUserDetails(USER_SURNAME).Text
    
    ' Title
    g_clsFormProcessing.HandleComboExtra cboTitle, vTmp, GET_CONTROL_VALUE
    clsOrgUser.SetUserTitle CStr(vTmp)
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub cboTitle_Validate(Cancel As Boolean)
    Cancel = Not cboTitle.ValidateData()
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    Dim colMatchData As Collection
    Dim sUserID As String
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        Set clsTableAccess = m_clsUserTable
        Set colMatchData = New Collection
                        
        sUserID = txtUserDetails(USER_ID).Text
        
        If Len(sUserID) > 0 Then
            colMatchData.Add sUserID
            
            bRet = Not clsTableAccess.DoesRecordExist(colMatchData)
            
            If bRet = False Then
                MsgBox "User ID " + sUserID + " already exists"
                txtUserDetails(USER_ID).SetFocus
                bRet = False
            End If
        Else
            MsgBox "ValidateScreenData:User ID is empty"
            bRet = False
        End If
    
        If bRet = True Then
            bRet = g_clsValidation.ValidateActiveFromTo(txtUserDetails(USER_ACTIVE_FROM), txtUserDetails(USER_ACTIVE_TO))
        End If
    End If

    ValidateScreenData = bRet
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()

    If bRet = True Then
        SaveUserDetails
        SaveOSGDetails
        ' Now load the userdetails screen
        LoadUserDetailsScreen
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub LoadUserDetailsScreen()
    Dim sUserID
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    Dim enumReturn As MSGReturnCode
    
    sUserID = Me.txtUserDetails(USER_ID).Text
    
    If Len(sUserID) > 0 Then
        Set clsTableAccess = m_clsUserTable
        colValues.Add sUserID
        clsTableAccess.SetKeyMatchValues colValues
        
        frmEditUser.SetIsEdit False
        frmEditUser.SetKeys colValues
        ' Always editing

        frmEditUser.Show vbModal, Me
        enumReturn = frmEditUser.GetReturnCode()
        Unload frmEditUser
        SetReturnCode enumReturn
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsUserTable = New OmigaUserTable
    g_clsFormProcessing.PopulateCombo "UserAccessType", cboAccessType
    g_clsFormProcessing.PopulateCombo "Title", cboTitle
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub txtUserDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtUserDetails(Index).ValidateData()
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Public Sub SaveOSGDetails()
Dim m_clsORGUSER As New OrgUser
    m_clsORGUSER.CallOSG_SaveUserMortgageAdminDetails
End Sub

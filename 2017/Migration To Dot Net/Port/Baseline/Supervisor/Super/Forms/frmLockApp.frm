VERSION 5.00

Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmCreateLock 
   Caption         =   "Lock Application"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmLockApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Lock"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2865
      TabIndex        =   5
      Top             =   1740
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtAppNo 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Top             =   120
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      TextType        =   4
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
      MaxLength       =   12
   End
   Begin MSGOCX.MSGComboBox cboUnit 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   2235
      _ExtentX        =   3942
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
   Begin MSGOCX.MSGDataCombo cboUser 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   840
      Width           =   2235
      _ExtentX        =   3942
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
      Mandatory       =   -1  'True
   End
   Begin MSGOCX.MSGComboBox cboChannel 
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Top             =   1200
      Width           =   2235
      _ExtentX        =   3942
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
   Begin VB.Label Channel 
      Caption         =   "Channel"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Unit"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label User 
      Caption         =   "User"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblAppNumber 
      Caption         =   "Application Number"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmCreateLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmCreateLocks
' Description   :   Form which allows the user to select a unit, user ID and
'                   Application number and create an application lock with those
'                   values.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Application or customer
Private m_enumLockType As LockType
Private m_clsUser As OmigaUserTable
Private m_ReturnCode As MSGReturnCode
Public Sub SetLockType(enumType As LockType)
    m_enumLockType = enumType
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboUnit_Click
' Description   :   When the user selects a new unit, populate the user combo
'                   with all valid users for that unit.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUnit_Click()
    On Error GoTo Failed
    Dim sUnit As String
    
    sUnit = cboUnit.SelText
    
    If Len(sUnit) > 0 Then
        cboUser.Enabled = True
        
        ' Populate the User Combo...
        m_clsUser.GetUsersFromUnit sUnit
        Dim rs As ADODB.Recordset
        Dim clsTableAccess As TableAccess
        Dim sField As String
        Set clsTableAccess = m_clsUser

        Set rs = clsTableAccess.GetRecordSet
        Set cboUser.RowSource = rs
        sField = m_clsUser.GetComboField()
        cboUser.ListField = sField
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   When the user presses OK, check that all mandatory fields
'                   have been entered the create the lock.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vAppNo As Variant
    Dim vUnit As Variant
    Dim vUser As Variant
    Dim vchannelID As Variant
    Dim clsLocks As New LockProcessing
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        vAppNo = txtAppNo.Text
        vUnit = Me.cboUnit.SelText
        vUser = Me.cboUser.SelText
                        
        g_clsFormProcessing.HandleComboExtra cboChannel, vchannelID, GET_CONTROL_VALUE
        
        If Len(vAppNo) > 0 And Len(vUnit) > 0 And Len(vUser) > 0 And Len(vchannelID) > 0 Then
            clsLocks.CreateLock vAppNo, vUnit, vUser, vchannelID
            SetReturnCode
        End If
        
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub Form_Initialize()
    Set m_clsUser = New OmigaUserTable
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Set the correct text values depending on whether we're creating
'                   an application or customer lock and populate the unit combo.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    Dim clsUnit As New UnitTable
    Dim colValues As New Collection
    
    SetReturnCode MSGFailure
    Select Case m_enumLockType
    Case Application
        Me.Caption = "Create Application Lock"
        Me.lblAppNumber.Caption = "Application Number"
    Case Customer
        Me.Caption = "Create Customer Lock"
        Me.lblAppNumber.Caption = "Customer Number"
    End Select

    ' Populate Unit combo
    clsUnit.GetUnitsAsCollection colValues, True
    cboUnit.SetListTextFromCollection colValues
    g_clsFormProcessing.PopulateChannel cboChannel

    ' Disable User combo
    Me.cboUser.Enabled = False
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

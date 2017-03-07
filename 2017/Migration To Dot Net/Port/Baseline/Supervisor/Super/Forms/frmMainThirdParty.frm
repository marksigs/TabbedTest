VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMainThirdParty 
   Caption         =   "Third Party Details"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "frmMainThirdParty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2220
      Width           =   1215
   End
   Begin MSGOCX.MSGComboBox cboAddressType 
      Height          =   315
      Left            =   1980
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      _ExtentX        =   7541
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
   Begin MSGOCX.MSGEditBox txtThirdParty 
      Height          =   315
      Index           =   0
      Left            =   1980
      TabIndex        =   1
      Top             =   540
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
      MaxLength       =   45
   End
   Begin MSGOCX.MSGComboBox cboOrganisationType 
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   900
      Width           =   4275
      _ExtentX        =   7541
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
      Text            =   ""
   End
   Begin MSGOCX.MSGEditBox txtThirdParty 
      Height          =   315
      Index           =   1
      Left            =   1980
      TabIndex        =   3
      Top             =   1260
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtThirdParty 
      Height          =   315
      Index           =   2
      Left            =   1980
      TabIndex        =   4
      Top             =   1620
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin VB.Label Label5 
      Caption         =   "Active To"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Active From"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Organisation Type"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Name && Address Type"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMainThirdParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmMainThirdParty
' Description   :
'
' Change history
' Prog      Date        Description
' STB       22/11/01    Amended contact details key's collection to use a constant
'                       and the address details key's collection.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const COMPANY_NAME = 0
Private Const ACTIVE_FROM = 1
Private Const ACTIVE_TO = 2

Private m_clsNameAndAddressDir As NameAndAddressDirTable
Private m_clsAddressTable As AddressTable
Private m_clsContactDetails As ContactDetailsTable
Private m_ReturnCode As MSGReturnCode
Private m_sDirectoryGUID As Variant
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(Optional bEdit As Boolean = False)

End Sub
Private Sub cboAddressType_Click()
    On Error GoTo Failed
    ' If the address type is bank/building society, enable organisation type and make
    ' it mandatory. Otherwise, disable it.
    Dim enumValueID As ThirdPartyCombo
    Dim vTmp As Variant
    Dim bEnable As Boolean
    
    bEnable = False
    
    g_clsFormProcessing.HandleComboExtra cboAddressType, vTmp, GET_CONTROL_VALUE

    If Len(vTmp) > 0 Then
        enumValueID = CInt(vTmp)
        If enumValueID = ThirdPartyLender Then
            ' Enable organisation type
            bEnable = True
        End If
    End If
    
    cboOrganisationType.Mandatory = bEnable
    cboOrganisationType.Enabled = bEnable
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cboOrganisationType_Validate(Cancel As Boolean)
    Cancel = Not cboOrganisationType.ValidateData()
End Sub
Private Sub cboAddressType_Validate(Cancel As Boolean)
    Cancel = Not cboAddressType.ValidateData()
End Sub
Private Sub Form_Load()
    Dim bRet As Boolean

    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsNameAndAddressDir = New NameAndAddressDirTable
    Set m_clsAddressTable = New AddressTable
    Set m_clsContactDetails = New ContactDetailsTable

    g_clsFormProcessing.PopulateCombo "OrganisationType", cboOrganisationType
    g_clsFormProcessing.PopulateCombo "ThirdPartyType", cboAddressType

    ' Organisation is disabled by default
    cboOrganisationType.Enabled = False

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Public Sub SetPanelType(enumPanelType As ThirdPartyCombo)
    g_clsFormProcessing.HandleComboExtra cboAddressType, CVar(enumPanelType), SET_CONTROL_VALUE
    cboAddressType.Enabled = False
    cboAddressType.TabStop = False
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            LoadThirdPartyDetails
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub LoadThirdPartyDetails()
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess
    Dim enumReturn As MSGReturnCode
    
    Set clsTableAccess = m_clsNameAndAddressDir
    colMatchValues.Add m_sDirectoryGUID

    clsTableAccess.SetKeyMatchValues colMatchValues

    Unload Me

    ' Now load the mortgage product details screen
    frmEditThirdParty.SetKeys colMatchValues
    frmEditThirdParty.SetIsEdit False
    frmEditThirdParty.Show vbModal, frmMain
    enumReturn = frmEditThirdParty.GetReturnCode()
    Unload frmEditThirdParty
    
    SetReturnCode enumReturn
End Sub
Private Function ValidateScreenData() As Boolean
    Dim nCount As Integer
    Dim bRet As Boolean

    bRet = g_clsValidation.ValidateActiveFromTo(txtThirdParty(ACTIVE_FROM), txtThirdParty(ACTIVE_TO))

    ValidateScreenData = bRet
End Function

Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim vStartDate As Variant
    Dim vEndDate As Variant
    Dim sLenderCode As String
    Dim sAddressGUID As Variant
    Dim vNextFeeSet As Variant
    Dim sLenderName As String
    Dim sContactDetailsGUID As Variant

    CreateTables

    ' Address
    m_clsAddressTable.SetAddressGUID
    sAddressGUID = m_clsAddressTable.GetAddressGUID()

    ' Contact Details
    m_clsContactDetails.SetContactDetailsGUID
    sContactDetailsGUID = m_clsContactDetails.GetContactDetailsGUID()
    
    ' Name and Address Directory
    g_clsFormProcessing.HandleDate txtThirdParty(ACTIVE_FROM), vStartDate, GET_CONTROL_VALUE
    g_clsFormProcessing.HandleDate txtThirdParty(ACTIVE_TO), vEndDate, GET_CONTROL_VALUE
    
    m_sDirectoryGUID = m_clsNameAndAddressDir.SetDirectoryGUID()
    m_clsNameAndAddressDir.SetAddressGUID sAddressGUID
    m_clsNameAndAddressDir.SetContactDetailsGUID sContactDetailsGUID
    m_clsNameAndAddressDir.SetActiveFrom vStartDate
    m_clsNameAndAddressDir.SetActiveTo vEndDate
    m_clsNameAndAddressDir.SetCompanyName Me.txtThirdParty(COMPANY_NAME).Text

    g_clsFormProcessing.HandleComboExtra cboOrganisationType, vTmp, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetOrganisationType vTmp

    g_clsFormProcessing.HandleComboExtra cboAddressType, vTmp, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetNameAndAddressType vTmp

    DoUpdates

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess

    Set clsTableAccess = m_clsContactDetails
    clsTableAccess.Update

    Set clsTableAccess = m_clsAddressTable
    clsTableAccess.Update

    Set clsTableAccess = m_clsNameAndAddressDir
    clsTableAccess.Update
    
    Exit Sub
Failed:
    'g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CreateTables()
    On Error GoTo Failed

    ' Name And Address Directory
    g_clsFormProcessing.CreateNewRecord m_clsNameAndAddressDir
    
    ' Address table
    g_clsFormProcessing.CreateNewRecord m_clsAddressTable

    ' Contact Details
    g_clsFormProcessing.CreateNewRecord m_clsContactDetails

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub txtThirdParty_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtThirdParty(Index).ValidateData()
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

VERSION 5.00
Object = "{C4FB9DA8-3561-4EF3-8006-C4C5037F7804}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPrintingDocument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Document Locations"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmEditPrintingDocument.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame fraPrintingDocument 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin MSGOCX.MSGMulti txtACEXML 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   2640
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Enabled         =   -1  'True
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
         MaxLength       =   250
      End
      Begin MSGOCX.MSGMulti txtACEXSL 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   3120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Enabled         =   -1  'True
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
         MaxLength       =   250
      End
      Begin MSGOCX.MSGComboBox cboDeliveryEngineType 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2160
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSGOCX.MSGMulti txtDescription 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
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
         Text            =   ""
         Mandatory       =   -1  'True
         MaxLength       =   256
      End
      Begin MSGOCX.MSGMulti txtFileLocation 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   5895
         _ExtentX        =   10398
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
         Text            =   ""
         Mandatory       =   -1  'True
         MaxLength       =   256
      End
      Begin MSGOCX.MSGEditBox txtTemplateName 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
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
         MaxLength       =   32
      End
      Begin MSGOCX.MSGEditBox txtDPSTemplateID 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
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
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "ACE XSL"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3180
         Width           =   660
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "ACE XML"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2700
         Width           =   690
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Engine Type"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Template name"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "File Location"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "DPS Template ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmEditPrintingDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
' Form          :   frmEditPrintingDocument
' Description   :   Form which allows the user to edit and add Print Template details
' Change history
' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply BBG1370 (New printing documents screen)
'------------------------------------------------------------------------------------
Option Explicit

Private m_ReturnCode    As MSGReturnCode
Private m_clsTemplates  As TemplateTable
Private m_colKeys       As Collection
Private m_bIsEdit       As Boolean

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

'SDS  04/10/04  BBG1542__START
'Private Sub cboDeliveryEngineType_Click()
'    Dim clsComboValidation As ComboValidationTable
'    Const ACEXML As Integer = 5
'    Const ACEXSL As Integer = 6
'    Set clsComboValidation = New ComboValidationTable
'
'    If cboDeliveryEngineType.GetExtra(cboDeliveryEngineType.ListIndex) = _
'            clsComboValidation.GetSingleValueFromValidation("DeliveryEngineType", "eKFI") Then
'        lblPrintingDocuments(ACEXML).Enabled = True
'        lblPrintingDocuments(ACEXSL).Enabled = True
'        txtACEXML.Enabled = True
'        txtACEXSL.Enabled = True
'    Else
'        lblPrintingDocuments(ACEXML).Enabled = False
'        lblPrintingDocuments(ACEXSL).Enabled = False
'        txtACEXML.Enabled = False
'        txtACEXSL.Enabled = False
'    End If
'End Sub
'SDS  04/10/04  BBG1542__END

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Failed
    
    Set m_clsTemplates = New TemplateTable
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub PopulateScreenControls()
On Error GoTo Failed
    
    g_clsFormProcessing.PopulateCombo "DeliveryEngineType", Me.cboDeliveryEngineType
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetEditState()
On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Set clsTableAccess = m_clsTemplates
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    clsTableAccess.GetTableData
    
    If clsTableAccess.RecordCount > 0 Then
        txtDPSTemplateID.Enabled = False
        PopulateScreenFields
    Else
        g_clsErrorHandling.RaiseError , "Empty Record Set Returned"
    End If
       
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenFields()
On Error GoTo Failed
      
    Dim vTmp As Variant
    
    txtDPSTemplateID.Text = m_clsTemplates.GetTemplateID()
    txtTemplateName.Text = m_clsTemplates.GetTemplateName()
    txtDescription.Text = m_clsTemplates.GetDescription()
    txtFileLocation.Text = m_clsTemplates.GetFileName()
    
    vTmp = m_clsTemplates.GetDeliveryEngineType
    g_clsFormProcessing.HandleComboExtra Me.cboDeliveryEngineType, vTmp, SET_CONTROL_VALUE
        
    txtACEXML.Text = m_clsTemplates.GetACEXML()
    txtACEXSL.Text = m_clsTemplates.GetACEXSL()
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetAddState()
On Error GoTo Failed
    
    g_clsFormProcessing.CreateNewRecord m_clsTemplates
              
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
               
    'Template ID
    m_clsTemplates.SetTemplateID txtDPSTemplateID.Text
    'Template Name
    m_clsTemplates.SetTemplateName txtTemplateName.Text
    'Description
    m_clsTemplates.SetDescription txtDescription.Text
    'File Name
    m_clsTemplates.SetFileName txtFileLocation.Text
        
    'Delivery Engine Type
    g_clsFormProcessing.HandleComboExtra Me.cboDeliveryEngineType, vTmp, GET_CONTROL_VALUE
    m_clsTemplates.SetDeliveryEngineType CStr(vTmp)
    'ACEXML
    m_clsTemplates.SetACEXML txtACEXML.Text
    'ACEXSL
    m_clsTemplates.SetACEXSL txtACEXSL.Text
    'SecurityLevel - Hardcoded to 10
    m_clsTemplates.SetSecurityLevel 10
    'StageNumber - Hardcoded to 0
    m_clsTemplates.SetStageNumber 0
    'Language - Hardcoded to 1
    m_clsTemplates.SetLanguage 1
            
    Set clsTableAccess = m_clsTemplates
    TableAccess(m_clsTemplates).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
            
        End If
        
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function

Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim colMatchValues As Collection

    Set colMatchValues = New Collection

    'sDesc = txtTemplateName.Text
    'SDS BBG1370 - Store template ID instead of template name
    sDesc = txtDPSTemplateID.Text

    colMatchValues.Add sDesc
    
    TableAccess(m_clsTemplates).SetKeyMatchValues colMatchValues

    g_clsHandleUpdates.SaveChangeRequest m_clsTemplates, sDesc

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If

    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number
End Function

Private Function DoesRecordExist() As Boolean
    Dim bRet As Boolean
    Dim sTemplateID As String
    Dim colValues As Collection
    Set colValues = New Collection

    sTemplateID = Trim(txtDPSTemplateID.Text)
    If Len(sTemplateID) > 0 Then
        colValues.Add sTemplateID

        bRet = TableAccess(m_clsTemplates).DoesRecordExist(colValues)

        If bRet = True Then
           g_clsErrorHandling.DisplayError "Template ID must be unique"
           txtDPSTemplateID.SetFocus
        End If
    End If
    DoesRecordExist = bRet
    Set colValues = Nothing
End Function

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function



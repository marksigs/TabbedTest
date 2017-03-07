VERSION 5.00
Begin VB.Form frmEditPrintingTemplate 
   Caption         =   "Add/Edit Printing Template"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmEditPrintingTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   6660
      Width           =   1215
   End
   Begin VB.Frame fraStages 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   4140
      Width           =   7575
   End
   Begin VB.Frame fraDetail 
      Height          =   3915
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   7575
      Begin VB.Frame fraPrintMenuAccess 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5460
         TabIndex        =   1
         Top             =   2580
         Width           =   1875
         Begin VB.OptionButton optPrintMenu 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   615
         End
         Begin VB.OptionButton optPrintMenu 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   60
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.CheckBox chkInactive 
         Enabled         =   0   'False
         Height          =   195
         Left            =   5520
         TabIndex        =   6
         Top             =   1800
         Width           =   555
      End
      Begin VB.Frame fraCustInd 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   2580
         Width           =   1875
         Begin VB.OptionButton optCustInd 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   60
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optCustInd 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   615
         End
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Print Data Manager Method Name"
         Height          =   375
         Index           =   14
         Left            =   4080
         TabIndex        =   3
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Remote Location"
         Height          =   195
         Index           =   13
         Left            =   4080
         TabIndex        =   4
         Top             =   3540
         Width           =   1335
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Printer Destination"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   5
         Top             =   3540
         Width           =   1335
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Recipient Type"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   7
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Print Menu Access"
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Customer Indicator"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Default Copies"
         Height          =   195
         Index           =   8
         Left            =   4080
         TabIndex        =   14
         Top             =   2220
         Width           =   1155
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Maximum Copies"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Set Inactive"
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   16
         Top             =   1740
         Width           =   1275
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Template Id"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Document Group"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Template Name"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   24
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Minimum Role Level"
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "DPS Template Id"
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEditPrintingTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPrintingTemplates
' Description   : Form which allows the adding/editing of Print Templates
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added Form
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private m_clsPrintTemplate As PrintingTemplateTable
Private m_bIsEdit As Boolean
Private m_colKeys As Collection
Private m_ReturnCode As MSGReturnCode
Private m_sTemplateID As String
Private m_clsComboValidation As ComboValidationTable
Private m_bScreenUpdated As Boolean
Private m_clsAvailableTemplates As AvailableTemplatesTable
Private m_bTaskManagementExists As Boolean

'Form Constants
Private Const TEMPLATEID = 0
Private Const MAXIMUM_COPIES = 2
Private Const TEMPLATE_NAME = 4
Private Const DEFAULT_COPIES = 6

'ComboGroup Constants
Private Const m_sComboMinimumRole As String = "UserRole"
Private Const m_sComboDocumentType As String = "PrintDocumentType"
Private Const m_sComboRecipientType As String = "PrintRecipientType"
Private Const m_sComboPrinterDestination As String = "PrinterDestination"
Private Const m_sComboApplicationStage As String = "ApplicationStage"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormEditState
' Description   :   Initialises the form in Edit Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormEditState()
    On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim bRet As Boolean
    
    Set clsTableAccess = m_clsPrintTemplate
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData(POPULATE_KEYS)
    bRet = True
    
    If Not rs Is Nothing Then
        If rs.RecordCount = 0 Then
            bRet = False
        End If
    Else
        bRet = False
    End If
    
    If Not bRet Then
        g_clsErrorHandling.RaiseError errRecordSetEmpty
    Else
        PopulateScreenFields
    End If
    
    txtDetail(TEMPLATEID).Enabled = False
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cboDocumentType_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cboMinRoleLevel_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cboRecipientType_Click()
    m_bScreenUpdated = True
End Sub

Private Sub chkInactive_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = DoOKProcessing
    If bRet Then
        SetReturnCode
        Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Activate()
    m_bScreenUpdated = False
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    
    Set m_clsPrintTemplate = New PrintingTemplateTable
    Set m_clsComboValidation = New ComboValidationTable
    Set m_clsAvailableTemplates = New AvailableTemplatesTable
    
    m_sTemplateID = ""
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetFormEditState
    Else
        SetFormAddState
    End If
    
    m_bScreenUpdated = False
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormAddState
' Description   :   Initialises the form in Edit Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormAddState()
    
    On Error GoTo Failed
        Dim rs As ADODB.Recordset
        
        'Populate Selected Items with an empty Recordset
        Set rs = TableAccess(m_clsAvailableTemplates).GetTableData(POPULATE_EMPTY)
        TableAccess(m_clsAvailableTemplates).SetRecordSet rs
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenControls()
    On Error GoTo Failed
    Dim sField As String
    Dim clsComboValue As ComboValueTable
    Dim rs As ADODB.Recordset
    
    Set clsComboValue = New ComboValueTable
    'Document Group
    g_clsFormProcessing.PopulateCombo m_sComboDocumentType, cboDocumentType
    
    'Minimun Role
    g_clsFormProcessing.PopulateCombo m_sComboMinimumRole, cboMinRoleLevel
    
    'Recipient Type
    g_clsFormProcessing.PopulateCombo m_sComboRecipientType, cboRecipientType
    
    'Printer Destination
    'Returns a Recordset with the corresponding validation types
    Set rs = m_clsComboValidation.GetComboGroupWithValidationType(m_sComboPrinterDestination)
    TableAccess(m_clsComboValidation).SetRecordSet rs
    
    Set rs = clsComboValue.GetComboGroupValues(m_sComboPrinterDestination)
    Set cboPrinterDestination.RowSource = rs
    sField = m_clsPrintTemplate.GetListField
    cboPrinterDestination.ListField = sField
    
    'Populate Stages SwapList
    SetStageHeaders
    'Does the task management schema exist?
    
    m_bTaskManagementExists = g_clsMainSupport.DoesTaskManagementExist
    If m_bTaskManagementExists Then
        PopulateAvailableItemsFromRS
    Else
        PopulateAvailableItemsFromComboGroup
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Private Sub SetStageHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim colLine As Collection

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Stage Name"
    headers.Add lvHeaders

    swapStage.SetFirstColumnHeaders headers
    swapStage.SetSecondColumnHeaders headers

    swapStage.SetFirstTitle "Select Stage(s)"
    swapStage.SetSecondTitle "Selected Stage(s)"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim sStageName As String
    Dim sStageID As String
    Dim colLine As Collection
    Dim bExists As Boolean
    
    Set colLine = New Collection
    Set clsSwapExtra = New SwapExtra
    
    If TableAccess(m_clsAvailableTemplates).RecordCount > 0 Then
        Set rs = TableAccess(m_clsAvailableTemplates).GetRecordSet
        
        Do While Not rs.EOF
            
            Set colLine = New Collection
            sStageID = m_clsAvailableTemplates.GetStageID
            sStageName = m_clsAvailableTemplates.GetStageName
            
            bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
            
            If Not bExists Then
                Set clsSwapExtra = New SwapExtra
                clsSwapExtra.SetValueID sStageID
                
                colLine.Add sStageName
                swapStage.AddLineSecond colLine, clsSwapExtra
            End If
            
            rs.MoveNext
        Loop
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim vVal As Variant
    Dim clsComboUtils As New ComboUtils
    
    'TEXTBOXES - Template ID
    txtDetail(TEMPLATEID).Text = m_clsPrintTemplate.GetHostTemplateID
    m_sTemplateID = txtDetail(TEMPLATEID).Text
    
    'DPS TemplateID
    txtDPSTemplateID.Text = m_clsPrintTemplate.GetDPSTemplateID
    
    'Template Name
    txtDetail(TEMPLATE_NAME).Text = m_clsPrintTemplate.GetHostTemplateName
    
    'Template Description
    txtDescription.Text = m_clsPrintTemplate.GetHostTemplateDescription
    
    ' Maximum Copies
    txtDetail(MAXIMUM_COPIES).Text = m_clsPrintTemplate.GetMaxCopies
    
    'Default Copies
    txtDetail(DEFAULT_COPIES).Text = m_clsPrintTemplate.GetDefaultCopies
    
    'Remote Location
    txtRemoteLocation.Text = m_clsPrintTemplate.GetRemotePrinterLocation
    
    'COMBO'S - Document Group
    vVal = m_clsPrintTemplate.GetTemplateGroupID
    g_clsFormProcessing.HandleComboExtra cboDocumentType, vVal, SET_CONTROL_VALUE
    
    'Minimum Role
    vVal = m_clsPrintTemplate.GetMinRoleLevel
    g_clsFormProcessing.HandleComboExtra cboMinRoleLevel, vVal, SET_CONTROL_VALUE
    
    'Recipient Type
    vVal = m_clsPrintTemplate.GetRecipientTypeID
    g_clsFormProcessing.HandleComboExtra cboRecipientType, vVal, SET_CONTROL_VALUE
    
    'Printer Destination
    vVal = clsComboUtils.GetValueNameFromValueID(m_clsPrintTemplate.GetPrinterDestinationType, m_sComboPrinterDestination)
    g_clsFormProcessing.HandleDataComboText cboPrinterDestination, vVal, SET_CONTROL_VALUE
    
    'Print Manager Method Name
    txtPrintMethodName.Text = m_clsPrintTemplate.GetPrintManagerMethod
    
    'Customer Indicator
    g_clsFormProcessing.HandleRadioButtons optCustInd(OPT_YES), optCustInd(OPT_NO), m_clsPrintTemplate.GetCustSpecificInd, SET_CONTROL_VALUE
    
    'Print Menu Access?
    g_clsFormProcessing.HandleRadioButtons optPrintMenu(OPT_YES), optPrintMenu(OPT_NO), m_clsPrintTemplate.GetPrintMenuAccessInd, SET_CONTROL_VALUE
    
    'Set Inative
    vVal = m_clsPrintTemplate.GetInactiveIndicator
    g_clsFormProcessing.HandleCheckBox chkInactive, vVal, SET_CONTROL_VALUE
    
    'Does the task managment schema exist?
    If m_bTaskManagementExists Then
        'Populate from Task Table
        m_clsAvailableTemplates.GetSelectedStages m_sTemplateID
        PopulateSelectedItems
        m_clsAvailableTemplates.GetSelectedStages m_sTemplateID, True
    Else
        'Populate From Combo Group
        m_clsAvailableTemplates.GetSelectedStagesFromComboGroup m_sTemplateID
        PopulateSelectedItems
        m_clsAvailableTemplates.GetSelectedStagesFromComboGroup m_sTemplateID, False
    End If
    chkInactive.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveScreenFields()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim vVal As Variant
    
    Set clsTableAccess = m_clsPrintTemplate
    
    'Next Set the TemplateID
    If m_bIsEdit Then
        m_sTemplateID = txtDetail(TEMPLATEID).Text
    Else
        g_clsFormProcessing.CreateNewRecord m_clsPrintTemplate
        m_sTemplateID = m_clsPrintTemplate.GetNextTemplateID
    End If
    
    m_clsPrintTemplate.SetTemplateID m_sTemplateID
    
    m_clsPrintTemplate.SetDPSTemplateID txtDPSTemplateID.Text
    
    g_clsFormProcessing.HandleComboExtra cboDocumentType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetTemplateGroupID vVal
    
    m_clsPrintTemplate.SetHostTemplateName txtDetail(TEMPLATE_NAME).Text
    
    m_clsPrintTemplate.SetHostTemplateDescription txtDescription.Text
    
    g_clsFormProcessing.HandleComboExtra cboMinRoleLevel, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetMinRoleLevel vVal
    
    g_clsFormProcessing.HandleCheckBox chkInactive, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetInactiveIndicator vVal
    
    m_clsPrintTemplate.SetMaxCopies txtDetail(MAXIMUM_COPIES).Text
    m_clsPrintTemplate.SetDefaultCopies txtDetail(DEFAULT_COPIES).Text
    
    g_clsFormProcessing.HandleRadioButtons optCustInd(OPT_YES), optCustInd(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetCustomerSpecificInd vVal
    
    'Print Menu
    g_clsFormProcessing.HandleRadioButtons optPrintMenu(OPT_YES), optPrintMenu(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPrintMenuAccessInd vVal
    
    'Recipient Type
    g_clsFormProcessing.HandleComboExtra cboRecipientType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetRecipientType vVal
    
    'Printer Destination
    vVal = m_clsComboValidation.GetValueID
    m_clsPrintTemplate.SetPrintDestinationType vVal
    
    'Remote Location
    m_clsPrintTemplate.SetRemotePrinterLocation txtRemoteLocation.Text
    
    'Print Manager Method Name
    m_clsPrintTemplate.SetPrintManagerMethodName txtPrintMethodName.Text
    
    clsTableAccess.Update
    
    'Stages
    SaveSelectedStages
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboPrinterDestination_Change
' Description   :   Checks if the selected item is Remote or Both, then enable Remote Location
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPrinterDestination_Change()

    Dim sValidation As String
    Dim nCnt As Integer
    On Error GoTo Failed
    Dim colValidationType As Collection
    
    Set colValidationType = New Collection
    'Get the validationType of the selected item
    sValidation = m_clsComboValidation.GetValidationTypeFromCombo(CLng(cboPrinterDestination.SelectedItem))
    

    Set colValidationType = m_clsComboValidation.GetComboValidationAsCollection(m_sComboPrinterDestination, m_clsComboValidation.GetValueID)
    
    For nCnt = 1 To colValidationType.Count
        sValidation = colValidationType(nCnt)
        Select Case sValidation
            Case PRINTER_LOCATION_REMOTE, PRINTER_LOCATION_FILE
                txtRemoteLocation.Enabled = True
                txtRemoteLocation.Mandatory = True
                Exit For
            Case Else
                txtRemoteLocation.Enabled = False
                txtRemoteLocation.Mandatory = False
        End Select
    Next
    m_bScreenUpdated = True
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean

    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim vSaveChanges As Variant
    
    bShowError = True

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If m_bIsEdit And bRet Then
        'This record is an edit and Valid
        vSaveChanges = MsgBox("You are about to update the selected Printing Template. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
         Select Case vSaveChanges
            Case vbYes
                If bRet Then
                    SaveScreenFields
                End If
            Case vbNo
                bRet = False
        End Select
    ElseIf Not m_bIsEdit Then
        SaveScreenFields
'        SaveChangeRequest
        bRet = True
    End If
        
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub optCustInd_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub optPrintMenu_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub
Private Sub txtDetail_Change(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub SaveSelectedStages()
    On Error GoTo Failed
    Dim nThisItem As Integer
    Dim colValues As Collection
    Dim sStageID As String
    Dim clsSwapExtra As SwapExtra
    
    If TableAccess(m_clsAvailableTemplates).RecordCount > 0 Then
        'Delete all Existing Records
        TableAccess(m_clsAvailableTemplates).DeleteAllRows
    End If
    
    If swapStage.GetSecondCount > 0 Then
        For nThisItem = 1 To swapStage.GetSecondCount
            'Add a row and set the fields
            TableAccess(m_clsAvailableTemplates).AddRow
        
            Set colValues = swapStage.GetLineSecond(nThisItem, clsSwapExtra)
    
            sStageID = clsSwapExtra.GetValueID
            
            m_clsAvailableTemplates.SetTemplateID m_sTemplateID
            m_clsAvailableTemplates.SetStageID sStageID
        Next
    End If
    
    TableAccess(m_clsAvailableTemplates).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateAvailableItemsFromComboGroup()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsComboValidation As ComboValidationTable
    Dim nThisItem As Integer
    Dim clsSwapExtra As SwapExtra
    Dim sStageID As String
    Dim sStageName As String
    Dim bExists As Boolean
    Dim colLine As Collection
    
    Set clsComboValidation = New ComboValidationTable
    
    Set rs = clsComboValidation.GetComboGroupWithValidationType(m_sComboApplicationStage)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For nThisItem = 1 To rs.RecordCount
                
                Set colLine = New Collection
                
                sStageID = clsComboValidation.GetValueID()
                sStageName = clsComboValidation.GetValueName()
                
                bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
                
                If Not bExists Then
                    Set clsSwapExtra = New SwapExtra
                    clsSwapExtra.SetValueID sStageID
                    
                    colLine.Add sStageName
                    swapStage.AddLineFirst colLine, clsSwapExtra
                
                End If
                rs.MoveNext
            Next
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesDPSTemplateIDExist
' Description   :   Checks to see if a DPS Document Record exists with the ID entered
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesDPSTemplateIDExist() As Boolean
    On Error GoTo Failed
    Dim clsTemplate As TemplateTable
    Dim bRet As Boolean
    Dim colValues As Collection
    
    Set colValues = New Collection
    
    Set clsTemplate = New TemplateTable
    
    colValues.Add CLng(txtDPSTemplateID.Text)
    
    bRet = TableAccess(clsTemplate).DoesRecordExist(colValues, TableAccess(clsTemplate).GetKeyMatchFields)
    
    
    DoesDPSTemplateIDExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateAvailableItemsFromRS
' Description   :   If the Task Management Schema exists then the available Tasks must be populated from the Task Table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItemsFromRS()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsComboValidation As ComboValidationTable
    Dim nThisItem As Integer
    Dim clsSwapExtra As SwapExtra
    Dim sStageID As String
    Dim sStageName As String
    Dim bExists As Boolean
    Dim colLine As Collection
    
    Set clsComboValidation = New ComboValidationTable
    If m_bIsEdit Then
        m_sTemplateID = m_colKeys(1)
    End If
    
    Set rs = m_clsAvailableTemplates.GetUnassignedStages(m_sTemplateID)
    TableAccess(m_clsAvailableTemplates).SetRecordSet rs
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For nThisItem = 1 To rs.RecordCount
                
                Set colLine = New Collection
                
                sStageID = m_clsAvailableTemplates.GetStageID()
                sStageName = m_clsAvailableTemplates.GetStageName()
                
                bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
                
                If Not bExists Then
                    Set clsSwapExtra = New SwapExtra
                    clsSwapExtra.SetValueID sStageID
                    
                    colLine.Add sStageName
                    swapStage.AddLineFirst colLine, clsSwapExtra
                
                End If
                rs.MoveNext
            Next
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub txtDPSTemplateID_Validate(Cancel As Boolean)
    On Error GoTo Failed
    
    Dim bRet As Boolean
        
    If Len(txtDPSTemplateID.Text) > 0 Then
        bRet = DoesDPSTemplateIDExist
        
        If Not bRet Then
            Cancel = True
            txtDPSTemplateID.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Your DPS template Id is not valid, please re-enter."
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

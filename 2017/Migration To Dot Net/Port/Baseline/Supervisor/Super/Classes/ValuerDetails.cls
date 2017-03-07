VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValuerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ValuerDetails
' Description   : Support class for the Panel type of Valuer. Contains all code to maintain a
'                 valuer.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' STB       29/11/2001  SYS2912 - Removed the frmMainThirdParty screen and
'                       ported its functionality to frmEditThirdParty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Constants.
Private Const PANEL_ID = 0
Private Const ASSOCIATE_PANEL_ID = 1

'Local Table Classes.
Private m_clsPanel As PanelTable
Private m_clsValuer As PanelValuerTable
Private m_clsValuationType As PanelValuerTypeTable

'Primary key for the third party's directory record.
Private m_vDirectoryGUID As Variant

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'Collections to hold Combo-related values.
Private m_colComboIDS As Collection
Private m_colComboValues As Collection


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if this tab-handler is in add or edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetDirectoryGUID
' Description   : Sets the primary key for the parent record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDirectoryGUID(ByVal vGUID As Variant)
    m_vDirectoryGUID = vGUID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the add/edit mode of this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Create underlying object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Class_Initialize()
    
    Set m_clsPanel = New PanelTable
    Set m_clsValuer = New PanelValuerTable
    Set m_clsValuationType = New PanelValuerTypeTable

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated accoding to the
'                 underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    Dim sError As String
    Dim rs As ADODB.Recordset
    Dim nRecordCount As Integer
    Dim vDirectoryGUID As Variant
    Dim colValues As New Collection
    
    On Error GoTo Failed
    
    sError = "Valuer:Initialise - "
    
    m_bIsEdit = bIsEdit
    
    If bIsEdit = True Then
        ' DJP SQL Server port
        colValues.Add m_vDirectoryGUID
        
        ' First, the Valuer details
        TableAccess(m_clsValuer).SetKeyMatchValues colValues
        TableAccess(m_clsValuer).GetTableData
        TableAccess(m_clsValuer).ValidateData
        
        nRecordCount = TableAccess(m_clsValuer).RecordCount()
        
        If nRecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Incorrect number of Valuer records returned: " & nRecordCount
        End If
    
        ' Next, the Panel details
        TableAccess(m_clsPanel).SetKeyMatchValues colValues
        TableAccess(m_clsPanel).GetTableData
        TableAccess(m_clsPanel).ValidateData
        
        nRecordCount = TableAccess(m_clsPanel).RecordCount
        
        If nRecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Incorrect number of Valuer records returned: " & nRecordCount
        End If
    Else
        g_clsFormProcessing.CreateNewRecord m_clsPanel
        g_clsFormProcessing.CreateNewRecord m_clsValuer
    End If

    SetupValuationGrid

    g_clsFormProcessing.PopulateCombo "PaymentMethod", frmEditThirdParty.cboValuerPaymentMethod
    g_clsFormProcessing.PopulateCombo "ValuerType", frmEditThirdParty.cboValuerType
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupValuationGrid
' Description   : Performs initialisation for the ValuationType table and binds it to the
'                 datagrid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupValuationGrid()

    Dim colValues As New Collection

    On Error GoTo Failed
    
    g_clsCombo.FindComboGroup "ValuationType", m_colComboValues, m_colComboIDS
    
    If m_bIsEdit = True Then
        ' DJP SQL Server port
        colValues.Add m_vDirectoryGUID
        TableAccess(m_clsValuationType).SetKeyMatchValues colValues
        TableAccess(m_clsValuationType).GetTableData
    Else
        TableAccess(m_clsValuationType).GetTableData POPULATE_EMPTY
    End If
    
    TableAccess(m_clsValuationType).ValidateData
    
    Set frmEditThirdParty.dgValuationType.DataSource = TableAccess(m_clsValuationType).GetRecordSet
    
    SetGridFields
    
    frmEditThirdParty.dgValuationType.Enabled = True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Initialises the datagrid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    
    Dim sGUID As String
    Dim bRet As Boolean
    Dim sCombo As String
    Dim fields As FieldData
    Dim colFields As New Collection
        
    On Error GoTo Failed
    
    bRet = True

    ' DJP SQL Server port
    sGUID = g_clsSQLAssistSP.GuidToString(CStr(m_vDirectoryGUID))
    
    ' Directory GUID
    fields.sField = "DirectoryGUID"
    fields.bRequired = True
    fields.bVisible = False
    ' DJP SQL Server port
    fields.sDefault = sGUID
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Valuation Type Text
    fields.sField = "ValuationTypeText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Valuation Type must be entered"
    fields.sTitle = "Valuation Type"
    fields.sOtherField = "ValuationType"

    If m_colComboValues.Count > 0 And m_colComboIDS.Count > 0 Then
        Set fields.colComboValues = m_colComboValues
        Set fields.colComboIDS = m_colComboIDS
        colFields.Add fields
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Valuer:SetGridFields, Unable to locate combo group" + sCombo
    End If
    
    ' Valuation Type
    fields.sField = "ValuationType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    frmEditThirdParty.dgValuationType.SetColumns colFields, "EditValuation", "Valuation"
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Loads all details from the tables into the screens controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    
    Dim vVal As Variant
    Dim vExpiry As Variant
    
    On Error GoTo Failed
    
    If m_bIsEdit = True Then
        ' Panel ID
        frmEditThirdParty.txtValuer(PANEL_ID).Text = m_clsPanel.GetPanelID()
        
        ' Associate Panel ID
        frmEditThirdParty.txtValuer(ASSOCIATE_PANEL_ID).Text = m_clsPanel.GetAssociatedPanelID()
        
        ' Payment Method
        vVal = m_clsPanel.GetPaymentMethod()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboValuerPaymentMethod, vVal, SET_CONTROL_VALUE
        
        'Valuer Type
        vVal = m_clsValuer.GetValuerType
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboValuerType, vVal, SET_CONTROL_VALUE
        
        ' Qualifications
        frmEditThirdParty.txtQualifications.Text = m_clsValuer.GetQualifications()
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all details from the relevant screen controls to table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim vVal As Variant
    Dim vExpiry As Variant
    
    On Error GoTo Failed
    
    ' Set the Directory GUID's
    m_clsPanel.SetDirectoryGUID m_vDirectoryGUID
    m_clsValuer.SetDirectoryGUID m_vDirectoryGUID
    
    ' Panel ID
    m_clsPanel.SetPanelID frmEditThirdParty.txtValuer(PANEL_ID).Text
    
    ' Associate Panel ID
    m_clsPanel.SetAssociatedPanelID frmEditThirdParty.txtValuer(ASSOCIATE_PANEL_ID).Text
    
    ' Payment Method
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboValuerPaymentMethod, vVal, GET_CONTROL_VALUE
     m_clsPanel.SetPaymentMethod vVal
    
    ' Qualifications
    m_clsValuer.SetQualifications frmEditThirdParty.txtQualifications.Text
    
    'ValuerType
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboValuerType, vVal, GET_CONTROL_VALUE
    m_clsValuer.SetValuerType vVal
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Update all underlying table objects handled by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    
    Dim clsTable As TableAccess
    
    On Error GoTo Failed

    'Update the Panel table.
    Set clsTable = m_clsPanel
    clsTable.Update

    'Update the PanelValuer table.
    Set clsTable = m_clsValuer
    clsTable.Update

    'Update the PanelValuerType table.
    Set clsTable = m_clsValuationType
    clsTable.Update

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates any control values relevant to this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = True
    bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsValuationType)
    
    If bRet = False Then
        frmEditThirdParty.SetActiveTab TabValuerDetails, True
        g_clsErrorHandling.RaiseError errGeneralError, "Valuation Type must be unique."
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Configure the screen controls to work in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Configure the screen controls to work in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    'Stub.
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelBankAccountDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PanelBankAccountDetails
' Description   : Contains all access routines for the PanelBankAccount table.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' STB       29/11/2001  SYS2912 - Removed the frmMainThirdParty screen and
'                       ported its functionality to frmEditThirdParty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MARS specific changes:
'
'GHun       29/07/2005  MARS27  Only allow one bank account to be created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Epsom specific changes:
' PB        10/05/2006  EP448   Legal rep bank details still mandatory - should not be
'--------------------------------------------------------------------------------------------

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'Current records key in the name and address directory table.
Private m_vDirectoryGUID As Variant

'Local table classes.
Private m_clsBankAccounts As PanelBankAccountTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if this tab-handler is in add or edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetDirectoryGUID(ByVal vGUID As Variant)
    m_vDirectoryGUID = vGUID
End Sub

Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Sub Class_Initialize()
    Set m_clsBankAccounts = New PanelBankAccountTable
End Sub

Public Sub Initialise(ByVal bIsEdit As Boolean)
        
    Dim sError As String
    Dim rs As ADODB.Recordset
    Dim vDirectoryGUID As Variant
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    sError = "Valuer:Initialise - "
    
    m_bIsEdit = bIsEdit
    
    'TODO: bring this called routine into this routine!
    SetupBankAccounts

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupBankAccounts()
    
    Dim rs As ADODB.Recordset
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsBankAccounts
    
    If m_bIsEdit = True Then
        ' Bank Accounts
        ' DJP SQL Server port
        ' Use the binary format, not string
        colValues.Add m_vDirectoryGUID
        clsTableAccess.SetKeyMatchValues colValues
        
        Set rs = clsTableAccess.GetTableData()
    Else
        Set rs = clsTableAccess.GetTableData(POPULATE_EMPTY)
    End If
    
    clsTableAccess.ValidateData
    
    Set frmEditThirdParty.dgBankAccounts.DataSource = rs
    
    SetGridFields
    
    frmEditThirdParty.dgBankAccounts.Enabled = True
    'MAR27 GHun Only allow 1 bank account to be added
    If rs.RecordCount = 0 Then
        frmEditThirdParty.dgBankAccounts.AllowAdd = True
    Else
        frmEditThirdParty.dgBankAccounts.AllowAdd = False
    End If
    'MAR27 End
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetGridFields()
    
    Dim bRet As Boolean
    Dim sGuid As String
    Dim sCombo As String
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboIDS As New Collection
    Dim colComboValues As New Collection
    
    On Error GoTo Failed
    
    bRet = True

    ' DJP SQL Server port
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vDirectoryGUID))
    
    ' Directory GUID
    fields.sField = "DirectoryGUID"
    fields.bRequired = True
    fields.bVisible = False
    ' DJP SQL Server port
    fields.sDefault = sGuid
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Bank Account Sequence Number
    fields.sField = "BankAccountSequenceNumber"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Bank Account Sequence Number must be entered"
    fields.sTitle = "Bank Account Sequence Number"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Bank Sort Code
    fields.sField = "BankSortCode"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Bank Sort Code must be entered"
    fields.sTitle = "Bank Sort Code"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Bank Name
    fields.sField = "BankName"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Bank Name must be entered"
    fields.sTitle = "Bank Name"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Account Name
    fields.sField = "AccountName"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Account Name must be entered"
    fields.sTitle = "Account Name"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Account Number
    fields.sField = "AccountNumber"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Account Number must be entered"
    fields.sTitle = "Account Number"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Account Currency
    fields.sField = "AccountCurrencyText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Account Currency must be entered"
    fields.sTitle = "Account Currency"
    g_clsCombo.FindComboGroup "AccountCurrency", colComboValues, colComboIDS
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    fields.sOtherField = "AccountCurrency"
    colFields.Add fields
    
    ' Account Currency
    fields.sField = "AccountCurrency"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields
        
    frmEditThirdParty.dgBankAccounts.SetColumns colFields, "EditBankAccounts", "Bank Accounts"
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetScreenFields()
    
    'This routine is stubbed in this class as it uses a bound datagrid.
    
End Sub

Public Sub SaveScreenData()
    'TODO: this may be 'stubbable' - but we might have to link the primary guid in...
End Sub

Public Sub DoUpdates()
    
    Dim clsTable As TableAccess
    Dim sSortCode As String
    Dim x, y, z As String
    
    On Error GoTo Failed

    'Update the PanelBankAccount table.
    Set clsTable = m_clsBankAccounts
        
    sSortCode = m_clsBankAccounts.GetBankSortCode
    
    ' EP448 PB
    If sSortCode = "" Then
        ' Do not save - no sort code
    Else
        If (Len(sSortCode) = 8) Then
            x = Mid(sSortCode, 3, 1)
            y = Mid(sSortCode, 5, 1)
            If x <> "-" And y <> "-" Then
                 g_clsErrorHandling.RaiseError errGeneralError, "Bank Acccount Sort code detail is incorrectly formatted."
            End If
        Else
            If (Len(sSortCode) = 6) Then
                x = Left(sSortCode, 2)
                y = Mid(sSortCode, 3, 2)
                z = Right(sSortCode, 2)
                sSortCode = x & "-" & y & "-" & z
                m_clsBankAccounts.SetBankSortCode sSortCode
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "Bank Acccount Sort code must be six characters"
            End If
        End If
        clsTable.Update
    End If
    ' EP448 End
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
        
    On Error GoTo Failed
    
    bRet = True
    bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsBankAccounts)

    If bRet = False Then
        frmEditThirdParty.SetActiveTab TabBankAccounts, True
        g_clsErrorHandling.RaiseError errGeneralError, "Bank Acccount Sequence Number has to be unique"
    End If
     
    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetAddState()
    'Stub.
End Sub

Public Sub SetEditState()
    'Stub.
End Sub


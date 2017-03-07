VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypicalAPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTypicalApr
' Description   : Provides functionality for the Typical APR tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
' JR        19/06/02    SYS4888 - Modify DoUpdates
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_colProductKeys As Collection
Private m_clsTypicalAPR As MortProdTypicalAPRTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    If m_clsTypicalAPR Is Nothing Then
        Set m_clsTypicalAPR = New MortProdTypicalAPRTable
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form to do initialisation of the Mortgage
'                 Product Typical APR's tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(Optional bIsEdit As Boolean)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim nPopulateType As PopulateType
    
    m_bIsEdit = bIsEdit
    Set clsTableAccess = m_clsTypicalAPR
    
    nPopulateType = POPULATE_EMPTY
    
    If m_bIsEdit Then
        nPopulateType = POPULATE_KEYS
        clsTableAccess.SetKeyMatchValues m_colProductKeys
    End If

    clsTableAccess.GetTableData nPopulateType
    clsTableAccess.ValidateData
    
    Set frmProductDetails.dgTypicalAPR.DataSource = clsTableAccess.GetRecordSet
    frmProductDetails.dgTypicalAPR.Enabled = True
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Initialises the datagrid, including field column names, whether they are mandatory
'                 etc.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection

    ' Mortgage Product Code
    fields.sField = "MortgageProductCode"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sProductCode
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Mortgage Product Start Date
    fields.sField = "StartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sProductStartDate
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Maximum Loan Amount
    fields.sField = "MaximumLoanAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Loan Amount must be entered"
    fields.sTitle = "Maximum Loan Amount"
    colFields.Add fields
    
    ' APR
    fields.sField = "APR"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "APR must be entered"
    fields.sTitle = "APR"
    colFields.Add fields
    
    frmProductDetails.dgTypicalAPR.SetColumns colFields, "TypicalAPR", "Typical APR's"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called externally, doesn't do anything for this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates the Typeical APRs table once all values have been written to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBanded As BandedTable
    Dim nRecordCount As Long
        
    nRecordCount = TableAccess(m_clsTypicalAPR).RecordCount
    If Not m_bIsEdit And nRecordCount > 0 Then
        Set clsBanded = m_clsTypicalAPR
        clsBanded.SetUpdateValues m_colProductKeys
        clsBanded.SetUpdateSets
        clsBanded.DoUpdateSets
    End If
    
    'JR SYS4888 - Remove check for recordcount and Update anyway
    'If nRecordCount > 0 Then
    TableAccess(m_clsTypicalAPR).Update
    'End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Called externally, updates the Typical APR table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    
    DoUpdates

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally when this tab is switched to.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Typical Apr tab - called externally.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = frmProductDetails.dgTypicalAPR.ValidateRows()
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTypicalAPR)
    
        If bRet = False Then
            frmProductDetails.dgTypicalAPR.SetGridFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Loan Amount must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "AdditionalMortProdParams.SetMortgageParameterKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

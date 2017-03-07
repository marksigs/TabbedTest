VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditInterestRateCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : EditInterestRateCS
' Description   : Client specific version of the EditInterestRate class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       13/12/01    SYS2831 Created for client variant support
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements EditInterestRate

' Private data
Private m_bIsEdit As Boolean
Private m_clsEditInterestRate As EditInterestRate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called this class is created. Need to create the EditInterestRate class first.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009. Amended :EditInterestRate_GetInterestRateTypeTable,EditInterestRate_SetEditState,EditInterestRate_SaveScreenData
'                                   Added :SaveScreenData,PopulateScreenFields
'BS         10/04/2003  BM0500 Removed duplicated call in EditInterestRate_SaveScreenData


'Private m_clsMortProdIntRates As MortProdIntRatesTable
'Private m_clsBaseRateSetTable As BaseRateSetTable
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    m_bIsEdit = False
    Set m_clsEditInterestRate = New EditInterestRate
    frmEditInterestRates.SetInterestRateClass Me
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function EditInterestRate_GetInterestRateTypeTable() As MortProdIntRatesTable
'GD        16/05/2002   BMIDS00009
    Set EditInterestRate_GetInterestRateTypeTable = m_clsEditInterestRate.GetInterestRateTable
End Function

Private Sub EditInterestRate_Show()
    On Error GoTo Failed
    
    ' Perform Interest Rate Load functionality first
    EditInterestRate_Load
    
    ' Now show the form
    m_clsEditInterestRate.Show
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub EditInterestRate_Load()
    On Error GoTo Failed
    
    ' Core Interest Rate load
    m_clsEditInterestRate.Load
    
    If m_bIsEdit = False Then
        EditInterestRate_SetAddState
    Else
        EditInterestRate_SetEditState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub EditInterestRate_SetAddState()
    On Error GoTo Failed
    
    m_clsEditInterestRate.SetAddState
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub EditInterestRate_SetEditState()
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed

    m_clsEditInterestRate.SetEditState
    'GD BMIDS
    PopulateScreenFields
    'END GD BMIDS
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub EditInterestRate_OK()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = EditInterestRate_DoOKProcessing()
    
    If bRet Then
        m_clsEditInterestRate.OK
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function EditInterestRate_DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = m_clsEditInterestRate.DoOKProcessing

    If bRet = True Then
        bRet = EditInterestRate_ValidateScreenData

        If bRet = True Then
            EditInterestRate_SaveScreenData
        End If
    End If

    EditInterestRate_DoOKProcessing = bRet

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub EditInterestRate_SaveScreenData()
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed

    
    m_clsEditInterestRate.SaveScreenData
    'GD BMIDS
    SaveScreenData
    'END GD BMIDS
    'BS BM0500 10/04/2003
    'SaveScreenData
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub EditInterestRate_Another()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = EditInterestRate_DoOKProcessing
    
    If bRet Then
        m_clsEditInterestRate.Another
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function EditInterestRate_ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bSuccess As Boolean
    
    bSuccess = m_clsEditInterestRate.ValidateScreenData
    
    EditInterestRate_ValidateScreenData = bSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub EditInterestRate_Cancel()
    On Error GoTo Failed
    
    m_clsEditInterestRate.Cancel
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub EditInterestRate_HandleType()
    m_clsEditInterestRate.HandleType
End Sub

Private Function EditInterestRate_IsEdit() As Boolean
    EditInterestRate_IsEdit = m_clsEditInterestRate.IsEdit
End Function

Private Function EditInterestRate_PopulateScreenFields() As Boolean
    EditInterestRate_PopulateScreenFields = m_clsEditInterestRate.PopulateScreenFields
End Function

Private Sub EditInterestRate_SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
    m_clsEditInterestRate.SetIsEdit bIsEdit
End Sub

Private Sub EditInterestRate_SetReturnCode(Optional enumReturn As MSGReturnCode = 1&)
    m_clsEditInterestRate.SetReturnCode enumReturn
End Sub

Private Function EditInterestRate_GetReturnCode() As MSGReturnCode
    EditInterestRate_GetReturnCode = m_clsEditInterestRate.GetReturnCode
End Function

Private Sub EditInterestRate_SetTableClass(clsTable As TableAccess)
    m_clsEditInterestRate.SetTableClass clsTable
End Sub

Private Sub EditInterestRate_UnloadForm()
    m_clsEditInterestRate.UnloadForm
End Sub

Private Function EditInterestRate_GetInterestRateTable() As MortProdIntRatesTable
    Set EditInterestRate_GetInterestRateTable = m_clsEditInterestRate.GetInterestRateTable
End Function
Private Sub SaveScreenData()
'GD        16/05/2002   BMIDS00009
    Dim rs As ADODB.Recordset
    Dim clsBaseRateSet As New BaseRateSetTable
    Dim clsTableAccess As TableAccess
    Dim iSelectedIndex As Integer
    Dim clsInterestRates As New MortProdIntRatesTable
    
    clsBaseRateSet.GetBaseRateSet
    'Set rs = clsTableAccess.GetRecordSet
    Set rs = TableAccess(clsBaseRateSet).GetRecordSet
    iSelectedIndex = frmEditInterestRates.cboBaseRateSet.SelectedItem
    rs.Move (iSelectedIndex - 1)
    '
    
    Set clsInterestRates = m_clsEditInterestRate.GetInterestRateTable
    clsInterestRates.SetBaseRateSet (clsBaseRateSet.GetFeeSet)
    

End Sub
Private Sub PopulateScreenFields()
'GD        16/05/2002   BMIDS00009
    Dim clsInterestRates As New MortProdIntRatesTable
    Dim clsBaseRateSetTable As BaseRateSetTable
    Dim vTmp As Variant
    Dim vTmp2 As Variant
    Dim colKeys As New Collection
    Set clsInterestRates = m_clsEditInterestRate.GetInterestRateTable
    vTmp = clsInterestRates.GetBaseRateSet()
    Set clsBaseRateSetTable = New BaseRateSetTable
    colKeys.Add vTmp
    TableAccess(clsBaseRateSetTable).SetKeyMatchValues colKeys
    TableAccess(clsBaseRateSetTable).GetTableData
    vTmp2 = clsBaseRateSetTable.GetTheBaseRateSet & " - " & clsBaseRateSetTable.GetDescription
    frmEditInterestRates.cboBaseRateSet.Text = vTmp2
End Sub


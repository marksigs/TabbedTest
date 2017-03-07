VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntMortProdProcuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : IntermediaryMortgageProdProcuration
' Description   : Procuration mortgage type tab
'
' Change history
' Prog      Date        Description
' AA        26/06/2001  created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsIntProcFee As IntProcFeeSplitTable
Private m_clsProcSplit As IntProcFeeSplitForIntTable
Private m_clsProcFee As IntermediaryProcFeeTable
Private m_nFeeValueID As Integer
Private m_vIntermediaryGUID As Variant
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Public Sub Initialise(bIsEdit As Boolean, Optional bGetData As Boolean = True)
    On Error GoTo Failed
    
    Set m_clsProcFee = New IntermediaryProcFeeTable
    m_bIsEdit = bIsEdit
    m_vIntermediaryGUID = m_colKeys(1)
    
    If bGetData Then
        GetSplitDetails
    End If
    SetListViewHeaders
    PopulateScreenControls
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetKeys(col As Collection)
    Set m_colKeys = col
End Sub

Public Function GetProcTypeID() As Variant

    Dim clsComboValidation As ComboValidationTable
    Dim nTypeID As Variant
    
    Set clsComboValidation = New ComboValidationTable
    
    nTypeID = clsComboValidation.GetSingleValueFromValidation("ProcurationFeeType", PROCURATION_FEE_NON_PRODUCT_SPECIFIC)
    
    GetProcTypeID = nTypeID
End Function

Public Sub PopulateScreenControls()
    On Error GoTo Failed
    
    Dim clsLender As MortgageLendersTable
    Dim clsIntermediary As Intermediary
    Dim rs As ADODB.Recordset
    Dim sComboText As String
    Dim sGuid As String
    'Product period Combo
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    Set clsLender = New MortgageLendersTable
    'Get all Records for intermediary and type
    m_clsProcFee.GetFeeTypesForIntermediary sGuid, CStr(m_nFeeValueID)

    
    Set clsIntermediary = New Intermediary
    
    'Populate product periods
    clsIntermediary.PopulateProcurationCombo frmEditProcurationFees.cboMortProdPeriod, m_clsProcFee, m_bIsEdit
    
    'Populate Lenders
    clsLender.GetLenders "LenderName"
    Set rs = TableAccess(clsLender).GetRecordSet
    
    
    With frmEditProcurationFees.cboMortLender
        Set .RowSource = rs
        .ListField = "LenderName"
        .BoundColumn = "LenderCode"
    End With
        
    If frmEditProcurationFees.cboMortProdPeriod.RowSource.RecordCount > 0 Then
        sComboText = frmEditProcurationFees.cboMortProdPeriod.RowSource(frmEditProcurationFees.cboMortProdPeriod.ListField)
        
        frmEditProcurationFees.cboMortProdPeriod.Text = ""
        
        frmEditProcurationFees.cboMortProdPeriod.Text = sComboText
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetIntermediaryGuid(vGuid As Variant)
    m_vIntermediaryGUID = vGuid
End Sub

Public Sub SetFeeType(nType As Integer)
    m_nFeeValueID = nType
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleLenderSelect
' Description   : A lender has been selected, so all products for the selected lender must be populated
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleLenderSelect(sLenderID As String)
    On Error GoTo Failed
    
    Dim clsProducts As MortgageProductTable

    Set clsProducts = New MortgageProductTable
    
    clsProducts.GetProducts SearchLender, sLenderID
    
    With frmEditProcurationFees.cboMortProduct
        .Text = ""
        Set .RowSource = TableAccess(clsProducts).GetRecordSet
        .ListField = "ProductName"
        .BoundColumn = "MORTGAGEPRODUCTCODE"
    End With
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetListViewHeaders
' Description   :   Sets the field header titles on the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "£ Range"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Fee"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "HO Split"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Company Split"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Individual Split"
    headers.Add lvHeaders
    
    frmEditProcurationFees.lvMortgageSpecific.AddHeadings headers
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSplitTableClass
' Description   : Sets the proc split table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSplitTableClass(clsSplit As IntProcFeeSplitTable)
    On Error GoTo Failed
    
    Set m_clsIntProcFee = clsSplit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the listview with all records from the proc fee split table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed

    Dim clsProcFee As IntProcFeeSplitTable
    Dim sGuid As String
    
    Set clsProcFee = New IntProcFeeSplitTable
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    clsProcFee.GetAllFeeSplitsForGuid sGuid, CStr(m_nFeeValueID)
    
    g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvMortgageSpecific, clsProcFee, True

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleProductPeriod
' Description   : Called when the insurance period is changed. Updates the LV to only show fees for that period
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleProductPeriod(Optional bGetData As Boolean = True)
    On Error GoTo Failed

    Dim sSequence As String
    Dim sProduct As String
    Dim sLender As String
    Dim clsFeeType As IntermediaryProcFeeTable
    Dim clsLender As MortgageLendersTable
    Dim col As Collection
    Dim clsProduct As MortgageProductTable
    
    sSequence = frmEditProcurationFees.cboMortProdPeriod.BoundText
    
    If Len(sSequence) > 0 Then
        TableAccess(m_clsIntProcFee).ApplyFilter "TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSequence)
        Set clsFeeType = New IntermediaryProcFeeTable
        
        TableAccess(clsFeeType).SetRecordSet frmEditProcurationFees.cboNonProdSpecificPeriod.RowSource
        
        If TableAccess(m_clsIntProcFee).RecordCount > 0 Then
            Set col = New Collection
            Set clsLender = New MortgageLendersTable
            Set clsProduct = New MortgageProductTable
                    
            col.Add clsFeeType.GetOrganisationID
            
            TableAccess(clsLender).GetTableData
            
            sLender = clsLender.GetLenderName
            frmEditProcurationFees.cboMortLender.Text = sLender
            
            sProduct = clsFeeType.GetProductID
            
            If Len(sProduct) > 0 Then
                'clsProduct.GetProducts SearchActive, , , sProduct
                clsProduct.GetProductByID sProduct
                If TableAccess(clsProduct).RecordCount > 0 Then
                    sProduct = clsProduct.GetProductName
                    frmEditProcurationFees.cboMortProduct.Text = sProduct
                End If
            End If
            g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvMortgageSpecific, m_clsIntProcFee, True
        End If
        TableAccess(m_clsIntProcFee).CancelFilter
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetSplitTableClass() As IntProcFeeSplitTable
    Set GetSplitTableClass = m_clsIntProcFee
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetSplitDetails
' Description   : Populates the split details
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSplitDetails()
    On Error GoTo Failed
    Dim sGuid As String
    
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    m_clsIntProcFee.GetAllFeeSplitsForGuid sGuid, CStr(m_nFeeValueID)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveProduct
' Description   : Saves the proc fee product to the type record
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveProduct()
    On Error GoTo Failed

    Dim vVal As Variant
    Dim clsFeeType As IntermediaryProcFeeTable
    Dim sSequence As String
    Dim sProd As String
    Dim vLender As Variant
    Dim sLender As String
    Dim clsProd As MortgageProductTable
    
    Set clsProd = New MortgageProductTable
    Set clsFeeType = New IntermediaryProcFeeTable

    sSequence = frmEditProcurationFees.cboMortProdPeriod.BoundText

    TableAccess(clsFeeType).SetRecordSet frmEditProcurationFees.cboMortProdPeriod.RowSource
    
    If Len(sSequence) > 0 Then
        TableAccess(clsFeeType).ApplyFilter "TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSequence)

        sProd = frmEditProcurationFees.cboMortProduct.BoundText

        If TableAccess(clsFeeType).RecordCount > 0 And Len(sProd) > 0 Then
            clsFeeType.SetProductID sProd
            TableAccess(clsProd).SetRecordSet frmEditProcurationFees.cboMortProduct.RowSource
            vLender = clsProd.GetOrganisationID
            
            sLender = g_clsSQLAssistSP.GuidToString(CStr(vLender))
            clsFeeType.SetOrganisationId sLender
            
            TableAccess(clsFeeType).Update
        End If


        TableAccess(clsFeeType).CancelFilter
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

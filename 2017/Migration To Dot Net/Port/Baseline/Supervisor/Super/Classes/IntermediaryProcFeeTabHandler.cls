VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntProcFeeTabHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : IntProcFeeTabHandler
' Description   : A generic tab-handler for procuration fee tabs. The fee-type
'                 (Insurance, Mortgage, Packaging or All) when set, will
'                 govern a filter used by this tab-handler before it
'                 manipulates any table objects.
'
' This tab-handler is fairly non-conventional, in-that, it uses filters on the
' table objects (based on fee-type) prior working with them and also provides
' a substantial more GUI helper routines. Most 'standard' routines are stubbed.
'
' Change history
' Prog      Date        Description
' STB       11/12/01    Class created.
' STB       22/04/02    SYS4401 Don't associate with a product if the fee type
'                       is packaging fee.
' STB       22/04/02    SYS4400 Display either the total amonut or the total
'                       percentage in the Fee column.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The number of intermediary split records.
Private Const NUMBER_SPLITS As Long = 3

'A collection of primary keys to identify the current intermediary record.
Private m_colKeys As Collection

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'This table holds periods.
Private m_clsProcFeeTypeTable As IntermediaryProcFeeTable

'This table holds fee-bands.
Private m_clsProcFeeSplitTable As IntProcFeeSplitTable

'This table holds band-splits.
Private m_clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable

'Store the fee-type this tab-handler is responsible for. This also corresponds
'to the tab-index on the screen (and hence all control indexes) as well as the
'fee-type ValueID. Most fortuanate.
Private m_uFeeType As ProcFeeTypeEnum

'Store a reference to the opening form.
Private m_frmProcFee As frmEditProcurationFees


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetForm
' Description   : Store a reference to the opener form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetForm(ByRef frmEditProcFees As frmEditProcurationFees)
    Set m_frmProcFee = frmEditProcFees
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Set the fee-type which this tab-handler will be responsible
'                 for.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeType(ByVal uProcFeeTypeEnum As ProcFeeTypeEnum)
    m_uFeeType = uProcFeeTypeEnum
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeTypeTable
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeTypeTable(ByRef clsProcFeeTypeTable As IntermediaryProcFeeTable)
    Set m_clsProcFeeTypeTable = clsProcFeeTypeTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplitTable
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplitTable(ByRef clsProcFeeSplitTable As IntProcFeeSplitTable)
    Set m_clsProcFeeSplitTable = clsProcFeeSplitTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplitByIntTable
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplitByIntTable(ByRef clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable)
    Set m_clsProcFeeSplitByIntTable = clsProcFeeSplitByIntTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated
'                 accoding to the underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)

    'Store the add/edit state.
    m_bIsEdit = bIsEdit

    'Perform state-related table operations.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Setup the columns for the fee-band lists.
    SetupListViewHeaders
    
    'Populate the combos this tab-handler is responsible for.
    PopulateScreenControls

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Add the column colHeaders to the listview serviced by this
'                 tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders()

    Dim colHeaders As Collection
    Dim lvHeaders As listViewAccess
    
    'Create a collection to hold column header objects.
    Set colHeaders = New Collection
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "£ Range"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Fee"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "HO Split"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Company Split"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Individual Split"
    colHeaders.Add lvHeaders
    
    'Associate the column headers collection with the listview.
    m_frmProcFee.lvFeeBandList(m_uFeeType).AddHeadings colHeaders

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate any combos or static lists from the table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenControls()
    
    'Stop a _Click() recursive event from taking place.
    m_frmProcFee.BlockComboRecursion True
    
    Select Case m_uFeeType
        Case MortgageFee
            'If this is a Mortgage type, populate the lender combo. It will
            'cascade and cause the product then period combos to be populated.
            PopulateLenderCombo
            PopulateProductCombo
            PopulatePeriodCombo
            PopulateFeeBandList
                    
        Case InsuranceFee
            'If this is an Insurance or Package type, populate the product
            'combo. It will cascade and cause the period combo to be populated.
            PopulateProductCombo
            PopulatePeriodCombo
            PopulateFeeBandList
                
        Case PackagingFee, NonSpecificFee
            'Otherwise, if this is a non-specific (All) type, just populate the
            'period combo. This will cause a cascaded refresh of the listview of
            'fee-bands.
            PopulatePeriodCombo
            PopulateFeeBandList
            
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "An invalid Procuration Fee Type was detected in IntermediaryProcFeeTabHandler"
    End Select
    
    'Allow combo _Click events to be processed.
    m_frmProcFee.BlockComboRecursion False

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateLenderCombo
' Description   : Populates the relevant lender combo control with all Mortgage
'                 Lenders. This will cause a cascaded population of the product
'                 combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateLenderCombo()
    
    Dim colFields As Collection
    Dim clsLenderTable As MortgageLendersTable
    
    'Create a lenders table object.
    Set clsLenderTable = New MortgageLendersTable
    
    'Get a sorted list of lenders.
    clsLenderTable.GetLenders "LenderName"
    
    'Build a collection of fieldnames to 'bind' the combo to.
    Set colFields = New Collection
    colFields.Add "LenderName"
    colFields.Add "LenderCode"
    
    'Populate the lenders combo from the table object.
    g_clsFormProcessing.PopulateComboFromTable m_frmProcFee.cboMortgageLender(m_uFeeType), TableAccess(clsLenderTable), colFields
    
    'Select <select>.
    g_clsFormProcessing.HandleComboText m_frmProcFee.cboMortgageLender(m_uFeeType), CStr(COMBO_NONE), SET_CONTROL_VALUE
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateProductCombo
' Description   : Populate the relevant product combo control. If required,
'                 using the values selected in the Organisation/Lender combo to
'                 filter the data shown. This will cause a cascaded population
'                 of the period combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateProductCombo()
    
    Dim sLenderCode As String
    Dim colFields As Collection
    Dim colValueIDs As Collection
    Dim clsProductTable As MortgageProductTable
    
    Select Case m_uFeeType
        Case MortgageFee
            'Create a mortgage product table object.
            Set clsProductTable = New MortgageProductTable
            
            'Read the LenderCode currently selected.
            g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboMortgageLender(m_uFeeType), sLenderCode, GET_CONTROL_VALUE
    
            'Get a list of products for the currently selected lender.
            clsProductTable.GetProducts SearchLender, sLenderCode
            
            'If there are no products, clear the combo.
            If TableAccess(clsProductTable).RecordCount = 0 Then
                'Populate the empty combo with only <select> selected.
                Set colFields = New Collection
                Set colValueIDs = New Collection
                
                colFields.Add COMBO_NONE
                colValueIDs.Add colFields.Count
                
                m_frmProcFee.cboProduct(m_uFeeType).SetListTextFromCollection colFields, colValueIDs
            Else
                'Build a collection of fieldnames to 'bind' the combo to.
                Set colFields = New Collection
                colFields.Add "ProductName"
                colFields.Add "MORTGAGEPRODUCTCODE"
                
                'Populate the products combo from the table object.
                g_clsFormProcessing.PopulateComboFromTable m_frmProcFee.cboProduct(m_uFeeType), TableAccess(clsProductTable), colFields
            End If
                
        Case InsuranceFee
            'Populate the insurance products from a combo group.
            g_clsFormProcessing.PopulateCombo "InsuranceType", m_frmProcFee.cboProduct(m_uFeeType)
        
        Case Else
            'Get no products.
            
    End Select
    
    'Select <select>.
    g_clsFormProcessing.HandleComboText m_frmProcFee.cboProduct(m_uFeeType), CStr(COMBO_NONE), SET_CONTROL_VALUE

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulatePeriodCombo
' Description   : Populate the relevant period combo control. If required,
'                 using the values selected in the Organisation/Lender and
'                 Product combos to filter the data shown. This will cause a
'                 cascaded populate of the fee/band listview control.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulatePeriodCombo()

    Dim sGuid As String
    Dim sFilter As String
    Dim sFeeType As String
    Dim sProductID As String
    Dim sLenderCode As String
    Dim sOrganisationID As String
    Dim colTextList As Collection
    Dim rsPeriods As ADODB.Recordset
    Dim colValueIDList As Collection
    
    'Get a local (more readable version of the GUID).
    sGuid = g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    
    'Get a local (more readable version of the Type).
    sFeeType = g_clsSQLAssistSP.FormatString(CStr(m_uFeeType))
    
    'Build a filter string based upon the fee-type.
    Select Case m_uFeeType
        Case NonSpecificFee, PackagingFee
            'The period combo should be filtered on GUID and Fee-Type.
            sFilter = GetPeriodFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, "")
    
        Case InsuranceFee
            'Get a local (more readable version of the ProductID).
            g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboProduct(m_uFeeType), sProductID, GET_CONTROL_VALUE
            sProductID = g_clsSQLAssistSP.FormatString(sProductID)
            
            'The period combo should be filtered on GUID, Fee-Type and Product ID.
            sFilter = GetPeriodFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, "")
            sFilter = sFilter & " AND ProductID = " & sProductID
            
        Case MortgageFee
            'Get a local (more readable version of the ProductID).
            g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboProduct(m_uFeeType), sProductID, GET_CONTROL_VALUE
            sProductID = g_clsSQLAssistSP.FormatString(sProductID)
            
            'Get a local (more readable version of the LenderCode).
            g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboMortgageLender(m_uFeeType), sLenderCode, GET_CONTROL_VALUE
            
            If sLenderCode = "" Then
                sOrganisationID = g_clsSQLAssistSP.FormatString("")
            Else
                sOrganisationID = g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(GetOrgIDFromLenderCode(sLenderCode))))
            End If
            
            'The period combo should be filtered on GUID, Fee-Type, Product ID and Lender ID.
            sFilter = GetPeriodFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, "")
            sFilter = sFilter & " AND ProductID = " & sProductID & " AND OrganisationID = " & sOrganisationID
    End Select
    
    'Apply the filter to restrict the data available.
    TableAccess(m_clsProcFeeTypeTable).ApplyFilter sFilter
    
    'Because the product strings must be in the format "dd/mm/yyyy - dd/mm/yyyy"
    '(i.e. an combination of field values), we must use a specialised routine
    'in the period table, to populate the combo control - not nice but its there!
    
    'Create two collections to hold the text-captions and the underlying value id's.
    Set colTextList = New Collection
    Set colValueIDList = New Collection

    'Get a reference to the underlying table object.
    Set rsPeriods = TableAccess(m_clsProcFeeTypeTable).GetRecordSet()
    
    'Ensure there is data present.
    If Not rsPeriods Is Nothing Then
        If rsPeriods.RecordCount > 0 Then
            'Move to the first record.
            rsPeriods.MoveFirst
            
            Do While Not rsPeriods.EOF
                'Add the necassary data into the collection(s).
                colValueIDList.Add CStr(rsPeriods.fields("TypeSequenceNumber").Value)
                
                'The text collection is the Start Date AND end date.
                colTextList.Add CVar(Format(rsPeriods.fields("ProcurationFeeTypeActiveFrom").Value, "dd/mm/yyyy") & " - " & Format(rsPeriods.fields("ProcurationFeeTypeActiveTo").Value, "dd/mm/yyyy"))
                
                'Move onto the next record.
                rsPeriods.MoveNext
            Loop
        End If
    End If

    colTextList.Add COMBO_NONE
    colValueIDList.Add colTextList.Count
    
    'Populate the combo from the generated collections.
    m_frmProcFee.cboPeriod(m_uFeeType).SetListTextFromCollection colTextList, colValueIDList
   
    'Select <select>.
    g_clsFormProcessing.HandleComboText m_frmProcFee.cboPeriod(m_uFeeType), CStr(COMBO_NONE), SET_CONTROL_VALUE

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateFeeBandList
' Description   : Populate the relevant listview control with fee-band records,
'                 based on the selected value of the relevant period combo.
'                 Each row corresponds to a row from the bands table and all
'                 its corresponding rows in the Split table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateFeeBandList()
    
    Dim sFilter As String
    Dim lRecordIdx As Long
    Dim colFeeLine As Collection
    Dim vTypeSequence As Variant
    Dim clsPopulateDetails As PopulateDetails
    
    'Get the Type/Period sequence number.
    g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboPeriod(m_uFeeType), vTypeSequence, GET_CONTROL_VALUE
    
    'If there is no period selected in the combo, then filter everything out
    'with an impossible value. Filter the bands on GUID, Type, TypeSequence.
    If CStr(vTypeSequence) = "" Then
        sFilter = GetPeriodFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, "-1")
    Else
        sFilter = GetPeriodFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, CStr(vTypeSequence))
    End If
    
    'Apply the filter to the bands table.
    TableAccess(m_clsProcFeeSplitTable).ApplyFilter sFilter

    'Note: We can't use the generic listview population routine as we have
    'a client-side cached parent/child table hierachry at this point. Therefore
    'we must do some extra work in this routine to build a row-collection and
    'populate-details collection object for each listview row.
    
    'Clear the listview.
    m_frmProcFee.lvFeeBandList(m_uFeeType).ListItems.Clear
    
    If TableAccess(m_clsProcFeeSplitTable).RecordCount > 0 Then
        'Ensure we're at the begining of the band table.
        TableAccess(m_clsProcFeeSplitTable).MoveFirst
        
        'Now populate the relevant listview from each record in this table object.
        For lRecordIdx = 1 To TableAccess(m_clsProcFeeSplitTable).RecordCount
            'Create a collection to hold row values.
            Set colFeeLine = New Collection
        
            'Low and High Band.
            colFeeLine.Add m_clsProcFeeSplitTable.GetLowBand & " - " & m_clsProcFeeSplitTable.GetHighBand
            
            'Total Amount / or Total Percentage.
            If m_clsProcFeeSplitTable.GetTotalAmount = "" Then
                colFeeLine.Add m_clsProcFeeSplitTable.GetTotalPercent
            Else
                colFeeLine.Add m_clsProcFeeSplitTable.GetTotalAmount
            End If
        
            'Get the LeadAgent split percentage.
            colFeeLine.Add GetSplitPercent(sFilter, m_clsProcFeeSplitTable.GetSplitSequenceNumber, LeadAgentType)
            
            'Get the Company split percentage.
            colFeeLine.Add GetSplitPercent(sFilter, m_clsProcFeeSplitTable.GetSplitSequenceNumber, CompanyType)
            
            'Get the Individual split percentage.
            colFeeLine.Add GetSplitPercent(sFilter, m_clsProcFeeSplitTable.GetSplitSequenceNumber, IndividualType)
        
            'Construct a 'populate details' object.
            Set clsPopulateDetails = New PopulateDetails
            clsPopulateDetails.SetKeyMatchValues TableAccess(m_clsProcFeeSplitTable).GetKeyMatchValues
            clsPopulateDetails.SetExtra m_clsProcFeeSplitTable.GetSplitSequenceNumber
            clsPopulateDetails.SetObjectDescription ""
        
            'Add the row values/collection into the listview control.
            m_frmProcFee.lvFeeBandList(m_uFeeType).AddLine colFeeLine, clsPopulateDetails
            
            'Move onto the next record.
            TableAccess(m_clsProcFeeSplitTable).MoveNext
        Next lRecordIdx
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DeleteFee
' Description   : Delete the selected record from the underlying table object
'                 and listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteFee()
    
    Dim sFilter As String
    Dim sTypeSequence As String
    Dim clsPopulateDetails As PopulateDetails
    
    'Get a reference to the selected item.
    Set clsPopulateDetails = m_frmProcFee.lvFeeBandList(m_uFeeType).SelectedItem.Tag
    
    'Ascertain what the currently selected TypeSequence (period) is.
    g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboPeriod(m_uFeeType), sTypeSequence, GET_CONTROL_VALUE
    
    'Build a filter to find the selected band and all related splits
    '(we can use the same filter!).
    sFilter = GetBandFilter(m_colKeys(INTERMEDIARY_KEY), m_uFeeType, sTypeSequence, CStr(clsPopulateDetails.GetExtra()))
    
    'Apply the filter to both the band and the split table.
    TableAccess(m_clsProcFeeSplitTable).ApplyFilter sFilter
    TableAccess(m_clsProcFeeSplitByIntTable).ApplyFilter sFilter
    
    'Ensure the correct number of records are left.
    If (TableAccess(m_clsProcFeeSplitTable).RecordCount = 1) And (TableAccess(m_clsProcFeeSplitByIntTable).RecordCount = NUMBER_SPLITS) Then
        'Delete those records.
        TableAccess(m_clsProcFeeSplitTable).DeleteAllRows
        TableAccess(m_clsProcFeeSplitByIntTable).DeleteAllRows
               
        'Note: Have to update the tables now. This means that we cannot delete
        'a band (and corresponding splits) if we are currently adding an intermediary
        'record. It also means a lock may be put on the database at this point until
        'the main intermediary screen is closed. Unfortuanately this is a feature of
        'using filters and batchupdates. It seems that if UpdateBatch is called on
        'a recordset where some filtered records have been deleted, an error is
        'thrown...
        TableAccess(m_clsProcFeeSplitByIntTable).Update
        TableAccess(m_clsProcFeeSplitTable).Update
        
        'Remove the line from the listview.
        m_frmProcFee.lvFeeBandList(m_uFeeType).RemoveLine m_frmProcFee.lvFeeBandList(m_uFeeType).SelectedItem.Index
        
        'Selection state may have altered, so reset the button state.
        SetFeeButtonState
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot delete the procuration fee and related splits as an incorrect number of records was returned from the filter"
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeButtonState
' Description   : Common routine to enable/disable the edit and delete fee buttons.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFeeButtonState()
    
    Dim vTmp As Variant

    'Read the period id.
    g_clsFormProcessing.HandleComboExtra m_frmProcFee.cboPeriod(m_uFeeType), vTmp, GET_CONTROL_VALUE
        
    'Only enable ADD if a period exists.
    m_frmProcFee.cmdAddFee(m_uFeeType).Enabled = (CStr(vTmp) <> "")

    'Only enable EDIT if an item is selected and a period exists.
    m_frmProcFee.cmdEditFee(m_uFeeType).Enabled = (Not m_frmProcFee.lvFeeBandList(m_uFeeType).SelectedItem Is Nothing)
    
    'Only enable DELETE if an item is selected and we're not adding a new record.
    'This is because deleting calls a DoUpdate prematurely.
    m_frmProcFee.cmdDeleteFee(m_uFeeType).Enabled = (Not m_frmProcFee.lvFeeBandList(m_uFeeType).SelectedItem Is Nothing) And (m_bIsEdit)
    
    'Also refresh the period add/edit button states.
    m_frmProcFee.cmdEditPeriod(m_uFeeType).Enabled = (CStr(vTmp) <> "")

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Terminate
' Description   : Release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Terminate()
    
    Set m_colKeys = Nothing
    Set m_frmProcFee = Nothing
    Set m_clsProcFeeTypeTable = Nothing
    Set m_clsProcFeeSplitTable = Nothing
    Set m_clsProcFeeSplitByIntTable = Nothing

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Instantiate any required table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Terminate
' Description   : Release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Terminate
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Ensure the controls this tab-handler is responsible for contain valid data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    'Stub
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Broker the call onto each tab-handler. Saving its data to the underlying
'                 table object(s).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup tables with new records and GUIDs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup tables and load data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOrgIDFromLenderCode
' Description   : From the specified LenderCode, return an OrganisationID. We must do this as
'                 the lender combo uses a lender code, whilst the proc fee records use an
'                 organisation id. It would be nice if the combos could use GUIDs...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetOrgIDFromLenderCode(ByVal sLenderCode As String) As Variant

    Dim vOrganisationID As Variant
    Dim clsLenderTable As MortgageLendersTable

    'If no LenderCode was passed in, then simply return a blank lender code.
    If sLenderCode = "" Then
        vOrganisationID = Empty
    Else
        'Create a lender table.
        Set clsLenderTable = New MortgageLendersTable

        'Populate the table object with ALL lenders.
        clsLenderTable.GetLenders

        'Get the desired record.
        clsLenderTable.FindLender sLenderCode

        'Ensure a single record was returned.
        If TableAccess(clsLenderTable).RecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errRecordNotFound, "Whilst trying to obtain an OrganisationID from a LenderCode: could not find a lender from the specified lender code - or too many records were returned (" & sLenderCode & ")"
        End If
        
        vOrganisationID = clsLenderTable.GetOrganisationID
    End If

    'Return the lender code to the caller.
    GetOrgIDFromLenderCode = vOrganisationID
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetSplitPercent
' Description   : Return the split amount for the specified intermediary type. This routine
'                 also accepts a filter to use on the Split table, it must simply add the
'                 intermediary type into the filter prior to using it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSplitPercent(ByVal sBandFilter As String, ByVal sSplitSequence As String, ByVal uIntermediaryType As IntermediaryTypeEnum) As String

    Dim sFilter As String
    Dim vTypeSequence As Variant
    
    'Build on-top of the specified filter.
    sFilter = sBandFilter & " AND SplitSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSplitSequence)
    sFilter = sFilter & " AND SplitForIntermediaryType = " & g_clsSQLAssistSP.FormatString(CStr(uIntermediaryType))
    
    'Apply the filter to the bands table.
    TableAccess(m_clsProcFeeSplitByIntTable).ApplyFilter sFilter
    
    'Ensure we have a single record.
    If TableAccess(m_clsProcFeeSplitByIntTable).RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "GetSplitPercent couldn't filter down to one record!"
    End If
    
    'Return the percentage amount.
    GetSplitPercent = m_clsProcFeeSplitByIntTable.GetSplitPercentage

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetPeriodFilter
' Description   : Build a recordset filter to be applied to the period table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPeriodFilter(ByVal vIntermediaryGUID As Variant, ByVal uFeeType As ProcFeeTypeEnum, ByVal sTypeSequence As String) As String

    Dim sFilter As String

    'IntermediaryGUID.
    sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vIntermediaryGUID)))
    
    'Type (proc fee).
    sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(uFeeType)
    
    'TypeSequence.
    If Len(sTypeSequence) = 0 Then
        'If no type sequence (PeriodID essentially) is specified, then don't
        'filter on that field.
    Else
        sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sTypeSequence)
    End If
    
    GetPeriodFilter = sFilter

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetBandFilter
' Description   : Build a recordset filter to be applied to the bands table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBandFilter(ByVal vIntermediaryGUID As Variant, ByVal uFeeType As ProcFeeTypeEnum, ByVal sTypeSequence As String, ByVal sSplitSequence As String) As String

    Dim sFilter As String

    'Start with a period filter.
    sFilter = GetPeriodFilter(vIntermediaryGUID, uFeeType, sTypeSequence)
    
    'SplitSequence.
    sFilter = sFilter & " AND SplitSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSplitSequence)
    
    GetBandFilter = sFilter
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetSplitFilter
' Description   : Build a recordset filter to be applied to the split table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSplitFilter(ByVal vIntermediaryGUID As Variant, ByVal uFeeType As ProcFeeTypeEnum, ByVal sTypeSequence As String, ByVal sSplitSequence As String, ByVal uIntermediaryType As IntermediaryTypeEnum) As String

    Dim sFilter As String
    
    'Start with a band filter.
    sFilter = GetBandFilter(vIntermediaryGUID, uFeeType, sTypeSequence, sSplitSequence)

    'SplitForIntermediaryType.
    sFilter = sFilter & " AND SplitForIntermediaryType = " & g_clsSQLAssistSP.FormatString(uIntermediaryType)

    GetSplitFilter = sFilter

End Function


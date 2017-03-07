VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageProduct
' Description   : Controlling class for the MortgageProduct form. This is the Core class which contains
'                 Core processing only. The client specific class MortgageProductCS will delegate to this
'                 class for Core functionality.
'
' Change history
' Prog      Date        Description
' DJP       10/12/01    SYS2831 Created for Client variants.
' DJP       07/01/02    SYS2831 Save Sets on Add and Edit, and set tabstops for the active tab.
' STB       30/01/02    SYS3771 Removing ProductGroupingTable.
' STB       02/05/02    SYS4496 Added method ShowInvalidTab called from ValidateScreenData.
' MC        25/05/2004  BMIDS775 - CC071 Additional Features dialog event
' MC        04/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
' MC        21/06/2004  BMIDS775 - Additional Features funtional bug fix.
' TW        09/10/2006  EP2_7 - Added handling for Additional Borrowing Fee and Credit Limit Increase Fee
' TW        30/11/2006  EP2_253 - Changes related to Mortgage Product Application Eligibility
' TW        11/12/2006  EP2_20 - Added handling for Transfer of Equity Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const INT_RATE_ADD = 0
Private Const INT_RATE_EDIT = 1
Private Const INT_RATE_DELETE = 2
Private Const START_DATE = 5

' Private data
Private m_clsMortgageProductTable As MortgageProductTable

Private m_colProductKeys As Collection
Private m_colTabs As Collection

Private m_bIsEdit As Boolean
Private m_nActiveTab  As Integer
Private m_ReturnCode As MSGReturnCode

Private m_sLenderCode As String
Private m_bCheckLenderChange As Boolean

Private m_clsAdditionalFeaturesTable As MortProdAddtFeaTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Birmingham Midshires Change history
' Prog      Date        Description
' MO        04/07/2002  BMIDS00088 Added new valuation fee columns to Sub SetupValuation()
' HMA       20/12/2004  BMIDS959   Remove ValidateRedemption


Private Sub Class_Initialize()
On Error GoTo Failed

    m_bIsEdit = False
    m_sLenderCode = ""
    m_bCheckLenderChange = True

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Load()
On Error GoTo Failed

    SetReturnCode MSGFailure

Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Public Sub Show()
    On Error GoTo Failed
    
    frmProductDetails.Show vbModal, frmMain
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm frmLenderDetails
End Sub

Public Sub UnloadForm()
    On Error GoTo Failed
    
    Unload frmProductDetails
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub OK()
On Error GoTo Failed
    
    SetReturnCode
    SaveChangeRequest
    frmProductDetails.Hide
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub Cancel()
    On Error GoTo Failed
    
    frmProductDetails.Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub SetEditState()
On Error GoTo Failed
    
    frmProductDetails.txtProductDetails(START_DATE).Enabled = False
    TableAccess(m_clsMortgageProductTable).SetKeyMatchValues m_colProductKeys

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CreateTabs()
On Error GoTo Failed
    Dim nTabCount As Integer
    Dim nThisTab As Integer
    Dim clsMortgageProduct As MortgageProductDetails
    Dim clsMortProdIncentive As MortProdIncentive
    Dim clsMortProdParams As AdditionalMortProdParams
    Dim clsMortProdPurpOfLoan As MortProdPurpOfLoan
    Dim clsMortProdChannelElib As MortProdChannelElig
    Dim clsMortProdEmpElig As MortProdEmpElig
    Dim clsMortProdSpecialGroup As MortProdSpecialGroup
    Dim clsMortProdAdminFee As MortProdAdminFees
    Dim clsMortProdValuationFee As MortProdValuationFee
' TW 09/10/2006 EP2_7
    Dim clsMortProdAddBorrFee As MortProdAddBorrFee
    Dim clsMortProdCredLimIncFee As MortProdCredLimIncFee
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Dim clsMortProdTranEquFee As MortProdTranEquFee
' TW 11/12/2006 EP2_20 End
    Dim clsMortProdOtherFee As MortProdOtherFee
    Dim clsMortProdIntRates As MortProdIntRates
    Dim clsMortProdBaseRates As MortProdBaseRates
    Dim clsMortProdTypicalAPR As MortProdTypicalAPR
    Dim clsMortProdPropLocation As MortProdPropLocation
    Dim clsMortProdTypeOfBuyer As MortProdTypeOfBuyer
    Dim clsTypeOfApplication As MortProdTypeofAppElig
    '*=[MC]BMIDS763 - CC075
    Dim clsMortProdTTFee As MortProdTTFees
    Dim clsMortProdIAFee As MortProdIAFees
    Dim clsMortProdSwitchFee As MortProdSwitchFees
' TW 30/11/2006 EP2_253
    Dim clsMortProdIncomeStatus As MortProdIncomeStatus
    Dim clsMortProdNatureOfLoan As MortProdNatureOfLoan
    Dim clsMortProdProdClass As MortProdProdClass
' TW 30/11/2006 EP2_253 End
    
    '*=END
    Set m_colTabs = New Collection
       
    ' Need to setup the tabs ListView fields.
    SetupListViewFields
    
    ' Add classes for all tabs
    Set clsMortgageProduct = New MortgageProductDetailsCS
    Set clsMortProdIncentive = New MortProdIncentiveCS
    Set clsMortProdParams = New MortProdParamsCS
    Set clsMortProdPurpOfLoan = New MortProdPurpOfLoanCS
    Set clsMortProdChannelElib = New MortProdChannelEligCS
    Set clsMortProdEmpElig = New MortProdEmpEligCS
    Set clsMortProdSpecialGroup = New MortProdSpecialGroupCS
    Set clsMortProdAdminFee = New MortProdAdminFeesCS
    Set clsMortProdValuationFee = New MortProdValuationFeeCS
    
' TW 09/10/2006 EP2_7
    Set clsMortProdAddBorrFee = New MortProdAddBorrFeeCS
    Set clsMortProdCredLimIncFee = New MortProdCredLimIncFeeCS
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Set clsMortProdTranEquFee = New MortProdTranEquFeeCS
' TW 11/12/2006 EP2_20 End
' TW 30/11/2006 EP2_253
    Set clsMortProdIncomeStatus = New MortProdIncomeStatus
    Set clsMortProdNatureOfLoan = New MortProdNatureOfLoan
    Set clsMortProdProdClass = New MortProdProdClass
' TW 30/11/2006 EP2_253 End
    
    '*=[MC]BMIDS763 - CC075 - FEE SETS
    Set clsMortProdTTFee = New MortProdTTFeesCS
    Set clsMortProdIAFee = New MortProdIAFeesCS
    Set clsMortProdSwitchFee = New MortProdSwitchFeesCS
    '*=[MC]SECTION END - BMIDS0763
    
    Set clsMortProdOtherFee = New MortProdOtherFeeCS
    Set clsMortProdIntRates = New MortProdIntRatesCS
    Set clsMortProdBaseRates = New MortProdBaseRatesCS
    Set clsMortProdTypicalAPR = New MortProdTypicalAprCS
    Set clsMortProdPropLocation = New MortProdPropLocationCS
    Set clsMortProdTypeOfBuyer = New MortProdTypeOfBuyerCS
    Set clsTypeOfApplication = New MortProdTypeOfAppCS
    
    m_colTabs.Add clsMortgageProduct
    m_colTabs.Add clsMortProdParams
    m_colTabs.Add clsMortProdPurpOfLoan
    m_colTabs.Add clsMortProdChannelElib
    m_colTabs.Add clsMortProdEmpElig
    m_colTabs.Add clsMortProdSpecialGroup
    m_colTabs.Add clsMortProdAdminFee
    m_colTabs.Add clsMortProdValuationFee
' TW 09/10/2006 EP2_7 End
    m_colTabs.Add clsMortProdOtherFee
    m_colTabs.Add clsMortProdIncentive
    m_colTabs.Add clsMortProdIntRates
    m_colTabs.Add clsMortProdBaseRates
    m_colTabs.Add clsMortProdTypicalAPR
    m_colTabs.Add clsMortProdPropLocation
    m_colTabs.Add clsMortProdTypeOfBuyer
    m_colTabs.Add clsTypeOfApplication
    '*=[MC]BMIDS763 - CC075 - FEE SETS
    m_colTabs.Add clsMortProdTTFee
    m_colTabs.Add clsMortProdIAFee
    m_colTabs.Add clsMortProdSwitchFee
' TW 09/10/2006 EP2_7
    m_colTabs.Add clsMortProdAddBorrFee
    m_colTabs.Add clsMortProdCredLimIncFee
' TW 11/12/2006 EP2_20
    m_colTabs.Add clsMortProdTranEquFee
' TW 11/12/2006 EP2_20 End
' TW 30/11/2006 EP2_253
    m_colTabs.Add clsMortProdIncomeStatus
    m_colTabs.Add clsMortProdNatureOfLoan
    m_colTabs.Add clsMortProdProdClass
' TW 30/11/2006 EP2_253 End
    
    
    '*=SECTION END
    
    
    If m_bIsEdit Then
        clsMortgageProduct.SetProductKeys m_colProductKeys
    End If
    Set m_clsMortgageProductTable = clsMortgageProduct.GetProductTable
    
    'BMIDS775 -ADDITIONALFEATURES
    Set m_clsAdditionalFeaturesTable = New MortProdAddtFeaTable
    
    ' Initialise the MortgageProduct table first as a pre-requisite for the rest.
    clsMortgageProduct.Initialise m_bIsEdit
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitialiseTabs()
    On Error GoTo Failed
    Dim nTabCount As Integer
    Dim nThisTab As Integer
    
    ' If we are editing, locate the keys to be used for the tabs that need to locate their own
    ' records
    If m_bIsEdit Then
        m_colTabs(mortAreaTabAdminFees).SetAdminFeeSet m_clsMortgageProductTable.GetAdminFeeSet
        m_colTabs(mortAreaTabBaseRates).SetBaseRateSet m_clsMortgageProductTable.GetBaseRateFeeSet
        m_colTabs(mortAreaTabValuationFees).SetFeeSet m_clsMortgageProductTable.GetValuationFeeSet
        m_colTabs(mortAreaOtherFee).SetLender m_clsMortgageProductTable.GetOrganisationID
        '*=[MC]BMIDS763 - CC075 - FEE SETS
        m_colTabs(mortAreaTabTTFees).SetTTFeeSet m_clsMortgageProductTable.GetTTFeeSet
        m_colTabs(mortAreaTabIAFees).SetIAFeeSet m_clsMortgageProductTable.GetIAFeeSet
        m_colTabs(mortAreaTabProdSwitchFees).SetProductSwitchFeeSet m_clsMortgageProductTable.GetProductSwitchFeeSet
' TW 09/10/2006 EP2_7
        m_colTabs(mortAreaTabAdditionalBorrowingFees).SetFeeSet m_clsMortgageProductTable.GetAdditionalBorrowingFeeSet
        m_colTabs(mortAreaTabCreditLimitIncreaseFees).SetFeeSet m_clsMortgageProductTable.GetCreditLimitIncreaseFeeSet
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        m_colTabs(mortAreaTabTransferOfEquityFees).SetFeeSet m_clsMortgageProductTable.GetTransferOfEquityFeeSet
' TW 11/12/2006 EP2_20 End

        'Dim var As Variant
        'Set var = m_colTabs(mortAreaTabProdSwitchFees)
        'Call MsgBox(TypeName(var))
        '*=SECTION END
    End If
    
    ' Loop through all tabs setting the product keys and saving the data
    nTabCount = m_colTabs.Count

    For nThisTab = mortAreaProdDetails + 1 To nTabCount
        m_colTabs(nThisTab).Initialise m_bIsEdit
    Next
    
    '*=BMIDS775
    If m_bIsEdit Then
        Call TableAccess(m_clsAdditionalFeaturesTable).SetKeyMatchValues(m_colProductKeys)
        Call TableAccess(m_clsAdditionalFeaturesTable).GetTableData
    Else
        Call TableAccess(m_clsAdditionalFeaturesTable).GetTableData(POPULATE_EMPTY)
    End If
    'END SECTION
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
    Exit Sub
    Resume
End Sub

Public Sub PopulateScreenFields()
On Error GoTo Failed
    Dim nTabCount As Integer
    Dim nThisTab As Integer
    
    If m_bIsEdit Then
        nTabCount = m_colTabs.Count
        
        For nThisTab = 1 To nTabCount
            m_colTabs(nThisTab).SetScreenFields
        Next
        
    End If
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    Exit Sub
End Sub

Public Sub SaveScreenData()
On Error GoTo Failed
    On Error GoTo Failed
    Dim nThisTab As Integer
    Dim nTabCount As Integer

    Dim clsMortgageProduct As MortgageProductDetails
    Dim clsMortgageProductTable As MortgageProductTable
    Dim colKeys As Collection
    
    Set clsMortgageProduct = m_colTabs(mortAreaProdDetails)
        
    SaveSets
    
    clsMortgageProduct.SaveScreenData
    Set clsMortgageProductTable = clsMortgageProduct.GetProductTable
    Set colKeys = TableAccess(clsMortgageProductTable).GetKeyMatchValues
    
    
    ' Loop through all tabs setting the product keys and saving the data
    nTabCount = m_colTabs.Count

    For nThisTab = mortAreaProdDetails + 1 To nTabCount
        m_colTabs(nThisTab).SetProductKeys colKeys
        m_colTabs(nThisTab).SaveScreenData
    Next
    
    '*=BMIDS775
    SaveFlexiFeaturesData colKeys
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveFlexiFeaturesData(colKeys As Collection)

On Error GoTo Failed
    
    'If frmProductDetails.optFlexibleProduct(1).Value Then
      'TableAccess(m_clsAdditionalFeaturesTable).SetKeyMatchValues colKeys
      'TableAccess(m_clsAdditionalFeaturesTable).DeleteRecords
    'Else
        m_clsAdditionalFeaturesTable.SetProductKeyValuesOnRS colKeys
        TableAccess(m_clsAdditionalFeaturesTable).Update
    'End If
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SaveSets()
    On Error GoTo Failed
    Dim sSetNo As String
    Dim clsMortgageProduct As MortgageProductDetails
    Dim clsMortProdAdminFees As MortProdAdminFees
    Dim clsMortProdValuationFee As MortProdValuationFee
    Dim clsMortProdBaseRates As MortProdBaseRates
    
' TW 09/10/2006 EP2_7
    Dim clsMortProdAddBorrFee As MortProdAddBorrFee
    Dim clsMortCredLimIncFee As MortProdCredLimIncFee
' TW 09/10/2006 EP2_7 End
' TW 30/11/2006 EP2_253
    Dim clsMortIncomeStatus As MortProdIncomeStatus
    Dim clsMortNatureOfLoan As MortProdNatureOfLoan
    Dim clsMortProdProdClass As MortProdProdClass
' TW 30/11/2006 EP2_253 End
' TW 11/12/2006 EP2_20
    Dim clsMortProdTranEquFee As MortProdTranEquFee
' TW 11/12/2006 EP2_20 End
    
    '*=[MC]BMIDS763
    Dim clsMortProdIAFees As MortProdIAFees
    Dim clsMortProdTTFees As MortProdTTFees
    Dim clsMortProdSwitchFees As MortProdSwitchFees
    '*=SECTION END

' TW 09/10/2006 EP2_7
    Set clsMortProdAddBorrFee = m_colTabs(mortAreaTabAdditionalBorrowingFees)
    Set clsMortCredLimIncFee = m_colTabs(mortAreaTabCreditLimitIncreaseFees)
' TW 09/10/2006 EP2_7 End
' TW 30/11/2006 EP2_253
    Set clsMortIncomeStatus = m_colTabs(mortAreaTabMortgageProductIncomeStatus)
    Set clsMortNatureOfLoan = m_colTabs(mortAreaTabMortgageProductNatureOfLoan)
    Set clsMortProdProdClass = m_colTabs(mortAreaTabMortgageProductProductClass)
' TW 30/11/2006 EP2_253 End
' TW 11/12/2006 EP2_20
    Set clsMortProdTranEquFee = m_colTabs(mortAreaTabTransferOfEquityFees)
' TW 11/12/2006 EP2_20 End

    Set clsMortProdAdminFees = m_colTabs(mortAreaTabAdminFees)
    Set clsMortProdValuationFee = m_colTabs(mortAreaTabValuationFees)
    Set clsMortProdBaseRates = m_colTabs(mortAreaTabBaseRates)
    Set clsMortgageProduct = m_colTabs(mortAreaProdDetails)
    
    '*=[MC]BMIDS763
    
    Set clsMortProdIAFees = m_colTabs(mortAreaTabIAFees)
    Set clsMortProdTTFees = m_colTabs(mortAreaTabTTFees)
    Set clsMortProdSwitchFees = m_colTabs(mortAreaTabProdSwitchFees)
    
    
    '*=SECTION END
    
    sSetNo = clsMortProdAdminFees.GetAdminFeeSet
    clsMortgageProduct.SetAdminFeeSet sSetNo
    
    sSetNo = clsMortProdValuationFee.GetValuationFeeSet
    clsMortgageProduct.SetValuationFeeSet sSetNo
    
    '*=[MC]BMIDS763 FEE SETS
    sSetNo = clsMortProdIAFees.GetIAFeeSet
    Call clsMortgageProduct.SetInsuranceAdminFeeSet(sSetNo)
    
    sSetNo = clsMortProdTTFees.GetTTFeeSet
    Call clsMortgageProduct.SetTTFeeSet(sSetNo)
    
    sSetNo = clsMortProdSwitchFees.GetProductSwitchFeeSet
    Call clsMortgageProduct.SetProductSwitchFeeSet(sSetNo)
    '*=SECTION END
    sSetNo = clsMortProdBaseRates.GetBaseRateSet
    clsMortgageProduct.SetBaseRateSet sSetNo

' TW 09/10/2006 EP2_7
    sSetNo = clsMortProdAddBorrFee.GetAdditionalBorrowingFeeSet
    clsMortgageProduct.SetAdditionalBorrowingFeeSet sSetNo
    
    sSetNo = clsMortCredLimIncFee.GetFeeSet
    clsMortgageProduct.SetCreditLimitIncreaseFeeSet sSetNo
    
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    sSetNo = clsMortProdTranEquFee.GetTransferOfEquityFeeSet
    clsMortgageProduct.SetTransferOfEquityFeeSet sSetNo
' TW 11/12/2006 EP2_20 End


Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetActiveTab(Optional nTab As ProductArea)
On Error GoTo Failed
    
    Dim bRet As Boolean
    bRet = False
    
    If Not IsMissing(nTab) Then
        If nTab > 0 Then
            m_nActiveTab = nTab
            bRet = True
        End If
    ElseIf m_nActiveTab > 0 Then
        bRet = True
    End If
    
    If bRet = True Then
       frmProductDetails.SSTab1.Tab = m_nActiveTab - 1
    Else
        MsgBox "SetActiveTab: Unable to set tab, invalid value of " & m_nActiveTab
    End If
    
    ' Make sure we set the tab stops for this tab too otherwise the tabbing will disappear onto the
    ' next tab (a problem with the tabbed dialog control)
    g_clsFormProcessing.SetTabstops frmProductDetails
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Function ValidateScreenData() As Boolean
On Error GoTo Failed
    
    Dim bSuccess As Boolean
    Dim nCount As Integer
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    
    bSuccess = g_clsFormProcessing.DoMandatoryProcessing(frmProductDetails)
    
    If bSuccess Then
        nTabCount = m_colTabs.Count
        
        nThisTab = 1
        Do While bSuccess = True And nThisTab <= nTabCount
            bSuccess = m_colTabs(nThisTab).ValidateScreenData()
            nThisTab = nThisTab + 1
        Loop
    End If
    
    ValidateScreenData = bSuccess
    
Exit Function
Failed:
    'Bring the offending tab to the forground.
    ShowInvalidTab nThisTab
        
    'Raise the error.
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetTabsCollection() As Collection
    On Error GoTo Failed
    
    If m_colTabs Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Mortgage Product tab list is empty"
    End If
    
    Set GetTabsCollection = m_colTabs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetupListViewFields()
On Error GoTo Failed
    
    SetupPurposeOfLoan
    SetupTypeOfApp
    SetupChannelElig
    SetupEmpElig
    SetupAdminFees
    SetupValuationFees
    SetupOtherFees
    SetupPropLocation
    
    '*=[MC]BMIDS763 - CC075 - FEE SETS
    Call SetupIAFees
    Call SetupTTFees
    Call SetupProductSwitchFees
    
' TW 09/10/2006 EP2_7
    SetupAdditionalBorrowingFees
    SetupCreditLimitIncreaseFees
' TW 09/10/2006 EP2_7 End    '*=SECTION END
' TW 30/11/2006 EP2_253
    SetupIncomeStatus
    SetupProductClass
    SetupNatureOfLoan
' TW 30/11/2006 EP2_253 End
' TW 11/12/2006 EP2_20
    SetupTransferOfEquityFees
' TW 11/12/2006 EP2_20 End



Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupOtherFees()
On Error GoTo Failed
    
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 50
    lvHeaders.sName = "Fee Name"
    headers.Add lvHeaders

    lvHeaders.nWidth = 50
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders


    frmProductDetails.MSGSwapOtherFees.SetFirstColumnHeaders headers
    frmProductDetails.MSGSwapOtherFees.SetSecondColumnHeaders headers

    frmProductDetails.MSGSwapOtherFees.SetFirstTitle "Available Other Fees"
    frmProductDetails.MSGSwapOtherFees.SetSecondTitle "Selected Other Fees"
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupValuationFees()
On Error GoTo Failed
    
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Set No"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Type of Valuation"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Location"
    headers.Add lvHeaders

    lvHeaders.nWidth = 10
    lvHeaders.sName = "Max Value"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"
    headers.Add lvHeaders

    ' MO - 04/07/2002 - BMIDS00088 - Start
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Fee Percentage"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Minimum Fee Value"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Maximum Fee Value"
    headers.Add lvHeaders
    ' MO - 04/07/2002 - BMIDS00088 - End

' TW 09/10/2006 EP2_7

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Type of Application"
    headers.Add lvHeaders
' TW 09/10/2006 EP2_7 End

    frmProductDetails.lvValuationFees.AddHeadings headers
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupAdminFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Type of App"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Location"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    frmProductDetails.lvAdminFees.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupTTFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Type of App"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Location"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    frmProductDetails.lvTTFeeSet.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupIncomeStatus()
On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Income Status"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapIncomeStatus.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Income Status"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapIncomeStatus.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetupProductClass()
On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Product Class"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapProductClass.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Product Class"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapProductClass.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetupNatureOfLoan()
On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Nature of Loan"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapNatureOfLoan.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Nature of Loan"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapNatureOfLoan.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetupTransferOfEquityFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Nature Of Loan"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    frmProductDetails.lvTransferOfEquityFees.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetupAdditionalBorrowingFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Nature Of Loan"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Max LTV"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Percentage"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Min Fee Value"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Max Fee Value"

    headers.Add lvHeaders

    frmProductDetails.lvAdditionalBorrowingFees.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub








Private Sub SetupCreditLimitIncreaseFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Nature Of Loan"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Max LTV"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Percentage"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Min Fee Value"

    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Max Fee Value"

    headers.Add lvHeaders

    frmProductDetails.lvCreditLimitIncreaseFees.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub SetupIAFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Type of App"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Location"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    frmProductDetails.lvIAFeeSet.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupProductSwitchFees()
On Error GoTo Failed
      
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Type of App"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Location"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Fee Amount"

    headers.Add lvHeaders

    frmProductDetails.lvProdSwitchFeeSet.AddHeadings headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub SetupPropLocation()
On Error GoTo Failed
        
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Property Location"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapPropLocation.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.sName = "Selected Property Location"

    headers.Add lvHeaders
    frmProductDetails.MSGSwapPropLocation.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

 Private Sub SetupEmpElig()
On Error GoTo Failed
        
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Employment Eligibility"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapEmpElig.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.sName = "Eligible Product Status"

    headers.Add lvHeaders
    frmProductDetails.MSGSwapEmpElig.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupChannelElig()
On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Sales Channel List"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapChannel.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.sName = "Sales Channel Eligibility"

    headers.Add lvHeaders
    frmProductDetails.MSGSwapChannel.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupTypeOfApp()
On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Application Eligibility"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapTypeOfAppEligibility.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Eligible Type of Application"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapTypeOfAppEligibility.SetSecondColumnHeaders headers
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupPurposeOfLoan()
On Error GoTo Failed

    Dim headers As New Collection
    Dim lvHeaders As listViewAccess

    ' Purpose of Loan
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Purpose of Loan"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapPurpOfLoan.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Purpose of Loan Eligibility"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapPurpOfLoan.SetSecondColumnHeaders headers

    ' Type of Buyer
    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Available Type of Buyer"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapTypeOfBuyer.SetFirstColumnHeaders headers

    Set headers = New Collection
    lvHeaders.nWidth = 100
    lvHeaders.sName = "Selected Type of Buyer"
    headers.Add lvHeaders
    frmProductDetails.MSGSwapTypeOfBuyer.SetSecondColumnHeaders headers

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Initialise()
On Error GoTo Failed
    Dim nTabCount As Long
    Dim nThisTab As Long
    
    ' Loop through all tabs setting the product keys and saving the data
    nTabCount = m_colTabs.Count

    For nThisTab = mortAreaProdDetails + 1 To nTabCount
        m_colTabs(nThisTab).SetProductKeys m_colProductKeys
    Next

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub HandleNewLender()
    On Error GoTo Failed
    Dim vOrgID As Variant
    Dim bSuccess As Boolean
    Dim clsMortProdOtherFee As MortProdOtherFee
    Dim clsMortgageProduct As MortgageProductDetails
    
    Set clsMortgageProduct = m_colTabs(mortAreaProdDetails)
    Set clsMortProdOtherFee = m_colTabs(mortAreaOtherFee)
    
    vOrgID = clsMortgageProduct.SetOrgIDFromCode
    
    bSuccess = clsMortProdOtherFee.SetLender(vOrgID)
    
    If bSuccess Then
        m_sLenderCode = frmProductDetails.cboLenderCode.SelText
    Else
        m_bCheckLenderChange = False
        frmProductDetails.cboLenderCode.SelText = m_sLenderCode
        m_bCheckLenderChange = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetKeyMatchValues() As Collection
    On Error GoTo Failed

    If Not m_colProductKeys Is Nothing Then
        Set GetKeyMatchValues = m_colProductKeys
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty
    End If

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortgageProduct.GetKeyMatchValues " & Err.DESCRIPTION
End Function

Public Sub SetKeys(colKeys As Collection)
On Error GoTo Failed

    Set m_colProductKeys = colKeys

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
On Error GoTo Failed
    
    m_bIsEdit = bIsEdit
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Function IsEdit() As Boolean
On Error GoTo Failed

    IsEdit = m_bIsEdit
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
End Function

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    On Error GoTo Failed
    
    m_ReturnCode = enumReturn
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetReturnCode() As MSGReturnCode
    On Error GoTo Failed
    
    GetReturnCode = m_ReturnCode
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SaveChangeRequest()
On Error GoTo Failed
    
    g_clsHandleUpdates.SaveChangeRequest m_clsMortgageProductTable
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitialiseTab()
    On Error GoTo Failed
    Dim nTab As Integer
    
    If Not m_colTabs Is Nothing Then
        nTab = frmProductDetails.SSTab1.Tab + 1
        
        Select Case nTab
            Case mortProdDetails
                m_colTabs(mortAreaProdDetails).InitTab
                
            Case mortProdParametersAndBands
                m_colTabs(mortAreaParameters).InitTab
            
            Case mortProdAppEligibility
                m_colTabs(mortAreaTypeOfApp).InitTab
                m_colTabs(mortAreaPurpOfLoan).InitTab
                m_colTabs(mortAreaTypeOfBuyer).InitTab
            
            Case mortProdOtherEligibility
                m_colTabs(mortAreaChannelElig).InitTab
                m_colTabs(mortAreaEmpElig).InitTab
                m_colTabs(mortAreaPropLocation).InitTab
            
            Case mortProdFees
                m_colTabs(mortAreaTabAdminFees).InitTab
                m_colTabs(mortAreaTabValuationFees).InitTab
' TW 09/10/2006 EP2_7
                m_colTabs(mortAreaTabAdditionalBorrowingFees).InitTab
' TW 09/10/2006 EP2_7 End
            Case mortProdFeeSet3
' TW 09/10/2006 EP2_7
                m_colTabs(mortAreaTabCreditLimitIncreaseFees).InitTab
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
                m_colTabs(mortAreaTabTransferOfEquityFees).InitTab
' TW 11/12/2006 EP2_20 End
            Case mortProdOtherFeesAndIncentives
                m_colTabs(mortAreaOtherFee).InitTab
                m_colTabs(mortAreaTabIncentives).InitTab

            Case mortProdInteresRates
                m_colTabs(mortAreaTabInterestRates).InitTab
                m_colTabs(mortAreaTabBaseRates).InitTab

            Case mortProdMisc
                m_colTabs(mortAreaTypicalAPR).InitTab
                m_colTabs(mortAreaSpecialGroup).InitTab
                
            Case mortProdAppEligibility3
' TW 30/11/2006 EP2_253
                m_colTabs(mortAreaTabMortgageProductIncomeStatus).InitTab
                m_colTabs(mortAreaTabMortgageProductNatureOfLoan).InitTab
                m_colTabs(mortAreaTabMortgageProductProductClass).InitTab
' TW 30/11/2006 EP2_253 End
        End Select
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveAdditionalFeaturesData(colKeys As Collection)

On Error GoTo Failed
    
    If frmProductDetails.optFlexibleProduct(1).Value Then
        TableAccess(m_clsAdditionalFeaturesTable).SetKeyMatchValues colKeys
        TableAccess(m_clsAdditionalFeaturesTable).DeleteRecords
    Else
        m_clsAdditionalFeaturesTable.SetProductKeyValuesOnRS colKeys
        TableAccess(m_clsAdditionalFeaturesTable).Update
    End If
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub LenderChange()
    On Error GoTo Failed
    
    If m_bCheckLenderChange And Len(frmProductDetails.cboLenderCode.SelText) > 0 Then
        HandleNewLender
    End If
   
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub ProductGroupingPressed()
    On Error GoTo Failed
    
    frmProductGrouping.SetTableClass m_clsMortgageProductTable
    frmProductGrouping.SetKeys m_colProductKeys
    frmProductGrouping.SetIsEdit m_bIsEdit
    frmProductGrouping.Show vbModal, frmProductDetails
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'*=[MC]BMIDS775 - CC071- Reg018 Additional Features dialog

Public Sub AdditionalFeaturesPressed(Optional ByVal vstrMortProdCode As String, Optional ByVal vstrStartDt As String)
    
    On Error GoTo ErrorHandler
    '*=For New Product
'    If Not m_bIsEdit Then
'        If m_colProductKeys Is Nothing Then
'            Set m_colProductKeys = New Collection
'            m_colProductKeys.Add vstrMortProdCode
'            m_colProductKeys.Add vstrStartDt
'        End If
'    End If
    
    '*=[MC]Set Table Class
    frmAdditionalFeatures.SetTableClass m_clsAdditionalFeaturesTable
    '*=Set Columns Collection
    'frmAdditionalFeatures.SetKeys m_colProductKeys
    'frmAdditionalFeatures.SetIsEdit m_bIsEdit
    '*=Show object as model dialog
    frmAdditionalFeatures.Show vbModal, frmProductDetails
    
    

ExitHandler:
    '*=Cleanup here
    Exit Sub
    
ErrorHandler:
    '*=Log and Throw error
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SelectItem(nItem As ProductArea)
    On Error GoTo Failed
    
    m_colTabs(nItem).Selected
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeselectItem(nItem As ProductArea)
    On Error GoTo Failed
    
    m_colTabs(nItem).DeSelect
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Delete(nItem As ProductArea)
    On Error GoTo Failed
    
    m_colTabs(nItem).Delete
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Edit(nItem As ProductArea)
    On Error GoTo Failed
    
    m_colTabs(nItem).Edit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Add(nItem As ProductArea)
    On Error GoTo Failed
    
    m_colTabs(nItem).Add
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'BMIDS959 Remove Sub ValidateRedemption

Private Sub ShowInvalidTab(ByVal uProductArea As ProductArea)

    Dim uTabToShow As ProductTabs
        
    'Store the original offending error message.
    g_clsErrorHandling.SaveError Err.Number, Err.DESCRIPTION
    
    On Error Resume Next

    Select Case uProductArea
        Case mortAreaProdDetails
            uTabToShow = mortProdDetails
            
        Case mortAreaParameters
            uTabToShow = mortProdParametersAndBands
        
        Case mortAreaTypeOfApp, mortAreaPurpOfLoan, mortAreaTypeOfBuyer
            uTabToShow = mortProdAppEligibility
        
        Case mortAreaChannelElig, mortAreaEmpElig, mortAreaPropLocation
            uTabToShow = mortProdOtherEligibility
        
        Case mortAreaTabAdminFees, mortAreaTabValuationFees
            uTabToShow = mortProdFees
        
        Case mortAreaOtherFee, mortAreaTabIncentives
            uTabToShow = mortProdOtherFeesAndIncentives
        
        Case mortAreaTabInterestRates, mortAreaTabBaseRates
            uTabToShow = mortProdInteresRates
        
        Case mortAreaTypicalAPR, mortAreaSpecialGroup
            uTabToShow = mortProdMisc
    End Select

    'Display the relevant tab.
    frmProductDetails.SSTab1.Tab = (uTabToShow - 1)

    'Bug fix. The tab control (although not raising an error) seems to set the
    'error description to 'object or property not supported' if .Tab is altered
    'to a new value. We've already stored the original error message to be
    'raised.
    Err.Clear

End Sub

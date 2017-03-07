VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerFinancialDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006 Added CreateAccountRelationships
'                                    DeleteAccountRelationships
'                                    UpdateAccountRelationships
'                                    CreateSpecialFeature
'                                    UpdateSpecialFeature
'                                    DeleteSpecialFeature
'                                    FindSpecialFeatureList
'                                    GetMortgageAccountDetails
'GHun   27/08/2002  BMIDS00190 DCWP3 BM076 Added FindLoansListForArrears
'MV     02/09/2002  BMIDS00353  CMWP4 - BM087 ; Added GetMtgAcctDataForFurtherAdv
'MDC    10/09/2002  BMIDS00336 CCWP1 BM062 - Add GetMortgageAccountLoanTotals
'GHun   14/11/2002  BMIDS00444 Added FindRemortgageAccountAddress
'SR     25/05/2004  BMIDS772 - new methods - GetFinancialSummaryView, GetCreditHistorySummaryView
'-------------------------------------------------------------------------
'Epsom History:
'
'Prog   Date        Description
'GHun   30/01/2007  EP2_954 Added new function GetAdditionalBorrowingDataForProjectionCalc
'-------------------------------------------------------------------------

Public Function GetMtgAcctDataForFurtherAdv(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

'SR 25/05/2004 : BMIDS772
Public Function GetFinancialSummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'SR 25/05/2004 : BMIDS772
Public Function GetCreditHistorySummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub CreateArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateBankCard(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateLiability(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                             Optional ByVal vblnSimpleCreate As Boolean = False)
End Sub

Public Sub CreateMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement, _
                                 ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub CreateRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
Attribute CreateRegularOutgoings.VB_Description = "Data  Process for creating & populating a new record in the ""RegularOutgoings"" table."
End Sub

Public Sub CreateMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement)
End Sub

Public Sub DeleteAllMortgageLoans(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteBankCard(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteLiability(ByVal vxmlTableElement As IXMLDOMElement, _
                           Optional ByVal vblnSimpleDelete As Boolean = False)
End Sub

Public Sub DeleteLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                             Optional ByVal vblnSimpleDelete As Boolean = False)
End Sub

Public Sub DeleteMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
Attribute DeleteRegularOutgoings.VB_Description = "Data Process for deleting Regular Outgoings"
End Sub

Public Function FindArrearsHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindBankCardList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindBankruptcyHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindCCJHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDeclinedMortgageList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindGroupConnectionList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindLiabilityList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindLifeProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageAccountList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindRegularOutgoingsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
Attribute FindRegularOutgoingsList.VB_Description = "Data Process to return relevant data from the ""RegularOutgoings"" table"
End Function

Public Function FindMortgageListForArrears(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageLoanList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindOutgoingsTotalList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAccountDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageLenderDirectoryData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub UpdateArrearsHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateBankCard(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateBankruptcyHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateCCJHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateDeclinedMortgage(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateGroupConnection(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateLiability(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateLifeProduct(ByVal vxmlTableElement As IXMLDOMElement, _
                             Optional ByVal vblnUpdatePolicyRelationship As Boolean = True)
End Sub

Public Sub UpdateMortgageAccount(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateRegularOutgoings(ByVal vxmlTableElement As IXMLDOMElement)
Attribute UpdateRegularOutgoings.VB_Description = "Data process to update Regular Outgoings"
End Sub

Public Sub UpdateMortgageLoan(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindLoansToBePorted(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindLoansAndLiabilitiesWithNoLender(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageAccountAndLoanList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetLoansNotRedeemedTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub CreateAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateAccountRelationships(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindAccountRelationshipList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub CreateSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement, ByVal vxmlKeys As IXMLDOMElement)
End Sub

Public Sub UpdateSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteSpecialFeature(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindSpecialFeatureList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageAccountDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS00190
Public Function FindLoansListForArrears(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00190 End

'BMIDS00336 MDC 10/09/2002
Public Function GetMortgageAccountLoanTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00336 MDC 10/09/2002 - End

'BMIDS00444
Public Function FindRemortgageAccountAddress(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00444 End

'BMIDS815
Public Function GetPortedStepAndPeriodFromMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS815 End

'EP2_954 GHun
Public Function GetAdditionalBorrowingDataForProjectionCalc(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'EP2_954 End

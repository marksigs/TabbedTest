VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerFinancialBO"
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
'GHun   22/08/2002  BMIDS00190 DCWP3 BM076 Added CalculateCreditCardRepayment,
'                                                FindLoansListForArrears
'MV     02/09/2002  BMIDS00353  CMWP4 - BM087 ; Added GetMtgAcctDataForFurtherAdv
'MDC    10/09/2002  BMIDS00336 CCWP1 BM062 - Add GetMortgageAccountLoanTotals
'GHun   14/11/2002  BMIDS00444 Added FindRemortgagAccountAddress
'SR     25/05/2004  BMIDS772 - new methods - GetFinancialSummaryView, GetCreditHistorySummaryView
'-------------------------------------------------------------------------
'Epsom History:
'
'Prog   Date        Description
'GHun   05/02/2007  EP2_954 Added new function GetAdditionalBorrowingDataForProjectionCalc
'-------------------------------------------------------------------------

Public Function GetMtgAcctDataForFurtherAdv(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' SR 25/05/2004 : BMIDS772
Public Function GetFinancialSummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' SR 25/05/2004 : BMIDS772
Public Function GetCreditHistorySummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute CreateRegularOutgoings.VB_Description = "Business Process for Creating & populating a new record in the ""RegularOutgoings"" table."
End Function

Public Function CreateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute DeleteRegularOutgoings.VB_Description = "Business Process for Deleting Regular Outgoings"
End Function

Public Function FindArrearsHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindBankCardList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindBankruptcyHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindCCJHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDeclinedMortgageList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindGroupConnectionList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindLiabilityList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindLifeProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindLoansToBePorted(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindOutgoingsTotalList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindMortgageAccountList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindMortgageAccountAndLoanList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindRegularOutgoingsList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute FindRegularOutgoingsList.VB_Description = "Business Process for finding  Regular Outgoings.  (Note that there is a similar function in AIP which also defaults from QQOuytgoings)"

End Function

Public Function FindMortgageListForArrears(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindMortgageLoanList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetAccountDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute UpdateRegularOutgoings.VB_Description = "Updates a record within the ""RegularOutgoings"" table."
End Function

Public Function UpdateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute ValidateRegularOutgoings.VB_Description = "Middle Tier base Validation for the RegularOutgoings business process."

End Function

Public Function ValidatePortedLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveMortgageAccountAndLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLoansNotRedeemedTotals(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindAccountRelationshipList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindSpecialFeatureList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageAccountDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS00190
Public Function CalculateCreditCardRepayment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindLoansListForArrears(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00190 End

'BMIDS00336 MDC 10/09/2002
Public Function GetMortgageAccountLoanTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00336 MDC 10/09/2002 - End

'BMIDS00444
Public Function FindRemortgageAccountAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
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

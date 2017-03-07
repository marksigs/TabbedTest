VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICFClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006 Added LoadAccountRelationshipData
'                                    LoadMortgageAccountSpecialFeatureData
'GHun   15/08/2002  BMIDS00190 DCWP3 BM076 Added LoadCustVerRegularOutgoingsData,
'                                                LoadCustVerBankruptcyHistoryData,
'                                                LoadCustVerCCJHistoryData,
'                                                LoadOtherArrearsAccountData,
'                                                LoadCFArrearsHistoryData,
'                                                LoadCFLoansLiabilitiesData,
'                                                LoadLoansLiabilitiesArrearsData
'MDC    10/09/2002  BMIDS00336 CCWP1 BM062 - Add GetMortgageAccountLoanTotals
'GHun   14/11/2002  BMIDS00444 Added FindRemortgageAccountAddessData
'GHun   19/11/2002  BMIDS00973 Added ArrearsHistoryDeleteData
'-------------------------------------------------------------------------


Public Function LoadMortgageAccDetsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAccountData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageLenderDirectoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadArrearsHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBankCreditCardData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBankruptcyHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCCJHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCurrentPropertyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerAddressData() As FreeThreadedDOMDocument40
End Function

Public Function LoadDeclinedMortgageData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFinancialOrganisationData() As FreeThreadedDOMDocument40
End Function

Public Function LoadGroupConnectionData() As FreeThreadedDOMDocument40
End Function

Public Function LoadHomeInsuranceData() As FreeThreadedDOMDocument40
End Function

Public Function LoadIndemnityInsuranceData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLifePolicyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLifeProductData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoansLiabilitiesData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoansLiabilitiesDirectoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoansLiabilitiesThirdPartyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadRegularOutgoingsData() As FreeThreadedDOMDocument40
Attribute LoadRegularOutgoingsData.VB_Description = "Interface definition for the  ""RegularOutgoings"" table."
End Function

Public Function LoadMortgageAccountData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountDetailsDataEx() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountDirectoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountThirdPartyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageLoanData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageRelatedContractsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadApplicationContractData() As FreeThreadedDOMDocument40
End Function

Public Function LoadPolicyRelationshipData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoansNotRedeemedTotalsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountAndLoanData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountArrearsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAccountRelationshipData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountSpecialFeatureData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageAccountListData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAccountRelationshipDetailsData() As FreeThreadedDOMDocument40
End Function

' BMIDS00190 DCWP3 BM076 GHun 24/07/2002
Public Function LoadCustVerRegularOutgoingsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustVerBankruptcyHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustVerCCJHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadOtherArrearsAccountData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCFArrearsHistoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCFLoansLiabilitiesData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoansLiabilitiesArrearsData() As FreeThreadedDOMDocument40
End Function
' BMIDS00190 End

'BMIDS00336 MDC 10/09/2002
Public Function LoadMortgageAccountLoanTotalsData() As FreeThreadedDOMDocument40
End Function
'BMIDS00336 MDC 10/09/2002 - End

'BMIDS00444
Public Function LoadRemortgageAccountAddressData() As FreeThreadedDOMDocument40
End Function
'BMIDS00444 End

'BMIDS00973
Public Function LoadArrearsHistoryDeleteData() As FreeThreadedDOMDocument40
End Function
'BMIDS00973 End


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Interface declarations for ApplicationBO.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     10/01/99 Added methods for Declaration table.
'SR     14/03/00 New method 'CreateApplicationAndCustomer'
'IW    16/03/00    Added LoadApplicationVerificationData
'IW    19/03/00    Added LoadApplicationCreditCardData
'SR    23/03/00  New method 'ValidateFinancialSummary'
'                Added MemoPad methods
'MC     26/04/00    Added OtherInsuranceCompany methods
'SR     02/06/00 Addes methods SaveNewLoan and ValidateNewLoan
'PSC    03/07/00 Remove GetLoansNotRedeemed
'MC     11/07/00 Added GetCurrentApplicationStage method
'APS    28/02/01 SYS1993 Added CreateApplicationPriority method
'PSC    07/03/01 SYS1879 Added GetMortgageAccountData and GetMortgagePropertyAddress
'PSC    09/03/01 SYS1879 Added ApproveRecommendApplication
'JR     22/03/01 SYS2048 Added GetApplicationOfferData
'SR     20/07/01  SYS2412 New method 'FindApplicationFactFindList'
'BG     21/12/01    SYS3268 Added UpdateApplicationOffer and CreateApplicationOffer methods
'------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
' MARS Specific History
'
'MF     19/08/2005  MAR20       Added CretePropertyDetails
'PSC    22/09/2005  MAR32       Added DeriveAndSetSLADate
'SR     20/10/2005  MAR24       Added SaveApplicationUnderwriting and GetApplicationUnderwriting
'JD     05/07/2006  MAR1890     Added GetDecisionDetails
'-------------------------------------------------------------------------------------------------

Option Explicit
Public Function ValidateDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDeclarationList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindOtherResidentDetailsList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteAllDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function Create(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteDependantForCustomer(ByVal vxmlRequest As IXMLDOMElement, Optional ByVal bRaiseErr As Boolean = True) As IXMLDOMNode

End Function

Public Function DeleteOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindMainApplicants(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindCustomerAddressListAndSalutation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDependantsForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindVerificationList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindUserHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindApplicationStageList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDependantForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetHighestEarners(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationPriority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function



Public Function IsAppUnderReview(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function


Public Function GetApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLoanPropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement, _
                                         Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveDependantForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreatePropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveOtherResident(ByVal vxmlRequest As IXMLDOMElement, _
                                  Optional ByRef rbotOperationType As Integer = 0, _
                                  Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveFinancialSummaryAndLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Update(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplicantCounts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCorrespondenceSalutation( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateLoanPropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateApplicationFactFind( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateApplicationFactFind( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationAndCustomer( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationCustomerAndVerification( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateCreateApplicationAndCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMemoPadData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMemoPadList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateNewLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationFactFindData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetCurrentApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMenuItemStatusList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveMenuItemStatus(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetUnitFromApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationPriority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
End Function

Public Function ValidateApplicationPriority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageAccountData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgagePropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ApproveRecommendApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationOfferData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindApplicationFactFindList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' PSC 22/09/2005 MAR32 - Start
Public Function DeriveAndSetSLADate(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 22/09/2005 MAR32 - End
'SR 20/10/2005 : MAR24
Public Function SaveApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetDecisionDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'SR 20/10/2005 : MAR24 - End

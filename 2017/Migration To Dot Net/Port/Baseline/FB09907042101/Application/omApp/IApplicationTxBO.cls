VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Interface declarations for ApplicationTxBO.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     10/01/99 Added methods for Declaration table.
'SR     14/03/00 New method 'CreateApplicationAndCustomer'
'IW     17/03/00 Added ApplicationVerification Code
'IW     20/03/00 Added ApplicationCreditCard Code
'SR     24/03/00 Added methods for MemoPad
'SR     02/06/00 Added method 'SaveNewLoan'
'APS    28/02/01 SYS1993 Added CreateApplicationPriority method
'PSC    09/03/01 SYS1879 Added ApproveRecommendApplication
'BG     21/12/01    SYS3268 Added UpdateApplicationOffer and CreateApplicationOffer methods
'------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' MARS Specific History
'
'MF     19/08/2005  MAR20       Added CretePropertyDetails
'PSC    22/09/2005  MAR32       Added DeriveAndSetSLADate
'SR     20/10/2005  MAR24       Added CreateApplicationUnderwriting UpdateApplicationUnderwriting
'-------------------------------------------------------------------------------------------------

Option Explicit
Public Function SaveDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreatePropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAllDeclaration(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function Create(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteDependantForCustomer(ByVal vxmlRequest As IXMLDOMElement, Optional ByVal bRaiseErr As Boolean = True) As IXMLDOMNode

End Function

Public Function DeleteOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveFinancialSummaryAndLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveNewLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

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

Public Function UpdateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicantCounts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCorrespondenceSalutation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicationFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationAndCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationCustomerAndVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveMenuItemStatus(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationPriority(vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ApproveRecommendApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

' PSC 22/09/2005 MAR32 - Start
Public Function DeriveAndSetSLADate(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 22/09/2005 MAR32 - End
' SR 20/10/2005 MAR24 - Start
Public Function CreateApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
Public Function UpdateApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
' SR 20/10/2005 MAR24 - End

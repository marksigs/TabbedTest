VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Interface declarations for Application Data Object
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     10/01/99    Added methods for Declaration table.
'IW     16/03/00    Added LoadApplicationVerificationData
'IW     9/03/00     Added LoadApplicationCreditCardData
'SR     24/03/00    Added MemoPad methods
'MC     26/04/00    Added OtherInsuranceCompany methods
'SR     02/06/00    Added methods 'CreateNewLoan' and 'UpdateNewLoan'
'MC     11/07/00    Added GetCurrentApplicationStage method
'APS    28/02/01    SYS1993 Added CreateApplicationPriority method
'PSC    07/03/01    SYS1879 Added GetMortgageAccountData
'JR     22/03/01    SYS2048 Added GetApplicationOfferData
'SR     20/07/01    SYS2412 New method 'FindApplicationFactFindList'
'BG     21/12/01    SYS3268 Added UpdateApplicationOffer and CreateApplicationOffer methods
'------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
' MARS Specific History
'
'MF     19/08/2005  MAR20       Added CretePropertyDetails
'SR     20/10/2005  MAR24       Added CreateApplicationUnderwriting, UpdateApplicationUnderwriting and
'                               GetApplicationUnderwriting
'JD     05/07/2006  MAR1890     Added GetDecisionDetails
'-------------------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     13/08/2002  BMIDS0331   Core Ref AQR : SYS 2372 Amended GetNextApplicationNumber
'------------------------------------------------------------------------------------------

Public Function GetNextApplicationNumber(ByVal vstrChannelID As String) As String
End Function


Public Function GetUnitFromApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Sub CreateCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateDeclaration(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateOtherResident(ByVal vxmlTableElement As IXMLDOMElement, _
                                               ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub Create(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateLoanProperty(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreatePropertyDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateSharedOwnershipDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateVerification(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateUserHistory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateNewLoan(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateApplicationFactFind(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteDependantForCustomer(ByVal vxmlTableElement As IXMLDOMElement, Optional ByVal bRaiseErr As Boolean = True)

End Sub

Public Sub DeleteOtherResident(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteVerification(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteUserHistory(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindCustomerAddressListAndSalutation(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDependantsForCustomer(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindMainApplicants(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindOtherResidentList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindOtherResidentDetailsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindUserHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindVerificationList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindApplicationStageList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDependantForCustomer(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLoanPropertyDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetOtherResident(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetUserHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationStage(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationPriority(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function IsAppUnderReview(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationFactFindData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub UpdateApplication(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateOtherResident(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateVerification(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateLoanProperty(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateSharedOwnershipDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateDeclaration(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteAllDeclaration(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindDeclarationList( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub Update(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateUserHistory(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateNewLoan(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateApplicationFactFind(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateMemoPad(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteMemoPad(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetMemoPadData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindMemoPadList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Sub UpdateMemoPad(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetOtherInsuranceCompany(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub UpdateOtherInsuranceCompany(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateOtherInsuranceCompany(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function GetNewLoan(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function GetCurrentApplicationStage(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMenuItemStatusList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Sub CreateMenuItemStatus(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateMenuItemStatus(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateApplicationPriority(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetMortgageAccountData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function GetApplicationOfferData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindApplicationFactFindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Sub UpdateApplicationOffer(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateApplicationOffer(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
'SR 20/10/2005 : MAR24
Public Sub CreateApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Function GetApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function
Public Function GetDecisionDetails(ByVal vxmlTableElement As IXMLDOMElement) As String

End Function
'SR 20/10/2005 : MAR24 - End

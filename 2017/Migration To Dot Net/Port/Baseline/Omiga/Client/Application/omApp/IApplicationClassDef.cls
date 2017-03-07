VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definitions interface for omApp.
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     07/01/99    Added LoadXXXData functions for NewPropertyDO.
'RF     10/01/99    Added LoadDeclarationData.
'SR     29/02/00    Added new value (AFFCounts) for Enum ApplicationFactFindData
'IW    16/03/00    Added LoadApplicationVerificationData
'IW    19/03/00    Added LoadApplicationCreditCardData
'SR    24/03/00     Added LoadMemoPadData
'MC     26/04/00    Added OtherInsuranceCompany methods
'JLD    09/05/00    Added AFFLatestDetailsView to applicationfactfinddata
'MC     30/05/00    Added LoadPropertyUnknownDetailsData method
' SR    15/06/00    Added LoadNewPropertyThirdPartyDetailsData & LoadFullNewPropertyDetailsData,
'JR     22/03/01    Added LoadApplicationOfferData
'INR    09/06/04    Added AFFThirdPartyData to applicationfactfinddata
'HMA    21/07/04    BMIDS758 Added LoadRemovedToECustomer
'INR    22/12/06    EP2_523 new Enum NPDFullNewPropertyDetails
'------------------------------------------------------------------------------------------

Option Explicit

Public Enum NEWPROPERTYDATA
    NPDDefault
    NPDGeneralDetails
    NPDDescriptionDetails
    NPDValuationTypeAndLocation
    NPDNewPropertyAndAddressDetails
    NPDFullNewPropertyDetails
End Enum

Public Enum NEWPROPERTYDETAILSDATA
    NPDDDefault
    NPDDChildrenOnly
End Enum

Public Enum APPLICATIONFACTFINDDATA
    AFFDefault
    AFFLoanProperty
    AFFApplicantCounts
    AFFLatestDetailsView
    AFFThirdPartyData
End Enum


Public Function LoadDeclarationData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyData( _
    Optional ByVal vClassDefType As NEWPROPERTYDATA = NPDDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyDetailsData( _
    Optional ByVal vClassDefType As NEWPROPERTYDETAILSDATA = NPDDDefault) _
    As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyDepositData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyAndAddressDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyRoomTypeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyThirdPartyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyLeaseholdData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyAddressData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNewPropertyAddressDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadVendorData() As FreeThreadedDOMDocument40
End Function

Public Function LoadApplicationData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoanPropertyDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadApplicationFactFindData( _
    Optional ByVal vClassDefType As APPLICATIONFACTFINDDATA = AFFDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadApplicationLatestData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerRelationshipData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerRoleData() As FreeThreadedDOMDocument40
End Function

Public Function LoadDependantDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFinancialSummaryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadOtherResidentData() As FreeThreadedDOMDocument40
End Function

Public Function LoadOtherResidentDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadVerificationData() As FreeThreadedDOMDocument40
End Function

Public Function LoaduserHistoryData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationStageData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationPriorityData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationReviewHistoryData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationVerificationData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationCreditCardData() As FreeThreadedDOMDocument40

End Function

Public Function LoadMortgageCalcDetailsData() As FreeThreadedDOMDocument40

End Function

Public Function LoadFindOmigaBusinessForCustomerData() As FreeThreadedDOMDocument40

End Function

Public Function LoadCustomerPackageData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationPackageData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationLockData() As FreeThreadedDOMDocument40

End Function

Public Function LoadCustomerLockApplicationLockData() As FreeThreadedDOMDocument40

End Function

Public Function LoadNewLoanData() As FreeThreadedDOMDocument40
End Function

Public Function LoadSharedOwnershipDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMemoPadData() As FreeThreadedDOMDocument40
End Function

Public Function LoadOtherInsuranceCompanyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadPropertyUnknownDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullNewPropertyDetails() As FreeThreadedDOMDocument40

End Function

Public Function LoadFindLegacyBusinessForCustomerData() As FreeThreadedDOMDocument40

End Function

Public Function LoadMenuItemStatusData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationEnquiryData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationOfferData() As FreeThreadedDOMDocument40

End Function

Public Function LoadApplicationDataForCostModelling() As FreeThreadedDOMDocument40
End Function

'BMIDS758
Public Function LoadRemovedToECustomer() As FreeThreadedDOMDocument40
End Function
'MAR24
Public Function LoadApplicationUnderwritingData() As FreeThreadedDOMDocument40
End Function

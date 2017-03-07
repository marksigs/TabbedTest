VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IApplicationClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definitions for omApp.
'
'Dependencies:  AddressClassDef, IAddressClassDef
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/99    Added LoadXXXData functions for NewPropertyDO.
' RF    10/01/99    Added LoadDeclarationData.
' RF    27/01/00    Added BCDECLARATIONNOTES to LoadApplicationFactFindData.
' SR    03/02/00    Removed fields IntermediaryGuid from class definition
' RF    15/02/00    Spec change to fields returned by NewPropertyDO.GetNewPropertyDescription
'                   and NewPropertyDO.GetNewPropertyGeneral.
' RF    15/02/00    Spec change: NewPropertyDO.GetNewPropertyDescription to return
'                   NUMBEROFOUTBUILDINGS.
' RF    15/02/00    Added methods NewPropertyBO/DO.GetValuationTypeAndLocation.
' SR    29/02/00    Modified LoadApplicationFactFindData to cater to the new
'                   value for enumeration (AFFCounts)
' SR    13/03/00    Modified - LoadApplicationStageData - added new field DateTime
' IW    16/03/00    Added LoadApplicationVerificationData
' IW    19/03/00    Added LoadApplicationCreditCardData
' SR    24/03/00    New method LoadMemoPadData
' MC    26/04/00    Added OtherInsuranceCompany methods
' IW    26/04/00    Moved StructureUseageChangedInd to General details
' SR    08/05/00    Added three new columns (CLIENTBCINDICATOR, CLIENTPPINDICATOR, APPLICATIONNOTES)
'                   to LoadApplicationFactFindData
' JLD   09/05/00    Added AFFLatestDetailsView to LoadApplicationFactFindData to return only the
'                   fields held in the ApplicationLatestDetails view. Updated LoadApplicationLatestData to use this.
' MC    11/05/00    Removed DirectoryGUID and VendorName from LoadVendorData
' MC    30/05/00    Added LoadPropertyUnknownDetailsData method
' BG    02/06/00    Added INTERESTONLYAMOUNT, OTHERLOANPURPOSE, ADDITIONALBORROWINGAMOUNT to LoadNewLoanData.
' IVW   08/09/00    Added LoadNewPropertyAndAddressDetailsData
' SR    15/06/00    Added LoadNewPropertyThirdPartyDetailsData & LoadFullNewPropertyDetailsData,
'                    modified LoadSharedOwnershipDetailsData
' JLD   31/07/00    Added INTENDEDAPPLICANTCOUNT to default view.
' JLD   02/08/00    SYS1379 Modified FindLegacyBusinessForCustomerData
' JLD   04/08/00    SYS1390 - LoadApplicationCreditCardData had wrong type for cardnumber
' BG    17/08/00    Added if statement to fix SYS0621 - not pulling back radio button settings on DC220.
' BG    12/09/00    Added SHAREDOWNERSHIPINDICATOR tag to LoadSharedOwnershipDetailsData ClassDef.
' DJP   24/01/01    DJP Added STAGESEQUENCENO and CASESEQUENCENO to LoadApplicationStageData and
'                   Added STAGESEQUENCENO and CASESEQUENCENO to LoadFindOmigaBusinessForCustomerData
' JLD   05/02/01    Added stagenumber to LoadFindOmigaBusinessForCustomerData
' JLD   22/02/01    SYS1879 added new fields to LoadApplicationFactFindData.
' APS   28/02/01    SYS1993 APPLICATIONPRIORITY column has changed to APPLICATIONPRIORITYVALUE
' PSC   08/03/01    SYS1879 Amend LoadApplicationData to make TypeOfBuyer a combo and
'                   Amend LoadApplicationPriorityData to make ApplicationPriorityValue a combo
' JLD   08/03/01    SYS1879 Amend LoadCustomerRoleData to bring back the combo text
' APS   11/03/01    SYS2030 Added ApplicationPriorityValue to LoadApplicationPriorityData()
' GD    15/03/01    SYS2047 Altered LoadNewPropertyData to return combo text
' GD    22/03/2001  SYS2047 OTHERSYSTEMACCOUNTNUMBER added to LoadApplicationData
' JR    22/03/01    SYS2048 Added LoadApplicationOfferData
' SR    11/06/01    SYS2362 Modified 'LoadFindOmigaBusinessForCustomerData'
' JR    25/06/01    Omiplus24 Modified 'LoadNewPropertyAddressData' to include Country/Area Code
' JR    24/07/01    SYS2505 - Modified 'LoadNewPropertyAddressData', comment out ref. to Country/Area Code
' JR    24/08/01    Omiplus24 Modified 'LoadNewPropertyAddressData', re-addedd ref. to Country/Area Code and USAGE
' BG    23/12/01    Changed OFFERSEQUENCENO to a primary key in Loadapplicationofferdata.
' DRC   24/04/02    SYS3017 Change of key status to several fields in LoadCustomerRelationshipData
' SG    27/05/02    SYS4767 MSMS to Core integration

'------------------------------------------------------------------------------------------------------------

Option Explicit

Implements IApplicationClassDef

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
' Prog  Date        Description
' MV    25/05/2002  BMIDS00013 - BM044 - Amended IApplicationClassDef_LoadApplicationFactFindData with BMDECLARATIONIND
' MV    10/06/2002  BMIDS00032 - BM078 - Amended IApplicationClassDef_LoadApplicationFactFindData with New columns
'                   FurtherAdvanceTermInd and FurtherAdvanceRepTypeInd
' GD    04/07/02    BMIDS00165 - Amended LoanNewPropertyData
' MO    19/07/02    BMIDS00496 - Modified LoadMemoPadData so that the memopad id, is generated as a sequence number for
'                       application and application fact find not as a unique number
' MO    23/09/2002  INWP1 - BMIDS00502 - Modified LoadApplicationFactFindData to include introducer data
' MO    17/10/2002  CC006 - BMIDS00663 - Modified LoadApplicationFactFindData to include introducer data
' GHun  18/11/2002  BMIDS00980 - Amended LoadFindOmigaBusinessForCustomerData to include AccountNumber
' DM    19/11/2002  BMIDS00935 Added NOLOCK hint to the following methods:
'                   IApplicationClassDef_LoadNewPropertyDepositData, IApplicationClassDef_LoadApplicationOfferData,
'                   IApplicationClassDef_LoadOtherResidentData, IApplicationClassDef_LoadVerificationData,
'                   IApplicationClassDef_LoadMemoPadData, IApplicationClassDef_LoadApplicationPriorityData
' SR    20/11/2002  BMIDS00500 [Core SYS5040] - modified method IApplicationClassDef_LoadApplicationLockData
'                   Add SQLNOLOCK hint at table level
' SR    02/12/2003  BM0209 - new method LoadApplicationDataForCostModelling()
' GD    31/01/2003  BM0256 - Added SQLNOLOCK to CUSTOMERROLE table in IApplicationClassDef_LoadCustomerRoleData()
' RF    12/05/2003  BM0536 - Added SQLNOLOCK to IApplicationClassDef_LoadUserHistoryData
' RF    14/05/2003  BM0558 - Changed value of SQLNOLOCK element from "Yes" to "TRUE" in 2 places
'                   for consistency
' PJO   27/06/2003  BM0563 - New fields added to Application Table
' MC    12/05/2004  BMIDS746    SourceReason Field added to NewPropertyDeposit DBSchema
' INR   14/05/2004  BMIDS757 - new introducer fields for LoadApplicationFactFindData
' MC    24/05/2004  BMIDS747 - CC063 Verification db table extended to stored additional information
' KRW   25/05/2004  BMIDS762 - FSA Validation
' INR   01/06/2004  BMIDS744 - OptOutIndicator, new field for GetCurrentFullBureauResults
' KRW   14/06/2004  BMIDS774 BMIDS776 - New fields for Regulation changes DC010 DC013
' INR   01/06/2004  BMIDS744 - OptOutIndicator  Required in the default call to LoadApplicationFactFindData
'                   as well for updateApplication from DC010
' MC    14/06/2004  bmids763    APPLICATIONDATE added to the classDEF routine LoadApplicationDataForCostModelling()
' SR    17/06/2004  BMIDs772 - Modified 'LoadFinancialSummaryData' - added new column 'REGULAROUTGOINGSINDICATOR'
' HMA   21/07/2004  BMIDS758 - Added 'LoadRemovedToECustomer'
' JD    03/08/2004  BMIDS830 add MAINRESIDENCEIND to LoadApplicationFactFindData
' SR    11/08/2004  BMIDS815 Modified method 'LoadApplicationData'- add new column 'LASTACCOUNTDOWNLOADDATE'
' SR    13/08/2004  BMIDS815 Modified method 'LoadApplicationDataForCostModelling'- add new column 'LASTACCOUNTDOWNLOADDATE'
' GHun  10/09/2004  BMIDS871 Amended LoadApplicationFactFindData to fix the spelling of ApplicationReceivedDate
' HMA   13/09/2004  BMIDS864 Modified LoadVerificationData
' JD    22/03/05    BMIDS975 added FutherAdvanceTermInd to LoadApplicationDataForCostModelling
'------------------------------------------------------------------------------------------
'MARS Specific History:
' HMA   25/07/2005  MAR18    Add LevelOfAdvice to LoadApplicationDataForCOstModelling
' PSC   22/09/2005  MAR32    Add SLAExpiryDate to LoadApplication
' SR    20/10/2005  MAR253   Modified LoadApplicationUnderwritingData
' PSC   04/11/2005  MAR427   Add OTHERSYSTEMCUSTOMERNUMBER to LoadFindOmigaBusinessForCustomerData
'------------------------------------------------------------------------------------------
' EPSOM Specific History:
' SAB   06/04/2006  EP348   Remove INGESTIONAPPLICATIONNUMBER from loadApplicationFactFindData
' SAB   13/13/2006  EP388   MAR1300
' IK    19/04/2006  EP371   Add SELFCERTIND to LoadApplicationFactFindData
' AW    09/05/2006  EP485   Add BROKERPROCFEEREFUND, PACKAGEDVALUATIONFEEREBATE, PACKAGEDVALUATIONFEE to LoadApplicationFactFindData
' PB    06/06/06    EP696   MAR1700 Add valuationtype to LoadNewPropertyData
' SAB   06/07/2006  EP945   Return PACKAGERINDIVIDUALNAME in LoadApplicationFactFindData
' PB    20/09/2006  EP1117  Return PACKAGERINDIVIDUALTELEPHONE and PACKAGERINDIVIDUALEMAIL
' GHun  08/11/2006  EP2_9   Changed LoadApplicationData to include IsLegalRepToBeUsed
' GHun  21/11/2006  EP2_19  Changed LoadNewPropertyData to include PreEmptionEndDate
' MHeys 22/11/2006  E2CR35  Changed LoadNewPropertyData to include new NewProperty fields
' MCh   28/11/2006  EP2_127 NumberOfDependants added to LoadApplicationData
' AShaw 30/11/2006  EP2_2   LoanProperty details - New columns
' MHeys 07/12/2006  EP2_339 Build Patch required EP2_127 removing
' GHun  18/12/2006  EP2_56  Added SequenceNumber, CustomerRole and Type to LoadRemovedToECustomer
' INR   20/12/2006  EP2_523 new Enum NPDFullNewPropertyDetails
' INR   09/01/2007  EP2_701 PRODUCTSCHEME added to LoadApplicationFactFindData
' INR   19/01/2007  EP2_677 changed allowed vClassDefType in LoadNewPropertyData
' SR    23/01/2007  EP2_858 - modified LoadApplicationFactFindData - new column 'ApplicationIncomeStatus'
' INR   23/01/2007  EP2_645 - modified LoadApplicationFactFindData - new column 'SELFCERTREASON & SELFCERTREASONDETAILS'"
' INR   25/01/2007  EP2_780 added SPECIALSCHEME to LoadApplicationDataForCostModelling
' IK    07/02/2007  EP2_1179 PRODUCTSCHEME added to LoadApplicationFactFindData, AFFLatestDetailsView
' LDM   19/02/2007  EP2_1165 Added DIRECTFINANCIALBENEFITIND to LoadNewLoanData.
' SR    01/03/2007  EP2_1196 modfied LoadApplicationData
' AW    19/03/2007  EP2_1897  Added QUOTATIONNUMBER to APPLICATIONOFFER
'------------------------------------------------------------------------------------------

Private Function IApplicationClassDef_LoadApplicationLockData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationLockData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONLOCK"
    ' RF 14/5/03 BM0558 Start
    'strXML = strXML & "<SQLNOLOCK>Yes</SQLNOLOCK>" ' SR 20/11/2002 : BMIDS00500 [Core SYS5040]
    strXML = strXML & "<SQLNOLOCK>TRUE</SQLNOLOCK>"
    ' RF 14/5/03 BM0558 End
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>UNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOCKDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MACHINEID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFLOCK<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationLockData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadApplicationOfferData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationOfferData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONOFFER"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>OFFERSEQUENCENO<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>OFFERISSUEDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OFFEREXPIRYDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SIGNEDOFFERRECEIVED<TYPE>dbdtDateTime</TYPE></OTHERS>"
    'AW    19/03/07  EP2_1897
    strXML = strXML & "<OTHERS>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationOfferData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationPackageData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationPackageData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONPACKAGE"
    strXML = strXML & "<PRIMARYKEY>PACKAGENUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COMPANYID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationPackageData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadApplicationPriorityData() As FreeThreadedDOMDocument40
'Header ----------------------------------------------------------------------------------
' description:
'   Used by ApplicationDO.
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationPriorityData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONPRIORITY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>APPPRIORITYSEQUENCENO<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>APPLICATIONPRIORITYVALUE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ApplicationPriority</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATETIMEASSIGNED<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationPriorityData = _
    objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function IApplicationClassDef_LoadApplicationReviewHistoryData() As FreeThreadedDOMDocument40
'Header ----------------------------------------------------------------------------------
' description:
'   Used by ApplicationDO.
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationReviewHistoryData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONREVIEWHISTORY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>REVIEWREASON<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>REVIEWDATETIME<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>REVIEWCOMMENTS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REVIEWUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REVIEWUNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERRIDENBYUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERRIDENBYUNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERRIDEREASON<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERRIDECOMMENTS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERRIDEDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationReviewHistoryData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
    
End Function

Private Function IApplicationClassDef_LoadApplicationStageData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by ApplicationDO.
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationStageData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONSTAGE"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CASESEQUENCENO<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>STAGESEQUENCENO<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAGENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationStageData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationVerificationData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationVerificationData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>APPLICATIONVERIFICATION"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>RESIDENCYOKINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>AGEOKINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LOANAMOUNTINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LOANTERMINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>PRIMARYRESIDENCEINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>KNOWNPROPERTYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationVerificationData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IApplicationClassDef_LoadApplicationCreditCardData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationCreditCardData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>APPLICATIONCREDITCARD"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>FEEPAYMENTMETHOD"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>NAMEOFCARDHOLDERS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CARDTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CreditCardType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CARDNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ISSUENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>EXPIRYDATE"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationCreditCardData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IApplicationClassDef_LoadCustomerLockApplicationLockData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IApplicationClassDef_LoadCustomerLockApplicationLockData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CUSTOMERLOCKAPPLICATIONLOCK"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"

    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadCustomerLockApplicationLockData = _
            objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadCustomerPackageData() As FreeThreadedDOMDocument40
 
    Const strFunctionName As String = "IApplicationClassDef_LoadCustomerPackageData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CUSTOMERPACKAGE"
    strXML = strXML & "<PRIMARYKEY>PACKAGENUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadCustomerPackageData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadDeclarationData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by ApplicationDO.
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadDeclarationData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>DECLARATION"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>DECLARATIONTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>QUESTIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>RESPONSETODECLARATIONQUESTION<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadDeclarationData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadFindLegacyBusinessForCustomerData() As FreeThreadedDOMDocument40
'Header------------------------------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : DOM Document
' **********************************************************************************************
    Const strFunctionName As String = "IApplicationClassDef_LoadFindLegacyBusinessForCustomerData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FINDLEGACYBUSINESSFORCUSTOMER"
    strXML = strXML & "<PRIMARYKEY>OTHERSYSTEMCUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATECREATED<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PACKAGENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)

    Set IApplicationClassDef_LoadFindLegacyBusinessForCustomerData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing


End Function

Private Function IApplicationClassDef_LoadFullNewPropertyDetails() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IApplicationClassDef_LoadFullNewPropertyDetails"

    Dim strXML As String
    Dim objAddressClassDef As IAddressClassDef
    Dim objXmlAssist As New XMLAssist

    Set objAddressClassDef = New AddressClassDef

    strXML = strXML & "<TABLENAME>FULLNEWPROPERTYDETAILS"
    'EP2_523
    strXML = strXML & IApplicationClassDef_LoadNewPropertyData(vClassDefType:=NPDFullNewPropertyDetails).xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyRoomTypeData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyDepositData.xml
    strXML = strXML & IApplicationClassDef_LoadSharedOwnershipDetailsData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyLeaseholdData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyAddressData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadFullNewPropertyDetails = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objAddressClassDef = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadLoanPropertyDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadLoanPropertyDetailsData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>LOANPROPERTYDETAILS"
    strXML = strXML & IApplicationClassDef_LoadNewLoanData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyData.xml
    strXML = strXML & IApplicationClassDef_LoadSharedOwnershipDetailsData.xml
    strXML = strXML & IApplicationClassDef_LoadApplicationFactFindData(vClassDefType:=AFFLoanProperty).xml
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadLoanPropertyDetailsData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadMemoPadData() As FreeThreadedDOMDocument40
     
    Const strFunctionName As String = "IApplicationClassDef_LoadMemoPadData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>MEMOPAD"
    
    ' MO    19/07/02    BMIDS00496 - Moved memopadid so that it is below application and fact find number
    'strXML = strXML & "<PRIMARYKEY>MEMOPADID<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' MO    19/07/02    BMIDS00496
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>MEMOPADID<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACTIONNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ENTRYDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MEMOENTRY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SCREENDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SCREENREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ENTRYTYPE<TYPE>dbdtComboId</TYPE><COMBO>MemoPadEntryType</COMBO></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadMemoPadData = _
                        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function


Private Function IApplicationClassDef_LoadMenuItemStatusData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadMenuItemStatusData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>MENUITEMSTATUS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MENUITEMID<TYPE>dbdtComboId</TYPE><COMBO>MenuItemId</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtComboId</TYPE><COMBO>MenuStatus</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadMenuItemStatusData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function IApplicationClassDef_LoadMortgageCalcDetailsData() As FreeThreadedDOMDocument40
     
    Const strFunctionName As String = "IApplicationClassDef_LoadMortgageCalcDetailsData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>MORTGAGECALCULATORDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PURCHASEPRICE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TERMMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TERMYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadMortgageCalcDetailsData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function IApplicationClassDef_LoadNewLoanData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadNewLoanData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>NEWLOAN"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIRASAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIRASELIGIBILITY<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtComboId</TYPE><COMBO>PurposeOfLoan</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>REPAYMENTTYPE<TYPE>dbdtComboId</TYPE><COMBO>RepaymentType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>RATECHANGENOTIFICATION<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SAMEMORTGAGETYPE<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SAMEFINISHINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SOLICITORINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TERMMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TERMYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERESTONLYAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERLOANPURPOSE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALBORROWINGAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTFINANCIALBENEFITIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadNewLoanData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadNewPropertyAndAddressDetailsData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyAndAddressDetailsData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>NEWPROPERTY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>GROUNDRENT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNEXPIREDTERMOFLEASEYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadNewPropertyAndAddressDetailsData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyData( _
    Optional ByVal vClassDefType As NEWPROPERTYDATA = NPDDefault) As FreeThreadedDOMDocument40

' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO.
' pass:
'   vClassDefType:
'       NPDDefault               - everything from NEWPROPERTY table.
'       NPDGeneralDetails        - general details only.
'       NPDDescriptionDetails    - description details only.
'       NPDValuationTypeAndLocation - PROPERTYLOCATION and VALUATIONTYPE only.
' return:
'------------------------------------------------------------------------------------------
'*************************************************************************************************
'NOTE : Additions to the NEWPROPERTY table will require LOANPROPERTYDETAILS view being updated too
'*************************************************************************************************
    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    'EP2_523 check for vClassDefType = NPDFullNewPropertyDetails
    If vClassDefType = NPDDefault Or vClassDefType = NPDDescriptionDetails Or vClassDefType = NPDFullNewPropertyDetails Then
    
        strXML = strXML & "<OTHERS>NUMBEROFSTOREYS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NEWPROPERTYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>YEARBUILT<TYPE>dbdtInt</TYPE></OTHERS>"
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>TENURETYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TENURETYPE<TYPE>dbdtComboId</TYPE><COMBO>PropertyTenure</COMBO></OTHERS>"
        
        strXML = strXML & "<OTHERS>FLOORNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATSERVICEDBYLIFTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>BUILDINGCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ROOFCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALLANDTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PARKINGARRANGEMENTSTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>LOCATIONOFGARAGETYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBEROFGARAGES<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESIDENTIALPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESIDENTIALPERCENTAGEDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>TYPEOFPROPERTY<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TYPEOFPROPERTY<TYPE>dbdtComboId</TYPE><COMBO>PropertyType</COMBO></OTHERS>"
        
        strXML = strXML & "<OTHERS>GUARANTEEEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>HOUSEBUILDERSGUARANTEE<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBEROFOUTBUILDINGS<TYPE>dbdtInt</TYPE></OTHERS>"
        
        'SG 27/05/02 SYS4767
        strXML = strXML & "<OTHERS>DATEOFENTRY<TYPE>dbdtDate</TYPE></OTHERS>"
        
    End If
    
    
    If vClassDefType = NPDNewPropertyAndAddressDetails Then
    
        strXML = strXML & "<OTHERS>BUILDINGCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>HOUSEBUILDERSGUARANTEE<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBEROFSTOREYS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PROPERTYLOCATION<TYPE>dbdtInt</TYPE></OTHERS>"
        
        strXML = strXML & "<OTHERS>ROOFCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>TENURETYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TENURETYPE<TYPE>dbdtComboId</TYPE><COMBO>PropertyTenure</COMBO></OTHERS>"
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>TYPEOFPROPERTY<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TYPEOFPROPERTY<TYPE>dbdtComboId</TYPE><COMBO>PropertyType</COMBO></OTHERS>"
        
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>VALUATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VALUATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>ValuationType</COMBO></OTHERS>"
        
        strXML = strXML & "<OTHERS>YEARBUILT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ENTIRELYOWNUSEINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESIDENTIALUSEONLYINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GROUNDRENT<TYPE>dbdtDouble</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>UNEXPIREDTERMOFLEASEYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
         
    End If
        
    'EP2_523 check for vClassDefType = NPDFullNewPropertyDetails
    If vClassDefType = NPDDefault Or vClassDefType = NPDGeneralDetails Or vClassDefType = NPDFullNewPropertyDetails Then
        
        strXML = strXML & "<OTHERS>APPLICANTTOOCCUPYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CURRENTLYOCCUPIEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EXLOCALAUTHORITYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FIRSTOCCUPIERINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GRANTAPPLIEDFORIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GRANTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>LETORPARTLETINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESALERESTRICTIONSIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESALERESTRICTIONSDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SELFBUILDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>STRUCTUREUSAGECHANGEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>STRUCTUREUSAGECHANGE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ENTIRELYOWNUSEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESIDENTIALUSEONLYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EXTRASINCLUDEDIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EXTRASINCLUDEDDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
        
        ' GD    04/07/02    BMIDS00165 - Amended LoanNewPropertyData START
        strXML = strXML & "<OTHERS>NUMBEROFOCCUPANTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TENANCYTYPE<TYPE>dbdtComboId</TYPE><COMBO>TenancyType</COMBO></OTHERS>"
        strXML = strXML & "<OTHERS>MONTHLYRENTALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RENTALINCOMESTATUS<TYPE>dbdtComboId</TYPE><COMBO>RentalIncomeStatus</COMBO></OTHERS>"
        strXML = strXML & "<OTHERS>EXCESSMONTHLYRENTALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
        ' GD    04/07/02    BMIDS00165 - Amended LoanNewPropertyData END
        'KRW    14/06/2004  BMIDS776
        strXML = strXML & "<OTHERS>FAMILYLET<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FAMILYMEMBER<TYPE>dbdtComboId</TYPE><COMBO>FamilyLetMember</COMBO></OTHERS>"
        strXML = strXML & "<OTHERS>CHANGEOFPROPERTY<TYPE>dbdtBoolean</TYPE></OTHERS>"   'MAR1300 GHun
    End If
    'EP2_532
    If vClassDefType = NPDDefault Or vClassDefType = NPDGeneralDetails Then
        strXML = strXML & "<OTHERS>PREEMPTIONENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"   'EP2_19 GHun
        'MAH 22/11/2006 E2CR35 Start
        strXML = strXML & "<OTHERS>BUSINESSCONNECTION<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FULLMARKETVALUE<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FINANCIALINCENTIVES<TYPE>dbdtBoolean</TYPE></OTHERS>"
    End If
    'EP2_677
    If vClassDefType = NPDDefault Or vClassDefType = NPDGeneralDetails Then
        strXML = strXML & "<OTHERS>NOFLATSINBLOCK<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATABOVECOMMERCIAL<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGRICULTURALRESTRICTIONS<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OCCUPYPROPERTY<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'MAH 22/11/2006 E2CR35 End
    End If
    
    ' PB EP696/MAR1700 add valuationtype to general details
    If vClassDefType = NPDGeneralDetails Then
        strXML = strXML & "<OTHERS>VALUATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>ValuationType</COMBO></OTHERS>"
    End If
    ' PB EP696/MAR1700 End
    
    'EP2_523 check for vClassDefType = NPDFullNewPropertyDetails
    If vClassDefType = NPDDefault Or vClassDefType = NPDValuationTypeAndLocation Or vClassDefType = NPDFullNewPropertyDetails Then
    
        'GD changed SYS2047
        'Original strXML = strXML & "<OTHERS>VALUATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VALUATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>ValuationType</COMBO></OTHERS>"
        
        strXML = strXML & "<OTHERS>PROPERTYLOCATION<TYPE>dbdtComboId</TYPE><COMBO>PropertyLocation</COMBO></OTHERS>"
    
    End If
    
    'BG 17/08/00 Added this if statement to fix SYS0621 - not pulling back radio button settings.
    If vClassDefType = NPDDescriptionDetails Then
    
        strXML = strXML & "<OTHERS>ENTIRELYOWNUSEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RESIDENTIALUSEONLYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    End If
    
    'EP2_523 check for vClassDefType = NPDFullNewPropertyDetails
    If vClassDefType = NPDDefault Or vClassDefType = NPDFullNewPropertyDetails Then
    
        strXML = strXML & "<OTHERS>ANYOTHERRESIDENTSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        'GD changed SYS 2047
        'Original strXML = strXML & "<OTHERS>DESCRIPTIONOFPROPERTY<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DESCRIPTIONOFPROPERTY<TYPE>dbdtComboId</TYPE><COMBO>PropertyDescription</COMBO></OTHERS>"
        strXML = strXML & "<OTHERS>DISCOUNTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FAMILYSALEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VALUATIONPRICE<TYPE>dbdtLong</TYPE></OTHERS>"
    End If
    'EP2_523
    If vClassDefType = NPDDefault Then
        ' EP2_2 Add new fields
        strXML = strXML & "<OTHERS>DATEOFORIGINALVALUATION<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYDISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYFLATNONAME<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYPOSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYSTREET<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYTOWN<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYCOUNTY<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ORIGINALVALUERCOMPANYCOUNTRY<TYPE>dbdtString</TYPE></OTHERS>"

    End If
    
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadNewPropertyDetailsData( _
    Optional ByVal vClassDefType As NEWPROPERTYDETAILSDATA = NPDDDefault) _
    As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO.
' pass:
'   vClassDefType:
'       NPDDDefault         - everything from NEWPROPERTY table, plus all the data from
'                             child tables NEWPROPERTYROOMTYPE, NEWPROPERTYLEASEHOLD and
'                             NEWPROPERTYDEPOSIT.
'       NPDDChildrenOnly    - all the data from child tables NEWPROPERTYROOMTYPE,
'                             NEWPROPERTYLEASEHOLD and NEWPROPERTYDEPOSIT.
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYDETAILS"
    Select Case vClassDefType
        Case NPDDDefault
            strXML = strXML & IApplicationClassDef_LoadNewPropertyData.xml
    End Select
    strXML = strXML & IApplicationClassDef_LoadNewPropertyDepositData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyRoomTypeData.xml
    strXML = strXML & IApplicationClassDef_LoadNewPropertyLeaseholdData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyDetailsData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadNewPropertyDepositData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyDepositData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYDEPOSIT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>NPDEPOSITSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>SOURCEOFFUNDING<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    '*=MC:BMIDS746 - CC062 Deposit source Reason
    strXML = strXML & "<OTHERS>SOURCEREASON<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyDepositData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyRoomTypeData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO.
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyRoomTypeData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYROOMTYPE"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ROOMTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>NUMBEROFROOMS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyRoomTypeData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyLeaseholdData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyLeaseholdData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYLEASEHOLD"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DATELEASESTARTED<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GROUNDRENT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INCREASEDRENTDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALTERMOFLEASEYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SERVICECHARGE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNEXPIREDTERMOFLEASEYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyLeaseholdData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyAddressDetailsData() _
    As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   NEWPROPERTYADDRESS and ADDRESS bounded by GETNEWPROPERTYADDRESS.
'   Used by NewPropertyDO.
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyAddressDetailsData"

    Dim strXML As String
    Dim objAddressClassDef As IAddressClassDef
    Dim objXmlAssist As New XMLAssist

    Set objAddressClassDef = New AddressClassDef

    strXML = strXML & "<TABLENAME>GETNEWPROPERTYADDRESS"
    strXML = strXML & IApplicationClassDef_LoadNewPropertyAddressData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyAddressDetailsData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objAddressClassDef = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyAddressData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO.
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyAddressData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYADDRESS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ARRANGEMENTSFORACCESS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCESSCONTACTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCESSTELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERARRANGEMENTSFORACCESS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set IApplicationClassDef_LoadNewPropertyAddressData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadNewPropertyThirdPartyData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadNewPropertyThirdPartyData"

    Dim strXML As String
    Dim objAddressClassDef As IAddressClassDef
    Dim objTPClassDef As IThirdPartyClassDef
    Dim objXmlAssist As New XMLAssist

    Set objAddressClassDef = New AddressClassDef
    Set objTPClassDef = New ThirdPartyClassDef

    strXML = strXML & "<TABLENAME>NEWPROPERTYTHIRDPARTYDETAILS"
    strXML = strXML & IApplicationClassDef_LoadVendorData.xml
    strXML = strXML & objTPClassDef.LoadThirdPartyData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadNewPropertyThirdPartyData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objAddressClassDef = Nothing
    Set objTPClassDef = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadOtherInsuranceCompanyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadOtherInsuranceCompanyData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>OTHERINSURANCECOMPANY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INSURANCECOMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>POLICYNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RENEWALDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadOtherInsuranceCompanyData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadPropertyUnknownDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadPropertyUnknownDetailsData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>NEWPROPERTY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ENTIRELYOWNUSEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RESIDENTIALUSEONLYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ROOFCONSTRUCTIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadPropertyUnknownDetailsData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadSharedOwnershipDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationClassDef_LoadSharedOwnershipDetailsData"
     
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>SHAREDOWNERSHIPDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
'    strXML = strXML & "<OTHERS>EQUITYAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYRENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SHAREDPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SHAREDOWNERSHIPINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SHAREDOWNERSHIPTYPE<TYPE>dbdtComboId</TYPE><COMBO>SharedOwnershipType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>SHAREDAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadSharedOwnershipDetailsData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IApplicationClassDef_LoadUserHistoryData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Class definition for table 'USERHISTORY'
' pass:
' return:
' history:
' RF    12/05/2003  BM0536 - Added SQLNOLOCK. Note that this appears at a different
'                   level in FindListXXX methods than in other ones!
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "IApplicationClassDef_LoadUserHistoryData"
    
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>USERHISTORY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>USERHISTORYDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<SQLNOLOCK>TRUE</SQLNOLOCK>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadUserHistoryData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
End Function

Private Function IApplicationClassDef_LoadVendorData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Used by NewPropertyDO.
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IApplicationClassDef_LoadVendorData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NEWPROPERTYVENDOR"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadVendorData = _
        objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

' PJO 27/06/2003 BM0563 - New fields added to Application Table
Private Function IApplicationClassDef_LoadApplicationData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATION"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>PACKAGENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFBUYER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TypeOfBuyerNewLoan</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>"
    'GD 22/03/2001
    strXML = strXML & "<OTHERS>OTHERSYSTEMACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
' PJO 27/06/2003 BM0563 - New fields added to Application Table
    strXML = strXML & "<OTHERS>PIPELINEBTLNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PIPELINEBTLBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMPLETEDBTLNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMPLETEDBTLBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PIPELINERESIDENTIALBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMPLETEDRESIDENTIALBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PIPELINEMANDATEBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PIPELINEAGGREGATEDLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INGESTIONAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" ' PCT EP15
'PJO BM0563 End
    'SR 11/08/2004 : BMIDS815 - Add new column 'LASTACCOUNTDOWNLOADDATE'
    strXML = strXML & "<OTHERS>LASTACCOUNTDOWNLOADDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    ' PSC 22/09/2005 MAR32
    strXML = strXML & "<OTHERS>SLAEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ISLEGALREPTOBEUSED<TYPE>dbdtBoolean</TYPE></OTHERS>" 'EP2_9 GHun
    strXML = strXML & "<OTHERS>NUMBEROFDEPENDANTS<TYPE>dbdtLong</TYPE></OTHERS>" 'EP2_127 MCh
    'SR 01/03/2007 : EP2_1196
    strXML = strXML & "<OTHERS>PROCFEETOCUST<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ProcFeeToCust</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>PROCFEETOCUSTAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    'SR 01/03/2007 : EP2_1196 - End
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationFactFindData( _
    Optional ByVal vClassDefType As APPLICATIONFACTFINDDATA = AFFDefault) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationFactFindData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    Dim strLatestDetailsView As String
    
    strXML = strXML & "<TABLENAME>APPLICATIONFACTFIND"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    'The ApplicationLatestDetailsView consists of the majority of applicationfactfind so
    'create a subset of this for use later.
    
    strLatestDetailsView = strXML
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ACTIVEQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONCURRENCY<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ACCEPTEDQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ADDITIONALINTERMEDIARYDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ATTITUDETOBORROWINGSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>CREDITSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>CREDITSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>CURRENTTOTALLOANS<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>DIRECTINDIRECTBUSINESS<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>Direct/Indirect</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>FUTURETOTALLOANS<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCINGAGENTREF<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MARKETINGSOURCE<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>MarketingSource</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MAXIMUMBORROWINGAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MIRASOPTION1AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MIRASOPTION2AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MIRASOPTION3AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>NEXTQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>NUMBEROFAPPLICANTS<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>NUMBEROFGUARANTORS<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>OUTSTANDINGLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONPACKAGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>QUICKQUOTEQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>RECOMMENDEDQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>SPECIALSCHEME<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>SpecialSchemes</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>TOTALMONTHLYREPAYMENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>NATUREOFLOAN<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>NatureOfLoan</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>GROUPCONNECTIONINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>BCDECLARATIONNOTES<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONAPPROVALDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONAPPROVALUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONAPPROVALUNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONRECOMMENDEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONRECOMMENDEDUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONRECOMMENDEDUNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>BMDECLARATIONIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>FURTHERADVANCETERMIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>FURTHERADVANCEREPTYPEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'MO - 23/09/2002 - INWP1 - BMIDS00502 - Start
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERIDLEVEL1<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERIDLEVEL2<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERIDLEVEL3<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERMCCBLEVEL1<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERMCCBLEVEL2<TYPE>dbdtString</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERMCCBLEVEL3<TYPE>dbdtString</TYPE></OTHERS>"
    'MO - 23/09/2002 - INWP1 - BMIDS00502 - End
    'MO - 17/10/2002 - CC006 - BMIDS00663 - Start
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERCORRESPONDINDLEVEL1<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERCORRESPONDINDLEVEL2<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>INTRODUCERCORRESPONDINDLEVEL3<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'MO - 17/10/2002 - CC006 - BMIDS00663 - End
        
    'KRW - 14/06/2004  BMIDS774 BMIDS776
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>LANDUSAGE<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>KFIRECEIVEDBYALLAPPS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>LEVELOFADVICE<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>LevelOfAdvice</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>REGULATIONINDICATOR<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>RegulationIndicator</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ESTIMATEDCOMPLETIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONRECEIVEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"   'BMIDS871 GHun
    ' pct 28/03/2006 EP15 - add line
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>APPLICATIONSIGNEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"   'BMIDS871 GHun
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ADDITIONALBROKERFEE<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ADDITIONALBROKERFEEDESC<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS744 Needed in here as well for GetApplicationData & UpdateApplication from DC010
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>OPTOUTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>EXTERNALSYSTEMTPDDECLARATION<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'JD BMIDS830 add MAINRESIDENCEIND
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>MAINRESIDENCEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        
    '~ MF MAR20 Add EARLIESTCOMPLETIONDATE & ARRANGEOWNCOVERINDICATOR
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>EARLIESTCOMPLETIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>ARRANGEOWNCOVERINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    'HM MAR29
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>WRAPUPPASSWORDTAKEN<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>SALESLEADSTATUS<TYPE>dbdtComboId</TYPE>"
    strLatestDetailsView = strLatestDetailsView & "  <COMBO>SalesLeadStatus</COMBO></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>CREDITCHECKWILLINGTOPROCEED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>FTTRANSFEROFEQUITYREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    ' PCT 29/03/2006 EP15 add 3 lines
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>TOTALPROCURATIONFEE<TYPE>dbdtInt</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>PACKAGERNETWORKNAME<TYPE>dbdtString</TYPE></OTHERS>"
    
    'IK EP371
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>SELFCERTIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    'AW EP485
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>PACKAGEDVALUATIONFEE<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>PACKAGEDVALUATIONFEEREBATE<TYPE>dbdtLong</TYPE></OTHERS>"
    strLatestDetailsView = strLatestDetailsView & "<OTHERS>BROKERPROCFEEREFUND<TYPE>dbdtLong</TYPE></OTHERS>"

    Select Case vClassDefType
        Case AFFLatestDetailsView
            'SAB EP945
            strXML = strLatestDetailsView & "<OTHERS>PACKAGERINDIVIDUALNAME<TYPE>dbdtString</TYPE></OTHERS>"
            'PB EP1117
            strXML = strXML & "<OTHERS>PACKAGERINDIVIDUALTELEPHONE<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>PACKAGERINDIVIDUALEMAIL<TYPE>dbdtString</TYPE></OTHERS>"
            'EP2_1179
            strXML = strXML & "<OTHERS>PRODUCTSCHEME<TYPE>dbdtInt</TYPE></OTHERS>"
            'EP2_1202
            strXML = strXML & "<OTHERS>APPLICATIONINCOMESTATUS<TYPE>dbdtComboId</TYPE>"
             strXML = strXML & "  <COMBO>ApplicationIncomeStatus</COMBO></OTHERS>"
            strXML = strXML & "<OTHERS>SELFCERTREASON<TYPE>dbdtComboId</TYPE>"
             strXML = strXML & "  <COMBO>SelfCertReason</COMBO></OTHERS>"
            strXML = strXML & "<OTHERS>SELFCERTREASONDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
            'EP2_1202 end
       Case AFFDefault
            strXML = strLatestDetailsView
            strXML = strXML & "<OTHERS>APPLICATIONNOTES<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>CLIENTBCINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>CLIENTPPINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTENDEDAPPLICANTCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
    
            'INR BMIDS757
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL1CODE<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL1FEES<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL1INDUCEMENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL2CODE<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL2FEES<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL2INDUCEMENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL3CODE<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL3FEES<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>INTRODUCERLEVEL3INDUCEMENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
            ' KRW BMIDS762
           strXML = strXML & "<OTHERS>FIRMFSANUMBERLEVEL1<TYPE>dbdtInt</TYPE></OTHERS>"
           strXML = strXML & "<OTHERS>FIRMFSANUMBERLEVEL2<TYPE>dbdtInt</TYPE></OTHERS>"
           strXML = strXML & "<OTHERS>FIRMFSANUMBERLEVEL3<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>PRODUCTSCHEME<TYPE>dbdtInt</TYPE></OTHERS>" 'EP2_701
            'SR : EP2_858
            strXML = strXML & "<OTHERS>APPLICATIONINCOMESTATUS<TYPE>dbdtComboId</TYPE>"
             strXML = strXML & "  <COMBO>ApplicationIncomeStatus</COMBO></OTHERS>"
            'EP2_645
            strXML = strXML & "<OTHERS>SELFCERTREASON<TYPE>dbdtComboId</TYPE>"
             strXML = strXML & "  <COMBO>SelfCertReason</COMBO></OTHERS>"
            strXML = strXML & "<OTHERS>SELFCERTREASONDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
              
        Case AFFLoanProperty
            strXML = strXML & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>FURTHERADVANCETERMIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>FURTHERADVANCEREPTYPEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
            '~ MF MAR20 Add EARLIESTCOMPLETIONDATE & ARRANGEOWNCOVERINDICATOR to LoanProperty
            strXML = strXML & "<OTHERS>EARLIESTCOMPLETIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>ARRANGEOWNCOVERINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
            
        Case AFFApplicantCounts
            strXML = strXML & "<OTHERS>NUMBEROFAPPLICANTS<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>NUMBEROFGUARANTORS<TYPE>dbdtInt</TYPE></OTHERS>"
            
        'BMIDS744 Don't want to have to get everything for GetCurrentFullBureauResults
        Case AFFThirdPartyData
            strXML = strXML & "<OTHERS>OPTOUTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
         
        End Select
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationFactFindData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadApplicationLatestData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationLatestData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONLATESTDETAILS"
    strXML = strXML & IApplicationClassDef_LoadApplicationData.xml
    strXML = strXML & IApplicationClassDef_LoadApplicationFactFindData(AFFLatestDetailsView).xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationLatestData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadCustomerRelationshipData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadCustomerRelationshipData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CUSTOMERRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
' AQR SYS3017 DRC - following fields changed from PRIMARY KEY to OTHERS
    strXML = strXML & "<OTHERS>OWNERCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OWNERCUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RELTOCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RELTOCUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RELATIONSHIPTOOWNER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadCustomerRelationshipData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationUnderwritingData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationUnderwritingData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONUNDERWRITING"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNDERWRITERSDECISION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>UnderwritersDecision</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>UNDERWRITERDECISIONDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTEROFFER<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadApplicationUnderwritingData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadCustomerRoleData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadCustomerRoleData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CUSTOMERROLE"
    ' RF 14/5/03 BM0558 Start
    'GD BM0256 START
    'strXML = strXML & "<SQLNOLOCK>Yes</SQLNOLOCK>"
    strXML = strXML & "<SQLNOLOCK>TRUE</SQLNOLOCK>"
    'GD BM0256 END
    ' RF 14/5/03 BM0558 End
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERROLETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>CustomerRoleType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERORDER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadCustomerRoleData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadDependantDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadDependantDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist

    Dim objICustomerClassDef As IClassDef
    Dim objIPersonClassDef As IPersonClassDef
    
    Set objICustomerClassDef = New ClassDef
    Set objIPersonClassDef = New PersonClassdef

    strXML = strXML & "<TABLENAME>DEPENDANTDETAILS"
    strXML = strXML & objICustomerClassDef.LoadDependantData.xml
    strXML = strXML & objIPersonClassDef.LoadPersonData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadDependantDetailsData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    Set objICustomerClassDef = Nothing
    Set objIPersonClassDef = Nothing
End Function

Private Function IApplicationClassDef_LoadFinancialSummaryData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadFinancialSummaryData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>FINANCIALSUMMARY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ARREARSHISTORYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>BANKCARDINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>BANKRUPTCYHISTORYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CCJHISTORYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DECLINEDMORTGAGEINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>EXISTINGMORTGAGEINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LIFEPRODUCTINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LOANLIABILITYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>TOTALREGULAROUTGOINGS"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    'SR 17/06/2004 : BMIDS772 - New columnm REGULAROUTGOINGSINDICATOR
    strXML = strXML & "<OTHERS>REGULAROUTGOINGSINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    'SR 17/06/2004 : BMIDS772 - End
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadFinancialSummaryData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadOtherResidentData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadOtherResidentData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>OTHERRESIDENT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>OTHERRESIDENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadOtherResidentData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IApplicationClassDef_LoadOtherResidentDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadOtherResidentDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist

    Dim objIPersonClassDef As IPersonClassDef

    Set objIPersonClassDef = New PersonClassdef

    strXML = strXML & "<TABLENAME>OTHERRESIDENTDETAILS"
    strXML = strXML & IApplicationClassDef_LoadOtherResidentData.xml
    strXML = strXML & objIPersonClassDef.LoadPersonData.xml
    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadOtherResidentDetailsData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    Set objIPersonClassDef = Nothing
End Function


Private Function IApplicationClassDef_LoadVerificationData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadVerificationData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>VERIFICATION"
    ' BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber from key
    strXML = strXML & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
  
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>VERIFICATIONSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>VERIFICATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IDENTIFICATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERIDENTIFICATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    
    '*=[MC] 18/02/2004  BMIDS0747 - CC063 - Regulation change Start
    strXML = strXML & "<OTHERS>ISSUER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VERIFICATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    '*=[MC] 18/02/2004  BMIDS0747 - CC063 - Regulation change End

    strXML = strXML & "</TABLENAME>"

    Set IApplicationClassDef_LoadVerificationData = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadFindOmigaBusinessForCustomerData() As FreeThreadedDOMDocument40
'Header------------------------------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : DOM Document
' **********************************************************************************************
    Const strFunctionName As String = "IApplicationClassDef_LoadFindOmigaBusinessForCustomerData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FINDOMIGABUSINESSFORCUSTOMER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATECREATED<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PACKAGENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAGESEQUENCENO<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CASESEQUENCENO<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAGENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERROLETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>CustomerRoleType</COMBO></OTHERS>"
    
    'BMIDS00980
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00980 End
    ' PSC 04/11/2005 MAR427 Start
    strXML = strXML & "<OTHERS>OTHERSYSTEMCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    ' PSC 04/11/2005 MAR427 End
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)

    Set IApplicationClassDef_LoadFindOmigaBusinessForCustomerData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationEnquiryData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationEnquiryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATION"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONSTAGENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONSTAGENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONSTAGESEQUENCENO<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONLOCKED<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONPRIORITYVALUE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    'strXML = strXML & "<COMBO>TypeOfMortgage</COMBO></OTHERS>"
    'strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>APPLICATIONSTAGE<TYPE>dbdtString</TYPE></OTHERS>"
    
    Set xmlClassDefDoc = objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    Set IApplicationClassDef_LoadApplicationEnquiryData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IApplicationClassDef_LoadApplicationDataForCostModelling() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IApplicationClassDef_LoadApplicationDataForCostModelling"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONLATESTDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
     '*=[MC]BMIDS763 APPLICATION DATE ADDED
    strXML = strXML & "<PRIMARYKEY>APPLICATIONDATE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    '*=SECTION END
    strXML = strXML & "<OTHERS>ACTIVEQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCEPTEDQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTACCOUNTDOWNLOADDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" 'SR 13/08/2004:BMIDS815
    strXML = strXML & "<OTHERS>FURTHERADVANCETERMIND<TYPE>dbdtBoolean</TYPE></OTHERS>" 'JD BMIDS975
    strXML = strXML & "<OTHERS>LEVELOFADVICE<TYPE>dbdtComboId</TYPE>"   'MAR18
    strXML = strXML & "  <COMBO>LevelOfAdvice</COMBO></OTHERS>"         'MAR18
    strXML = strXML & "<OTHERS>SPECIALSCHEME<TYPE>dbdtComboId</TYPE>"   'EP2_780
    strXML = strXML & "  <COMBO>SpecialSchemes</COMBO></OTHERS>"   'EP2_780
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadApplicationDataForCostModelling = _
                    objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

'BMIDS758  New function
Private Function IApplicationClassDef_LoadRemovedToECustomer() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IApplicationClassDef_LoadRemovedToECustomer"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>REMOVEDTOECUSTOMER"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>OMIGACUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"   'EP2_56 GHun
    strXML = strXML & "<OTHERS>CIFNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERFORENAMES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    'EP2_56 GHun
    strXML = strXML & "<OTHERS>CUSTOMERROLE<TYPE>dbdtComboId</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPE<TYPE>dbdtString</TYPE></OTHERS>"
    'EP2_56 End
    strXML = strXML & "</TABLENAME>"
    
    Set IApplicationClassDef_LoadRemovedToECustomer = _
                    objXmlAssist.Load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function



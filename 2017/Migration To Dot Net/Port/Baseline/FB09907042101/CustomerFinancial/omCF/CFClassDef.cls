VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
'SR     25/02/00        Removed field 'TobeRedeemed' from LoadMortgageLoan
'                       AQR - sys0235
'SR     28/04/00        Removed field 'ASSIGNTONEWLOANINDICATOR' from
'                       ICFClassDef_LoadMortgageRelatedContractsData
'                       Removed field 'RELATIONSHIP' from
'                       ICFClassDef_LoadPolicyRelationshipData()
'IW     02/05/00        Added LIFEASSUREDINDICATOR field to
'                       ICFClassDef_LoadMortgageRelatedContractsData()
'IW     02/05/00        SYS0601 - Removed refferences to PolicyRelationship table
'IW     05/05/00        SYS0504 - Added ApplicationContract Link Table
'IW     05/05/00        SYS0676 - Added MonthlyPremium Format Mask
'MC     08/05/00        Fixed SYS0681 & SYS0679. Amended 3 types to dbdtCurrency
'MC     09/06/00        SYS0866 - standardise format of balance and payment fields
'MC     08/06/00        SYS0865 - Added Agreement Type combo to LoansLiabilities
'MC     21/06/00        SYS0866 - Amend IndemnityPremium type to dbdtCurrency
'PSC    29/06/00        SYS1012 - Add LoadLoansNotRedeemedTotalsData
'BG     03/10/00        SYS0419 - Added LoadMortgageAccountArrearsData
'GD     14/03/01        SYS2047 Added GetMortgageAccount
'MH     24/04/06        MAR1448 Add IndemnityAccountGiud
'AShaw  11/12/06        EP2_19 Add OriginalLTV column data.
'-------------------------------------------------------------------------
Option Explicit

Implements ICFClassDef




'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006 Added LoadAccountRelationshipsData
'                                    LoadMortgageAccountSpecialFeatureData
'                                    LoadMortgageAccountListData
'                              Amend LoadMortgageAccountData
'                                    LoadMortgageAccountArrearsData
'GHun   24/07/2002  BMIDS00190 DCWP3 BM076 Added LoadCustVerRegularOutgoingsData,
'                                                LoadCustVerBankruptcyHistoryData,
'                                                LoadCustVerCCJHistoryData,
'                                                LoadOtherArrearsAccountData,
'                                                LoadCFArrearsHistoryData,
'                                                LoadCFLoansLiabilitiesData,
'                                                LoadLoansLiabilitiesArrearsData
'                                          Amend LoadBankruptcyHistoryData,
'                                                LoadCCJHistoryData,
'                                                LoadRegularOutgoingsData,
'                                                LoadLoansLiabilitiesData,
'                                                LoadArrearsHistoryData,
'                                                LoadLoansLiabilitiesThirdPartyData,
'                                                LoadLoansLiabilitiesDirectoryData,
'MDC    10/09/2002  BMIDS00336 CCWP1 BM062 Added LoadMortgageAccountLoanTotalsData
'MV     12/09/2002  BMIDS00387 Amended ICFClassDef_LoadMortgageLoanData()
'MDC    13/09/2002  BMIDS00336 CCWP1 BM062 Amended LoadAccountData
'MDC    16/09/2002  BMIDS00336 CCWP1 BM062 Amended  LoadBankruptcyHistoryData
'                                                   LoadCCJHistoryData
'                                                   LoadArrearsHistoryData
'                                                   LoadCFArrearsHistoryData
'                                                   LoadCFLoansLiabilitiesHistoryData
'GHun   19/09/2002  BMIDS00499 Amended LoadMortgageLoanData
'MV     23/09/2002  BMIDS00511 Modified Corrupted ClassDefs
'MV     23/09/2002  BMIDS00511 Modified Corrupted ClassDefs
'MDC    30/09/2002  BMIDS00520 Reapply missing LoadMortgageAccountArrearsData
'MDC    30/09/2002  BMIDS00510 Reapply missing LoadMortgageAccountDetailsDataEx
'GHun   03/10/2002  BMIDS00499 Reapply changes to LoadMortgageLoanData as they were lost due to corruption
'MDC    08/10/2002  BMIDS00561 Add CreditSearch to LoadCFLoansLiabilitiesData
'ASu    16/10/2002  BMIDS00621 Change database field REGULAROUTGOINGSTYPE type from a primary key to ordinary field in LoadRegularOutgoingsData classdef
'GHun   31/10/2002  BMIDS00447 Add LASTTWOYEARSINARREARS to LoadArrearsHistoryData
'GHun   13/11/2002  BMIDS00444 Add RemortgageIndicator to LoanMortgageAccountData, LoadMortgageAccountDetailsDataEx & LoadMortgageAccountListData
'                              Add LoadRemortgageAccountAddressData
'GHun   19/11/2002  BMIDS00950 Amended LoadMortgageAccountArrearsData to include BMidsAccount
'GHun   19/11/2002  BMIDS00973 Add LoadArrearsHistoryDeleteData
'DM     19/11/2002  BMIDS00935 Added NOLOCK hint to the following methods
'                   ICFClassDef_LoadDeclinedMortgageData, ICFClassDef_LoadGroupConnectionData
'MDC    03/01/2003  BM0222 Add missing fields to LoadMortgageAccountDetailsDataEx & LoadMortgageAccountListData

'GD     21/05/2003  BM0370 Change Type of DXID
'DRC    23/12/2003  BMIDS667 added Repossession flag to LoadMortgageAccountData
'HMA    11/03/2004  BMIDS712 Added CAISACCOUNTTYPE to CCJHISTORY table.
'SR     10/08/2004  BMIDS815  Modified LoadMortgageLoanData
'-------------------------------------------------------------------------
'MARS History:
'HMA    27/09/2005  MAR46     Add CurrentRateExpiryDate to LoadMortgageLoanData
'HMA    14/11/2005  MAR102    Add OriginalLoanAmount to LoadCFLoansLiabilitiesData & LoadLoansLiabilitiesData
'PE     25/01/2006  MAR1109   DC070 - Should be change to display the outstanding balance (Collateral Balance) from Mortgage Account
'JD     01/02/2006  MAR1040   Add MortgageAccountRedemptionStatus for MortgageAccountAndLoan view
'PE     20/03/2006  MAR1437   Added CardType to loansliabilities definitions.
'AW     25/07/2006  EP1023    Added CONTACTTITLEOTHER to LoadMortgageAccountDetailsDataEx()
'-------------------------------------------------------------------------
'Epsom History
'MAH    21/11/2006 E2CR35      Added OTHERPROPERTYMORTGAGE
'MAH    21/11/2006 E2CR35      Added PAIDFORBYBUSINESS
'GHun   22/11/2006  EP2_19     Added PreEmptionEndDate to LoadMortgageAccountDetailsDataEx and LoadMortgageAccountData
'PSC    16/01/2007  EP2_741    Added BANKNAME, ORIGINALNATUREOFLOAN, ORIGINALCREDITSCHEME to LoadMortgageAccountDetailsDataEx and LoadMortgageAccountData
'                              Added FLEXIBLEPRODUCTINDICATOR and ORIGINALINCOMESTATUS to LoadMortgageLoanData
'PSC    06/02/2007  EP2_1260   Change ProductStartDate to datetime
'PSC    07/02/2007  EP2_1260   Backout previous change
'-------------------------------------------------------------------------


Private Function ICFClassDef_LoadAccountRelationshipData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "ICFClassDef_LoadAccountRelationshipData"

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    Const cstrXML As String = "<TABLENAME>ACCOUNTRELATIONSHIP" & _
             "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
             "<OTHERS>CUSTOMERROLETYPE<TYPE>dbdtComboId</TYPE><COMBO>CustomerRoleType</COMBO></OTHERS>" & _
             "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(cstrXML, TypeName(Me), cstrFunctionName)
    
    Set ICFClassDef_LoadAccountRelationshipData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadAccountRelationshipDetailsData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "ICFClassDef_LoadAccountRelationshipDetailsData"

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    Const cstrXML As String = "<TABLENAME>ACCOUNTRELATIONSHIPDETAILS" & _
             "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
             "<OTHERS>CUSTOMERROLETYPE<TYPE>dbdtComboId</TYPE><COMBO>CustomerRoleType</COMBO></OTHERS>" & _
             "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
             "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(cstrXML, TypeName(Me), cstrFunctionName)
    
    Set ICFClassDef_LoadAccountRelationshipDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function


Private Function ICFClassDef_LoadCFArrearsHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCFArrearsHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CF_ARREARSHISTORY"
    strXML = strXML & "<TABLENAME>ACCOUNTRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    strXML = strXML & "<PRIMARYKEY>ARREARSHISTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATECLEARED<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTIONOFLOAN<TYPE>dbdtComboId</TYPE><COMBO>ArrearsLoanType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT<TYPE>dbdtCurrency</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMNUMBEROFMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00336 MDC 16/09/2002
    strXML = strXML & "<OTHERS>UNASSIGNED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REPOSSESSIONIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00336 MDC 16/09/2002 - End
    'BMIDS00447
    strXML = strXML & "<OTHERS>CURRENTYEARSINARREARS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTTWOYEARSINARREARS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IMPORTEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00447 End
    strXML = strXML & "<TABLENAME>OTHERARREARSACCOUNT"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCFArrearsHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCFLoansLiabilitiesData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCFLoansLiabilitiesData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CF_LOANSLIABILITIES"
    strXML = strXML & "<TABLENAME>ACCOUNTRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    strXML = strXML & "<TABLENAME>ACCOUNT"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IMPORTEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    'BMIDS00336 MDC 16/09/2002
    strXML = strXML & "<OTHERS>UNASSIGNED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00336 MDC 16/09/2002 - End
    'BMIDS00561 MDC 08/10/2002
    strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00561 MDC 08/10/2002 - End
    strXML = strXML & "</TABLENAME>"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOANREPAYMENTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT<TYPE>dbdtCurrency</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOBEREPAIDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALOUTSTANDINGBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AGREEMENTTYPE<TYPE>dbdtComboId</TYPE><COMBO>AgreementType</COMBO></OTHERS>"
    'MAR102
    strXML = strXML & "<OTHERS>ORIGINALLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    
    'MAR1437 - 20/03/2006
    strXML = strXML & "<OTHERS>CARDTYPE<TYPE>dbdtInt</TYPE></OTHERS>"

    strXML = strXML & "<OTHERS>PAIDFORBYBUSINESS<TYPE>dbdtBoolean</TYPE></OTHERS>" 'MAH 21/11/2006  E2CR35
    strXML = strXML & "</TABLENAME>"

'    strXML = "<TABLENAME>CF_LOANSLIABILITIES"
'    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
'    strXML = strXML & "<TYPE>dbdtString</TYPE>"
'    strXML = strXML & "</PRIMARYKEY>"
'    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
'    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
'    strXML = strXML & "</PRIMARYKEY>"
'    strXML = strXML & "<OTHERS>ACCOUNTGUID"
'    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>ACCOUNTNUMBER"
'    strXML = strXML & "<TYPE>dbdtString</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>IMPORTEDINDICATOR"
'    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>DIRECTORYGUID"
'    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>THIRDPARTYGUID"
'    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
'    strXML = strXML & "<TYPE>dbdtString</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
'    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>ENDDATE"
'    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>LOANREPAYMENTINDICATOR"
'    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
'    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>TOBEREPAIDINDICATOR"
'    strXML = strXML & "<OTHERS>TOTALOUTSTANDINGBALANCE"
'    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>AGREEMENTTYPE"
'    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
'    strXML = strXML & "<COMBO>AgreementType</COMBO>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCFLoansLiabilitiesData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
'    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>TOTALOUTSTANDINGBALANCE"
'    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>AGREEMENTTYPE"
'    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
'    strXML = strXML & "<COMBO>AgreementType</COMBO>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCFLoansLiabilitiesData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCustVerBankruptcyHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCustVerBankruptcyHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CUSTOMERVERSIONBANKRUPTCYHISTORY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BANKRUPTCYHISTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCustVerBankruptcyHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCustVerCCJHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCustVerCCJHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CUSTOMERVERSIONCCJHISTORY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CCJHISTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCustVerCCJHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCustVerRegularOutgoingsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCustVerRegularOutgoingsData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CUSTOMERVERSIONREGULAROUTGOINGS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>REGULAROUTGOINGSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCustVerRegularOutgoingsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadLoansLiabilitiesArrearsData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICFClassDef_LoadLoansLiabilitiesArrearsData"

    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Dim strXML          As String
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>LOANSLIABILITIESARREARS" & _
             "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
             "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<TABLENAME>ACCOUNTRELATIONSHIP" & _
             "<OTHERS>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" & _
             "</TABLENAME>" & _
             "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set ICFClassDef_LoadLoansLiabilitiesArrearsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageAccDetsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccDetsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>MORTGAGEACCOUNT"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERADDRESSSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set ICFClassDef_LoadMortgageAccDetsData = xmlClassDefDoc
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadAccountData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadAccountData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = "<TABLENAME>ACCOUNT"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>IMPORTEDINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 13/09/2002
    strXML = strXML & "<OTHERS>UNASSIGNED"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CREDITSEARCH"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CAISRECORDREF"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 13/09/2002 - End
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadAccountData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadMortgageAccountListData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICFClassDef_LoadMortgageAccountListData"

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    Dim strXML As String
    
    'BM0222 MDC 03/01/2003 - Add UNASSIGNED
    'BMIDS00444 Added REMORTGAGEINDICATOR
    'MAR1109 - Peter Edney - 25/01/2006
    'Added TOTALCOLLATERALBALANCE node.
    strXML = "<TABLENAME>MORTGAGEACCOUNTSUMMARY" & _
             "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
             "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
             "<OTHERS>SECONDCHARGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
             "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>REMORTGAGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
             "<OTHERS>UNASSIGNED<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
             "<OTHERS>TOTALCOLLATERALBALANCE<TYPE>dbdtString</TYPE></OTHERS>" & _
             "</TABLENAME>"
 
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountListData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

'BMIDS00336 MDC 10/09/2002
Private Function ICFClassDef_LoadMortgageAccountLoanTotalsData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' procedure:    ICFClassDef_LoadMortgageAccountLoanTotalsData
' author:       Mark Coleman
' date:         12/09/2002
' description:
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountLoanTotalsData"
    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = "<TABLENAME>MORTGAGEACCOUNTLOANTOTALS"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>OUTSTANDINGBALANCE"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>STARTDATE"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountLoanTotalsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function
'BMIDS00336 MDC 10/09/2002 - End

Private Function ICFClassDef_LoadMortgageAccountSpecialFeatureData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICFClassDef_LoadMortgageAccountSpecialFeatureData"

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    Const cstrXML As String = "<TABLENAME>MORTGAGEACCOUNTSPECIALFEATURE" & _
             "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
             "<PRIMARYKEY>MORTGAGEACCOUNTSPECIALFEATURESEQNO<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
             "<OTHERS>MORTGAGEACCOUNTSPECIALFEATUREIND<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>MORTGAGEACCOUNTSPECIALFEATUREDESC<TYPE>dbdtString</TYPE></OTHERS>" & _
             "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(cstrXML, TypeName(Me), cstrFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountSpecialFeatureData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageLenderDirectoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageLenderDirectoryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = "<TABLENAME>MORTGAGELENDERDIRECTORY"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ORGANISATIONID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MAINMORTGAGELENDERIND"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageLenderDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadApplicationContractData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadApplicationContractData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>APPLICATIONCONTRACT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ASSIGNEDTOLOANINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadApplicationContractData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadArrearsHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadArrearsHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>ARREARSHISTORY"
    'BMIDS00190
    strXML = strXML & "<PRIMARYKEY>ARREARSHISTORYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    'strXML = strXML & "<TYPE>dbdtString</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<OTHERS>ACCOUNTGUID"
    'strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    'strXML = strXML & "</OTHERS>"
    'strXML = strXML & "<PRIMARYKEY>ARREARSSEQUENCENUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'BMIDS00190
    strXML = strXML & "<OTHERS>DATECLEARED"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTIONOFLOAN"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ArrearsLoanType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMBALANCE"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>OTHERDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMNUMBEROFMONTHS"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 16/09/2002
    strXML = strXML & "<OTHERS>REPOSSESSIONIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNASSIGNED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CURRENTYEARSINARREARS<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00336 MDC 16/09/2002 - End
    'BMIDS00447
    strXML = strXML & "<OTHERS>LASTTWOYEARSINARREARS<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00447 End
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadArrearsHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadBankCreditCardData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadBankCreditCardData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>BANKCREDITCARD"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>AVERAGEMONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CARDPROVIDER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CARDTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CreditCardType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>TOTALOUTSTANDINGBALANCE"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadBankCreditCardData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadBankruptcyHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadBankruptcyHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>BANKRUPTCYHISTORY"
    'BMIDS00190
    strXML = strXML & "<PRIMARYKEY>BANKRUPTCYHISTORYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    'strXML = strXML & "<TYPE>dbdtString</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'BMIDS00190 End
    strXML = strXML & "<OTHERS>AMOUNTOFDEBT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEDECLARED"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFDISCHARGE"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>OTHERDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 16/09/2002
    strXML = strXML & "<OTHERS>UNASSIGNED"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>IVA"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CREDITSEARCH"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>COURTPLAINTIFF"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 16/09/2002 - End
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadBankruptcyHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCCJHistoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCCJHistoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CCJHISTORY"
    'BMIDS00190
    strXML = strXML & "<PRIMARYKEY>CCJHISTORYGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    'strXML = strXML & "<TYPE>dbdtString</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'BMIDS00190 End
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATECLEARED"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFJUDGEMENT"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>OTHERDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>PLAINTIFF"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>VALUEOFJUDGEMENT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 16/09/2002
    strXML = strXML & "<OTHERS>CCJTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CCJType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DEFAULTRECORD"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CREDITSEARCH"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>UNASSIGNED"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00336 MDC 16/09/2002 - End
    'BMIDS712 Added CAISAccountType
    strXML = strXML & "<OTHERS>CAISACCOUNTTYPE"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCCJHistoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCurrentPropertyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCurrentPropertyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>CURRENTPROPERTY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTION"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyDescription</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ESTIMATEDCURRENTVALUE"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>INTENDEDACTION"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CurrentPropertyIntendedAction</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUATIONDATE"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERID"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>TENURE"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>YEARBUILT"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>REINSTATEMENTAMOUNT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERNAME"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCurrentPropertyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadCustomerAddressData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadCustomerAddressData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>CUSTOMERADDRESS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERADDRESSSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSTYPE"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDIN"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDOUT"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>NATUREOFOCCUPANCY"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>NatureOfOccupancy</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadCustomerAddressData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function


Private Function ICFClassDef_LoadDeclinedMortgageData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadDeclinedMortgageData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>DECLINEDMORTGAGE"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'DM     19/11/2002  BMIDS00935
    strXML = strXML & "<SQLNOLOCK>TRUE</SQLNOLOCK>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DATEDECLINED"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>DECLINEDDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadDeclinedMortgageData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadFinancialOrganisationData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadFinancialOrganisationData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>FINANCIALORGANISATION"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ORGANISATIONNAME"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>THISORGANISATIONINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ORGANISATIONTYPE"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadFinancialOrganisationData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function


Private Function ICFClassDef_LoadGroupConnectionData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadGroupConnectionData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>GROUPCONNECTION"

    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM     19/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>GROUPCONNECTIONSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>ACCOUNTTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>GroupConnectionAccountType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadGroupConnectionData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadHomeInsuranceData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadHomeInsuranceData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>HOMEINSURANCE"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>HOMEINSURANCETYPE"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadHomeInsuranceData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadIndemnityInsuranceData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadIndemnityInsuranceData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>INDEMNITYINSURANCE"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INDEMNITYCOMPANYNAME"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>INDEMNITYAMOUNT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>INDEMNITYMORTGAGEAMOUNT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>INDEMNITYDATEPAID"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadIndemnityInsuranceData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadLifePolicyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLifePolicyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>LIFEPOLICY"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>MATURITYVALUE"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMDEATHBENEFIT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYPREMIUM"
    strXML = strXML & "<TYPE>dbdtDouble</TYPE>"
    strXML = strXML & "<FORMATMASK>0.00</FORMATMASK>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>POLICYTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PolicyType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>YEAROFMATURITY"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLifePolicyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadLifeProductData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLifeProductData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = "<TABLENAME>LIFEPRODUCT"
    strXML = strXML & ICFClassDef_LoadLifePolicyData.xml
    strXML = strXML & ICFClassDef_LoadMortgageRelatedContractsData.xml
    strXML = strXML & ICFClassDef_LoadApplicationContractData.xml
    strXML = strXML & ICFClassDef_LoadAccountData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLifeProductData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadLoansLiabilitiesData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLoansLiabilitiesData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>LOANSLIABILITIES"
    'BMIDS00190
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    'strXML = strXML & "<TYPE>dbdtString</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER"
    'strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    'strXML = strXML & "</PRIMARYKEY>"
    'BMIDS00190 End
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ENDDATE"
    strXML = strXML & "<TYPE>dbdtDate</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00190
    'strXML = strXML & "<OTHERS>LOANACCOUNTNUMBER"
    'strXML = strXML & "<TYPE>dbdtString</TYPE>"
    'strXML = strXML & "</OTHERS>"
    'BMIDS00190 End
    strXML = strXML & "<OTHERS>LOANREPAYMENTINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT"
    strXML = strXML & "<TYPE>dbdtCurrency</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>TOBEREPAIDINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>TOTALOUTSTANDINGBALANCE"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    'BMIDS00190
    'strXML = strXML & "<OTHERS>THIRDPARTYGUID"
    'strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    'strXML = strXML & "</OTHERS>"
    'strXML = strXML & "<OTHERS>DIRECTORYGUID"
    'strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    'strXML = strXML & "</OTHERS>"
    'BMIDS00190 End
    strXML = strXML & "<OTHERS>AGREEMENTTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AgreementType</COMBO>"
    strXML = strXML & "</OTHERS>"
    'MAR102
    strXML = strXML & "<OTHERS>ORIGINALLOANAMOUNT"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    
    'MAR1437 - 20/03/2006
    strXML = strXML & "<OTHERS>CARDTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAIDFORBYBUSINESS<TYPE>dbdtBoolean</TYPE></OTHERS>" 'MAH 21/11/2006 E2CR35
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLoansLiabilitiesData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadLoansLiabilitiesDirectoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLoansLiabilitiesDirectoryData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Dim objTPClassDef   As IThirdPartyClassDef

    Set objXmlAssist = New XMLAssist
    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>LOANSLIABILITIESDIRECTORY"
    'BMIDS00190
    'strXML = strXML & ICFClassDef_LoadLoansLiabilitiesData.xml
    'BMIDS00190 End
    strXML = strXML & objTPClassDef.LoadDirectorySummaryData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLoansLiabilitiesDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function

Private Function ICFClassDef_LoadLoansLiabilitiesThirdPartyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLoansLiabilitiesThirdPartyData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Dim objTPClassDef   As IThirdPartyClassDef

    Set objXmlAssist = New XMLAssist
    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>LOANSLIABILITIESTHIRDPARTY"
    'BMIDS00190
    'strXML = strXML & ICFClassDef_LoadLoansLiabilitiesData.xml
    'BMIDS00190 End
    strXML = strXML & objTPClassDef.LoadThirdPartyDetailsData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLoansLiabilitiesThirdPartyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function

Private Function ICFClassDef_LoadLoansNotRedeemedTotalsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadLoansNotRedeemedTotalsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>MORTGAGELOAN"
    strXML = strXML & "<OTHERS>LOANSNOTREDEEMED"
    strXML = strXML & "<TYPE>dbdtLong</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTSONLOANSNOTREDEEMED"
    strXML = strXML & "<TYPE>dbdtDouble</TYPE>"
    strXML = strXML & "<FORMATMASK>0.00</FORMATMASK>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadLoansNotRedeemedTotalsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageAccountAndLoanData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountAndLoanData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef

    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>MORTGAGEACCOUNTANDLOAN"
    strXML = strXML & ICFClassDef_LoadAccountRelationshipData.xml
    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
    strXML = strXML & ICFClassDef_LoadMortgageLoanData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountAndLoanData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function


Private Function ICFClassDef_LoadMortgageAccountData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>MORTGAGEACCOUNT"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<OTHERS>SECONDCHARGEINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    
    ' PSC 22/07/02 BMIDS00006 - Start
    strXML = strXML & "<OTHERS>MONTHLYRENTALINCOME<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPE<TYPE>dbdtComboId</TYPE><COMBO>PropertyType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtComboId</TYPE><COMBO>PropertyDescription</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ESTIMATEDCURRENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTENDEDACTION<TYPE>dbdtComboId</TYPE><COMBO>CurrentPropertyIntendedAction</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUATIONAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TENURE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEARBUILT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REINSTATEMENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HOMEINSURANCETYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BMIDSACCOUNT<TYPE>dbdtBoolean</TYPE></OTHERS>"

'    strXML = strXML & "<OTHERS>CUSTOMERNUMBER"
'    strXML = strXML & "<TYPE>dbdtString</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER"
'    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
'    strXML = strXML & "</OTHERS>"
'    strXML = strXML & "<OTHERS>CUSTOMERADDRESSSEQUENCENUMBER"
'    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
'    strXML = strXML & "</OTHERS>"
    ' PSC 22/07/02 BMIDS00006 - End
    
    '~ MF 11/08/2005 MAR20
    strXML = strXML & "<OTHERS>REDEMPTIONSTATUS<TYPE>dbdtComboId</TYPE><COMBO>RedemptionStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMONTHLYCOST<TYPE>dbdtCurrency</TYPE></OTHERS>"
    
    'JD MAR1040 for MortgageAccountAndLoan view
    strXML = strXML & "<OTHERS>MORTGAGEACCOUNTREDEMPTIONSTATUS<TYPE>dbdtComboId</TYPE><COMBO>RedemptionStatus</COMBO></OTHERS>"

    
    'BMIDS00444
    strXML = strXML & "<OTHERS>REMORTGAGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00444 End
    'BMIDS667
    strXML = strXML & "<OTHERS>REPOSSESSIONFLAG<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS667 End
    '*=[MC] BMIDS756
    strXML = strXML & "<OTHERS>COLLATERALID<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALCOLLATERALBALANCE<TYPE>dbdtCurrency</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTDUEDATE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RMCFLAG<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DAILYMONTHLYANNUALINTEREST<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DSSFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKACCOUNTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUSINESSCHANNEL<TYPE>dbdtString</TYPE></OTHERS>"
    'END BMIDS756
    strXML = strXML & "<OTHERS>OTHERPROPERTYMORTGAGE<TYPE>dbdtBoolean</TYPE></OTHERS>" ' MAH E2CR35 21/11/2006
    strXML = strXML & "<OTHERS>PREEMPTIONENDDATE<TYPE>dbdtDate</TYPE></OTHERS>" 'EP2_19 GHun
    ' PSC 16/01/2007 EP2_741 - Start
    strXML = strXML & "<OTHERS>BANKNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALNATUREOFLOAN<TYPE>dbdtComboId</TYPE><COMBO>NatureOfLoan</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALCREDITSCHEME<TYPE>dbdtComboId</TYPE><COMBO>NatureOfLoan</COMBO></OTHERS>"
    ' PSC 16/01/2007 EP2_741 - End
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadMortgageAccountDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = "<TABLENAME>MORTGAGEACCOUNTDETAILS"
    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
    strXML = strXML & ICFClassDef_LoadAccountData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function


Private Function ICFClassDef_LoadMortgageAccountDirectoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountDirectoryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef

    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>MORTGAGEACCOUNTDIRECTORY"
    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
    strXML = strXML & ICFClassDef_LoadAccountData.xml
    strXML = strXML & ICFClassDef_LoadIndemnityInsuranceData.xml
    strXML = strXML & objTPClassDef.LoadDirectorySummaryData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function

Private Function ICFClassDef_LoadOtherArrearsAccountData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadOtherArrearsAccountData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>OTHERARREARSACCOUNT"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadOtherArrearsAccountData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function ICFClassDef_LoadRegularOutgoingsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadRegularOutgoingsData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>REGULAROUTGOINGS"
    'BMIDS00190
    strXML = strXML & "<PRIMARYKEY>REGULAROUTGOINGSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    'strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'BMIDS00190 End
    'ASu BMIDS00621 - Start
    'strXML = strXML & "<PRIMARYKEY>REGULAROUTGOINGSTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>REGULAROUTGOINGSTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    'ASu - End
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTFREQUENCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>RegularOutgoingsPaymentFreq</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
   
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadRegularOutgoingsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageAccountArrearsData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountArrearsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef

    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>MORTGAGEACCOUNTARREARS"
    'BMIDS00520 MDC 30/09/2002
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    
    'BMIDS00950
    strXML = strXML & "<OTHERS>BMIDSACCOUNT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00950 End
    
    strXML = strXML & "<TABLENAME>ACCOUNTRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
'    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
'    strXML = strXML & ICFClassDef_LoadAccountData.xml
'    strXML = strXML & ICFClassDef_LoadIndemnityInsuranceData.xml
'    strXML = strXML & objTPClassDef.LoadThirdPartyDetailsData.xml
    'BMIDS00520 MDC 30/09/2002 - End
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountArrearsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function


Private Function ICFClassDef_LoadMortgageAccountThirdPartyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountThirdPartyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef
    Dim objContactClassDef As IContactClassDef

    Set objTPClassDef = New ThirdPartyClassDef
    Set objContactClassDef = New ContactClassDef

    strXML = "<TABLENAME>MORTGAGEACCOUNTTHIRDPARTY"
    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
    strXML = strXML & ICFClassDef_LoadAccountData.xml
    strXML = strXML & ICFClassDef_LoadIndemnityInsuranceData.xml
    strXML = strXML & objTPClassDef.LoadThirdPartyDetailsData.xml
    strXML = strXML & objContactClassDef.LoadContactTelephoneData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountThirdPartyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
    Set objContactClassDef = Nothing
End Function


Private Function ICFClassDef_LoadMortgageAccountDetailsDataEx() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageAccountDetailsDataEx"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    'BMIDS00510 MDC 30/09/2002 - Reinstate overwritten change for BMIDS00006
'    strXML = "<TABLENAME>MORTGAGEACCOUNTDETAILS"
'    strXML = strXML & ICFClassDef_LoadMortgageAccountData.xml
'    strXML = strXML & ICFClassDef_LoadAccountData.xml
'    strXML = strXML & ICFClassDef_LoadIndemnityInsuranceData.xml
'    strXML = strXML & "</TABLENAME>"

    strXML = "<TABLENAME>MORTGAGEACCOUNTDETAILS" & _
             "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
             "<OTHERS>SECONDCHARGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
             "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
             "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>MONTHLYRENTALINCOME<TYPE>dbdtCurrency</TYPE></OTHERS>" & _
             "<OTHERS>TYPE<TYPE>dbdtComboId</TYPE><COMBO>PropertyType</COMBO></OTHERS>" & _
             "<OTHERS>DESCRIPTION<TYPE>dbdtComboId</TYPE><COMBO>PropertyDescription</COMBO></OTHERS>" & _
             "<OTHERS>ESTIMATEDCURRENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>" & _
             "<OTHERS>INTENDEDACTION<TYPE>dbdtComboId</TYPE><COMBO>CurrentPropertyIntendedAction</COMBO></OTHERS>" & _
             "<OTHERS>LASTVALUATIONAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" & _
             "<OTHERS>LASTVALUATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>" & _
             "<OTHERS>LASTVALUERID<TYPE>dbdtString</TYPE></OTHERS>" & _
             "<OTHERS>TENURE<TYPE>dbdtInt</TYPE></OTHERS>"
    
    '~ MF 10/08/2005 MAR20
    strXML = strXML & "<OTHERS>REDEMPTIONSTATUS<TYPE>dbdtComboId</TYPE><COMBO>RedemptionStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMONTHLYCOST<TYPE>dbdtCurrency</TYPE></OTHERS>"
    
    'BMIDS00444 Added REMORTGAGEINDICATOR
    strXML = strXML & "<OTHERS>YEARBUILT<TYPE>dbdtInt</TYPE></OTHERS>" & _
                      "<OTHERS>REINSTATEMENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" & _
                      "<OTHERS>LASTVALUERNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtLong</TYPE></OTHERS>" & _
                      "<OTHERS>HOMEINSURANCETYPE<TYPE>dbdtInt</TYPE></OTHERS>" & _
                      "<OTHERS>MAADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>BMIDSACCOUNT<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "<OTHERS>REMORTGAGEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS667
    strXML = strXML & "<OTHERS>REPOSSESSIONFLAG<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS667 End
    
    'BMIDS756
    strXML = strXML & "<OTHERS>COLLATERALID<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALCOLLATERALBALANCE<TYPE>dbdtCurrency</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTDUEDATE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RMCFLAG<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DAILYMONTHLYANNUALINTEREST<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DSSFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKACCOUNTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUSINESSCHANNEL<TYPE>dbdtString</TYPE></OTHERS>"
    
    'MAH 21/11/2006 E2CR35
    strXML = strXML & "<OTHERS>OTHERPROPERTYMORTGAGE<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    'BMIDS756 END
                           
    strXML = strXML & "<OTHERS>PREEMPTIONENDDATE<TYPE>dbdtDate</TYPE></OTHERS>" 'EP2_19 GHun
                           
    ' PSC 16/01/2007 EP2_741 - Start
    strXML = strXML & "<OTHERS>BANKNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALNATUREOFLOAN<TYPE>dbdtComboId</TYPE><COMBO>NatureOfLoan</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALCREDITSCHEME<TYPE>dbdtComboId</TYPE><COMBO>NatureOfLoan</COMBO></OTHERS>"
    ' PSC 16/01/2007 EP2_741 - End
                           
    'BM0222 MDC 03/01/2003 - Add CREDITSEARCH & UNASSIGNED
    strXML = strXML & "<TABLENAME>ACCOUNT" & _
                      "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>IMPORTEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "<OTHERS>UNASSIGNED<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "</TABLENAME>"
                       
    strXML = strXML & "<TABLENAME>NAMEANDADDRESSDIRECTORY" & _
                      "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>NAMEANDADDRESSTYPE<TYPE>dbdtComboId</TYPE><COMBO>ThirdPartyType</COMBO></OTHERS>" & _
                      "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>HEADOFFICEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "<OTHERS>NAMEANDADDRESSACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" & _
                      "<OTHERS>NAMEANDADDRESSACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" & _
                      "<OTHERS>DXID<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>NOTES<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>ORGANISATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>OrganisationType</COMBO></OTHERS>" & _
                      "<OTHERS>BRANCHNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>NAMEANDADDRESSBANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "</TABLENAME>"
'GD BM0370 - Changed type for DXID (Above)
    strXML = strXML & "<TABLENAME>THIRDPARTY" & _
                      "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>THIRDPARTYTYPE<TYPE>dbdtComboId</TYPE><COMBO>ThirdPartyType</COMBO></OTHERS>" & _
                      "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>ORGANISATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>OrganisationType</COMBO></OTHERS>" & _
                      "<OTHERS>BRANCHNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>THIRDPARTYBANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "</TABLENAME>"
                      
    strXML = strXML & "<TABLENAME>ADDRESS" & _
                      "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>COUNTRY<TYPE>dbdtComboId</TYPE><COMBO>Country</COMBO></OTHERS>" & _
                      "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>DELIVERYPOINTSUFFIX<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>MAILSORTCODE<TYPE>dbdtLong</TYPE></OTHERS>" & _
                      "<OTHERS>PAFINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                      "</TABLENAME>"
                      
    strXML = strXML & "<TABLENAME>CONTACTDETAILS" & _
                      "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTTYPE<TYPE>dbdtInt</TYPE></OTHERS>" & _
                      "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>CONTACTTITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "</TABLENAME>"
'//MAR1448 M Heys 24/04/2006 start
'    strXML = strXML & "<TABLENAME>INDEMNITYINSURANCE" & _
'                      "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
'                      "<OTHERS>INDEMNITYCOMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
'                      "<OTHERS>INDEMNITYAMOUNT<TYPE>dbdtCurrency</TYPE></OTHERS>" & _
'                      "<OTHERS>INDEMNITYMORTGAGEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" & _
'                      "<OTHERS>INDEMNITYDATEPAID<TYPE>dbdtDate</TYPE></OTHERS>" & _
'                      "</TABLENAME>" & _
'                      "</TABLENAME>"
    strXML = strXML & "<TABLENAME>INDEMNITYINSURANCE" & _
                      "<OTHERS>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "<OTHERS>INDEMNITYCOMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                      "<OTHERS>INDEMNITYAMOUNT<TYPE>dbdtCurrency</TYPE></OTHERS>" & _
                      "<OTHERS>INDEMNITYMORTGAGEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" & _
                      "<OTHERS>INDEMNITYDATEPAID<TYPE>dbdtDate</TYPE></OTHERS>" & _
                      "<OTHERS>INDEMNITYACCOUNTGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                      "</TABLENAME>" & _
                      "</TABLENAME>"
'//MAR1448 M Heys 24/04/2006 end
                  
    'BMIDS00510 MDC 30/09/2002 - End
                  
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageAccountDetailsDataEx = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageLoanData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "ICFClassDef_LoadMortgageLoanData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>MORTGAGELOAN"
    strXML = strXML & "<PRIMARYKEY>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INTERESTRATE<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"
    strXML = strXML & "<OTHERS>LOANACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIRASAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIRASELIGIBILITY<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYREPAYMENT<TYPE>dbdtCurrency</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALLOANAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALPARTANDPARTINTONLYAMT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALTERMMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALTERMYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OUTSTANDINGBALANCE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtComboId</TYPE><COMBO>PurposeOfLoan</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>REDEMPTIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REDEMPTIONSTATUS<TYPE>dbdtComboId</TYPE><COMBO>RedemptionStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>REPAYMENTTYPE<TYPE>dbdtComboId</TYPE><COMBO>RepaymentType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>STARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOBEREPAIDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISBURSEDAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    'BMIDS00499 CCIINDICATOR and CCAINDICATOR are alphanumeric fields
    'strXML = strXML & "<OTHERS>CCIINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>CCAINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CCIINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CCAINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00499 End
    strXML = strXML & "<OTHERS>ICBSCALCULATEDOUTSTANDINGTERM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PENALTYPLANDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOANCLASSTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PENALTYPLANCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PENALTYPLANENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OVERPAYMENTS<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOANTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXISTINGRETENTIONS<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AVAILABLEFORDISBURSEMENT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REDEMPTIONPENALTY<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOANENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    
    '*=[MC] BMIDS756 Mortgage Regulations Change Start
    'strXML = strXML & "<OTHERS>INTERESTONLYAMOUNT<TYPE>dbdtCurrency</TYPE></OTHERS>"
    'PSC 06/02/2007 EP2_1260
    strXML = strXML & "<OTHERS>PRODUCTSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INDEXCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VARIANCE<TYPE>dbdtInt</TYPE></OTHERS>"
    '*=[MC] BMIDS756 Mortgage Regulations Change End
    
    'SR 10/08/2004 : BMIDS815
    strXML = strXML & "<OTHERS>PRODUCTSTEP<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REMAININGSTEPDURATION<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REMAININGINTERESTONLYAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REMAININGCAPITALINTERESTAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    'SR 10/08/2004 : BMIDS815 - End
    
    'MAR46
    strXML = strXML & "<OTHERS>CURRENTRATEEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    
    'EP2_19
    strXML = strXML & "<OTHERS>ORIGINALLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
    
    ' PSC 16/01/2007 EP2_741 - Start
    strXML = strXML & "<OTHERS>FLEXIBLEPRODUCTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORIGINALINCOMESTATUS<TYPE>dbdtComboId</TYPE><COMBO>ApplicationIncomeStatus</COMBO></OTHERS>"
    ' PSC 16/01/2007 EP2_741 - End

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageLoanData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function ICFClassDef_LoadMortgageRelatedContractsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadMortgageRelatedContractsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>MORTGAGERELATEDCONTRACTS"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
'    strXML = strXML & "<OTHERS>ASSIGNTONEWLOANINDICATOR"
'    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
'    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>NEWPOLICYINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>LIFEASSUREDINDICATOR"
    strXML = strXML & "<TYPE>dbdtBoolean</TYPE>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadMortgageRelatedContractsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function


Private Function ICFClassDef_LoadPolicyRelationshipData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadPolicyRelationshipData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>POLICYRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID"
    strXML = strXML & "<TYPE>dbdtGuid</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER"
    strXML = strXML & "<TYPE>dbdtString</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER"
    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "</PRIMARYKEY>"
'    strXML = strXML & "<OTHERS>RELATIONSHIP"
'    strXML = strXML & "<TYPE>dbdtInt</TYPE>"
'    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadPolicyRelationshipData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

'BMIDS00444
Private Function ICFClassDef_LoadRemortgageAccountAddressData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadRemortgageAccountAddressData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
       
    strXML = "<TABLENAME>FINDREMORTGAGEACCOUNTADDRESS"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<TABLENAME>ACCOUNTRELATIONSHIP"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    strXML = strXML & "<TABLENAME>ADDRESS"
    strXML = strXML & "<PRIMARYKEY>ADDRESSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Country</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DELIVERYPOINTSUFFIX<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAILSORTCODE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAFINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME></TABLENAME>"
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadRemortgageAccountAddressData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function
'BMIDS00444 End

'BMIDS00973 ClassDef used for deleting ArrearsHistory records by AccountGUID
Private Function ICFClassDef_LoadArrearsHistoryDeleteData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICFClassDef_LoadArrearsHistoryDeleteData"

    Dim strXML          As String
    Dim xmlClassDefDoc  As FreeThreadedDOMDocument40
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
       
    strXML = "<TABLENAME>ARREARSHISTORY"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set ICFClassDef_LoadArrearsHistoryDeleteData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function
'BMIDS00973 End




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AgreementInPrincipleBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AgreementInPrincipleBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (Omiga4TxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'RF     30/09/99    Applied changes raised by code review of 30/09/99, including:
'                   removed AnonInterfaceFunction
'                   removed calls to Validate
'                   improved error handling
'RF     04/10/99    Added profiling
'SR     15/03/01    SYS1840 - New method 'FindThirdPartyForCustomer'
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'PSC    11/12/01    SYS3361 Bring back correct data for accountant
'
'------------------------------------------------------------------------------------------

Option Explicit
 
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
 
Private Const cstrBANKCREDITCARDLIST_NAME = "BANKCREDITCARDLIST"
Private Const cstrBANKRUPTCYHISTORYLIST_NAME = "BANKRUPTCYHISTORYLIST"
Private Const cstrDECLINEDMORTGAGELIST_NAME = "DECLINEDMORTGAGELIST"
Private Const cstrCCJHISTORYLIST_NAME = "CCJHISTORYLIST"
Private Const cstrPOLICYRELATIONSHIPLIST_NAME = "POLICYRELATIONSHIPLIST"
Private Const cstrARREARSHISTORYLIST_NAME = "ARREARSHISTORYLIST"
Private Const cstrMORTGAGEACCOUNTLIST_NAME = "MORTGAGEACCOUNTLIST"
Private Const cstrMORTGAGELOANLIST_NAME = "MORTGAGELOANLIST"
Private Const cstrLOANSLIABILITIESLIST_NAME = "LOANSLIABILITIESLIST"
Private Const cstrFINANCIALSUMMARY_NAME = "FINANCIALSUMMARY"
Private Const cstrREGULAROUTGOINGS_NAME = "REGULAROUTGOINGS"
Private Const cstrQQREGULAROUTGOINGS_NAME = "QUICKQUOTEOUTGOINGS"

Private Const cstrSHOWNOTREDEEMED_NAME = "SHOWNOTREDEEMED" 'BM0492 LDM 15/05/2003

'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    14/08/2002  BMIDS00336 CCWP1 BM062 Credit Check & Bureau Download
'MDC    26/09/2002  BMIDS00525 Amend ImportBureauData
'MDC    08/10/2002  BMIDS00561 Fixes for ImportBureauData
'MDC    08/11/2002  BMIDS00801 Add Amountrequested to BankruptcyHistory in ImportBureauData
'MDC    11/11/2002  BMIDS00919 Correct MortgageAccount creation in ImportBureauData
'MDC    14/11/2002  BMIDS00946 Match LoanLiability Types against Bureau Account Types in ImportBureauData
'MDC    14/11/2002  BMIDS00945 Calculate Monthly Repayment for Credit Cards in ImportBureauData
'MDC    14/11/2002  BMIDS00949 Fix creation of Aliases in ImportBureauData
'MDC    16/11/2002  BMIDS00917 Retrieve all Employments with an Accountant in GetAccountantDetailsForCustomer
'MDC    18/11/2002  BMIDS00972 Duplication of loan liability records in ImportBureauData
'MDC    19/11/2002  BMIDS00985 Voters Roll data in ImportBureauData
'GD     19/11/2002  BMIDS00997  Use validationtypes and not valueids in GetEmploymentdetailsForCustomer
'MDC    20/11/2002  BMIDS01027 Handle 'other' CAIS Account Types in ImportBureauData
'MDC    22/11/2002  BMIDS01055 Handle apostrophes in customer names
'MDC    02/12/2002  BMIDS01124
'MDC    05/12/2002  BM0158
'MDC    05/12/2002  BM0162
'MDC    06/12/2002  BM0168
'MDC    06/12/2002  BM0172
'MDC    16/01/2003  BM0251
'GD     24/02/2003  BM0247 Fixes for ImportBureauData()
'GD     26/02/2003  BM0373 Fixes for ImportBureauData() - ENHANCE THE DE-DUPLICATE EXPERIAN RECORD PROCESS
'GHun   10/03/2003  BM0376 CC026 Added RunIncomeCalcs
'GD     21/03/2003  BM0482 Fixes to ImportBureauData() - Alter data that is passed to CreateCCJHistory
'GD     24/03/2003  BM0479 Fixes to ImportBureauData() - END DATE IS DEFAULTD ON ALL LOANS WHEN IT SHOULD ONLY BE CREDIT CARD
'GD     24/03/2003  BM0486 Fixes to ImportBureauData() - ARREARS HISTORY - REPOSSESSIONS
'LDM    15/05/2003  BM0492 for dc70 bring back value for outstanding balance on total loans not redeemed
'HMA    26/06/2003  CC037  Include all of the Public Information codes returned by Experian when importing the bureau
'                          to create or update CCJ's Bankruptcy History details.
'HMA    26/06/2003  BM0529 Create CCJ record for status 9 Bad Debt.
'HMA    26/06/2003  BMIDS595 and BMIDS596 Correct handling of the FBCAISAccountStatus string in ImportBureauData()
'HMA    30/06/2003  BMIDS590 (CC037) Ensure IVA flag is set for all necessary Public Information codes.
'HMA    30/06/2003  BMIDS593 (CC041) When selecting CAIS accounts to process also select CAIS accounts
'                            with arrears which have been settled in the last 'n' months.
'HMA    30/06/2003  BMIDS597 Default values on Bankruptcy and CCJ history.
'HMA    03/07/2003  BMIDS590 Use CLng instead of CInt to convert strCAISCurrentDefaultBal.
'INR    04/07/2003  BMIDS593 (CC041) Type mismatch on CDate with empty string & check xmlNewCAISAccounts is not nothing
'INR    07/07/2003  BMIDS597 Need to create or update financial summary even if no records exist, for
'                   the No indicator to display correctly on screens DC090 130 140 150
'INR    17/07/2003  BMIDS593 (CC041)FBCAISWorstStatus should be all uppercase.
'GD     25/07/2003  BMIDS610 Ensure COURTPLAINTIFF is populated.
'GD     28/07/2003  BMIDS617 Change redemption status to be 'Other Lender' from 'Not to be Redeemed'
'GD     31/07/2003  BMIDS610 Further fixes
'GD     31/07/2003  BMIDS610 Further fixes to remove duplicate BankruptcyHistory records
'GD     31/07/2003  BMIDS610 Further fixes to remove duplicate BankruptcyHistory records
'INR    05/09/2003  BMIDS630 Fix date declared/discharged by InfoDataType and ensure that
'                   only matched PublicInfo records are used.
'INR    05/09/2003  BMIDS633 Redemption status not showing satisfied mortgages/loans/liabilities.
'INR    10/09/2003  BMIDS630 Fixes to allow updates to bankruptcy data.
'INR    12/09/2003  BMIDS630 Changed to allow a bankruptcy record per client where we have a type
'                   of 24,25,26,27,28,30,32,33,34,35,38,41,43,44,47,48,49,50,51,52,53 or 54.
'GHun   26/11/2003  BMIDS658 Amend ImportBureauData to reduce locking
'DRC -  09/12/2003  BMIDS669 For a settled CAIS record with arrears the monthly payment is set to 0
'HMA    28/01/2004  BMIDS661 Allow for date of bankruptcy and date of discharge being the same.
'DRC    18/02/2004  BMIDS693 Add Credit search flag & matching of names by Uppper Case alpha characters only
'HMA    04/03/2004  BMIDS712 Add CAISAccountType to CCJ History record.
'SR     26/05/2004  BMIDS772 New methods - 'GetFinancialSummary' and 'GetCrditHistorySummaryView'
'                            Modified methods 'FindArrearsHistorySummary', 'FindBankCardSummary', 'FindBankruptcyHistorySummary'
'                            'FindCCJHistorySummary', 'FindLiabilitySummary', 'FindMortgageAccountSummary',
'                            'FindDeclinedMortgageSummary'
'JD     05/10/2004  BMIDS897 Added private method GetDirectoryGUIDForBureau. Code no longer creates a new
'                            ThirdParty every time, now gets an existing directory based on the company name and NameAndAddressType.
'-------------------------------------------------------------------------------------------------------------------------------------


'BMIDS00336 MDC 14/08/2002
Public Function ImportBureauData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' procedure:    ImportBureauData
' author:       Mark Coleman
' date:         14/08/2002
' description:
' pass:         vstrXMLRequest  XML string request specifying the data and operation
' return:       XML response string
'------------------------------------------------------------------------------------------
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlTempRequestElem As IXMLDOMElement
Dim xmlTempElem As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlBureauData As IXMLDOMNode
Dim xmlAliasListNode As IXMLDOMNode
Dim xmlAliasNode As IXMLDOMNode
Dim xmlTempParentElem As IXMLDOMNode
Dim xmlCustomerElem As IXMLDOMElement
Dim xmlPersonNode As IXMLDOMNode
Dim xmlMortgageAccountListNode As IXMLDOMNode
Dim xmlLiabilityListNode As IXMLDOMNode
Dim xmlArrearsListNode As IXMLDOMNode
Dim xmlTempChildElem As IXMLDOMElement
Dim xmlCustomerNodeList As IXMLDOMNodeList
Dim xmlMortgageLoanList As IXMLDOMNode
Dim xmlAccount As IXMLDOMNode
Dim xmlCAISNode As IXMLDOMNode
Dim xmlBankruptcyListNode As IXMLDOMNode
Dim xslDoc As FreeThreadedDOMDocument40
Dim xmlExistingOwnersList As IXMLDOMNode
Dim xmlNewRecordsNodeList As IXMLDOMNodeList
Dim xmlApplicationElement As IXMLDOMElement
Dim xmlCustomerListElement As IXMLDOMElement
Dim xmlAccountGuid As IXMLDOMNode   'BMIDS00561 MDC 08/10/2002
Dim xmlTransformedDoc As FreeThreadedDOMDocument40    'BMIDS00801 MDC 09/11/2002
'BM0247 START
Dim strFullCAISAccountStatus As String
Dim intFullCAISAccountStatus As Integer
Dim intMaxAccountStatus As Integer
Dim intTempCAISStatus As Integer
Dim strTempCAISStatus As String
'BM0247 END
'BM0482 START
Dim strCaisCurrentDefaultBal As String
Dim strCCJType As String
Dim strPublicInfoSatisfiedDate As String
'BM0482 END
'BM0479 START
Dim strCaisStartDate As String
Dim strCAISRepayPeriod As String
'BM0479 END
'Dim objCreditCheckBO As CreditCheckBO
Dim objCustFinancialBO As CustomerFinancialBO
Dim objCustomerBO As CustomerBO
Dim objGlobalParamBO As GlobalParameterBO
Dim objComboDO As ComboDO
Dim objApplicationBO As ApplicationBO
Dim colValueIds As Collection

Dim strResponse As String
Dim strForename As String
Dim strSurname As String
Dim strTempValue As String
Dim strTitle As String
Dim strCustomerNumber As String
Dim strCustomerVersionNumber As String
Dim blnUnAssigned As Boolean
Dim strAccStatus As String
Dim strAccType As String
Dim strStartDate As String
Dim strRepayment As String
Dim strBalance As String
Dim intTempValue As Integer
Dim intIndex As Integer
Dim blnCreateNew As Boolean
Dim blnArrearsHistory As Boolean
Dim blnBankruptcyHistory As Boolean
Dim blnCCJHistory As Boolean
Dim blnExistingMortgage As Boolean
Dim blnLoanLiability As Boolean
Dim intTotalCAISRecords As Integer  'BMIDS00561
Dim strDateOfChange As String   'BMIDS00949 MDC 14/11/2002
Dim strCondition As String  'BMIDS00972 MDC 18/11/2002
Dim strTempDate As String   'BM0162 MDC 05/12/2002
'Variables used For BMIDS593 START
Dim xmlCAISListDoc As FreeThreadedDOMDocument40
Dim xmlNewCAISAccounts As IXMLDOMNode
Dim strSettledDate As String
Dim strWorstStatus As String
Dim strTimePeriod As String
Dim bInTimePeriod As Boolean
Dim bArrears As Boolean
Dim intExperianMonths As Integer
Dim strExperianMonths As String
'BMIDS593 END
'Variables used For BM0373 START
Dim xmlAppFFRequest As IXMLDOMElement
Dim xmlAppFFElem As IXMLDOMElement
Dim xmlAppFFNum As IXMLDOMElement
Dim xmlAppNum As IXMLDOMElement
Dim strAppFFResponse As String
Dim xmlGetAppFFDataResponseDoc As FreeThreadedDOMDocument40
Dim strDirectIndirectBusiness As String
Dim strRedemptionStatus As String
Dim strSearch As String
Dim xmlComboResponseDoc As FreeThreadedDOMDocument40
Dim xmlRedemptionElement As IXMLDOMElement
Dim xmlDirectIndirectBusiness  As IXMLDOMElement
Dim blnToBeRedeemed As Boolean
'Variables used For BM0373 END
'GD BM0486 START
Dim xmlTempAccountRelationshipElem As IXMLDOMElement
'GD BM0486 END
'BMIDS630
Dim xmlTempType46Elem As IXMLDOMNode
Dim xmlTempNonType46Elem As IXMLDOMNode
Dim xmlType46NodeList As IXMLDOMNodeList
Dim xmlNonDischargedToRemoveList As IXMLDOMNodeList
Dim xmlFBPublicInfoList As IXMLDOMElement
Dim xmlBHistDoc As FreeThreadedDOMDocument40
Set xmlBHistDoc = New FreeThreadedDOMDocument40
xmlBHistDoc.validateOnParse = False
xmlBHistDoc.setProperty "NewParser", True
Dim xmlBanktuptcyToCreateList As IXMLDOMElement

Dim intType46Index As Integer
Dim int46RemoveIndex As Integer
Dim strDateDeclared As String
Dim strNonDischRemovalPattern As String
Dim strCourtPlaintiff As String
Dim strnon46CreditCheckGUID As String
Dim strnon46FBBlockID As String
Dim strnon46FBHeaderSequence As String
Dim strnon46FBPublicInfoSequence As String
Dim strDeclaredMostRecentNonDisch As String
Dim intDataType As Integer
Dim bNonDischargedToRemove As Boolean
'BMIDS630

Const cstrFunctionName As String = "ImportBureauData"

    On Error GoTo ImportBureauDataVbErr
    
    '=======================================================================================
    ' Validate
    '=======================================================================================
    Set xmlRequestDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
    Set xmlApplicationElement = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, _
            "APPLICATION[(APPLICATIONNUMBER != '') and (APPLICATIONFACTFINDNUMBER != '')]")
    Set xmlCustomerListElement = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, _
            "CUSTOMERLIST[CUSTOMER[(CUSTOMERNUMBER != '') and (CUSTOMERVERSIONNUMBER != '')]]")


    '=======================================================================================
    ' Initialise
    '=======================================================================================
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    'GD 27/02/2003 BM0373 START
    Set xmlComboResponseDoc = New FreeThreadedDOMDocument40
    xmlComboResponseDoc.validateOnParse = False
    xmlComboResponseDoc.setProperty "NewParser", True
    'GD 27/02/2003 BM0373 END
    Set xmlCAISListDoc = New FreeThreadedDOMDocument40   'BMIDS593
    xmlCAISListDoc.validateOnParse = False
    xmlCAISListDoc.setProperty "NewParser", True
    
    xmlResponseDoc.async = False
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseElem

    If Not m_objContext Is Nothing Then
        'Set objCreditCheckBO = m_objContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
        Set objCustFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objGlobalParamBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Else
        'Set objCreditCheckBO = New CreditCheckBO
        Set objCustFinancialBO = New CustomerFinancialBO
        Set objCustomerBO = New CustomerBO
        Set objGlobalParamBO = New GlobalParameterBO
        Set objComboDO = New ComboDO
        Set objApplicationBO = New ApplicationBO
    End If
    
    
    '=======================================================================================
    ' Step A - Get deduplicated bureau data
    '=======================================================================================
    'Call Credit Check component
    
    'BMIDS658 The call to GetAndDeduplicateBureauRecords is now made earlier in omTMBO, and
    'the XML passed through
    'strResponse = objCreditCheckBO.GetAndDeduplicateBureauRecords(vstrXMLRequest)
    'Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    'If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
    
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True

    'Copy the Bureau Data from the request
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "BUREAUDATA")
    Set xmlBureauData = xmlNode.cloneNode(True)
    
    'Remove the bureau data from the request, as the request is used later on and does not
    'expect the bureau data to be there
    xmlRequestDoc.documentElement.removeChild xmlNode

    'BMIDS658 End
        
    If Not xmlBureauData Is Nothing Then
    
        '=======================================================================================
        ' Step B - Alias/Association data
        '=======================================================================================
        
        Set xmlNewRecordsNodeList = xmlBureauData.selectNodes( _
                                            ".//FULLBUREAUASSOCIATIONSLIST/FULLBUREAUASSOCIATIONS")
        
        If xmlNewRecordsNodeList.length > 0 Then
            'Create request
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlNodeList = xmlRequestDoc.selectNodes(".//CUSTOMERLIST/CUSTOMER")
            For Each xmlNode In xmlNodeList
                m_objXmlAssist.CopyNode "CUSTOMER", "ALIAS", xmlNode, xmlTempRequestElem, True
            Next
            
            'Call FindAliasList to find all current Aliases held in Omiga for each customer
            strResponse = objCustomerBO.FindAliasList(xmlTempRequestElem.xml)
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                'Alias records found
                Set xmlAliasListNode = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, _
                                                                       "ALIASPERSONLIST")
            'BMIDS 693 DRC  need to ensure that the aliases are Upper case and no other characters
            Dim xmlCustomerAliasList As IXMLDOMNodeList
                Set xmlCustomerAliasList = xmlAliasListNode.selectNodes(".//ALIASPERSON/ALIAS/PERSON")
                For Each xmlNode In xmlCustomerAliasList
                    Set xmlTempElem = xmlNode.selectSingleNode(".//SURNAME")
                    xmlTempElem.Text = ToUpperAlphaOnly(xmlTempElem.Text)
                    Set xmlTempElem = xmlNode.selectSingleNode(".//FIRSTFORENAME")
                    xmlTempElem.Text = ToUpperAlphaOnly(xmlTempElem.Text)
                Next
            End If
            'BMIDS 693 DRC End
            
            'Create or Update each Alias record
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("ALIASPERSONLIST")
            xmlTempRequestElem.appendChild xmlTempParentElem
            
            For Each xmlNode In xmlNewRecordsNodeList
                'Find owner of this record (alias can only have a single owner)
                Set xmlCustomerElem = m_objXmlAssist.GetNode(xmlNode, "CUSTOMER", False)
                If Not xmlCustomerElem Is Nothing Then
                    If Not xmlAliasListNode Is Nothing Then
                        'Check if there is already a matching record in Omiga
                        strCustomerNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERNUMBER")
                        strCustomerVersionNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERVERSIONNUMBER")
                        'BMIDS 693 DRC  need to ensure that the names are Upper case and no other characters
                        strForename = ToUpperAlphaOnly(m_objXmlAssist.GetElementText(xmlNode, "FBASSOCIATIONFORENAME"))
                        strSurname = ToUpperAlphaOnly(m_objXmlAssist.GetElementText(xmlNode, "FBASSOCIATIONSURNAME"))
                        'BMIDS00949 MDC 14/11/2002
                        strDateOfChange = m_objXmlAssist.GetElementText(xmlNode, "FBASSOCIATIONINFODATE")
                        'BMIDS01055 MDC 22/11/2002 - Handle names with apostrophes
                        Set xmlAliasNode = m_objXmlAssist.GetNode(xmlAliasListNode, _
                                ".//ALIASPERSON/ALIAS[CUSTOMERNUMBER = '" & strCustomerNumber _
                                & "' and CUSTOMERVERSIONNUMBER = '" & strCustomerVersionNumber _
                                & "' and translate(PERSON/FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strForename) _
                                & """ and translate(PERSON/SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strSurname) _
                                & """ and DATEOFCHANGE = '" & strDateOfChange & "']", False)
'                                & "' and PERSON/SURNAME $ieq$ '" & strSurname & "']", False)
                        'BMIDS01055 MDC 22/11/2002 - End
                        'BMIDS00949 MDC 14/11/2002 - End
                    End If
                    
                    'Build request for call to SaveAlias ---------------------------------------
                    Set xmlTempChildElem = xmlRequestDoc.createElement("ALIASPERSON")
                    xmlTempParentElem.appendChild xmlTempChildElem
                    If xmlAliasNode Is Nothing Then
                        Set xmlAliasNode = xmlRequestDoc.createElement("ALIAS")
                        'Customer Number
                        m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlAliasNode, False
                        'Customer Version Number
                        m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlAliasNode, False
                        'Person Guid
                        Set xmlTempElem = xmlRequestDoc.createElement("PERSONGUID")
                        xmlAliasNode.appendChild xmlTempElem
                        ' DRC BMIDS693  - add CreditSearch flag for a new Alias record
                        m_objXmlAssist.SetMandatoryChildText xmlAliasNode, "CREDITSEARCH", "1"
                    End If
                    
                    'Alias Type
                    strTempValue = m_objXmlAssist.GetNodeValue(xmlNode, "FBASSOCIATIONINFOTYPE", False, False)
                    Select Case strTempValue
                    Case "A", "J"
                        strTempValue = "20" 'Association
                    Case Else
                        strTempValue = "10" 'Default to Alias
                    End Select
                    m_objXmlAssist.SetMandatoryChildText xmlAliasNode, "ALIASTYPE", strTempValue
                    
                    'Date of Change
                    strTempValue = m_objXmlAssist.GetElementText(xmlNode, "FBASSOCIATIONINFODATE")
                    If IsDate(strTempValue) Then
                        m_objXmlAssist.SetMandatoryChildText xmlAliasNode, "DATEOFCHANGE", strTempValue
                    End If

                    
                    Set xmlPersonNode = m_objXmlAssist.GetNode(xmlAliasNode, "PERSON", False)
                    If xmlPersonNode Is Nothing Then
                        'Create Person node
                        Set xmlPersonNode = xmlRequestDoc.createElement("PERSON")
                        xmlAliasNode.appendChild xmlPersonNode
                        'Person Guid
                        m_objXmlAssist.CopyNode "PERSONGUID", "PERSONGUID", xmlAliasNode, xmlPersonNode, False
                    End If
                    
                    'Forename
                    strTempValue = m_objXmlAssist.GetNodeValue(xmlNode, "FBASSOCIATIONFORENAME", False, False)
                    m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "FIRSTFORENAME", strTempValue
                    'Surname
                    strTempValue = m_objXmlAssist.GetNodeValue(xmlNode, "FBASSOCIATIONSURNAME", False, False)
                    m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "SURNAME", strTempValue
                    'Other forenames
                    strTempValue = m_objXmlAssist.GetNodeValue(xmlNode, "FBASSOCIATIONINITIALS", False, False)
                    m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "OTHERFORENAMES", strTempValue
                    'Title
                    strTitle = m_objXmlAssist.GetNodeValue(xmlNode, "FBASSOCIATIONTITLE", False, False)
                    
                    'BMIDS00919 MDC 12/11/2002
                    If Len(Trim$(strTitle)) > 0 Then
                        strTempValue = objComboDO.GetFirstComboValueIdFromValueName("Title", strTitle)
                    Else
                        strTempValue = vbNullString
                    End If
                    'BMIDS00919 MDC 12/11/2002 - End
                    
                    If Len(Trim$(strTempValue)) = 0 Then
                        'Not found so use 'Other'
                        m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "TITLE", objComboDO.GetFirstComboValueId("Title", "O")
                        m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "TITLEOTHER", strTitle
                    Else
                        m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "TITLE", strTempValue
                        m_objXmlAssist.SetMandatoryChildText xmlPersonNode, "TITLEOTHER", vbNullString
                    End If
                    
                    xmlTempChildElem.appendChild xmlAliasNode.cloneNode(True)
                    Set xmlAliasNode = Nothing
                End If
            Next
                
            If xmlTempParentElem.hasChildNodes Then
                'Call SaveAlias
                strResponse = objCustomerBO.SaveAlias(xmlTempRequestElem.xml)
                m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
            End If
            
        End If
        
        'Update the CustomerVersion with Alias indicator and Voters Roll
        Set xmlNodeList = xmlRequestDoc.selectNodes(".//CUSTOMERLIST/CUSTOMER")
        For Each xmlNode In xmlNodeList
            'BMIDS00985 MDC 18/11/2002
            strCustomerNumber = m_objXmlAssist.GetElementText(xmlNode, "CUSTOMERNUMBER")
            strCustomerVersionNumber = m_objXmlAssist.GetElementText(xmlNode, "CUSTOMERVERSIONNUMBER")
            'BMIDS00985 MDC 18/11/2002 - End
        
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("UPDATE")
            xmlTempRequestElem.appendChild xmlTempParentElem
            
            Set xmlCustomerElem = xmlRequestDoc.createElement("CUSTOMER")
            xmlTempParentElem.appendChild xmlCustomerElem
            
            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlNode, xmlCustomerElem, False
            
            Set xmlTempChildElem = xmlRequestDoc.createElement("CUSTOMERVERSION")
            xmlCustomerElem.appendChild xmlTempChildElem
            
            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlNode, xmlTempChildElem, False
            m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlNode, xmlTempChildElem, False
            
            Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
            xmlTempElem.Text = "1"
            xmlTempChildElem.appendChild xmlTempElem
            
            'Are there one or more aliases for this customer?
            Set xmlTempElem = xmlRequestDoc.createElement("ALIASINDICATOR")
            If xmlBureauData.selectNodes("FULLBUREAUASSOCIATIONSLIST/FULLBUREAUASSOCIATIONS/CUSTOMER" _
                & "[CUSTOMERNUMBER = '" & strCustomerNumber & "' and CUSTOMERVERSIONNUMBER = '" _
                & strCustomerVersionNumber & "']").length > 0 Then
                'An alias exists
                xmlTempElem.Text = "1"
            Else
                xmlTempElem.Text = "0"
            End If
            xmlTempChildElem.appendChild xmlTempElem
            
            strTempValue = m_objXmlAssist.GetNodeValue(xmlBureauData, _
                "FULLBUREAUVOTERSROLLLIST/FULLBUREAUVOTERSROLL[(CUSTOMERLIST/CUSTOMER/" _
                & "CUSTOMERNUMBER='" & strCustomerNumber & "') and (CUSTOMERLIST/CUSTOMER/" _
                & "CUSTOMERVERSIONNUMBER = '" & strCustomerVersionNumber & "')]/FBREGISTRATIONDATE1", False, False)
                
            If Len(Trim$(strTempValue)) > 0 And IsNumeric(strTempValue) Then
                'Convert 2 digit year to 4 digit year
                If CInt(strTempValue) <= 30 Then    'BMIDS00985 MDC 19/11/2002
                    strTempValue = "20" & strTempValue
                Else
                    strTempValue = "19" & strTempValue
                End If
                Set xmlTempElem = xmlRequestDoc.createElement("YEARADDEDTOVOTERSROLL")
                xmlTempElem.Text = strTempValue
                xmlTempChildElem.appendChild xmlTempElem
            End If
                
            'Save the updated Customer details
            strResponse = objCustomerBO.UpdateCustomerDetails(xmlTempRequestElem.xml)
            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        Next

        
        '=======================================================================================
        ' Step C - Account data
        '=======================================================================================
    
        'BMIDS00561 MDC 07/10/2002
        intTotalCAISRecords = xmlBureauData.selectNodes(".//FULLBUREAUCAISLIST/FULLBUREAUCAIS").length
        'BMIDS00561 MDC 07/10/2002 - End
        
        'Get each Bureau CAIS record

'       BMIDS0593 START
'                 When selecting the CAIS accounts to process, also need to select the records with
'                 arrears which have been settled in the last 'ExperianNoOfMonthsCAISSettled' months
'                 Get all the records where OwnDataFlag <> 'Y' here.

'       Set xmlNewRecordsNodeList = xmlBureauData.selectNodes(".//FULLBUREAUCAISLIST/" _
'                & "FULLBUREAUCAIS[FBCAISOWNDATAFLAG != 'Y' and (FBCAISSETTLEDDATE = '' or " _
'                & "FBCAISWORSTSTATUS = '8' or FBCAISWORSTSTATUS = '9')]")
        Set xmlNewRecordsNodeList = xmlBureauData.selectNodes(".//FULLBUREAUCAISLIST/" _
                 & "FULLBUREAUCAIS[FBCAISOWNDATAFLAG != 'Y']")
    
        If xmlNewRecordsNodeList.length > 0 Then
        
            'Now create a node to contain all the records we need
            
            Set xmlNewCAISAccounts = xmlCAISListDoc.createElement("FULLBUREAUCAISLIST")
            
            For Each xmlCAISNode In xmlNewRecordsNodeList
            
                strSettledDate = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISSETTLEDDATE")
                'INR BMIDS593 FBCAISWorstStatus should be all uppercase
                strWorstStatus = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISWORSTSTATUS")

                'We need this account if the worst status is 8 or 9 OR the settled date is zero or blank
                If (strWorstStatus = "8" Or strWorstStatus = "9" Or _
                    (Len(strSettledDate) = 0)) Then
                    'INR BMIDS593
'                    strSettledDate = "0" Or strSettledDate = " ") Then
                    
                    xmlNewCAISAccounts.appendChild xmlCAISNode
                Else
                    'Check if this is an account with arrears which has been settled within the last 'n' months
                
                    strAccStatus = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTSTATUS")

                    'Check if this account has been in arrears by looking atFBCAISAccountStatus
                    bArrears = False
                    For intIndex = 1 To Len(strAccStatus)
                        strTempCAISStatus = Mid$(strAccStatus, intIndex, 1)
                        If IsNumeric(strTempCAISStatus) Then
                            intTempCAISStatus = CInt(strTempCAISStatus)
                            If (intTempCAISStatus > 0) And (intTempCAISStatus < 8) Then
                                bArrears = True
                                Exit For
                            End If
                        End If
                    Next
                
                    'Calculate the time we are interested in: ExperianNoOfMonthsCAISSettled months before today.
                    bInTimePeriod = False
                    
                    strExperianMonths = objGlobalParamBO.GetCurrentParameter("ExperianNoOfMonthsCAISSettled")
                    Set xmlTempResponseDoc = m_objXmlAssist.load(strExperianMonths, TypeName(Me), cstrFunctionName)
                    If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
                        strExperianMonths = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='" & "ExperianNoOfMonthsCAISSettled" & "']/AMOUNT")
                    End If
     
                    intExperianMonths = CInt(strExperianMonths)
                    
                    strTimePeriod = DateAdd("m", (-1 * intExperianMonths), Now)
                
                    'If this account has been settled in the last 'n' months,
                    If CDate(strSettledDate) >= CDate(strTimePeriod) Then
                        bInTimePeriod = True
                    End If
                
                    If (bInTimePeriod = True And bArrears = True) Then
                    
                        'we need this record
                        ' DRC - BMIDS669 settled CAIS record with arrears the monthly payment is set to 0
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlCAISNode, "FBCAISMONTHLYPAYMENT", False)
                        If Not xmlTempElem Is Nothing Then
                            xmlTempElem.Text = "0"
                        End If
                        xmlNewCAISAccounts.appendChild xmlCAISNode
                    End If
                End If
            Next
        End If

        ' BMIDSINR Check this is not nothing before trying to use it
        If Not xmlNewCAISAccounts Is Nothing Then
          If xmlNewCAISAccounts.hasChildNodes Then
        'BMIDS593 End
                 
            'Get existing MortgageAccounts
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("MORTGAGEACCOUNTLIST")
            xmlTempRequestElem.appendChild xmlTempParentElem
            Set xmlNodeList = xmlRequestDoc.documentElement.selectNodes(".//CUSTOMERLIST/CUSTOMER")
            For Each xmlNode In xmlNodeList
                m_objXmlAssist.CopyNode "CUSTOMER", "MORTGAGEACCOUNT", xmlNode, xmlTempParentElem, True
            Next
            strResponse = objCustFinancialBO.FindMortgageAccountList(xmlTempRequestElem.xml)
            
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                'MortgageAccount records found
                Set xmlMortgageAccountListNode = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, _
                                                                                        "MORTGAGEACCOUNTLIST")
                'Find list of MortgageLoans for each Account
                For Each xmlTempElem In xmlMortgageAccountListNode.selectNodes("MORTGAGEACCOUNT")
                    Set xmlTempParentElem = xmlRequestDoc.createElement("MORTGAGEACCOUNTLOANTOTALS")
                    m_objXmlAssist.CopyNode "ACCOUNTGUID", "ACCOUNTGUID", xmlTempElem, xmlTempParentElem, False
                    strResponse = objCustFinancialBO.GetMortgageAccountLoanTotals(xmlTempParentElem.xml)
                    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                    If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                        Set xmlNode = m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, "MORTGAGEACCOUNTLOANTOTALS", False)
                        If Not xmlNode Is Nothing Then
                            xmlTempElem.appendChild xmlNode
                        End If
                    End If
                Next
            
            End If
        
            'Get existing LoanLiabilities
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("LOANSLIABILITIESLIST")
            xmlTempRequestElem.appendChild xmlTempParentElem
            Set xmlNodeList = xmlRequestDoc.documentElement.selectNodes(".//CUSTOMERLIST/CUSTOMER")
            For Each xmlNode In xmlNodeList
                m_objXmlAssist.CopyNode "CUSTOMER", "LOANSLIABILITIES", xmlNode, xmlTempParentElem, True
            Next
            strResponse = objCustFinancialBO.FindLiabilityList(xmlTempRequestElem.xml)
            
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                'LoanLiabilities records found
                Set xmlLiabilityListNode = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, _
                                                                                        "LOANSLIABILITIESLIST")
            End If
        
            'Get existing ArrearsHistory
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("ARREARSHISTORYLIST")
            Set xmlNodeList = xmlRequestDoc.documentElement.selectNodes(".//CUSTOMERLIST/CUSTOMER")
            For Each xmlNode In xmlNodeList
                m_objXmlAssist.CopyNode "CUSTOMER", "ARREARSHISTORY", xmlNode, xmlTempParentElem, True
            Next
            xmlTempRequestElem.appendChild xmlTempParentElem
            strResponse = objCustFinancialBO.FindArrearsHistoryList(xmlTempRequestElem.xml)
            
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                'ArrearsHistory records found
                Set xmlArrearsListNode = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, _
                                                                                        "ARREARSHISTORYLIST")
            End If
        
'            For Each xmlCAISNode In xmlNewRecordsNodeList  'BMIDS593
            For Each xmlCAISNode In xmlNewCAISAccounts.selectNodes(".//FULLBUREAUCAIS")
            
                'Find owner of this record
                Set xmlCustomerNodeList = xmlCAISNode.selectNodes("CUSTOMERLIST/CUSTOMER")
                If xmlCustomerNodeList.length > 0 Then
                    'Customer(s) found
                    blnUnAssigned = False
                Else
                    'Default assignment is to Customer 1
                    Set xmlCustomerElem = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "CUSTOMERLIST/CUSTOMER[1]")
                    Set xmlTempParentElem = m_objXmlAssist.GetMandatoryNode(xmlCAISNode, "CUSTOMERLIST")
                    xmlTempParentElem.appendChild xmlCustomerElem.cloneNode(True)
                    blnUnAssigned = True
                    Set xmlCustomerNodeList = xmlCAISNode.selectNodes("CUSTOMERLIST/CUSTOMER")
                End If
        
                strTempValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCMLADDRESSTYPE")
                strAccStatus = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTSTATUS")
                
                'Check if a Repossession
                If strTempValue = "R" Then
                    'GD BM0486 START
                    'Check for existing record in Omiga
                    strTempValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISSETTLEDDATE")
                    If Not xmlArrearsListNode Is Nothing Then
                        Set xmlTempParentElem = m_objXmlAssist.GetNode(xmlArrearsListNode, "ARREARSHISTORY[DATECLEARED = '" & strTempValue & "']", False)
                    Else
                        Set xmlTempParentElem = Nothing
                    End If
                    If Not xmlTempParentElem Is Nothing Then
                        'Matching Omiga record found so update it
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlCAISNode, "FBCAISBALANCE", False)
                        If Not xmlTempElem Is Nothing Then
                            xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlTempParentElem, "MAXIMUMBALANCE")
                        Else
                            m_objXmlAssist.CopyNode "FBCAISBALANCE", "MAXIMUMBALANCE", xmlCAISNode, xmlTempParentElem, False
                        End If
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlCAISNode, "FBCAISMONTHLYPAYMENT", False)
                        If Not xmlTempElem Is Nothing Then
                            xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlTempParentElem, "MONTHLYREPAYMENT")
                        Else
                            m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                        End If
                        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                        xmlTempRequestElem.appendChild xmlTempParentElem
                        
                        'GD BM0486 START
                        'Remove ACCOUNTRELATIONSHIP node so that a record is not inserted into ACCOUNTRELATIONSHIP table
                        ' when the UpdateArrearsHistory is called.

                        Set xmlTempAccountRelationshipElem = xmlTempRequestElem.selectSingleNode(".//ACCOUNTRELATIONSHIP")
                        If Not (xmlTempAccountRelationshipElem Is Nothing) Then
                            xmlTempRequestElem.selectSingleNode(".//ARREARSHISTORY").removeChild xmlTempAccountRelationshipElem
                        End If
                        'GD BM0486 END
                        
                        'Update the ArrearsHistory record
                        strResponse = objCustFinancialBO.UpdateArrearsHistory(xmlTempRequestElem.xml)
                        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                        Debug.Print "Updated Arrears History record"

                    Else
                    'GD BM0486 END
                        'Create Other Arrears Account and Arrears History record
                        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                        Set xmlTempParentElem = xmlRequestDoc.createElement("ARREARSHISTORY")
                        xmlTempRequestElem.appendChild xmlTempParentElem
                        
                        Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                        xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                        xmlTempParentElem.appendChild xmlTempElem
                        Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                        xmlTempElem.Text = "1"
                        xmlTempParentElem.appendChild xmlTempElem
                        Set xmlTempElem = xmlRequestDoc.createElement("REPOSSESSIONIND")
                        xmlTempElem.Text = "1"
                        xmlTempParentElem.appendChild xmlTempElem
                        m_objXmlAssist.CopyNode "FBCAISSETTLEDDATE", "DATECLEARED", xmlCAISNode, xmlTempParentElem, False
                        m_objXmlAssist.CopyNode "FBCAISBALANCE", "MAXIMUMBALANCE", xmlCAISNode, xmlTempParentElem, False
                        m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                        
                        Set xmlTempChildElem = xmlRequestDoc.createElement("OTHERARREARSACCOUNT")
                        xmlTempParentElem.appendChild xmlTempChildElem
                        Set xmlTempElem = xmlRequestDoc.createElement("DESCRIPTIONOFLOAN")
                        'BM0158 MDC 05/12/2002
                        xmlTempElem.Text = objComboDO.GetFirstComboValueId("ArrearsLoanType", "R")

                        'xmlTempElem.Text = "Repossession"
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ADDITIONALINDICATOR", "0"
                        'BM0158 MDC 05/12/2002 - End
                        xmlTempChildElem.appendChild xmlTempElem
                        'GD BM0486 START
                        'Ensure DESCRIPTIONOFLOAN lives under ARREARSHISTORY, as well as OTHERARREARSACCOUNT
                        xmlTempParentElem.appendChild xmlTempElem.cloneNode(True)
                        'GD BM0486 END
                        If blnUnAssigned Then
                            Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                            xmlTempParentElem.appendChild xmlTempChildElem
                            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                        Else
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                xmlTempParentElem.appendChild xmlTempChildElem
                                m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                                m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            Next
                        End If
                        
                        strResponse = objCustFinancialBO.CreateArrearsHistory(xmlTempRequestElem.xml)
                        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                        Debug.Print "Created Arrears History record"
                        blnArrearsHistory = True
                    End If
                'Check if Arrears
                ElseIf strTempValue = "A" Or strTempValue = "F" Then
                    'Check for existing record in Omiga
                    strTempValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISSETTLEDDATE")
                    'BMIDS01124 MDC 02/12/2002
                    If Not xmlArrearsListNode Is Nothing Then
                        Set xmlTempParentElem = m_objXmlAssist.GetNode(xmlArrearsListNode, "ARREARSHISTORY[DATECLEARED = '" & strTempValue & "']", False)
                    Else
                        Set xmlTempParentElem = Nothing
                    End If
                    'BMIDS01124 MDC 02/12/2002 - End
                    If Not xmlTempParentElem Is Nothing Then
                        'Matching Omiga record found so update it
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlCAISNode, "FBCAISBALANCE", False)
                        If Not xmlTempElem Is Nothing Then
                            xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlTempParentElem, "MAXIMUMBALANCE")
                        Else
                            m_objXmlAssist.CopyNode "FBCAISBALANCE", "MAXIMUMBALANCE", xmlCAISNode, xmlTempParentElem, False
                        End If
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlCAISNode, "FBCAISMONTHLYPAYMENT", False)
                        If Not xmlTempElem Is Nothing Then
                            xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlTempParentElem, "MONTHLYREPAYMENT")
                        Else
                            m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                        End If
                        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                        xmlTempRequestElem.appendChild xmlTempParentElem
                        
                        'GD BM0486 START
                        'Remove ACCOUNTRELATIONSHIP node so that a record is not inserted into ACCOUNTRELATIONSHIP table
                        ' when the UpdateArrearsHistory is called.

                        Set xmlTempAccountRelationshipElem = xmlTempRequestElem.selectSingleNode(".//ACCOUNTRELATIONSHIP")
                        If Not (xmlTempAccountRelationshipElem Is Nothing) Then
                            xmlTempRequestElem.selectSingleNode(".//ARREARSHISTORY").removeChild xmlTempAccountRelationshipElem
                        End If
                        'GD BM0486 END
                        
                        'Update the ArrearsHistory record
                        
                        
                        strResponse = objCustFinancialBO.UpdateArrearsHistory(xmlTempRequestElem.xml)
                        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                        Debug.Print "Updated Arrears History record"
                        
                    Else
                        'Create a new arrears history record
                        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                        Set xmlTempParentElem = xmlRequestDoc.createElement("ARREARSHISTORY")
                        xmlTempRequestElem.appendChild xmlTempParentElem
                        
                        Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                        xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                        xmlTempChildElem.appendChild xmlTempElem
                        Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                        xmlTempElem.Text = "1"
                        xmlTempChildElem.appendChild xmlTempElem
                        m_objXmlAssist.CopyNode "FBCAISSETTLEDDATE", "DATECLEARED", xmlCAISNode, xmlTempParentElem, False
                        m_objXmlAssist.CopyNode "FBCAISBALANCE", "MAXIMUMBALANCE", xmlCAISNode, xmlTempParentElem, False
                        m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                        'BM0158 MDC 05/12/2002
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ADDITIONALINDICATOR", "0"
                        'BM0158 MDC 05/12/2002 - End
                        
                        Set xmlTempChildElem = xmlRequestDoc.createElement("OTHERARREARSACCOUNT")
                        xmlTempParentElem.appendChild xmlTempChildElem
                        Set xmlTempElem = xmlRequestDoc.createElement("DESCRIPTIONOFLOAN")
                        'BM0158 MDC 05/12/2002
                        xmlTempElem.Text = objComboDO.GetFirstComboValueId("ArrearsLoanType", "R")
                        'xmlTempElem.Text = "Repossession"
                        'BM0158 MDC 05/12/2002 - End
                        xmlTempChildElem.appendChild xmlTempElem
                        'GD BM0486 START
                        'Ensure DESCRIPTIONOFLOAN lives under ARREARSHISTORY, as well as OTHERARREARSACCOUNT
                        xmlTempParentElem.appendChild xmlTempElem.cloneNode(True)
                        'GD BM0486 END
                        '#TASK - Do we need to distinguish between Unassigned and not
                        If blnUnAssigned Then
                            Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                            xmlTempParentElem.appendChild xmlTempChildElem
                            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                        Else
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                xmlTempParentElem.appendChild xmlTempChildElem
                                m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                                m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            Next
                        End If
                        
                        strResponse = objCustFinancialBO.CreateArrearsHistory(xmlTempRequestElem.xml)
                        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                        Debug.Print "Created Arrears History record"
                        blnArrearsHistory = True
                    End If
                    
                'Check if CCJ
                'BM0529 Include check for status 9 Bad Debt.
                ElseIf (Left$(strAccStatus, 1) = "8") Or (Left$(strAccStatus, 1) = "9") Then
                    'Create CCJHistory record
                    Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                    Set xmlTempParentElem = xmlRequestDoc.createElement("CCJHISTORY")
                    xmlTempRequestElem.appendChild xmlTempParentElem
                    
                    Set xmlTempElem = xmlRequestDoc.createElement("DEFAULTRECORD")
                    xmlTempElem.Text = "1"
                    xmlTempParentElem.appendChild xmlTempElem
                    Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                    xmlTempElem.Text = "1"
                    xmlTempParentElem.appendChild xmlTempElem
                    Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                    xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                    xmlTempParentElem.appendChild xmlTempElem
                    'GD BM0482 START 1
                    'Original Code Start
'                    m_objXmlAssist.CopyNode "FBCAISSETTLEDDATE", "DATECLEARED", xmlCAISNode, xmlTempParentElem, False
'                    m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
'                    m_objXmlAssist.CopyNode "FBCAISBALANCE", "VALUEOFJUDGEMENT", xmlCAISNode, xmlTempParentElem, False
'                    m_objXmlAssist.CopyNode "FBCAISSTARTDATE", "DATEOFJUDGEMENT", xmlCAISNode, xmlTempParentElem, False
                    'Original Code End
                    'Make DATECLEARED = NULL, so remove this line : m_objXmlAssist.CopyNode "FBCAISSETTLEDDATE", "DATECLEARED", xmlCAISNode, xmlTempParentElem, False
                    m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                    m_objXmlAssist.CopyNode "FBCAISBALANCE", "VALUEOFJUDGEMENT", xmlCAISNode, xmlTempParentElem, False
                    'Set DATEOFJUDGEMENT = FBCAISSETTLEDDATE
                    m_objXmlAssist.CopyNode "FBCAISSETTLEDDATE", "DATEOFJUDGEMENT", xmlCAISNode, xmlTempParentElem, False
                    'GD BM0482 END 1
                    'GD BM0482 START 2
                    'Get FBCAISCURRENTDEFAULTBAL
                    strCaisCurrentDefaultBal = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCURRENTDEFAULTBAL")
                    If CLng(strCaisCurrentDefaultBal) = 0 Then      'BMIDS590
                        'Set CCJTYPE = Combo Value ID where validation type is 'S'
                        strCCJType = objComboDO.GetFirstComboValueId("CCJType", "S")
                    Else ' greater than zero
                        'Set CCJTYPE = Combo Value ID where validation type is 'U'
                        strCCJType = objComboDO.GetFirstComboValueId("CCJType", "U")
                    End If
                    m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "CCJTYPE", strCCJType
                    'GD BM0482 END 2
                    'BM0158 MDC 05/12/2002
                    m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ADDITIONALINDICATOR", "0"
                    'BM0158 MDC 05/12/2002 - End
                    
                    'BMIDS712 Add CAISACCOUNTTYPE to CCJ History Record
                    m_objXmlAssist.CopyNode "FBCAISACCOUNTTYPE", "CAISACCOUNTTYPE", xmlCAISNode, xmlTempParentElem, False
                    
                    For Each xmlCustomerElem In xmlCustomerNodeList
                        Set xmlTempChildElem = xmlRequestDoc.createElement("CUSTOMERVERSIONCCJHISTORY")
                        xmlTempParentElem.appendChild xmlTempChildElem
                        m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                        m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                    Next
                    
                    strResponse = objCustFinancialBO.CreateCCJHistory(xmlTempRequestElem.xml)
                    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                    Debug.Print "Created CCJ History record"
                    blnCCJHistory = True
                    
                Else
                    'Check if Mortgage Account
                    strAccType = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTTYPE")
                    If strAccType = "03" Or strAccType = "16" Then
                        'Mortgage Account
                        
                        strStartDate = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISSTARTDATE", False, False)
                        strRepayment = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISMONTHLYPAYMENT", False, False)
                        strBalance = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISBALANCE", False, False)
                        
                        'BMIDS00919 MDC 12/11/2002
                        If Not xmlMortgageAccountListNode Is Nothing Then
                            'Try to find existing matching Omiga 4 record
                            Set xmlAccount = m_objXmlAssist.GetNode(xmlMortgageAccountListNode, _
                                "MORTGAGEACCOUNT[MORTGAGEACCOUNTLOANTOTALS/STARTDATE = '" & strStartDate & "'" _
                                & " and MORTGAGEACCOUNTLOANTOTALS/OUTSTANDINGBALANCE = '" & Format$(strBalance, "#######0.00") & "'" _
                                & " and MORTGAGEACCOUNTLOANTOTALS/MONTHLYREPAYMENT = '" & Format$(strRepayment, "#######0.00") & "']", False)
                        Else
                            Set xmlAccount = Nothing
                        End If
                        'BMIDS00919 MDC 12/11/2002 - End
                        
                        If Not xmlAccount Is Nothing Then
                            'Update existing account --------------------------------------------------
                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            
                            'BMIDS00919 MDC 12/11/2002
                            Set xmlTempElem = xmlRequestDoc.createElement("DIRECTORYADDRESSINDICATOR")
                            If Len(Trim$(m_objXmlAssist.GetElementText(xmlAccount, "DIRECTORYGUID"))) > 0 Then
                                xmlTempElem.Text = "1"
                            Else
                                xmlTempElem.Text = "0"
                            End If
                            xmlTempRequestElem.appendChild xmlTempElem
                            'BMIDS00919 MDC 12/11/2002 - End
                            
                            'Mortgage Account
                            xmlTempRequestElem.appendChild xmlAccount
                            
                            'Account
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ACCOUNT")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            m_objXmlAssist.CopyNode "ACCOUNTGUID", "ACCOUNTGUID", xmlAccount, xmlTempParentElem, False
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                            xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("IMPORTEDINDICATOR")
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            'BMIDS00561 MDC 07/10/2002
'                            Set xmlTempElem = xmlRequestDoc.createElement("ACCOUNTNUMBER")
'                            xmlTempElem.Text = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBBUREAUREFERENCE", False, False)
'                            xmlTempParentElem.appendChild xmlTempElem
                            strTempValue = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBBUREAUREFERENCE", False, False)
                            strTempValue = strTempValue & " (" _
                                    & m_objXmlAssist.GetMandatoryElementText(xmlCAISNode, "FBHEADERSEQUENCE") _
                                    & " of " & CStr(intTotalCAISRecords) & ")"
                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ACCOUNTNUMBER", Left$(strTempValue, 20)
                            'BMIDS00561 MDC 07/10/2002 - End
                            
                            'Account Relationships
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIPLIST")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                xmlTempParentElem.appendChild xmlTempChildElem
                                m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                                m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            Next
                            
                            strResponse = objCustFinancialBO.UpdateMortgageAccount(xmlTempRequestElem.xml)
                            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                            Debug.Print "Updated Mortgage Account record"
                            
                            'Get MortgageLoan list
                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            Set xmlTempParentElem = xmlRequestDoc.createElement("MORTGAGELOANLIST")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            Set xmlTempChildElem = xmlRequestDoc.createElement("MORTGAGELOAN")
                            xmlTempParentElem.appendChild xmlTempChildElem
                            m_objXmlAssist.CopyNode "ACCOUNTGUID", "ACCOUNTGUID", xmlAccount, xmlTempChildElem, False
                            
                            strResponse = objCustFinancialBO.FindMortgageLoanList(xmlTempRequestElem.xml)
                            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                            m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
                            
                            intTempValue = CInt(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISREPAYPERIOD"))
                            If intTempValue > 0 And IsDate(strStartDate) Then
                                strTempValue = DateAdd("m", intTempValue, strStartDate)
                            Else
                                strTempValue = DateAdd("yyyy", 5, Now)  'BM0168 MDC 06/12/2002 - correct year format
                            End If
                            'BM0168 MDC 05/12/2002 - Ensure date is formated as date only (i.e. no time element)
                            strTempValue = Format$(strTempValue, "dd/mm/yyyy")
                            'BM0168 MDC 05/12/2002 - End
                                
                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            Set xmlTempParentElem = xmlRequestDoc.createElement("MORTGAGELOANLIST")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            
                            'Update Mortgage Loans -------------------------------------------------------
                            For Each xmlNode In xmlTempResponseDoc.documentElement.selectNodes("MORTGAGELOANLIST/MORTGAGELOAN")
                                Set xmlTempChildElem = xmlNode.cloneNode(True)
                                If IsDate(strStartDate) Then
                                    m_objXmlAssist.SetMandatoryChildText xmlTempChildElem, "STARTDATE", strStartDate
                                End If
                                m_objXmlAssist.SetMandatoryChildText xmlTempChildElem, "LOANENDDATE", strTempValue
                                xmlTempParentElem.appendChild xmlTempChildElem
'''''                                'GD 26/02/2003 BM0373 START

'''''                                'Build xmlAppFFRequest
'''''                                Set xmlAppFFRequest = xmlRequestDoc.documentElement.cloneNode(False)
'''''                                Set xmlAppFFElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
'''''                                Set xmlAppFFNum = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "APPLICATION/APPLICATIONFACTFINDNUMBER")
'''''                                Set xmlAppNum = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "APPLICATION/APPLICATIONNUMBER")
'''''                                xmlAppFFRequest.appendChild xmlAppFFElem
'''''                                xmlAppFFElem.appendChild xmlAppFFNum
'''''                                xmlAppFFElem.appendChild xmlAppNum
'''''
'''''                                'Call omApp.ApplicationBO.GetApplicationFactFindData(ApplicationNumber,ApplicationfactFindNumber)
'''''                                strAppFFResponse = objApplicationBO.GetApplicationFactFinddata(xmlAppFFRequest.xml)
'''''                                'Load Response into a FreeThreadedDOMDocument40
'''''                                'Check if call has failed, and raise an error if it has
'''''                                Set xmlGetAppFFDataResponseDoc = m_objXmlAssist.load(strAppFFResponse, TypeName(Me), cstrFunctionName)
'''''                                m_objErrAssist.CheckResponse xmlGetAppFFDataResponseDoc.documentElement, xmlResponseElem, True
'''''                                'ApplicationFactFind Data returned OK, if we reach here
'''''                                'Extract strDirectIndirectBusiness
'''''                                'strDirectIndirectBusiness = Get ApplicationfactFind.DirectIndirectBusiness
'''''                                strDirectIndirectBusiness = xmlGetAppFFDataResponseDoc.documentElement.selectSingleNode(".//DIRECTINDIRECTBUSINESS").Text
'''''                                'if strDirectIndirectBusiness has a ValidationType of 7 (aXess) then
'''''                                If objComboDO.IsItemInValidation("Direct/InDirect", strDirectIndirectBusiness, "7") Then
'''''                                '   strRemdemptionStatus = ValueID with ValidationTypes = 7,R (='To be Redeemed - this advance')
'''''
'''''                                    strSearch = "LIST/LISTNAME/LISTENTRY[(VALIDATIONTYPELIST/VALIDATIONTYPE='R') or ( VALIDATIONTYPELIST/VALIDATIONTYPE='7')]/VALUEID"
'''''                                    strResponse = objComboDO.GetComboList("<LIST><LISTNAME>RedemptionStatus</LISTNAME></LIST>")
'''''                                    xmlComboResponseDoc.loadXML (strResponse)
'''''                                    Set xmlRedemptionElement = xmlComboResponseDoc.selectSingleNode(strSearch)
'''''                                    If xmlRedemptionElement Is Nothing Then
'''''                                        'RaiseError
'''''                                        m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4NoComboTagValue, "ValueID for validation types '7' and 'R'  not found"
'''''                                    Else
'''''                                        strRedemptionStatus = xmlRedemptionElement.Text
'''''                                    End If
'''''
'''''                                Else
'''''                                '   strRemdemptionStatus = ValueID with ValidationTypes = N (='NOT To be Redeemed')
'''''                                    strRedemptionStatus = objComboDO.GetComboValueId("Direct/Indirect", "N")
'''''                                End If
'''''
'''''                                'Create REDEMPTIONSTATUS Element with value = strRedemptionStatus
'''''                                Set xmlRedemptionElement = xmlRequestDoc.createElement("REDEMPTIONSTATUS")
'''''                                xmlRedemptionElement.Text = strRedemptionStatus
'''''                                'Append REDEMPTIONSTSTUS Element to xmlTempParentElem
'''''                                xmlTempParentElem.appendChild xmlRedemptionElement
'''''                                'GD 26/02/2003 BM0373 END
                                strResponse = objCustFinancialBO.UpdateMortgageLoan(xmlTempRequestElem.xml)
                                m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                                Debug.Print "Updated Mortgage loan record"
                                
                                xmlTempParentElem.removeChild xmlTempChildElem
                            Next
                        
                        Else
                            'Create new account -----------------------------------------------------
                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            
                            'BMIDS00919 MDC 12/11/2002
                            Set xmlTempElem = xmlRequestDoc.createElement("DIRECTORYADDRESSINDICATOR")
                            xmlTempElem.Text = "1" 'JD BMIDS897
                            xmlTempRequestElem.appendChild xmlTempElem
                            
                            Dim xmlMortgageAccountNode As IXMLDOMNode
                            Set xmlMortgageAccountNode = xmlRequestDoc.createElement("MORTGAGEACCOUNT")
                            xmlTempRequestElem.appendChild xmlMortgageAccountNode
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("BMIDSACCOUNT")
                            xmlTempElem.Text = "0"
                            xmlMortgageAccountNode.appendChild xmlTempElem
                            
                            'BM0158 MDC 05/12/2002
                            m_objXmlAssist.SetMandatoryChildText xmlMortgageAccountNode, "ADDITIONALINDICATOR", "0"
                            'BM0158 MDC 05/12/2002 - End
                            
                            'Account
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ACCOUNT")
                            xmlMortgageAccountNode.appendChild xmlTempParentElem
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                            xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            Set xmlTempElem = xmlRequestDoc.createElement("IMPORTEDINDICATOR")
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            'BMIDS00561 MDC 07/10/2002
'                            Set xmlTempElem = xmlRequestDoc.createElement("ACCOUNTNUMBER")
'                            xmlTempElem.Text = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBBUREAUREFERENCE", False, False)
'                            xmlTempParentElem.appendChild xmlTempElem
                            strTempValue = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBBUREAUREFERENCE", False, False)
                            strTempValue = strTempValue & " (" _
                                    & m_objXmlAssist.GetMandatoryElementText(xmlCAISNode, "FBHEADERSEQUENCE") _
                                    & " of " & CStr(intTotalCAISRecords) & ")"
                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ACCOUNTNUMBER", Left$(strTempValue, 20)
                            'BMIDS00561 MDC 07/10/2002 - End
                            
                            'JD BMIDS897 No longer create a new thirdparty. Instead, find
                            ' an existing DirectoryGuid to attach to the Account
                            strTempValue = Trim$(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCOMPANYTYPE"))
                            Set xmlTempElem = xmlRequestDoc.createElement("DIRECTORYGUID")
                            xmlTempElem.Text = GetDirectoryGUIDForBureau(strTempValue)
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            
                            'Third Party
'                            Set xmlTempChildElem = xmlRequestDoc.createElement("THIRDPARTY")
'                            xmlTempParentElem.appendChild xmlTempChildElem
'                            Set xmlTempElem = xmlRequestDoc.createElement("THIRDPARTYTYPE")
'                            'xmlTempElem.Text = GetFirstComboValueId("ThirdPartyType", "T")
'                            xmlTempElem.Text = "3"
'                            xmlTempChildElem.appendChild xmlTempElem
'
'                            Set xmlTempElem = xmlRequestDoc.createElement("COMPANYNAME")
'                            xmlTempChildElem.appendChild xmlTempElem
'                            strTempValue = Trim$(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCOMPANYTYPE"))
'                            If strTempValue = "01" Then
'                                strTempValue = "BureauDefaultUKBank"
'                            ElseIf strTempValue = "08" Then
'                                strTempValue = "BureauDefaultBuildingSociety"
'                            Else
'                                strTempValue = "BureauDefaultOtherLender"
'                            End If
'                            strResponse = objGlobalParamBO.GetCurrentParameter(strTempValue)
'                            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
'                            If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
'                                xmlTempElem.Text = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='" & strTempValue & "']/STRING")
'                            End If
                            
                            'Account Relationships
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIPLIST")
                            xmlMortgageAccountNode.appendChild xmlTempParentElem
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                Set xmlTempChildElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                xmlTempParentElem.appendChild xmlTempChildElem
                                m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                                m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            Next
                            
                            'Mortgage Account
                            'Set xmlAccount = xmlRequestDoc.createElement("MORTGAGEACCOUNT")
                            'xmlTempRequestElem.appendChild xmlAccount
                            
                            'Indemnity Insurance
                            Set xmlTempParentElem = xmlRequestDoc.createElement("INDEMNITYINSURANCE")
                            xmlMortgageAccountNode.appendChild xmlTempParentElem
                            
                            'BMIDS00919 MDC 12/11/2002 - End
                            
                            'Create Mortgage Account
                            strResponse = objCustFinancialBO.CreateMortgageAccount(xmlTempRequestElem.xml)
                            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                            m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
                            Debug.Print "Created Mortgage Account record"
                            blnExistingMortgage = True
                            
                            'Create a corresponding MortgageLoan record -------------------------------------
                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            Set xmlTempParentElem = xmlRequestDoc.createElement("MORTGAGELOAN")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            
                            xmlTempParentElem.appendChild m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, ".//ACCOUNTGUID")
                            m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlTempParentElem, False
                            m_objXmlAssist.CopyNode "FBCAISBALANCE", "OUTSTANDINGBALANCE", xmlCAISNode, xmlTempParentElem, False
                        
                            intTempValue = CInt(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISREPAYPERIOD"))
                            If intTempValue > 0 And IsDate(strStartDate) Then
                                strTempValue = DateAdd("m", intTempValue, strStartDate)
                            Else
                                strTempValue = DateAdd("yyyy", 5, Now)  'BM0158 MDC 05/12/2002 - correct year format
                            End If
                            'BM0158 MDC 05/12/2002 - Ensure date is formated as date only (i.e. no time element)
                            strTempValue = Format$(strTempValue, "dd/mm/yyyy")
                            'BM0158 MDC 05/12/2002 - End
                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "LOANENDDATE", strTempValue
                                
                            If IsDate(strStartDate) Then
                                m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "STARTDATE", strStartDate
                            End If
                            
                            'BMIDS00919 MDC 12/11/2002 - Default values for mandatory REDEMPTIONSTATUS field

                            
                            'GD 26/02/2003 BM0373 START -REMOVAL OF EXISTING CODE
'                            If IsDate(strTempValue) Then
'                                'BM0251 MDC 15/01/2003
'                                'If strTempValue > Now Then
'                                If CDate(strTempValue) > Now Then
'                                'BM0251 MDC 15/01/2003 - End
'                                    strTempValue = objComboDO.GetFirstComboValueId("RedemptionStatus", "R")
'                                Else
'                                    strTempValue = objComboDO.GetFirstComboValueId("RedemptionStatus", "A")
'                                End If
'                            Else
'                                strTempValue = objComboDO.GetFirstComboValueId("RedemptionStatus", "R")
'                            End If
'                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "REDEMPTIONSTATUS", strTempValue
                            'GD 26/02/2003 BM0373 END -REMOVAL OF EXISTING CODE
                            
                            'BMIDS00919 MDC 12/11/2002 - End
                            'GD 26/02/2003 BM0373 START
                           'Build xmlAppFFRequest
                            Set xmlAppFFRequest = xmlRequestDoc.documentElement.cloneNode(False)
                            Set xmlAppFFElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
'                            Set xmlAppFFNum = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "APPLICATION/APPLICATIONFACTFINDNUMBER")
'                            Set xmlAppNum = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, "APPLICATION/APPLICATIONNUMBER")
                            m_objXmlAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", xmlApplicationElement, _
                                                                        xmlAppFFElem, False
                            m_objXmlAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", _
                                                    xmlApplicationElement, xmlAppFFElem, False
                            
                            xmlAppFFRequest.appendChild xmlAppFFElem
'                            xmlAppFFElem.appendChild xmlAppFFNum
'                            xmlAppFFElem.appendChild xmlAppNum

                            
                            'Call omApp.ApplicationBO.GetApplicationFactFindData(ApplicationNumber,ApplicationfactFindNumber)
                            strAppFFResponse = objApplicationBO.GetApplicationFactFindData(xmlAppFFRequest.xml)
                            'Load Response into a FreeThreadedDOMDocument40
                            'Check if call has failed, and raise an error if it has
                            Set xmlGetAppFFDataResponseDoc = m_objXmlAssist.load(strAppFFResponse, TypeName(Me), cstrFunctionName)
                            m_objErrAssist.CheckResponse xmlGetAppFFDataResponseDoc.documentElement, xmlResponseElem, True
                            'ApplicationFactFind Data returned OK, if we reach here
                            'Extract strDirectIndirectBusiness
                            'strDirectIndirectBusiness = Get ApplicationfactFind.DirectIndirectBusiness
                            
                            'BMIDS633
                            strSettledDate = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISSETTLEDDATE", False, False)
                            
                            If Len(strSettledDate) > 0 Then
                                strResponse = objComboDO.GetComboValueId("RedemptionStatus", "A")
                                xmlComboResponseDoc.loadXML (strResponse)
                                Set xmlRedemptionElement = xmlComboResponseDoc.selectSingleNode(".//VALUEID")
                                If xmlRedemptionElement Is Nothing Then
                                    'RaiseError
                                    m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4NoComboTagValue, "ValueID for validation type 'A'  not found"
                                Else
                                    strRedemptionStatus = xmlRedemptionElement.Text
                                End If
                            
                            Else
                                'strDirectIndirectBusiness = xmlGetAppFFDataResponseDoc.documentElement.selectSingleNode(".//DIRECTINDIRECTBUSINESS").Text
                                Set xmlDirectIndirectBusiness = xmlGetAppFFDataResponseDoc.documentElement.selectSingleNode(".//DIRECTINDIRECTBUSINESS")
                                If Not (xmlDirectIndirectBusiness Is Nothing) Then
                                    strDirectIndirectBusiness = xmlDirectIndirectBusiness.Text
                                Else
                                    strDirectIndirectBusiness = vbNullString
                                End If
                                
                                
                                'if strDirectIndirectBusiness has a ValidationType of 7 (aXess) then
                                blnToBeRedeemed = False
                                If Len(strDirectIndirectBusiness) > 0 Then
                                    blnToBeRedeemed = objComboDO.IsItemInValidation("Direct/InDirect", strDirectIndirectBusiness, "7")
                                End If
                                
                                
                                If blnToBeRedeemed Then
    ''''                            'strRedemptionStatus = Value ID with Validation Type = 'R' and '7' (='To Be Redeemed - this advance')
                                    strSearch = "LIST/LISTNAME/LISTENTRY[(VALIDATIONTYPELIST/VALIDATIONTYPE='R') and ( VALIDATIONTYPELIST/VALIDATIONTYPE='7')]/VALUEID"
                                    strResponse = objComboDO.GetComboList("<LIST><LISTNAME>RedemptionStatus</LISTNAME></LIST>")
                                    xmlComboResponseDoc.loadXML (strResponse)
                                    Set xmlRedemptionElement = xmlComboResponseDoc.selectSingleNode(strSearch)
                                    If xmlRedemptionElement Is Nothing Then
                                        'RaiseError
                                        m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4NoComboTagValue, "ValueID for validation types '7' and 'R'  not found"
                                    Else
                                        strRedemptionStatus = xmlRedemptionElement.Text
                                    End If
                                Else
                                    'BMIDS633 Els 'Other Lender
    '                                If IsDate(strTempValue) Then
    '                                    If CDate(strTempValue) > Now Then
                                               'BMIDS617 START
    '                                           'strRemdemptionStatus = Value ID with Validation Type = N (='NOT To be Redeemed')
    '                                            strResponse = objComboDO.GetComboValueId("RedemptionStatus", "N")
                                               'strRemdemptionStatus = Value ID with Validation Type = 70 (='Other Lender')
                                                strResponse = objComboDO.GetComboValueId("RedemptionStatus", "70")
                                                'BMIDS617 END
                                                xmlComboResponseDoc.loadXML (strResponse)
                                                Set xmlRedemptionElement = xmlComboResponseDoc.selectSingleNode(".//VALUEID")
                                                If xmlRedemptionElement Is Nothing Then
                                                    'RaiseError
                                                    m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4NoComboTagValue, "ValueID for validation type '70'  not found"
                                                Else
                                                    strRedemptionStatus = xmlRedemptionElement.Text
                                                End If
    '                                    Else
    '                                            'strRedemptionStatus = Value ID with Validation Type = 'A' (='Already Redeemed')
    '                                            strResponse = objComboDO.GetComboValueId("RedemptionStatus", "A")
    '                                            xmlComboResponseDoc.loadXML (strResponse)
    '                                            Set xmlRedemptionElement = xmlComboResponseDoc.selectSingleNode(".//VALUEID")
    '                                            If xmlRedemptionElement Is Nothing Then
    '                                                'RaiseError
    '                                                m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4NoComboTagValue, "ValueID for validation type 'A'  not found"
    '                                            Else
    '                                                strRedemptionStatus = xmlRedemptionElement.Text
    '                                            End If
    '
    '                                    End If
    '                                End If
                                End If
                            End If
                            
                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "REDEMPTIONSTATUS", strRedemptionStatus
                            
                            
                            'GD 26/02/2003 BM0373 END
                            strResponse = objCustFinancialBO.CreateMortgageLoan(xmlTempRequestElem.xml)
                            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                            Debug.Print "Created Mortgage Loan record"
                        
                        End If
                        strTempValue = objComboDO.GetFirstComboValueId("ArrearsLoanType", "M")
                        'BM0168 MDC 06/12/2002
                        'BM0172 MDC 06/12/2002
                        If Not xmlAccount Is Nothing Then
                            m_objXmlAssist.CopyNode "ACCOUNTGUID", "ACCOUNTGUID", xmlAccount, xmlTempRequestElem, False
                        End If
                        'BM0172 MDC 06/12/2002 - End
                        'BM0168 MDC 06/12/2002 - End
                    Else
                        'Otherwise its a LoanLiability --------------------------------------------------
                        strRepayment = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISMONTHLYPAYMENT", False, False)
                        strBalance = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISBALANCE", False, False)
                        
                        'Try to find existing matching Omiga 4 record
                        
                        'BMIDS00972 MDC 18/11/2002 - Exclude monthly repayment from match when Credit Card as repayment is calculated later
'                        Set xmlAccount = m_objXmlAssist.GetNode(xmlLiabilityListNode, _
                            "LOANSLIABILITIES[TOTALOUTSTANDINGBALANCE = '" & strBalance & "'" _
                            & " and MONTHLYREPAYMENT = '" & Format$(strRepayment, "#####0.00") & "']", False)
                        strCondition = "LOANSLIABILITIES[TOTALOUTSTANDINGBALANCE = '" & strBalance & "'"
                        
                        strTempValue = objComboDO.GetFirstComboValueId("AgreementType", strAccType)
                        
                        'BMIDS01027 MDC 20/11/2002
                        If Len(Trim$(strTempValue)) > 0 Then
                            If Not objComboDO.IsItemInValidation("AgreementType", strTempValue, "CC") Then
                                strCondition = strCondition & " and MONTHLYREPAYMENT = '" & Format$(strRepayment, "#####0.00") & "'"
                            End If
                        End If
                        'BMIDS01027 MDC 20/11/2002 - End
                        
                        strCondition = strCondition & "]"
                        Set xmlAccount = m_objXmlAssist.GetNode(xmlLiabilityListNode, strCondition, False)
                        'BMIDS00972 MDC 18/11/2002 - End
                        
                        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                        
                        strTempValue = "0"
                        If xmlAccount Is Nothing Then
                            'Create a new LoansLiabilites record
                            blnCreateNew = True
                            
                            'LoansLiabilities
                            Set xmlAccount = xmlRequestDoc.createElement("LOANSLIABILITIES")
'                            m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlAccount, False
                            m_objXmlAssist.CopyNode "FBCAISBALANCE", "TOTALOUTSTANDINGBALANCE", xmlCAISNode, xmlAccount, False
                            If IsDate(strStartDate) Then
                                m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "STARTDATE", strStartDate
                            End If
                            
                            'BM0162 MDC 05/12/2002
                            intTempValue = CInt(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISREPAYPERIOD"))
                            If intTempValue > 0 And IsDate(strStartDate) Then
                                strTempDate = DateAdd("m", intTempValue, strStartDate)
                                'Ensure End Date is not in the past (fails CreateLiability validation)
                                If CDate(strTempDate) < Now Then
                                    strTempDate = CStr(Now)
                                End If
                            Else
                                strTempDate = DateAdd("yyyy", 5, Now)
                            End If
                            strTempDate = Format$(strTempDate, "dd/mm/yyyy")
                            m_objXmlAssist.SetMandatoryChildText xmlAccount, "ENDDATE", strTempDate
                            
                            m_objXmlAssist.SetMandatoryChildText xmlAccount, "LOANREPAYMENTINDICATOR", "0"
                            'BM0162 MDC 05/12/2002 - End
                            
                            'BM0158 MDC 05/12/2002
                            m_objXmlAssist.SetMandatoryChildText xmlAccount, "ADDITIONALINDICATOR", "0"
                            'BM0158 MDC 05/12/2002 - End
                            
                            'Account
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ACCOUNT")
                            Set xmlTempElem = xmlRequestDoc.createElement("DIRECTORYGUID")
                            'JD BMIDS897 No longer create a new thirdparty. Instead, find
                            ' an existing DirectoryGuid to attach to the Account
                            strTempValue = Trim$(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCOMPANYTYPE"))
                            xmlTempElem.Text = GetDirectoryGUIDForBureau(strTempValue)
                            xmlTempParentElem.appendChild xmlTempElem
                            Set xmlTempElem = xmlRequestDoc.createElement("THIRDPARTYGUID")
                            xmlTempParentElem.appendChild xmlTempElem
                            
                            xmlTempRequestElem.appendChild xmlAccount
                            xmlAccount.appendChild xmlTempParentElem
                            
                        Else
                            'Account
                            blnCreateNew = False    'BMIDS00970 MDC 16/11/2002
                            Set xmlTempParentElem = m_objXmlAssist.GetMandatoryNode(xmlAccount, "ACCOUNT")
                            If Len(Trim$(m_objXmlAssist.GetElementText(xmlTempParentElem, "DIRECTORYGUID"))) > 0 Then
                                strTempValue = "1"
                            End If
                        End If
                        
                        Set xmlTempElem = xmlRequestDoc.createElement("DIRECTORYADDRESSINDICATOR")
                        xmlTempElem.Text = strTempValue
                        xmlTempRequestElem.appendChild xmlTempElem
                        
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "UNASSIGNED", CStr(Abs(CInt(blnUnAssigned)))
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "CREDITSEARCH", "1"
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "IMPORTEDINDICATOR", "1"
                        'BMIDS00561 MDC 07/10/2002
                        strTempValue = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBBUREAUREFERENCE", False, False)
                        strTempValue = strTempValue & " (" _
                                & m_objXmlAssist.GetMandatoryElementText(xmlCAISNode, "FBHEADERSEQUENCE") _
                                & " of " & CStr(intTotalCAISRecords) & ")"
                        m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ACCOUNTNUMBER", Left$(strTempValue, 20)
                        'BMIDS00561 MDC 07/10/2002 - End
                            
                        xmlTempRequestElem.appendChild xmlAccount
                        
                        If Not blnCreateNew Then
                            'Get existing owners
                            Set xmlExistingOwnersList = xmlRequestDoc.createElement("OWNERS")
                            For Each xmlTempElem In xmlAccount.selectNodes("ACCOUNTRELATIONSHIP")
                                xmlExistingOwnersList.appendChild xmlTempElem
                            Next
                            
                            'Add in any new owners
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                strCustomerNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERNUMBER")
                                strCustomerVersionNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERVERSIONNUMBER")
                                Set xmlTempElem = m_objXmlAssist.GetNode(xmlExistingOwnersList, _
                                        "ACCOUNTRELATIONSHIP[CUSTOMERNUMBER='" & strCustomerNumber & "' and " _
                                        & "CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']", False)
                                If Not xmlTempElem Is Nothing Then
                                    xmlExistingOwnersList.removeChild xmlTempElem
                                Else
                                    'Add new owner
                                    Set xmlTempElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                    xmlAccount.appendChild xmlTempElem
                                    m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempElem, False
                                    m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempElem, False
                                End If
                            Next
                            
                            'Delete existing owners not now on this account
                            If xmlExistingOwnersList.selectNodes("ACCOUNTRELATIONSHIP").length > 0 Then
                                Set xmlTempChildElem = xmlRequestDoc.createElement("DELETE")
                                xmlAccount.appendChild xmlTempChildElem
                            End If
                            For Each xmlCustomerElem In xmlExistingOwnersList.selectNodes("ACCOUNTRELATIONSHIP")
                                xmlTempChildElem.appendChild xmlCustomerElem
                            Next
                            
                            'BM0168 MDC 06/12/2002 - Do not update End Date in case this is in the past (and will fail validation)
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempRequestElem, ".//ENDDATE", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.parentNode.removeChild xmlTempElem
                            End If
                            'BM0168 MDC 06/12/2002 - End
                            
                            'Update existing Loan liability
                            strResponse = objCustFinancialBO.UpdateLiability(xmlTempRequestElem.xml)
                            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                            Debug.Print "Updated LoansLiability record"
                            
                        Else
                            'Create new Loan liability
                            
                            'BMIDS00561 MDC 08/10/2002
                            Set xmlTempElem = xmlRequestDoc.createElement("AGREEMENTTYPE")
                            'BMIDS00946 MDC 14/11/2002
                            strTempValue = objComboDO.GetFirstComboValueId("AgreementType", strAccType)
                            If Len(Trim$(strTempValue)) = 0 Then
                                strTempValue = objComboDO.GetFirstComboValueId("AgreementType", "O")
                            End If
                            xmlTempElem.Text = strTempValue
'                            If strAccType = "05" Or strAccType = "06" Then
'                                'Credit Card
'                                xmlTempElem.Text = objComboDO.GetFirstComboValueId("AgreementType", "CC")
'                            Else
'                                'Other
'                                xmlTempElem.Text = objComboDO.GetFirstComboValueId("AgreementType", "O")
'                            End If
                            
                            xmlAccount.appendChild xmlTempElem
                            'BMIDS00946 MDC 14/11/2002 - End

                            'BMIDS00945 MDC 14/11/2002
                            'If Credit Card, calculate monthly repayment
                            If objComboDO.IsItemInValidation("AgreementType", strTempValue, "CC") Then
                                Set xmlTempElem = xmlRequestDoc.documentElement.cloneNode(False)
                                m_objXmlAssist.CopyNode "TOTALOUTSTANDINGBALANCE", "TOTALOUTSTANDINGBALANCE", xmlAccount, xmlTempElem, False
                                m_objXmlAssist.CopyNode "FBCAISBALANCE", "TOTALOUTSTANDINGBALANCE", xmlCAISNode, xmlTempElem, False
                                strResponse = objCustFinancialBO.CalculateCreditCardRepayment(xmlTempElem.xml)
                                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                                If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
                                    m_objXmlAssist.CopyNode "MONTHLYREPAYMENT", "MONTHLYREPAYMENT", xmlTempResponseDoc.documentElement, xmlAccount, False
                                End If
                            Else
                                m_objXmlAssist.CopyNode "FBCAISMONTHLYPAYMENT", "MONTHLYREPAYMENT", xmlCAISNode, xmlAccount, False
                            End If
                            'BMIDS00945 MDC 14/11/2002 - End
                            
                            'BMIDS00561 MDC 08/10/2002 - End
                            
                            
                            'Owners
                            For Each xmlCustomerElem In xmlCustomerNodeList
                                Set xmlTempElem = xmlRequestDoc.createElement("ACCOUNTRELATIONSHIP")
                                xmlAccount.appendChild xmlTempElem
                                m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempElem, False
                                m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempElem, False
                            Next
                            'JD BMIDS897 Don't create a new ThirdParty
'                            'Third Party
'                            Set xmlTempChildElem = xmlRequestDoc.createElement("THIRDPARTY")
'                            xmlAccount.appendChild xmlTempChildElem
'
'                            Set xmlTempElem = xmlRequestDoc.createElement("THIRDPARTYTYPE")
'                            '#TASK - Need unique validation type for Bank/Building Soc/Other Lender
'                            'xmlTempElem.Text = objComboDO.GetFirstComboValueId("ThirdPartyType", "T")
'                            xmlTempElem.Text = "3"
'                            xmlTempChildElem.appendChild xmlTempElem
'
'                            Set xmlTempElem = xmlRequestDoc.createElement("COMPANYNAME")
'                            xmlTempChildElem.appendChild xmlTempElem
'                            strTempValue = Trim$(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISCOMPANYTYPE"))
'                            If strTempValue = "01" Then
'                                strTempValue = "BureauDefaultUKBank"
'                            ElseIf strTempValue = "08" Then
'                                strTempValue = "BureauDefaultBuildingSociety"
'                            Else
'                                strTempValue = "BureauDefaultOtherLender"
'                            End If
'                            strResponse = objGlobalParamBO.GetCurrentParameter(strTempValue)
'                            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
'                            If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
'                                xmlTempElem.Text = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='" & strTempValue & "']/STRING")
'                            End If
                            'GD BM0479 START
                            'Get FBCAISACCOUNTTYPE
                            strTempValue = Trim$(m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTTYPE"))
                            'ENDDATE already set to 5 years in the future
                            If Not (strTempValue = "05" Or strTempValue = "06") Then 'NOT Credit Card or Charge Card
                                'Get ENDDATE Element
                                strTempValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISREPAYPERIOD")
                                'Handle the sutuation where FBCAISREPAYPERIOD = empty string
                                If Len(strTempValue) = 0 Then
                                    strTempValue = "0"
                                End If
                                intTempValue = CInt(strTempValue)
                                If intTempValue > 0 Then
                                    'strStartDate = m_objXmlAssist.GetNodeValue(xmlCAISNode, "FBCAISSTARTDATE", False, False)
                                    strCaisStartDate = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISSTARTDATE")
                                    'strCAISRepayPeriod = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISREPAYPERIOD")
                                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempRequestElem, ".//ENDDATE", False)
                                    If IsDate(strCaisStartDate) Then
                                        strTempDate = DateAdd("m", intTempValue, strCaisStartDate)
                                    End If
                                    If xmlTempElem Is Nothing Then 'No ENDDATE element exists
                                        'Create One.
                                        Set xmlTempElem = xmlRequestDoc.createElement("ENDDATE")
                                        xmlAccount.appendChild xmlTempElem
                                    End If
                                    strTempDate = Format$(strTempDate, "dd/mm/yyyy")
                                    'If ENDDATE is in the past, the set it to ""
                                    If CDate(strTempDate) < Now Then
                                        strTempDate = vbNullString
                                    End If
                                    xmlTempElem.Text = strTempDate
                                End If
                            End If
                            'GD BM0479 END
                            'Create LoansLiabilities
                            strResponse = objCustFinancialBO.CreateLiability(xmlTempRequestElem.xml)
                            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                            Debug.Print "Created Liability record"
                            blnLoanLiability = True
                        
                        End If
                        strTempValue = objComboDO.GetFirstComboValueId("ArrearsLoanType", "L")
                        
                    End If
                
                    'Is there an Arrears History? -------------------------------------------------------
                    
                    'BMIDS595 & 596  START
                    'Change code below to look at the whole 12 character string.
                    'Create arrears history if any of the characters in the string are > 0.
            
'                    'GD BM0247 24/02/2003 START
'                    strValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTSTATUS")
'                    strFullCAISAccountStatus = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTSTATUS")
'                    'Get First Character of strValue
'                    strValue = Mid$(strValue, 1, 1)
'                    'GD BM0247 ORIGINAL strValue = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISWORSTSTATUS")
'                    'GD BM0247 24/02/2003 END
'
'                    If IsNumeric(strValue) Then
'                       If CInt(strValue) > 0 Then

                    strFullCAISAccountStatus = m_objXmlAssist.GetElementText(xmlCAISNode, "FBCAISACCOUNTSTATUS")

                    'Initialise intMaxAccountStatus
                    intMaxAccountStatus = 0
                    'Length of string
                    intFullCAISAccountStatus = Len(strFullCAISAccountStatus)
                            
                    For intIndex = 1 To intFullCAISAccountStatus
                        strTempCAISStatus = Mid$(strFullCAISAccountStatus, intIndex, 1)
                        If IsNumeric(strTempCAISStatus) Then
                            intTempCAISStatus = CInt(strTempCAISStatus)
                            If intTempCAISStatus > intMaxAccountStatus Then
                                intMaxAccountStatus = intTempCAISStatus
                            End If
                        End If
                    Next
                                   
                    If intMaxAccountStatus > 0 Then
                    'BMIDS595 & 596  END
                            'Create Arrears History record for this Account
                            
                            'BMIDS00561 MDC 08/10/2002
                            Set xmlAccountGuid = m_objXmlAssist.GetNode(xmlTempRequestElem, ".//ACCOUNTGUID", False)
                            'BMIDS00561 MDC 08/10/2002 - End

                            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                            Set xmlTempParentElem = xmlRequestDoc.createElement("ARREARSHISTORY")
                            xmlTempRequestElem.appendChild xmlTempParentElem
                            Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                            xmlTempParentElem.appendChild xmlTempElem
                            xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                            Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                            xmlTempParentElem.appendChild xmlTempElem
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                            Set xmlTempElem = xmlRequestDoc.createElement("REPOSSESSIONIND")
                            xmlTempElem.Text = "0"
                            
                            'BM0158 MDC 05/12/2002
                            m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ADDITIONALINDICATOR", "0"
                            'BM0158 MDC 05/12/2002 - End
                            
                            'BMIDS00561 MDC 08/10/2002
                            'Get Account Guid to link this record to
                            If xmlAccountGuid Is Nothing Then
                                Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
                                Set xmlAccountGuid = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc.documentElement, ".//GENERATEDKEYS/ACCOUNTGUID")
                            End If
                            xmlTempParentElem.appendChild xmlAccountGuid
                            'm_objXmlAssist.CopyNode "ACCOUNTGUID", "ACCOUNTGUID", xmlAccount, xmlTempParentElem, False
                            'BMIDS00561 MDC 08/10/2002 - End
                            
                            
                            'GD BM0247 24/02/2003 START
                            'Create a MAXIMUMNUMBEROFMONTHS Element
                            Set xmlTempElem = xmlRequestDoc.createElement("MAXIMUMNUMBEROFMONTHS")
                            
                            'Assign largest value in  FBCAISACCOUNTSTATUS to the MAXIMUMNUMBEROFMONTHS
                            xmlTempElem.Text = CStr(intMaxAccountStatus)
                            xmlTempParentElem.appendChild xmlTempElem
                            'ORIGINAL m_objXmlAssist.CopyNode "FBCAISWORSTSTATUS", "MAXIMUMNUMBEROFMONTHS", xmlCAISNode, xmlTempParentElem, False
                            'GD BM0247 24/02/2003 END
                            
                            m_objXmlAssist.CopyNode "FBCAISACCOUNTSTATUS", "CURRENTYEARSINARREARS", xmlCAISNode, xmlTempParentElem, False
                            'DESCRIPTIONOFLOAN is mandatory
                            Set xmlTempElem = xmlRequestDoc.createElement("DESCRIPTIONOFLOAN")
                            xmlTempParentElem.appendChild xmlTempElem
                            xmlTempElem.Text = strTempValue
                            strResponse = objCustFinancialBO.CreateArrearsHistory(xmlTempRequestElem.xml)
                            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                            Debug.Print "Created Arrears History record"
                            blnArrearsHistory = True
                        
                    End If
                End If
            Next
            
           End If 'If xmlNewCAISAccounts.hasChildNodes Then
        End If 'If Not xmlNewCAISAccounts Is Nothing Then
    
        '=======================================================================================
        ' Step D - CCJ/Bankruptcy data
        '=======================================================================================
        
        blnCreateNew = False    'BMIDS00946 MDC 14/11/2002
        
        'BMIDS630 Do we have any type 46's that might have an associated non type 46 from Experian?
        'If there is we need to get the DATEDECLARED from the non discharged entry. Add this
        'to the discharged entry and remove the non discharged entry. All of the required info
        'is now in the discharged entry and we then only need to compare against the database
        'to decide whether it will be an update or create.
        Set xmlType46NodeList = xmlBureauData.selectNodes(".//FULLBUREAUPUBLICINFOLIST/" _
                    & "FULLBUREAUPUBLICINFO[FBPUBLICINFODATATYPE = '46']")
        
        If xmlType46NodeList.length > 0 Then
        
            bNonDischargedToRemove = False
            
            strNonDischRemovalPattern = ".//FULLBUREAUPUBLICINFOLIST/" _
                & "FULLBUREAUPUBLICINFO[(FBPUBLICINFODATATYPE != '46' and "

            'Are there any
            For intType46Index = 0 To xmlType46NodeList.length - 1
                'Get the courtplaintiff to compare
                strCourtPlaintiff = m_objXmlAssist.GetElementText(xmlType46NodeList.Item(intType46Index), "FBCOURTPLAINTIFF")
                Set xmlTempNonType46Elem = m_objXmlAssist.GetNode(xmlBureauData, ".//FULLBUREAUPUBLICINFO" _
                                        & "[FBCOURTPLAINTIFF = '" & strCourtPlaintiff & "' and FBPUBLICINFODATATYPE != '46']", False)
                
                
                If Not xmlTempNonType46Elem Is Nothing Then
                
                    'We have a non type 46, which matches our type 46 the
                    'FBPUBLICINFOREGISTRATIONDATE from the non type 46 is the Date Declared
                    strDateDeclared = m_objXmlAssist.GetElementText(xmlTempNonType46Elem, "FBPUBLICINFOREGISTRATIONDATE")
                    'Add the Declared Date Back to the Type 46 record
                    Set xmlTempType46Elem = m_objXmlAssist.GetNode(xmlBureauData, ".//FULLBUREAUPUBLICINFO" _
                                        & "[FBCOURTPLAINTIFF = '" & strCourtPlaintiff & "' and FBPUBLICINFODATATYPE = '46']", False)
                    
                    Set xmlTempElem = xmlRequestDoc.createElement("DATEDECLARED")
                    xmlTempElem.Text = strDateDeclared
                    xmlTempType46Elem.appendChild xmlTempElem

                    'Get the info required to ensure that we remove the correct record
                    strnon46CreditCheckGUID = m_objXmlAssist.GetElementText(xmlTempNonType46Elem, "CREDITCHECKGUID")
                    strnon46FBBlockID = m_objXmlAssist.GetElementText(xmlTempNonType46Elem, "FBBLOCKID")
                    strnon46FBHeaderSequence = m_objXmlAssist.GetElementText(xmlTempNonType46Elem, "FBHEADERSEQUENCE")
                    strnon46FBPublicInfoSequence = m_objXmlAssist.GetElementText(xmlTempNonType46Elem, "FBPUBLICINFOSEQUENCE")

                    If bNonDischargedToRemove = False Then
                        'First one to find
                        strNonDischRemovalPattern = strNonDischRemovalPattern & "((CREDITCHECKGUID = '" & strnon46CreditCheckGUID & "') and " _
                        & "(FBBLOCKID = '" & strnon46FBBlockID & "') and " _
                        & "(FBHEADERSEQUENCE = '" & strnon46FBHeaderSequence & "') and " _
                        & "(FBPUBLICINFOSEQUENCE = '" & strnon46FBPublicInfoSequence & "') )"
                        'set so we don't come in here again
                        bNonDischargedToRemove = True
                        
                    Else
                        strNonDischRemovalPattern = strNonDischRemovalPattern & " or ((CREDITCHECKGUID = '" & strnon46CreditCheckGUID & "') and " _
                        & "(FBBLOCKID = '" & strnon46FBBlockID & "') and " _
                        & "(FBHEADERSEQUENCE = '" & strnon46FBHeaderSequence & "') and " _
                        & "(FBPUBLICINFOSEQUENCE = '" & strnon46FBPublicInfoSequence & "') )"
                    End If

                End If
            Next
            
            strNonDischRemovalPattern = strNonDischRemovalPattern & ")]"
                
            If bNonDischargedToRemove Then
                Set xmlNonDischargedToRemoveList = xmlBureauData.selectNodes(strNonDischRemovalPattern)
            End If
                
            If Not (xmlNonDischargedToRemoveList Is Nothing) Then
                If xmlNonDischargedToRemoveList.length > 0 Then
                    Set xmlFBPublicInfoList = xmlBureauData.selectSingleNode("FULLBUREAUPUBLICINFOLIST")
                    For int46RemoveIndex = 0 To (xmlNonDischargedToRemoveList.length - 1)
                        xmlFBPublicInfoList.removeChild xmlNonDischargedToRemoveList.Item(int46RemoveIndex)
                    Next
                End If
            End If
            
        End If 'End if xmlType46NodeList.length > 0 then

        'Get each Bureau PublicInfo record
        'BMIDS630 Only want to include in xmlNewRecordsNodeList, the records which are
        'matched to one of our customers. i.e. CUSTOMERLIST/CUSTOMER is attached to
        '<FULLBUREAUPUBLICINFOLIST><FULLBUREAUPUBLICINFO>
        Dim strCustBasedPatternMatch As String
        Dim strCustComponentOfPatternMatch As String
        Dim strCustomerNum As String
        Dim xmlCustNodeList As IXMLDOMNodeList
        Dim strXslPath As String
        Dim blnLoadedOk As Boolean
      
        Set xmlCustNodeList = xmlRequestDoc.documentElement.selectNodes(".//CUSTOMERLIST/CUSTOMER/CUSTOMERNUMBER")
    
        If xmlCustNodeList.length > 0 Then
            Dim nCustCount As Integer
            
            strCustBasedPatternMatch = ".//FULLBUREAUPUBLICINFOLIST/" _
                    & "FULLBUREAUPUBLICINFO[((FBPUBLICINFODATATYPE != '56') and " _
                    & "(FBPUBLICINFODATATYPE != '57')"

            strCustComponentOfPatternMatch = " and ("
            For nCustCount = 0 To (xmlCustNodeList.length - 1)
                strCustomerNum = xmlCustNodeList.Item(nCustCount).Text
                If nCustCount < (xmlCustNodeList.length - 1) Then
                    strCustComponentOfPatternMatch = strCustComponentOfPatternMatch & _
                        "(CUSTOMERLIST/CUSTOMER/CUSTOMERNUMBER = '" & strCustomerNum & "') or "
                Else
                    strCustComponentOfPatternMatch = strCustComponentOfPatternMatch & _
                        "(CUSTOMERLIST/CUSTOMER/CUSTOMERNUMBER = '" & strCustomerNum & "'))"
                End If
            Next
    
            strCustBasedPatternMatch = strCustBasedPatternMatch & strCustComponentOfPatternMatch & ")]"
            
        End If
                
        Set xmlNewRecordsNodeList = xmlBureauData.selectNodes(strCustBasedPatternMatch)
        
        'If new records to be imported...
        If xmlNewRecordsNodeList.length > 0 Then
            
            'Get existing BankruptcyHistory
            Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
            Set xmlTempParentElem = xmlRequestDoc.createElement("BANKRUPTCYHISTORYLIST")
            'BMIDS00525 MDC 26/09/2002
            xmlTempRequestElem.appendChild xmlTempParentElem
            'BMIDS00525 MDC 26/09/2002 - End
            Set xmlNodeList = xmlRequestDoc.documentElement.selectNodes(".//CUSTOMERLIST/CUSTOMER")
            For Each xmlNode In xmlNodeList
                m_objXmlAssist.CopyNode "CUSTOMER", "BANKRUPTCYHISTORY", xmlNode, xmlTempParentElem, True
            Next
            strResponse = objCustFinancialBO.FindBankruptcyHistoryList(xmlTempRequestElem.xml)
            
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
                'BMIDS00525 MDC 26/09/2002 - Check that BANKRUPTCYHISTORY have been returned
                'as FindBankruptcyHistoryList will not return 'Record Not Found'
                Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempResponseDoc.documentElement, _
                                        "BANKRUPTCYHISTORYLIST/BANKRUPTCYHISTORY", False)
                If Not xmlTempElem Is Nothing Then
                    'BankruptcyHistory records found - sort by declared date
                    Set xslDoc = New FreeThreadedDOMDocument40
                    xslDoc.validateOnParse = False
                    xslDoc.setProperty "NewParser", True
                    xslDoc.async = False
                    Set xmlTransformedDoc = New FreeThreadedDOMDocument40 'BMIDS00801 MDC 09/11/2002
                    xmlTransformedDoc.validateOnParse = False
                    xmlTransformedDoc.setProperty "NewParser", True
                    
                    strXslPath = App.Path & "\SortBankruptcyHistory.xslt"
                    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
                    blnLoadedOk = xslDoc.load(strXslPath)
                    
                    If Not blnLoadedOk Then
                        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
                    End If
                    
                    xmlTempResponseDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc 'xmlTempResponseDoc
                    
                    'BMIDS00801 MDC 09/11/2002
                    'Set xmlBankruptcyListNode = m_objXmlAssist.GetMandatoryNode(xmlTransformedDoc.documentElement, _
                                                                                            ".//BANKRUPTCYHISTORYLIST")
                    If Not xmlTransformedDoc Is Nothing Then
                        Set xmlBankruptcyListNode = xmlTransformedDoc.documentElement
                    Else
                        Set xmlBankruptcyListNode = Nothing
                    End If
                    'BMIDS00801 MDC 09/11/2002 - End
                End If
                'BMIDS00525 MDC 26/09/2002 - End
                
            End If
            '---------------------------  GD BMIDS610 FURTHER FIXES START  ---------------------------
            Dim xmlNewRecordsByDataType As IXMLDOMNodeList
            Dim strMostRecentDate As String
            Dim intDataTypeIndex As Integer
            Dim strPublicInfoRegDate As String
            Dim xmlPublicInfoRegDateElem As IXMLDOMElement
            'BMIDS630
            Dim xmlPublicInfoDataTypeElem As IXMLDOMElement
            Dim xmlPublicInfoCustomerElem As IXMLDOMElement
            Dim strCurCustomer As String
            Dim xmlTempCVBHistoryElem As IXMLDOMElement
            Dim strPatternMatch As String
            Dim xmlTempBHistoryElem As IXMLDOMElement
            ''Search for full bureau records with these FBPUBLICINFODATATYPEs (below) and find the one with the Max FBPUBLICINFOREGISTRATIONDATE
            '' 24, 25, 26, 27, 28, 30, 32, 33, 34, 35, 38, 41, 43, 44, 45, 47, 48, 49, 50, 51, 52, 53, 54
            strMostRecentDate = "01/01/1900" ''**NEW**
            'BMIDS630
            strDeclaredMostRecentNonDisch = "01/01/1900"
            ''**
            strPatternMatch = ".//FULLBUREAUPUBLICINFOLIST/" _
            & "FULLBUREAUPUBLICINFO[(FBPUBLICINFODATATYPE != '56' and " _
            & "FBPUBLICINFODATATYPE != '57') " _
            & " and ((FBPUBLICINFODATATYPE = '24') " _
            & "or (FBPUBLICINFODATATYPE = '25') or (FBPUBLICINFODATATYPE = '26') or (FBPUBLICINFODATATYPE = '27') " _
            & "or (FBPUBLICINFODATATYPE = '28') or (FBPUBLICINFODATATYPE = '30') or (FBPUBLICINFODATATYPE = '32') " _
            & "or (FBPUBLICINFODATATYPE = '33') or (FBPUBLICINFODATATYPE = '34') or (FBPUBLICINFODATATYPE = '35') " _
            & "or (FBPUBLICINFODATATYPE = '38') or (FBPUBLICINFODATATYPE = '41') or (FBPUBLICINFODATATYPE = '43') " _
            & "or (FBPUBLICINFODATATYPE = '44') or (FBPUBLICINFODATATYPE = '45') or (FBPUBLICINFODATATYPE = '47') " _
            & "or (FBPUBLICINFODATATYPE = '48') or (FBPUBLICINFODATATYPE = '49') or (FBPUBLICINFODATATYPE = '50') " _
            & "or (FBPUBLICINFODATATYPE = '51') or (FBPUBLICINFODATATYPE = '52') or (FBPUBLICINFODATATYPE = '53') or (FBPUBLICINFODATATYPE = '54') " _
            & " )"  ''**NEW**
            
            'BMIDS630 INR Make sure we only take into account the assigned records (i.e. they
            'have a customer attached), otherwise we might end up with an unassigned record
'            strPatternMatch = strPatternMatch & strCustComponentOfPatternMatch & "]"
            strPatternMatch = strPatternMatch & strCustComponentOfPatternMatch & "]"

            Dim xmlBanktuptcyToCreate As IXMLDOMElement ''** NEW **
            Dim strRemovalPattern As String ''** NEW **
            'BMIDS630
            'Dim strCustRemovalPattern As String ''** NEW **
            Dim bRemoved As Boolean
            
            Set xmlBanktuptcyToCreateList = xmlBHistDoc.createElement("BANKRUPTCYTOCREATELIST")
            xmlBHistDoc.appendChild xmlBanktuptcyToCreateList
            Set xmlBanktuptcyToCreateList = xmlBHistDoc.selectSingleNode(".//BANKRUPTCYTOCREATELIST")
                    
            Set xmlNewRecordsByDataType = xmlBureauData.selectNodes(strPatternMatch)
            For intDataTypeIndex = 0 To xmlNewRecordsByDataType.length - 1
        
                Set xmlPublicInfoRegDateElem = xmlNewRecordsByDataType.Item(intDataTypeIndex).selectSingleNode(".//FBPUBLICINFOREGISTRATIONDATE")
                Set xmlPublicInfoCustomerElem = xmlNewRecordsByDataType.Item(intDataTypeIndex).selectSingleNode(".//CUSTOMERNUMBER")

                'BMIDS661  Get Data Type here
                intDataType = 0
                Set xmlPublicInfoDataTypeElem = xmlNewRecordsByDataType.Item(intDataTypeIndex).selectSingleNode(".//FBPUBLICINFODATATYPE")
                If Not (xmlPublicInfoDataTypeElem Is Nothing) Then
                    If Len(xmlPublicInfoDataTypeElem.Text) > 0 Then
                        intDataType = CInt(xmlPublicInfoDataTypeElem.Text)
                    End If
                End If

                'if the Customer is nothing we shouldn't be using it anyway as it is unmatched
                If ((Not (xmlPublicInfoRegDateElem Is Nothing)) And (Not (xmlPublicInfoCustomerElem Is Nothing))) Then
                    strPublicInfoRegDate = xmlPublicInfoRegDateElem.Text
                    strCurCustomer = xmlPublicInfoCustomerElem.Text
                    If IsDate(strPublicInfoRegDate) Then
                        
                        strPatternMatch = ".//FULLBUREAUPUBLICINFO[CUSTOMERLIST/CUSTOMER/CUSTOMERNUMBER = '" & strCurCustomer & "']"
                        Set xmlBanktuptcyToCreate = xmlBanktuptcyToCreateList.selectSingleNode(strPatternMatch)
                        If Not xmlBanktuptcyToCreate Is Nothing Then
                            'we have matched an existing bankruptcy record
                            
                                Set xmlTempElem = xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDATE")
                                If Not xmlTempElem Is Nothing Then
                                
                                    bRemoved = False
                                
                                    strMostRecentDate = xmlTempElem.Text
                                    Set xmlTempElem = xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDECLAREDDATE")
                                    If Not xmlTempElem Is Nothing Then
                                        strDeclaredMostRecentNonDisch = xmlTempElem.Text
                                    End If
                                    
                                    'BMIDS661
                                    'If the earlier record and new record have the same date, we need to keep any discharged bankruptcy.
                                    If CDate(strPublicInfoRegDate) = CDate(strMostRecentDate) Then
                                     
                                        'If this record is discharged bankruptcy, remove earlier record
                                        If ((intDataType = 26) Or (intDataType = 27) Or (intDataType = 28) _
                                            Or (intDataType = 41) Or (intDataType = 52) Or (intDataType = 53)) Then
                                            bRemoved = True
                                        End If
                                        
                                    ElseIf CDate(strPublicInfoRegDate) > CDate(strMostRecentDate) Then
                                        bRemoved = True
                                                                                
                                    End If
                                    
                                    If (bRemoved = True) Then
 
                                        'BMIDS630 INR Need our list on a per client basis
                                        'If we have a later one which has the same client, remove the
                                        'earlier record and append the new one.
                                
                                        xmlBanktuptcyToCreateList.removeChild xmlBanktuptcyToCreate
                                        'bRemoved = True
                                        Set xmlBanktuptcyToCreate = xmlNewRecordsByDataType.Item(intDataTypeIndex) ''** NEW **
                                        
                                        If xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDATE") Is Nothing Then
                                        
                                            Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDATE")
                                            xmlTempElem.Text = strPublicInfoRegDate
                                            xmlBanktuptcyToCreate.appendChild xmlTempElem
                                        Else
                                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlBanktuptcyToCreate, "MOSTRECENTDATE", False)
                                            If Not xmlTempElem Is Nothing Then
                                                xmlTempElem.Text = strPublicInfoRegDate
                                            End If
                                        
                                        End If
                                        
                                        'Need to keep any existing MOSTRECENTDECLAREDDATE with the current client record
                                        If xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDECLAREDDATE") Is Nothing Then
                                        
                                            Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDECLAREDDATE")
                                            xmlTempElem.Text = strDeclaredMostRecentNonDisch
                                            xmlBanktuptcyToCreate.appendChild xmlTempElem
                                        Else
                                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlBanktuptcyToCreate, "MOSTRECENTDECLAREDDATE", False)
                                            If Not xmlTempElem Is Nothing Then
                                                xmlTempElem.Text = strDeclaredMostRecentNonDisch
                                            End If
                                        
                                        End If
                                        
                                    End If
                                   
                                    'BMIDS630 Need datatype to work out most recent non discharged declared date
                                    Set xmlPublicInfoDataTypeElem = xmlNewRecordsByDataType.Item(intDataTypeIndex).selectSingleNode(".//FBPUBLICINFODATATYPE")
                                    If Not (xmlPublicInfoDataTypeElem Is Nothing) Then
                                        If Len(xmlPublicInfoDataTypeElem.Text) > 0 Then
                                            intDataType = CInt(xmlPublicInfoDataTypeElem.Text)
                                            If ((intDataType = 24) Or (intDataType = 25) Or _
                                            (intDataType = 30) Or (intDataType = 32) Or (intDataType = 33) Or _
                                            (intDataType = 34) Or (intDataType = 35) Or (intDataType = 38) Or _
                                            (intDataType = 43) Or (intDataType = 44) Or (intDataType = 45) Or _
                                            (intDataType = 47) Or (intDataType = 48) Or (intDataType = 49) Or _
                                            (intDataType = 50) Or (intDataType = 51) Or (intDataType = 54)) Then
                                    
                                                'These Data Types indicate that strPublicInfoRegDate is a Declared Date
                                                If CDate(strPublicInfoRegDate) > CDate(strDeclaredMostRecentNonDisch) Then
                                                    
                                                    If xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDECLAREDDATE") Is Nothing Then
                                                    
                                                        Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDECLAREDDATE")
                                                        xmlTempElem.Text = strPublicInfoRegDate
                                                        xmlBanktuptcyToCreate.appendChild xmlTempElem
                                                    Else
                                                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlBanktuptcyToCreate, "MOSTRECENTDECLAREDDATE", False)
                                                        If Not xmlTempElem Is Nothing Then
                                                            xmlTempElem.Text = strPublicInfoRegDate
                                                        End If
                                                    
                                                    End If
                                                
                                                End If
                                            End If
                                        End If
                                        If bRemoved Then
                                            'Add the recreated record
                                            xmlBanktuptcyToCreateList.appendChild xmlBanktuptcyToCreate.cloneNode(True)
                                        
                                        End If

                                    End If
                                   
                               
                            End If
                            
                        Else
                            'Else this is a new bankruptcy record
                            Set xmlBanktuptcyToCreate = xmlNewRecordsByDataType.Item(intDataTypeIndex)

                            'New record so no need to do the compare
                            If CDate(strPublicInfoRegDate) Then
                                strMostRecentDate = strPublicInfoRegDate
                                'BMIDS630 INR Need our list on a per client basis append the new one.
                                If xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDATE") Is Nothing Then
                                    Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDATE")
                                    xmlTempElem.Text = strMostRecentDate
                                    xmlBanktuptcyToCreate.appendChild xmlTempElem
                                Else
                                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlBanktuptcyToCreate, "MOSTRECENTDATE", False)
                                    If Not xmlTempElem Is Nothing Then
                                        xmlTempElem.Text = strMostRecentDate
                                    End If
                                
                                End If
                                
                            End If
                            
                            'BMIDS630 Need datatype to work out most recent non discharged declared date
                            Set xmlPublicInfoDataTypeElem = xmlNewRecordsByDataType.Item(intDataTypeIndex).selectSingleNode(".//FBPUBLICINFODATATYPE")
                            If Not (xmlPublicInfoDataTypeElem Is Nothing) Then
                                If Len(xmlPublicInfoDataTypeElem.Text) > 0 Then
                                    intDataType = CInt(xmlPublicInfoDataTypeElem.Text)
                                    If ((intDataType = 24) Or (intDataType = 25) Or _
                                    (intDataType = 30) Or (intDataType = 32) Or (intDataType = 33) Or _
                                    (intDataType = 34) Or (intDataType = 35) Or (intDataType = 38) Or _
                                    (intDataType = 43) Or (intDataType = 44) Or (intDataType = 45) Or _
                                    (intDataType = 47) Or (intDataType = 48) Or (intDataType = 49) Or _
                                    (intDataType = 50) Or (intDataType = 51) Or (intDataType = 54)) Then
                            
                                        'These Data Types indicate that strPublicInfoRegDate is a Declared Date
                                        If CDate(strPublicInfoRegDate) Then
                                            
                                            If xmlBanktuptcyToCreate.selectSingleNode("MOSTRECENTDECLAREDDATE") Is Nothing Then
                                            
                                                Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDECLAREDDATE")
                                                xmlTempElem.Text = strPublicInfoRegDate
                                                xmlBanktuptcyToCreate.appendChild xmlTempElem
                                            Else
                                                Set xmlTempElem = m_objXmlAssist.GetNode(xmlBanktuptcyToCreate, "MOSTRECENTDECLAREDDATE", False)
                                                If Not xmlTempElem Is Nothing Then
                                                    xmlTempElem.Text = strPublicInfoRegDate
                                                End If
                                            
                                            End If
                                        
                                        End If
                                    End If
                                End If
                            End If
                            
                            'Add the newly created record
                            xmlBanktuptcyToCreateList.appendChild xmlBanktuptcyToCreate.cloneNode(True)
                            
                        End If  'If Not xmlBanktuptcyToCreate Is Nothing
                        
                    End If 'If IsDate(strPublicInfoRegDate)
                    
                
'                xmlBanktuptcyToCreateList.appendChild xmlBanktuptcyToCreate.cloneNode(True)

                
                End If 'If ((Not (xmlPublicInfoRegDateElem Is Nothing))
                
               
            Next
            
'BMIDS630 This is now dealt with further down
'            If strMostRecentDate = "01/01/1900" Then
'                strMostRecentDate = vbNullString
'            End If
'            'BMIDS630
'            If strDeclaredMostRecentNonDisch = "01/01/1900" Then
'                strDeclaredMostRecentNonDisch = vbNullString
'            End If
            
            ''BMIDS610 Further fixes to remove duplicate BankRuptcyHistory Records START
            ''If there are some Public Info records RECORDS''**NEW **
            Dim strCreateCreditCheckGUID As String ''**NEW **
            Dim strCreateFBBlockID As String ''**NEW **
            Dim strCreateFBHeaderSequence As String ''**NEW **
            Dim strCreateFBPublicInfoSequence As String ''**NEW **
            Dim xmlListToRemove As IXMLDOMNodeList ''**NEW **
            Dim intRemoveIndex As Integer
            'BMIDS630 declared & used further up
'            Dim xmlFBPublicInfoList As IXMLDOMElement ''**NEW**
            Dim strGUID As String
            Dim strID  As String
            Dim strHeaderSequence  As String
            Dim strPublicInfoSequence  As String
            Dim xmlBankruptcyList As IXMLDOMNodeList
            'Dim strMostRecent As String
            'Dim strMostDischRecent As String
            Dim xmlMatchNode As IXMLDOMElement
            Dim xmlBHMatchNode As IXMLDOMElement
            
            Set xmlFBPublicInfoList = Nothing
           
            'BMIDS630 Could have more than one record if we have multiple clients
            Set xmlBankruptcyList = xmlBanktuptcyToCreateList.selectNodes(".//FULLBUREAUPUBLICINFO")
            
            For intDataTypeIndex = 0 To xmlBankruptcyList.length - 1
                    strCreateCreditCheckGUID = xmlBankruptcyList.Item(intDataTypeIndex).selectSingleNode(".//CREDITCHECKGUID").Text ''**NEW **
                    strCreateFBBlockID = xmlBankruptcyList.Item(intDataTypeIndex).selectSingleNode(".//FBBLOCKID").Text ''**NEW **
                    strCreateFBHeaderSequence = xmlBankruptcyList.Item(intDataTypeIndex).selectSingleNode(".//FBHEADERSEQUENCE").Text ''**NEW **
                    strCreateFBPublicInfoSequence = xmlBankruptcyList.Item(intDataTypeIndex).selectSingleNode(".//FBPUBLICINFOSEQUENCE").Text ''**NEW **
                        

                    Dim strBHMatchMostRecentDec As String
                    Dim strBureauMatchMostRecentDec As String

                    Set xmlBHMatchNode = xmlBankruptcyList.Item(intDataTypeIndex).selectSingleNode(".//MOSTRECENTDECLAREDDATE")
                    If Not xmlBHMatchNode Is Nothing Then
                        If Len(xmlBHMatchNode.Text) > 0 Then
                            strBHMatchMostRecentDec = xmlBHMatchNode.Text
                        End If
                    End If


                    Set xmlMatchNode = xmlBureauData.selectSingleNode(".//FULLBUREAUPUBLICINFOLIST/" _
                        & "FULLBUREAUPUBLICINFO[(CREDITCHECKGUID = '" & strCreateCreditCheckGUID & "') and " _
                        & "(FBBLOCKID = '" & strCreateFBBlockID & "') and" _
                        & "(FBHEADERSEQUENCE = '" & strCreateFBHeaderSequence & "') and" _
                        & "(FBPUBLICINFOSEQUENCE = '" & strCreateFBPublicInfoSequence & "') ]")
                    
                    If xmlMatchNode.selectSingleNode("MOSTRECENTDECLAREDDATE") Is Nothing Then
                    
                        Set xmlTempElem = xmlRequestDoc.createElement("MOSTRECENTDECLAREDDATE")
                        xmlTempElem.Text = strBHMatchMostRecentDec
                        xmlMatchNode.appendChild xmlTempElem
                    Else
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlMatchNode, "MOSTRECENTDECLAREDDATE", False)
                        If Not xmlTempElem Is Nothing Then
                            strBureauMatchMostRecentDec = xmlTempElem.Text
                        End If
                    
                        If CDate(strBHMatchMostRecentDec) > CDate(strBureauMatchMostRecentDec) Then
                            xmlTempElem.Text = strBHMatchMostRecentDec
                        End If
                    End If
                    



'                If Not (xmlBanktuptcyToCreate Is Nothing) Then ''**NEW **
'                    strCreateCreditCheckGUID = xmlBanktuptcyToCreate.selectSingleNode(".//CREDITCHECKGUID").Text ''**NEW **
'                    strCreateFBBlockID = xmlBanktuptcyToCreate.selectSingleNode(".//FBBLOCKID").Text ''**NEW **
'                    strCreateFBHeaderSequence = xmlBanktuptcyToCreate.selectSingleNode(".//FBHEADERSEQUENCE").Text ''**NEW **
'                    strCreateFBPublicInfoSequence = xmlBanktuptcyToCreate.selectSingleNode(".//FBPUBLICINFOSEQUENCE").Text ''**NEW **
        
                If intDataTypeIndex < (xmlBankruptcyList.length - 1) Then
                    strGUID = strGUID & _
                        "(CREDITCHECKGUID != '" & strCreateCreditCheckGUID & "') and "
                    strID = strID & _
                        "(FBBLOCKID != '" & strCreateFBBlockID & "') and "
                    strHeaderSequence = strHeaderSequence & _
                        "(FBHEADERSEQUENCE != '" & strCreateFBHeaderSequence & "') and "
                    strPublicInfoSequence = strPublicInfoSequence & _
                        "(FBPUBLICINFOSEQUENCE != '" & strCreateFBPublicInfoSequence & "')  and "

                Else
                    strGUID = strGUID & _
                        "(CREDITCHECKGUID != '" & strCreateCreditCheckGUID & "')"
                    strID = strID & _
                        "(FBBLOCKID != '" & strCreateFBBlockID & "')"
                    strHeaderSequence = strHeaderSequence & _
                        "(FBHEADERSEQUENCE != '" & strCreateFBHeaderSequence & "')"
                    strPublicInfoSequence = strPublicInfoSequence & _
                        "(FBPUBLICINFOSEQUENCE != '" & strCreateFBPublicInfoSequence & "')"

                End If
                
             Next
            
            If (xmlBankruptcyList.length > 0) Then
       
                    strRemovalPattern = ".//FULLBUREAUPUBLICINFOLIST/" _
                    & "FULLBUREAUPUBLICINFO[(FBPUBLICINFODATATYPE != '56' and " _
                    & "FBPUBLICINFODATATYPE != '57') " _
                    & " and ((FBPUBLICINFODATATYPE = '24') " _
                    & "or (FBPUBLICINFODATATYPE = '25') or (FBPUBLICINFODATATYPE = '26') or (FBPUBLICINFODATATYPE = '27') " _
                    & "or (FBPUBLICINFODATATYPE = '28') or (FBPUBLICINFODATATYPE = '30') or (FBPUBLICINFODATATYPE = '32') " _
                    & "or (FBPUBLICINFODATATYPE = '33') or (FBPUBLICINFODATATYPE = '34') or (FBPUBLICINFODATATYPE = '35') " _
                    & "or (FBPUBLICINFODATATYPE = '38') or (FBPUBLICINFODATATYPE = '41') or (FBPUBLICINFODATATYPE = '43') " _
                    & "or (FBPUBLICINFODATATYPE = '44') or (FBPUBLICINFODATATYPE = '45') or (FBPUBLICINFODATATYPE = '47') " _
                    & "or (FBPUBLICINFODATATYPE = '48') or (FBPUBLICINFODATATYPE = '49') or (FBPUBLICINFODATATYPE = '50') " _
                    & "or (FBPUBLICINFODATATYPE = '51') or (FBPUBLICINFODATATYPE = '52') or (FBPUBLICINFODATATYPE = '53') or (FBPUBLICINFODATATYPE = '54')) " _
                    & " and ((" & strGUID & ") or (" & strID & ") or (" & strHeaderSequence & ") or (" & strPublicInfoSequence & ") )]"
                    
                    Set xmlListToRemove = xmlBureauData.selectNodes(strRemovalPattern) ''**NEW**
                        
    '                End If ''**NEW**
                    
                    If Not (xmlListToRemove Is Nothing) Then
                        If xmlListToRemove.length > 0 Then ''**NEW**
                            Set xmlFBPublicInfoList = xmlBureauData.selectSingleNode("FULLBUREAUPUBLICINFOLIST") ''**NEW**
                            For intRemoveIndex = 0 To (xmlListToRemove.length - 1) ''**NEW**
                                xmlFBPublicInfoList.removeChild xmlListToRemove.Item(intRemoveIndex) ''**NEW**
                            Next ''**NEW**
                            'Refresh
                            'BMIDS630 INR any unassigned records (i.e. they don't have a customer attached),
                            'have been removed bythe time we get here
                            Set xmlNewRecordsNodeList = xmlBureauData.selectNodes(".//FULLBUREAUPUBLICINFOLIST/" _
                            & "FULLBUREAUPUBLICINFO[FBPUBLICINFODATATYPE != '56' and " _
                            & "FBPUBLICINFODATATYPE != '57']") ''**NEW **
                            
                            
                        End If ''**NEW**
                    End If
            Else
                        'Refresh
                        'BMIDS630 INR any unassigned records (i.e. they don't have a customer attached),
                        'have been removed bythe time we get here
                        Set xmlNewRecordsNodeList = xmlBureauData.selectNodes(".//FULLBUREAUPUBLICINFOLIST/" _
                        & "FULLBUREAUPUBLICINFO[FBPUBLICINFODATATYPE != '56' and " _
                        & "FBPUBLICINFODATATYPE != '57']") ''**NEW **
            
            End If

            
           
            ''BMIDS610 Further fixes to remove duplicate BankRuptcyHistory Records END
    ''---------------------------  GD BMIDS610 FURTHER FIXES END  ---------------------------
            For intIndex = 0 To xmlNewRecordsNodeList.length - 1
                
                Set xmlNode = xmlNewRecordsNodeList.Item(intIndex)
                
                'Find owner of this record
                Set xmlCustomerNodeList = xmlNode.selectNodes("CUSTOMERLIST/CUSTOMER")
                If xmlCustomerNodeList.length > 0 Then
                    'Customer(s) found
                    blnUnAssigned = False
                Else
                    'Default assignment is to Customer 1
                    Set xmlCustomerElem = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, _
                                                                        "CUSTOMERLIST/CUSTOMER[1]")
                    blnUnAssigned = True
                End If
        
                Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                intTempValue = CInt(m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFODATATYPE"))
                Select Case intTempValue
                'CC037  Change case selections.
                'Case 42, 43, 44, 45, 46
                Case 29, 31, 37, 40, 42, 46
                    'Check if matching record already exists in Omiga
                    'BMIDS00801 MDC 09/11/2002
                    Set xmlTempParentElem = Nothing
                    'BMIDS00801 MDC 09/11/2002 - End
                    
                    'BMIDS00525 MDC 26/09/2002
                    If Not xmlBankruptcyListNode Is Nothing Then
                        'BMIDS610 START 25/07/2003
                        'strTempValue = m_objXmlAssist.GetElementText(xmlNode, "FBCOURTPLAINFIFF")
                        strTempValue = m_objXmlAssist.GetElementText(xmlNode, "FBCOURTPLAINTIFF")
                        'BMIDS610 END 25/07/2003
                        Set xmlTempParentElem = m_objXmlAssist.GetNode(xmlBankruptcyListNode, "BANKRUPTCYHISTORY" _
                                                & "[COURTPLAINTIFF = '" & strTempValue & "']", False)
                                                                 
                    End If
                    'BMIDS00525 MDC 26/09/2002 - End
                    
                    If xmlTempParentElem Is Nothing Then
                        'Create new BankruptcyHistory record
                        Set xmlTempParentElem = xmlRequestDoc.createElement("BANKRUPTCYHISTORY")
                        blnCreateNew = True
                        'BMIDS630 DATEDECLARED & DATEOFDISCHARGE are filled in based on publicinfodatatype.
                        If intTempValue = 46 Then
                            m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEOFDISCHARGE", xmlNode, xmlTempParentElem, False
                            'if we are creating a '46' and we have a DATEDECLARED node, this has
                            'come from its associated non discharged (which was removed) and needs
                            'to be included.
                                If Not xmlNode.selectSingleNode("DATEDECLARED") Is Nothing Then
                                
                                    'BMIDS630 Only create if it doesn't already exist
                                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEDECLARED", False)
                                    If Not xmlTempElem Is Nothing Then
                                        'Special case, xmlNode has a DATEDECLARED here.
                                        xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "DATEDECLARED")
                                    Else
                                        m_objXmlAssist.CopyNode "DATEDECLARED", "DATEDECLARED", xmlNode, xmlTempParentElem, False
                                    End If
                            
                            End If
                        Else
                            m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEDECLARED", xmlNode, xmlTempParentElem, False
                        End If
                    Else
                        'Update existing BankruptcyHistory record
                        If intTempValue = 46 Then
                            'BMIDS630 Only create if it doesn't already exist
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEOFDISCHARGE", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                            Else
                                m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEOFDISCHARGE", xmlNode, xmlTempParentElem, False
                            End If
                        Else
                            'BMIDS630 Only create if it doesn't already exist
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEDECLARED", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                            Else
                                m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEDECLARED", xmlNode, xmlTempParentElem, False
                            End If
                        End If
                    End If
                    
                'Case 26, 27, 28, 34, 21, 49, 53
                'GD BMIDS610 FURTHER FIXES START
                ''*** NOTE NOTE NOTE ***if the items in this Case statement change, the strPatternMatch variable defined above WILL ALSO NEED CHANGING
                'GD BMIDS610 FURTHER FIXES END
                Case 24, 25, 26, 27, 28, 30, 32, 33, 34, 35, 38, 41, 43, 44, 45, 47, 48, 49, 50, 51, 52, 53, 54

                    'Check if matching record already exists in Omiga
                    Set xmlTempParentElem = Nothing
                    
                    'GD BMIDS610 FURTHER FIXES If intIndex = 0 Then
                        'Public Info records are sorted by date descending so this is the most recent
                        'Update most recent record preceding this date
                        'GD BMIDS610 FURTHER FIXES START
                        'strStartDate = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                       
                       'BMIDS630 Changed where strMostRecentDate & strDeclaredMostRecentNonDisch reside
'                        strStartDate = strMostRecentDate
                        strStartDate = m_objXmlAssist.GetElementText(xmlNode, "MOSTRECENTDATE")
                        If strStartDate = "01/01/1900" Then
                            strStartDate = vbNullString
                        End If
                        strDeclaredMostRecentNonDisch = m_objXmlAssist.GetElementText(xmlNode, "MOSTRECENTDECLAREDDATE")
                        If strDeclaredMostRecentNonDisch = "01/01/1900" Then
                            strDeclaredMostRecentNonDisch = vbNullString
                        End If

                        
                        'GD BMIDS610 FURTHER FIXES END
                        'BMIDS00525 MDC 26/09/2002
                        If Not xmlBankruptcyListNode Is Nothing Then
                            For Each xmlTempElem In xmlBankruptcyListNode.childNodes
                                strTempValue = m_objXmlAssist.GetElementText(xmlTempElem, "DATEDECLARED")
                                If IsDate(strTempValue) And IsDate(strStartDate) Then
                                    
                                    If CDate(strTempValue) <= CDate(strStartDate) Then
                                        'Do an update
                                        Set xmlTempParentElem = xmlTempElem.cloneNode(False)
                                        'Add BANKRUPTCYHISTORYGUID
                                        Set xmlTempBHistoryElem = xmlTempElem.cloneNode(True).selectSingleNode(".//BANKRUPTCYHISTORYGUID")
                                        If Not (xmlTempBHistoryElem Is Nothing) Then
                                            xmlTempParentElem.appendChild xmlTempBHistoryElem
                                        End If
                                        
                                        'Add CUSTOMERVERSIONBANKRUPTCYHISTORY
                                        Set xmlTempCVBHistoryElem = xmlTempElem.cloneNode(True).selectSingleNode(".//CUSTOMERVERSIONBANKRUPTCYHISTORY")
                                        If Not (xmlTempCVBHistoryElem Is Nothing) Then
                                            xmlTempParentElem.appendChild xmlTempCVBHistoryElem
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                        'BMIDS00525 MDC 26/09/2002 - End
                    'GD BMIDS610 FURTHER FIXES End If
                        
                    If xmlTempParentElem Is Nothing Then
                        'Create new BankruptcyHistory record
                        Set xmlTempParentElem = xmlRequestDoc.createElement("BANKRUPTCYHISTORY")
                        blnCreateNew = True
                        'BMIDS630 DATEDECLARED & DATEOFDISCHARGE are filled in based on publicinfodatatype.
                        If ((intTempValue = 26) Or (intTempValue = 27) Or (intTempValue = 28) _
                            Or (intTempValue = 41) Or (intTempValue = 52) Or (intTempValue = 53)) Then
                            'To get here, we have no bankruptcy records on the database,
                            'so only need to have the FBPUBLICINFOREGISTRATIONDATE from most recent
                            'non discharged bankruptcy in this import
                            Set xmlTempElem = xmlTempParentElem.selectSingleNode("DATEDECLARED")
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = strDeclaredMostRecentNonDisch
                            Else
                                Set xmlTempElem = xmlRequestDoc.createElement("DATEDECLARED")
                                xmlTempElem.Text = strDeclaredMostRecentNonDisch
                                xmlTempParentElem.appendChild xmlTempElem
                            End If
                            m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEOFDISCHARGE", xmlNode, xmlTempParentElem, False
                        Else
                            'BMIDS630 Only create if it doesn't already exist
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEDECLARED", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                            Else
                                m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEDECLARED", xmlNode, xmlTempParentElem, False
                            End If
                        End If
                    Else
                        'Update BankruptcyHistory record
                        If ((intTempValue = 26) Or (intTempValue = 27) Or (intTempValue = 28) _
                            Or (intTempValue = 41) Or (intTempValue = 52) Or (intTempValue = 53)) Then
                            
                            'BMIDS630 Only create if it doesn't already exist
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEOFDISCHARGE", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                            Else
                                m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEOFDISCHARGE", xmlNode, xmlTempParentElem, False
                            End If
                        Else
                            'Remove DATEOFDISCHARGE
                            Set xmlTempElem = xmlTempParentElem.selectSingleNode("DATEOFDISCHARGE")
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = vbNullString
                            Else
                                Set xmlTempElem = xmlRequestDoc.createElement("DATEOFDISCHARGE")
                                xmlTempElem.Text = vbNullString
                                xmlTempParentElem.appendChild xmlTempElem
                            End If
                            'BMIDS630 Only create if it doesn't already exist
                            Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "DATEDECLARED", False)
                            If Not xmlTempElem Is Nothing Then
                                xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOREGISTRATIONDATE")
                            Else
                                m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEDECLARED", xmlNode, xmlTempParentElem, False
                            End If
                        End If

                    End If
                        
                        
                'Case 2, 17
                Case 2, 17, 39
                    'Create CCJ History
                    'GD BM0482 START 3
                    strPublicInfoSatisfiedDate = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOSATISFIEDDATE")
                    'if FBPUBLICINFOSATISFIEDDATE is Populated then
                    If Len(strPublicInfoSatisfiedDate) = 0 Then
                    '   Set CCJType = 'S'
                        strCCJType = objComboDO.GetFirstComboValueId("CCJType", "U")
                    'else
                    Else
                    '   Set CCJType = 'U'
                        strCCJType = objComboDO.GetFirstComboValueId("CCJType", "S")
                    'end if
                    End If
                    'GD BM0482 END 3
                    

                    
                    Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
                    Set xmlTempParentElem = xmlRequestDoc.createElement("CCJHISTORY")
                    'BMIDS00525 MDC 26/09/2002
                    xmlTempRequestElem.appendChild xmlTempParentElem
                    'BMIDS00525 MDC 26/09/2002 - End
                    
                    'BM0158 MDC 05/12/2002
                    m_objXmlAssist.SetMandatoryChildText xmlTempParentElem, "ADDITIONALINDICATOR", "0"
                    'BM0158 MDC 05/12/2002 - End
                    
                    m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEOFJUDGEMENT", xmlNode, xmlTempParentElem, False
                    m_objXmlAssist.CopyNode "FBPUBLICINFOSATISFIEDDATE", "DATECLEARED", xmlNode, xmlTempParentElem, False
                    
                    strBalance = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOAMOUNTINPOUNDS")
                    If Len(Trim$(strBalance)) > 0 Then
                        strTempValue = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOAMOUNTINPENCE")
                        If Len(Trim$(strTempValue)) = 0 Then
                            strTempValue = "00"
                        'BMIDS00525 MDC 27/09/2002
                        Else
                            strTempValue = Format$(strTempValue, "00")
                        End If
                        'BMIDS00525 MDC 27/09/2002 - End
                        strBalance = strBalance & "." & strTempValue
                    End If
                    Set xmlTempElem = xmlRequestDoc.createElement("VALUEOFJUDGEMENT")
                    xmlTempElem.Text = strBalance
                    xmlTempParentElem.appendChild xmlTempElem
                    
                    'GD BM0482 START 4
                    'Add CCJType to CreateCCJHistory Request
                    Set xmlTempElem = xmlRequestDoc.createElement("CCJTYPE")
                    xmlTempElem.Text = strCCJType
                    xmlTempParentElem.appendChild xmlTempElem
                    
                    'Set DEFAULTRECORD = 0
                    'Add DEFAULTRECORD to CreateCCJHistory
                    Set xmlTempElem = xmlRequestDoc.createElement("DEFAULTRECORD")
                    xmlTempElem.Text = "0"
                    xmlTempParentElem.appendChild xmlTempElem
                    'GD BM0482 END 4
                    
                    'BMIDS00525 MDC 27/09/2002 - Monthly Repayment is mandatory
                    Set xmlTempElem = xmlRequestDoc.createElement("MONTHLYREPAYMENT")
                    xmlTempElem.Text = "0"  'This data is not returned by bureau so default to 0
                    xmlTempParentElem.appendChild xmlTempElem
                    'BMIDS00525 MDC 27/09/2002 - End
                    
                    Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                    xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                    xmlTempParentElem.appendChild xmlTempElem
                    
                    Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                    xmlTempElem.Text = "1"
                    xmlTempParentElem.appendChild xmlTempElem
                    
                    'Owners
                    If blnUnAssigned Then
                        Set xmlTempChildElem = xmlRequestDoc.createElement("CUSTOMERVERSIONCCJHISTORY")
                        xmlTempParentElem.appendChild xmlTempChildElem
                        m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                        m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                    Else
                        For Each xmlCustomerElem In xmlCustomerNodeList
                            Set xmlTempChildElem = xmlRequestDoc.createElement("CUSTOMERVERSIONCCJHISTORY")
                            xmlTempParentElem.appendChild xmlTempChildElem
                            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempChildElem, False
                            m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempChildElem, False
                        Next
                    End If
                    
                    strResponse = objCustFinancialBO.CreateCCJHistory(xmlTempRequestElem.xml)
                    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                    Debug.Print "Created CCJ History record"
                    blnCCJHistory = True
                                    
                    Set xmlTempParentElem = Nothing
                'BM0158 MDC 05/12/2002
                Case Else
                    Set xmlTempParentElem = Nothing
                'BM0158 MDC 05/12/2002 - End
                    
                End Select
                
                If Not xmlTempParentElem Is Nothing Then
                    'Create/Update BankruptcyHistory record
                    xmlTempRequestElem.appendChild xmlTempParentElem
                    
                    'BMIDS630 This is now done above based on FBPublicInfoDataType
'                    m_objXmlAssist.CopyNode "FBPUBLICINFOREGISTRATIONDATE", "DATEDECLARED", xmlNode, xmlTempParentElem, False
'                    m_objXmlAssist.CopyNode "FBPUBLICINFOSATISFIEDDATE", "DATEOFDISCHARGE", xmlNode, xmlTempParentElem, False
                    
                    strBalance = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOAMOUNTINPOUNDS")
                    If Len(Trim$(strBalance)) > 0 Then
                        strTempValue = m_objXmlAssist.GetElementText(xmlNode, "FBPUBLICINFOAMOUNTINPENCE")
                        If Len(Trim$(strTempValue)) = 0 Then
                            strTempValue = "00"
                        End If
                        strBalance = strBalance & "." & strTempValue
                    End If
                    'BMIDS630 Only create if it doesn't already exist
                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "AMOUNTOFDEBT", False)
                    If Not xmlTempElem Is Nothing Then
                        xmlTempElem.Text = strBalance
                    Else
                        Set xmlTempElem = xmlRequestDoc.createElement("AMOUNTOFDEBT")
                        xmlTempElem.Text = strBalance
                        xmlTempParentElem.appendChild xmlTempElem
                    End If
                    
                    'BMIDS630 Only create if it doesn't already exist
                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "UNASSIGNED", False)
                    If Not xmlTempElem Is Nothing Then
                        xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                    Else
                        Set xmlTempElem = xmlRequestDoc.createElement("UNASSIGNED")
                        xmlTempElem.Text = CStr(Abs(CInt(blnUnAssigned)))
                        xmlTempParentElem.appendChild xmlTempElem
                    End If
                    
                    'BMIDS630 Only create if it doesn't already exist
                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "CREDITSEARCH", False)
                    If Not xmlTempElem Is Nothing Then
                        xmlTempElem.Text = "1"
                    Else
                        Set xmlTempElem = xmlRequestDoc.createElement("CREDITSEARCH")
                        xmlTempElem.Text = "1"
                        xmlTempParentElem.appendChild xmlTempElem
                    End If
                    
                    'BMIDS630 Only create if it doesn't already exist
                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "MONTHLYREPAYMENT", False)
                    If Not xmlTempElem Is Nothing Then
                        xmlTempElem.Text = "0"
                    Else
                        'BMIDS00801 MDC 08/11/2002
                        Set xmlTempElem = xmlRequestDoc.createElement("MONTHLYREPAYMENT")
                        xmlTempElem.Text = "0"
                        xmlTempParentElem.appendChild xmlTempElem
                        'BMIDS00801 MDC 08/11/2002 - End
                    End If
                    
                    'BMIDS630 Only create if it doesn't already exist
                    Set xmlTempElem = m_objXmlAssist.GetNode(xmlTempParentElem, "COURTPLAINTIFF", False)
                    If Not xmlTempElem Is Nothing Then
                        xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBCOURTPLAINTIFF")
                    Else
                        'BMIDS610 GD 25/07/2003 START
                        Set xmlTempElem = xmlRequestDoc.createElement("COURTPLAINTIFF")
                        xmlTempElem.Text = m_objXmlAssist.GetElementText(xmlNode, "FBCOURTPLAINTIFF")
                        xmlTempParentElem.appendChild xmlTempElem
                        'BMIDS610 GD 25/07/2003 END
                    End If
                    
                    
                    'Owners
                    'Get existing owners
                    Set xmlExistingOwnersList = xmlRequestDoc.createElement("OWNERS")
                    For Each xmlTempElem In xmlTempParentElem.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
                        xmlExistingOwnersList.appendChild xmlTempElem
                    Next
                    
                    'Add in any new owners
                    For Each xmlCustomerElem In xmlCustomerNodeList
                        strCustomerNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERNUMBER")
                        strCustomerVersionNumber = m_objXmlAssist.GetElementText(xmlCustomerElem, "CUSTOMERVERSIONNUMBER")
                        'BMIDS00801 MDC 09/11/2002
                        Set xmlTempElem = m_objXmlAssist.GetNode(xmlExistingOwnersList, _
                                "CUSTOMERVERSIONBANKRUPTCYHISTORY[CUSTOMERNUMBER='" & strCustomerNumber & "' and " _
                                & "CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']", False)
                                
                        If Not xmlTempElem Is Nothing Then
                            xmlExistingOwnersList.removeChild xmlTempElem
                        Else
                            'Add new owner
                            Set xmlTempElem = xmlRequestDoc.createElement("CUSTOMERVERSIONBANKRUPTCYHISTORY")
                            xmlTempParentElem.appendChild xmlTempElem
                            m_objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerElem, xmlTempElem, False
                            m_objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomerElem, xmlTempElem, False
                        End If
                    Next
                    
                    'Delete existing owners not now on this account
                    If xmlExistingOwnersList.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY").length > 0 Then
                        Set xmlTempChildElem = xmlRequestDoc.createElement("DELETE")
                        xmlTempParentElem.appendChild xmlTempChildElem
                    End If
                    For Each xmlCustomerElem In xmlExistingOwnersList.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY")
                        xmlTempChildElem.appendChild xmlCustomerElem
                    Next
                    
                    If blnCreateNew Then
                        'Create new BankruptcyHistory record  BMIDS590  Add codes 29, 31, 37, 40
                        If intTempValue = 42 Or intTempValue = 46 Or intTempValue = 29 Or _
                           intTempValue = 31 Or intTempValue = 37 Or intTempValue = 40 Then
                            Set xmlTempElem = xmlRequestDoc.createElement("IVA")
                            xmlTempElem.Text = "1"
                            xmlTempParentElem.appendChild xmlTempElem
                        End If
                        strResponse = objCustFinancialBO.CreateBankruptcyHistory(xmlTempRequestElem.xml)
                        blnBankruptcyHistory = True
                        Debug.Print "Created Bankruptcy History record"
                    Else
                        'Update matching BankruptcyHistory record
                        strResponse = objCustFinancialBO.UpdateBankruptcyHistory(xmlTempRequestElem.xml)
                        Debug.Print "Updated Bankruptcy History record"
                    End If
                    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
                    
                End If
            Next
        End If
        
    End If

    '=======================================================================================
    ' Step E - Update Financial Summary
    '=======================================================================================
    'INR BMIDS597 Need to create or update this even if none are true for
    'the No indicator to display correctly on screens DC090 130 140 150
'    If blnArrearsHistory Or blnBankruptcyHistory Or blnCCJHistory Or blnExistingMortgage Or blnLoanLiability Then
        
        'Update the Financial Summary to reflect newly created records
        Set xmlTempRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
        Set xmlTempParentElem = xmlRequestDoc.createElement("FINANCIALSUMMARY")
        xmlTempRequestElem.appendChild xmlTempParentElem
        
        m_objXmlAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", xmlApplicationElement, _
                                                                        xmlTempParentElem, False
        m_objXmlAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", _
                                                    xmlApplicationElement, xmlTempParentElem, False
        
        'BMIDS01124 MDC 03/12/2002 - Check if Financial Summary exists already?
        strResponse = objApplicationBO.GetFinancialSummary(xmlTempRequestElem.xml)
        Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        If CheckSuccessOrRecordNotFound(xmlTempResponseDoc.documentElement, cstrFunctionName) Then
            'Record found
            blnCreateNew = False
        Else
            'Need to create a new record
            blnCreateNew = True
        End If
        'BMIDS01124 MDC 03/12/2002 - End
        
        'BMIDS597 When no records created need to set each of the questions for financial summary to false.
        Set xmlTempElem = xmlRequestDoc.createElement("ARREARSHISTORYINDICATOR")
        If blnArrearsHistory Then
            xmlTempElem.Text = "1"
        Else
             xmlTempElem.Text = "0"
        End If
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("BANKRUPTCYHISTORYINDICATOR")
        If blnBankruptcyHistory Then
            xmlTempElem.Text = "1"
        Else
            xmlTempElem.Text = "0"
        End If
        xmlTempParentElem.appendChild xmlTempElem
            
        Set xmlTempElem = xmlRequestDoc.createElement("CCJHISTORYINDICATOR")
        If blnCCJHistory Then
            xmlTempElem.Text = "1"
        Else
            xmlTempElem.Text = "0"
        End If
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("EXISTINGMORTGAGEINDICATOR")
        If blnExistingMortgage Then
            xmlTempElem.Text = "1"
        Else
            xmlTempElem.Text = "0"
        End If
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("LOANLIABILITYINDICATOR")
        If blnLoanLiability Then
            xmlTempElem.Text = "1"
        Else
            xmlTempElem.Text = "0"
        End If
        xmlTempParentElem.appendChild xmlTempElem
        
        'BMIDS01124 MDC 03/12/2002 - Check if Financial Summary exists already?
        'strResponse = objApplicationBO.UpdateFinancialSummary(xmlTempRequestElem.xml)
        If blnCreateNew Then
            strResponse = objApplicationBO.CreateFinancialSummary(xmlTempRequestElem.xml)
        Else
            strResponse = objApplicationBO.UpdateFinancialSummary(xmlTempRequestElem.xml)
        End If
        'BMIDS01124 MDC 03/12/2002 - End
        
        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        
    'INR BMIDS597
'    End If
    
    '=======================================================================================
    ' Response
    '=======================================================================================
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    ImportBureauData = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ImportBureauDataExit:
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlTempRequestElem = Nothing
    Set xmlTempElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlBureauData = Nothing
    Set xmlAliasListNode = Nothing
    Set xmlAliasNode = Nothing
    Set xmlTempParentElem = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlPersonNode = Nothing
    Set xmlMortgageAccountListNode = Nothing
    Set xmlLiabilityListNode = Nothing
    Set xmlArrearsListNode = Nothing
    Set xmlTempChildElem = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlMortgageLoanList = Nothing
    Set xmlAccount = Nothing
    Set xmlCAISNode = Nothing
    Set xmlBankruptcyListNode = Nothing
    Set xslDoc = Nothing
    Set xmlExistingOwnersList = Nothing
    Set xmlNewRecordsNodeList = Nothing
    Set xmlApplicationElement = Nothing
    Set xmlCustomerListElement = Nothing
    'Set objCreditCheckBO = Nothing
    Set objCustFinancialBO = Nothing
    Set objCustomerBO = Nothing
    Set objGlobalParamBO = Nothing
    Set objComboDO = Nothing
    Set objApplicationBO = Nothing
    Set colValueIds = Nothing
    Set xmlTransformedDoc = Nothing 'BMIDS00801 MDC 09/11/2002
    'Variables used For BM0373 START
    Set xmlAppFFRequest = Nothing
    Set xmlAppFFElem = Nothing
    Set xmlAppFFNum = Nothing
    Set xmlAppNum = Nothing
    Set xmlGetAppFFDataResponseDoc = Nothing
    Set xmlComboResponseDoc = Nothing
    Set xmlRedemptionElement = Nothing
    'Variables used For BM0373 END
    'Variables used For BMIDS593 START
    Set xmlCAISListDoc = Nothing
    Set xmlNewCAISAccounts = Nothing
    'BMIDS593 END
    'BMIDS630 END
    Set xmlTempType46Elem = Nothing
    Set xmlTempNonType46Elem = Nothing
    Set xmlType46NodeList = Nothing
    Set xmlNonDischargedToRemoveList = Nothing
    Set xmlFBPublicInfoList = Nothing
    'BMIDS630 END
    
    Exit Function

ImportBureauDataVbErr:

    m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    'GD TESTING
    Debug.Print "ERROR LOGGED : " & " " & Err.Number & " " & Err.Description & " " & vbLogEventTypeError
    m_objErrAssist.AddToErrSource cstrFunctionName

    ImportBureauData = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ImportBureauDataExit

End Function
'BMIDS00336 MDC 14/08/2002 - End

'BMIDS00336 MDC 14/08/2002
Private Function CheckSuccessOrRecordNotFound(ByVal xmlResponseToCheck As IXMLDOMNode, _
                                            ByVal strFunction As String) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    CheckSuccessOrRecordNotFound
' author:       Mark Coleman
' date:         14/08/2002
' description:  Helper function to check if a response is Success or Record Not Found error
'               Used to ignore 'Record Not Found' errors
' pass:         xmlResponseToCheck  XML Input specifying the XML Response data to be checked
'               strFunction         Name of the calling function
' return:       True if success, False if record Not Found
' errors:       If an error other than Record Not Found, raise it
'------------------------------------------------------------------------------------------
On Error GoTo CheckSuccessOrRecordNotFoundVbErr

Const strFunctionName As String = "CheckSuccessOrRecordNotFound"
    
Dim lngErr As Long

    'Check response ignoring record not found error
    lngErr = m_objErrAssist.CheckResponse(xmlResponseToCheck, xmlResponseToCheck, False)
    If lngErr <> 0 Then
        lngErr = m_objErrAssist.GetOmigaErrorNumber(lngErr)
    End If
    
    If lngErr <> 0 And lngErr <> oeRecordNotFound Then
        m_objErrAssist.ThrowError TypeName(Me), strFunction, lngErr
    ElseIf lngErr = oeRecordNotFound Then
        CheckSuccessOrRecordNotFound = False
    Else
        CheckSuccessOrRecordNotFound = True
    End If
    
    Exit Function
    
CheckSuccessOrRecordNotFoundVbErr:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'BMIDS00336 MDC 14/08/2002 - End

Friend Function FindQQRegularOutgoings(ByVal vstrXMLRequest As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Called from FindCFRegularOutgoings:
'               If no Customer Financial Regular Outgoings then use anything available from
'               the Quick Quote process.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindQQRegularOutgoingsVbErr
    
    Const strFunctionName As String = "FindQQRegularOutgoings"


    Dim objQuickQuoteDO As QuickQuoteOutgoingsDO
    Dim objContext As ObjectContext
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTemp As FreeThreadedDOMDocument40
    
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set objContext = GetObjectContext()
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
 
    If objContext Is Nothing Then
        Set objQuickQuoteDO = New QuickQuoteOutgoingsDO
    Else
        Set objQuickQuoteDO = objContext.CreateInstance(gstrQUICK_QUOTE & ".QuickQuoteOutgoingsDO")
    End If
    
    Set xmlTemp = objXmlAssist.load(objQuickQuoteDO.FindList(vstrXMLRequest), TypeName(Me), strFunctionName)
    If Not xmlTemp Is Nothing Then
        Set xmlDataNode = xmlTemp.firstChild
    End If
    
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    Set FindQQRegularOutgoings = xmlOut.firstChild
   
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindQQRegularOutgoingsExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteDO = Nothing

    Exit Function

FindQQRegularOutgoingsVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindQQRegularOutgoings = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindQQRegularOutgoingsExit

End Function

Public Function FindCFRegularOutgoings(ByVal vstrXMLRequest As String) As String
Attribute FindCFRegularOutgoings.VB_Description = "Gets data from RegularOutgoings table, if None exists gets defaults from QuickQuoteoutgoings table."
Attribute FindCFRegularOutgoings.VB_UserMemId = 0
' header ----------------------------------------------------------------------------------
' description:  If no Customer Financial Regular Outgoings then use anything available from
'               the Quick Quote process.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindCFRegularOutgoingsVbErr
    
    Const strFunctionName As String = "FindCFRegularOutgoings"


    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim objContext As ObjectContext
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim sReturnXML As String
    
    Set objContext = GetObjectContext()
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
 
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
    
    ' Check validity of input XML:
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrREGULAROUTGOINGS_NAME).Item(0)
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrREGULAROUTGOINGS_NAME & " tag not found"
    End If
     
    
    ' See if there are any real Regular Outgoings:
    Set xmlDataNode = objICustomerFinancialDO.FindRegularOutgoingsList(xmlIn.firstChild)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
        xmlResponseElement.setAttribute "TYPE", "SUCCESS"
        FindCFRegularOutgoings = xmlOut.xml
    Else
        'If not look in Quick Quotes for defaults:
         vstrXMLRequest = Replace(vstrXMLRequest, "REGULAROUTGOINGS", "QUICKQUOTEOUTGOINGS")
        sReturnXML = FindQQRegularOutgoings(vstrXMLRequest).xml
        sReturnXML = Replace(sReturnXML, "QUICKQUOTEOUTGOINGS", "REGULAROUTGOINGS")
        sReturnXML = Replace(sReturnXML, "FREQUENCYSEQUENCENUMBER", "PAYMENTFREQUENCY")
        FindCFRegularOutgoings = sReturnXML
    End If
   
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindRegularOutgoingsExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindCFRegularOutgoingsVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCFRegularOutgoings = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindRegularOutgoingsExit

End Function

Public Function FindBankCardSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindBankCardSummaryVbErr
    
    Const strFunctionName As String = "FindBankCardSummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'Dim objIApplicationDO As IApplicationDO  'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO  'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO") 'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
    
'****SR 25/05/2004 : BMIDS772
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlDataNode = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlDataNode Is Nothing) Then
'        xmlResponseElement.appendChild xmlDataNode
'    End If
'****SR 25/05/2004 : BMIDS772  - End

    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrBANKCREDITCARDLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrBANKCREDITCARDLIST_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindBankCardList(xmlRequestNode)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    FindBankCardSummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindBankCardSummaryExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindBankCardSummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBankCardSummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindBankCardSummaryExit

End Function

Public Function FindBankruptcyHistorySummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindBankruptcyHistorySummaryVbErr
    
    Const strFunctionName As String = "FindBankruptcyHistorySummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlFinancialSummaryRequestNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'Dim objIApplicationDO As IApplicationDO    'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO    'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")     'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
        
'*** SR 25/05/2004 : BMIDS772
'    Dim strXMLData As String
'
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlData = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlData Is Nothing) Then
'        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
'    End If
'*** SR 25/05/2004 : BMIDS772 - End
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrBANKRUPTCYHISTORYLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrBANKRUPTCYHISTORYLIST_NAME & " tag not found"
    End If
    
    Set xmlData = objICustomerFinancialDO.FindBankruptcyHistoryList(xmlRequestNode)
    If Not (xmlData Is Nothing) Then
        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    objXmlAssist.WriteXMLToFile xmlOut.xml
    FindBankruptcyHistorySummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindBankruptcyHistorySummaryExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindBankruptcyHistorySummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        xmlErrDoc.validateOnParse = False
        xmlErrDoc.setProperty "NewParser", True
        Set xmlErrDoc = objXmlAssist.load( _
            objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBankruptcyHistorySummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindBankruptcyHistorySummaryExit

End Function

Public Function FindDeclinedMortgageSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindDeclinedMortgageSummaryVbErr
    
    Const strFunctionName As String = "FindDeclinedMortgageSummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlFinancialSummaryRequestNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'Dim objIApplicationDO As IApplicationDO     'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO   'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")     'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
'***SR 25/05/2004 : BMIDS772
'    Dim strXMLData As String
'
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlData = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlData Is Nothing) Then
'        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
'    End If
'***SR 25/05/2004 : BMIDS772
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrDECLINEDMORTGAGELIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrDECLINEDMORTGAGELIST_NAME & " tag not found"
    End If
    
    Set xmlData = objICustomerFinancialDO.FindDeclinedMortgageList(xmlRequestNode)
    If Not (xmlData Is Nothing) Then
        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    objXmlAssist.WriteXMLToFile xmlOut.xml
    FindDeclinedMortgageSummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindDeclinedMortgageSummaryExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindDeclinedMortgageSummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        xmlErrDoc.validateOnParse = False
        xmlErrDoc.setProperty "NewParser", True
        Set xmlErrDoc = objXmlAssist.load( _
            objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindDeclinedMortgageSummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindDeclinedMortgageSummaryExit

End Function

Public Function FindCCJHistorySummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindCCJHistorySummaryVbErr
    
    Const strFunctionName As String = "FindCCJHistorySummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlFinancialSummaryRequestNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'Dim objIApplicationDO As IApplicationDO    'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO  'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")    'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If

'***SR 25/05/2004 : BMIDS772
'    Dim strXMLData As String
'
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlData = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlData Is Nothing) Then
'        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
'    End If
'***SR 25/05/2004 : BMIDS772 - End

    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrCCJHISTORYLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrCCJHISTORYLIST_NAME & " tag not found"
    End If
    
    Set xmlData = objICustomerFinancialDO.FindCCJHistoryList(xmlRequestNode)
    If Not (xmlData Is Nothing) Then
        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    FindCCJHistorySummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindCCJHistorySummaryExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindCCJHistorySummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        xmlErrDoc.validateOnParse = False
        xmlErrDoc.setProperty "NewParser", True
        Set xmlErrDoc = objXmlAssist.load( _
            objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCCJHistorySummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindCCJHistorySummaryExit

End Function

Public Function FindLifeProductSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindLifeProductSummaryVbErr
    
    Const strFunctionName As String = "FindLifeProductSummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objIApplicationDO As IApplicationDO
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
        
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrPOLICYRELATIONSHIPLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrPOLICYRELATIONSHIPLIST_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objICustomerFinancialDO.FindLifeProductList(xmlRequestNode)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    FindLifeProductSummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindLifeProductSummaryExit:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    Set xmlRequestNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindLifeProductSummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLifeProductSummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindLifeProductSummaryExit

End Function

 'SR 25/05/2004 : BMIDS772 - new method
Public Function GetFinancialSummaryView(ByVal vstrXMLRequest As String) As String

On Error GoTo GetFinancialSummaryViewVbErr
    
    Const cstrFunctionName As String = "GetFinancialSummaryView"
    
    Dim xmlIn As FreeThreadedDOMDocument40, xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElement As IXMLDOMElement, xmlElement As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode, xmlFinancialSummary As IXMLDOMElement
    
    Dim objApplicationBO As ApplicationBO, objCFBO As CustomerFinancialBO
    
    Dim strResponse As String, strApplNumber As String, strAFFNumber As String
    Dim lngErrorNo As Long
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
        
    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objCFBO = New CustomerFinancialBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
    End If
    
    ' Get Financial Summary data for all customers and attach to Response
    strResponse = objCFBO.GetFinancialSummaryView(vstrXMLRequest)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    lngErrorNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrorNo <> 0 And lngErrorNo <> oeRecordNotFound Then
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElement, True
    End If
    
    m_objXmlAssist.AttachResponseData xmlResponseElement, xmlOut.documentElement
    
    '--------------------------------------------------------------------------
    ' Get Financial Summary data specific to application and attach to Response
    '--------------------------------------------------------------------------
    strApplNumber = xmlIn.selectSingleNode("//APPLICATIONNUMBER").Text
    strAFFNumber = xmlIn.selectSingleNode("//APPLICATIONFACTFINDNUMBER").Text
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(xmlIn.documentElement)
    Set xmlFinancialSummary = xmlIn.createElement("FINANCIALSUMMARY")
    xmlNewRequest.appendChild xmlFinancialSummary
    
    Set xmlElement = xmlIn.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlFinancialSummary.appendChild xmlElement
    
    Set xmlElement = xmlIn.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlFinancialSummary.appendChild xmlElement
    
    strResponse = objApplicationBO.GetFinancialSummary(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    lngErrorNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrorNo <> 0 And lngErrorNo <> oeRecordNotFound Then
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElement, True
    End If
    
    m_objXmlAssist.AttachResponseData xmlResponseElement, xmlOut.documentElement
        
    GetFinancialSummaryView = xmlResponseElement.xml

    m_objContext.SetComplete

GetFinancialSummaryViewExit:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    Set xmlNewRequest = Nothing
    Set xmlFinancialSummary = Nothing
    Set objApplicationBO = Nothing
    Set objCFBO = Nothing
    
    Exit Function

GetFinancialSummaryViewVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        m_objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetFinancialSummaryView = m_objErrAssist.CreateErrorResponse
            
    m_objContext.SetComplete
    
    '   go to clean-up section
    Resume GetFinancialSummaryViewExit

End Function

 'SR 25/05/2004 : BMIDS772 - new method
Public Function GetCreditHistorySummaryView(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCreditHistorySummaryViewVbErr
    
    Const cstrFunctionName As String = "GetCreditHistoryView"
    
    Dim xmlIn As FreeThreadedDOMDocument40, xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElement As IXMLDOMElement, xmlElement As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode, xmlFinancialSummary As IXMLDOMElement
    
    Dim objApplicationBO As ApplicationBO, objCFBO As CustomerFinancialBO
    
    Dim strResponse As String, strApplNumber As String, strAFFNumber As String
    Dim lngErrorNo As Long
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
        
    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objCFBO = New CustomerFinancialBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
    End If
    
    ' Get Credit History Summary data for all customers and attach to Response
    strResponse = objCFBO.GetCreditHistorySummaryView(vstrXMLRequest)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    lngErrorNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrorNo <> 0 And lngErrorNo <> oeRecordNotFound Then
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElement, True
    End If
    
    m_objXmlAssist.AttachResponseData xmlResponseElement, xmlOut.documentElement
    
    '--------------------------------------------------------------------------
    ' Get Financial Summary data specific to application and attach to Response
    '--------------------------------------------------------------------------
    strApplNumber = xmlIn.selectSingleNode("//APPLICATIONNUMBER").Text
    strAFFNumber = xmlIn.selectSingleNode("//APPLICATIONFACTFINDNUMBER").Text
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(xmlIn.documentElement)
    Set xmlFinancialSummary = xmlIn.createElement("FINANCIALSUMMARY")
    xmlNewRequest.appendChild xmlFinancialSummary
    
    Set xmlElement = xmlIn.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlFinancialSummary.appendChild xmlElement
    
    Set xmlElement = xmlIn.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlFinancialSummary.appendChild xmlElement
    
    strResponse = objApplicationBO.GetFinancialSummary(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    lngErrorNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrorNo <> 0 And lngErrorNo <> oeRecordNotFound Then
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElement, True
    End If
    
    m_objXmlAssist.AttachResponseData xmlResponseElement, xmlOut.documentElement
        
    GetCreditHistorySummaryView = xmlResponseElement.xml

    m_objContext.SetComplete

GetCreditHistorySummaryViewExit:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    Set xmlNewRequest = Nothing
    Set xmlFinancialSummary = Nothing
    Set objApplicationBO = Nothing
    Set objCFBO = Nothing
    
    Exit Function

GetCreditHistorySummaryViewVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        m_objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetCreditHistorySummaryView = m_objErrAssist.CreateErrorResponse
            
    m_objContext.SetComplete
    
    '   go to clean-up section
    Resume GetCreditHistorySummaryViewExit

End Function

Public Function FindArrearsHistorySummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindArrearsHistorySummaryVbErr
    
    Const strFunctionName As String = "FindArrearsHistorySummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlFinancialSummaryRequestNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    ' Dim objIApplicationDO As IApplicationDO  'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    
    If objContext Is Nothing Then
        ' Set objIApplicationDO = New ApplicationDO   'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")  'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
        
'***SR 25/05/2004 : BMIDS772
'    Dim strXMLData As String
    
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
    

'    Set xmlData = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlData Is Nothing) Then
'        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
'    End If
'****SR 25/05/2004 : BMIDS772 - End
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrARREARSHISTORYLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrARREARSHISTORYLIST_NAME & " tag not found"
    End If
    
    Set xmlData = objICustomerFinancialDO.FindArrearsHistoryList(xmlRequestNode)
    If Not (xmlData Is Nothing) Then
        Set xmlResponseNode = xmlDataNode.appendChild(xmlData)
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    FindArrearsHistorySummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindArrearsHistorySummaryExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindArrearsHistorySummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        xmlErrDoc.validateOnParse = False
        xmlErrDoc.setProperty "NewParser", True
        Set xmlErrDoc = objXmlAssist.load( _
            objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindArrearsHistorySummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindArrearsHistorySummaryExit

End Function

Public Function FindMortgageAccountSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'15/05/2003 LDM BM0492 if extra paramenter "SHOWNOTREDEEMED" has been tacked onto the end of the XML input
'                               make sure it is passed down to DO.FindMortgageAccountList
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageAccountSummaryVbErr
    
    Const strFunctionName As String = "FindMortgageAccountSummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objIApplicationDO As IApplicationDO
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO      'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")    'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
'***SR 25/05/2004 : BMIDS772
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlDataNode = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlDataNode Is Nothing) Then
'        xmlResponseElement.appendChild xmlDataNode
'    End If
'***SR 25/05/2004 : BMIDS772
    'BM0492 LDM 15/05/2003
    Set xmlDataNode = xmlIn.getElementsByTagName(cstrSHOWNOTREDEEMED_NAME).Item(0)
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrMORTGAGEACCOUNTLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrMORTGAGEACCOUNTLIST_NAME & " tag not found"
    End If
    
    'BM0492 LDM 15/05/2003
    ' if we have passed in this extra "parameter" make sure it get passed down to the DO
    If Not (xmlDataNode Is Nothing) Then
        xmlRequestNode.appendChild xmlDataNode
    End If
    
    Set xmlDataNode = objICustomerFinancialDO.FindMortgageAccountList(xmlRequestNode)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    FindMortgageAccountSummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindMortgageAccountSummaryExit:
    
    Set xmlDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElement = Nothing

    Set xmlIn = Nothing
    Set xmlOut = Nothing
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindMortgageAccountSummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageAccountSummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindMortgageAccountSummaryExit

End Function


Public Function FindLiabilitySummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindLiabilitySummaryVbErr
    
    Const strFunctionName As String = "FindLiabilitySummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'Dim objIApplicationDO As IApplicationDO        'SR 25/05/2004 : BMIDS772
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlElement = xmlOut.createElement("RESPONSE")
    Set xmlResponseElement = xmlOut.appendChild(xmlElement)
    
    If objContext Is Nothing Then
        'Set objIApplicationDO = New ApplicationDO      'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        'Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")  'SR 25/05/2004 : BMIDS772
        Set objICustomerFinancialDO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialDO")
    End If
    
'***SR 25/05/2004 : BMIDS772
'    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrFINANCIALSUMMARY_NAME).Item(0)
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
'            omiga4MissingPrimaryTag, cstrFINANCIALSUMMARY_NAME & " tag not found"
'    End If
'
'    Set xmlDataNode = objIApplicationDO.GetFinancialSummary(xmlRequestNode)
'    If Not (xmlDataNode Is Nothing) Then
'        xmlResponseElement.appendChild xmlDataNode
'    End If
'***SR 25/05/2004 : BMIDS772 - End
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrLOANSLIABILITIESLIST_NAME).Item(0)
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrLOANSLIABILITIESLIST_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objICustomerFinancialDO.FindLiabilityList(xmlRequestNode)
    If Not (xmlDataNode Is Nothing) Then
        xmlResponseElement.appendChild xmlDataNode
    End If
    
    '   if we get here, everything has completed OK
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    FindLiabilitySummary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindLiabilitySummaryExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseElement = Nothing
    Set xmlElement = Nothing
    Set xmlRequestNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

FindLiabilitySummaryVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLiabilitySummary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindLiabilitySummaryExit

End Function

Public Function FindThirdPartyForCustomer(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'CustomerRole'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindThirdPartyForCustomerVbErr

    Const cstrFunctionName As String = "FindThirdPartyForCustomer"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMElement

    Dim strThirdPartyType As String

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                            
    If xmlIn.selectSingleNode("//THIRDPARTY") Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, "THIRDPARTY" & " tag not found"
    End If
        
    strThirdPartyType = m_objXmlAssist.GetMandatoryElementText(xmlIn.documentElement, ".//THIRDPARTYTYPE")
    Select Case strThirdPartyType
        Case "10"  'Employer
            Set xmlTempResponseNode = GetEmploymentdetailsForCustomer(xmlIn.documentElement)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case "20"  'Accountant
            Set xmlTempResponseNode = GetAccountantDetailsForCustomer(xmlIn.documentElement)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case "30"  'Lender
            Set xmlTempResponseNode = GetMortgageAccountDetailsForCustomer(xmlIn.documentElement)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case "40"  'Landlord
            Set xmlTempResponseNode = GetLandLordDetailsForCustomer(xmlIn.documentElement)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeInvalidParameter, "Invalid Third Party Type"
    End Select
    
    FindThirdPartyForCustomer = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindThirdPartyForCustomerExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindThirdPartyForCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    FindThirdPartyForCustomer = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindThirdPartyForCustomerExit

End Function
Private Function GetDirectoryGUIDForBureau(ByVal vstrFBCaisCompanyType As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the thirdPartyGuid for this companytype. If it does not exist, create it
' pass:         vstrFBCaisCompanyType
' return:       the GUID
'------------------------------------------------------------------------------------------
 On Error GoTo GetDirectoryGUIDForBureauVbErr

    Const cstrFunctionName As String = "GetDirectoryGUIDForBureau"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTPRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTPCreateReqDoc As FreeThreadedDOMDocument40
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTPCreateNode As IXMLDOMNode
    Dim xmlTPNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    
    Dim objGlobalParamBO As GlobalParameterBO
    Dim objTPBO As ThirdPartyBO
    
    Dim strThirdPartyType As String
    Dim strTempValue As String
    Dim strCompanyName As String
    Dim strResponse As String
    Dim strReturnGUID As String
    
    If Not m_objContext Is Nothing Then
        Set objGlobalParamBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objTPBO = m_objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    Else
        Set objGlobalParamBO = New GlobalParameterBO
        Set objTPBO = New ThirdPartyBO
    End If

    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get the company name from FBCAISCOMPANYTYPE passed in
    If vstrFBCaisCompanyType = "01" Then
        strTempValue = "BureauDefaultUKBank"
    ElseIf vstrFBCaisCompanyType = "08" Then
        strTempValue = "BureauDefaultBuildingSociety"
    Else
        strTempValue = "BureauDefaultOtherLender"
    End If
    strResponse = objGlobalParamBO.GetCurrentParameter(strTempValue)
    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
        strCompanyName = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='" & strTempValue & "']/STRING")
    End If
    
    'Get the Directory GUID for this company name
    Set xmlTPRequestDoc = New FreeThreadedDOMDocument40
    xmlTPRequestDoc.validateOnParse = False
    xmlTPRequestDoc.setProperty "NewParser", True
    Set xmlRequestElem = xmlTPRequestDoc.createElement("REQUEST")
    xmlTPRequestDoc.appendChild xmlRequestElem
    Set xmlTPNode = xmlRequestElem.appendChild(xmlTPRequestDoc.createElement("NAMEANDADDRESSDIRECTORY"))
    Set xmlTempNode = xmlTPRequestDoc.createElement("COMPANYNAME")
    xmlTempNode.Text = strCompanyName
    xmlTPNode.appendChild xmlTempNode
    Set xmlTempNode = xmlTPRequestDoc.createElement("NAMEANDADDRESSTYPE")
    xmlTempNode.Text = "3"
    xmlTPNode.appendChild xmlTempNode
    
    strResponse = objTPBO.GetDirectoryGUID(xmlTPRequestDoc.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    strReturnGUID = m_objXmlAssist.GetElementText(xmlOut.documentElement, "./NAMEANDADDRESSDIRECTORY/DIRECTORYGUID")
    If strReturnGUID = "" Then
        'need to create a NameAndAddressDirectory then get the GUID
        Set xmlTPCreateReqDoc = New FreeThreadedDOMDocument40
        xmlTPCreateReqDoc.validateOnParse = False
        xmlTPCreateReqDoc.setProperty "NewParser", True
        xmlTPCreateReqDoc.appendChild xmlRequestElem.cloneNode(True)
        strResponse = objTPBO.CreateDirectory(xmlTPCreateReqDoc.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        
        strResponse = objTPBO.GetDirectoryGUID(xmlTPRequestDoc.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
        strReturnGUID = m_objXmlAssist.GetElementText(xmlOut.documentElement, "./NAMEANDADDRESSDIRECTORY/DIRECTORYGUID")
    End If
    
    GetDirectoryGUIDForBureau = strReturnGUID
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDirectoryGUIDForBureauExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDirectoryGUIDForBureauVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetDirectoryGUIDForBureau = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDirectoryGUIDForBureauExit

End Function

Private Function GetEmploymentdetailsForCustomer( _
                    ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMElement


On Error GoTo GetVbErr
    
    Const cstrFunctionName As String = "GetEmploymentdetailsForCustomer"
    'GD BMIDS00997
    Dim objComboDO As ComboDO
    
    Dim objCEBO As CustomerEmploymentBO, objTPBO As ThirdPartyBO
    Dim xmlNewRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement, xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlEmploymentList As IXMLDOMNodeList, xmlThirdParty As IXMLDOMElement
    Dim xmlElem2 As IXMLDOMElement
   
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strResponse As String, lngErrNo As Long
    
    Set objCEBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    
    'GD BMIDS00997
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlListElem = xmlOut.createElement("EMPLOYMENTLIST")
    xmlNewRequest.appendChild xmlListElem
    
    Set xmlTableElem = xmlOut.createElement("EMPLOYMENT")
    xmlListElem.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERNUMBER")
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
    
    strResponse = objCEBO.FindEmploymentList(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrNo <> 0 And lngErrNo <> oeRecordNotFound Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, lngErrNo
    End If
        
    'include only the nodes with EmploymentStatus as 10 or 30 - NO ! Use validation types!!
    
    
    Dim strEmploymentStatus As String
    Set xmlEmploymentList = xmlOut.selectNodes("//EMPLOYMENT")
    For Each xmlElement In xmlEmploymentList
        strEmploymentStatus = xmlElement.selectSingleNode(".//EMPLOYMENTSTATUS").Text
        'GD BMIDS00997 If strEmploymentStatus = "10" Or strEmploymentStatus = "20" Then
        If objComboDO.IsItemInValidation("EmploymentStatus", strEmploymentStatus, "E") Then
            Set xmlThirdParty = xmlOut.createElement("THIRDPARTY")
            
            Set xmlElem2 = xmlOut.createElement("CONTEXT")
            xmlElem2.Text = m_objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSEQUENCENUMBER")
            xmlThirdParty.appendChild xmlElem2
            
            'Fetch ThirdParty details and add CompanyName to Response
            Set xmlElem2 = xmlOut.createElement("NAME")
            xmlElem2.Text = GetTPCompanyName(vXMLRequest, xmlElement)
            xmlThirdParty.appendChild xmlElem2
            
            xmlResponseElem.appendChild xmlThirdParty
        End If
    Next xmlElement
        
    Set GetEmploymentdetailsForCustomer = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
       
GetVbExit:
    Set objCEBO = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlListElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    
    Exit Function
GetVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set GetEmploymentdetailsForCustomer = m_objErrAssist.CreateErrorResponseEx()
    Resume GetVbExit
End Function


Private Function GetAccountantDetailsForCustomer( _
                    ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMElement

On Error GoTo GetVbErr
    
    Const cstrFunctionName As String = "GetAccountantDetailsForCustomer"

    Dim objCEBO As CustomerEmploymentBO
    Dim xmlNewRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement, xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlEmploymentList As IXMLDOMNodeList
    Dim xmlThirdParty As IXMLDOMNode, xmlElem2 As IXMLDOMElement
    
    Dim strResponse As String, lngErrNo As Long
    
    Dim xmlAccountantRequest As IXMLDOMNode
    Dim xmlAccountant As IXMLDOMNode
    Dim xmlAccountantGuid As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlAccountantResponse As FreeThreadedDOMDocument40
   
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objCEBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlListElem = xmlOut.createElement("EMPLOYMENTLIST")
    xmlNewRequest.appendChild xmlListElem
    
    Set xmlTableElem = xmlOut.createElement("EMPLOYMENT")
    xmlListElem.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERNUMBER")
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
    
    strResponse = objCEBO.FindEmploymentList(xmlNewRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)

    'BMIDS00917 MDC 16/11/2002 - Include any employment with an Accountant
    'Dim strEmploymentStatus As String
    'Set xmlEmploymentList = xmlOut.selectNodes("//EMPLOYMENT")
    Set xmlEmploymentList = xmlOut.selectNodes("//EMPLOYMENT[ACCOUNTANTGUID != '']")
    'BMIDS00917 MDC 16/11/2002 - End
    
    ' PSC 11/12/01 SYS3361 - Start
    ' Set up Accountant Request
    Set xmlAccountantRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlAccountant = xmlOut.createElement("ACCOUNTANT")
    xmlAccountantRequest.appendChild xmlAccountant
    Set xmlAccountantGuid = xmlOut.createElement("ACCOUNTANTGUID")
    xmlAccountant.appendChild xmlAccountantGuid
    ' PSC 11/12/01 SYS3361 - End
            
    For Each xmlElement In xmlEmploymentList
        'BMIDS00917 MDC 16/11/2002 - Include any employment with an Accountant
        'strEmploymentStatus = xmlElement.selectSingleNode(".//EMPLOYMENTSTATUS").Text
        'If strEmploymentStatus = "20" Or strEmploymentStatus = "30" Then
        'BMIDS00917 MDC 16/11/2002 - End
            Set xmlThirdParty = xmlOut.createElement("THIRDPARTY")
            'Append EmploymentSequenceNumber to Response
            Set xmlElem2 = xmlOut.createElement("CONTEXT")
            xmlElem2.Text = xmlElement.selectSingleNode("EMPLOYMENTSEQUENCENUMBER").Text
            xmlThirdParty.appendChild xmlElem2
            
            ' get ThirdParty.CompanyName/NameAndAddressDirectory.CompanyName and add to response
            Set xmlElem2 = xmlOut.createElement("NAME")
            
            ' PSC 11/12/01 SYS3361 - Start
            xmlAccountantGuid.Text = m_objXmlAssist.GetNodeValue(xmlElement, ".//ACCOUNTANTGUID")
            strResponse = objCEBO.GetAccountantDetails(xmlAccountantRequest.xml)
            m_objErrAssist.CheckXMLResponse strResponse, True
            Set xmlAccountantResponse = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            xmlElem2.Text = m_objXmlAssist.GetNodeValue(xmlAccountantResponse, ".//COMPANYNAME")
            ' PSC 11/12/01 SYS3361 - End
            
            
            xmlThirdParty.appendChild xmlElem2
            
            xmlResponseElem.appendChild xmlThirdParty
        'End If     'BMIDS00917 MDC 16/11/2002
    Next xmlElement
        
    Set GetAccountantDetailsForCustomer = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
       
GetVbExit:
    Set objCEBO = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlListElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlAccountantRequest = Nothing
    Set xmlAccountant = Nothing
    Set xmlAccountantGuid = Nothing
    Set xmlNode = Nothing
    Set xmlAccountantResponse = Nothing

    Exit Function

GetVbErr:
    
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrNo = 0 Or lngErrNo = oeRecordNotFound Then
        Resume Next
    End If

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set GetAccountantDetailsForCustomer = m_objErrAssist.CreateErrorResponseEx()
    Resume GetVbExit

End Function

Private Function GetMortgageAccountDetailsForCustomer( _
                    ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMElement

On Error GoTo GetVbErr
    
    Const cstrFunctionName As String = "GetMortgageAccountDetailsForCustomer"

    Dim xmlOut As New FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlNewRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlTableListElem As IXMLDOMElement, xmlElement As IXMLDOMElement
    Dim xmlElem2 As IXMLDOMElement
        
    Dim strResponse As String, lngErrNo As Long
   
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Build Request and call MortgageAccountSummary
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTableElem = xmlOut.createElement("FINANCIALSUMMARY")
    xmlNewRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    Set xmlTableListElem = xmlOut.createElement("MORTGAGEACCOUNTLIST")
    xmlNewRequest.appendChild xmlTableListElem
    Set xmlTableElem = xmlOut.createElement("MORTGAGEACCOUNT")
    xmlTableListElem.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
    
    strResponse = FindMortgageAccountSummary(xmlNewRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    For Each xmlElement In xmlOut.selectNodes("//MORTGAGEACCOUNT")
        Set xmlTableElem = xmlOut.createElement("THIRDPARTY")
        Set xmlElem2 = xmlOut.createElement("NAME")
        xmlElem2.Text = m_objXmlAssist.GetTagValue(xmlElement, "COMPANYNAME")
        xmlTableElem.appendChild xmlElem2
        
        Set xmlElem2 = xmlOut.createElement("CONTEXT")
        xmlElem2.Text = m_objXmlAssist.GetTagValue(xmlElement, "ACCOUNTGUID")
        xmlTableElem.appendChild xmlElem2
                
        xmlResponseElem.appendChild xmlTableElem
    Next xmlElement
    
    Set GetMortgageAccountDetailsForCustomer = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
       
GetVbExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableListElem = Nothing
    Set xmlElement = Nothing
    Set xmlElem2 = Nothing
        
    Exit Function
GetVbErr:
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrNo = 0 Or lngErrNo = oeRecordNotFound Then
        Resume Next
    End If

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set GetMortgageAccountDetailsForCustomer = m_objErrAssist.CreateErrorResponseEx()
    Resume GetVbExit
End Function

Private Function GetLandLordDetailsForCustomer( _
                    ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMElement

On Error GoTo GetVbErr
    
    Const cstrFunctionName As String = "GetLandLordDetailsForCustomer"

    Dim xmlOut As New FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement, xmlNewRequest As IXMLDOMElement
    Dim xmlCustomerAddressList As IXMLDOMNodeList, xmlCustomerAddressElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement, xmlElem2 As IXMLDOMElement
    Dim xmlThirdParty As IXMLDOMElement
    
    Dim strResponse As String, lngErrNo As Long
        
    Dim objCustomerBO As CustomerBO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If Not m_objContext Is Nothing Then
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    Else
        Set objCustomerBO = New CustomerBO
    End If

    'Build Request and find customer address list
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTableElem = xmlOut.createElement("CUSTOMERADDRESS")
    xmlNewRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode("//CUSTOMERNUMBER")
    xmlTableElem.appendChild vXMLRequest.selectSingleNode("//CUSTOMERVERSIONNUMBER")
    
    strResponse = objCustomerBO.FindCustomerAddressList(xmlNewRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    ' Get tenancy data
    Dim strNatureOfOccupancy As String
    
    Set xmlCustomerAddressList = xmlOut.selectNodes("//CUSTOMERADDRESS")
    For Each xmlCustomerAddressElem In xmlCustomerAddressList
        strNatureOfOccupancy = xmlCustomerAddressElem.selectSingleNode(".//NATUREOFOCCUPANCY").Text
        
        If strNatureOfOccupancy = "2" Or strNatureOfOccupancy = "3" Then
            Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
            
            Set xmlTableElem = xmlOut.createElement("TENANCY")
            xmlTableElem.appendChild xmlCustomerAddressElem.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
            xmlTableElem.appendChild xmlCustomerAddressElem.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
            xmlTableElem.appendChild xmlCustomerAddressElem.selectSingleNode(".//CUSTOMERADDRESSSEQUENCENUMBER").cloneNode(True)
            
            xmlNewRequest.appendChild xmlTableElem
            strResponse = objCustomerBO.GetTenancy(xmlNewRequest.xml)
            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
            Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
            
            Set xmlTableElem = xmlOut.selectSingleNode("//TENANCY")
            If Not xmlTableElem Is Nothing Then
                Set xmlThirdParty = xmlOut.createElement("THIRDPARTY")
                xmlResponseElem.appendChild xmlThirdParty
                
                Set xmlElem2 = xmlOut.createElement("NAME")
                xmlElem2.Text = xmlTableElem.selectSingleNode(".//COMPANYNAME").Text
                xmlThirdParty.appendChild xmlElem2
                
                Set xmlElem2 = xmlOut.createElement("CONTEXT")
                xmlElem2.Text = xmlTableElem.selectSingleNode(".//CUSTOMERADDRESSSEQUENCENUMBER").Text
                xmlThirdParty.appendChild xmlElem2
            End If
        End If
    Next xmlCustomerAddressElem
    
    Set GetLandLordDetailsForCustomer = xmlResponseElem

GetVbExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCustomerAddressList = Nothing
    Set xmlCustomerAddressElem = Nothing
    Set xmlElement = Nothing
    Set xmlElem2 = Nothing
    Set xmlThirdParty = Nothing
        
    Set objCustomerBO = Nothing
    
    Exit Function
GetVbErr:
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement, True)
    If lngErrNo = 0 Or lngErrNo = oeRecordNotFound Then
        Resume Next
    End If

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set GetLandLordDetailsForCustomer = m_objErrAssist.CreateErrorResponseEx()
    Resume GetVbExit
End Function

Private Function GetTPCompanyName(ByVal vXMLRequest As IXMLDOMElement, _
                                  ByVal xmlEmploymentElem As IXMLDOMElement) As String
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlNewRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement, xmlElement As IXMLDOMElement
    
    Dim strTPGuid As String, strResponse As String, strCompanyName As String
    
    Dim objTPBO As ThirdPartyBO
    Set objTPBO = m_objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    strTPGuid = m_objXmlAssist.GetTagValue(xmlEmploymentElem, "THIRDPARTYGUID")
    
    If Len(strTPGuid) > 0 Then
        Set xmlTableElem = xmlOut.createElement("THIRDPARTY")
        xmlNewRequest.appendChild xmlTableElem
        
        Set xmlElement = xmlOut.createElement("THIRDPARTYGUID")
        xmlElement.Text = strTPGuid
        xmlTableElem.appendChild xmlElement
    Else
        strTPGuid = m_objXmlAssist.GetTagValue(xmlEmploymentElem, "DIRECTORYGUID")
        
        Set xmlTableElem = xmlOut.createElement("NAMEANDADDRESSDIRECTORY")
        xmlNewRequest.appendChild xmlTableElem
        
        Set xmlElement = xmlOut.createElement("DIRECTORYGUID")
        xmlElement.Text = strTPGuid
        xmlTableElem.appendChild xmlElement
    End If
    
    strResponse = objTPBO.GetThirdParty(xmlNewRequest.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), "GetTPCompanyName")
    
    Dim lngErrNo As Long
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlOut.documentElement)
    If lngErrNo <> 0 Then
        m_objErrAssist.ReRaiseResponseError xmlOut.documentElement
    End If
    
    strCompanyName = xmlOut.selectSingleNode(".//COMPANYNAME").Text
    
    GetTPCompanyName = strCompanyName
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlOut = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set objTPBO = Nothing
End Function

'BM0376 GHun 10/03/2003
Public Function RunIncomeCalcs(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Call RunIncomeCalculation for active quotes
' pass:         vstrXmlRequest  xml Request data stream containing applicationnumber
' return:                       xml Response data stream containing results of operation
'                               either TYPE="SUCCESS" or TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "RunIncomeCalcs"
    Dim xmlRequestDoc           As FreeThreadedDOMDocument40
    Dim xmlResponseDoc          As FreeThreadedDOMDocument40
    Dim xmlApplicationElement   As IXMLDOMElement
    Dim xmlRequestElem          As IXMLDOMElement
    Dim xmlTempElem             As IXMLDOMElement
    Dim xmlQuoteData            As IXMLDOMNode
    Dim xmlCustomerList         As IXMLDOMNode
    Dim xmlCustomer             As IXMLDOMNode
    Dim xmlNode                 As IXMLDOMNode
    Dim objApplicationQuoteBO   As IApplicationQuoteBO
    Dim objIncomeCalcsBO        As IIncomeCalcsBO
    Dim objApplicationManagerBO As IApplicationManagerBO

On Error GoTo RunIncomeCalcsVbErr

    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Set xmlRequestDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
    Set xmlApplicationElement = m_objXmlAssist.GetMandatoryNode(xmlRequestDoc.documentElement, _
            "APPLICATION[(APPLICATIONNUMBER != '') and (APPLICATIONFACTFINDNUMBER != '')]")

    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlTempElem = xmlResponseDoc.createElement("RESPONSE")
    xmlTempElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseDoc.appendChild xmlTempElem

    If Not m_objContext Is Nothing Then
        Set objApplicationQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
        Set objIncomeCalcsBO = m_objContext.CreateInstance(gstrINCOMECALCS & ".IncomeCalcsBO")
        Set objApplicationManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    Else
        Set objApplicationQuoteBO = New ApplicationQuoteBO
        Set objIncomeCalcsBO = New IncomeCalcsBO
        Set objApplicationManagerBO = New ApplicationManagerBO
    End If

    'Check if there are any Accepted or Active quotes for the application
    Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(xmlRequestDoc.documentElement)
    xmlRequestElem.appendChild xmlApplicationElement.cloneNode(True)
'    Set xmlNode = xmlRequestDoc.createElement("APPLICATION")
'    xmlRequestElem.appendChild xmlNode
'    m_objXmlAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", _
'                                            xmlApplicationElement, xmlNode, False
'    m_objXmlAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", _
'                                            xmlApplicationElement, xmlNode, False
        
    Set xmlQuoteData = objApplicationQuoteBO.GetAcceptedOrActiveQuoteData(xmlRequestElem)
    
    m_objErrAssist.CheckResponse xmlQuoteData, xmlResponseDoc.documentElement, False
    
    'Find CustomerRole data for the application
    m_objXmlAssist.ChangeNodeName xmlRequestElem, "APPLICATION", "CUSTOMERROLE"
    Set xmlCustomerList = objApplicationManagerBO.FindCustomerRoleList(xmlRequestElem)
    m_objErrAssist.CheckResponse xmlCustomerList, xmlResponseDoc.documentElement, False
    
    'Run Income Calcs
    Set xmlRequestElem = xmlRequestDoc.documentElement.cloneNode(False)
    Set xmlTempElem = xmlRequestDoc.createElement("INCOMECALCULATION")
    xmlRequestElem.appendChild xmlTempElem
    
    If xmlQuoteData.selectSingleNode(".//MORTGAGESUBQUOTE") Is Nothing Then
        xmlTempElem.setAttribute "CALCULATEMAXBORROWING", "0"
    Else
        xmlTempElem.setAttribute "CALCULATEMAXBORROWING", "1"
    End If
    
    'Append Application Node
    xmlTempElem.appendChild m_objXmlAssist.GetNode(xmlRequestDoc.documentElement, "APPLICATION", True)
    'Append CustomerList node
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomerList, "//CUSTOMERROLELIST")
    m_objXmlAssist.ChangeNodeName xmlNode, "CUSTOMERROLELIST", "CUSTOMERLIST"
    For Each xmlCustomer In xmlNode.selectNodes("CUSTOMERROLE")
        m_objXmlAssist.ChangeNodeName xmlCustomer, "CUSTOMERROLE", "CUSTOMER"
    Next
    xmlTempElem.appendChild xmlNode
    
    Set xmlTempElem = objIncomeCalcsBO.RunIncomeCalculation(xmlRequestElem)
    m_objErrAssist.CheckResponse xmlTempElem, xmlResponseDoc.documentElement, True
    
    RunIncomeCalcs = xmlResponseDoc.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RunIncomeCalcsExit:
    Set objApplicationQuoteBO = Nothing
    Set objIncomeCalcsBO = Nothing
    Set objApplicationManagerBO = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlApplicationElement = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTempElem = Nothing
    Set xmlQuoteData = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomer = Nothing
    Set xmlNode = Nothing
    Exit Function

RunIncomeCalcsVbErr:
    m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    m_objErrAssist.AddToErrSource cstrFunctionName

    RunIncomeCalcs = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    '   go to clean-up section
    Resume RunIncomeCalcsExit
End Function
'BM0376 End


Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub

Private Function ToUpperAlphaOnly(ByVal vstrIn As String) As String
' header ----------------------------------------------------------------------------------
' procedure:    ToUpperAlphaOnly
' author:       Dave Crossley
' date:         12/02/2004
' description:  Helper function to remove any non alpha characters & return
'               result in uppercase
' pass:         Raw string
' return:       String containing only upper case letters
' errors:       If an error found, raise it
'------------------------------------------------------------------------------------------
On Error GoTo ToUpperAlphaOnlyErr

Const strFunctionName As String = "ToUpperAlphaOnly"

Dim strOut As String
strOut = ""
Dim intIndex As Integer
intIndex = 1
Dim strChar As String
Dim intTextLen, intAscValue As Integer
intTextLen = Len(vstrIn)

   
    While intIndex <= intTextLen
       strChar = Mid$(vstrIn, intIndex, 1)
       intAscValue = Asc(strChar)
       If (intAscValue >= 65 And intAscValue <= 90) Then
         ' uppercase
          strOut = strOut & strChar
       Else
          If (intAscValue >= 97 And intAscValue <= 122) Then
          ' lowercase
           strOut = strOut & UCase$(strChar)
          End If
       End If
       intIndex = intIndex + 1
    
    Wend
       

  ToUpperAlphaOnly = strOut
  Exit Function
    
ToUpperAlphaOnlyErr:
    
    Err.Raise Err.Number, Err.Source, Err.Description

  
End Function





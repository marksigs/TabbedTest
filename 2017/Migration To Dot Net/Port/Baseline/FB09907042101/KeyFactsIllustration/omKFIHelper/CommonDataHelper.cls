VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         CommonDataHelper
'** Created by:     Andy Maggs
'** Date:           14/04/2004
'** Description:    Provides access to data that is required across more than
'**                 one section of the KFI documents
'********************************************************************************
'Prog   Date        Description
'BC     05/10/2005  MAR88 - Project MARS
'TW     11/11/2005  MAR525 - Error when establishing type of Application
'BC     16/12/2005  MAR907 - Remortgages, ToE and Product Switch
'BC     02/03/2006  MAr1347  Use first Validation Type fro RepaymentMethod, rather the second
'********************************************************************************
'Prog   Date        Description
'PB     15/05/2006  EP529 / MAR1738 Sec 12 - Payment holidays allowed until - date does not tie up with product expiry date
'AW     22/05/2006  EP590 Amendments for Section 8 (Fees yo will pay)
'PM     23/05/2006  EP584 New subroutine SetBrokerAndPackager. New PropertyGet for BrokerName and PackagerName.
'                   Updates to Initialise and IsIntroducedByIntermediary
'PB     24/05/2006  EP603 / MAR1777 Changed SetApplicationFactFindBasedProperties to default ExpectedCompletionDate correctly
'PB     25/05/2006  EP603 / MAR1788 Retrieve APR and AmountPerUnitBorrowedfrom MortgageSubQuote
'PB     05/06/2006  EP651 / MAR1284 KFI TEMPLATE NEEDS CHANGES FOR TRANSFER OF EQUITY (CHARTER 1)
'PB     06/06/2006  EP651 / MAR1590 KFI Part and Part printing issue
'PB     15/06/2006  EP772 Now takes an average of APR and Amount Per Unit
'PB     24/08/2006  EP1082  Updated to give correct 'Other' title in sect 2 & 9 in offer doc
'PE     04/09/2006  EP1113  CC118 - Valuation moved from ‘fees payable to db mortgages’ section to ‘other fees’ section
'AW     06/10/2006  EP1187  Rework of CC118
'AW     12/10/2006  EP1212  All variations on Valuation Fees to appear in Other Fees Section for Phase 1
'LH     13/10/2006  EP1226 / CC132 - Added lender code
'********************************************************************************
'Prog   Date        Description
'PB     21/11/2006  EP2_139 Added Transfer of Equity to SetApplicationFactFindBasedProperties
'EP2_139 30/11/2006 Offer and KFI document changes..
' INR   12/12/2006  EP2_422   Section13 changes
' INR   25/01/2007  EP2_704 Compare against INTRODUCERFEETYPE validationtype
' INR   01/02/2007  EP2_718 Should be FREELEGALFEES
' INR   02/02/2007  EP2_704 Need PackageAssoc for assembling IntermediaryNames
' INR   19/02/2007  EP2_1464 Need to set MainMortgageTypeGroup for TOE
' INR   21/02/2007  EP2_1450 AdditionalBorrowing =>CheckForValidationType "F" not "F," & new mortgagetypetext
' INR   28/02/2007  EP2_1449 More Additional Borrowing stuff
' INR   06/03/2007  EP2_1667 AddCustomerAddedRemovedData required by Offer & KFI, so moved here
' INR   06/03/2007  EP2_1530/EP2_1667 Use new validationtype to identify TypeOfApplication for AddInsuranceLoanPurposeElement
' PB    07/03/2007  EP2_1861 Added IsRefundOfValuation
' PB    20/03/2007  EP2_1861 Corrected bug - reading wrong attribute
' INR   21/03/2007  EP2_1753 Match against the whole validationtype when excluding mortgageoneoffcosts
' INR   25/03/2007  EP2_1753 If we have a legal fee with an amount, we need it for SOLFEEAMOUNT in KFIDataHelper
' PE    27/03/2007  EP2_1994 / EP2_1992 / EP2_1990 / EP2_1979
' INR   28/03/2007  EP2_1994 PrincipalFirmName would not be assigned to BrokerName
' TW    09/04/2007  EP2_1743 - kfi - not appearing correctly in a case with drawdown
' SR    11/04/2007  EP2_2270 modified function SetBrokerAndPackager
' INR   13/04/2007  EP2_2395 fixes to SetBrokerAndPackager
' INR   15/04/2007  EP2_2395 changed CheckForValidationType for "REPAYMENTMETHOD"
' INR   17/04/2007  EP2_2448 new functions for specialscheme
' GHun  20/04/2007  EP2_2508 changed SetMortgageSubQuoteProperties to use correct APR and AMOUNTPERUNITBORROWED
' INR   20/04/2007  EP2_2478 New check for additional borrowing and/or TOE
'********************************************************************************
Option Explicit
Option Compare Text
Public Enum MortgageInterestRateType
    mrtStandardVariableRate = 1
    mrtFixedRate = 2
    mrtDiscountedRate = 3
    mrtCappedRate = 4
    mrtCollaredRate = 5
    mrtCappedAndCollaredRate = 6
    mrtTrackerAbove = 7
End Enum

Public Enum enumRepaymentType
    enumInterestOnly = 0
    enumCapitalAndInterest
End Enum

Public blnFixedRateMortgage As Boolean
Public blnIsTransferofEquity As Boolean ' BC MAR907
Public blnIsProductSwitch As Boolean ' BC MAR907
Public blnFixedforWholeMortgage As Boolean ' PB EP529 / MAR1738
' PB 07/06/2006 EP651/MAR1590 Begin
Public mstrMortgageRepaymentTypeElement As String
Public mlngTotalCapitalandInterest As Double
Public mlngTotalInterestOnly As Double
' EP651/MAR1590 End

Public Enum FeePremiumType
    fptNone = 1
    fptAllPaid = 2
    fptPartPart = 3
    fptAllAdded = 4
End Enum

Private mxmlDoc As FreeThreadedDOMDocument40
Private mxmlData As IXMLDOMNode
' PB 07/06/2006 EP651/MAR1590
'Private mstrMortgageRepaymentTypeElement As String
' EP561/MAR1590 End
Private mstrLoanPurposeElement As String
Private mstrComponentsTypeElement As String
Private mblnHasInsurances As Boolean
Private mblnIsVariableRate As Boolean
Private mblnIsAdditionalBorrowing As Boolean
Private mblnIsRemortgage As Boolean ' BC MAR907
Private mlngLoanAmount As Long
Private mlngTotalLoanAmount As Long
Private mdblLTV As Double
' PB 25/05/2006 EP603/MAR1788
Private mdblAPR As Double
Private mdblAmountPerPound As Double
' EP603/MAR1788 End
Private mlngLoanAmountRepayable As Long
Private mlngFeesAddedToLoanAmount As Long
Private mlngTotalFees As Long
Private meFeesType As FeePremiumType
Private mdblPremiumsAddedToLoanAmount As Double
Private mdblTotalInsurancePremiums As Double
Private meInsurancePremiumType As FeePremiumType
Private mdtExpectedCompletionDate As Date
Private mstrProvider As String
Private mstrProviderCode As String 'LH 13/10/2006 EP1226 / CC132
Private mlngMortgageCompletionFee As Long
Private mstrMainMortgageTypeGroup As String
Private mastrFeesAddedToLoan() As String  'Array of the fee amounts that are added to loan
Private mintTermInYears As Integer
Private mintTermInMonths As Integer
Private mblnIsMultiComponent As Boolean
'SR 10/09/2004 : CORE82
Private mblnIsLifeTimeDocGroup As Boolean 'TRUE for all LifeTime docs else FALSE
Private mstrMortgageTypeText As String
Private mstrMortgageTypeText2 As String
Private mstrMortgageTypeText4 As String
'SR 10/09/2004 : CORE82 - End
Private mstrMortgageTypeText_A As String 'SR 14/10/2004 : BBG1596
Private mdblStandardVariableRate As Double  'SR 12/09/2004 : CORE82
Private mdblAmountPerUnitBorrowed As Double  'INR 17/09/2004 : CORE82
'SR 21/09/2004 : CORE82
Private mblnIntroducedByIntermediary As Boolean
Private mstrIntermediaryContactName As String
Private mstrIntermediaryOrganisationName As String
'SR 21/09/2004 : CORE82 - End
Private mblnPackagedApplication As Boolean  'SR 07/10/2004 : BBG1545
'IK 01/12/2004 E2EM00003125, E2EM00003126
Private mblnHasDrawDown As Boolean
Private mlngDrawDown As Long
Private mlngAmountRequestedLessDrawDown As Long
Private mlngTotalLoanAmountLessDrawDown As Long
'IK 01/12/2004 E2EM00003125, E2EM00003126 ends
' TW 09/04/2007 EP2_1743
Private mdblMonthlyCostLessDrawdown As Double
' TW 09/04/2007 EP2_1743 End

'PM 23/05/2006 EP584 Start
Private mstrBrokerName      As String
Private mstrPackagerName    As String
'PM 23/05/2006 EP584 End
'EP2_139
Private mstrIntroducerName      As String
Private mblnIndividualIntroducer As Boolean
Private mblnIsFlexible As Boolean
Private mblnIsIncentive As Boolean
Private mblnIsCashback As Boolean
Private mblnIsFreeLegal As Boolean
Private mblnIsFreevaluation As Boolean
Private mdblTotalNetMonthlyCost As Double
Private mFreeValuationAmount As Double
Private mstrIntermediaryNames      As String
Private mstrFeePayableToThirdParty      As String
Private mblnIsProcFeeRefund As Boolean
Private mProcFeeRefundAmount As Long
Private mblnIsRefundOfValuation As Boolean 'EP2_1861
Private mstrSpecialSchemeGroupType As String 'EP2_2448
Private mstrSpecialSchemeValidation As String
Private mblnIsAdditionalBorrowingTOE As Boolean 'EP2_2478


Private Sub Class_Initialize()
    '
    'PB 09/03/2007 EP2_1627
    gblnSection2 = False
    '
End Sub

Private Sub Class_Terminate()
    Set mxmlDoc = Nothing
    Set mxmlData = Nothing
End Sub

'********************************************************************************
'** Function:       Initialise
'** Created by:     Andy Maggs
'** Date:           14/04/2004
'** Description:    Initialises the class with the data it is wrapping. It parses
'**                 the data to create the elements that are required in multiple
'**                 sections within the KFI document.
'** Parameters:     vxmlDoc - the return KFI document that is used in this class
'**                 to create new elements with.
'**                 vxmlDataNode - the input data for the document.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub Initialise(ByVal vxmlDataNode As IXMLDOMNode, ByVal strDocGroup As String)
    Const cstrFunctionName As String = "Initialise"

    On Error GoTo ErrHandler

    '*-create a new document and store the data
    Set mxmlDoc = New FreeThreadedDOMDocument40
    mxmlDoc.async = False
    Set mxmlData = vxmlDataNode
    
    Call SetDocGroup(strDocGroup)  'SR 10/09/2004 : CORE82
    Call SetIsAdditionalBorrowing
    Call SetMortgageTypeText 'SR 10/09/2004 : CORE82
    Call SetApplicationFactFindBasedProperties
    Call SetLoanComponentBasedProperties
    Call SetHasInsurances
    Call SetMortgageSubQuoteProperties
    Call SetFeesInfo
    Call SetInsurancePremiumInfo
    Call SetMortgageLenderProperties
'PM 23/05/2006 EP584 Start
    Call SetBrokerAndPackager
    'EP2_139
    mblnIntroducedByIntermediary = IsIntroducedByIntermediary(mstrIntermediaryContactName, True, mstrIntermediaryOrganisationName, True)  'SR 21/09/2004 : CORE82
    'EP2_139
    Call SetFlexible
    'Set this to false, then it may be set true with SetCashback, SetFreeLegal or SetFreevaluation
    mblnIsIncentive = False
    Call SetCashback
    Call SetFreeLegal
    Call SetFreeValuation
    'EP2_1861
    Call SetRefundOfValuation

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Public Property Get Document() As FreeThreadedDOMDocument40
    Set Document = mxmlDoc
End Property

Public Property Get Data() As IXMLDOMNode
    Set Data = mxmlData
End Property

Public Property Get BrokerName() As String
    BrokerName = mstrBrokerName
End Property

Public Property Get PackagerName() As String
    PackagerName = mstrPackagerName
End Property
'EP2_139
Public Property Get IntroducerName() As String
    IntroducerName = mstrIntroducerName
End Property
'EP2_139
Public Property Get IntermediaryNames() As String
    IntermediaryNames = mstrIntermediaryNames
End Property
'EP2_139
Public Property Get IsProcFeeRefund() As Boolean
    IsProcFeeRefund = mblnIsProcFeeRefund
End Property
'EP2_139
Public Property Get ProcFeeRefundAmount() As String
    ProcFeeRefundAmount = mProcFeeRefundAmount
End Property
'EP2_139
Public Property Get FeePayableToThirdPartyAmount() As String
    FeePayableToThirdPartyAmount = mstrFeePayableToThirdParty
End Property
Public Property Get HasInsurances() As Boolean
    HasInsurances = mblnHasInsurances
End Property

Public Property Get IsVariableRate() As Boolean
    IsVariableRate = mblnIsVariableRate
End Property

Public Property Get MortgageTypeText() As String
    MortgageTypeText = mstrMortgageTypeText
End Property
'SR 10/09/2004 : CORE82
Public Property Get MortgageTypeText2() As String
    MortgageTypeText2 = mstrMortgageTypeText2
End Property
'SR 10/09/2004 : CORE82 - End

'SR 13/10/2004 : BBG1596
Public Property Get MortgageTypeText_A() As String
    MortgageTypeText_A = mstrMortgageTypeText_A
End Property
'SR 13/10/2004 : BBG1596 - End
'EP2_1450
Public Property Get MortgageTypeText4() As String
    MortgageTypeText4 = mstrMortgageTypeText4
End Property

Public Property Get IsAdditionalBorrowing() As Boolean
    IsAdditionalBorrowing = mblnIsAdditionalBorrowing
End Property

Public Property Get IsAdditionalBorrowingTOE() As Boolean
    IsAdditionalBorrowingTOE = mblnIsAdditionalBorrowingTOE
End Property

Public Property Get IsProductSwitch() As Boolean
    IsProductSwitch = blnIsProductSwitch
End Property

Public Property Get IsTransferOfEquity() As Boolean
    IsTransferOfEquity = blnIsTransferofEquity
End Property

Public Property Get LoanAmount() As Long
    LoanAmount = mlngLoanAmount
End Property

Public Property Get TotalLoanAmount() As Long
    TotalLoanAmount = mlngTotalLoanAmount
End Property
'EP2_139
Public Property Get TotalNetMonthlyCost() As Double
    TotalNetMonthlyCost = mdblTotalNetMonthlyCost
End Property

Public Property Get LTV() As Double
    LTV = mdblLTV
End Property

'PB EP603/MAR1788
Public Property Get APR() As Double
    APR = mdblAPR
End Property
Public Property Get AmountPerPound() As Double
    AmountPerPound = mdblAmountPerPound
End Property
' EP603/MAR1788 End

Public Property Get LoanAmountRepayable() As Long
    LoanAmountRepayable = mlngLoanAmountRepayable
End Property

Public Property Get FeesAddedToLoanAmount() As Long
    FeesAddedToLoanAmount = mlngFeesAddedToLoanAmount
End Property


Public Property Get ArrayOfFeesAddedToLoan() As String()
    ArrayOfFeesAddedToLoan = mastrFeesAddedToLoan
End Property

Public Property Get TotalFees() As Long
    TotalFees = mlngTotalFees
End Property

Public Property Get MortgageCompletionFee() As Long
    MortgageCompletionFee = mlngMortgageCompletionFee
End Property

Public Property Get FeesType() As FeePremiumType
    FeesType = meFeesType
End Property

Public Property Get PremiumsAddedToLoanAmount() As Double
    PremiumsAddedToLoanAmount = mdblPremiumsAddedToLoanAmount
End Property

Public Property Get TotalInsurancePremiums() As Double
    TotalInsurancePremiums = mdblTotalInsurancePremiums
End Property

Public Property Get InsurancePremiumType() As FeePremiumType
    InsurancePremiumType = meInsurancePremiumType
End Property

Public Property Get ExpectedCompletionDate() As Date
    ExpectedCompletionDate = mdtExpectedCompletionDate
End Property

Public Property Get LoanPurposeText() As String
    LoanPurposeText = mstrLoanPurposeElement
End Property
'SR 07/10/2004 : BBG1545
Public Property Get IsApplicationPackaged() As Boolean
    IsApplicationPackaged = mblnPackagedApplication
End Property
'SR 07/10/2004 : BBG1545 - End

'IK 01/12/2004 E2EM00003125, E2EM00003126
Public Property Get HasDrawDown() As Boolean
    HasDrawDown = mblnHasDrawDown
End Property

Public Property Get DrawDownAmount() As Long
    DrawDownAmount = mlngDrawDown
End Property

Public Property Get AmountRequestedLessDrawDown() As Long
    AmountRequestedLessDrawDown = mlngAmountRequestedLessDrawDown
End Property

Public Property Get MonthlyCostLessDrawDown() As Long
    MonthlyCostLessDrawDown = mdblMonthlyCostLessDrawdown
End Property
Public Property Get TotalLoanAmountLessDrawDown() As Long
    TotalLoanAmountLessDrawDown = mlngTotalLoanAmountLessDrawDown
End Property

'IK 01/12/2004 E2EM00003125, E2EM00003126 ends

'********************************************************************************
'** Function:       AddMortgageRepaymentTypeElement
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Adds the appropriate mortgage type element to the supplied
'**                 node and returns it.
'** Parameters:     vxmlNode - the node to create the mortgage type element on.
'** Returns:        The newly created node.
'** Errors:         None Expected
'********************************************************************************
Public Function AddMortgageRepaymentTypeElement(ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "AddMortgageRepaymentTypeElement"
    Dim xmlItem As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    'EP2_1449
    Set xmlItem = CreateNewElement(mstrMortgageRepaymentTypeElement, vxmlNode)
    Call xmlSetAttributeValue(xmlItem, "MORTGAGETYPE", MortgageTypeText())

    Set AddMortgageRepaymentTypeElement = xmlItem

    Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       AddLoanPurposeElement
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Adds the loan purpose element to the specified node.
'** Parameters:     vxmlNode - the node to add the loan purpose element to.
'** Returns:        The new loan purpose node.
'** Errors:         None Expected
'********************************************************************************
Public Function AddLoanPurposeElement(ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "AddLoanPurposeElement"

    On Error GoTo ErrHandler

    Set AddLoanPurposeElement = CreateNewElement(mstrLoanPurposeElement, vxmlNode)

    Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       AddComponentsTypeElement
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Adds the components type element (i.e. SINGLECOMPONENT or
'**                 MULTICOMPONENT).
'** Parameters:     vxmlNode - the node to add the components type element to.
'** Returns:        The newly created node.
'** Errors:         None Expected
'********************************************************************************
Public Function AddComponentsTypeElement(ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "AddComponentsTypeElement"

    On Error GoTo ErrHandler

    Set AddComponentsTypeElement = CreateNewElement(mstrComponentsTypeElement, vxmlNode)

    Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       LoanComponents
'** Created by:     Andy Maggs
'** Date:           05/04/2004
'** Description:    Gets and returns the collection of LoanComponent nodes.
'** Parameters:     None.
'** Returns:        The collection of LoanComponent nodes.
'** Errors:         None Expected
'********************************************************************************
Public Property Get LoanComponents() As IXMLDOMNodeList
    Set LoanComponents = mxmlData.selectNodes(gcstrLOANCOMPONENT_PATH)
End Property

Public Property Get SingleLoanComponent() As IXMLDOMNode
    Set SingleLoanComponent = mxmlData.selectSingleNode(gcstrLOANCOMPONENT_PATH)
End Property
Public Property Get Provider() As String
    Provider = mstrProvider
End Property
'LH 13/10/2006 EP1226 / CC132: start
Public Property Get ProviderCode() As String
    ProviderCode = mstrProviderCode
End Property
'LH 13/10/2006 EP1226 / CC132: end

'SR 21/09/2004 : CORE82
Public Property Get IntroducedByIntermediaryProperty() As Boolean
    IntroducedByIntermediaryProperty = mblnIntroducedByIntermediary
End Property
'SR 21/09/2004 : CORE82 - End

'SR 21/09/2004 : CORE82
Public Property Get IntermediaryContactName() As String
    IntermediaryContactName = mstrIntermediaryContactName
End Property
'SR 21/09/2004 : CORE82 - End

'SR 21/09/2004 : CORE82
Public Property Get IntermediaryOrganisationName() As String
    IntermediaryOrganisationName = mstrIntermediaryOrganisationName
End Property
'SR 21/09/2004 : CORE82 - End
'EP2_139
Public Property Get IndividualIntroducerProperty() As Boolean
    IndividualIntroducerProperty = mblnIndividualIntroducer
End Property


'********************************************************************************
'** Function:       GetLoanComponentInterestRates
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Gets and returns all the interest rate records for the
'**                 specified loan component.
'** Parameters:     vxmlLoanComponentNode - the loan component node.
'** Returns:        The list of interest rate records.
'** Errors:         None Expected
'********************************************************************************
Public Function GetLoanComponentInterestRates( _
        ByVal vxmlLoanComponentNode As IXMLDOMNode) As IXMLDOMNodeList
    Const cstrFunctionName As String = "GetLoanComponentInterestRates"
    
    On Error GoTo ErrHandler
    
    Set GetLoanComponentInterestRates = vxmlLoanComponentNode.selectNodes( _
            ".//INTERESTRATETYPE")

    Exit Function
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       GetLoanComponentFirstInterestRate
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Gets the first interest rate record from the loan component.
'** Parameters:     vxmlLoanComponentNode - the loan component node to get it
'**                 from.
'** Returns:        The first interest rate record node.
'** Errors:         None Expected
'********************************************************************************
Public Function GetLoanComponentFirstInterestRate( _
        ByVal vxmlLoanComponentNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "GetLoanComponentFirstInterestRate"
    
    On Error GoTo ErrHandler
        
    Set GetLoanComponentFirstInterestRate = _
            vxmlLoanComponentNode.selectSingleNode( _
            ".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]")
    'MAR88 - BC 05Oct
    '"//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]") is changed to
    '".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]")

    Exit Function
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       GetIntroducerFee
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Gets the fee charged by the introducer.
'** Parameters:     None
'** Returns:        The introducer fee.
'** Errors:         None Expected
'********************************************************************************
Public Function GetIntroducerFee() As String
    Const cstrFunctionName As String = "GetIntroducerFee"
    Dim xmlItem As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Set xmlItem = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    GetIntroducerFee = xmlGetAttributeText(xmlItem, "INTRODUCERFEE")

    Set xmlItem = Nothing
    Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       IsIntroducedByIntermediary
'** Created by:     Srini Rao
'** Date:           06/2004
'** Description:    Determines whether the client was introduced by an
'**                 intermediary or not and if so, optionally returns the name
'**                 of the intermediary.

'** Parameters:     rstrContactName - the returned intermediary name.
'**                 vblnNeedContactName - whether the contact name should be
'**                 returned or not
'** Returns:        True if introduced by intermediary or False if not.
'** Errors:         None Expected
'********************************************************************************
'PM 23/05/2006 EP584 Start
'EP2_139 30/11/2006  Offer and KFI document changes.
Public Function IsIntroducedByIntermediary(ByRef rstrBrokerName As String, _
        Optional ByVal vblnNeedBrokerName As Boolean = True, _
        Optional ByRef rstrPackagerName As String = "", _
        Optional ByVal vblnNeedPackager As Boolean = False) As Boolean
        
    Const cstrFunctionName As String = "IsIntroducedByIntermediary"
  
  
    On Error GoTo ErrHandler
    
    Dim xmlItem As IXMLDOMNode
    Dim xmlList As IXMLDOMNodeList
    Dim strIntermediaryType As String
    Dim strTemp As String
    Dim blnReturn As Boolean
    Dim xmlAppFF As IXMLDOMNode  'SAB 24/04/2006 - EPSOM EP417

    Dim rstrContactName As String
    Dim vblnNeedOrganisationName As Boolean
    blnReturn = False
    rstrContactName = ""
    mblnIsProcFeeRefund = False
    
    Set xmlItem = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    strTemp = xmlGetAttributeText(xmlItem, "DIRECTINDIRECTBUSINESS")

    'MAR44 - BC 12 Sep 05
    'If strTemp =  "D," Then
    If InStr(strTemp, "D,") > 0 Then
        IsIntroducedByIntermediary = blnReturn
        Exit Function
    End If

    Set xmlList = mxmlData.selectNodes(gcstrAPPINTRODUCER_PATH)
    If xmlList.length = 0 Then
        IsIntroducedByIntermediary = blnReturn 'SR 25/10/2004 : BBG1684
    Else
        'EP2_422 Have Intermediary, check if there is a refund of proc fee
        Dim ProcFeeRefundAmount As Long
        If (xmlGetAttributeAsLong(xmlItem, "BROKERPROCFEEREFUND")) Then
            ProcFeeRefundAmount = xmlGetAttributeAsLong(xmlItem, "BROKERPROCFEEREFUND")
            If ProcFeeRefundAmount > 0 Then
                mblnIsProcFeeRefund = True
                mProcFeeRefundAmount = ProcFeeRefundAmount
            End If
        End If
        blnReturn = True
    End If
    
    If vblnNeedBrokerName Then
        rstrBrokerName = mstrIntroducerName
    Else
        rstrBrokerName = ""
    End If

    If vblnNeedPackager Then
        rstrPackagerName = mstrBrokerName
    Else
        rstrPackagerName = ""
    End If

    IsIntroducedByIntermediary = blnReturn

    Set xmlItem = Nothing
    Set xmlList = Nothing
    Set xmlAppFF = Nothing      'SAB 24/04/2006 - EPSOM EP417
'PM 23/05/2006 EP584 End
    Exit Function

ErrHandler:
'PM 23/05/2006 EP584 Start
    Set xmlItem = Nothing
    Set xmlList = Nothing
    Set xmlAppFF = Nothing      'SAB 24/04/2006 - EPSOM EP417
'PM 23/05/2006 EP584 End
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'SR 21/09/2004 : CORE82
Private Function GetIntermediaryOrganisationName() As String
    
    Const cstrFunctionName As String = "GetIntermediaryOrganisationName"
    
    Dim xmlIntemediaryOrganisation As IXMLDOMNode
    Dim xmlIntemediaryOrganisationList As IXMLDOMNodeList, xmlNode As IXMLDOMNode  'SR 11/09/2004 : CORE82
    Dim strValidationType As String 'SR 11/09/2004 : CORE82
    
    On Error GoTo ErrHandler
    
    'SR 11/09/2004 : CORE82
    Set xmlIntemediaryOrganisationList = mxmlData.selectNodes(gcstrAPPLICATIONINTERMEDIARY_PATH)
    If xmlIntemediaryOrganisationList.length > 1 Then
        For Each xmlNode In xmlIntemediaryOrganisationList
            strValidationType = xmlGetAttributeText(xmlNode, "INTERMEDIARYTYPE")
            If (InStr((UCase$(strValidationType)), "ECROSD,") > 0) Then
                Set xmlIntemediaryOrganisation = xmlNode.selectSingleNode(".//INTERMEDIARY/INTERMEDIARYORGANISATION")
                Exit For
            End If
        Next xmlNode
    Else ' this should not happen. just in case the data is wrong
        Set xmlIntemediaryOrganisation = mxmlData.selectSingleNode(gcstrINTERMEDIARYORGANISATION_PATH)
    End If
    'SR 11/09/2004 : CORE82
    
    If Not xmlIntemediaryOrganisation Is Nothing Then
        GetIntermediaryOrganisationName = xmlGetAttributeText(xmlIntemediaryOrganisation, "NAME")
    Else
        GetIntermediaryOrganisationName = ""
    End If

    Set xmlIntemediaryOrganisation = Nothing
    Set xmlIntemediaryOrganisationList = Nothing
    Set xmlNode = Nothing
Exit Function
ErrHandler:
    Set xmlIntemediaryOrganisation = Nothing
    Set xmlIntemediaryOrganisationList = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'SR 21/09/2004 : CORE82 - End

'********************************************************************************
'** Function:       GetFeesPaidToProvider
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Gets the list of fees that are actually paid to the provider.
'** Parameters:     None.
'** Returns:        The list of fees.
'** Errors:         None Expected
'********************************************************************************
Public Function GetFeesPaidToProvider() As IXMLDOMNodeList
    Const cstrFunctionName As String = "GetFeesPaidToProvider"
    'CORE82 Broker Fees are not paid to the provider
    'EP2_1753 Need legal fees amount
    Const cstrExcludeList As String = "TID|STA|MIG|BRK|PRC|PAK|MKT"  'AW 22/05/2006 : EP590 - removed TPV from string
    Dim xmlList As IXMLDOMNodeList
    Dim xmlFeesList As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim strType As String
    Dim astrType() As String
    Dim intIndex As Integer
    Dim blnExclude As Boolean
    Dim xmlPackaged As IXMLDOMNode
    'AW EP1187  06/10/2006
    Dim xmlValuation As IXMLDOMNode
    Dim blnRetype As Boolean
    Dim strValuationType As String
    'EP2_1753
    Dim xmlTPVNode As IXMLDOMNode
    Dim rebateAmount As Long
    Dim feeAmount As Long
    Dim legalAmount As Long
    rebateAmount = 0
    feeAmount = 0
    legalAmount = 0
    
    On Error GoTo ErrHandler

    Set xmlFeesList = mxmlDoc.createElement("FEES")
    Set xmlList = mxmlData.selectNodes(gcstrMORTGAGEONEOFFCOSTS_PATH)
    If xmlList.length > 0 Then
        For Each xmlItem In xmlList
            strType = xmlGetAttributeText(xmlItem, "MORTGAGEONEOFFCOSTTYPE")
            
            astrType = Split(strType, ",")
            '*-ensure exclude flag is always reset
            blnExclude = False
            For intIndex = 0 To UBound(astrType)
                If Len(astrType(intIndex)) > 0 Then
                    'EP2_1753 Match against the whole validationtype
                    If InStr("|" + cstrExcludeList + "|", "|" + astrType(intIndex) + "|") > 0 Then
                        blnExclude = True
                        Exit For
                    Else
                        'EP2_1753 Won't have a TPVA without a TPV
                        If ((astrType(intIndex) = "TPV") Or (astrType(intIndex) = "TPVA")) Then
                            rebateAmount = rebateAmount + xmlGetAttributeAsLong(xmlItem, "REBATEAMOUNT", "")
                            feeAmount = feeAmount + xmlGetAttributeAsLong(xmlItem, "REFUNDAMOUNT", "")
                            'Initialise TPV value which will be set below
                            If astrType(intIndex) = "TPV" Then
                                Call xmlSetAttributeValue(xmlItem, "COMBINEDFEE", "0")
                                Call xmlSetAttributeValue(xmlItem, "COMBINEDREBATE", "0")
                            End If
                        End If
                        'EP2_1753 If we have a legal fee with an amount, we need it for SOLFEEAMOUNT
                        'in KFIDataHelper, but don't want to display it as a legal fee, so set "AMOUNT"
                        'to zero
                        If (astrType(intIndex) = "LEG") Then
                            legalAmount = xmlGetAttributeAsLong(xmlItem, "AMOUNT", "")
                            Call xmlSetAttributeValue(xmlItem, "LEGALAMOUNT", CStr(legalAmount))
                            Call xmlSetAttributeValue(xmlItem, "AMOUNT", "0")
                        End If

                        'SR 07/10/2004 : BBG1545
                        'Peter Edney - 04/08/2006 - EP1113
                        'If astrType(intIndex) = "VAL" And mblnPackagedApplication Then
                        'If (astrType(intIndex) = "VAL") And (Not mblnPackagedApplication) Then
                        'AW EP1187  12/10/2006
                        'AW     04/10/2006  EP1212
'                        If (astrType(intIndex) = "VAL") Then
'
'                            strValuationType = ""
'                            blnRetype = False
'                            Set xmlValuation = mxmlData.selectSingleNode(gcstrVALUATIONINSTRUCTIONS_PATH & "[last()]")
'                            If Not (xmlValuation Is Nothing) Then
'                                strValuationType = UCase(xmlGetAttributeText(xmlValuation, "VALUATIONTYPE_VALIDID", ""))
'                                If strValuationType <> "" Then
'                                    blnRetype = CheckForValidationType(strValuationType, "RT")
'                                End If
'                            End If
'
'                            If Not blnRetype Then
'                                Set xmlPackaged = xmlItem.ownerDocument.createElement("packaged")
'                                xmlItem.appendChild xmlPackaged
'                            End If
'                            'AW EP1187  06/10/2006 - End
'                            'blnExclude = True
'                            'Exit For
'                        End If 'SR 07/10/2004 : BBG1545 - End
                    End If
                End If
            Next intIndex
            
            If Not blnExclude Then
                '*-add this item to the fees list
                Call xmlFeesList.appendChild(xmlItem)
            End If
        Next xmlItem
    End If
    
    Set xmlTPVNode = xmlFeesList.selectSingleNode("//MORTGAGEONEOFFCOST[@COMBINEDFEE]")
    If Not xmlTPVNode Is Nothing Then
        If feeAmount > 0 Then
            Call xmlSetAttributeValue(xmlTPVNode, "COMBINEDFEE", CStr(feeAmount))
        End If
        If rebateAmount > 0 Then
            Call xmlSetAttributeValue(xmlTPVNode, "COMBINEDREBATE", CStr(rebateAmount))
        End If
    End If
    
    Set GetFeesPaidToProvider = xmlFeesList.selectNodes("//MORTGAGEONEOFFCOST")

    Set xmlList = Nothing
    Set xmlFeesList = Nothing
    Set xmlItem = Nothing
    Set xmlTPVNode = Nothing
    Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlFeesList = Nothing
    Set xmlItem = Nothing
    Set xmlTPVNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       GetLatestValnReportPropertyDetails
'** Created by:     Andy Maggs
'** Date:           24/05/2004
'** Description:    Gets the most recent valuation report property details node.
'** Parameters:     None
'** Returns:        The node if there is one.
'** Errors:         None Expected
'********************************************************************************
Public Function GetLatestValnReportPropertyDetails() As IXMLDOMNode
    Const cstrFunctionName As String = "GetLatestValnReportPropertyDetails"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlDetails As IXMLDOMNode

    On Error GoTo ErrHandler
    
    '*-get the full list of valuation report property details records
    Set xmlList = mxmlData.selectNodes(gcstrAPPLICATIONFACTFIND_PATH & "/VALNREPPROPERTYDETAILS")
    '*-and select the one with the highest instruction sequence numberF
    Set xmlDetails = xmlListHelper.GetElementWithMaxLngAttribValue(xmlList, "INSTRUCTIONSEQUENCENO")
    
    Set GetLatestValnReportPropertyDetails = xmlDetails

    Set xmlList = Nothing
    Set xmlDetails = Nothing
    Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlDetails = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'********************************************************************************
'** Function:       CreateNewElement
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Utility function to create a new element and add it to the
'**                 specified node.
'** Parameters:     vstrName - the name of the element to create.
'**                 vxmlNode - the node to add it to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Function CreateNewElement(ByVal vstrName As String, _
        ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "CreateNewElement"

    Dim xmlElem As IXMLDOMElement

    On Error GoTo ErrHandler

    Set xmlElem = mxmlDoc.createElement(vstrName)
    Set CreateNewElement = vxmlNode.appendChild(xmlElem)
    
    Set xmlElem = Nothing
    Exit Function
ErrHandler:
    Set xmlElem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'SR 10/09/2004 : CORE82
Private Sub SetDocGroup(ByVal strDocGroup As String)
    'BBG1489 Take into account NonRegulated Mortgage
    If strDocGroup = "STANDARD" Or strDocGroup = "OFFER" Or _
                    strDocGroup = "TRANSFEROFFER" Or strDocGroup = "NONREG" Then
        mblnIsLifeTimeDocGroup = False
    Else
        mblnIsLifeTimeDocGroup = True
    End If
End Sub

Private Sub SetMortgageTypeText()
    If mblnIsAdditionalBorrowing Then
        mstrMortgageTypeText = "additional borrowing"
        mstrMortgageTypeText_A = "additional borrowing" 'SR 13/10/2004 : BBG1596
        mstrMortgageTypeText4 = "additional borrowing on a repayment basis"
    Else
        If Not mblnIsLifeTimeDocGroup Then
            mstrMortgageTypeText = "mortgage"
            mstrMortgageTypeText_A = "a mortgage" 'SR 13/10/2004 : BBG1596
            mstrMortgageTypeText4 = "a repayment mortgage"
        Else
            mstrMortgageTypeText = "lifetime mortgage"
            mstrMortgageTypeText2 = "mortgage"
            mstrMortgageTypeText_A = "lifetime mortgage" 'SR 13/10/2004 : BBG1596
        End If
    End If
End Sub
'SR 10/09/2004 : CORE82 - End

'********************************************************************************
'** Function:       SetApplicationFactFindBasedProperties
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Sets the appropriate loan purpose element (PURPOSEPURCHASE,
'**                 PURPOSEFURTHERADVANCE or PURPOSEREMORTGAGE).
'** Parameters:     None.
'** Returns:        N/A.
'** Errors:         None Expected
'********************************************************************************
Private Sub SetApplicationFactFindBasedProperties()
    Const cstrFunctionName As String = "SetApplicationFactFindBasedProperties"
    Dim xmlItem As IXMLDOMNode, xmlLC As IXMLDOMNode 'SR 23/09/2004 : CORE82
    Dim strValue As String
    Dim astrTemp() As String
    Dim strTemp As String
    Dim intCount As Integer
    Dim intArrLength As Integer
    Dim groupType As String
    
    On Error GoTo ErrHandler
    
    Set xmlItem = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        '*-get the purpose of the loan
        strValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
'SR 23/09/2004 : CORE82 - Changed the logic. due to the way BBG considers re-mortgages
        If InStr(strValue, "F,") > 0 Then
            mstrMainMortgageTypeGroup = "F"
        Else
'TW 11/11/2005 MAR525
'            Set xmlLC = SingleLoanComponent()
'            strValue = xmlGetAttributeText(xmlLC, "PURPOSEOFLOAN")
'TW 11/11/2005 MAR525 End
            If InStr(strValue, "PSW,") > 0 Then
                mstrMainMortgageTypeGroup = "P"
'PB 22/11/2006 EP2_139 Begin
            ElseIf InStr(strValue, "NP") > 0 Then
                mstrMainMortgageTypeGroup = "NP"
'PB 22/11/2006 EP2_139 End
            ElseIf InStr(strValue, "N,") > 0 Then
                mstrMainMortgageTypeGroup = "N"
            ElseIf InStr(strValue, "R,") Then
                mstrMainMortgageTypeGroup = "R"
            'EP2_1464
            ElseIf InStr(strValue, "TOE,") Then
                mstrMainMortgageTypeGroup = "TOE"
            End If
'TW 11/11/2005 MAR525
            If mstrMainMortgageTypeGroup = "" Then
                Set xmlLC = SingleLoanComponent()
                strValue = xmlGetAttributeText(xmlLC, "PURPOSEOFLOAN")
                If InStr(strValue, "N,") > 0 Then
                    mstrMainMortgageTypeGroup = "N"
                ElseIf InStr(strValue, "R,") Then
                    mstrMainMortgageTypeGroup = "R"
                End If
            End If
'TW 11/11/2005 MAR525 End
        End If

'        astrTemp = Split(strValue, ",")
'        intArrLength = UBound(astrTemp)
'
'        For intCount = 0 To intArrLength Step 1
'            strTemp = astrTemp(intCount)
'            If strTemp = "N" Or strTemp = "R" Or strTemp = "F" Then
'                mstrMainMortgageTypeGroup = strTemp
'                Exit For
'            End If
'        Next intCount
'SR 23/09/2004 : CORE82 - End

        Select Case mstrMainMortgageTypeGroup
            Case "N"
                mstrLoanPurposeElement = "PURPOSEPURCHASE"
            Case "F"
                mstrLoanPurposeElement = "PURPOSEFURTHERADVANCE"
            Case "R"
                mstrLoanPurposeElement = "PURPOSEREMORTGAGE"
            Case "P"
                mstrLoanPurposeElement = "PURPOSEPRODUCTSWITCH"
'PB 21/11/2006 EP2_139 Begin
            Case "TOE"
                mstrLoanPurposeElement = "PURPOSETOE"
            Case "NP"
                mstrLoanPurposeElement = "PURPOSEPORTED"
'PB 21/11/2006 EP2_139 End
        End Select
                    
        '*-get the expected completion date, default to Now + 90 if it has not been set
        If xmlAttributeValueExists(xmlItem, "EXPECTEDCOMPLETIONDATE") Then
            mdtExpectedCompletionDate = xmlGetAttributeAsDate(xmlItem, "EXPECTEDCOMPLETIONDATE")
        Else
            ' PB 24/05/2006 EP603/MAR1777
            ' mdtExpectedCompletionDate = DateAdd("d", 90, Now)
            mdtExpectedCompletionDate = CalcExpectedCompletionDate(Date)
            ' EP603/MAR1777 End
        End If
                
        mblnPackagedApplication = xmlGetAttributeAsBoolean(xmlItem, _
                                    "APPLICATIONPACKAGEINDICATOR") 'SR 07/10/2004 : BBG1545
        
        'EP2_2448 get the special scheme, presently only used for pre-dip KFI
        If xmlAttributeValueExists(xmlItem, "SPECIALSCHEME") Then
            groupType = xmlGetAttributeText(xmlItem, "SPECIALSCHEME")
            mstrSpecialSchemeGroupType = GetComboText("SpecialGroup", groupType)
            mstrSpecialSchemeValidation = GetValidationTypeForValueID("SpecialGroup", groupType)

        End If

    End If
    
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetHasInsurances
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Sets up whether the quote has insurances attached or not.
'** Parameters:     None.
'** Returns:        N/A.
'** Errors:         None Expected
'********************************************************************************
Private Sub SetHasInsurances()
    Const cstrFunctionName As String = "SetHasInsurances"
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler
    
    mblnHasInsurances = False
    Set xmlItem = mxmlData.selectSingleNode(gcstrOTHERINSURANCECOMPANY_PATH)
    If Not xmlItem Is Nothing Then
        '*-we have found at least one insurance, so stop looking
        mblnHasInsurances = True
        Exit Sub
    End If
    Set xmlItem = mxmlData.selectSingleNode(gcstrLIFESUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        '*-we have found at least one insurance, so stop looking
        mblnHasInsurances = True
        Exit Sub
    End If
    Set xmlItem = mxmlData.selectSingleNode(gcstrPAYMENTPROTECTIONSUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        '*-we have found at least one insurance, so stop looking
        mblnHasInsurances = True
        Exit Sub
    End If
    Set xmlItem = mxmlData.selectSingleNode(gcstrBUILDINGSANDCONTENTSSUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        '*-we have found at least one insurance, so stop looking
        mblnHasInsurances = True
        Exit Sub
    End If
    
    Set xmlItem = Nothing
    Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetIsAdditionalBorrowing
'** Created by:     Andy Maggs
'** Date:           07/04/2004
'** Description:    Determines whether the application is for additional
'**                 borrowing or not.
'** Parameters:     vxmlData - the source data XML
'** Returns:        True if it is, else False
'** Errors:         None Expected
'********************************************************************************
Private Sub SetIsAdditionalBorrowing()
    Const cstrFunctionName As String = "SetIsAdditionalBorrowing"
    Dim strValue As String

    On Error GoTo ErrHandler

    '*-get the value for the TYPEOFAPPLICATION
    strValue = GetTypeOfMortgage()
    
'PB 01/12/2006 EP2_139
    '
    'Validation type F does not necessarily mean additional borrowing.
    'New combo is required so we can hav TOE(additional borrowing) & TOE(not additional borrowing).
    '
    'Secondly, 'strValue = "whatever"' will not work as there is probably more than one
    'associated validation type, so could be "CC,F,M,N,P,T,TOE,TOEXA", in which case we need
    'to seach the values for ABO. We can't just use InStr because at some point someone might
    'create type 'ABOXA', which would also be included but might not be expected.
    '
    'Better to call CheckForValidationType function.
    '
    'mblnIsAdditionalBorrowing = (strValue = "F,")
    'EP2_1450
    mblnIsAdditionalBorrowing = CheckForValidationType(strValue, "F")
    'EP2_2478
    mblnIsAdditionalBorrowingTOE = CheckForValidationType(strValue, "RET")
    
    'BC MAR907 Begin
    ' PB 05/06/2006 EP651/MAR1590 Begin
    'mblnIsRemortgage = (strValue = "R,") Or (strValue = "TOE,")
    'blnIsTransferofEquity = (strValue = "TOE,")
    ' PB EP651/MAR1590 End
    
    If InStr(strValue, "PSW,") > 0 Then
        blnIsProductSwitch = True 'BC MAR907
    End If
    
    ' PB 05/06/2006 EP651/MAR1590 Begin
    If InStr(strValue, "TOE,") > 0 Then
        blnIsTransferofEquity = True 'BC MAR907
    End If
    ' PB EP651/MAR1590 End
    
    mblnIsRemortgage = (strValue = "R,") Or (blnIsTransferofEquity)
    'BC MAR907 End
    
    Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       GetTypeOfMortgage
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Gets the type of mortgage from the TYPEOFAPPLICATION
'**                 attribute on the APPLICATIONFACTFIND node.
'** Parameters:     None
'** Returns:        The type of mortgage.
'** Errors:         None Expected
'********************************************************************************
Private Function GetTypeOfMortgage() As String
    Const cstrFunctionName As String = "GetTypeOfMortgage"
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String

    On Error GoTo ErrHandler
    
    Set xmlItem = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        strValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
    End If
    
    GetTypeOfMortgage = strValue
    
    Set xmlItem = Nothing
    Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function

Public Property Get GetMainMortgageTypeGroup() As String
    GetMainMortgageTypeGroup = mstrMainMortgageTypeGroup
End Property

Public Property Get GetSpecialSchemeGroupType() As String
    GetSpecialSchemeGroupType = mstrSpecialSchemeGroupType
End Property

Public Property Get GetSpecialSchemeValidation() As String
    GetSpecialSchemeValidation = mstrSpecialSchemeValidation
End Property


'********************************************************************************
'** Function:       SetMortgageSubQuoteProperties
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Sets the properties extracted from the MORTGAGESUBQUOTE.
'** Parameters:     None
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub SetMortgageSubQuoteProperties()
    Const cstrFunctionName As String = "SetMortgageSubQuoteProperties"
    Dim xmlItem As IXMLDOMNode
    Dim xmlSpecialGroup As IXMLDOMNode 'TK 29/03/2005 BBG1968
    Dim strSpecialGroup As String
    Dim bLifetimeCases As Boolean
    '
    'EP2_2508 GHun
    '' PB 15/06/2006 EP772 Begin
    'Dim xmlLoanComponent As IXMLDOMElement    ' Individual loan component
    'Dim dblTotalAPR As Double                 ' Sum total of each loan component APR values
    'Dim dblTotalAPU As Double                 ' Sum total of each loan component 'Amount Per Unit' values
    'Dim dblAvgAPR As Double                   ' Average APR (total APR / component count)
    'Dim dblAvgAPU As Double                   ' Average APU (total APU / component count)
    'Dim intComponentCount As Integer          ' Count of components
    '' EP772 End
    'EP2_2508 End

    On Error GoTo ErrHandler

    'TK 29/03/2005 BBG1968
    Set xmlSpecialGroup = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    bLifetimeCases = False
    If Not xmlSpecialGroup Is Nothing Then
        strSpecialGroup = xmlGetAttributeText(xmlSpecialGroup, "SPECIALGROUP")
        If InStr(1, strSpecialGroup, "EQ", vbTextCompare) Then
            bLifetimeCases = True
        End If
    End If
    'TK 29/03/2005 BBG1968 End

    Set xmlItem = mxmlData.selectSingleNode(gcstrMORTGAGESUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        mlngTotalLoanAmount = xmlGetAttributeAsLong(xmlItem, "TOTALLOANAMOUNT")
        'EP2_139
        mdblTotalNetMonthlyCost = xmlGetAttributeAsDouble(xmlItem, "TOTALNETMONTHLYCOST")
        mlngLoanAmount = xmlGetAttributeAsLong(xmlItem, "AMOUNTREQUESTED")
        mdblLTV = xmlGetAttributeAsDouble(xmlItem, "LTV")
        
        'EP2_2508 GHun
        mdblAPR = xmlGetAttributeAsDouble(xmlItem, "APR")
        mdblAmountPerPound = xmlGetAttributeAsDouble(xmlItem, "AMOUNTPERUNITBORROWED")
        
        '' PB EP603/MAR1788
        '' PB 15/06/2006 EP772 Begin
        '' Need to loop thru all loan components and get an average of APR and AmountPerUnit
        'For Each xmlLoanComponent In xmlItem.childNodes
        '    If xmlLoanComponent.nodeName = "LOANCOMPONENT" Then
        '        dblTotalAPR = dblTotalAPR + xmlGetAttributeAsDouble(xmlLoanComponent, "APR")
        '        dblTotalAPU = dblTotalAPU + xmlGetAttributeAsDouble(xmlLoanComponent, "AMOUNTPERUNITBORROWED")
        '        intComponentCount = intComponentCount + 1
        '    End If
        'Next
        'If intComponentCount > 0 Then
        '    dblAvgAPR = dblTotalAPR / intComponentCount
        '    dblAvgAPU = dblTotalAPU / intComponentCount
        'End If
        ''
        ''mdblAPR = xmlGetAttributeAsDouble(xmlItem, "APR")
        'mdblAPR = dblAvgAPR
        ''mdblAmountPerPound = xmlGetAttributeAsDouble(xmlItem, "AMOUNTPERUNITBORROWED")
        'mdblAmountPerPound = dblAvgAPU
        '' PB EP772 End
        '' EP603/MAR1788 End
        'EP2_2508 End
        
'IK 01/12/2004 E2EM00003125, E2EM00003126
        mlngDrawDown = Int(xmlGetAttributeAsDouble(xmlItem, "DRAWDOWN"))
        If mlngDrawDown > 0 Then
            mblnHasDrawDown = True
            mdblMonthlyCostLessDrawdown = xmlGetAttributeAsDouble(xmlItem, "MONTHLYCOSTLESSDRAWDOWN")
            'TK 29/03/2005 BBG1968
            If bLifetimeCases Then
                'mlngTotalLoanAmount = mlngTotalLoanAmount
                mlngLoanAmount = mlngLoanAmount + mlngDrawDown
                mdblLTV = xmlGetAttributeAsDouble(xmlItem, "LTVWITHDRAWDOWN")
                mlngAmountRequestedLessDrawDown = mlngLoanAmount - mlngDrawDown

                If mlngTotalLoanAmount > mlngDrawDown Then
                    mlngTotalLoanAmountLessDrawDown = mlngTotalLoanAmount - mlngDrawDown
                Else
                    mlngTotalLoanAmountLessDrawDown = mlngTotalLoanAmount
                End If
            Else
                mlngAmountRequestedLessDrawDown = mlngLoanAmount - mlngDrawDown
'IK 03/12/2004 E2EM00003125
                mlngTotalLoanAmountLessDrawDown = mlngTotalLoanAmount - mlngDrawDown
'IK 03/12/2004 E2EM00003125 ends
            End If
            'TK 29/03/2005 BBG1968 End
        End If
'IK 01/12/2004 E2EM00003125, E2EM00003126 ends
        
    End If
    
    Set xmlItem = Nothing
    Set xmlSpecialGroup = Nothing
    Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlSpecialGroup = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetLoanComponentBasedProperties
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Sets the properties extracted from the LOANCOMPONENTS.
'** Parameters:     None
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub SetLoanComponentBasedProperties()
    Const cstrFunctionName As String = "SetLoanComponentBasedProperties"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim eType As MortgageInterestRateType
    Dim intInterestOnly As Integer
    Dim intRepayment As Integer
    Dim intPart As Integer
    Dim strValue As String
    Dim strRepayChar As String
    Dim intTempTermInYears As Integer
    Dim intTempTermInMonths As Integer
    
'PB 04/12/2006 EP2_139
    Dim lngTotalAmountPayable As Long
    Dim lngAmountBorrowed As Long
    
    Dim eRepaymentType As enumRepaymentType
    
    On Error GoTo ErrHandler
    
    mblnIsVariableRate = False
    mlngLoanAmountRepayable = 0
    
    Set xmlList = mxmlData.selectNodes(gcstrLOANCOMPONENT_PATH)
    
    If xmlList.length > 1 Then
        mstrComponentsTypeElement = "MULTICOMPONENT"
        mblnIsMultiComponent = True
    Else
        mstrComponentsTypeElement = "SINGLECOMPONENT"
        mblnIsMultiComponent = False
    End If
    
    For Each xmlItem In xmlList
        
        'EP2_2395
        '*-check whether this component is interest only or capital and interest
        If CheckForValidationType(xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "1") _
        Or CheckForValidationType(xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "I") Then
            'Interest only
            eRepaymentType = enumInterestOnly
        ElseIf CheckForValidationType(xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "2") _
        Or CheckForValidationType(xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "C") Then
            'Capital and interest
            eRepaymentType = enumCapitalAndInterest
        Else
            'Something else - ??
        End If
        '
        'BBG1565 This will return a non fixed rate interest type, unless
        'only a fixed rate int type exists, GetMortgageInterestRateType only
        'looks at the first interest rate type, when there could be others
'        eType = GetMortgageInterestRateType(Me, xmlItem)
        eType = CheckIfVariableMortgageInterestRateType(Me, xmlItem)
        If eType <> mrtFixedRate Then
            '*-there is at least one component that is not fixed
            mblnIsVariableRate = True
        End If
        
        'CORE82
        mdblAmountPerUnitBorrowed = xmlGetAttributeAsDouble(xmlItem, "AMOUNTPERUNITBORROWED")
        
        lngTotalAmountPayable = xmlGetAttributeAsLong(xmlItem, "TOTALAMOUNTPAYABLE")
        'EP2_2395
        lngAmountBorrowed = xmlGetAttributeAsLong(xmlItem, "TOTALLOANCOMPONENTAMOUNT")
        '
        '*-add the amount to the total repayable
        mlngLoanAmountRepayable = mlngLoanAmountRepayable + lngTotalAmountPayable
        '
        '*-add the amount borrowed
        If eRepaymentType = enumInterestOnly Then
            mlngTotalInterestOnly = mlngTotalInterestOnly + lngAmountBorrowed
        ElseIf eRepaymentType = enumCapitalAndInterest Then
            mlngTotalCapitalandInterest = mlngTotalCapitalandInterest + lngAmountBorrowed
        End If
        
        '*-determine the overall repayment method
        strValue = xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD")
        strRepayChar = Mid(strValue, 1, 1)
        
        Select Case strRepayChar
            Case "1"
                intInterestOnly = intInterestOnly + 1
            Case "2"
                intRepayment = intRepayment + 1
            Case Else
                intPart = intPart + 1
        End Select
        
        intTempTermInYears = xmlGetAttributeAsInteger(xmlItem, "TERMINYEARS")
        intTempTermInMonths = xmlGetAttributeAsInteger(xmlItem, "TERMINMONTHS")
        
        If intTempTermInYears > mintTermInYears Then
            mintTermInYears = intTempTermInYears
            mintTermInMonths = intTempTermInMonths
        Else
            If intTempTermInYears = mintTermInYears Then
                If intTempTermInMonths > mintTermInMonths Then
                    mintTermInMonths = intTempTermInMonths
                End If
            End If
        End If
    
    Next xmlItem
    
    If intInterestOnly = xmlList.length Then
        '*-all loan components are interest only
        mstrMortgageRepaymentTypeElement = "INTERESTONLY"
    ElseIf intRepayment = xmlList.length Then
        '*-all loan components are repayment
        mstrMortgageRepaymentTypeElement = "REPAYMENT"
    Else
        '*-there is a mixture of interest only and repayment loan components
        mstrMortgageRepaymentTypeElement = "PARTANDPART"
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetFeesInfo
'** Created by:     Andy Maggs
'** Date:           30/03/2004
'** Description:    Gets the information about fees associated with the mortgage
'**                 quote and returns whether all fees are added to the loan, paid
'**                 separately, or part and part.
'** Parameters:     vxmlData - the source data.
'**                 reFeeType - the type of fees, i.e. none, all paid, part paid
'**                 part added, or all added.
'**                 rdblFeesAddedAmount - the total amount of fees added to the
'**                 loan.
'**                 rdblTotalFees - the total amount of fees whether added or not.
'** Returns:        Data returned via ByRef parameters.
'** Errors:         None Expected
'********************************************************************************
Private Sub SetFeesInfo()
    Const cstrFunctionName As String = "SetFeesInfo"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim lngAmount As Long
    Dim strType As String
    Dim intNoFeeAddedToLoan As Integer

    On Error GoTo ErrHandler
   
    '*-zero values
    mlngFeesAddedToLoanAmount = 0
    mlngTotalFees = 0
    
    '*-need to iterate through the mortgage one off costs to identify all those that
    '*-have been added to the loan
    Set xmlList = mxmlData.selectNodes(gcstrMORTGAGEONEOFFCOSTS_PATH)
    intNoFeeAddedToLoan = 0
    If xmlList.length = 0 Then
        meFeesType = fptNone
    Else
        'Array of the fee amounts that are added to loan
        ReDim mastrFeesAddedToLoan(xmlList.length - 1)
        
        For Each xmlItem In xmlList
            lngAmount = xmlGetAttributeAsLong(xmlItem, "AMOUNT")
            '*-add this amount to the total
            mlngTotalFees = mlngTotalFees + lngAmount
            If xmlGetAttributeAsBoolean(xmlItem, "ADDTOLOAN") Then
                '*-also add the amount to the fees added to loan total
                mlngFeesAddedToLoanAmount = mlngFeesAddedToLoanAmount + lngAmount
                mastrFeesAddedToLoan(intNoFeeAddedToLoan) = lngAmount
                intNoFeeAddedToLoan = intNoFeeAddedToLoan + 1
            End If
            
            '*-get the type
            strType = UCase$(xmlGetAttributeText(xmlItem, "MORTGAGEONEOFFCOSTTYPE"))
            If InStr(strType, "SEA,") > 0 Or InStr(strType, "DEE,") > 0 Then
                mlngMortgageCompletionFee = mlngMortgageCompletionFee + lngAmount
            End If
            
        Next xmlItem
        
        If mlngFeesAddedToLoanAmount = mlngTotalFees Then
            meFeesType = fptAllAdded
        ElseIf mlngFeesAddedToLoanAmount = 0 Then
            meFeesType = fptAllPaid
        Else
            meFeesType = fptPartPart
        End If
    End If
    
    'Redimenion the array of fees added to loan (removing all element with no
    'values assigned)
    If intNoFeeAddedToLoan > 0 Then
        ReDim Preserve mastrFeesAddedToLoan(intNoFeeAddedToLoan - 1)
    Else
        ReDim Preserve mastrFeesAddedToLoan(0)
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetInsurancePremiumInfo
'** Created by:     Andy Maggs
'** Date:           30/03/2004
'** Description:    Gets the information about insurance premiums associated with
'**                 the mortgage quote and returns whether all premiums are added
'**                 to the loan, paid separately, or part and part.
'** Parameters:     vxmlData - the source data.
'**                 rePremType - the type of premiums, i.e. none, all paid, part
'**                 paid part added, or all added.
'**                 rdblInsurancePremiumsAdded - the total amount of premiums
'**                 added to the loan.
'**                 rdblTotalPremiums - the total amount of premiums whether
'**                 added or not.
'** Returns:        Data returned via ByRef parameters.
'** Errors:         None Expected
'********************************************************************************
Private Sub SetInsurancePremiumInfo()
    Const cstrFunctionName As String = "SetInsurancePremiumInfo"

    On Error GoTo ErrHandler

    '*-zero values
    mdblPremiumsAddedToLoanAmount = 0
    mdblTotalInsurancePremiums = 0
    '*-insurance premiums are not actually relevant to Omiga because it
    '*-does not support adding them to the loan
    meInsurancePremiumType = fptNone

    Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       SetMortgageLenderProperties
'** Created by:     Andy Maggs
'** Date:           01/06/2004
'** Description:    Sets the properties from the Mortgage Lender section.
'** Parameters:     None
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub SetMortgageLenderProperties()
    Const cstrFunctionName As String = "SetMortgageLenderProperties"
    
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler

    '*-get the value for and add the PROVIDER attribute
    Set xmlItem = mxmlData.selectSingleNode(gcstrMORTGAGELENDER_PATH)
    If Not xmlItem Is Nothing Then
        mstrProvider = xmlGetAttributeText(xmlItem, "LENDERNAME")
        mstrProviderCode = xmlGetAttributeText(xmlItem, "LENDERCODE") 'LH 13/10/2006 EP1226 / CC132
    End If
    
    Set xmlItem = Nothing
    Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


'PM 19/05/2006 : EPSOM - EP584 : Start
'********************************************************************************
'** Function:       StandardContactName
'** Created by:     Pat Morse
'** Date:           22/05/2006
'** Description:    Builds the standard contact name from the title, forename and surname
'**                 EPSOM EP584
'** Parameters:     strTitle
'**                 strForename
'**                 strSurname
'** Returns:        Standard contact name in the format e.g. Mr T Higgins
'** Errors:         None Expected
'********************************************************************************
Public Function StandardContactName( _
    ByVal strTitle As String _
    , ByVal strForename As String _
    , ByVal strSurname As String) _
    As String
    
    Dim strContactName As String
    
    strContactName = strTitle

    If Len(strContactName) > 0 Then
        If Len(strForename) > 0 Then
            strContactName = strContactName + " " + Mid(strForename, 1, 1)
        End If
    Else
        strContactName = Mid(strForename, 1, 1)
    End If
    
    If Len(strContactName) > 0 Then
        If Len(strSurname) > 0 Then
            strContactName = strContactName + " " + strSurname
        End If
    Else
        strContactName = strSurname
    End If
    
    StandardContactName = strContactName
End Function


'********************************************************************************
'** Function:       SetBrokerAndPackager
'** Created by:     Pat Morse
'** Date:           23/05/2006
'** Description:    Sets up whether the quote has insurances attached or not.
'** Parameters:     None.
'** Returns:        N/A.
'** Errors:         None Expected
'********************************************************************************
'EP2_139   Clean up for Offer and KFI document changes.
'EP2_422   Section13 changes
Private Sub SetBrokerAndPackager()
    Const cstrFunctionName As String = "SetBrokerAndPackager"
    Dim xmlIntroducer       As IXMLDOMNode
    Dim xmlIntroducerFirm   As IXMLDOMNode
    Dim xmlARFirm           As IXMLDOMNode
    Dim xmlPrincipalFirm       As IXMLDOMNode
    Dim xmlAppIntroducer       As IXMLDOMNode
    Dim xmlFirmTradingName       As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim xmlList As IXMLDOMNodeList
    Dim xmlMortgageClub As IXMLDOMNode
    
    Dim strBrokerName           As String
    Dim strPackagerName           As String
    Dim strARName           As String
    Dim strPrincipalName           As String
    Dim strMortgageClubName           As String
    Dim strPackageAssocName           As String
    Dim strIntermediaryType As String
    'EP2_704
    Dim fsaRefL_G As String
    Dim fsaRefAR As String
    Dim blnLandG As Boolean
    
    '------------------------------Introducer details----------------------------------
    Set xmlIntroducer = mxmlData.selectSingleNode(gcstrORGANISATIONUSER_PATH)
    If Not xmlIntroducer Is Nothing Then
            mstrIntroducerName = StandardContactName _
                (xmlGetAttributeText(xmlIntroducer, "USERTITLE_TEXT") _
                , xmlGetAttributeText(xmlIntroducer, "USERFORENAME") _
                , xmlGetAttributeText(xmlIntroducer, "USERSURNAME"))
        mblnIndividualIntroducer = True
    Else
        mblnIndividualIntroducer = False
    End If

    'EP2_704
    Set xmlItem = mxmlData.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "FSARefNoLandG" & Chr$(34) & "]")
    fsaRefL_G = xmlGetAttributeText(xmlItem, "STRING")
    blnLandG = False

    'use the FirmTradingName if it exists
    Set xmlIntroducer = mxmlData.selectSingleNode(gcstrINTRODUCER_PATH)
    Set xmlAppIntroducer = mxmlData.selectSingleNode(gcstrARFIRM_PATH)
    If Not xmlAppIntroducer Is Nothing Then
        'Set xmlFirmTradingName = xmlAppIntroducer.selectSingleNode("FIRMTRADINGNAME") ' SR EP2_2270
        Set xmlARFirm = xmlAppIntroducer.selectSingleNode("ARFIRM")
        If Not xmlARFirm Is Nothing Then
            fsaRefAR = xmlGetAttributeText(xmlARFirm, "FSAARFIRMREF")
            If (StrComp(fsaRefL_G, fsaRefAR) = 0) Then
                blnLandG = True
            End If
        End If

        Set xmlIntroducerFirm = xmlIntroducer.selectSingleNode(".//INTRODUCERFIRM[@ARFIRMID='" & xmlGetAttributeText(xmlARFirm, "ARFIRMID") & "']")
        If Not xmlIntroducerFirm Is Nothing Then
            strARName = xmlGetAttributeText(xmlIntroducerFirm, "TRADINGAS")
        End If
        If (strARName = "") Then
            If Not xmlARFirm Is Nothing Then
                strARName = xmlGetAttributeText(xmlARFirm, "ARFIRMNAME")
            End If
        End If
    End If

    Set xmlList = mxmlData.selectNodes(gcstrPRINCIPALFIRM_PATH)
    'could be two principal firms one packager, one not
    For Each xmlItem In xmlList
        Set xmlPrincipalFirm = xmlItem.selectSingleNode("PRINCIPALFIRM")
        If Not xmlPrincipalFirm Is Nothing Then
            strIntermediaryType = xmlGetAttributeText(xmlPrincipalFirm, "PACKAGERINDICATOR")
        End If
        If strIntermediaryType = "1" Then
            '------------------------------Packager details----------------------------------
            If Not xmlPrincipalFirm Is Nothing Then
                strPackagerName = xmlGetAttributeText(xmlPrincipalFirm, "PRINCIPALFIRMNAME")
             End If
             mstrPackagerName = strPackagerName
        Else
            'EP2395
            If Not xmlIntroducer Is Nothing Then
                Set xmlIntroducerFirm = xmlIntroducer.selectSingleNode(".//INTRODUCERFIRM[@PRINCIPALFIRMID='" & xmlGetAttributeText(xmlPrincipalFirm, "PRINCIPALFIRMID") & "']")
                If Not xmlIntroducerFirm Is Nothing Then
                    strPrincipalName = xmlGetAttributeText(xmlIntroducerFirm, "TRADINGAS")
                End If
                    
                If strPrincipalName = "" Then
                    If Not xmlPrincipalFirm Is Nothing Then
                        strPrincipalName = xmlGetAttributeText(xmlPrincipalFirm, "PRINCIPALFIRMNAME")
                    End If
                End If
            End If
        End If
    Next xmlItem
    
    'MortgageClubNetworkAssoc
    Set xmlList = mxmlData.selectNodes(gcstrMORTGAGECLUBNETWORKASSOC_PATH)
    
    For Each xmlItem In xmlList
        Set xmlMortgageClub = xmlItem.selectSingleNode("MORTGAGECLUBNETWORKASSOCIATION")
        If Not xmlMortgageClub Is Nothing Then
            'EP2_704 Need PackageAssoc
            strIntermediaryType = xmlGetAttributeText(xmlMortgageClub, "PACKAGERINDICATOR")
            If strIntermediaryType = "1" Then
                strPackageAssocName = xmlGetAttributeText(xmlMortgageClub, "MORTGAGECLUBNETWORKASSOCNAME")
            Else
                strMortgageClubName = xmlGetAttributeText(xmlMortgageClub, "MORTGAGECLUBNETWORKASSOCNAME")
            End If
        End If
    Next xmlItem
    
    'EP2_704 IntermediaryNames displayed varies according to E2CR24_26 flowchart
    If Len(strPackageAssocName) > 0 Then
    
    
        mstrIntermediaryNames = SetIntermediaryNames(strPrincipalName, strPackagerName _
                        , "", strARName, strPackageAssocName, blnLandG)
    
    
    
    ElseIf Len(strPackagerName) > 0 Then
        
        
        
        mstrIntermediaryNames = SetIntermediaryNames(strPrincipalName, strPackagerName _
                        , "", strARName, "", blnLandG)
    
    
    
    ElseIf Len(strMortgageClubName) > 0 Then
        
        
        
        mstrIntermediaryNames = SetIntermediaryNames(strPrincipalName, "" _
                        , strMortgageClubName, strARName, "", False)
    
    
    
    ElseIf (Len(strPrincipalName) Or Len(strARName)) Then
        
        
        
        mstrIntermediaryNames = SetIntermediaryNames(strPrincipalName, "" _
                        , "", strARName, "", False)
    
    
    
    ElseIf Len(mstrIntroducerName) > 0 Then
    
        mstrIntermediaryNames = SetIntermediaryNames(strPrincipalName, mstrIntroducerName, "", strARName, "", False)
    
    End If

    'If we don't have an ARFirmName, use the PrincipalFirmName
    If Len(strARName) > 0 Then
        mstrBrokerName = strARName
    'EP2_1994
    ElseIf Len(strPrincipalName) > 0 Then
        mstrBrokerName = strPrincipalName
    ElseIf Len(mstrIntroducerName) > 0 Then
        mstrBrokerName = mstrIntroducerName
    End If
    
    Set xmlIntroducer = Nothing
    Set xmlIntroducerFirm = Nothing
    Set xmlMortgageClub = Nothing
    Set xmlARFirm = Nothing
    Set xmlPrincipalFirm = Nothing
    Set xmlAppIntroducer = Nothing
    Set xmlFirmTradingName = Nothing
    Set xmlItem = Nothing
    Set xmlList = Nothing

    Exit Sub
ErrHandler:
    Set xmlIntroducer = Nothing
    Set xmlIntroducerFirm = Nothing
    Set xmlMortgageClub = Nothing
    Set xmlARFirm = Nothing
    Set xmlPrincipalFirm = Nothing
    Set xmlAppIntroducer = Nothing
    Set xmlFirmTradingName = Nothing
    Set xmlItem = Nothing
    Set xmlList = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'PM 19/05/2006 : EPSOM - EP584 : End


'********************************************************************************
'** Function:       GetProcurationFee
'** Created by:     Ian
'** Date:           29/06/2004
'** Description:    Gets the Procuration fee.
'** Parameters:     None
'** Returns:        The introducer fee.
'** Errors:         None Expected
'********************************************************************************
Public Function GetProcurationFee() As String
    Const cstrFunctionName As String = "GetProcurationFee"
    Dim xmlItem As IXMLDOMNode
    
    On Error GoTo ErrHandler
    Set xmlItem = mxmlData.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'PRC')]")
    If Not xmlItem Is Nothing Then
        GetProcurationFee = xmlGetAttributeText(xmlItem, "AMOUNT")
    End If

    Set xmlItem = Nothing
    Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function

Public Property Get TermInYears() As Integer
    TermInYears = mintTermInYears
End Property

Public Property Get TermInMonths() As Integer
    TermInMonths = mintTermInMonths
End Property

Public Property Get IsMultiComponent() As Boolean
    IsMultiComponent = mblnIsMultiComponent
End Property

'SR 12/09/2004 : CORE82
Public Property Get StandardVariableRate() As Double
    StandardVariableRate = mdblStandardVariableRate
End Property

Public Property Let StandardVariableRate(ByVal dblRate As Double)
    mdblStandardVariableRate = dblRate
End Property
'SR 12/09/2004 : CORE82 - End

'INR 17/09/2004 : CORE82
Public Property Get AmountPerUnitBorrowed() As Double
    AmountPerUnitBorrowed = mdblAmountPerUnitBorrowed
End Property

'EP2_139
Public Property Get IsFlexible() As Boolean
    IsFlexible = mblnIsFlexible
End Property
'EP2_139
Public Sub SetFlexible()
    Const cstrFunctionName As String = "SetFlexible"
    Dim xmlLoanComponents As IXMLDOMNodeList
    Dim xmlMortgageProduct As IXMLDOMNode
    Dim strFlexOption As String
    
    On Error GoTo ErrHandler
    
    Set xmlLoanComponents = LoanComponents
    If xmlLoanComponents.length = 0 Then
        Exit Sub
    End If
    
    'If one product is Flexible, they all have to be flexible
    If xmlLoanComponents.length > 0 Then
        Set xmlMortgageProduct = mxmlData.selectSingleNode(".//MORTGAGEPRODUCT")
        If Not xmlMortgageProduct Is Nothing Then
                strFlexOption = xmlGetAttributeText(xmlMortgageProduct, "FLEXIBLEMORTGAGEPRODUCT")
                If StrComp(strFlexOption, "1") = 0 Then
                    mblnIsFlexible = True
                Else
                    mblnIsFlexible = False
                End If
        End If
    End If

    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'EP2_139
Public Property Get IsIncentive() As Boolean
    IsIncentive = mblnIsIncentive
End Property

'EP2_139
Public Property Get IsCashback() As Boolean
    IsCashback = mblnIsCashback
End Property
'EP2_139
Public Sub SetCashback()
    Const cstrFunctionName As String = "SetCashback"
    Dim xmlLoanComponents As IXMLDOMNodeList
    Dim xmlMortgageProduct As IXMLDOMNode
    Dim strFlexOption As String
    
    On Error GoTo ErrHandler
    
    Set xmlLoanComponents = LoanComponents
    If xmlLoanComponents.length = 0 Then
        Exit Sub
    End If
    
    'If one product is Flexible, they all have to be flexible
    If xmlLoanComponents.length > 0 Then
        Set xmlMortgageProduct = mxmlData.selectSingleNode(".//MORTGAGEPRODUCT")
        If Not xmlMortgageProduct Is Nothing Then
                strFlexOption = xmlGetAttributeText(xmlMortgageProduct, "CASHBACKPRODUCT")
                If StrComp(strFlexOption, "1") = 0 Then
                    mblnIsCashback = True
                    mblnIsIncentive = True
                Else
                    mblnIsCashback = False
                End If
        End If
    End If

    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


'EP2_139
Public Property Get IsFreeLegal() As Boolean
    IsFreeLegal = mblnIsFreeLegal
End Property
'EP2_139
Public Sub SetFreeLegal()
    Const cstrFunctionName As String = "SetFreeLegal"
    Dim xmlLoanComponents As IXMLDOMNodeList
    Dim xmlMortgageProduct As IXMLDOMNode
    Dim strFlexOption As String
    
    On Error GoTo ErrHandler
    
    Set xmlLoanComponents = LoanComponents
    If xmlLoanComponents.length = 0 Then
        Exit Sub
    End If
    
    'If one product is Flexible, they all have to be flexible
    If xmlLoanComponents.length > 0 Then
        Set xmlMortgageProduct = mxmlData.selectSingleNode(".//MORTGAGEPRODUCT")
        If Not xmlMortgageProduct Is Nothing Then
            'EP2_718 Should be FREELEGALFEES
            strFlexOption = xmlGetAttributeText(xmlMortgageProduct, "FREELEGALFEES")
            If StrComp(strFlexOption, "1") = 0 Then
                mblnIsFreeLegal = True
                mblnIsIncentive = True
            Else
                mblnIsFreeLegal = False
            End If
        End If
    End If

    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'EP2_139
Public Property Get IsFreevaluation() As Boolean
    IsFreevaluation = mblnIsFreevaluation
End Property
'EP2_139
Public Sub SetFreeValuation()
    Const cstrFunctionName As String = "SetFreevaluation"
    Dim xmlLoanComponents As IXMLDOMNodeList
    Dim xmlMortgageProduct As IXMLDOMNode
    Dim xmlCostList As IXMLDOMNodeList
    Dim xmlCost As IXMLDOMNode
    Dim strFlexOption As String
    Dim strValuationType As String
    
    On Error GoTo ErrHandler
    
    Set xmlLoanComponents = LoanComponents
    If xmlLoanComponents.length = 0 Then
        Exit Sub
    End If
    
    'If one product is Flexible, they all have to be flexible
    If xmlLoanComponents.length > 0 Then
        Set xmlMortgageProduct = mxmlData.selectSingleNode(".//MORTGAGEPRODUCT")
        If Not xmlMortgageProduct Is Nothing Then
                strFlexOption = xmlGetAttributeText(xmlMortgageProduct, "REFUNDOFVALUATION")
                If StrComp(strFlexOption, "1") = 0 Then
                    mblnIsFreevaluation = True
                    mblnIsIncentive = True
                Else
                    mblnIsFreevaluation = False
                End If
        End If
    End If

    If mblnIsFreevaluation Then
        Set xmlCostList = mxmlData.selectNodes("//MORTGAGEONEOFFCOST")
        For Each xmlCost In xmlCostList
            strValuationType = UCase(xmlGetAttributeText(xmlCost, "MORTGAGEONEOFFCOSTTYPE", ""))
            If strValuationType <> "" Then
                If CheckForValidationType(strValuationType, "VAL") Then
                    mFreeValuationAmount = xmlGetAttributeAsDouble(xmlCost, "AMOUNT")
                End If
            End If
        Next xmlCost
    End If
    
    Set xmlCost = Nothing
    Set xmlCostList = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlLoanComponents = Nothing
    Set xmlMortgageProduct = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'EP2_139
Public Property Get GetFreeValuationAmount() As Double
    GetFreeValuationAmount = mFreeValuationAmount
End Property

'EP2_422
'EP2_704 reworked. Have included boolean to identify whether AR Broker is L&G as per Flowchart in E2CR24_26
'Not required as yet as if L&G is principal name will come thru anyway, but Mike says there is a change control where it will be required
Public Function SetIntermediaryNames( _
    ByVal strPrincipalName As String _
    , ByVal strPackagerName As String _
    , ByVal strMortgageClubName As String _
    , ByVal strARName As String _
    , ByVal strPackageAssocName As String _
    , ByVal blnLandG As Boolean) _
    As String
    
    Const cstrFunctionName As String = "FeePayableToThirdParty"
    Dim strCombineNames As String
    Dim strCombineText As String
    
On Error GoTo ErrHandler
' FeePayableToThirdParty(ByVal checkPR As Boolean _
'                                        , ByVal checkA As Boolean _
'                                        , ByVal checkPA As Boolean)
    If FeePayableToThirdParty(True, True, True) > 0 Then
        strCombineText = " and "
    Else
        strCombineText = " or "
    End If
    
    strCombineNames = strPackageAssocName

    If Len(strCombineNames) > 0 Then
        If Len(strPackagerName) > 0 Then
            strCombineNames = strCombineNames + strCombineText + strPackagerName
        End If
    Else
        If Len(strPackagerName) > 0 Then
            strCombineNames = strPackagerName
        End If
    End If

    If Len(strCombineNames) > 0 Then
        If Len(strMortgageClubName) > 0 Then
            strCombineNames = strCombineNames + strCombineText + strMortgageClubName
        End If
    Else
        If Len(strMortgageClubName) > 0 Then
            strCombineNames = strMortgageClubName
        End If
    End If
    
    If Len(strCombineNames) > 0 Then
        If Len(strPrincipalName) > 0 Then
            strCombineNames = strCombineNames + strCombineText + strPrincipalName
        End If
    Else
        strCombineNames = strPrincipalName
    End If
    
    If Len(strCombineNames) > 0 Then
        If Len(strARName) > 0 Then
            strCombineNames = strCombineNames + strCombineText + strARName
        End If
    Else
        strCombineNames = strARName
    End If
    
    SetIntermediaryNames = strCombineNames
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'EP2_422
Private Function FeePayableToThirdParty(ByVal checkPR As Boolean _
                                        , ByVal checkA As Boolean _
                                        , ByVal checkPA As Boolean) As Double
    
    Const cstrFunctionName As String = "FeePayableToThirdParty"
    Dim xmlMortgageIntroducerFeeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim dblTPPayment As Double, strValidationType As String
    
On Error GoTo ErrHandler

    Set xmlMortgageIntroducerFeeList = mxmlData.selectNodes(gcstrAPPPROCFEES_PATH)
    For Each xmlNode In xmlMortgageIntroducerFeeList
        'EP2_704 Compare against INTRODUCERFEETYPE validationtype
        'EP2_704 Won't always require all amounts
        strValidationType = xmlGetAttributeText(xmlNode, "INTRODUCERFEETYPE_VALIDID")
        If ((checkPR And CheckForValidationType(strValidationType, "PR")) Or _
            (checkA And CheckForValidationType(strValidationType, "A")) Or _
            (checkPA And CheckForValidationType(strValidationType, "PA"))) Then
            dblTPPayment = dblTPPayment + xmlGetAttributeAsLong(xmlNode, "FEEAMOUNT")
        End If
    Next xmlNode
    
    mstrFeePayableToThirdParty = dblTPPayment
   
    FeePayableToThirdParty = dblTPPayment
    
ErrHandler:
    Set xmlMortgageIntroducerFeeList = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'EP2_1449
Public Function AddInsuranceLoanPurposeElement(ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "AddInsuranceLoanPurposeElement"
    Dim loanPurposeElement As String
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    
    On Error GoTo ErrHandler
    'EP2_1530/EP2_1667
    Set xmlItem = mxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        '*-get the purpose of the loan
        strValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
    End If
        
    If CheckForValidationType(strValue, "KFI") Then
        loanPurposeElement = "PURPOSEMORTGAGE"
    Else
        loanPurposeElement = "PURPOSEREMORTGAGE"
    End If

    Set AddInsuranceLoanPurposeElement = CreateNewElement(loanPurposeElement, vxmlNode)

    Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'EP2_1667
Public Sub AddCustomerAddedRemovedData(ByVal vstrXPath As String, _
        ByVal vstrNodeName As String, ByVal vstrAttribName As String, _
        ByVal vxmlNode As IXMLDOMNode)
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim strNamesList As String
    Dim xmlNames As IXMLDOMNode

    Set xmlList = mxmlData.selectNodes(vstrXPath)
    If Not xmlList Is Nothing Then
        For Each xmlItem In xmlList
            If Len(strNamesList) > 0 Then
                strNamesList = strNamesList & vbCrLf
            End If
            strNamesList = strNamesList & BuildName(xmlGetAttributeText(xmlItem, "TITLE_TEXT"), _
                    xmlGetAttributeText(xmlItem, "FORENAME"), _
                    xmlGetAttributeText(xmlItem, "SURNAME"))
        Next xmlItem

        If Len(strNamesList) > 0 Then
            '*-add the ADDNAMES element
            Set xmlNames = CreateNewElement(vstrNodeName, vxmlNode)
            '*-add the NAMESTOADD attribute
            Call xmlSetAttributeValue(xmlNames, vstrAttribName, strNamesList)
        End If
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlNames = Nothing
End Sub

'EP2_1861
Public Sub SetRefundOfValuation()
    '
    Const cstrTHIS_FUNCTION As String = "SetRefundOfValuation"
    Dim strErrorAssist As String
    On Error GoTo eh_SetRefundOfValuation
    '
    Dim xmlItem As IXMLDOMNode
    Dim xmlNodes As IXMLDOMNodeList
    '
    mblnIsRefundOfValuation = False
    strErrorAssist = "Getting node list"
    Set xmlNodes = mxmlData.selectNodes(gcstrMORTGAGEPRODUCT_PATH)
    strErrorAssist = "Beginning loop through node list"
    For Each xmlItem In xmlNodes
        If xmlGetAttributeAsBoolean(xmlItem, "REFUNDOFVALUATION") Then
            strErrorAssist = "Setting result=true"
            mblnIsRefundOfValuation = True
            Exit For
        End If
    Next
    '
CleanUp:
    Set xmlItem = Nothing
    Set xmlNodes = Nothing
    Exit Sub
    '
eh_SetRefundOfValuation:
    'call common error handler
    GoTo CleanUp
End Sub
Public Function IsRefundOfValuation() As Boolean
    IsRefundOfValuation = mblnIsRefundOfValuation
End Function

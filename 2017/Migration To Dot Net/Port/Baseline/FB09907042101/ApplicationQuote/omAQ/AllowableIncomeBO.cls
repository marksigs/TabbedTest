VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AllowableIncomeBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Dependencies:  List any other dependent components
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     19/09/01    Created
'SG     21/02/02    SYS3738 Amended CalculateAllowableIncome for ComfirmedEarnedIncome changes
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist

'BMIDS
'---------------------------------------------------------------------------------------------
'AW     21/10/02    BMIDS00653 - BM029, Amended CalcAllowableInc, CalculateAllowableIncome
'                                       Added CalculateOutGoings for Net Allowable Income Calcs
' MDC   29/10/02    BMIDS00653 - Further updates to Allowable Income Calcs
' MDC   05/11/2002  BMIDS00653 - Replace hardcoded valueId's with ValidationTypes
' MDC   06/11/2002  BMIDS00861 - Exclude Unearned Income from Allowable Income calculation
' MDC   07/11/2002  BMIDS00865 - Include Retired Income in Allowable Income calculation
' MDC   07/11/2002  BMIDS00863 - Correction for Self Employed Net Profit Calculation
' MDC   08/11/2002  BMIDS00879 - Correction of typo in Combo group name in Allowable Income Calc
' AW    13/11/2002  BMIDS00886 - Retired Income not taken into account because it has a MainStatus = 0
' MV    31/01/2003  BM0304 - Amended CalcAllowableIncome()
' GD    19/02/2003  BM0360  -   Amended CalculateAllowableIncome()
' GHun  03/03/2003  BM0376 - CC026 Include guarantors in calcs
' GHun  11/03/2003  BM0376 - CC026 Ignore unassigned LoansLiabilities in CalculateOutgoings
'---------------------------------------------------------------------------------------------
'INGDUK
'---------------------------------------------------------------------------------------------
' JD    28/09/2005  MAR30   Added underwriter override processing to CalculateAllowableIncome. Added TotalMortgages to CalculateOutgoings
' MahaT 16/11/2005  MAR272  1. Ammended CalcAllowableIncome to set NetConfirmedAllowbleIncome only if ConfirmedAllowableIncome > 0
'                           2. Ammended CalculateOutGoins for calculating mortage account outgoings (XPath problem)
' MahaT 17//1/2005  MAR174  Ammended CalculateOutGoins for bNumApplicants (XPath problem)
' INR   02/12/2005  MAR175 Global LoanTermMonthsRemaining changed to Number of days basis
' JD    16/12/2005  MAR118 Only include in Loans and Liabilities balance those with validation type IIB (CalculateOutgoings)
' JD    22/01/2006  MAR1040 changed CalculateAllowableIncome to calculate both declared and confirmed.
'                           changed CalculateOutgoings to get the mortgage outgoing amount from the mortgageaccount not the mortgageloan.
' JD    08/02/2006  MAR1040 CalculateOutgoings - use the amount from mortgageloan if present, else use the amount from mortgageaccount
' PSC   17/05/2006  MAR1801 Correct calculation of Mortgage outgoings
' JD    13/06/06    MAR688  Changes for performance - added use of stored procedures.
' GHun  10/07/2006  MAR1891 Changed CalculateAllowableIncome & CalculateOutGoings & added IsItemInValidationXML
' JD    20/07/06    MAR1912 CalculateAllowableIncome - fixed self-employed calculation to do confirmed AND declared calc.
' JD    20/07/06    MAR1919 ClaculateAllowableIncome - add check for global UseAllEmployments
'---------------------------------------------------------------------------------------------
'EPSOM
'---------------------------------------------------------------------------------------------
' PE    01/08/2006  EP1061  dc160 - net allowable income not coming accross for self employed/contract workers. (Fixed XPath expressions)
' HMA   08/08/2006  EP1053  Correct CalculateOutgoings.
' LDM   17/08/2006  EP1087  Correct CalculateAllowableIncome. self employed income error
' MCh   04/12/2006  EP2_128 Changes made to CalculateAllowableIncome
' PSC   26/01/2007  EP2_976 Bring code in line with specification
' PSC   28/02/2007  EP2_1410 Get Dependants Income factor correctly
' AW    04/04/2007  EP2_2250 Corrected logic around liability end dates
'---------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXMLAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing
End Sub

Public Function CalcAllowableInc(vstrXMLRequest As String) As String

'On Error GoTo CalcAllowableIncVbErr

    'Const strFunctionName As String = "CalcAllowableInc"

    '
    ' JD MAR688 CalcAllowableInc NO LONGER CALLED
    '
    
'    Dim xmlIn As FreeThreadedDOMDocument40
'    Dim xmlOut As New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'    Dim xmlResponseElem As IXMLDOMElement, xmlNewRequest As IXMLDOMNode, xmlTable As IXMLDOMElement
'    Dim xmlElement As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode, xmlMortgageLender As IXMLDOMNode
'    Dim xmlAllowableIncFactors As IXMLDOMNode, xmlTempResponseElem As IXMLDOMNode
'
'    Dim strOrganisationId As String, strResponse As String
'
'    Dim objMortgageLenderBO As MortgageLenderBO
'
'    ' Create default response block
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
'
'    strOrganisationId = m_objXMLAssist.GetTagValue(xmlIn.documentElement, ".//ORGANISATIONID")
'
'    If Len(strOrganisationId) = 0 Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 557, "ORGANISATIONID"
'    End If
'
'    'Get MortgageLender record
'    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(xmlIn.documentElement)
'    Set xmlTable = xmlOut.createElement("MORTGAGELENDER")
'    xmlNewRequest.appendChild xmlTable
'
'    Set xmlElement = xmlOut.createElement("ORGANISATIONID")
'    xmlElement.Text = strOrganisationId
'    xmlTable.appendChild xmlElement
'
'    If Not m_objContext Is Nothing Then
'        Set objMortgageLenderBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
'    Else
'        Set objMortgageLenderBO = New MortgageLenderBO
'    End If
'
'    strResponse = objMortgageLenderBO.GetLenderDetails(xmlNewRequest.xml)
'    Set xmlOut = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
'
'    Set xmlMortgageLender = xmlOut.selectSingleNode(".//MORTGAGELENDER")
'
'    'Get Allowable Income Factors
'    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(xmlIn.documentElement)
'    Set xmlTable = xmlOut.createElement("ALLOWABLEINCOMEFACTORS")
'    xmlNewRequest.appendChild xmlTable
'
'    Set xmlElement = xmlOut.createElement("ORGANISATIONID")
'    xmlElement.Text = strOrganisationId
'    xmlTable.appendChild xmlElement
'
'    strResponse = objMortgageLenderBO.FindAllowableIncomeFactors(xmlNewRequest.xml)
'    Set xmlOut = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
'
'    Set xmlAllowableIncFactors = xmlOut.selectSingleNode("//ALLOWABLEINCOMEFACTORSLIST")
'
'    'Append OrganisationId to Response
'    Set xmlElement = xmlOut.createElement("ORGANISATIONID")
'    xmlElement.Text = strOrganisationId
'    xmlResponseElem.appendChild xmlElement
'
'    ' JD MAR688 CalcAllowableInc no longer called   Set xmlTempResponseElem = CalculateAllowableIncome(xmlIn.documentElement, xmlAllowableIncFactors, xmlMortgageLender)
'    Debug.Print xmlTempResponseElem.xml
'    m_objErrAssist.CheckResponse xmlTempResponseElem, xmlResponseElem, True
'
'    CalcAllowableInc = xmlTempResponseElem.xml
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'CalcAllowableIncExit:
'
'    Set xmlIn = Nothing
'    Set xmlOut = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlNewRequest = Nothing
'    Set xmlTable = Nothing
'    Set xmlElement = Nothing
'    Set xmlTempResponseNode = Nothing
'    Set xmlMortgageLender = Nothing
'    Set xmlAllowableIncFactors = Nothing
'    Set xmlTempResponseElem = Nothing
'
'    Exit Function
'
'CalcAllowableIncVbErr:
'
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'
'    m_objErrAssist.AddToErrSource strFunctionName
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    CalcAllowableInc = m_objErrAssist.CreateErrorResponse
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' go to clean-up section
'    Resume CalcAllowableIncExit

End Function


Public Function CalculateAllowableIncome(strCalcAllowInputs As String, strCalcOutgoingsInputs As String) As String

On Error GoTo CalculateAllowableIncomeVbErr

    Const strFunctionName As String = "CalculateAllowableIncome"

    Dim xmlCustomerList As IXMLDOMNodeList, xmlCustomer As IXMLDOMNode
    Dim xmlUnearnedIncomeList As IXMLDOMNodeList, xmlUnearnedIncome As IXMLDOMNode
    Dim xmlEmploymentList As IXMLDOMNodeList, xmlEmployment As IXMLDOMNode
    Dim xmlRespCustomers As IXMLDOMElement, xmlRespCustomer As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement, xmlNodeList As IXMLDOMNodeList, xmlFactorElem As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40, xmlOut As FreeThreadedDOMDocument40
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlCalcOutgoing As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Dim xmlCalcOutgoingsNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    
    Dim objComboDO As ComboDO
    '   AW  21/10/02    BMIDS00653
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim dblAdditionalIncRatio As Double
    '   AW    BMIDS00653 -End
    
    '   AW  21/10/02    BMIDS00653
    Dim dblTotalRegOutgoings As Double, dblTotalLoansLiabilities As Double
    
    'BMIDS00653 MDC 28/10/2002
    Dim objCustEmpBO As CustomerEmploymentBO
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlFieldNode As IXMLDOMNode
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim strResponse As String
    'BMIDS00653 MDC 28/10/2002 - End
    
    'BM0376
    'Dim strExcludeGuarantors As String
    Dim blnIncludeGuarantors As Boolean
    'BM0376
    Dim strCustRoleType As String
    Dim strCondition As String, strMainStatus As String, strEmploymentStatus As String
    Dim strType As String
    Dim dblAllowableIncome As Double, dblFactor As Double
    Dim dblAmount As Double, dblFrequency As Double
    Dim dblYear1Amount As Double, dblYear2Amount As Double, dblYear3Amount As Double
    '   AW  21/10/02    BMIDS00653
    Dim dbNetlAllowableIncome As Double
    Dim xmlOutGoingsElem As IXMLDOMElement
    
    'BMIDS00653 MDC 05/11/2002
    Dim strEmployedConfirmedIncomeGroup As String
    Dim strEmployedDeclaredIncomeGroup As String
    Dim strSelfEmployedConfirmedIncomeGroup As String
    Dim strSelfEmployedDeclaredIncomeGroup As String
    Dim strEmploymentSequenceNumber As String 'JD MAR688
    'Dim strUnearnedIncomeGroup As String   'BMIDS00861 MDC 06/11/2002
    Dim strIncomeType As String
    Dim dblIncome As Double
    Dim dblIncome1 As Double
    Dim dblIncome2 As Double
    Dim dblIncome3 As Double
    Dim xmlConfirmedEarnedIncomeList As IXMLDOMNodeList
    Dim xmlConfirmedEarnedIncome As IXMLDOMElement
    Dim xmlEmployersRef As IXMLDOMNode
    Dim blnIncludeNonGuaranteed As Boolean
    Dim dblAdditionalIncome As Double
    Dim dblBasicIncome As Double
    Dim bExcludeOtherIncome As Boolean  ' JD MAR30
    Dim dblConfirmedAllowableIncome As Double  'JD MAR30
    Dim dblNetConfirmedAllowableIncome As Double 'JD MAR30
    Dim strUnderwriterOverrideIncludeOtherInc As String  'JD MAR30
    Dim strUnderwriterOtherIncomePercentage As String  'JD MAR30
    'BMIDS00653 MDC 05/11/2002 - End
    Dim bUsedConfirmedDetails As Boolean 'JD MAR1040
    Dim dblConfirmedBasicInc As Double 'JD MAR1040
    Dim dblConfirmedAdditionalInc As Double ' JD MAR1040
    Dim dblTotalBalanceMortgages As Double
    Dim dblTotalBalanceLoansLiabilities As Double 'JD MAR688
    Dim dblTotalMortgageOutgoings As Double ' JD MAR30
    Dim dblTotalDependantsOutgoings As Double 'EP2_128
    
    Dim intNoOfYearsAccounts As Integer 'BMIDS00863 MDC 07/11/2002
    
    'MAR1891 GHun
    Dim blnIncomeSummaryExists As Boolean
    'XML Elements for caching combo groups
    Dim xmlIncomeTypeCombo As IXMLDOMElement
    Dim xmlConfirmedIncomeTypeCombo As IXMLDOMElement
    Dim xmlEmploymentStatusCombo As IXMLDOMElement
    'MAR1891 End
    Dim bUseAllEmployments As Boolean 'JD MAR1919

    'EP2_128 --------------------------------
    Dim blnAFAdditionalIncomeCheck As Boolean
    Dim blnAFDBAffordability As Boolean
    Dim intAFAllowableIncPercentage As Integer
    'End EP2_128 -----------------------------

    Set xmlDoc = xmlCreateDOMObject()

    Set xmlIn = xmlLoad(strCalcAllowInputs, strFunctionName)
    Set xmlRequest = xmlIn.selectSingleNode("RESPONSE")
    
    Set xmlCalcOutgoing = xmlCreateDOMObject()
    xmlCalcOutgoing.loadXML strCalcOutgoingsInputs
    Set xmlCalcOutgoingsNode = xmlCalcOutgoing.selectSingleNode("RESPONSE")
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlElement = xmlOut.createElement("LENDERDETAILS")
    xmlResponseElem.appendChild xmlElement
    
    Set xmlTempNode = xmlRequest.selectNodes(".//ALLOWABLEINCOMEFACTORS").Item(0)
    
    Set xmlRespCustomers = xmlOut.createElement("CUSTOMERLIST")
    xmlElement.appendChild xmlRespCustomers
    
    Set xmlCustomerList = xmlRequest.selectNodes(".//CUSTOMER")
    
    If Not m_objContext Is Nothing Then
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    Else
        Set objComboDO = New ComboDO
        Set objGlobalParameterDO = New GlobalParameterDO
    End If
    
    'BM0376 Use IncludeGuarantors global parameter instead
    'strExcludeGuarantors = m_objXMLAssist.GetTagValue(xmlMortgageLender, "EXCLUDEGUARANTORS")
    blnIncludeGuarantors = CSafeBool(objGlobalParameterDO.GetCurrentParameterByType("IncludeGuarantors", "BOOLEAN"))
    'BM0376
    
    'JD MAR30 get global parameter ExcludeOtherIncomeUnlessUnderwriter
    bExcludeOtherIncome = CSafeBool(objGlobalParameterDO.GetCurrentParameterByType("ExcludeOthIncUnlessUnderwriter", "BOOLEAN"))
    
    'JD MAR1919 get global parameter UseAllEmployments
    bUseAllEmployments = CSafeBool(objGlobalParameterDO.GetCurrentParameterByType("UseAllEmployments", "BOOLEAN"))
    
    'EP2_128
    blnAFAdditionalIncomeCheck = CSafeBool(objGlobalParameterDO.GetCurrentParameterByType("AFAdditionalIncomeCheck", "BOOLEAN"))
    blnAFDBAffordability = CSafeBool(objGlobalParameterDO.GetCurrentParameterByType("AFDBAffordability", "BOOLEAN"))
    intAFAllowableIncPercentage = CSafeInt(objGlobalParameterDO.GetCurrentParameterByType("AFAllowableIncPercentage", "PERCENTAGE"))
    
    'BMIDS00863 MDC 07/11/2002
'    strEmployedID = objComboDO.GetFirstComboValueId("EmploymentStatus", "E")
'    strSelfEmployedID = objComboDO.GetFirstComboValueId("EmploymentStatus", "S")
'    strContractID = objComboDO.GetFirstComboValueId("EmploymentStatus", "C")
    'BMIDS00863 MDC 07/11/2002 - End
    'BMIDS00653 MDC 05/11/2002
    strEmployedConfirmedIncomeGroup = objComboDO.GetFirstComboValueId("IncomeGroupType", "EC")
    strEmployedDeclaredIncomeGroup = objComboDO.GetFirstComboValueId("IncomeGroupType", "ED")
    strSelfEmployedConfirmedIncomeGroup = objComboDO.GetFirstComboValueId("IncomeGroupType", "SC")
    strSelfEmployedDeclaredIncomeGroup = objComboDO.GetFirstComboValueId("IncomeGroupType", "SD")
    'strUnearnedIncomeGroup = objComboDO.GetFirstComboValueId("IncomeGroupType", "O")   'BMIDS00861 MDC 06/11/2002
    'BMIDS00653 MDC 05/11/2002 - End
    
    'MAR1891 GHun
    'Retrieve global parameter outside of the loop
    dblAdditionalIncRatio = objGlobalParameterDO.GetCurrentParameterByType("AdditionalIncomeRatio", "AMOUNT")

    Set xmlTempRequestNode = xmlRequest.cloneNode(False)    'Moved outside of loop
    'Cache frequently used combo groups in xml elements, to prevent multiple database calls
    Set xmlIncomeTypeCombo = xmlRequest.selectSingleNode("COMBO[@GROUPNAME='IncomeType']")
    Set xmlConfirmedIncomeTypeCombo = xmlRequest.selectSingleNode("COMBO[@GROUPNAME='ConfirmedIncomeType']")
    Set xmlEmploymentStatusCombo = xmlRequest.selectSingleNode("COMBO[@GROUPNAME='EmploymentStatus']")
    'MAR1891 End
    
    ' PSC 26/01/2007 EP2_976
    ' Calculate dependant outgoings per customer
    ' PSC 28/02/2007 EP2_1410
    dblTotalDependantsOutgoings = xmlGetNodeAsDouble(xmlRequest, "DEPENDANTSINCOMEFACTOR/@AMOUNT") / xmlCustomerList.length

    For Each xmlCustomer In xmlCustomerList
        'GD BM0360 19/02/03 START
        'Initialise FOR EACH customer
        intNoOfYearsAccounts = 0
        dblIncome = 0
        'GD BM0360 19/02/03 END
    
        strCustRoleType = xmlGetAttributeText(xmlCustomer, "CUSTOMERROLETYPE")
        'BMIDS00653 MDC 06/11/2002
        dblBasicIncome = 0
        dblAdditionalIncome = 0
        'BMIDS00653 MDC 06/11/2002
        dblConfirmedBasicInc = 0        'JD MAR1040
        dblConfirmedAdditionalInc = 0   'JD MAR1040
        
        dblAllowableIncome = 0
        dbNetlAllowableIncome = 0
        'BMIDS00653 MDC 28/10/2002
        dblTotalRegOutgoings = 0
        dblTotalLoansLiabilities = 0
        'BMIDS00653 MDC 28/10/2002 - End
        dblTotalMortgageOutgoings = 0  'JD MAR30
        
        'JD MAR30
        dblConfirmedAllowableIncome = 0
        bUsedConfirmedDetails = False  'JD MAR1040
        blnIncomeSummaryExists = False  'MAR1891 GHun
        
        Set xmlTempNode = xmlCustomer.selectSingleNode(".//INCOMESUMMARY")
        If Not xmlTempNode Is Nothing Then
            strUnderwriterOverrideIncludeOtherInc = xmlGetAttributeText(xmlTempNode, "UNDERWRITEROVERRIDEINCLUDEOTHERINC")
            strUnderwriterOtherIncomePercentage = xmlGetAttributeText(xmlTempNode, "UNDERWRITEROTHERINCOMEPERCENTAGE")
            blnIncomeSummaryExists = (xmlGetAttributeText(xmlTempNode, "EXISTS") = 1)   'MAR1891 GHun
        End If
        
        ' Exclude Guarantors
        'BMIDS00653 MDC 05/11/2002
        'If strCustRoleType = "2" And strExcludeGuarantors = "1" Then
        'BM0376
        'If objComboDO.IsItemInValidation("CustomerRoleType", strCustRoleType, "G") And strExcludeGuarantors = "1" Then
        If objComboDO.IsItemInValidation("CustomerRoleType", strCustRoleType, "G") And Not blnIncludeGuarantors Then
        'BM0376 End
        'BMIDS00653 MDC 05/11/2002- End
            '================================================================================================
            ' Exclude Guarantor
            '================================================================================================
            Set xmlRespCustomer = xmlDoc.createElement("CUSTOMER")
            Set xmlElement = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlElement.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERNUMBER")
            xmlRespCustomer.appendChild xmlElement
            Set xmlElement = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlElement.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERVERSIONNUMBER")
            xmlRespCustomer.appendChild xmlElement
            'xmlRespCustomer.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
            'xmlRespCustomer.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
                        
            Set xmlElement = xmlOut.createElement("MONTHLYALLOWABLEINCOME")
            xmlElement.Text = 0
            xmlRespCustomer.appendChild xmlElement
            
            Set xmlElement = xmlOut.createElement("ANNUALALLOWABLEINCOME")
            xmlElement.Text = 0
            xmlRespCustomer.appendChild xmlElement
            
            '   AW  21/10/02    BMIDS00653
            Set xmlElement = xmlOut.createElement("NETANNUALALLOWABLEINCOME")
            xmlElement.Text = 0
            xmlRespCustomer.appendChild xmlElement
        
        Else
            'BMIDS00861 MDC 06/11/2002 - Exclude Unearned Income
'            '================================================================================================
'            ' Unearned Incomes
'            '================================================================================================
'            Set xmlUnearnedIncomeList = xmlCustomer.selectNodes(".//UNEARNEDINCOME")
'            For Each xmlUnearnedIncome In xmlUnearnedIncomeList
'                strUnearnedIncomeType = xmlUnearnedIncome.selectSingleNode(".//UNEARNEDINCOMETYPE").Text
'
'                'BMIDS00653 MDC 05/11/2002
'                'strCondition = ".//ALLOWABLEINCOMEFACTORS[INCOMEGROUP='5' and TYPE='" & strUnearnedIncomeType & "']"
'                strCondition = ".//ALLOWABLEINCOMEFACTORS[INCOMEGROUP='" & strUnearnedIncomeGroup & _
'                                                                "' and TYPE='" & strUnearnedIncomeType & "']"
'                'BMIDS00653 MDC 05/11/2002 - End
'                Set xmlFactorElem = xmlAllowableIncFactors.selectSingleNode(strCondition)
'
'                If xmlFactorElem Is Nothing Then
'                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 307, "IncomeGroup=5 and Type=" & strUnearnedIncomeType
'                Else
'                    dblUnearnedIncomeAmount = CSafeDbl(m_objXMLAssist.GetTagValue(xmlUnearnedIncome, "UNEARNEDINCOMEAMOUNT"))
'                    dblFactor = CSafeDbl(m_objXMLAssist.GetTagValue(xmlFactorElem, "FACTOR"))
'                    dblFactor = dblFactor / 100
'                    dblFrequency = CSafeDbl(m_objXMLAssist.GetTagValue(xmlUnearnedIncome, "PAYMENTFREQUENCY"))
'
'                    dblAllowableIncome = dblAllowableIncome + (dblUnearnedIncomeAmount * dblFactor * dblFrequency)
'                End If
'            Next xmlUnearnedIncome
            'BMIDS00861 MDC 06/11/2002 - End
            
            '================================================================================================
            'Income through Employment
            '================================================================================================
            Set xmlEmploymentList = xmlCustomer.selectNodes(".//EMPLOYMENT")
            For Each xmlEmployment In xmlEmploymentList
                strMainStatus = xmlGetAttributeText(xmlEmployment, "MAINSTATUS")
                'strUseAllEmployments = m_objXMLAssist.GetTagValue(xmlEmployment, "USEALLEMPLOYMENTS")
                'BMIDS00886 AW 13/11/2002
                strEmploymentStatus = xmlGetAttributeText(xmlEmployment, "EMPLOYMENTSTATUS")
                strEmploymentSequenceNumber = xmlGetAttributeText(xmlEmployment, "EMPLOYMENTSEQUENCENUMBER")

                'BMIDS00653 MDC 29/10/2002
                'If strUseAllEmployments = "0" And strMainStatus = "0" Then
                'JD MAR1919 include check for global UseAllEmployments
                If bUseAllEmployments = False And _
                    CSafeBool(strMainStatus) = False And _
                                IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "R") = False Then 'BMIDS00886 AW 13/11/2002
                'BMIDS00653 MDC 29/10/2002 - End
                    ' Consider next employment
                
                Else
                    'strEmploymentStatus = xmlGetAttributeText(xmlEmployment, "EMPLOYMENTSTATUS")
                    'BMIDS00865 MDC 07/11/2002 - Include ALL Employed status
                    'If strEmploymentStatus = strEmployedID Then
                    If IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "E") Or _
                        IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "R") Then     'BMIDS00886 AW 13/11/2002
                    'BMIDS00865 MDC 07/11/2002 - End
                    
                        '====================================================================================
                        ' Employed
                        '====================================================================================
                        Set xmlEmployersRef = m_objXMLAssist.GetNode(xmlCustomer, "CURRENTEMPLOYERSREF[@EMPLOYMENTSEQUENCENUMBER=" + strEmploymentSequenceNumber + "]", False)
                        If Not xmlEmployersRef Is Nothing Then
                            
                            '====================================================================================
                            ' Employed Confirmed Income
                            '====================================================================================
                            bUsedConfirmedDetails = True 'JD MAR1040
                            blnIncludeNonGuaranteed = CSafeBool(m_objXMLAssist.GetAttributeValue(xmlEmployersRef, "CURRENTEMPLOYERSREF", "NONGUARANTEEDPAYMENTSIND"))
                            
                            Set xmlConfirmedEarnedIncomeList = xmlCustomer.selectNodes(".//CONFIRMEDEARNEDINCOME[@EMPLOYMENTSEQUENCENUMBER=" + strEmploymentSequenceNumber + "]")
                            For Each xmlConfirmedEarnedIncome In xmlConfirmedEarnedIncomeList
                                'Confirmed Earned Income
                                strIncomeType = m_objXMLAssist.GetMandatoryAttribute(xmlConfirmedEarnedIncome, "EARNEDINCOMETYPE")
                                strCondition = ".//ALLOWABLEINCOMEFACTORS[@INCOMEGROUP='" & strEmployedConfirmedIncomeGroup & _
                                                                                "' and @TYPE='" & Trim$(strIncomeType) & "']"
                                
                                'Ignore non guaranteed income if not confirmed by employer
                                If (IsItemInValidationXML(xmlConfirmedIncomeTypeCombo, strIncomeType, "RO") Or _
                                      IsItemInValidationXML(xmlConfirmedIncomeTypeCombo, strIncomeType, "RB")) And _
                                      Not blnIncludeNonGuaranteed Then
                                    'Ignore this income
                                ElseIf (bExcludeOtherIncome And _
                                        CSafeBool(strUnderwriterOverrideIncludeOtherInc) <> True And _
                                        IsItemInValidationXML(xmlConfirmedIncomeTypeCombo, strIncomeType, "OTH")) Then
                                    'Ignore this income as the underwriter has chosen to not include other income type.
                                Else
                                    'Include this income
                                    dblFactor = 0
                                    'JD MAR30 Use the Other income percentage as entered by the underwriter if present
                                    If CSafeBool(strUnderwriterOverrideIncludeOtherInc) = True And _
                                        IsItemInValidationXML(xmlConfirmedIncomeTypeCombo, strIncomeType, "OTH") Then
                                        If CSafeDbl(strUnderwriterOtherIncomePercentage) > 0 Then
                                            dblFactor = CSafeDbl(strUnderwriterOtherIncomePercentage)
                                        End If
                                    End If
                                        
                                    Set xmlFactorElem = xmlRequest.selectSingleNode(strCondition)
                                    If Not xmlFactorElem Is Nothing Then
                                        If dblFactor = 0 Then 'JD MAR30 if we aren't using the underwriters factor use the allowable income factor
                                            dblFactor = xmlGetAttributeAsDouble(xmlFactorElem, "FACTOR")
                                        End If
                                        dblFactor = dblFactor / 100
                                        dblAmount = xmlGetAttributeAsDouble(xmlConfirmedEarnedIncome, "EARNEDINCOMEAMOUNT", "0")
                                        dblFrequency = xmlGetAttributeAsDouble(xmlConfirmedEarnedIncome, "PAYMENTFREQUENCYTYPE", "0")
                                        dblIncome = (dblAmount * dblFrequency) * dblFactor
                                        'dblAllowableIncome = dblAllowableIncome + dblIncome  JD MAR1040 use Confirmed var so incomesummary is not overwritten
                                        dblConfirmedAllowableIncome = dblConfirmedAllowableIncome + dblIncome
                                        
                                        If IsItemInValidationXML(xmlConfirmedIncomeTypeCombo, strIncomeType, "B") Then
                                            'Basic Income
                                            dblConfirmedBasicInc = dblConfirmedBasicInc + dblIncome
                                        Else
                                            'Additional Income
                                            dblConfirmedAdditionalInc = dblConfirmedAdditionalInc + dblIncome
                                        End If
                                    Else
                                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 307, "IncomeGroup=" & _
                                                                    strEmployedConfirmedIncomeGroup & " and Type=" & strIncomeType
                                    End If
                                End If
                                
                            Next xmlConfirmedEarnedIncome
                            'JD MAR30 save this as confirmed allowable income
                            'JD MAR1040 unneccessary now
                            'dblConfirmedAllowableIncome = dblAllowableIncome
                        'Else
                        End If  'JD MAR1040 Do declared as well as confirmed
                            '====================================================================================
                            ' Employed Declared Income
                            '====================================================================================
                            Set xmlNodeList = xmlCustomer.selectNodes(".//DECLAREDEMPEARNINGS[@EMPLOYMENTSEQUENCENUMBER=" + strEmploymentSequenceNumber + "]")
                            For Each xmlElement In xmlNodeList
                                strType = xmlGetAttributeText(xmlElement, "EARNEDINCOMETYPE")
                                
                                strCondition = ".//ALLOWABLEINCOMEFACTORS[@INCOMEGROUP='" & strEmployedDeclaredIncomeGroup & _
                                                                                        "' and @TYPE='" & strType & "']"
                                'JD MAR30 check underwriter override
                                If (bExcludeOtherIncome And _
                                    CSafeBool(strUnderwriterOverrideIncludeOtherInc) <> True And _
                                    IsItemInValidationXML(xmlIncomeTypeCombo, strType, "OTH")) Then
                                    'Ignore this income as the underwriter has chosen to not include other income type.
                                Else
                                    dblFactor = 0
                                    'JD MAR30 Use the Other income percentage as entered by the underwriter if present
                                    If CSafeBool(strUnderwriterOverrideIncludeOtherInc) = True And _
                                        IsItemInValidationXML(xmlIncomeTypeCombo, strType, "OTH") Then
                                        If CSafeDbl(strUnderwriterOtherIncomePercentage) > 0 Then
                                            dblFactor = CSafeDbl(strUnderwriterOtherIncomePercentage)
                                        End If
                                    End If
                                    Set xmlFactorElem = xmlRequest.selectSingleNode(strCondition)
                                    
                                    If Not xmlFactorElem Is Nothing Then
                                        If dblFactor = 0 Then
                                            'JD MAR30 if underwriter percentage not used use factor
                                            dblFactor = xmlGetAttributeAsDouble(xmlFactorElem, "FACTOR")
                                        End If
                                        dblFactor = dblFactor / 100
                                        dblAmount = xmlGetAttributeAsDouble(xmlElement, "EARNEDINCOMEAMOUNT")
                                        dblFrequency = xmlGetAttributeAsDouble(xmlElement, "PAYMENTFREQUENCYTYPE")
                                        
                                        dblIncome = (dblAmount * dblFrequency) * dblFactor
                                        dblAllowableIncome = dblAllowableIncome + dblIncome
                                        'dblAllowableIncome = dblAllowableIncome + (dblAmount * dblFactor * dblFrequency)
                                        
                                        If IsItemInValidationXML(xmlIncomeTypeCombo, strType, "B") Then
                                            'Basic Income
                                            dblBasicIncome = dblBasicIncome + dblIncome
                                        Else
                                            'Additional Income
                                            dblAdditionalIncome = dblAdditionalIncome + dblIncome
                                        End If
                                    Else
                                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 307, "IncomeGroup=" & strEmployedDeclaredIncomeGroup & " and Type=" & strType
                                    End If
                                End If
                            Next xmlElement
                        'End If  JD MAR1040
                        
                    'BMIDS00863 MDC 07/11/2002
                    'ElseIf strEmploymentStatus = strSelfEmployedID Or strEmploymentStatus = strContractID Then
                    ElseIf IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "S") Or _
                            IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "C") Then
                    'BMIDS00863 MDC 07/11/2002 - End
                        '====================================================================================
                        ' Self Employed
                        '====================================================================================
                        dblIncome = 0               'LDM 17/8/2006 EP1087 initialise for running totals used for declared and confirmed
                        intNoOfYearsAccounts = 0    '//
                        Set xmlElement = xmlCustomer.selectSingleNode(".//CONFIRMEDSELFEMPEARNINGS[@EMPLOYMENTSEQUENCENUMBER=" + strEmploymentSequenceNumber + "]")
                        'BMIDS00863 MDC 06/11/2002
                        'strYearsAccountForSelfEmp = m_objXmlAssist.GetTagValue(xmlMortgageLender, "YEARSACCOUNTFORSELFEMP")
                        'BMIDS00863 MDC 06/11/2002 - End
                        If Not xmlElement Is Nothing Then
                            '====================================================================================
                            ' Self Employed Confirmed Income
                            '====================================================================================
                            
                            'BMIDS00653 MDC 05/11/2002
                            dblFactor = 0
                            If IsItemInValidationXML(xmlEmploymentStatusCombo, strEmploymentStatus, "Z33") Then
                                'Self Employed as a director with shares > 33%
                                strIncomeType = objComboDO.GetFirstComboValueId("SelfEmpConfirmedIncomeType", "SD")  'Salary & Dividends
                                dblIncome1 = CSafeDbl(xmlGetAttributeText(xmlElement, "SALARY1")) _
                                                + CSafeDbl(xmlGetAttributeText(xmlElement, "DIVIDEND1"))
                                dblIncome2 = CSafeDbl(xmlGetAttributeText(xmlElement, "SALARY2")) _
                                                + CSafeDbl(xmlGetAttributeText(xmlElement, "DIVIDEND2"))
                                dblIncome3 = CSafeDbl(xmlGetAttributeText(xmlElement, "SALARY3")) _
                                                + CSafeDbl(xmlGetAttributeText(xmlElement, "DIVIDEND3"))
                            Else
                                'Self Employed (Sole Trader or Partnership)
                                strIncomeType = objComboDO.GetFirstComboValueId("SelfEmpConfirmedIncomeType", "ANP")  'Average Net Profit
                                dblIncome1 = CSafeDbl(xmlGetAttributeText(xmlElement, "NETPROFITENTITLEMENT1"))
                                dblIncome2 = CSafeDbl(xmlGetAttributeText(xmlElement, "NETPROFITENTITLEMENT2"))
                                dblIncome3 = CSafeDbl(xmlGetAttributeText(xmlElement, "NETPROFITENTITLEMENT3"))
                            End If
                            
                            'EP1061
                            strCondition = ".//ALLOWABLEINCOMEFACTORS[@INCOMEGROUP='" & strSelfEmployedConfirmedIncomeGroup & "' and @TYPE='" _
                                                                            & strIncomeType & "']"
                            Set xmlFactorElem = xmlRequest.selectSingleNode(strCondition)
                            If Not xmlFactorElem Is Nothing Then
                                dblFactor = CSafeDbl(xmlGetAttributeAsDouble(xmlFactorElem, "FACTOR"))
                                dblFactor = dblFactor / 100
                            Else
                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 307, "IncomeGroup=" & _
                                                            strSelfEmployedConfirmedIncomeGroup & " and Type=" & strIncomeType
                            End If
                                
                            'BMIDS00863 MDC 07/11/2002
                            If dblIncome1 > 0 Then
                                dblIncome = dblIncome + dblIncome1
                                intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                            End If
                            If dblIncome2 > 0 Then
                                dblIncome = dblIncome + dblIncome2
                                intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                            End If
                            If dblIncome3 > 0 Then
                                dblIncome = dblIncome + dblIncome3
                                intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                            End If
                            If intNoOfYearsAccounts < 1 Then
                                intNoOfYearsAccounts = 1
                            End If
                            'JD MAR1912 set confirmed allowable income
                            bUsedConfirmedDetails = True
                            dblConfirmedAllowableIncome = dblConfirmedAllowableIncome + ((dblIncome * dblFactor) / intNoOfYearsAccounts)
'                            If strYearsAccountForSelfEmp = "1" Or strYearsAccountForSelfEmp = "" Then
'                                dblAmount = dblIncome1 * dblFactor
'                            ElseIf strYearsAccountForSelfEmp = "2" Then
'                                dblAmount = ((dblIncome1 + dblIncome2) * dblFactor) / 2
'                            ElseIf strYearsAccountForSelfEmp = "3" Then
'                                dblAmount = ((dblIncome1 + dblIncome2 + dblIncome3) * dblFactor) / 3
'                            End If
'                            dblAllowableIncome = dblAllowableIncome + dblAmount
                            'BMIDS00863 MDC 07/11/2002 - End
                            'BMIDS00653 MDC 05/11/2002 - End

                        'Else JD MAR1912 Do declared as well as confirmed
                        End If
                            '====================================================================================
                            ' Self Employed Declared Income
                            '====================================================================================
                            dblIncome = 0               'LDM 17/8/2006 EP1087 initialise for running totals used for declared and confirmed
                            intNoOfYearsAccounts = 0    '//
                            Set xmlElement = xmlCustomer.selectSingleNode(".//NETPROFIT[@EMPLOYMENTSEQUENCENUMBER=" + strEmploymentSequenceNumber + "]")
                            If Not xmlElement Is Nothing Then
                                'BMIDS00653 MDC 05/11/2002
                                strIncomeType = objComboDO.GetFirstComboValueId("SelfEmpConfirmedIncomeType", "ANP")  'Average Net Profit
                                'EP1061
                                strCondition = ".//ALLOWABLEINCOMEFACTORS[@INCOMEGROUP='" & strSelfEmployedDeclaredIncomeGroup & _
                                                                                            "' and @TYPE='" & strIncomeType & "']"
                                Set xmlFactorElem = xmlRequest.selectSingleNode(strCondition)
                                
                                If xmlFactorElem Is Nothing Then
                                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 307, "IncomeGroup=" & _
                                                                strSelfEmployedDeclaredIncomeGroup & " and Type=" & strIncomeType
                                Else
                                    dblFactor = CSafeDbl(xmlGetAttributeAsDouble(xmlFactorElem, "FACTOR"))
                                    dblFactor = dblFactor / 100
                                    dblYear1Amount = CSafeDbl(xmlGetAttributeText(xmlElement, "YEAR1AMOUNT"))
                                    dblYear2Amount = CSafeDbl(xmlGetAttributeText(xmlElement, "YEAR2AMOUNT"))
                                    dblYear3Amount = CSafeDbl(xmlGetAttributeText(xmlElement, "YEAR3AMOUNT"))
                                    
                                    
                                    'BMIDS00863 MDC 07/11/2002
                                    If dblYear1Amount > 0 Then
                                        dblIncome = dblIncome + dblYear1Amount
                                        intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                                    End If
                                    If dblYear2Amount > 0 Then
                                        dblIncome = dblIncome + dblYear2Amount
                                        intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                                    End If
                                    If dblYear3Amount > 0 Then
                                        dblIncome = dblIncome + dblYear3Amount
                                        intNoOfYearsAccounts = intNoOfYearsAccounts + 1
                                    End If
                                    If intNoOfYearsAccounts < 1 Then
                                        intNoOfYearsAccounts = 1
                                    End If
                                    dblAllowableIncome = dblAllowableIncome + ((dblIncome * dblFactor) / intNoOfYearsAccounts)
'                                    If strYearsAccountForSelfEmp = "1" Or strYearsAccountForSelfEmp = "" Then
'                                        dblAllowableIncome = dblAllowableIncome + (dblYear1Amount * dblFactor)
'                                    ElseIf strYearsAccountForSelfEmp = "2" Then
'                                        dblAmount = (dblYear1Amount + dblYear2Amount) * dblFactor / 2
'                                        dblAllowableIncome = dblAllowableIncome + dblAmount
'                                    ElseIf strYearsAccountForSelfEmp = "3" Then
'                                        dblAmount = (dblYear1Amount + dblYear2Amount + dblYear3Amount) * dblFactor / 3
'                                        dblAllowableIncome = dblAllowableIncome + dblAmount
'                                    End If
                                    'BMIDS00863 MDC 07/11/2002 - End
                                End If
                                'BMIDS00653 MDC 05/11/2002 - End
                            End If
                        'End If  JD MAR1912
                    End If
                End If
            Next xmlEmployment
        End If
        
        '===============================================================================================
        'Check Additional Income Ratio (MAR1040) for confirmed and declared
        '===============================================================================================
        'EP2_128
        If blnAFAdditionalIncomeCheck Then
            If dblBasicIncome > 0 Then
                If dblAdditionalIncome / dblBasicIncome > dblAdditionalIncRatio Then
                    dblAllowableIncome = dblAllowableIncome - (dblAdditionalIncome / 2)
                End If
            End If
            If dblConfirmedBasicInc > 0 Then
                'JD MAR30 alter confirmed allowable inc if necessary
                If dblConfirmedAdditionalInc / dblConfirmedBasicInc > dblAdditionalIncRatio Then
                    dblConfirmedAllowableIncome = dblConfirmedAllowableIncome - (dblConfirmedAdditionalInc / 2)
                End If
            End If
        End If
        
        'BMIDS00653 MDC 28/10/2002
        Set xmlOutGoingsElem = xmlDoc.createElement("REQUEST")
        'Set xmlRespCustomer = m_objXMLAssist.GetMandatoryNode(xmlRequest, "CUSTOMERLIST").cloneNode(True)
        Set xmlRespCustomer = xmlDoc.createElement("CUSTOMERLIST")
        Dim xmlTempNodeList As IXMLDOMNodeList
        Dim xmlCustListNode As IXMLDOMNode
        Set xmlTempNodeList = xmlRequest.selectNodes(".//CUSTOMER")
        For Each xmlCustListNode In xmlTempNodeList
            xmlRespCustomer.appendChild xmlCustListNode.cloneNode(True)
        Next
        xmlOutGoingsElem.appendChild xmlRespCustomer
        Call CalculateOutGoings(xmlCalcOutgoingsNode, xmlCustomer, dblTotalRegOutgoings, _
                                 dblTotalLoansLiabilities, dblTotalMortgageOutgoings, dblTotalBalanceMortgages, dblTotalBalanceLoansLiabilities, blnAFDBAffordability)
        
        ' create REQUEST details
'        Set xmlOutGoingsElem = xmlDoc.createElement("REQUEST")
'        Set vxmlOutGoingsRequest = xmlDoc.appendChild(xmlOutGoingsElem)
        Dim xmlTempElem As IXMLDOMElement
        Set xmlRespCustomer = xmlOut.createElement("CUSTOMER")
        Set xmlTempElem = xmlOut.createElement("CUSTOMERNUMBER")
        xmlTempElem.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERNUMBER")
        xmlRespCustomer.appendChild xmlTempElem
        Set xmlTempElem = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlTempElem.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        xmlRespCustomer.appendChild xmlTempElem
        'm_objXMLAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomer, xmlRespCustomer, False
        'm_objXMLAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomer, xmlRespCustomer, False
'        vxmlOutGoingsRequest.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
'        vxmlOutGoingsRequest.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
'        Call CalculateOutGoings(vxmlOutGoingsRequest, dblTotalRegOutgoings, _
                                  dblTotalLoansLiabilities)
        
        'BMIDS00653 MDC 28/10/2002 - End
        
        'EP2_128 MCh 30/11/2006
        If blnAFDBAffordability Then
            dbNetlAllowableIncome = (dblAllowableIncome * (intAFAllowableIncPercentage / 100)) - _
                                   (dblTotalRegOutgoings + dblTotalLoansLiabilities + dblTotalMortgageOutgoings + dblTotalDependantsOutgoings)
        Else
            dbNetlAllowableIncome = dblAllowableIncome - (dblTotalRegOutgoings + dblTotalLoansLiabilities + dblTotalMortgageOutgoings + dblTotalDependantsOutgoings) 'JD MAR30
        End If
        'End EP2_128
        
        ' JD MAR30 find net confirmed allowable income too
        'START: MAR272 Maha T (set NetConfirmedAllowbleIncome only if ConfirmedAllowableIncome > 0)
        
        'EP2_128 MCh
        If dblConfirmedAllowableIncome > 0 Then
            If blnAFDBAffordability Then
                dblNetConfirmedAllowableIncome = dblConfirmedAllowableIncome * (intAFAllowableIncPercentage / 100) - _
                                                 (dblTotalRegOutgoings + dblTotalLoansLiabilities + dblTotalMortgageOutgoings + dblTotalDependantsOutgoings)
            Else
                dblNetConfirmedAllowableIncome = dblConfirmedAllowableIncome - (dblTotalRegOutgoings + dblTotalLoansLiabilities + dblTotalMortgageOutgoings + dblTotalDependantsOutgoings)
            End If
        Else
            dblNetConfirmedAllowableIncome = 0
        End If
        ' END MAR272
                                
        Set xmlElement = xmlOut.createElement("MONTHLYALLOWABLEINCOME")
        xmlElement.Text = dblAllowableIncome / 12
        xmlRespCustomer.appendChild xmlElement
            
        Set xmlElement = xmlOut.createElement("ANNUALALLOWABLEINCOME")
        xmlElement.Text = dblAllowableIncome
        xmlRespCustomer.appendChild xmlElement
        '   AW  21/10/02    BMIDS00653
        Set xmlElement = xmlOut.createElement("NETANNUALALLOWABLEINCOME")
        xmlElement.Text = dbNetlAllowableIncome
        xmlRespCustomer.appendChild xmlElement
        'JD MAR30 return ConfirmedAllowableIncome
        Set xmlElement = xmlOut.createElement("CONFIRMEDALLOWABLEINCOME")
        If bUsedConfirmedDetails = True Then
            xmlElement.Text = dblConfirmedAllowableIncome
        Else
            xmlElement.Text = vbNullString
        End If
        xmlRespCustomer.appendChild xmlElement
        'JD MAR1040 return NetConfirmedAllowableIncome
        Set xmlElement = xmlOut.createElement("NETCONFIRMEDALLOWABLEINCOME")
        If bUsedConfirmedDetails = True Then
            xmlElement.Text = dblNetConfirmedAllowableIncome
        Else
            xmlElement.Text = vbNullString
        End If
        xmlRespCustomer.appendChild xmlElement
        
        xmlRespCustomers.appendChild xmlRespCustomer
        
        'Insert or Update IncomeSummary
        Set xmlTableNode = xmlDoc.createElement("INCOMESUMMARY")
        'MAR1891 GHun
        If blnIncomeSummaryExists Then
            xmlSetAttributeValue xmlTableNode, "EXISTS", 1
        End If
        'MAR1891 End
        xmlTempRequestNode.appendChild xmlTableNode
        Set xmlTempElem = xmlDoc.createElement("CUSTOMERNUMBER")
        xmlTempElem.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERNUMBER")
        xmlTableNode.appendChild xmlTempElem
        Set xmlTempElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlTempElem.Text = xmlGetAttributeText(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        xmlTableNode.appendChild xmlTempElem
        'm_objXMLAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomer, xmlTableNode, False
        'm_objXMLAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCustomer, xmlTableNode, False
        Set xmlFieldNode = xmlDoc.createElement("ALLOWABLEANNUALINCOME")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblAllowableIncome)
        Set xmlFieldNode = xmlDoc.createElement("NETALLOWABLEANNUALINCOME")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dbNetlAllowableIncome)
        Set xmlFieldNode = xmlDoc.createElement("ALLOWABLEMONTHLYINCOME")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblAllowableIncome / 12)
        Set xmlFieldNode = xmlDoc.createElement("GROSSCONFIRMEDALLOWABLEINCOME") 'JD MAR30
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblConfirmedAllowableIncome)
        Set xmlFieldNode = xmlDoc.createElement("NETCONFIRMEDALLOWABLEINCOME") 'JD MAR30
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblNetConfirmedAllowableIncome)
        ' save values from calculateOutgoings
        Set xmlFieldNode = xmlDoc.createElement("TOTALREGULAROUTGOINGSANNUALOUTGOINGS")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblTotalRegOutgoings)
        Set xmlFieldNode = xmlDoc.createElement("TOTALLOANSLIABILITYANNUALOUTGOINGS")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblTotalLoansLiabilities)
        Set xmlFieldNode = xmlDoc.createElement("TOTALMORTGAGEANNUALOUTGOINGS")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblTotalMortgageOutgoings)
        Set xmlFieldNode = xmlDoc.createElement("TOTALMORTGAGEOUTSTANDINGBALANCE")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblTotalBalanceMortgages)
        Set xmlFieldNode = xmlDoc.createElement("TOTALLOANSLIABILITYOUTSTANDINGBALANCE")
        xmlTableNode.appendChild xmlFieldNode
        xmlFieldNode.Text = CStr(dblTotalBalanceLoansLiabilities)
               
    Next xmlCustomer
    
    'MAR1891 GHun Only create objCustEmpBO just before it is needed
    If Not m_objContext Is Nothing Then
        Set objCustEmpBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    Else
        Set objCustEmpBO = New CustomerEmploymentBO 'BMIDS00653 MDC 28/10/2002
    End If
    'MAR1891 End
    
    'MAR1891 GHun Moved outside of loop so SaveTaxDetails to needs to be called once
    xmlSetAttributeValue xmlTempRequestNode, "EXISTENCECHECKED", 1
    strResponse = objCustEmpBO.SaveTaxDetails(xmlTempRequestNode.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    'MAR1891 End
                      
    CalculateAllowableIncome = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
         m_objContext.SetComplete
    End If
   
CalculateAllowableIncomeExit:

    Set xmlCustomerList = Nothing
    Set xmlCustomer = Nothing
    Set xmlUnearnedIncomeList = Nothing
    Set xmlUnearnedIncome = Nothing
    Set xmlEmploymentList = Nothing
    Set xmlEmployment = Nothing
    Set xmlRespCustomers = Nothing
    Set xmlRespCustomer = Nothing
    Set xmlElement = Nothing
    Set xmlNodeList = Nothing
    Set xmlFactorElem = Nothing
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objComboDO = Nothing
    Set objGlobalParameterDO = Nothing
    'BMIDS00653 MDC 28/10/2002
    Set objCustEmpBO = Nothing
    Set xmlTableNode = Nothing
    Set xmlFieldNode = Nothing
    Set xmlTempRequestNode = Nothing
    'BMIDS00653 MDC 28/10/2002 - End
    'MAR1891 GHun
    Set xmlIncomeTypeCombo = Nothing
    Set xmlConfirmedIncomeTypeCombo = Nothing
    Set xmlEmploymentStatusCombo = Nothing
    'MAR1891 End
    
    Exit Function
    
CalculateAllowableIncomeVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    'BMIDS00863 MDC 07/11/2002
    'CalculateAllowableIncome = m_objErrAssist.CreateErrorResponse
    CalculateAllowableIncome = m_objErrAssist.CreateErrorResponseEx().xml
    'BMIDS00863 MDC 07/11/2002 - End
   
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
   
    ' go to clean-up section
    Resume CalculateAllowableIncomeExit

End Function

'MAR1891 GHun Checks if a ValueId of a combo value has a specific validation type. The combo group is passed in as cached XML, rather than going to the database every time.
Private Function IsItemInValidationXML(ByVal vxmlCombo As IXMLDOMElement, ByVal vstrValueId As String, ByVal vstrValidationType As String) As Boolean
    If Not vxmlCombo.selectSingleNode("COMBOVALUE[@VALUEID='" & vstrValueId & "']/COMBOVALIDATION[@VALIDATIONTYPE='" & vstrValidationType & "']") Is Nothing Then
        IsItemInValidationXML = True
    Else
        IsItemInValidationXML = False
    End If
End Function
'MAR1891 End

Private Sub CalculateOutGoings(ByVal vxmlOutGoingsResponse As IXMLDOMNode, _
                               ByVal xmlCurrentCustomer As IXMLDOMNode, _
                               ByRef dblTotalRegOutgoings As Double, _
                               ByRef dblTotalLoansLiabilities As Double, _
                               ByRef dblTotalMortgageOutgoings As Double, _
                               ByRef dblTotalBalanceMortgages As Double, _
                               ByRef dblTotalBalanceLoansLiabilities As Double, _
                               ByVal blnAFDBAffordability As Boolean)
'-------------------------------------------------------------------------------------------------
'Description :
'Called from : CalculateAllowableIncome
'JD MAR30 added input param dblTotalMortgages
'JD MAR688 all outgoings now already extracted with stored procedure and passed in as
' vxmlOutGoingsResponse. Calls to methods removed.
'-------------------------------------------------------------------------------------------------
On Error GoTo CalculateVbErr
    
    Const strFunctionName As String = "CalculateOutGoings"
        
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objComboDO As ComboDO   'BMIDS00653 MDC 06/11/2002
    Dim xmlNode2 As IXMLDOMNode
    Dim xmlNodeThisCustomer As IXMLDOMNode
    
    'BMIDS00653 MDC 06/11/2002 - Validate outstanding term
    Dim strTempDate As String
    'MAR175
    Dim strNoOfDays As String
    Dim blnSkip As Boolean
    'BMIDS00653 MDC 06/11/2002 - End
    
    'Dim strResponse As String
    'BMIDS00653 MDC 28/10/2002
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    
    'Dim dblTotalBalanceLoansLiabilities As Double 'JD MAR30
    Dim strRedemptionStatus As String 'JD MAR30
    'Dim dblTotalBalanceMortgages As Double 'JD MAR30
    'Dim xmlTempRequestNode As IXMLDOMNode 'JD MAR30
    'Dim xmlFieldNode As IXMLDOMNode ' JD MAR30
    Dim strAccountguid As String 'JD MAR1040
    
    ' PSC 17/05/2006 MAR1801 - Start
    Dim xmlMortgageAccountList As IXMLDOMNodeList
    Dim xmlMortgageAccount As IXMLDOMNode
    Dim xmlMortgageLoanList As IXMLDOMNodeList
    Dim xmlMortgageLoan As IXMLDOMNode
    Dim strSearchPattern As String
    ' PSC 17/05/2006 MAR1801 - End

    'EP2_128 MCh
    Dim blnAFExcLoanLiabilityPaidbyBus As Boolean
    Dim blnPaidForByBusiness As Boolean
    
    ' PSC 26/01/2007 EP2_976 - Start
    Dim dblResidentialMortPercent As Double
    Dim dblBTLMortPercent As Double
    Dim dblOutstandingBalance As Double
    Dim dblTotalCollateralBalance As Double
    Dim blnOwnAccount As Boolean
    ' PSC 26/01/2007 EP2_976 - End
    
    strCustomerNumber = xmlGetAttributeText(xmlCurrentCustomer, "CUSTOMERNUMBER")
    strCustomerVersionNumber = xmlGetAttributeText(xmlCurrentCustomer, "CUSTOMERVERSIONNUMBER")
    'BMIDS00653 MDC 28/10/2002 - End
    
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objComboDO = New ComboDO    'BMIDS00653 MDC 06/11/2002
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")   'BMIDS00653 MDC 06/11/2002
    End If
    
    'EP2_128
    blnAFExcLoanLiabilityPaidbyBus = objGlobalParameterDO.GetCurrentParameterByType("AFExcLoanLiabilityPaidbyBus", "BOOLEAN")
    
    ' PSC 26/01/2007 EP2_976 - Start
    dblResidentialMortPercent = objGlobalParameterDO.GetCurrentParameterByType("AFResidentialMortPercent", "PERCENTAGE")
    dblBTLMortPercent = objGlobalParameterDO.GetCurrentParameterByType("AFBTLMortPercent ", "PERCENTAGE")
    ' PSC 26/01/2007 EP2_976 - End
    
    'JD MAR688 choose this customer
    Set xmlNodeThisCustomer = vxmlOutGoingsResponse.selectSingleNode(".//CUSTOMER[@CUSTOMERNUMBER='" + strCustomerNumber + "' and @CUSTOMERVERSIONNUMBER='" + strCustomerVersionNumber + "']")
    
    '-----------------------------------------------------------------------------------
    ' Find Total of Monthly RegularOutGoings this customer
    '----------------------------------------------------------------------------------
    Dim iPayFrequency As Integer, strAmount As String
    Dim strFrequency As String
    Dim bNumApplicants As Integer
    
    'MAR688
    For Each xmlNode2 In xmlNodeThisCustomer.selectNodes(".//REGULAROUTGOINGS")

        Dim strThisRegOutgoingsGuid As String
        strThisRegOutgoingsGuid = xmlGetAttributeText(xmlNode2, "REGULAROUTGOINGSGUID")
        
        bNumApplicants = IIf(vxmlOutGoingsResponse.selectNodes(".//CUSTOMER/REGULAROUTGOINGS[@REGULAROUTGOINGSGUID='" + strThisRegOutgoingsGuid + "']").length <> 1, 2, 1)
        
        strFrequency = xmlGetAttributeText(xmlNode2, "PAYMENTFREQUENCY")
        iPayFrequency = CSafeInt(strFrequency)
        strAmount = xmlGetAttributeText(xmlNode2, "AMOUNT")
        
        If (CSafeDbl(strAmount)) Then
            dblTotalRegOutgoings = dblTotalRegOutgoings _
                               + (iPayFrequency * CSafeDbl(strAmount) / bNumApplicants)
        End If
        
    Next xmlNode2
    

    '-----------------------------------------------------------------------------------
    ' Find Total of Monthly Loans And Liabilities this customer
    '----------------------------------------------------------------------------------
    Dim strLiabilityType As String
    Dim dTotAmount As Double, dCCPercentage As Double
    
    'JD MAR30 initialise
    dblTotalBalanceLoansLiabilities = 0
    
    strNoOfDays = vbNullString  'MAR1891 GHun
    dCCPercentage = -1        'MAR1891 GHun
    
    Dim xmlNodeAccountRelationship As IXMLDOMNode
    For Each xmlNode2 In xmlNodeThisCustomer.selectNodes(".//LOANSLIABILITIES[@LOANREPAYMENTINDICATOR='0']")
        
        strAccountguid = xmlGetAttributeText(xmlNode2, "ACCOUNTGUID")
        blnPaidForByBusiness = xmlGetAttributeAsBoolean(xmlNode2, "PAIDFORBYBUSINESS")
        
        blnSkip = False
        
        Set xmlNodeAccountRelationship = xmlNodeThisCustomer.selectSingleNode("ACCOUNTRELATIONSHIP[@ACCOUNTGUID='" + strAccountguid + "']")
        If xmlGetAttributeText(xmlNodeAccountRelationship, "UNASSIGNED") = "1" Then
            blnSkip = True
        End If
        
        'BMIDS00653 MDC 06/11/2002 - Validate outstanding term
        strTempDate = xmlGetAttributeText(xmlNode2, "ENDDATE")
        
        If IsDate(strTempDate) Then
            'MAR1891 GHun Only retrieve global parameter if necessary
            If Len(strNoOfDays) = 0 Then
                'MAR175 Global changed to Number of days basis
                strNoOfDays = objGlobalParameterDO.GetCurrentParameterByType("LoanTermMonthsRemaining", "AMOUNT")
            End If
            'MAR1891 End
            'AW EP2_2250
            If DateDiff("d", Now, strTempDate) <= CSafeInt(strNoOfDays) Then
                    blnSkip = True
            Else
                If blnAFExcLoanLiabilityPaidbyBus And blnPaidForByBusiness Then
                    blnSkip = True
                End If
            End If
            'AW EP2_2250  - End
        End If
        
        If Not blnSkip Then
        'BMIDS00653 MDC 06/11/2002 - End
        
            'BMIDS00653 MDC 28/10/2002
            'bNumApplicants = 1
            bNumApplicants = IIf(vxmlOutGoingsResponse.selectNodes(".//CUSTOMER/ACCOUNTRELATIONSHIP[@ACCOUNTGUID='" + strAccountguid + "']").length <> 1, 2, 1)
            'BMIDS00653 MDC 28/10/2002 - End
            
            strLiabilityType = xmlGetAttributeText(xmlNode2, "AGREEMENTTYPE")
            
            'BMIDS00653 MDC 06/11/2002
            'If (strLiabilityType = "6") Then 'Credit card
            If objComboDO.IsItemInValidation("AgreementType", strLiabilityType, "CC") Then 'Credit card
            'BMIDS00653 MDC 06/11/2002 - End
            
                dTotAmount = CSafeDbl(xmlGetAttributeText(xmlNode2, "TOTALOUTSTANDINGBALANCE"))
                'MAR1891 GHun Only retrieve global parameter if necessary
                If dCCPercentage = -1 Then
                    dCCPercentage = CSafeDbl(objGlobalParameterDO.GetCurrentParameterByType("CreditCardPercentage", "PERCENTAGE"))
                End If
                'MAR1891 End
            
                If ((dTotAmount > 0#) And (dCCPercentage > 0#)) Then
                
                    dblTotalLoansLiabilities = dblTotalLoansLiabilities + _
                    (((dTotAmount * dCCPercentage / 100#) * 12) / bNumApplicants)
                
                End If
                
            Else
               dblTotalLoansLiabilities = dblTotalLoansLiabilities + _
               ((CSafeDbl(xmlGetAttributeText(xmlNode2, "MONTHLYREPAYMENT")) * 12) / bNumApplicants)
                
            End If
            
            'JD MAR30
            'MAR118 Only include in balance those with validation type IIB
            If objComboDO.IsItemInValidation("AgreementType", strLiabilityType, "IIB") Then
                dblTotalBalanceLoansLiabilities = dblTotalBalanceLoansLiabilities + CSafeDbl(xmlGetAttributeText(xmlNode2, "TOTALOUTSTANDINGBALANCE"))
            End If
        End If  'BMIDS00653 MDC 06/11/2002
                
    Next xmlNode2
    
    'JD MAR30 get the mortgage loans
    '-----------------------------------------------------------------------------------
    ' Find Total outgoings and total balance of mortgage loans for this customer
    '----------------------------------------------------------------------------------
    dblTotalMortgageOutgoings = 0
    dblTotalBalanceMortgages = 0
            
    ' PSC 17/05/2006 MAR1801 - Start
    'strSearchPattern = ".//MORTGAGEACCOUNTANDLOAN[ACCOUNTRELATIONSHIP[CUSTOMERNUMBER='" & strCustomerNumber & "' and CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and not(CUSTOMERNUMBER=preceding::ACCOUNTRELATIONSHIP/CUSTOMERNUMBER and CUSTOMERVERSIONNUMBER=preceding::ACCOUNTRELATIONSHIP/CUSTOMERVERSIONNUMBER and ACCOUNTGUID=preceding::ACCOUNTRELATIONSHIP/ACCOUNTGUID)]]/MORTGAGEACCOUNT"
    
    'Set xmlMortgageAccountList = xmlDoc.selectNodes(strSearchPattern)
    Set xmlMortgageAccountList = xmlNodeThisCustomer.selectNodes(".//MORTGAGEACCOUNT")
    
    ' Process each unique mortgage account
    For Each xmlMortgageAccount In xmlMortgageAccountList
        strAccountguid = xmlGetAttributeText(xmlMortgageAccount, "ACCOUNTGUID")
        
        blnOwnAccount = xmlGetAttributeAsBoolean(xmlMortgageAccount, "OWNACCOUNT")
        
        bNumApplicants = vxmlOutGoingsResponse.selectNodes(".//CUSTOMER/MORTGAGEACCOUNT[@ACCOUNTGUID='" + strAccountguid + "']").length
        
        If bNumApplicants > 2 Then
            bNumApplicants = 2
        End If
        
        ' Get any loans for this account
        ' EP1053  Correct search pattern.
        strSearchPattern = ".//MORTGAGELOAN[@ACCOUNTGUID='" & strAccountguid & "' and @MORTGAGELOANGUID!='']"
        Set xmlMortgageLoanList = xmlNodeThisCustomer.selectNodes(strSearchPattern)
        
        ' Use data from loans if there are any else use the account figures
        If xmlMortgageLoanList.length > 0 Then
            For Each xmlMortgageLoan In xmlMortgageLoanList
                strRedemptionStatus = xmlGetAttributeText(xmlMortgageLoan, "REDEMPTIONSTATUS")
                
                ' Not other security
                If Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "O") Then
                    ' Total the outstanding balances
                    ' PSC 26/01/2007 EP2_976 - Start
                    dblOutstandingBalance = xmlGetAttributeAsDouble(xmlMortgageLoan, "OUTSTANDINGBALANCE")
                    dblTotalBalanceMortgages = dblTotalBalanceMortgages + dblOutstandingBalance
                    ' PSC 26/01/2007 EP2_976 - End
                    
                    ' Not to be redeemed so calculate outgoings
                    If Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "R") Then
                        
                        ' PSC 26/01/2007 EP2_976 - Start
                        If blnAFDBAffordability Then

                            If IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "BTLI") And blnOwnAccount Then
                                dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + ((dblOutstandingBalance * dblBTLMortPercent / 100) / bNumApplicants)
                            ElseIf Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "OBTL") And _
                                   Not (IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "BTLR") And blnOwnAccount) Then
                                dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + ((dblOutstandingBalance * dblResidentialMortPercent / 100) / bNumApplicants)
                            End If
                          
                        Else
                            dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + (xmlGetAttributeAsDouble(xmlMortgageLoan, "MONTHLYREPAYMENT") * 12 / bNumApplicants)
                        End If
                        ' PSC 26/01/2007 EP2_976 - End
                        
                    End If
                End If
            Next
        Else
            strRedemptionStatus = xmlGetAttributeText(xmlMortgageAccount, "MORTGAGEACCOUNTREDEMPTIONSTATUS")
            
            ' Not other security
            If Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "O") Then
                ' Total the outstanding balances
                ' PSC 26/01/2007 EP2_976 - Start
                dblTotalCollateralBalance = xmlGetAttributeAsDouble(xmlMortgageAccount, "TOTALCOLLATERALBALANCE")
                dblTotalBalanceMortgages = dblTotalBalanceMortgages + dblTotalCollateralBalance
                ' PSC 26/01/2007 EP2_976 - End
                    
                ' Not to be redeemed so calculate outgoings
                If Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "R") Then
                
                    ' PSC 26/01/2007 EP2_976 - Start
                    If blnAFDBAffordability Then

                        If IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "BTLI") And blnOwnAccount Then
                            dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + ((dblTotalCollateralBalance * dblBTLMortPercent / 100) / bNumApplicants)
                        ElseIf Not IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "OBTL") And _
                               Not (IsValidationType("RedemptionStatus", CSafeInt(strRedemptionStatus), "BTLR") And blnOwnAccount) Then
                            dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + ((dblTotalCollateralBalance * dblResidentialMortPercent / 100) / bNumApplicants)
                        End If
                
                    Else
                        dblTotalMortgageOutgoings = dblTotalMortgageOutgoings + (xmlGetAttributeAsDouble(xmlMortgageAccount, "TOTALMONTHLYCOST") * 12 / bNumApplicants)
                    End If
                    ' PSC 26/01/2007 EP2_976 - End
                End If
            End If
        End If
    Next
    ' PSC 17/05/2006 MAR1801 - End
    
    'MAR688 Income Summary is saved in CalcAllowableIncome after call to this method
    'therefore don't do a save here.
    'JD MAR30 save subtotals to the incomesummary table
'    Set xmlTempRequestNode = vxmlOutGoingsRequest.cloneNode(False)
'    Set xmlTableNode = xmlDoc.createElement("INCOMESUMMARY")
'    xmlTempRequestNode.appendChild xmlTableNode
'    Set xmlTempElem = xmlDoc.createElement("CUSTOMERNUMBER")
'    xmlTempElem.Text = xmlGetAttributeText(xmlCurrentCustomer, "CUSTOMERNUMBER")
'    xmlTableNode.appendChild xmlTempElem
'    Set xmlTempElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
'    xmlTempElem.Text = xmlGetAttributeText(xmlCurrentCustomer, "CUSTOMERVERSIONNUMBER")
'    xmlTableNode.appendChild xmlTempElem
'    'm_objXMLAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCurrentCustomer, xmlTableNode, False
'    'm_objXMLAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlCurrentCustomer, xmlTableNode, False
'    Set xmlFieldNode = xmlDoc.createElement("TOTALREGULAROUTGOINGSANNUALOUTGOINGS")
'    xmlTableNode.appendChild xmlFieldNode
'    xmlFieldNode.Text = CStr(dblTotalRegOutgoings)
'    Set xmlFieldNode = xmlDoc.createElement("TOTALLOANSLIABILITYANNUALOUTGOINGS")
'    xmlTableNode.appendChild xmlFieldNode
'    xmlFieldNode.Text = CStr(dblTotalLoansLiabilities)
'    Set xmlFieldNode = xmlDoc.createElement("TOTALMORTGAGEANNUALOUTGOINGS")
'    xmlTableNode.appendChild xmlFieldNode
'    xmlFieldNode.Text = CStr(dblTotalMortgageOutgoings)
'    Set xmlFieldNode = xmlDoc.createElement("TOTALMORTGAGEOUTSTANDINGBALANCE")
'    xmlTableNode.appendChild xmlFieldNode
'    xmlFieldNode.Text = CStr(dblTotalBalanceMortgages)
'    Set xmlFieldNode = xmlDoc.createElement("TOTALLOANSLIABILITYOUTSTANDINGBALANCE")
'    xmlTableNode.appendChild xmlFieldNode
'    xmlFieldNode.Text = CStr(dblTotalBalanceLoansLiabilities)
'    strResponse = objCustEmpBO.SaveTaxDetails(xmlTempRequestNode.xml)
'    m_objErrAssist.CheckXMLResponse strResponse, True
   
    Set objGlobalParameterDO = Nothing
    
    Set xmlNode2 = Nothing
    Set objComboDO = Nothing    'BMIDS00653 MDC 06/11/2002
    
    ' PSC 17/05/2006 MAR1801 - Start
    Set xmlMortgageAccountList = Nothing
    Set xmlMortgageAccount = Nothing
    Set xmlMortgageLoanList = Nothing
    Set xmlMortgageLoan = Nothing
    ' PSC 17/05/2006 MAR1801 - End
   
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

CalculateVbErr:
    
    Set objGlobalParameterDO = Nothing
    
    Set xmlNode2 = Nothing
    Set objComboDO = Nothing    'BMIDS00653 MDC 06/11/2002
    
    ' PSC 17/05/2006 MAR1801 - Start
    Set xmlMortgageAccountList = Nothing
    Set xmlMortgageAccount = Nothing
    Set xmlMortgageLoanList = Nothing
    Set xmlMortgageLoan = Nothing
    ' PSC 17/05/2006 MAR1801 - End

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    Set m_objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


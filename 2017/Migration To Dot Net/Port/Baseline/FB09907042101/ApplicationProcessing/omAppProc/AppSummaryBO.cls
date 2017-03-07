VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AppSummaryBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AppSummaryBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Application Summary Business Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    26/02/01    SYS1879 Created
'PSC    13/03/01    SYS1879 Amend GetLoanAndRisksData to get the Valuation Report details
'PSC    10/04/01    SYS2115 Amend GetLoanAndRisksData to change message if accepted quote
'                   number not present
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IAppSummaryBO

Private m_objContext As ObjectContext

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR     Description
'MV     04/02/2003  BM0288  Amended IAppSummaryBO_GetLoanAndRisksData()
'MV     18/02/2003  BM0288  Amended IAppSummaryBO_GetLoanAndRisksData()
'------------------------------------------------------------------------------------------
'MARS Specific History:
'JD     17/02/2006  MAR1040 GetLoanAndRiskData ammended to include income multiples and riskassessmentdecision
'JD     09/03/2006  MAR1061 GetLoanAndRisksData - added retrieval of MortgageSubquote.PurchasePriceOrEstimatedValue
'PSC    25/03/2006  MARS1485 Apply BBG1939 Performance fixes
'JD     16/05/2006  MAR1700 GetLoanAndRisksData - get ChangeOfProperty and ValuationType
'------------------------------------------------------------------------------------------
Private Sub IAppSummaryBO_CreateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified APPLICATIONREVIEWHISTORY entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo CreateApplicationReviewHistoryExit
    
    Const cstrFunctionName As String = "IAppSummaryBO_CreateApplicationReviewHistory"
    
    Dim objAppSummaryTxBO As IAppSummaryTxBO    ' PSC 25/03/2006 MARS1485

    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "APPLICATIONREVIEWHISTORY"
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objAppSummaryTxBO = m_objContext.CreateInstance(App.Title & ".AppSummaryTxBO")
    Else
        Set objAppSummaryTxBO = New AppSummaryTxBO
    End If

    objAppSummaryTxBO.CreateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
    ' PSC 25/03/2006 MARS1485 - End
    
CreateApplicationReviewHistoryExit:
    
    Set objAppSummaryTxBO = Nothing     ' PSC 25/03/2006 MARS1485
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAppSummaryBO_FindApplicationReviewHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve a list of APPLICATIONREVIEWHISTORY entities
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       APPLICATIONREVIEWHISTORYLIST node appended to vxmlResponseNode on exit
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo FindApplicationReviewHistoryListExit

    Const cstrFunctionName As String = "IAppSummaryBO_FindApplicationReviewHistoryList"

    Dim xmlAppReviewNode As IXMLDOMNode
    Dim objAppSummaryDO As IAppSummaryDO    ' PSC 25/03/2006 MARS1485

    ' Get Application Review History node
    Set xmlAppReviewNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONREVIEWHISTORY")

    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objAppSummaryDO = m_objContext.CreateInstance(App.Title & ".AppSummaryDO")
    Else
        Set objAppSummaryDO = New AppSummaryDO
    End If

    objAppSummaryDO.FindApplicationReviewHistoryList xmlAppReviewNode, vxmlResponseNode
    ' PSC 25/03/2006 MARS1485 - End
    
FindApplicationReviewHistoryListExit:

    Set objAppSummaryDO = Nothing       ' PSC 25/03/2006 MARS1485
    Set xmlAppReviewNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAppSummaryBO_GetLoanAndRisksData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetLoanAndRisksDataExit
    
    Const cstrFunctionName As String = "IAppSummaryBO_GetLoanAndRisksData"
    
    Dim xmlAppFactFind As IXMLDOMElement
    Dim xmlQuotationLoanSummary As IXMLDOMElement
    Dim xmlValuation As IXMLDOMElement
    Dim xmlLoanData As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlApplication As IXMLDOMElement
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponse As IXMLDOMElement
    
    ' PSC 25/03/2006 MARS1485 - Start
    Dim objApplicationBO As ApplicationBO
    Dim objQuotationBO As QuotationBO
    Dim objValuationDO  As IValuationDO
    Dim objRABO As RiskAssessmentBO
    ' PSC 25/03/2006 MARS1485 - End
    Dim objNewPropertyBO As NewPropertyBO 'MAR1700

    Dim strResponse As String
    Dim strValue As String
    Dim lngErrorNo As Long
    Dim strSource As String
    Dim strDescription As String
    Dim strQuotationNo As String        ' PSC 25/03/2006 MARS1485

    If vxmlRequestNode.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequestNode
    Else
        Set xmlRequestNode = xmlGetMandatoryNode(xmlRequestNode, "//REQUEST")
    End If
    
    ' Validate application data present
    Set xmlApplication = xmlGetMandatoryNode(xmlRequestNode, "APPLICATION")
    
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONFACTFINDNUMBER"
    
    ' Put loan data onto the response
    Set xmlLoanData = vxmlResponseNode.ownerDocument.createElement("LOANDATA")
    vxmlResponseNode.appendChild xmlLoanData
    
    ' Create request to get the application fact find data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "APPLICATIONFACTFIND")
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Else
        Set objApplicationBO = New ApplicationBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Get the application fact find data to retrieve accepted quotation number
    ' PSC 25/03/2006 MARS1485
    strResponse = objApplicationBO.GetApplicationFactFindData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
            
    ' PSC 10/04/01 SYS2115
    strQuotationNo = xmlGetMandatoryNodeText(xmlResponseDoc, ".//ACCEPTEDQUOTENUMBER", 308)
    
    ' Convert response to attribute based so that required attributes can be copied to
    ' the response
    Set xmlAppFactFind = xmlGetMandatoryNode(xmlResponseDoc, ".//APPLICATIONFACTFIND")
    Set xmlAppFactFind = xmlCreateAttributeBasedResponse(xmlAppFactFind, True)

    xmlCopyAttribute xmlAppFactFind, xmlLoanData, "APPLICATIONNUMBER"
    xmlCopyAttribute xmlAppFactFind, xmlLoanData, "APPLICATIONFACTFINDNUMBER"
    xmlCopyAttribute xmlAppFactFind, xmlLoanData, "DIRECTINDIRECTBUSINESS"
    xmlCopyAttribute xmlAppFactFind, xmlLoanData, "DIRECTINDIRECTBUSINESS_TEXT"
    xmlCopyAttribute xmlAppFactFind, xmlLoanData, "PURCHASEPRICEORESTIMATEDVALUE"
    
    'Create request to get the accepted quotation loan summary
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "QUOTATION")

    Set xmlNode = xmlGetMandatoryNode(xmlTempRequestDoc, ".//QUOTATION")
    Set xmlElement = xmlTempRequestDoc.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strQuotationNo
    xmlNode.appendChild xmlElement
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    Else
        Set objQuotationBO = New QuotationBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Get the quotation data to retrieve the mortgage sub quote number
    ' PSC 25/03/2006 MARS1485
    strResponse = objQuotationBO.GetQuotationLoanSummary(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
       
    ' Convert response to attribute based so that required attributes can be copied to
    ' the response
    Set xmlQuotationLoanSummary = xmlGetMandatoryNode(xmlResponseDoc, ".//QUOTATIONLOANSUMMARY")
    Set xmlQuotationLoanSummary = xmlCreateAttributeBasedResponse(xmlQuotationLoanSummary, True)
    
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "TOTALLOANAMOUNT"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "LTV"
    'MAR1061 Get Purchaseprice
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "PURCHASEPRICEORESTIMATEDVALUE"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "PURPOSEOFLOAN"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "PURPOSEOFLOAN_TEXT"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "TERMINMONTHS"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "TERMINYEARS"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "INDEMNITYPREMIUM"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "INDEMNITYAMOUNT"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "BENEFITTYPE_TEXT"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "LOANSANDLIABILITIES"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "MONTHLYMORTGAGEPAYMENTS"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "MORTGAGERELATEDINSURANCE"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "OTHEROUTGOINGS"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "TOTALMONTHLYINCOME"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "SUBPURPOSEOFLOAN"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "SUBPURPOSEOFLOAN_TEXT"
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "INCOMEMULTIPLE" 'JD MAR1040
    xmlCopyAttribute xmlQuotationLoanSummary, xmlLoanData, "CONFIRMEDCALCULATEDINCMULTIPLE" 'JD MAR1040
    
    ' Create request to get the Valuation Report data
    Set xmlValuation = xmlTempRequestDoc.createElement("VALUATION")
    xmlCopyAttribute xmlApplication, xmlValuation, "APPLICATIONNUMBER"
    xmlCopyAttribute xmlApplication, xmlValuation, "APPLICATIONFACTFINDNUMBER"

    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objValuationDO = m_objContext.CreateInstance(App.Title & ".ValuationDO")
    Else
        Set objValuationDO = New ValuationDO
    End If
    ' PSC 25/03/2006 MARS1485 - End

    Set xmlTempResponse = xmlResponseDoc.createElement("RESPONSE")
    
    On Error Resume Next
    ' PSC 25/03/2006 MARS1485
    objValuationDO.GetValuationReport xmlValuation, xmlTempResponse

    lngErrorNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
    
    On Error GoTo GetLoanAndRisksDataExit
    
    ' Re raise error if not ok and not record not found
    If errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And lngErrorNo <> 0 Then
        Err.Raise lngErrorNo, strSource, strDescription
    End If

    ' Append the valuation report data to the output
    If lngErrorNo = 0 Then
        Set xmlValuation = xmlGetMandatoryNode(xmlTempResponse, "GETVALUATIONREPORT")
        xmlCopyAttribute xmlValuation, xmlLoanData, "PRESENTVALUATION"
        xmlCopyAttribute xmlValuation, xmlLoanData, "POSTWORKSVALUATION"
    End If

    ' Create request to get the Risk Assessment Score  data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "RISKASSESSMENT")
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objRABO = m_objContext.CreateInstance(gstrRISK_ASSESSMENT & ".RiskAssessmentBO")
    Else
        Set objRABO = New RiskAssessmentBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
        
    ' Get Risk Assessment Score
    'MAR1040 get the decision as well as the score.
    'strResponse = m_objRABO.GetLatestRiskAssessmentScore(xmlTempRequestDoc.xml)
    ' PSC 25/03/2006 MARS1485
    strResponse = objRABO.GetLatestRiskAssessment(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlResponseDoc.documentElement, vxmlResponseNode, False)
   
    ' Raise error if any error other than record not found
    If errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And lngErrorNo <> 0 Then
        errRaiseXMLResponseNode xmlResponseDoc.documentElement
    End If
    
    ' Append the decision to the output
    ' Append the score to the output
    If lngErrorNo = 0 Then
        Set xmlElement = xmlGetMandatoryNode(xmlResponseDoc, ".//RISKASSESSMENTDECISION")
        xmlLoanData.setAttribute "RISKASSESSMENTDECISION", xmlElement.getAttribute("TEXT")
        Set xmlElement = xmlGetMandatoryNode(xmlResponseDoc, ".//RISKASSESSMENTSCORE")
        xmlLoanData.setAttribute "RISKASSESSMENTSCORE", xmlElement.Text
    End If
    
    'MAR1700 Get the NewProperty ChangeOfProperty flag
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "NEWPROPERTY")
    If Not m_objContext Is Nothing Then
        Set objNewPropertyBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
    Else
        Set objNewPropertyBO = New NewPropertyBO
    End If
    strResponse = objNewPropertyBO.GetNewPropertyGeneral(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlResponseDoc.documentElement, vxmlResponseNode, False)
    ' Raise error if any error other than record not found
    If errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And lngErrorNo <> 0 Then
        errRaiseXMLResponseNode xmlResponseDoc.documentElement
    End If
    'append to output
    If lngErrorNo = 0 Then
        Set xmlElement = xmlGetNode(xmlResponseDoc, ".//CHANGEOFPROPERTY")
        If Not xmlElement Is Nothing Then
            xmlLoanData.setAttribute "CHANGEOFPROPERTY", xmlElement.Text
        End If
        Set xmlElement = xmlGetNode(xmlResponseDoc, ".//VALUATIONTYPE")
        If Not xmlElement Is Nothing Then
            xmlLoanData.setAttribute "VALUATIONTYPE", xmlElement.Text
        End If
    End If

    
    
   
GetLoanAndRisksDataExit:

    Set xmlAppFactFind = Nothing
    Set xmlQuotationLoanSummary = Nothing
    Set xmlValuation = Nothing
    Set xmlLoanData = Nothing
    Set xmlRequestNode = Nothing
    Set xmlApplication = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempResponse = Nothing
    
    ' PSC 25/03/2006 MARS1485 - Start
    Set objApplicationBO = Nothing
    Set objQuotationBO = Nothing
    Set objValuationDO = Nothing
    Set objRABO = Nothing
    ' PSC 25/03/2006 MARS1485 - End
    Set objNewPropertyBO = Nothing 'MAR1700

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAppSummaryBO_UpdateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Updates specified APPLICATIONREVIEWHISTORY entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo CreateApplicationReviewHistoryExit
    
    Const cstrFunctionName As String = "IAppSummaryBO_UpdateApplicationReviewHistory"
    
    Dim objAppSummaryTxBO As IAppSummaryTxBO    ' PSC 25/03/2006 MARS1485
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "APPLICATIONREVIEWHISTORY"
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objAppSummaryTxBO = m_objContext.CreateInstance(App.Title & ".AppSummaryTxBO")
    Else
        Set objAppSummaryTxBO = New AppSummaryTxBO
    End If
    
    objAppSummaryTxBO.UpdateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
    ' PSC 25/03/2006 MARS1485 - End
    
CreateApplicationReviewHistoryExit:
    
    Set objAppSummaryTxBO = Nothing     ' PSC 25/03/2006 MARS1485
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAppSummaryBO_ValidateUserLogon(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Validates the user password and that the user has the correct mandate level for
'   approving/reccomending the application
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo ValidateUserLogonExit
    
    Const cstrFunctionName As String = "IAppSummaryBO_ValidateUserLogon"
    
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlApplication As IXMLDOMElement
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlMortSubQuoteNo As IXMLDOMElement
    Dim xmlAmtRequested As IXMLDOMElement
    Dim xmlLTV As IXMLDOMElement
    Dim xmlUser As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    
    ' PSC 25/03/2006 MARS1485 - Start
    Dim objOrganisationBO  As OrganisationBO
    Dim objApplicationBO  As ApplicationBO
    Dim objQuotationBO  As QuotationBO
    Dim objMortSubQuoteBO  As MortgageSubQuoteBO
    ' PSC 25/03/2006 MARS1485 - End

    Dim strResponse As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequestNode
    Else
        Set xmlRequestNode = xmlGetMandatoryNode(xmlRequestNode, "//REQUEST")
    End If
    
    ' Validate request
    xmlCheckMandatoryAttribute xmlRequestNode, "USERID"
    
    ' Validate application data present
    Set xmlApplication = xmlGetMandatoryNode(xmlRequestNode, "APPLICATION")
    
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONFACTFINDNUMBER"
    
    ' Validate user data present
    Set xmlUser = xmlGetMandatoryNode(xmlRequestNode, "USER")
    
    xmlCheckMandatoryAttribute xmlUser, "USERID"
    xmlCheckMandatoryAttribute xmlUser, "PASSWORDVALUE"
    xmlCheckMandatoryAttribute xmlUser, "UNITID"
    
    ' Create a request to validate the user logon
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "USER", True, "OMIGAUSER")
    Set xmlTempRequest = xmlGetMandatoryNode(xmlTempRequestDoc, ".//REQUEST")
    xmlCopyAttributeValue xmlUser, xmlTempRequest, "USERID", "USERID"
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objOrganisationBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    Else
        Set objOrganisationBO = New OrganisationBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Validate the user logon
    ' PSC 25/03/2006 MARS1485
    strResponse = objOrganisationBO.ValidateUserLogon(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True

    ' Create request to get the application fact find data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "APPLICATIONFACTFIND")
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Else
        Set objApplicationBO = New ApplicationBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Get the application fact find data to retrieve accepted quotation number
    ' PSC 25/03/2006 MARS1485
    strResponse = objApplicationBO.GetApplicationFactFindData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
    
    Dim strQuotationNo As String
    
    strQuotationNo = xmlGetMandatoryNodeText(xmlResponseDoc, ".//ACCEPTEDQUOTENUMBER")
    
    'Create request to get the accepted quotation data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "QUOTATION")

    Set xmlNode = xmlGetMandatoryNode(xmlTempRequestDoc, ".//QUOTATION")
    Set xmlElement = xmlTempRequestDoc.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strQuotationNo
    xmlNode.appendChild xmlElement
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    Else
        Set objQuotationBO = New QuotationBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Get the quotation data to retrieve the mortgage sub quote number
    ' PSC 25/03/2006 MARS1485
    strResponse = objQuotationBO.GetData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
   
    Set xmlMortSubQuoteNo = xmlGetMandatoryNode(xmlResponseDoc, ".//MORTGAGESUBQUOTENUMBER")

    ' Create request to get the mortgage sub quote
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "MORTGAGESUBQUOTE")

    Set xmlNode = xmlGetMandatoryNode(xmlTempRequestDoc, ".//MORTGAGESUBQUOTE")
    xmlNode.appendChild xmlMortSubQuoteNo
    
    ' PSC 25/03/2006 MARS1485 - Start
    If Not m_objContext Is Nothing Then
        Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    Else
        Set objMortSubQuoteBO = New MortgageSubQuoteBO
    End If
    ' PSC 25/03/2006 MARS1485 - End
    
    ' Get the mortgage sub quote data to retrieve the amount requested and LTV
    ' PSC 25/03/2006 MARS1485
    strResponse = objMortSubQuoteBO.GetData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
   
    Set xmlAmtRequested = xmlGetMandatoryNode(xmlResponseDoc, ".//AMOUNTREQUESTED")
    Set xmlLTV = xmlGetMandatoryNode(xmlResponseDoc, ".//LTV")
    
    ' Create request to validate mandate level
    Set xmlTempRequest = xmlGetMandatoryNode(xmlTempRequestDoc, ".//REQUEST")
    xmlTempRequest.removeChild xmlTempRequest.firstChild
    xmlTempRequest.appendChild xmlAmtRequested
    xmlTempRequest.appendChild xmlLTV
    xmlCopyAttributeValue xmlUser, xmlTempRequest, "USERID", "USERID"
     
    ' Validate User Mandate Level
    strResponse = objOrganisationBO.ValidateUserMandateLevel(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
   
    Dim strValidUser As String
    strValidUser = xmlGetMandatoryNodeText(xmlResponseDoc, ".//VALIDUSER")
    
    Set xmlElement = xmlResponseDoc.createElement("USER")
    xmlElement.setAttribute "VALIDUSER", strValidUser
    vxmlResponseNode.appendChild xmlElement
    
ValidateUserLogonExit:

    Set xmlRequestNode = Nothing
    Set xmlApplication = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlMortSubQuoteNo = Nothing
    Set xmlAmtRequested = Nothing
    Set xmlLTV = Nothing
    Set xmlUser = Nothing
    Set xmlElement = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlNode = Nothing
    
    ' PSC 25/03/2006 MARS1485 - Start
    Set objOrganisationBO = Nothing
    Set objApplicationBO = Nothing
    Set objQuotationBO = Nothing
    Set objMortSubQuoteBO = Nothing
    ' PSC 25/03/2006 MARS1485 - End
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub


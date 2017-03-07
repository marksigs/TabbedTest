VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Preprocessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <PreProcessingBO>
' File:         <Preprocessing.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionBO.doc
' Purpose:      To perform any client based pre processing tasks for the ingestion engine
' History:
'   AQR         AUTHOR              DESCRIPTION                             DATE
'   ---         ------              -----------                             ------
'   IDWP1       Martin O'Hanlon     Created                                 17/07/2002
'   BMIDS00462  Martin O'Hanlon     Made changes to ingest purpose of loan  28/10/2002
'   BMIDS00999  Martin O'Hanlon     Changed to take stings from the core    20/11/2002
'                                   component
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'------------------------------------------------------------------------
' Procedure OmRequest
' Author: MO
' Date: 20/11/2002
' Purpose: Pre-processing public method, this will be called at the
'    pre-processing phase by the DataIngestion component
' Input parameters: XML Request string
' Output parameters: XML Response string
'------------------------------------------------------------------------
Public Function OmRequest(ByVal vstrXmlIn As String) As String

On Error GoTo OmRequestErrorOccured
        
    Dim xmlRequest As IXMLDOMNode
    Dim xmlOmigaFormattedRequest As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    
    'extract the main nodes from the rules request
    ExtractPreprocessingRulesRequest vstrXmlIn, xmlRequest, xmlOmigaFormattedRequest, xmlResponse
    
    'do the request
    DoRequest xmlRequest, xmlOmigaFormattedRequest, xmlResponse
    
    'return the response xml string
    OmRequest = BuildPreprocessingRulesResponse(xmlRequest, xmlOmigaFormattedRequest, xmlResponse)
    
OmRequestExit:

    'destroy objects
    Set xmlRequest = Nothing
    Set xmlOmigaFormattedRequest = Nothing
    Set xmlResponse = Nothing
    
Exit Function

OmRequestErrorOccured:

    Err.Raise Err.Number, "DoRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume OmRequestExit

End Function

'------------------------------------------------------------------------
' Procedure DoRequest
' Author: MO
' Date: 19/07/2002
' Purpose: This is the method which actual performs the correct tasks,
'           based on the operation
' Input parameters: vxmlRequestNode, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DoRequestErrorOccured
        
    'run specific pre-processing function dependant on operation
    Select Case xmlGetMandatoryNode(vxmlRequestNode, "@OPERATION").Text
        
        Case "CREATEAPPLICATION"
        
            'run validation
            Validate xmlGetMandatoryNode(vxmlRequestNode, "@OPERATION").Text, vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
        
            'generate the keys needed for the ingestion
            GenerateKeys vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
            
            'set the number of dependencies
            SetNumberOfDependants vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
            
            'set up loan components
            SetupLoanComponents vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
            
            'find the additional mortgage product details
            FindMortgageProductDetails vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
        
    End Select
    
DoRequestExit:

    'destroy objects
    
Exit Sub

DoRequestErrorOccured:

    Err.Raise Err.Number, "DoRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume DoRequestExit

End Sub

'------------------------------------------------------------------------
' Procedure GenerateKeys
' Author: MO
' Date: 17/07/2002
' Purpose: Generates the unique keys that cannot be gained from the CRUD
' Input parameters: OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Function GenerateKeys(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo GenerateKeysErrorOccured
    
    'variable declaration
    Dim objCustomerNode As IXMLDOMNode
    Dim objCustomerNumber As IXMLDOMNode
    Dim objOmCustBO As Object
    Dim objOmApplicationBO As Object
    Dim strNewCustomerNumberXML As String
    Dim objApplicationNumber As IXMLDOMNode
    Dim strNewApplicationNumber As String
    Dim objChannelID As IXMLDOMNode
    Dim objResponseDOM As FreeThreadedDOMDocument40
    
    ' Create new Customer numbers
    Set objOmCustBO = gobjContext.CreateInstance("omCust.CustomerBO")
    
    For Each objCustomerNode In vxmlOmigaFormatRequestNode.selectNodes("//CUSTOMER")
        Set objCustomerNumber = xmlGetMandatoryAttribute(objCustomerNode, "CUSTOMERNUMBER")
        'if a customer number hasnt been assigned, get one for it
        If objCustomerNumber.Text = "" Then
        
            'Get the next customer number
            strNewCustomerNumberXML = objOmCustBO.GetNextCustomerNumber
            
            errCheckXMLResponse strNewCustomerNumberXML
            
            Set objResponseDOM = New FreeThreadedDOMDocument40
            objResponseDOM.validateOnParse = False
            objResponseDOM.setProperty "NewParser", True
            
            objResponseDOM.loadXML strNewCustomerNumberXML
            
            'set the attribute value for this node
            objCustomerNumber.Text = xmlGetMandatoryNode(objResponseDOM.documentElement, "CUSTOMER/@CUSTOMERNUMBER").Text
        End If
       
    Next
    
    Set objOmCustBO = Nothing
    
    'get the application number
    Set objApplicationNumber = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "//APPLICATION/@APPLICATIONNUMBER")
    ' Create new application number
    Set objOmApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    
    ' get the channel id
    Set objChannelID = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "//@CHANNELID")
    
    strNewApplicationNumber = objOmApplicationBO.GetNextApplicationNumber(objChannelID.Text)
    'set the attribute value for this node
    objApplicationNumber.Text = strNewApplicationNumber

GenerateKeysExit:
    
    'destroy objects
    Set objApplicationNumber = Nothing
    Set objChannelID = Nothing
    Set objOmApplicationBO = Nothing
    Set objCustomerNode = Nothing
    Set objCustomerNumber = Nothing
    Set objOmCustBO = Nothing
    Set objResponseDOM = Nothing

Exit Function

GenerateKeysErrorOccured:
    
    Err.Raise Err.Number, "GenerateKeys." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume GenerateKeysExit
    
End Function

'------------------------------------------------------------------------
' Procedure Validate
' Author: MO
' Date: 30/07/2002
' Purpose: Generates an error if the request doesnt pass validation
' Input parameters: strOperation, RequestXML, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub Validate(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidateErrorOccured
    
    Select Case strOperation
        
        Case "CreateApplication"
        
            'There is currently no validation - but it could be put in here like this
            'if [something] <> [avalue] then
                'err.raise direValidationFailed, "Validation", "Validation Failed - The something value was not equal to the value."
            'end if
    
    End Select
    
ValidateExit:

    'destory objects

Exit Sub
    
ValidateErrorOccured:

    Err.Raise Err.Number, "Validate." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume ValidateExit

End Sub

'------------------------------------------------------------------------
' Procedure FindMortgageProductDetails
' Author: MO
' Date: 13/08/2002
' Purpose: Performs a searches of mortgage products to find the product
'           information for the SalesProductCode
' Input parameters: RequestXML, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub FindMortgageProductDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindMortgageProductDetailsErrorOccured
    
    Dim objOmMPBO As Object
    Dim xmlMPSearchRequestDOM As FreeThreadedDOMDocument40
    Dim xmlMPSearchResponseDOM As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlLoanComponentNode As IXMLDOMNode
    Dim strMPSearchResponse As String
    Dim strMortgageProductCode As String
    
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
    
    Set objOmMPBO = gobjContext.CreateInstance("omMP.MortgageProductBO")
    
    For Each xmlLoanComponentNode In vxmlOmigaFormatRequestNode.selectNodes("//MORTGAGESUBQUOTE/LOANCOMPONENT")
        
        'build mortgage product search request
        Set xmlMPSearchRequestDOM = New FreeThreadedDOMDocument40
        xmlMPSearchRequestDOM.validateOnParse = False
        xmlMPSearchRequestDOM.setProperty "NewParser", True
        
        Set xmlNode = xmlMPSearchRequestDOM.createElement("REQUEST")
        xmlMPSearchRequestDOM.appendChild xmlNode
        
        Set xmlNode = xmlMPSearchRequestDOM.createElement("SEARCH")
        xmlMPSearchRequestDOM.selectSingleNode("REQUEST").appendChild xmlNode
        
        Set xmlNode = xmlMPSearchRequestDOM.createElement("MORTGAGEPRODUCT")
        xmlMPSearchRequestDOM.selectSingleNode("REQUEST/SEARCH").appendChild xmlNode
        
        Set xmlNode = xmlMPSearchRequestDOM.createElement("MORTGAGEPRODUCTCODE")
        strMortgageProductCode = xmlGetMandatoryAttribute(xmlLoanComponentNode, "MORTGAGEPRODUCTCODE").Text
        xmlNode.Text = strMortgageProductCode
        xmlMPSearchRequestDOM.selectSingleNode("REQUEST/SEARCH/MORTGAGEPRODUCT").appendChild xmlNode
        
        Set xmlNode = xmlMPSearchRequestDOM.createElement("STARTDATE")
        xmlNode.Text = Now
        xmlMPSearchRequestDOM.selectSingleNode("REQUEST/SEARCH/MORTGAGEPRODUCT").appendChild xmlNode
        
        'try and find the product in omiga, for the product specified within the input XML
        strMPSearchResponse = objOmMPBO.FindProductList(xmlMPSearchRequestDOM.xml)
        
        'check to see if a no record found error was generated
        On Error Resume Next
        
        errCheckXMLResponse strMPSearchResponse
                
        'get the error details
        lngErrorNumber = Err.Number
        strErrorSource = Err.Source
        strErrorDescription = Err.Description
                
        On Error GoTo FindMortgageProductDetailsErrorOccured
        
        'has an error has occured
        If lngErrorNumber <> 0 Then
            
            'is this error that the record cuold not be found?
            If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
                Err.Raise direMortgageProductDoesntExist, "FindMortgageProductDetails." & strErrorSource, "The mortgage product " & strMortgageProductCode & " could not be found in the Omiga 4 database."
            Else
                Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
            End If
            
        End If
        
        'clear the error
        Err.Clear
        
        'load response in DOM
        Set xmlMPSearchResponseDOM = New FreeThreadedDOMDocument40
        xmlMPSearchResponseDOM.validateOnParse = False
        xmlMPSearchResponseDOM.setProperty "NewParser", True
        
        xmlMPSearchResponseDOM.loadXML strMPSearchResponse
        
        'get the start date from the mortgage product details and add it to the loan component
        xmlSetAttributeValue xmlLoanComponentNode, "STARTDATE", xmlGetMandatoryNode(xmlMPSearchResponseDOM.documentElement, "MORTGAGEPRODUCTLIST/MORTGAGEPRODUCT/STARTDATE").Text
        
        Set xmlMPSearchResponseDOM = Nothing
        Set xmlMPSearchRequestDOM = Nothing
    
    Next
    
FindMortgageProductDetailsExit:

    'destory objects
    Set objOmMPBO = Nothing
    Set xmlMPSearchRequestDOM = Nothing
    Set xmlMPSearchResponseDOM = Nothing

Exit Sub
    
FindMortgageProductDetailsErrorOccured:

    Err.Raise Err.Number, "FindMortgageProductDetails." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume FindMortgageProductDetailsExit

End Sub

'------------------------------------------------------------------------
' Procedure SetNumberOfDependancies
' Author: MO
' Date: 14/08/2002
' Purpose: Sets the number of dependancies, this is done in code as the
'           dependancies data within the BM extract includes all other residents,
'           and the count of these dependancies where the individual was under 17 years
'           of age proved to complicated within XSL
' Input parameters: RequestXML, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub SetNumberOfDependants(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo SetNumberOfDependantsErrorOccured
    
    Const constMaximumAgeOfDependant As Long = 17
    
    Dim lngNumberOfDependants As Long
    Dim strDOB As String
    Dim lngAge As Long
    Dim objApplicantDependantNode As IXMLDOMNode
    Dim objNumberOfDependantsNode As IXMLDOMNode
    
    lngNumberOfDependants = 0
    
    'loop through all the ApplicantDependant nodes and count all the ones
    For Each objApplicantDependantNode In vxmlRequestNode.selectNodes("//ProposalRoot/Proposal/ApplicantRoot/Applicant/ApplicantDependantRoot/ApplicantDependant")
        
        strDOB = xmlGetMandatoryNode(objApplicantDependantNode, "DOB").Text
        
        If Len(strDOB) > 3 And IsDate(strDOB) = True Then
            
            'calculate the dependants age
            lngAge = Year(Now) - Year(strDOB)
            
            If Month(Now) < Month(strDOB) Then
                lngAge = lngAge - 1
            Else
                If Month(Now) = Month(strDOB) Then
                    If Day(Now) < Day(strDOB) Then
                        lngAge = lngAge - 1
                    End If
                End If
            End If
            
            If lngAge < constMaximumAgeOfDependant Then
                lngNumberOfDependants = lngNumberOfDependants + 1
            End If
            
        End If
        
    Next
    
    'set the number of dependants against the first customer in the XML
    Set objNumberOfDependantsNode = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "//APPLICATION/APPLICATIONFACTFIND/CUSTOMER[1]/CUSTOMERVERSION/@NUMBEROFDEPENDANTS")
    
    objNumberOfDependantsNode.Text = lngNumberOfDependants
    
SetNumberOfDependantsExit:

    'destory objects
    
    Set objApplicantDependantNode = Nothing
    Set objNumberOfDependantsNode = Nothing

Exit Sub
    
SetNumberOfDependantsErrorOccured:

    Err.Raise Err.Number, "SetNumberOfDependants." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume SetNumberOfDependantsExit

End Sub

'------------------------------------------------------------------------
' Procedure SetupLoanComponents
' Author: MO
' Date: 21/08/2002
' Purpose: Parses the data passed into ingestion engine in the products and
'           capital raising nodes to create the loan component nodes.
'           The loan components are created in the following order:
'               1. Capital Raising values
'                   1.1. against Products in order
'                       1.2.1 against capital And interest
'                       1.2.2 then interest only
'               2. Product amount remaining (that havent been taken up by capital raising values)
'
' Input parameters: RequestXML, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub SetupLoanComponents(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo SetupLoanComponentsErrorOccured
    
    Dim objMortgageSubQuoteNode As IXMLDOMNode
    Dim objProductsRootNode As IXMLDOMNode
    Dim objProductNode As IXMLDOMNode
    Dim objProductNodeList As IXMLDOMNodeList
    Dim objSortedProductNodeList As IXMLDOMNode
    Dim objCapitalRaisingRootNode As IXMLDOMNode
    Dim objCapitalRaisingNode As IXMLDOMNode
    Dim blnLoanComponentCreated As Boolean
    Dim dblProductCIAmount As Double
    Dim dblProductIOAmount As Double
    Dim strTermInYears As String
    Dim strTermInMonths As String
    Dim strPurposeOfLoan As String
    
    'one loan component must be created
    blnLoanComponentCreated = False
    
    'get the mortgage subquote node, the loan component nodes will be attached to this node
    Set objMortgageSubQuoteNode = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE")
    
    'create a clone of the products root, because im going to be changing some values
    Set objProductsRootNode = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/ProductsRoot").cloneNode(True)
    
    'create a clone of the capital raising root node cause im going to be changing some values
    Set objCapitalRaisingRootNode = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/ApplicantRoot/Applicant[1]/CapitalRaisingRoot").cloneNode(True)
    '[CapitalRaising/RemortgageCapitalAmount > '' and CapitalRaising/RemortgageCapitalAmount != '0']
    
    'get the mortgage term out of the request XML
    strTermInYears = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/Term").Text
    strTermInMonths = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/TermMonths").Text
    
    'sort the products, so that the product which the fees are added too is first
    Set objSortedProductNodeList = SortProducts(vxmlRequestNode, objProductsRootNode.cloneNode(True))
    
    'get the purpose of loan for this application, this will be put onto all loan components
    strPurposeOfLoan = DeterminePurposeOfLoan(vxmlRequestNode)
    
    'loop through the products and create the loan components
    For Each objProductNode In objSortedProductNodeList.selectNodes("Products")
        
        'get the product amounts out of the xml
        dblProductCIAmount = CDbl(xmlGetMandatoryNode(objProductNode, "ProductAmount").Text)
        dblProductIOAmount = CDbl(xmlGetMandatoryNode(objProductNode, "ProductAmountInterestOnly").Text)
        
        'loop through all the capital raising nodes, where they have an amount
        For Each objCapitalRaisingNode In objCapitalRaisingRootNode.selectNodes("CapitalRaising[RemortgageCapitalAmount > '' and RemortgageCapitalAmount != '0']")
                
            'try and create a loan component for the capital and raising amount
            ' against the products capital and interest amount
            If dblProductCIAmount > 0 Then
                'can this capital raising amount fit inside this products capital and interest amount
                If CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text) <= dblProductCIAmount Then
                    ' create the loan component for this capital raising amount
                    CreateLoanComponent objMortgageSubQuoteNode, xmlGetMandatoryNode(objProductNode, "SalesProductCode").Text, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text, 0, strTermInYears, strTermInMonths, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalReasonTypeId").Text, strPurposeOfLoan
                    blnLoanComponentCreated = True
                    ' and minus it off the products capital and interest amount
                    dblProductCIAmount = dblProductCIAmount - CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text)
                    ' set the capital and raising amount to zero as Ive now distributed this into a loan component
                    xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text = 0
                Else
                    ' create the loan component for the product capital and interest amount, against this capital raising
                    CreateLoanComponent objMortgageSubQuoteNode, xmlGetMandatoryNode(objProductNode, "SalesProductCode").Text, dblProductCIAmount, 0, strTermInYears, strTermInMonths, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalReasonTypeId").Text, strPurposeOfLoan
                    blnLoanComponentCreated = True
                    ' and minus it off the capital raising amount
                    xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text = CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text) - dblProductCIAmount
                    'set the capital and interest amount for the product to 0 as it has been dealt with
                    dblProductCIAmount = 0
                End If
            End If
            
            'if the capital raising amount hasnt been used yet, try and assigned it to the interest only portion
            If CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text) > 0 Then
            
                ' against the products interest only amount
                If dblProductIOAmount > 0 Then
                    
                    'can this capital raising amount fit inside this products capital and interest amount
                    If xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount") <= dblProductIOAmount Then
                        ' create the loan component for this capital raising amount
                        CreateLoanComponent objMortgageSubQuoteNode, xmlGetMandatoryNode(objProductNode, "SalesProductCode").Text, 0, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text, strTermInYears, strTermInMonths, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalReasonTypeId").Text, strPurposeOfLoan
                        blnLoanComponentCreated = True
                        ' and minus it off the products capital and interest amount
                        dblProductIOAmount = dblProductIOAmount - CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text)
                        ' set the capital and raising amount to zero as Ive now distributed this into a loan component
                        xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text = 0
                    Else
                        ' create the loan component for the product capital and interest amount, against this capital raising
                        CreateLoanComponent objMortgageSubQuoteNode, xmlGetMandatoryNode(objProductNode, "SalesProductCode").Text, 0, dblProductIOAmount, strTermInYears, strTermInMonths, xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalReasonTypeId").Text, strPurposeOfLoan
                        blnLoanComponentCreated = True
                        ' and minus it off the capital raising amount
                        xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text = CDbl(xmlGetMandatoryNode(objCapitalRaisingNode, "RemortgageCapitalAmount").Text) - dblProductIOAmount
                        ' set the capital and interest amount for the product to 0 as it has been dealt with
                        dblProductIOAmount = 0
                    End If
                    
                End If
                
            End If
            
        Next
        
        'if there is any value left for this product assign a loan component for it
        If dblProductCIAmount > 0 Or dblProductIOAmount > 0 Then
            
            CreateLoanComponent objMortgageSubQuoteNode, xmlGetMandatoryNode(objProductNode, "SalesProductCode").Text, dblProductCIAmount, dblProductIOAmount, strTermInYears, strTermInMonths, "", strPurposeOfLoan
            blnLoanComponentCreated = True
            
        End If
        
    Next
    
    If blnLoanComponentCreated = False Then
        Err.Raise direLoanComponentWasNotCreated, "SetupLoanComponents", "A loan component was not created, this could be due to missing Products details in the request xml."
    End If
    
SetupLoanComponentsExit:

    'destory objects
    Set objMortgageSubQuoteNode = Nothing
    Set objProductsRootNode = Nothing
    Set objProductNode = Nothing
    Set objCapitalRaisingRootNode = Nothing
    Set objCapitalRaisingNode = Nothing

Exit Sub

SetupLoanComponentsErrorOccured:

    Err.Raise Err.Number, "SetupLoanComponents." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume SetupLoanComponentsExit

End Sub

'------------------------------------------------------------------------
' Procedure CreateLoanComponent
' Author: MO
' Date: 21/08/2002
' Purpose: Creates a loan component xml structure within the Omiga XML Request that
'           will be inserted into the database
' Input parameters: objMortgageSubQuoteNode, strProductCode, dblCapitalAndInterestAmount,
'                   dblInterestOnlyAmount, strTermYears, strTermMonths, strSubPurposeOfLoan,
'                   strPurposeOfLoan
' Output parameters: None
'------------------------------------------------------------------------
Private Sub CreateLoanComponent(ByVal objMortgageSubQuoteNode As IXMLDOMNode, ByVal strProductCode As String, ByVal dblCapitalAndInterestAmount As Double, ByVal dblInterestOnlyAmount As Double, ByVal strTermYears As String, ByVal strTermMonths As String, ByVal strSubPurposeOfLoan As String, ByVal strPurposeOfLoan As String)
    
On Error GoTo CreateLoanComponentErrorOccured
    
    Dim objXMLOut As FreeThreadedDOMDocument40
    Dim objLoanComponent As IXMLDOMElement
    Dim strRepaymentMethod As String
    
    Set objXMLOut = New FreeThreadedDOMDocument40
    objXMLOut.validateOnParse = False
    objXMLOut.setProperty "NewParser", True
    
    'create a new loan component node and append it to the mortgage subquote
    Set objLoanComponent = objXMLOut.createElement("LOANCOMPONENT")
    objMortgageSubQuoteNode.appendChild objLoanComponent
    
    objLoanComponent.setAttribute "MORTGAGEPRODUCTCODE", strProductCode
    objLoanComponent.setAttribute "STARTDATE", ""
    
    'Work out repayment type
    ' interest only
    If dblInterestOnlyAmount <> 0 And dblCapitalAndInterestAmount = 0 Then
        strRepaymentMethod = 1
    End If
    'capital and interest
    If dblInterestOnlyAmount = 0 And dblCapitalAndInterestAmount <> 0 Then
        strRepaymentMethod = 2
    End If
    'part and part
    If dblInterestOnlyAmount <> 0 And dblCapitalAndInterestAmount <> 0 Then
        strRepaymentMethod = 3
    End If
    
    objLoanComponent.setAttribute "REPAYMENTMETHOD", strRepaymentMethod
    objLoanComponent.setAttribute "LOANAMOUNT", dblInterestOnlyAmount + dblCapitalAndInterestAmount
    
    'only populate the capitalandinterest and interestonly elements if it part and part
    If strRepaymentMethod = 3 Then
        objLoanComponent.setAttribute "CAPITALANDINTERESTELEMENT", dblCapitalAndInterestAmount
        objLoanComponent.setAttribute "INTERESTONLYELEMENT", dblInterestOnlyAmount
    Else
        objLoanComponent.setAttribute "CAPITALANDINTERESTELEMENT", ""
        objLoanComponent.setAttribute "INTERESTONLYELEMENT", ""
    End If
    
    'purpose of loan
    objLoanComponent.setAttribute "PURPOSEOFLOAN", strPurposeOfLoan
    
    'subpurpose of loan
    objLoanComponent.setAttribute "SUBPURPOSEOFLOAN", strSubPurposeOfLoan
    objLoanComponent.setAttribute "TERMINMONTHS", strTermMonths
    objLoanComponent.setAttribute "TERMINYEARS", strTermYears
    objLoanComponent.setAttribute "TOTALLOANCOMPONENTAMOUNT", dblInterestOnlyAmount + dblCapitalAndInterestAmount
    
CreateLoanComponentExit:

    'destory objects
    Set objXMLOut = Nothing
    Set objLoanComponent = Nothing

Exit Sub

CreateLoanComponentErrorOccured:

    Err.Raise Err.Number, "CreateLoanComponent." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume CreateLoanComponentExit

End Sub

'------------------------------------------------------------------------
' Procedure SortProducts
' Author: MO
' Date: 27/08/2002
' Purpose: Sorts the products into the correct order so that they are in
'           the correct order to be entered into the database, basically
'           so that the product that the fees are added to, goes in first
' Input parameters: objProductsList
' Output parameters: objSortedProductsList
'------------------------------------------------------------------------
Private Function SortProducts(ByVal vxmlRequestNode As IXMLDOMNode, ByVal objProductsList As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo SortProductsErrorOccured
    
    Dim objSortedProductsList As IXMLDOMNode
    Dim objSortedProductsListDOM As FreeThreadedDOMDocument40
    Dim objFirstProduct As IXMLDOMNode
    Dim objProductNode As IXMLDOMNode
    Dim strFeeControlProduct As String
    
    Set objSortedProductsListDOM = New FreeThreadedDOMDocument40
    objSortedProductsListDOM.validateOnParse = False
    objSortedProductsListDOM.setProperty "NewParser", True
    Set objSortedProductsList = objSortedProductsListDOM.createElement("PRODUCTLIST")
    
    'get the first product out and add it to the sorted list
    strFeeControlProduct = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/FeeControlProduct").Text
    
    'if there isnt a fee control product, dont find the first product
    If Len(strFeeControlProduct) > 0 Then
    
        Set objFirstProduct = xmlGetMandatoryNode(objProductsList, "Products[Id = '" & strFeeControlProduct & "']")
        
        objSortedProductsList.appendChild objFirstProduct.cloneNode(True)
        
        'ive dealt with the first product, so ill remove it from the productslist
        objProductsList.removeChild objFirstProduct
        
    End If
        
    'loop through the remaining products and add them to the sorted product list
    For Each objProductNode In objProductsList.selectNodes("Products")
        objSortedProductsList.appendChild objProductNode.cloneNode(True)
    Next
    
    Set SortProducts = objSortedProductsList
    
SortProductsExit:

    'destory objects
    Set objSortedProductsListDOM = Nothing
    Set objSortedProductsList = Nothing
    Set objFirstProduct = Nothing
    Set objProductNode = Nothing

Exit Function

SortProductsErrorOccured:

    Err.Raise Err.Number, "SortProducts." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume SortProductsExit

End Function

'------------------------------------------------------------------------
' Procedure DeterminePurposeOfLoan
' Author: MO
' Date: 28/10/2002
' Purpose: determines what the purpose of the loan should be for the loan components
' Input parameters: vxmlRequestNode as IXMLDOMNode
' Output parameters: strPurposeOfLoan
'------------------------------------------------------------------------
Private Function DeterminePurposeOfLoan(ByVal vxmlRequestNode As IXMLDOMNode) As String
    
On Error GoTo DeterminePurposeOfLoanErrorOccured
        
    'POS Constants
    Const cstrPOSRemortgageValue As String = "2"
    Const cstrPOSBTLValue As String = "2"
    
    'O4 Constants
    Const cstrO4HousePurchaseNonFTBValue As String = "1"
    Const cstrO4HousePurchaseFTBValue As String = "2"
    Const cstrO4HousePurchaseBTLValue As String = "3"
    Const cstrO4RemortReplaceOrigValue As String = "4"
    Const cstrO4RemortReplaceBTLValue As String = "5"
    Const cstrO4RemortCapitalRaisingValue As String = "8"
    Const cstrO4MortgageFreeSPLCapitalRaisingValue As String = "20"
    
    Dim blnRemortgage As Boolean
    Dim blnBTL As Boolean
    Dim blnCapitalRaising As Boolean
    Dim blnHaveAMortgage As Boolean
    Dim blnFirstTimeBuyer As Boolean
    
    Dim objProposalNode As IXMLDOMNode
    
    Set objProposalNode = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal")
    
    'determine if this is a house purchase or a remortgage
    If xmlGetMandatoryNode(objProposalNode, "MortgageStatusTypeId").Text = cstrPOSRemortgageValue Then
        blnRemortgage = True
    Else
        blnRemortgage = False
    End If
    
    'determine if this is a buy to let
    If xmlGetMandatoryNode(objProposalNode, "MortgageTypeId").Text = cstrPOSBTLValue Then
        blnBTL = True
    Else
        blnBTL = False
    End If
    
    'determine if this is a capital raising mortgage
    If objProposalNode.selectSingleNode("ApplicantRoot/Applicant[1]/CapitalRaisingRoot/CapitalRaising[RemortgageCapitalAmount > '' and RemortgageCapitalAmount != '0']") Is Nothing Then
        blnCapitalRaising = False
    Else
        blnCapitalRaising = True
    End If
    
    'determine if they have a mortgage on there current address
    If objProposalNode.selectSingleNode("ApplicantRoot/Applicant/ApplicantAddressRoot/ApplicantAddress[AddressTypeId = '2' and HaveAMortgageYN = '1']") Is Nothing Then
        blnHaveAMortgage = False
    Else
        blnHaveAMortgage = True
    End If
    
    'determine if this is a first time buyer case - all applicants need to be first time buyer
    If objProposalNode.selectNodes("ApplicantRoot/Applicant").length = objProposalNode.selectNodes("ApplicantRoot/Applicant[FirstTimeBuyerYN = '1']").length Then
        blnFirstTimeBuyer = True
    Else
        blnFirstTimeBuyer = False
    End If
    
    If blnRemortgage = False Then
        'house purchase
        If blnBTL = True Then
            'buy to let
            DeterminePurposeOfLoan = cstrO4HousePurchaseBTLValue
        ElseIf blnFirstTimeBuyer = True Then
            'first time buyer
            DeterminePurposeOfLoan = cstrO4HousePurchaseFTBValue
        Else
            'non first time buyer
            DeterminePurposeOfLoan = cstrO4HousePurchaseNonFTBValue
        End If
    Else
        'remortgage
        If blnBTL = True Then
            'buy to let
            DeterminePurposeOfLoan = cstrO4RemortReplaceBTLValue
        ElseIf blnHaveAMortgage = False And blnCapitalRaising = True Then
            'mortgage free spl - capital raising
            DeterminePurposeOfLoan = cstrO4MortgageFreeSPLCapitalRaisingValue
        ElseIf blnCapitalRaising = True Then
            'capital raising
            DeterminePurposeOfLoan = cstrO4RemortCapitalRaisingValue
        Else
            'just remortgage
            DeterminePurposeOfLoan = cstrO4RemortReplaceOrigValue
        End If
    End If
    
DeterminePurposeOfLoanExit:

    'destory objects
    Set objProposalNode = Nothing
    
Exit Function

DeterminePurposeOfLoanErrorOccured:

    Err.Raise Err.Number, "DeterminePurposeOfLoan." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume DeterminePurposeOfLoanExit

End Function

'------------------------------------------------------------------------
' Procedure ExtractPreprocessingRulesRequest
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the request from preprocessing and extracts the 3 main nodes
'           of the operation
' Input parameters: vstrRulesResponse holding the response from the rules component
'                   vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Sub ExtractPreprocessingRulesRequest(ByVal strRulesResponse As String, ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ExtractPreprocessingRulesRequestErrorOccured
    
    Dim xmlRulesRequestDOM As FreeThreadedDOMDocument40
    
    Set xmlRulesRequestDOM = xmlLoad(strRulesResponse)
    
    'get the original request node
    Set vxmlRequestNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "ORIGINALREQUEST/REQUEST").cloneNode(True)
    
    'get the omiga formatted request node
    Set vxmlOmigaFormatRequestNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "OMIGAFORMATTEDREQUEST/REQUEST").cloneNode(True)
    
    'get the current response node
    Set vxmlResponseNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "CURRENTRESPONSE/RESPONSE").cloneNode(True)
    
ExtractPreprocessingRulesRequestExit:
    
    'destroy objects
    Set xmlRulesRequestDOM = Nothing
    
    Exit Sub
    
ExtractPreprocessingRulesRequestErrorOccured:

    Err.Raise Err.Number, "ExtractPreprocessingRulesRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume ExtractPreprocessingRulesRequestExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildPreprocessingRulesResponse
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the 3 main nodes of the operation and converts them into
'           the string response needed for the Rules component
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildPreprocessingRulesResponse(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As String

    On Error GoTo BuildPreprocessingRulesResponseErrorOccured
    
    Dim xmlRulesResponseDOM As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOriginalRequestNode As IXMLDOMNode
    Dim xmlOmigaFormattedRequestNode As IXMLDOMNode
    Dim xmlCurrentResponseNode As IXMLDOMNode
    
    Set xmlRulesResponseDOM = New FreeThreadedDOMDocument40
    xmlRulesResponseDOM.validateOnParse = False
    xmlRulesResponseDOM.setProperty "NewParser", True
    'create the top level request node
    Set xmlRequestNode = xmlRulesResponseDOM.createElement("REQUEST")
    xmlRulesResponseDOM.appendChild xmlRequestNode
    
    'create the original request node
    Set xmlOriginalRequestNode = xmlRulesResponseDOM.createElement("ORIGINALREQUEST")
    xmlRequestNode.appendChild xmlOriginalRequestNode
    'append the original request node to it!
    xmlOriginalRequestNode.appendChild vxmlRequestNode.cloneNode(True)
    
    'create the omiga formatted request node
    Set xmlOmigaFormattedRequestNode = xmlRulesResponseDOM.createElement("OMIGAFORMATTEDREQUEST")
    xmlRequestNode.appendChild xmlOmigaFormattedRequestNode
    'append the omiga formatted request to it
    xmlOmigaFormattedRequestNode.appendChild vxmlOmigaFormatRequestNode.cloneNode(True)
    
    'create the current response node
    Set xmlCurrentResponseNode = xmlRulesResponseDOM.createElement("CURRENTRESPONSE")
    xmlRequestNode.appendChild xmlCurrentResponseNode
    'append the response node to it
    xmlCurrentResponseNode.appendChild vxmlResponseNode.cloneNode(True)
        
    'return the xml as a string
    BuildPreprocessingRulesResponse = xmlRulesResponseDOM.xml

BuildPreprocessingRulesResponseExit:
    
    'destroy objects
    Set xmlRulesResponseDOM = Nothing
    Set xmlRequestNode = Nothing
    Set xmlOriginalRequestNode = Nothing
    Set xmlOmigaFormattedRequestNode = Nothing
    Set xmlCurrentResponseNode = Nothing

    Exit Function
    
BuildPreprocessingRulesResponseErrorOccured:

    Err.Raise Err.Number, "BuildPreprocessingRulesResponse." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume BuildPreprocessingRulesResponseExit
    
End Function

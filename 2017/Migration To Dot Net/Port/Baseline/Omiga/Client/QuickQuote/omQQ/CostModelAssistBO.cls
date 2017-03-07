VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CostModelAssistBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CostModelAssistBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Component to link cost modelling to AiP
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' MCS   29/11/99 Created
' DJP   28/02/00 Added FindMortgageProducts and related functions
' DJP   29/02/00 Added GetDefaultsForNewLoanComponent
' DJP   16/03/00 Added CalcAndSaveLifeCover, BuildLifeCoverInterfaceXML, GetTermAndEndowmentNeeds,
'                BuildLifeCoverCalcInputInterface, CalculateAndSaveLifeCoverCosts, CalculateMortgageCosts,
'                CalcCostModelLTV
' PSC   07/04/00 SetComplete rather than SetAbort in GetLoansNotRedeemed
' JLD   11/04/00 SYS0465 - FindMortgageProducts() should use MortgageProductBO rather than DO so the
'                Product list is verified.
' JLD   12/04/00 Allow SharedOwnershipAmount to be empty in CalcCostModelLTV - set it to '0'
' PSC   14/04/00 AQR SYS0602: Amend to allow DiscountAmount and SharedOwnershipAmount to be empty
' JLD   27/04/00 Match the context correctly to "Quick Quote" or "Cost Modelling"
' JLD   15/05/00 SYS0715 added updateMortgageSubquote to be used by QQ and AQ
' JLD   18/05/00 SYS0725 add customerlist to save of life benefits
' MCS   26/05/00 SYS0724 Change to match spec
' DJP   31/05/00 SYS0610 - removed CalculateAndSaveLifeCoverCosts as it wasn't being used.
' PSC   08/06/00 Correct strTypeOfApplication to strTypeOfBuyer
' PSC   19/06/00 SYS0792 Amend Shared Ownership Tag names due to database changes
' PSC   30/06/00 SYS1012 Amend to use revised GetLoansNotRedeemed method and distinguish
'                        between Quick Quote and Cost Modelling modes
' MS    11/07/00 SYS1198 - Read comments 1 & 2 for modification to ported lender for phase 2.
' MS    14/07/00 SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
' MS    19/07/00 Modified for performance enhancements
' DM    11/08/00 SYS1173 Cant reinstate subquote
' DJP   13/09/00 SYS1031 - Change FindMortgageProducts to correctly handle self employed applicant
' CL    30/11/00 SYS1685 Change to call amended Mortgage Lender
' ASm   12/01/01 SYS1817: oeInvalidNumberOfCustomers error enum removed and replaced by its respective error number
' DRC   23/05/01 SYS2093: Change to CalcCostModelLTV to allow for Valuation Report amount as well as Purchase Price
' DRC    3/10/01 SYS2745: Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
' CL    07/05/02 SYS4510  Modifications to class initialise
' CL    10/05/02 SYS4510  Remove class initialize & class terminate
' SG    28/05/02 SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------
Option Explicit
Implements ICostModelAssistBO
Implements ObjectControl

Private Const cstrLIFESUBQUOTE_COMP As String = "omCM"

Private Const cINTEREST_ONLY As String = "I"
Private Const cCAPITAL_AND_INTEREST As String = "C"
Private Const cPART_AND_PART As String = "P"

' Private member variables
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist
Private m_objConvertAssist As ConvertAssist
Private m_objContext As ObjectContext
Private m_objGeneralAssist As GeneralAssist

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     10/06/2002  BMIDS00032  Modified ICostModelAssistBO_GetDefaultsForNewLoanComponent
'PSC    11/06/2002  BMIDS00043  Get Completion Date to be used in CalculateMortgageCosts
'                               and RecostMortgageComponentAndLifeCover
'MV     17/06/2002  BMIDS00032  Modified ICostModelAssistBO_GetDefaultsForNewLoanComponent
'PSC    17/06/2002  BMIDS00043  Late bind omRotBO due to build problems
'MV     20/06/2002  BMIDS00077  Core Upgrade - Modified ICostModelAssistBO_CalculateMortgageCosts
'GD     02/07/2002  BMIDS00166  Removed GOODHEALTH and SMOKER from BuildLifeCoverInterfaceXML
'PSC    11/07/2002  BMIDS00062  Amend RecostMortgageComponentAndLifeCover to use composite rates
'MV     18/07/2002  BMIDS00179  Core Upgrade Rollback - Modified ReInstateQuotation , ValidateAndReInstateQuotation()
'                               FindMortgageProducts(),CalcCostModelLTV(), CalculateMortgageCosts()
'MV     04/09/2002  BMIDS00353  Modified ICostModelAssistBO_CalculateMortgageCosts
'GD     06/09/2002  BMIDS00256  Modified ICostModelAssistBO_CalcCostModelLTV
'MV     10/09/2002  BMIDS00353  Amended ICostModelAssistBO_CalculateMortgageCosts at Stage7 to handle No Record found Error
'MV     26/09/2002  BMIDS00522  Amended ICostModelAssistBO_CalculateMortgageCosts at Stage7
'MV     03/10/2002  BMIDS00430,BMIDS00536   Amended ICostModelAssistBO_CalculateMortgageCosts at Stage1 and Stage2
'MV     07/10/2002  BMIDS00588  Amended ICostModelAssistBO_ProcessAddedOneOffCosts
'SA     08/10/2002  BMIDS00246  Amended FindMortgageProducts to call new method  objICustomerEmploymentDO.GetEmploymentStatusList
'SA     09/10/2002  BMIDS00246  Removed all code pertaining to IsSelfEmployed as it's no longer used for finding products
'MV     11/10/2002  BMIDS00590  Amended ICostModelAssistBO_ProcessAddedOneOffCosts,ICostModelAssistBO_CalculateMortgageCosts
'                               Modified ICostModelAssistBO_CalcCostModelLTV
'SA     17/10/2002  BMIDS00662  Modified FindProductList - building of employment status list string
'MO     22/10/2002  BMIDS00450  Modified CalcCostModelLTV to fix type mismatch bug.
'MV     23/10/2002  BMIDS00430  Modified at Stage1 in ICostModelAssistBO_CalculateMortgageCosts
'                               Modified at Stage3 and Stage 5 in ICostModelAssistBO_CalcCostModelLTV
'MO     24/10/2002  BMIDS00711  Modifed CalcCostModelLTV to fix invalid parameter bug.
'SA     05/11/2002  BMIDS00818  Removed Validate user mandate level from CalulateMortgageCosts
'DPF    11/11/2002  BMIDS00808  added processing for recalculating Drawdown costs within ICostModelAssistBO_RecostMortgageComponentAndLifeCover
'MDC    18/11/2002  BMIDS00938  CC014
'MDC    20/11/2002  BMIDS01024
'DPF    20/11/2002  BMIDS01030  Fix added to CostModelAssistBO.ICostModelAssistBO_CalculateMortgageCosts
'SR     02/01/2003  BM0209      Modified method CalculateMortgageCosts
'MV     05/02/2003  BM0235      Amended ICostModelAssistBO_GetDefaultsForNewLoanComponent
'MV     05/02/2003  BM0235      Amended ICostModelAssistBO_GetDefaultsForNewLoanComponent
'GHun   02/07/2003  BMIDS591    CC040 Amended CalcCostModelLTV and CalculateMortgageCosts to added extra boolean parameter
'GHun   22/04/2004  BMIDS736    Amend RecostMortgageComponentAndLifeCover for Alpha+ integration
'MC     13/06/2004  BMIDS763    ApplicationDate data element added to the request.
'INR    05/07/2004  BMIDS774    Use ApplicationFactFind EstimatedCompletion Date instead of ROT Completion date.
'MC     07/07/2004  BMIDS763    ApplicateDate is passed with every request with in "ProcessAddedOneOffCosts()"
'GHun   15/07/2004  BMIDS793    Amended RecostMortgageComponentAndLifeCover to handle drawdown = 0
'JD     19/07/04    BMIDS763    CalculateMortgageCosts, XML change
'KRW    29/09/04    BM0561      FURTHER ADV / SPL - Existing term / repayment method not transferred for 2nd loan component.
'------------------------------------------------------------------------------------------
'MARS Specific History:
'HMA    26/07/2005  MAR18       Default Employment status in FindProducts if not found.
'JD     14/10/2005  MAR41       Changed CalcCostModelLTV and CalucaulteMortgageCosts to get valuation
'                               from Hometrack or valuationreport depending on valuation type.
'JD     11/11/2005  MAR506      hometrackBO expects attribute based xml in CalcLTV method
'AE     14/11/2005  MAR506      declare strAppNo and StrAppFFNo
'JD     18/11/2005  MAR299      in CalcCostModelLTV if the hometrack valuation is 0 use the purchase price.
'JD     22/11/2005  MAR299      set correct appno in calculateMortageCosts and check that hometrack valuation amount is not 0
'HMA    29/11/2005  MAR696      Allow for no application number in FindMortgageProducts
'JD     09/03/2006  MAR1061     ValidateAndReinstateQuotation - update AFF with reinstated quote's purchaseprice
'GHun   05/04/2006  MAR1300     Change of Property changes
'JD     25/04/2006  MAR1640     CalcCostModelLTV - if the presentvaluation is 0 then set the LTV as default 999
'PE     08/05/2006  MAR1702     Maximum Applicant Age Incorrect
'PE     12/05/2006  MAR1757     Handle date of birth not set
'GHun   10/07/2006  MAR1893     Commented out redundant LifeSubQuote code
'------------------------------------------------------------------------------------------

Public Function IsSelfEmployed(ByVal vXMLElemRequest As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
'   No records are trapped and just IsSelfEmployed is returned as FALSE
'
' pass:
'   vXMLElemRequest  xml Request element containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   IsSelfEmployed  Boolean indicating if ANY of the customers passed in are SelfEmployed
'
'
' Raise Errors: none
'------------------------------------------------------------------------------------------
On Error GoTo IsSelfEmployedVbErr

    Const strFunctionName As String = "IsSelfEmployed"

    
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim objXmlAssist As New XMLAssist
    Dim objConvertAssist As New ConvertAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement

    Dim strXML As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerEmploymentDO = objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentDO")
    End If

    Dim strSQL As String
    Dim objClassDef As FreeThreadedDOMDocument40
    
    Set objClassDef = LoadDataForQuickQuoteApplicantDetails

    Set xmlElement = objDOAssist.FindListMultipleEx(vXMLElemRequest, objClassDef)
    'DO NOT error on no record found

        Dim nloop As Integer
        Dim xmlNodeList As IXMLDOMNodeList
        Dim blnEmploymentStatus As Boolean
        blnEmploymentStatus = False
    If Not xmlElement Is Nothing Then
        'get the number of employment records out
        Set xmlNodeList = xmlElement.getElementsByTagName("QUICKQUOTEAPPLICANTDETAILS")
    
        For nloop = 0 To (xmlNodeList.length - 1)
            
            Set xmlElement = xmlNodeList.Item(nloop)
            'check the employment status we want selfemployed
            'this is inefficient why not get the value from the combo once then compare ??
            If objConvertAssist.CSafeBool(objICustomerEmploymentDO.IsEmploymentStatus(SelfEmployed, objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS"))) Then
                'if any are true
                blnEmploymentStatus = True
    
                Exit For ' get out of this loop
            End If
        Next

    End If
    
    IsSelfEmployed = blnEmploymentStatus
    
IsSelfEmployedExit:

    Set objDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objClassDef = Nothing
    Set objConvertAssist = Nothing
    Set objICustomerEmploymentDO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function

IsSelfEmployedVbErr:
    
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    Set objDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objClassDef = Nothing
    Set objConvertAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing

'    objErrAssist.AddToErrSource strFunctionName

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function LoadDataForQuickQuoteApplicantDetails() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
    Const strFunctionName As String = "LoadDataForQuickQuoteApplicantDetails"

    
    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
       
    strXML = strXML & "<TABLENAME>QUICKQUOTEAPPLICANTDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>EMPLOYMENTSTATUS"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>EmploymentStatus</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set LoadDataForQuickQuoteApplicantDetails = xmlClassDefDoc
    
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing

End Function

Private Function ICostModelAssistBO_CalculateAndSaveBCSubQuote( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal strCustNumber1 As String, _
                            ByVal strCustNumber2 As String, _
                            ByVal strCustVersionNumber1 As String, _
                            ByVal strApplNumber As String, _
                            ByVal strAFFNumber As String, _
                            Optional ByVal strContext As String = "QUICK QUOTE", _
                            Optional ByVal blnBakruptcyApplicant1 As Boolean, _
                            Optional ByVal blnBakruptcyApplicant2 As Boolean, _
                            Optional ByVal strDeclarationValue As String, _
                            Optional xmlNewProperty As IXMLDOMElement) As IXMLDOMElement

On Error GoTo CalculateAndSaveBCSubQuoteVbErr

    Const strFunctionName As String = "ICostModelAssistBO_CalculateAndSaveBCSubQuote"


    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement, xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode, xmlTableElem As IXMLDOMNode
    Dim xmlQuoteStatusList As IXMLDOMNodeList
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlValuablesOverLimit As IXMLDOMNode
    Dim xmlPolarisOutput As IXMLDOMNode
    Dim xmlBCDetailsStatus As IXMLDOMNode
    
    Dim objICustClassDef As IomCustClassDef
    Dim objIDOAssist As IDOAssist, objBCSubQuoteBO As BuildingsAndContentsSubQuoteBO
    Dim objBCBO As BuildingsAndCOntentsBO
    Dim objComboDO As ComboDO
    
    Dim strResponse As String, strDOB As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
     If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objBCSubQuoteBO = New BuildingsAndContentsSubQuoteBO
        Set objBCBO = New BuildingsAndCOntentsBO
        Set objComboDO = New ComboDO
    Else
        Set objIDOAssist = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objBCSubQuoteBO = _
            m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".BuildingsAndContentsSubQuoteBO")
        Set objBCBO = _
            m_objContext.CreateInstance(gstrBUILDINGSANDCONTENTS & ".BuildingsAndContentsBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
        
    '--------------------------------------------------------------------
    'Get date of birth of the first applicant required for the interface
    '--------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("CUSTOMERVERSION")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCustNumber1
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCustVersionNumber1
    xmlTableElem.appendChild xmlElement

    Set objICustClassDef = New omCustClassDef
    Set xmlTempResponseNode = _
        objIDOAssist.GetData(xmlNewRequest, objICustClassDef.LoadCustomerVersionData(omCustVersionDOB))
    
    strDOB = xmlTempResponseNode.selectSingleNode(".//DATEOFBIRTH").Text
    
    '------------------------------------------------------------------
    'Retrieve valuables over limit details
    '------------------------------------------------------------------
    Set xmlValuablesOverLimit = vxmlRequest.selectSingleNode(".//VALUABLESOVERLIMITLIST")
    
    '-------------------------------------------------------------------------------
    'Build Input Interface to Polaris
    '-------------------------------------------------------------------------------
    ' Build the request block
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER1")
    xmlElement.Text = strCustNumber1
    xmlNewRequest.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER2")
    xmlElement.Text = strCustNumber2
    xmlNewRequest.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("DATEOFBIRTH")
    xmlElement.Text = strDOB
    xmlNewRequest.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("BANKRUPTCYAPPLICANT1")
    ' Assign value only if it was passed in
    If strContext <> "QUICK QUOTE" Then
        'APS 04/05/2000 - Changed the text value from TRUE/FALSE to 1/0
        If blnBakruptcyApplicant1 Then
            xmlElement.Text = "1"
        Else
            xmlElement.Text = "0"
        End If
    End If
    xmlNewRequest.appendChild xmlElement
    
    ' APS 04/05/2000 - Changed Tag name
    Set xmlElement = xmlOut.createElement("BANKRUPTCYAPPLICANT2")
    If strContext <> "QUICK QUOTE" Then
        'APS 04/05/2000 - Changed the text value from TRUE/FALSE to 1/0
        If blnBakruptcyApplicant1 Then
            xmlElement.Text = "1"
        Else
            xmlElement.Text = "0"
        End If
    End If
    xmlNewRequest.appendChild xmlElement
    
    ' APS 04/05/2000 - Changed Tag name
    Set xmlElement = xmlOut.createElement("DECLARATIONVALUE")
    If strContext <> "QUICK QUOTE" Then
        If blnBakruptcyApplicant1 Then
            xmlElement.Text = strDeclarationValue
        End If
    End If
    xmlNewRequest.appendChild xmlElement
    
    ' Add Buildings&Contents XML passed the Request
    xmlNewRequest.appendChild vxmlRequest.selectSingleNode(".//BCSUBQUOTEDETAILSNOTES").cloneNode(True)
    
    ' Add NewProperty XML only, if it was passed in
    If strContext <> "QUICK QUOTE" Then
        xmlNewRequest.appendChild xmlNewProperty
    End If
    
    If Not xmlValuablesOverLimit Is Nothing Then
        'm_objXMLAssist.AttachResponseData xmlValuablesOverLimit, xmlNewRequest
        xmlNewRequest.appendChild xmlValuablesOverLimit
    End If
    
    strResponse = objBCSubQuoteBO.BuildBCCalcInterface(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    '-------------------------------------------------------------------------------
    'Run Polaris
    '-------------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    m_objXMLAssist.AttachResponseData xmlNewRequest, xmlDoc.documentElement
    strResponse = objBCBO.Calculate(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    Set xmlPolarisOutput = xmlDoc.documentElement
        
    '-------------------------------------------------------------------------
    'Save Quote Details
    '-------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    m_objXMLAssist.AttachResponseData xmlNewRequest, xmlDoc.documentElement
    xmlNewRequest.appendChild vxmlRequest.selectSingleNode(".//BCSUBQUOTEDETAILSNOTES").cloneNode(True)

    strResponse = objBCSubQuoteBO.SaveSubQuote(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
    xmlResponseElem.appendChild xmlPolarisOutput
    
    Set ICostModelAssistBO_CalculateAndSaveBCSubQuote = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAndSaveBCSubQuoteExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlValuablesOverLimit = Nothing
    Set xmlBCDetailsStatus = Nothing
    
    Set objICustClassDef = Nothing
    Set objIDOAssist = Nothing
    Set objBCSubQuoteBO = Nothing
    Set objBCBO = Nothing
    
    Exit Function

CalculateAndSaveBCSubQuoteVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_CalculateAndSaveBCSubQuote = _
                            m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit

End Function

Private Function ICostModelAssistBO_FindStoredQuoteDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
'-------------------------------------------------------------------------------------------
'Description : This does all the processing required for the methods
'              'findStoredQuoteDetails' in the objects ApplicationQuote and QuickQuote
'-------------------------------------------------------------------------------------------
On Error GoTo FindStoredQuoteDetailsVbErr

    Const strFunctionName As String = "ICostModelAssistBO_FindStoredQuoteDetails"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlApplData As IXMLDOMNode
    
    Dim objApplicationBO As ApplicationBO
    Dim objQuotationBO As QuotationBO
    
    Dim strResponse As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objQuotationBO = New QuotationBO
    Else
        Set objApplicationBO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objQuotationBO = _
                m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    
    '--------------------------------------------------------------------------
    'Retrieve ApplicationFactFind details
    '--------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableNode = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableNode
    
    Dim strTag As String
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableNode.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableNode.appendChild xmlElement
    
    strResponse = objApplicationBO.GetApplicationData(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set xmlApplData = xmlDoc.documentElement
    
    '-------------------------------------------------------------------------
    'Retrieve Quotation details
    '-------------------------------------------------------------------------
    m_objXMLAssist.ChangeNodeName xmlNewRequest, "APPLICATION", "QUOTATION"
    strResponse = objQuotationBO.FindStoredQuoteDetails(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    '--------------------------------------------------------------------------
    'Append the Application & Quotation data to Response and Exit
    '---------------------------------------------------------------------------
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlApplData
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlDoc.documentElement
    
    Set ICostModelAssistBO_FindStoredQuoteDetails = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    Set xmlElement = Nothing
    Set xmlApplData = Nothing
    
    Set objApplicationBO = Nothing
    Set objQuotationBO = Nothing

    Exit Function
FindStoredQuoteDetailsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_FindStoredQuoteDetails = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Private Function ICostModelAssistBO_ProcessAddedOneOffCosts( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            Optional ByVal vstrContext As String = "QUICK QUOTE", _
                            Optional ByVal vstrLocation As String, _
                            Optional ByVal vstrValuationType As String, _
                            Optional ByVal vstrLoansNotRedeemed As String, _
                            Optional ByVal vstrSharedPercentage As String) As IXMLDOMElement

    On Error GoTo ProcessAddedOneOffCostsVbErr

    Const strFunctionName As String = "ICostModelAssistBO_ProcessAddedOneOffCosts"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlApplData As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode
    Dim xmlAppStaffReqElem  As IXMLDOMElement
    Dim xmlAppStaffElement  As IXMLDOMElement
    Dim xmlAppStaffRespDoc  As New FreeThreadedDOMDocument40
    xmlAppStaffRespDoc.validateOnParse = False
    xmlAppStaffRespDoc.setProperty "NewParser", True
    Dim xmlAppStaffNode  As IXMLDOMNode
    Dim xmlAppNo  As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objComboDO As ComboDO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    Dim objApplicationManagerBO As ApplicationManagerBO
        
    Dim strTag As String
    Dim strResponse As String
    
    'Prepare Response Object
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Create Object Instance
    If m_objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objComboDO = New ComboDO
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objApplicationManagerBO = New ApplicationManagerBO
    Else
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objApplicationManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    End If
    
'Stage1:
    
    'Retreive AIP & Application Data
    
    Set xmlApplData = objICostModelAssistDO.GetApplicationData(vxmlRequest)
       
    Set xmlAppNo = xmlApplData.selectSingleNode(".//APPLICATIONNUMBER")
    Set xmlAppFactFindNo = xmlApplData.selectSingleNode(".//APPLICATIONFACTFINDNUMBER")
    
    ' Find Loans not redeemed, Location and valuationType for QuickQuote.
    ' For AppQuote, they are passed in
    If StrComp(vstrContext, "QUICK QUOTE", vbTextCompare) = 0 Then
        vstrLocation = objComboDO.GetFirstComboValueId("PropertyLocation", _
                    objGlobalParameterDO.GetCurrentParameterByType("QQLocation", "STRING"))
        vstrValuationType = objComboDO.GetFirstComboValueId("ValuationType", _
                    objGlobalParameterDO.GetCurrentParameterByType("QQValuationType", "STRING"))
        
        vstrLoansNotRedeemed = xmlApplData.selectSingleNode(".//OUTSTANDINGLOANAMOUNT").Text
    End If
    
'Stage2:
    
    'Is application Member Of Staff
    Set xmlAppStaffReqElem = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlAppStaffElement = xmlOut.createElement("APPLICATION")
    xmlAppStaffReqElem.appendChild xmlAppStaffElement
    
    xmlAppStaffElement.appendChild xmlAppNo
    xmlAppStaffElement.appendChild xmlAppFactFindNo
    
    strResponse = objApplicationManagerBO.IsStaffApplication(xmlAppStaffReqElem.xml)
    Set xmlAppStaffRespDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlAppStaffRespDoc.documentElement, xmlResponseElem, True
    
    Set xmlAppStaffNode = xmlAppStaffRespDoc.selectSingleNode(".//MEMBEROFSTAFF")

'Stage3:

    'Process the One-Off costs

    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlNewRequest.appendChild xmlTableElem
    
    'ApplicationNumber
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    '*=[MC]BMIDS763 - APPLICATIONDATE
    strTag = "APPLICATIONDATE"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    '*=BMIDS763 SECTION END
    
    strTag = "MORTGAGESUBQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//ONEOFFCOSTLIST").cloneNode(True)
    
    'Location
    strTag = "LOCATION"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vstrLocation
    xmlTableElem.appendChild xmlElement
    
    'TypeOfValuation
    strTag = "TYPEOFVALUATION"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vstrValuationType
    xmlTableElem.appendChild xmlElement
    
    'TypeOfApplication
    xmlTableElem.appendChild xmlApplData.selectSingleNode(".//TYPEOFAPPLICATION")
    
    'PurchasePrice
    strTag = "PURCHASEPRICE"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = xmlApplData.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE").Text
    xmlTableElem.appendChild xmlElement
    
    'LoansNotRedeemed
    strTag = "LOANSNOTREDEEMED"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vstrLoansNotRedeemed
    If xmlElement.Text = "" Then
        xmlElement.Text = "0"
    End If
    xmlTableElem.appendChild xmlElement
    
    'SharedPercentage
    strTag = "SHAREDPERCENTAGE"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vstrSharedPercentage
    If xmlElement.Text = "" Then
        xmlElement.Text = "0"
    End If
    xmlTableElem.appendChild xmlElement
    
    xmlTableElem.appendChild xmlAppStaffNode
    
    strResponse = objMortgageSubQuoteBO.ProcessAddedOneOffCost(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
'Stage4:

    ' Attach the required data to Response and Exit

    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlDoc.documentElement
    xmlResponseElem.appendChild xmlApplData.selectSingleNode(".//APPLICATIONDATE")
    
    Set ICostModelAssistBO_ProcessAddedOneOffCosts = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ProcessAddedOneOffCostsExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplData = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    
    Set objICostModelAssistDO = Nothing
    Set objComboDO = Nothing
    Set objGlobalParameterDO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    
    Exit Function
    
ProcessAddedOneOffCostsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_ProcessAddedOneOffCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostsExit

End Function

Private Function ICostModelAssistBO_RecostMortgageComponentAndLifeCover( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
'-------------------------------------------------------------------------------------------
'Description : This does all the processing required for the methods
'              'RecostMortgageComponentAndLifeCover' in ApplicationQuote and QuickQuote
'              NOTE : Re-costing of the associate Life cover is done in the respective TxBos
'              i.e., AppQuoteTxBo or QuickQuoteTxBo. This is done to avoid the references of
'              AppQuoteBO or QuickQuoteBO is this class.
'-------------------------------------------------------------------------------------------
On Error GoTo RecostMortgageComponentAndLifeCoverVbErr

    Const strFunctionName As String = "ICostModelAssistBO_RecostMortgageComponentAndLifeCover"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim xmlLoanComponent As IXMLDOMNode, xmlLoanComponents As IXMLDOMNode
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode, xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement ', xmlCalcEngineOutput As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strContext As String
    Dim strDrawDown As String
    'Dim strApplicationDate As String
    Dim strSearchPattern As String
    Dim lngInterestOnlyAmount As Long, lngCapitalAndInterestAmount As Long
    'Dim strPurchasePriceOrEstimatedValue As String, strTypeofApplication As String
    'Dim strLocation As String, strValuationType As String
    'Dim strLegalFeeType As String
    Dim strResponse As String
    
    Dim objMortgageSQBO As MortgageSubQuoteBO
    
    ' PSC 11/06/2002 BMIDS00043 - Start
'BMIDS774    Dim objROTBO As Object
'    Dim xmlROTRequest As IXMLDOMElement
'    Dim xmlROTRoot As IXMLDOMElement
'    Dim xmlCompDate As IXMLDOMAttribute
'BMIDS774    Dim xmlROTResponse As FreeThreadedDOMDocument40
    ' PSC 11/06/2002 BMIDS00043 - End
     
    ' BMIDS774
    Dim objIApplClassDef As IApplicationClassDef
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlAFFData As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlAFFElement As IXMLDOMElement
    Dim xmlCompDateElem As IXMLDOMElement
    Dim objIDOAssist As IDOAssist
    
    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
    Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    ' BMIDS774

     
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' PSC 11/07/2002 BMIDS00062 - Start
    Set xmlNewRequest = vxmlRequest.cloneNode(True)
    strContext = xmlNewRequest.selectSingleNode(".//CONTEXT").Text

        
    If UCase$(strContext) = "COST MODELLING" Then
        m_objXMLAssist.ChangeNodeName xmlNewRequest, "APPLICATIONQUOTE", "MORTGAGESUBQUOTE"
    Else
        m_objXMLAssist.ChangeNodeName xmlNewRequest, "QUICKQUOTE", "MORTGAGESUBQUOTE"
    End If
    ' PSC 11/07/2002 BMIDS00062 - End
    
    '----------------------------------------------------------------------------------------------
    ' Retrieve the various parameters from the passed in XML
    '----------------------------------------------------------------------------------------------
    ' PSC 11/07/2002 BMIDS00062
    'strSearchPattern = IIf(UCase(strContext) = "QUICK QUOTE", "QUICKQUOTE", "APPLICATIONQUOTE")
    lngInterestOnlyAmount = Val(xmlNewRequest.selectSingleNode(".//INTERESTONLYAMOUNT").Text)
    lngCapitalAndInterestAmount = Val(xmlNewRequest.selectSingleNode(".//CAPITALANDINTERESTAMOUNT").Text)
    'strApplicationDate = xmlNewRequest.selectSingleNode(".//APPLICATIONDATE").Text
    
    ' Find the LOANCOMPONENT in the list which corresponds to the individual LOANCOMPONENT passed in
    Set xmlLoanComponents = xmlNewRequest.selectSingleNode(".//LOANCOMPONENTLIST")
    Set xmlLoanComponentList = xmlLoanComponents.selectNodes(".//LOANCOMPONENT")
    
     ' PSC 11/06/2002 BMIDS00043 - Start
'    Set xmlLoanComponent = vxmlRequest.selectSingleNode(".//" & strSearchPattern & "/LOANCOMPONENT")
'    For Each xmlNode In xmlLoanComponentList
'        If m_objXmlAssist.NodesMatch(xmlNode, xmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER") Then
'            Set xmlLoanComponent = xmlNode
'        End If
'    Next xmlNode
      ' PSC 11/06/2002 BMIDS00043 - End
    '--------------------------------------------------------------------------------------------------------
    ' 1) Update Loan Component for New Part & Part Split
    '--------------------------------------------------------------------------------------------------------
    
    ' PSC 11/07/2002 BMIDS00062 - Start
    Dim intMinSeqNo As Integer
    Dim strSeqNo As String
    Dim strPorted  As String
    
    ' Find first non ported loan to add costs too
    intMinSeqNo = 9999
    
    For Each xmlNode In xmlLoanComponentList
        strSeqNo = m_objXMLAssist.GetMandatoryElementText(xmlNode, "LOANCOMPONENTSEQUENCENUMBER")
        strPorted = m_objXMLAssist.GetMandatoryElementText(xmlNode, "PORTEDLOAN")
        
        If strPorted <> "1" And CInt(strSeqNo) < intMinSeqNo Then
            intMinSeqNo = CInt(strSeqNo)
        End If
    Next

    strSearchPattern = "LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER='" & CStr(intMinSeqNo) & "']"
    Set xmlLoanComponent = xmlLoanComponents.selectSingleNode(strSearchPattern)
    ' PSC 11/07/2002 BMIDS00062 - End
    
    If lngInterestOnlyAmount <> 0 Then
    
        ' PSC 11/06/2002 BMIDS00043 - Start
'        vxmlRequest.selectSingleNode("./" & strSearchPattern & "/LOANCOMPONENT/INTERESTONLYELEMENT").Text = _
'                                                                CStr(lngInterestOnlyAmount)
'        vxmlRequest.selectSingleNode("./" & strSearchPattern & "/LOANCOMPONENT/CAPITALANDINTERESTELEMENT").Text = _
'                                                                CStr(lngCapitalAndInterestAmount)
        ' PSC 11/06/2002 BMIDS00043 - End
        
        m_objXMLAssist.SetMandatoryChildText xmlLoanComponent, _
                            "INTERESTONLYELEMENT", CStr(lngInterestOnlyAmount)
        m_objXMLAssist.SetMandatoryChildText xmlLoanComponent, _
                            "CAPITALANDINTERESTELEMENT", CStr(lngCapitalAndInterestAmount)
    End If
    
    '--------------------------------------------------------------------------------
    'Re-cost the Loan Component Cost
    '--------------------------------------------------------------------------------
    ' PSC 11/06/2002 BMIDS00043 - Start
'    If UCase(strContext) = "COST MODELLING" Then
'        Set xmlElement = SetForRecostAQLoan(vxmlRequest, strPurchasePriceOrEstimatedValue, _
'                                strTypeofApplication, strLocation, strValuationType, _
'                                strLegalFeeType)
'    Else
'        Set xmlElement = SetForRecostQQLoan(vxmlRequest, strPurchasePriceOrEstimatedValue, _
'                                strTypeofApplication, strLocation, strValuationType, _
'                                strLegalFeeType)
'    End If
'
'    m_objErrAssist.CheckResponse xmlElement, xmlResponseElem, True
'
'    ' Build the request for MortgageSubQuoteBO.RecostComponent
'    Set xmlNewRequest = vxmlRequest.cloneNode(True)
'
'    If UCase(strContext) = "COST MODELLING" Then
'        m_objXmlAssist.ChangeNodeName xmlNewRequest, "APPLICATIONQUOTE", "MORTGAGESUBQUOTE"
'    Else
'        m_objXmlAssist.ChangeNodeName xmlNewRequest, "QUICKQUOTE", "MORTGAGESUBQUOTE"
'    End If
'
'
'    Set xmlNode = xmlNewRequest.selectSingleNode("./MORTGAGESUBQUOTE")
'    Set xmlElement = xmlOut.createElement("PURCHASEPRICE")
'    xmlElement.Text = strPurchasePriceOrEstimatedValue
'    xmlNode.appendChild xmlElement
'
'    Set xmlElement = xmlOut.createElement("TYPEOFAPPLICATION")
'    xmlElement.Text = strTypeofApplication
'    xmlNode.appendChild xmlElement
'
'    Set xmlElement = xmlOut.createElement("LOCATION")
'    xmlElement.Text = strLocation
'    xmlNode.appendChild xmlElement
'
'    Set xmlElement = xmlOut.createElement("TYPEOFVALUATION")
'    xmlElement.Text = strValuationType
'    xmlNode.appendChild xmlElement
'
'    Set xmlElement = xmlOut.createElement("FEETYPE")
'    xmlElement.Text = strLegalFeeType
'    xmlNode.appendChild xmlElement

    ' PSC 11/06/2002 BMIDS00043 - End
    
    ' PSC 11/06/2002 BMIDS00043 - Start
    ' Call ROT to get completion date
'BMIDS774    Set xmlROTRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    xmlROTRequest.setAttribute "OPERATION", "GetReportOnTitleData"
'    Set xmlROTRoot = xmlOut.createElement("REPORTONTITLE")
'    xmlROTRequest.appendChild xmlROTRoot
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    
    '*=[MC]BMIDS763 - APPLICATIONDATE
    Dim strAppDate As String
    '*=SECTION END
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlLoanComponent, "APPLICATIONNUMBER")
    strAppFFNo = m_objXMLAssist.GetTagValue(xmlLoanComponent, "APPLICATIONFACTFINDNUMBER")
    
    '*=[MC]BMIDS763 - APPLICATIONDATE
    strAppDate = m_objXMLAssist.GetTagValue(vxmlRequest, "APPLICATIONDATE")
    '*=SECTION END
    
'BMIDS774    xmlROTRoot.setAttribute "APPLICATIONNUMBER", strAppNo
'    xmlROTRoot.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
'
'    Set objROTBO = m_objContext.CreateInstance(gstrREPORTONTITLE_COMPONENT & ".omROTBO")
'
'    strResponse = objROTBO.OmRotRequest(xmlROTRequest.xml)
'    Set xmlROTResponse = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'    Dim lngErrorNo As Long
'
'    lngErrorNo = m_objErrAssist.CheckResponse(xmlROTResponse.documentElement, xmlResponseElem, False)
'
'    ' If not success or record not found re-raise the error
'    If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
'        m_objErrAssist.ReRaiseResponseError xmlROTResponse.documentElement
'    End If
'
'BMIDS774    Set xmlCompDate = xmlROTResponse.selectSingleNode("RESPONSE/REPORTONTITLE/@COMPLETIONDATE")
    'BMIDS774 Need to get the ApplicationFactFind EstimatedCompletionDate
    Set xmlInputElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlAFFElement.Text = strAppNo
    xmlTableElem.appendChild xmlAFFElement
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlAFFElement.Text = strAppFFNo
    xmlTableElem.appendChild xmlAFFElement
    
    '*=[MC]BMIDS763 - APPLICATION DATE
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONDATE")
    xmlAFFElement.Text = strAppDate
    xmlTableElem.appendChild xmlAFFElement
    '*=SECTION END
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData()
    Set xmlAFFData = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
   
    Set xmlCompDateElem = xmlAFFData.selectSingleNode("ESTIMATEDCOMPLETIONDATE")
    
    ' PSC 11/07/2002 BMIDS00062 Add completion date to each component
    If Not xmlCompDateElem Is Nothing Then
        'BMIDS774 If EstimatedCompletionDate is >= Sys Date, use it
        'otherwise don't create it so we get the alpha+ default.
        If DateDiff("d", Now, CSafeDate(xmlCompDateElem.Text)) >= 0 Then
            For Each xmlNode In xmlLoanComponentList
                Set xmlElement = xmlOut.createElement("COMPLETIONDATE")
                xmlElement.Text = xmlCompDateElem.Text
                xmlNode.appendChild xmlElement
            Next
        End If
    End If
    ' PSC 11/06/2002 BMIDS00043 - End
    
    ' PSC 11/07/2002 BMIDS00062 - Start
    ' Move one off costs to the correct component for calculation
    Set xmlNode = m_objXMLAssist.GetMandatoryNode(xmlNewRequest, "MORTGAGESUBQUOTE/ONEOFFCOSTLIST")
    xmlLoanComponent.appendChild xmlNode
    ' PSC 11/07/2002 BMIDS00062 - End
    
    
    If m_objContext Is Nothing Then
        Set objMortgageSQBO = New MortgageSubQuoteBO
    Else
        Set objMortgageSQBO = _
            m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    'DPF 7/11/02 - BMIDS00808
    'If a Drawdown amount has been passed in we calculate our monthly costs less drawdown
    'BMIDS793 GHun Only calculate drawdown if the amount is > 0
    Set xmlNode = xmlNewRequest.selectSingleNode("//DRAWDOWN")
    If Not xmlNode Is Nothing Then
        If CSafeDbl(xmlNode.Text) > 0 Then
    'BMIDS793 End
            Dim xmlDrawDownDoc As FreeThreadedDOMDocument40
            Set xmlDrawDownDoc = New FreeThreadedDOMDocument40
            xmlDrawDownDoc.validateOnParse = False
            xmlDrawDownDoc.setProperty "NewParser", True
            strDrawDown = objMortgageSQBO.CalculateDrawDownCosts(xmlNewRequest.xml)
            Set xmlDrawDownDoc = m_objXMLAssist.load(strDrawDown, TypeName(Me), strFunctionName)
        'BMIDS793 GHun
        Else
            strDrawDown = vbNullString
        End If
        'BMIDS793 End
    Else
        strDrawDown = ""
    End If
     
    ' PSC 11/07/2002 BMIDS00062 - Start
    'DPF:  If we have calculated drawdown costs we pass in a slightly amended request
    If Len(strDrawDown) = 0 Then
        strResponse = objMortgageSQBO.RunCompositeRateCalc(xmlNewRequest.xml)
    Else
        Set xmlNewRequest = m_objXMLAssist.GetMandatoryNode(xmlDrawDownDoc.documentElement, "REQUEST")
        strResponse = objMortgageSQBO.RunCompositeRateCalc(xmlNewRequest.xml)
    End If

    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'If we have calculated a monthly cost less drawdown append it to response to GUI
    If Len(strDrawDown) > 0 Then
        Dim xmlMSQMCLessDrawdown As IXMLDOMNode
        Set xmlMSQMCLessDrawdown = m_objXMLAssist.GetNode(xmlDrawDownDoc, "//REQUEST/MONTHLYCOSTLESSDRAWDOWN", False)
        xmlDoc.documentElement.appendChild xmlMSQMCLessDrawdown
    End If
    
    'DPF END
    
'    Set xmlCalcEngineOutput = xmlDoc.documentElement
'
'    AttachCalcsEngineOutputToResponse xmlResponseElem, xmlCalcEngineOutput
        
    '-----------------------------------------------------------------------------------------
    ' 3) Update the appropriate Loan Component in xmlLoanComponents with Calcs Engine results
    '------------------------------------------------------------------------------------------
'    m_objXmlAssist.SetMandatoryChildText xmlLoanComponent, "APR", _
'        m_objXmlAssist.GetTagValue(xmlCalcEngineOutput, "OUT030_APR")
'    m_objXmlAssist.SetMandatoryChildText xmlLoanComponent, "NETMONTHLYCOST", _
'        m_objXmlAssist.GetTagValue(xmlCalcEngineOutput, "OUT060_TOTALNETMORTGAGEPAYMENT1")
'    m_objXmlAssist.SetMandatoryChildText xmlLoanComponent, "GROSSMONTHLYCOST", _
'        m_objXmlAssist.GetTagValue(xmlCalcEngineOutput, "OUT130_TOTALGROSSMORTGAGEPAYMENT1")
    
    '**NOTE**----------------------------------------------------------------------------
    ' Re-cost the associated Life Cover --  is done in AppQuoteTxBO or QuickQuoteTxBO
    '------------------------------------------------------------------------------------
    
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlDoc.documentElement
    ' PSC 11/07/2002 BMIDS00062 - End
    
    Set ICostModelAssistBO_RecostMortgageComponentAndLifeCover = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostMortgageComponentAndLifeCoverExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Set xmlLoanComponent = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    'Set xmlCalcEngineOutput = Nothing
    Set xmlDoc = Nothing
    Set xmlDrawDownDoc = Nothing
    Set objMortgageSQBO = Nothing
    
    ' PSC 11/06/2002 BMIDS00043 - Start
'BMIDS774    Set objROTBO = Nothing
'    Set xmlROTRequest = Nothing
'    Set xmlROTRoot = Nothing
'    Set xmlCompDate = Nothing
'BMIDS774    Set xmlROTResponse = Nothing
    ' PSC 11/06/2002 BMIDS00043 - End
    'BMIDS774
    Set xmlAFFData = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlAFFElement = Nothing
    Set xmlCompDateElem = Nothing
    Set objIApplClassDef = Nothing
    Set xmlApplClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    Exit Function
    
RecostMortgageComponentAndLifeCoverVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_RecostMortgageComponentAndLifeCover = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit

End Function


Private Function SetForRecostAQLoan(ByVal vxmlRequest As IXMLDOMNode, _
                            ByRef strPurchasePriceOrEstimatedValue As String, _
                            ByRef strTypeofApplication As String, _
                            ByRef strLocation As String, _
                            ByRef strValuationType As String, _
                            ByRef strLegalFeeType As String) As IXMLDOMElement
'--------------------------------------------------------------------------------------------
'Description : Sets the request for re-costing the loan component for ApplicationQuote
'              Called from - RecostMortgageComponentAndLifeCover
'--------------------------------------------------------------------------------------------
On Error GoTo RecostVbErr
    
    Const strFunctionName As String = "SetForRecostAQLoan"


    Dim objIApplicationDO As IApplicationDO, objNewPropertyBO As NewPropertyBO
    Dim objComboDO As ComboDO
    Dim strApplNumber As String, strAFFNumber As String
    Dim strTag As String, strResponse As String
   
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode, xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objNewPropertyBO = New NewPropertyBO
        Set objIApplicationDO = New ApplicationDO
        Set objComboDO = New ComboDO
    Else
        Set objIApplicationDO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objNewPropertyBO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
        
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    '--------------------------------------------------------------------------
    ' Find PurchasePriceOrEstiated value, TypeOfApplication
    '--------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = objIApplicationDO.GetApplicationData(xmlNewRequest)
    
    strPurchasePriceOrEstimatedValue = xmlElement.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE").Text
    strTypeofApplication = xmlElement.selectSingleNode(".//TYPEOFAPPLICATION").Text
    
    '----------------------------------------------------------------------------------
    'Set LegalFeeType
    '----------------------------------------------------------------------------------
    strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "R")
    
    For Each xmlElement In _
        vxmlRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/REPAYMENTMETHOD")
        
        If Not objComboDO.IsItemInValidation("RepaymentType", xmlElement.Text, "C") Then
            strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "D")
            Exit For
        End If
    Next xmlElement
    
    '-----------------------------------------------------------------------------------
    ' Find ValuationType, LegalFeeType of New Property
    '-----------------------------------------------------------------------------------
    m_objXMLAssist.ChangeNodeName xmlNewRequest, "APPLICATION", "NEWPROPERTY"
    strResponse = objNewPropertyBO.GetValuationTypeAndLocation(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    strValuationType = xmlDoc.selectSingleNode(".//VALUATIONTYPE").Text
    strLocation = xmlDoc.selectSingleNode(".//PROPERTYLOCATION").Text
    
    Set SetForRecostAQLoan = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostVbExit:
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    Set objIApplicationDO = Nothing
    Set objNewPropertyBO = Nothing
    Set objComboDO = Nothing
    
    Exit Function
RecostVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set SetForRecostAQLoan = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostVbExit

End Function

Private Function SetForRecostQQLoan(ByVal vxmlRequest As IXMLDOMNode, _
                            ByRef strPurchasePriceOrEstimatedValue As String, _
                            ByRef strTypeofApplication As String, _
                            ByRef strLocation As String, _
                            ByRef strValuationType As String, _
                            ByRef strLegalFeeType As String) As IXMLDOMElement
'--------------------------------------------------------------------------------------------
'Description : Sets the request for re-costing the loan component for ApplicationQuote
'              Called from - RecostMortgageComponentAndLifeCover
'--------------------------------------------------------------------------------------------
On Error GoTo RecostVbErr
    
    Const strFunctionName As String = "SetForRecostQQLoan"


    Dim objIApplicationDO As IApplicationDO
    Dim objComboDO As ComboDO, objGlobalParamDO As GlobalParameterDO
    Dim strApplNumber As String, strAFFNumber As String
    Dim strTag As String, strResponse As String
   
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode, xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objComboDO = New ComboDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objIApplicationDO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParamDO = _
                m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
        
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    '--------------------------------------------------------------------------
    ' Find PurchasePriceOrEstiated value, TypeOfApplication
    '--------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = objIApplicationDO.GetApplicationData(xmlNewRequest)
    
    strPurchasePriceOrEstimatedValue = xmlElement.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE").Text
    strTypeofApplication = xmlElement.selectSingleNode(".//TYPEOFAPPLICATION").Text
    
    '----------------------------------------------------------------------------------
    'Set LegalFeeType
    '----------------------------------------------------------------------------------
    strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "R")
    For Each xmlElement In _
        vxmlRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/REPAYMENTMETHOD")
        
        If Not objComboDO.IsItemInValidation("RepaymentType", xmlElement.Text, "C") Then
            strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "D")
            Exit For
        End If
    Next xmlElement
    
    '-----------------------------------------------------------------------------------
    ' Find ValuationType, Location of New Property
    '-----------------------------------------------------------------------------------
    Dim strTemp As String
    
    strTemp = objGlobalParamDO.GetCurrentParameterByType("QQLocation", "STRING")
    strLocation = objComboDO.GetFirstComboValueId("PropertyLocation", strTemp)
    
    strTemp = objGlobalParamDO.GetCurrentParameterByType("QQValuationType", "STRING")
    strValuationType = objComboDO.GetFirstComboValueId("ValuationType", strTemp)
    
    Set SetForRecostQQLoan = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostVbExit:
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    Set objIApplicationDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objComboDO = Nothing
    
    Exit Function
RecostVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set SetForRecostQQLoan = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostVbExit

End Function

Private Sub AttachCalcsEngineOutputToResponse( _
                                    ByVal xmlResponseElem As IXMLDOMElement, _
                                    ByVal xmlCalcEngineOutput As IXMLDOMElement)
'------------------------------------------------------------------------------
'Description : Attaches the output from Recosting (CalcEngineOutputInterface XML)
'              to Response XML.
'              Called from ICostModelAssistBO.RecostMortgageCostAndLifeCover
'------------------------------------------------------------------------------
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlElement = xmlOut.createElement("COMPONENTNETMONTHLYCOST")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT060_TOTALNETMORTGAGEPAYMENT1").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("COMPONENTGROSSMONTHLYCOST")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT130_TOTALGROSSMORTGAGEPAYMENT1").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT020_ACCRUEDINTEREST")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT020_ACCRUEDINTEREST").Text
    xmlResponseElem.appendChild xmlElement

    Set xmlElement = xmlOut.createElement("APR")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT030_APR").Text
    xmlResponseElem.appendChild xmlElement

    Set xmlElement = xmlOut.createElement("OUT040_TOTALAMOUNTPAYABLE")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT040_TOTALAMOUNTPAYABLE").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT050_TOTALMORTGAGEPAYMENTS")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT050_TOTALMORTGAGEPAYMENTS").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT070_TOTALNETMORTGAGEPAYMENT2")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT070_TOTALNETMORTGAGEPAYMENT2").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT080_TOTALNETMORTGAGEPAYMENT3")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT080_TOTALNETMORTGAGEPAYMENT3").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT090_TOTALNETMORTGAGEPAYMENT4")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT090_TOTALNETMORTGAGEPAYMENT4").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT100_TOTALNETMORTGAGEPAYMENT5")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT100_TOTALNETMORTGAGEPAYMENT5").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT110_TOTALNETMORTGAGEPAYMENT6")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT110_TOTALNETMORTGAGEPAYMENT6").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT120_TOTALNETMORTGAGEPAYMENT7")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT120_TOTALNETMORTGAGEPAYMENT7").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT140_TOTALGROSSMORTGAGEPAYMENT2")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT140_TOTALGROSSMORTGAGEPAYMENT2").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT150_TOTALGROSSMORTGAGEPAYMENT3")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT150_TOTALGROSSMORTGAGEPAYMENT3").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT160_TOTALGROSSMORTGAGEPAYMENT4")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT160_TOTALGROSSMORTGAGEPAYMENT4").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT170_TOTALGROSSMORTGAGEPAYMENT5")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT170_TOTALGROSSMORTGAGEPAYMENT5").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT180_TOTALGROSSMORTGAGEPAYMENT6")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT180_TOTALGROSSMORTGAGEPAYMENT6").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT190_TOTALGROSSMORTGAGEPAYMENT7")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT190_TOTALGROSSMORTGAGEPAYMENT7").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT200_UNROUNDEDAPR")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT200_UNROUNDEDAPR").Text
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("OUT210_FINALGROSSPAYMENT")
    xmlElement.Text = _
            xmlCalcEngineOutput.selectSingleNode(".//OUT210_FINALGROSSPAYMENT").Text
    xmlResponseElem.appendChild xmlElement
    
End Sub

Private Function ICostModelAssistBO_ReinstateQuotation( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal strContext = "QUICK QUOTE") As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Common method for AQ and QQ which Reinstates Quotation
' pass:         vxmlRequest
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateQuotationVbErr
    
    Const strFunctionName As String = "ICostModelAssistBO_ReinstateQuotation"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlDoc As FreeThreadedDOMDocument40 ' Document to load Response from other objects
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    'Rollback AQR 0977 Dim xmlResponseNode As IXMLDOMNode
    Dim xmlSelfEmpRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
        
    'MV - 18/07/2002 - BMIS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767
    'Dim xmlQuoteElem As IXMLDOMElement
    
    Dim xmlNewRequest As IXMLDOMNode, xmlCustomerElem As IXMLDOMNode
    
    Dim strQuotationNumber As String, strTag As String
    Dim blnIsMemerOfStaff As Boolean, blnIsSelfEmployed As Boolean
    
    Dim objApplMgrBO As ApplicationManagerBO
    Dim objCFBO As CustomerFinancialBO
    Dim objICEDO As ICustomerEmploymentDO
    'Rollback AQR 0977 Dim objCostModelAssistBO As ICostModelAssistBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'Rollback AQR 0977 Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objApplMgrBO = New ApplicationManagerBO
    Else
        Set objApplMgrBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    End If

    ' Get CustomersForApplication
    Dim strResponse As String
    strResponse = objApplMgrBO.FindCustomersForApplication(vxmlRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    'Rollback AQR 0977
''
'' build a customer list
''
'    Dim xmlCustomerList As IXMLDOMNodeList
'    Dim xmlCustomerNode As IXMLDOMNode
'
'    Set xmlCustomerList = xmlDoc.getElementsByTagName("CUSTOMER")
'
'    If xmlCustomerList.length = 0 Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
'            oeMissingPrimaryTag, "List of customers for application is empty"
'
'    End If
    'End Rollback AQR 0977
    
    If strContext = "COST MODELLING" Then  ' store the customerVersion details
        Set xmlCustomerElem = xmlDoc.documentElement.firstChild.firstChild
    End If
    
    'Check if any customer on this application is member of staff
    blnIsMemerOfStaff = IsMemberOfStaff(xmlDoc.documentElement)
    
    'Check if any customer on this application is self-employed. Build the required
    'REQUEST node first
    If strContext = "QUICK QUOTE" Then
        Set xmlSelfEmpRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableElem = xmlOut.createElement("QUICKQUOTEAPPLICANTDETAILS")
        xmlSelfEmpRequest.appendChild xmlTableElem
    
        strTag = "APPLICATIONNUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElement
    
        strTag = "APPLICATIONFACTFINDNUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
        xmlTableElem.appendChild xmlElement
    
        blnIsSelfEmployed = IsSelfEmployed(xmlSelfEmpRequest)
    Else
        If m_objContext Is Nothing Then
            Set objICEDO = New CustomerEmploymentDO
        Else
            Set objICEDO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentDO")

        End If
        'MS     14/07/00    SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
        'blnIsSelfEmployed = objICEDO.IsSelfEmployed(xmlDoc.documentElement)
        ' DM 11/08/00   SYS1173 Cant reinstate subquote
        ' We need to pass an EMPLOYMENT node but we have a CUSTOMER node.
        ' We will change it before the call and back after the call.
        Dim objXmlAssist As XMLAssist
        Set objXmlAssist = New XMLAssist
        objXmlAssist.ChangeNodeName xmlDoc, "CUSTOMER", "EMPLOYMENT"
        blnIsSelfEmployed = objICEDO.IsMainSelfEmployed(xmlDoc.documentElement)
        objXmlAssist.ChangeNodeName xmlDoc, "EMPLOYMENT", "CUSTOMER"

    End If
  
        
    '------------------------------------------------------------------------------------
    'In case of APPLICATIONQUOTE, retrieve existing mortgage loans that are to be ported
    'and, add then to the Request
    '-----------------------------------------------------------------------------------
    If strContext = "COST MODELLING" Then
        If m_objContext Is Nothing Then
            Set objCFBO = New CustomerFinancialBO
        Else
            Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        End If
        
        Set xmlNewRequest = vxmlRequest.cloneNode(True)
        xmlNewRequest.appendChild xmlCustomerElem
    End If
    'Rollback AQR 0977
'            ' DM AQR sys0977
'        Dim xmlTempNode As IXMLDOMNode
'        Dim lngError As Long
'        If m_objContext Is Nothing Then
'            Set objCostModelAssistBO = New CostModelAssistBO
'        Else
'            Set objCostModelAssistBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistBO")
'
'        End If
'        Set xmlTempNode = objCostModelAssistBO.GetLTV(vxmlRequest, "COST MODELLING")
'        ' Check the response
'        lngError = m_objErrAssist.CheckResponse(xmlTempNode, xmlResponseNode, True)
'
'        ' Create a temporary document
'        Dim xmlDocCalcLTV As FreeThreadedDOMDocument40
'        Dim xmlDocElement As IXMLDOMElement
'        Set xmlDocCalcLTV = New FreeThreadedDOMDocument40
'
'        'copy over the request node no children
'        Set xmlDocElement = objXmlAssist.GetRequestNodeEx(vxmlRequest)
'        Set xmlDocCalcLTV = xmlDocElement.ownerDocument
'
'        ' copy over the LTV response node above including children.
'        objXmlAssist.CopyNode "LTV", "LTV", xmlTempNode.firstChild, xmlDocCalcLTV.firstChild, True
'        ' copy over the Customerlist from the Request with children
'
'  '      objXmlAssist.CopyNode "CUSTOMERROLELIST", "CUSTOMERLIST", xmlDoc.firstChild, xmlDocCalcLTV.firstChild.firstChild, True
'
'
''
'    ' Create a customer list node
'    Set xmlDocElement = xmlDocCalcLTV.firstChild.firstChild
'    Set xmlDocElement = xmlDocElement.appendChild(xmlDocCalcLTV.createElement("CUSTOMERLIST"))
'
'    For Each xmlCustomerNode In xmlCustomerList
'        xmlDocElement.appendChild xmlCustomerNode
'    Next
'        ' Call the CalcCostModelLTV method on this object
'        Dim xmlLTV As FreeThreadedDOMDocument40
'        Set xmlLTV = New FreeThreadedDOMDocument40
'        Dim objApplicationQuoteBO As ApplicationQuoteBO
'        If m_objContext Is Nothing Then
'            Set objApplicationQuoteBO = New ApplicationQuoteBO
'        Else
'            Set objApplicationQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
'        End If
'        strResponse = objApplicationQuoteBO.CalcCostModelLTV(xmlDocCalcLTV.xml)
'        ' Load the string into xml
'        Set xmlLTV = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'        ' Check the response
'        lngError = m_objErrAssist.CheckResponse(xmlLTV.documentElement, xmlResponseNode, True)
'        ' end of DM AQR sys0977
'      ' DC AQR sys0977
'        ' add on the LTV to the request
'        Dim xmlNode As IXMLDOMNode
'        ' DM Need to add the LTV xml to the xml being sent to the GetValidatedQuotationDetails method below.
'        Set xmlNode = xmlLTV.selectSingleNode("//LTV")
'
    'End Rollback AQR 0977
    'Check validity of quotation; reinstate if valid
    If strContext = "COST MODELLING" Then
    'Rollback AQR 0977         objXmlAssist.CopyNode "LTV", "LTV", xmlNode, xmlNewRequest, True
        Call ValidateAndReinstateQuotation(xmlNewRequest, blnIsMemerOfStaff, blnIsSelfEmployed)
    Else
    'Rollback AQR 0977         objXmlAssist.CopyNode "LTV", "LTV", xmlNode, vxmlRequest, True
        
        'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
        'SG 28/05/02 SYS4767 START
        'strTag = "ISQUICKQUOTE"  'JLD MSMS0092
        'Set xmlQuoteElem = vxmlRequest.selectSingleNode(".//QUOTATION")
        'Set xmlElement = xmlOut.createElement(strTag)
        'xmlElement.Text = "1"
        'xmlQuoteElem.appendChild xmlElement
        'SG 28/05/02 SYS4767 END
       
        Call ValidateAndReinstateQuotation(vxmlRequest, blnIsMemerOfStaff, blnIsSelfEmployed)
    End If
    'Rollback AQR 0977           'End DC AQR sys0977
    Set ICostModelAssistBO_ReinstateQuotation = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReinstateQuotationExit:

    '   clean-up section, release object references
    Set objApplMgrBO = Nothing
    Set objCFBO = Nothing
    Set xmlSelfEmpRequest = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    Exit Function

ReinstateQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_ReinstateQuotation = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateQuotationExit
End Function

Private Sub ValidateAndReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement, _
                                          ByVal blnIsMemerOfStaff As Boolean, _
                                          ByVal blnIsSelfEmployed As Boolean)
'-------------------------------------------------------------------------------
'Description : This checks whether the quotation is valid; if so, updates the
'              'ActiveQuoteNumber' in ApplicationFactFind table
'-------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateAndReinstateQuotation"
    
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40 ' Document to load the response
    Dim strTag As String, strResponse As String
    Dim strApplNumber As String, strAFFNumber As String, strQuoteNumber As String
    
    Dim xmlApplNumberElem As IXMLDOMElement, xmlAFFNumberElem As IXMLDOMElement
    
    Dim objQuotationBO As QuotationBO
    Dim objApplicationBO As ApplicationBO
    
    ' Build the Request
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("QUOTATION")
    xmlNewRequest.appendChild xmlTableElem
    
    strTag = "APPLICATIONNUMBER"
    Set xmlApplNumberElem = xmlOut.createElement(strTag)
    strApplNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlApplNumberElem.Text = strApplNumber
    xmlTableElem.appendChild xmlApplNumberElem
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlAFFNumberElem = xmlOut.createElement(strTag)
    strAFFNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlAFFNumberElem.Text = strAFFNumber
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strTag = "QUOTATIONNUMBER"
    Set xmlElement = xmlOut.createElement("ACTIVEQUOTENUMBER")
    strQuoteNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlElement.Text = strQuoteNumber
    xmlTableElem.appendChild xmlElement
    
    strTag = "ISMEMBEROFSTAFF"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = IIf(blnIsMemerOfStaff, "1", "0")
    xmlTableElem.appendChild xmlElement
    
    strTag = "ISSELFEMPLOYED"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = IIf(blnIsSelfEmployed, "1", "0")
    xmlTableElem.appendChild xmlElement
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767 START
    'JLD MSMS0092 if ISQUICKQUOTE tag is present copy that one too.
    'If Not vxmlRequest.selectSingleNode(".//ISQUICKQUOTE") Is Nothing Then
    '    Set xmlElement = xmlOut.createElement("ISQUICKQUOTE")
    '    xmlElement.Text = vxmlRequest.selectSingleNode(".//ISQUICKQUOTE").Text
    '    xmlTableElem.appendChild xmlElement
    'End If
    'SG 28/05/02 SYS4767 END
    
    If m_objContext Is Nothing Then
        Set objQuotationBO = New QuotationBO
    Else
        Set objQuotationBO = _
                m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    
    strResponse = objQuotationBO.ValidateQuotationForReinstatement(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    
'    If m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDMORTGAGESUBQUOTE") = "1" Or _
'       m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDBCSUBQUOTE") = "1" Or _
'       m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDPPSUBQUOTE") = "1" Then
    'SYS0724 MCS 26/05/00
    If StrComp(m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDMORTGAGESUBQUOTE"), "0", vbTextCompare) = 0 Or _
       StrComp(m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDBCSUBQUOTE"), "0", vbTextCompare) = 0 Or _
       StrComp(m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "VALIDPPSUBQUOTE"), "0", vbTextCompare) = 0 Then

        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 216
    Else
        
        'MAR1061 Get the mortgagesubquote for the new quote
        m_objXMLAssist.ChangeNodeName xmlNewRequest, "ACTIVEQUOTENUMBER", "QUOTATIONNUMBER"
        strResponse = objQuotationBO.GetQuotationSummary(xmlNewRequest.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        Dim sPurchasePrice As String
        sPurchasePrice = m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "PURCHASEPRICEORESTIMATEDVALUE")
        
       ' Update the table 'ApplicationFactFind'
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableElem = xmlOut.createElement("APPLICATIONFACTFIND")
        xmlNewRequest.appendChild xmlTableElem
                
        xmlTableElem.appendChild xmlApplNumberElem
        xmlTableElem.appendChild xmlAFFNumberElem

        strTag = "ACTIVEQUOTENUMBER"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = strQuoteNumber
        xmlTableElem.appendChild xmlElement
        
        'MAR1061 update purchaseprice
        strTag = "PURCHASEPRICEORESTIMATEDVALUE"
        Set xmlElement = xmlOut.createElement(strTag)
        xmlElement.Text = sPurchasePrice
        xmlTableElem.appendChild xmlElement
                
        If m_objContext Is Nothing Then
            Set objApplicationBO = New ApplicationBO
        Else
            Set objApplicationBO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        End If
        
        strResponse = objApplicationBO.UpdateApplicationFactFind(xmlNewRequest.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True
    'Rollback AQR 0977         'DC AQR SYS0977
'        'reset recalculated LTV
'             If m_objXMLAssist.GetNodeValue(vxmlRequest, ".//LTV", False) <> "" Then
'                Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
'                Dim lngError As Long
'                Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
'                Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTEUPDATE")
'                xmlNewRequest.appendChild xmlTableElem
'                xmlTableElem.appendChild xmlApplNumberElem
'                xmlTableElem.appendChild xmlAFFNumberElem
'                m_objXMLAssist.CopyNode "MORTGAGESUBQUOTE", "MORTGAGESUBQUOTE", vxmlRequest, xmlTableElem, True
'                m_objXMLAssist.CopyNode "LTV", "LTV", vxmlRequest, xmlTableElem, True
'                strResponse = objMortgageSubQuoteBO.Update(xmlDoc.xml)
'                Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'        ' Check the response
'                lngError = m_objErrAssist.CheckXMLResponse(xmlDoc.firstChild, False)
'        ' if the error wasn't Record Not Found then raise it
'                If lngError <> 0 Then
'                    lngError = m_objErrAssist.GetOmigaErrorNumber(lngError)
'                    If lngError <> oeRecordNotFound Then
'                       m_objErrAssist.ThrowError gstrAPPLICATIONQUOTE, strFunctionName, lngError
'                    End If
'                End If
'
'
'            End If
' end Rollback AQR 0977
        
    End If

    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing

    Set xmlApplNumberElem = Nothing
    Set xmlAFFNumberElem = Nothing
    
    Set objQuotationBO = Nothing
    Set objApplicationBO = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub
ValidateVbErr:
    
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing

    Set xmlApplNumberElem = Nothing
    Set xmlAFFNumberElem = Nothing
    
    Set objQuotationBO = Nothing
    Set objApplicationBO = Nothing
    
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
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICostModelAssistBO_UpdateMortgageSubquote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates Mortgage Life Sub quote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGESUBQUOTE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'           :
'   </MORTGAGESUBQUOTE>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "ICostModelAssistBO_UpdateMortgageSubquote"


    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objMSQBO As MortgageSubQuoteBO
    Dim objConvertAssist As ConvertAssist
    Set objConvertAssist = New ConvertAssist
    Dim strAmtRequested As String
    Dim lngDeposit As Long
    
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objMSQBO = New MortgageSubQuoteBO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objMSQBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    '-------------------------------------------------------------------------------
    ' Recalculate Deposit           15/05/00 JLD SYS0715
    '-------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIApplicationDO.GetApplicationData(vxmlRequest)
    Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE")
    lngDeposit = 0
    If Not xmlTempNode Is Nothing Then
        strAmtRequested = vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").Text
        If objConvertAssist.CSafeLng(xmlTempNode.Text) > objConvertAssist.CSafeLng(strAmtRequested) Then
            lngDeposit = objConvertAssist.CSafeLng(xmlTempNode.Text) - objConvertAssist.CSafeLng(strAmtRequested)
        End If
    End If
            
    '------------------------------------------------------------------------------
    ' update the mortgage sub-quote
    '-------------------------------------------------------------------------------
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlElement = vxmlRequest.selectSingleNode(".//DEPOSIT")
    If xmlElement Is Nothing Then
        'create one if the deposit is not null
        If lngDeposit > 0 Then
            Set xmlTempNode = vxmlRequest.selectSingleNode("MORTGAGESUBQUOTE")
            Set xmlTempNode = xmlTempNode.appendChild(xmlDoc.createElement("DEPOSIT"))
            xmlTempNode.Text = lngDeposit
        End If
    Else
        xmlElement.Text = lngDeposit
    End If
    
    Dim strResponse As String
    strResponse = objMSQBO.Update(vxmlRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set ICostModelAssistBO_UpdateMortgageSubquote = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlElement = Nothing
    Set objIApplicationDO = Nothing
    Set objMSQBO = Nothing
    Set objConvertAssist = Nothing
    
    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_UpdateMortgageSubquote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
End Function

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    If m_objErrAssist Is Nothing Then
        Set m_objErrAssist = New ErrAssist
    End If

    ' DJP HERE
    Set m_objGeneralAssist = New GeneralAssist
    If m_objXMLAssist Is Nothing Then
        Set m_objXMLAssist = New XMLAssist
    End If

    If m_objConvertAssist Is Nothing Then
        Set m_objConvertAssist = New ConvertAssist
    End If

    Set m_objContext = GetObjectContext()
End Sub

Private Sub Class_Terminate()
    'SYS4510  Modifications to class terminate
    'Set m_objErrAssist = Nothing
    'Set m_objXMLAssist = Nothing
    'Set m_objContext = Nothing
    'END SYS4510 CL
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing
    Set m_objContext = Nothing
End Sub

Private Sub Class_Initialize()
    
    'SYS4510 CL
    'If m_objErrAssist Is Nothing Then
    '    Set m_objErrAssist = New ErrAssist
    'End If

    'If m_objXMLAssist Is Nothing Then
    '    Set m_objXMLAssist = New XMLAssist
    'End If

    'If m_objConvertAssist Is Nothing Then
    '    Set m_objConvertAssist = New ConvertAssist
    'End If
    'END SYS4510 CL
    
End Sub

Private Function ICostModelAssistBO_GetBorrowingAndLenderName(ByVal vstrXMLRequest As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Shared method between omQQ and omAppQuote
'
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------
On Error GoTo GetBorrowingAndLenderNameVbErr
    
    Const strFunctionName As String = "ICostModelAssistBO_GetBorrowingAndLenderName"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlTemp As New FreeThreadedDOMDocument40
    xmlTemp.validateOnParse = False
    xmlTemp.setProperty "NewParser", True

    Dim objMaximumBorrowingBO As MaximumBorrowingBO
    Dim objMortgageLenderBO As MortgageLenderBO

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                                            
    Set xmlResponseElem = xmlTemp.createElement("MAXBORROWINGLIST")
    Set xmlDataNode = xmlTemp.appendChild(xmlResponseElem)
         
    If objContext Is Nothing Then
        Set objMaximumBorrowingBO = New MaximumBorrowingBO
        Set objMortgageLenderBO = New MortgageLenderBO
    Else
        Set objMaximumBorrowingBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MaximumBorrowingBO")
        Set objMortgageLenderBO = objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
    End If
                                    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim xmlElement As IXMLDOMElement
    
    Set xmlTemp = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Dim strLenderNameAndMaxBorrowing As String

    Dim xmlNode As IXMLDOMNode

    Dim xmlCustANode As IXMLDOMNode
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlElementA As IXMLDOMElement
    'Get customer A element
    Set xmlElementA = xmlTemp.getElementsByTagName("CUSTOMERA").Item(0)

    If xmlElementA Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERA tag not found"
    End If
    ' income summary for cust a
    Set xmlNode = xmlElementA.getElementsByTagName("INCOMESUMMARY").Item(0)

    Dim lngEarnedIncomeA As Long
    Dim lngEarnedIncomeB As Long

    lngEarnedIncomeA = objXmlAssist.GetTagValue(xmlElementA, "TOTALGROSSEARNEDINCOME")

    Dim xmlTemp2 As New FreeThreadedDOMDocument40
    xmlTemp2.validateOnParse = False
    xmlTemp2.setProperty "NewParser", True
    ' create earnedincome node for cust a
    Set xmlCustANode = xmlTemp2.createElement("EARNEDINCOME")
    xmlCustANode.Text = lngEarnedIncomeA

    'raise error if lngEarnedIncomeA = 0 ???
    'get cust B element
    Set xmlElement = xmlTemp.getElementsByTagName("CUSTOMERB").Item(0)

    Dim strMortageLenders As String
    Dim objLenderDoc As New FreeThreadedDOMDocument40
    objLenderDoc.validateOnParse = False
    objLenderDoc.setProperty "NewParser", True
    Dim objNodeList As IXMLDOMNodeList
    Dim xmlElement2 As IXMLDOMElement
    'Get all the lender names and income multiples for each lender
    strMortageLenders = objMortgageLenderBO.FindLenderNameAndIncomeMultiples("<MORTGAGELENDER></MORTGAGELENDER>")
    Set objLenderDoc = objXmlAssist.load(strMortageLenders, TypeName(Me), strFunctionName)
    
    objErrAssist.CheckResponse objLenderDoc.documentElement, xmlResponseElem, True
            
    'get the list of lenders
    Set xmlElement2 = objLenderDoc.getElementsByTagName("MORTGAGELENDERLIST").Item(0)

    If xmlElement2 Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGELENDERLIST tag not found"
    End If
    'now list of lenders
    Set objNodeList = xmlElement2.getElementsByTagName("MORTGAGELENDER")

    Dim nloop As Integer
    Dim xmlElement3 As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode

    Dim xmlMaxBorrowingDoc As New FreeThreadedDOMDocument40
    xmlMaxBorrowingDoc.validateOnParse = False
    xmlMaxBorrowingDoc.setProperty "NewParser", True
    
    ' loop round all the lenders
    
    For nloop = 0 To (objNodeList.length - 1)
        'create the xml to pass down to calculate the maximum borrowing
        Set xmlRootNode = xmlMaxBorrowingDoc.createElement("CALCULATEMAXIMUMBORROWING")
        Set xmlNode = xmlMaxBorrowingDoc.appendChild(xmlRootNode)
        'append the cust a node
        Set xmlTempNode = xmlRootNode.appendChild(xmlElementA.cloneNode(False))
        'and the created node  EARNEDINCOME for cust a
        Set xmlNode = xmlTempNode.appendChild(xmlCustANode)

        Set xmlElement2 = objNodeList.Item(nloop)
        'if we have cust B
        If Not xmlElement Is Nothing Then
            'append cust B
            Set xmlTempNode = xmlElement.cloneNode(False)
            
            lngEarnedIncomeB = objXmlAssist.GetTagValue(xmlElement, "TOTALGROSSEARNEDINCOME")
            'create EARNEDINCOME node for Cust B
            Set xmlNode = xmlTemp2.createElement("EARNEDINCOME")
            xmlNode.Text = lngEarnedIncomeB
            'append the TOTALGROSSEARNEDINCOME for cust B
            Set xmlNode = xmlTempNode.appendChild(xmlNode)

            Set xmlNode = xmlRootNode.appendChild(xmlTempNode)
            'get the multipleincome multiplier out here
            Set xmlElement3 = xmlElement2.getElementsByTagName("HIGHESTEARNINCMULTIPLE").Item(0)
            Set xmlTempNode = xmlRootNode.appendChild(xmlElement3.cloneNode(True))

            Set xmlElement3 = xmlElement2.getElementsByTagName("SECONDEARNINCMULTIPLE").Item(0)
            Set xmlTempNode = xmlRootNode.appendChild(xmlElement3.cloneNode(True))

            Set xmlElement3 = xmlElement2.getElementsByTagName("JOINTINCOMEMULTIPLE").Item(0)
            Set xmlTempNode = xmlRootNode.appendChild(xmlElement3.cloneNode(True))

        Else
            'and the single income multiplier out here
            Set xmlElement3 = xmlElement2.getElementsByTagName("SINGLEINCOMEMULTIPLE").Item(0)

            'add this CUSTOMERA node
            Set xmlTempNode = xmlRootNode.appendChild(xmlElement3.cloneNode(True))

        End If

    Dim lngMaxBorrowing As Long

    lngMaxBorrowing = objMaximumBorrowingBO.CalculateMaximumBorrowing(xmlMaxBorrowingDoc.xml)   'This will return the MaxBorrowing Amount

     If lngMaxBorrowing > 0 Then

        Set xmlRootNode = xmlTemp2.createElement("MAXBORROWING")
        Set xmlTempNode = xmlDataNode.appendChild(xmlRootNode)

        Set xmlTemp2 = Nothing

        Set xmlTempNode = xmlTemp2.createElement("MAXBORROWINGAMOUNT")
        xmlTempNode.Text = lngMaxBorrowing
        Set xmlTempNode = xmlRootNode.appendChild(xmlTempNode)

        Set xmlElement3 = xmlElement2.getElementsByTagName("LENDERNAME").Item(0)
        Set xmlTempNode = xmlRootNode.appendChild(xmlElement3.cloneNode(True))

     End If

     Set xmlMaxBorrowingDoc = Nothing
    Next

    Set ICostModelAssistBO_GetBorrowingAndLenderName = xmlOut.documentElement
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetBorrowingAndLenderNameExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objMaximumBorrowingBO = Nothing
    Set objMortgageLenderBO = Nothing
    Set xmlTemp = Nothing
    Set xmlMaxBorrowingDoc = Nothing
    Set objLenderDoc = Nothing
                                
    Exit Function

GetBorrowingAndLenderNameVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_GetBorrowingAndLenderName = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetBorrowingAndLenderNameExit

End Function

Private Function ICostModelAssistBO_FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vxmlRequest  xml Request node
' <REQUEST>
'   <MORTGAGEPRODUCTREQUEST>
'        <SEARCHCONTEXT>TEST01</SEARCHCONTEXT>
'        <CHANNELID>1</CHANNELID>
'        <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
'        <APPLICATIONNUMBER></APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER></MORTGAGESUBQUOTENUMBER>
'        <PURPOSEOFLOAN></PURPOSEOFLOAN>
'        <TERMINYEARS></TERMINYEARS>
'        <TERMINMONTHS></TERMINMONTHS>
'        <AMOUNTREQUESTED></AMOUNTREQUESTED>
'   </MORTGAGEPRODUCTREQUEST>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "ICostModelAssistBO_FindMortgageProducts"
    Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDOMRequest As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim objAppManagerBO As ApplicationManagerBO
    Dim objCostModelAssistDO As ICostModelAssistDO
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim objApplicationDO As IApplicationDO
    Dim objIMortgageProductDO As IMortgageProductDO
    Dim objIMortgageProductBO As IMortgageProductBO
    Dim objComboDO As ComboDO
    Dim sQuotationType As String
    Dim objDOAssist As DOAssist             '++ BMIDS00627
    Dim strDOB() As String                  'MAR1702
    Dim intIndex As Integer                 'MAR1702
    
    On Error GoTo FindMortgageProductsErr

    strNodeName = "MORTGAGEPRODUCTREQUEST"

    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
        Set objAppManagerBO = New ApplicationManagerBO
        Set objCostModelAssistDO = New CostModelAssistDO
        Set objIMortgageProductDO = New MortgageProductDO
        Set objIMortgageProductBO = New MortgageProductBO
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
        Set objDOAssist = New DOAssist          '++SA BMIDS00627
        
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objAppManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
        ' omQQ or omAQ
        Set objCostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objICustomerEmploymentDO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentDO")
        Set objIMortgageProductDO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductDO")
        Set objIMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
        Set objDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")           '++SA BMIDS00627
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    Set xmlDOMRequest = vxmlRequest.ownerDocument
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    ' Validate and get what's passed in.
    Dim strAppNo As String
    Dim strAppFactFind As String
    Dim strQuoteNo As String
    
    'MAR696 If we are coming from the web, application number may not be present.
    'Make the following non-mandatory
    strAppNo = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONNUMBER", False)                 ' MAR696
    strAppFactFind = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER", False)   ' MAR696
    strQuoteNo = m_objXMLAssist.GetNodeValue(xmlRequestElem, "MORTGAGESUBQUOTENUMBER", False)          ' MAR696

    Dim bMemberOfStaff As Boolean
    'Dim bSelfEmployed As Boolean       '--BMIDS00246
    Dim intMemberOfStaff As Integer
    'Dim intSelfEmployed As Integer     '--BMIDS00246
    Dim strSearchContext As String      '++BMIDS00627
    
    bMemberOfStaff = False
    'bSelfEmployed = False      '--BMIDS00246
    intMemberOfStaff = 0
    'intSelfEmployed = 0        '--BMIDS00246
    
    Dim xmlQuoteNodeResp As IXMLDOMElement
    Dim xmlQuoteNodeRoot As IXMLDOMNode
    Dim xmlQuoteNodeReq As IXMLDOMNode
    Dim xmlQuoteNodeItem As IXMLDOMNode
    Dim xmlQQDetailsRequest As IXMLDOMElement       '++BMIDS00627
    Dim xmlTableElem As IXMLDOMElement              '++BMIDS00627
    Dim xmlOutDoc As FreeThreadedDOMDocument40        '++BMIDS00627
    
    Set xmlQuoteNodeRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlQuoteNodeReq = xmlQuoteNodeRoot.appendChild(xmlDOMRequest.createElement("BASICQUOTATIONDETAILS"))
    Set xmlQuoteNodeItem = xmlQuoteNodeReq.appendChild(xmlDOMRequest.createElement("APPLICATIONNUMBER"))
    
    xmlQuoteNodeItem.Text = strAppNo
    
    'MAR696 allow for no application number
    If (Len(strAppNo) > 0) Then
    
    Set xmlQuoteNodeResp = objCostModelAssistDO.GetDataForSubQuoteValidation(xmlQuoteNodeRoot)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Member of Staff
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    bMemberOfStaff = IsMemberOfStaff(xmlQuoteNodeResp)
    
    If bMemberOfStaff = True Then
        intMemberOfStaff = 1
    End If
    '--BMIDS00246 IsSelfEmployed no longer used in finding mortgage products {
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Self Employed
'    Dim xmlSelfEmployedRoot As IXMLDOMNode
'    Dim xmlSelfEmployedElem As IXMLDOMNode
'
'    ' SYS1031 Do we have self employed already passed in?
'    Set xmlSelfEmployedElem = m_objXmlAssist.GetNode(vxmlRequest, ".//APPLICANTSELFEMPLOYED", False)
'
'    If xmlSelfEmployedElem Is Nothing Then
'        Set xmlSelfEmployedRoot = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'        xmlSelfEmployedRoot.appendChild xmlQuoteNodeResp
'
'        'MS     14/07/00    SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
'        bSelfEmployed = objICustomerEmploymentDO.IsMainSelfEmployed(xmlSelfEmployedRoot)
'    Else
'        Dim sSelfEmployed As String
'        sSelfEmployed = xmlSelfEmployedElem.Text
'
'        If Len(sSelfEmployed) > 0 Then
'            bSelfEmployed = CBool(sSelfEmployed)
'        End If
'
'    End If
'
'    If bSelfEmployed = True Then
'        intSelfEmployed = 1
'    End If
'--BMIDS00246 }

    '++ BMIDS00246 SA Get a comma separated list of employment status for the
    ' main employment of each applicant to pass into stored procedure. {
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Employment Status Types
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strEmploymentStatusList As String
    Dim xmlEmploymentStatusRoot As IXMLDOMNode
                
    Const cstrEmploymentDefaultValidation As String = "E"   ' MAR18
    
    '++BMIDS00627 {
    strSearchContext = m_objXMLAssist.GetTagValue(vxmlRequest, "SEARCHCONTEXT")     '++BMIDS00627
    If UCase$(strSearchContext) = "QUICK QUOTE" Then
        'If in Quick Quote mode, we need to get the information from the QuickQuoteApplicantDetails
        Dim objClassDef As FreeThreadedDOMDocument40
        Dim xmlElement As IXMLDOMElement
        Set objClassDef = LoadDataForQuickQuoteApplicantDetails
        Set xmlOutDoc = New FreeThreadedDOMDocument40
        xmlOutDoc.validateOnParse = False
        xmlOutDoc.setProperty "NewParser", True
        
        'Build up request
        Set xmlQQDetailsRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableElem = xmlOutDoc.createElement("QUICKQUOTEAPPLICANTDETAILS")
        xmlQQDetailsRequest.appendChild xmlTableElem
    
        Set xmlElement = xmlOutDoc.createElement("APPLICATIONNUMBER")
        xmlElement.Text = strAppNo
        xmlTableElem.appendChild xmlElement
    
        Set xmlElement = xmlOutDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlElement.Text = strAppFactFind
        xmlTableElem.appendChild xmlElement
        
        Set xmlElement = objDOAssist.FindListMultipleEx(xmlQQDetailsRequest, objClassDef)
        'DO NOT error on no record found
        Dim nloop As Integer
        Dim xmlNodeList As IXMLDOMNodeList
        If Not xmlElement Is Nothing Then
            'get the number of employment records out
            Set xmlNodeList = xmlElement.getElementsByTagName("QUICKQUOTEAPPLICANTDETAILS")
            For nloop = 0 To (xmlNodeList.length - 1)
                '--BMIDS00662 SA allow for single applicants with multiple employment history.
                'If nloop = 0 Then
                If Len(strEmploymentStatusList) = 0 Then                                '++BMIDS00662
                    strEmploymentStatusList = m_objXMLAssist.GetTagValue(xmlNodeList(nloop), "EMPLOYMENTSTATUS")
                Else
                    strEmploymentStatusList = strEmploymentStatusList & "," & m_objXMLAssist.GetTagValue(xmlNodeList(nloop), "EMPLOYMENTSTATUS")
                End If
            Next
        End If
    Else
        '++BMIDS00627 }
        Set xmlEmploymentStatusRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        xmlEmploymentStatusRoot.appendChild xmlQuoteNodeResp
        strEmploymentStatusList = objICustomerEmploymentDO.GetEmploymentStatusList(xmlEmploymentStatusRoot)
        '++ BMIDS00246 }
        
        'MAR18
        If Len(strEmploymentStatusList) = 0 Then
            ' No employment record found (eg reached here from stand alone KFI)
            ' Convert the default validation type of "E" (Employed) to a value
            strEmploymentStatusList = objComboDO.GetFirstComboValueId("EmploymentStatus", cstrEmploymentDefaultValidation)
            
        End If
        
    End If  '++BMIDS00627
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get application data
    Dim xmlAppDataResp As IXMLDOMNode
    Set xmlAppDataResp = objCostModelAssistDO.GetApplicationData(vxmlRequest)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Establish Lender
    ' In a multi-loan component mortgage sub-quote, all components must have products from
    ' the same lender, so read the first one.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlLenderResp As IXMLDOMNode
    Dim strLender As String
    
    strLender = EstablishLender(vxmlRequest, _
                                xmlResponseRoot, _
                                objIMortgageProductDO, _
                                strAppNo, _
                                strAppFactFind, _
                                strQuoteNo)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Comment 2
    '' MS 11/07/00 SYS1198. For Phase 2. See also comment 1.
    ''
    '' We need to check the XML returned here. We need to do a search
    '' for the default non-panel lender product, so probably best to modify
    '' FindProductList method called below to Return a Non-Panel Product.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
                                
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Establish Default Country of Origin and Product Language
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strProductLanguage As String
    Dim strCountryOfOrigin As String
    
    EstablishCountryAndLanguage objComboDO, strProductLanguage, strCountryOfOrigin
    
    'MAR1702
    'Get applicants date of birth
    GetDOB strAppNo, strAppFactFind, strDOB
    
    End If   'MAR696
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Find the mortgage product
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlProdListResp As IXMLDOMNode
    Dim xmlProdListRoot As IXMLDOMNode
    Dim xmlProdListElem As IXMLDOMNode
        
    Set xmlProdListRoot = xmlRequestElem.cloneNode(True)
    m_objXMLAssist.ChangeNodeName xmlProdListRoot, strNodeName, "MORTGAGEPRODUCT"
    
    'MAR696  Allow for no application number
    If (Len(strAppNo) > 0) Then
    
    ' Member of staff
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("MEMBEROFSTAFF"))
    xmlProdListElem.Text = intMemberOfStaff
    '--BMIDS00246 {
    ' Is self employed
    'Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("SELFEMPLOYED"))
    'xmlProdListElem.Text = intSelfEmployed
    '-- BMIDS00246 }
    ' Type of application
    Dim strTypeofApplication As String
    
    '++ BMIDS00246 SA Add in list of employmentstatus {
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("EMPLOYMENTSTATUSLIST"))
    xmlProdListElem.Text = strEmploymentStatusList
    '++ BMIDS00246 }

    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    strTypeofApplication = m_objXMLAssist.GetNodeValue(xmlAppDataResp, "//TYPEOFAPPLICATION", False)
    'SYS 28/05/02 SYS4767 START
    ' JLD MSMS0034 Use the values passed in if they are there.
    'If Not xmlRequestElem.selectSingleNode(".//TYPEOFAPPLICATION") Is Nothing Then
    '    strTypeofApplication = m_objXMLAssist.GetNodeValue(xmlRequestElem, "//TYPEOFAPPLICATION", False)
    'Else
    '    strTypeofApplication = m_objXMLAssist.GetNodeValue(xmlAppDataResp, "//TYPEOFAPPLICATION", False)
    'End If
    'SG 28/05/02 SYS4767 END
    
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("TYPEOFAPPLICATION"))
    xmlProdListElem.Text = strTypeofApplication
    
    ' Type of buyer
    ' PSC 08/06/00 Correct strTypeOfApplication to strTypeOfBuyer
    Dim strTypeOfBuyer As String
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    strTypeOfBuyer = m_objXMLAssist.GetNodeValue(xmlAppDataResp, "//TYPEOFBUYER", False)
    'SG 28/05/02 SYS4767 START
    ' JLD MSMS0034 Use the values passed in if they are there.
    'If Not xmlRequestElem.selectSingleNode(".//TYPEOFBUYER") Is Nothing Then
    '    strTypeOfBuyer = m_objXMLAssist.GetNodeValue(xmlRequestElem, "//TYPEOFBUYER", False)
    'Else
    '    strTypeOfBuyer = m_objXMLAssist.GetNodeValue(xmlAppDataResp, "//TYPEOFBUYER", False)
    'End If
    'SG 28/05/02 SYS4767 END
    
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("TYPEOFBUYER"))
    xmlProdListElem.Text = strTypeOfBuyer
    
    ' Lender
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("LENDERID"))
    xmlProdListElem.Text = strLender
    
    ' Country of Origin
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("COUNTRYOFORIGIN"))
    xmlProdListElem.Text = strCountryOfOrigin
    
    ' Product Language
    Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("LANGUAGE"))
    xmlProdListElem.Text = strProductLanguage
    
    'MAR1702
    For intIndex = 1 To UBound(strDOB)
        Set xmlProdListElem = xmlProdListRoot.appendChild(xmlDOMRequest.createElement("APPLICANT" & CStr(intIndex) & "DOB"))
        xmlProdListElem.Text = strDOB(intIndex)
    Next
        
    End If  'MAR696
    
    Dim objErrAssist As New ErrAssist
    Set xmlProdListResp = objIMortgageProductBO.FindProductList(xmlProdListRoot)
    objErrAssist.CheckResponse xmlProdListResp, xmlResponseRoot, True
    xmlResponseRoot.appendChild xmlProdListResp
    
    Set ICostModelAssistBO_FindMortgageProducts = xmlResponseRoot
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
FindMortgageProductsExit:

    '   clean-up section, release object references

    Exit Function

FindMortgageProductsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_FindMortgageProducts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume FindMortgageProductsExit

End Function

Private Function IsMemberOfStaff(ByVal xmlQuoteNodeResp As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Establishes if any of the applicants are members of staff or not.
'
' pass:         xmlQuoteNodeResp  xml Request node
' return:       True if any applicants are members of staff, false if not.
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IsMemberOfStaff"""
    Dim bMemberOfStaff As Boolean
    Dim xmlElemList As IXMLDOMNodeList
    Dim xmlQuoteNodeItem As IXMLDOMElement
    
    On Error GoTo IsMemberOfStaffErr
    
    bMemberOfStaff = False

    Set xmlElemList = xmlQuoteNodeResp.getElementsByTagName("MEMBEROFSTAFF")
    m_objXMLAssist.CheckNode xmlElemList, TypeName(Me), strFunctionName
    
    For Each xmlQuoteNodeItem In xmlElemList
        bMemberOfStaff = m_objConvertAssist.CSafeBool(xmlQuoteNodeItem.Text)
        
        If bMemberOfStaff Then
            Exit For
        End If
    Next

    IsMemberOfStaff = bMemberOfStaff
    Exit Function

IsMemberOfStaffErr:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub EstablishCountryAndLanguage(objComboDO As ComboDO, strProdLanguage As String, strCountryOfOrigin As String)
' header ----------------------------------------------------------------------------------
' description:  Establishes the default country and language for the "UK", "ENGLISH"
'
' pass:         objComboDO the comboDO passed in, product language variable which will be set by this sub,
'               and strCountryOfOrigin which is also set by this sub.
' return:       Language is set in strProdLanguage, country of origin in strCountryOfOrigin
'------------------------------------------------------------------------------------------
    Dim strGroup As String
    Dim strValue As String
    Const strFunctionName As String = "EstablishCountryAndLanguage"
    
    On Error GoTo Failed
    
    ' Establish country of origin
    strGroup = "CountryOfOrigin"
    strValue = "UK"
    strCountryOfOrigin = objComboDO.GetFirstComboValueId(strGroup, strValue)
    
    If Len(strCountryOfOrigin) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeRecordNotFound, "Combo Value not found for: " + strGroup + ", " + strValue
    End If

    ' Establish Language
    strGroup = "MortgageProductLanguage"
    strValue = "Eng"

    strProdLanguage = objComboDO.GetFirstComboValueId(strGroup, strValue)
    
    If Len(strProdLanguage) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeRecordNotFound, "Combo Value " + strValue + " not found"
    End If
    
    Exit Sub

Failed:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function EstablishLender(vxmlRequest As IXMLDOMNode, _
                                 xmlResponseRoot As IXMLDOMNode, _
                                 objIMortgageProductDO As IMortgageProductDO, _
                                 strAppNo As String, _
                                 strAppFactFind As String, _
                                 strQuoteNo As String) As String
    Dim xmlLoanCompRoot As IXMLDOMNode
    Dim xmlLoanCompElem As IXMLDOMNode
    Dim xmlLoanCompItem As IXMLDOMNode
    Dim xmlLoanCompResp As IXMLDOMNode
    Dim domDoc As FreeThreadedDOMDocument40
    Dim objMortSubQuoteDO As IMortgageSubQuoteDO
    Dim objComboDO As ComboDO
    Const strFunctionName As String = "EstablishLender"
    Dim lngLoanAmount As Long                   'BMIDS00656
    Dim lngTotalLoanComponentAmount As Long     'BMIDS00656
    
    On Error GoTo EstablishLenderErr
    
    If m_objContext Is Nothing Then
        Set objMortSubQuoteDO = New MortgageSubQuoteDO
        Set objComboDO = New ComboDO
    Else
        Set objMortSubQuoteDO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    Set domDoc = vxmlRequest.ownerDocument
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Lowest loan component
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set xmlLoanCompRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlLoanCompItem = xmlLoanCompRoot.appendChild(domDoc.createElement("APPLICATIONNUMBER"))
    xmlLoanCompItem.Text = strAppNo
    
    Set xmlLoanCompItem = xmlLoanCompRoot.appendChild(domDoc.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlLoanCompItem.Text = strAppFactFind
    
    Set xmlLoanCompItem = xmlLoanCompRoot.appendChild(domDoc.createElement("MORTGAGESUBQUOTENUMBER"))
    xmlLoanCompItem.Text = strQuoteNo
    
    'BMIDS00656 We only want to get the lender if it's a multi component loan.
    lngLoanAmount = CLng(m_objXMLAssist.GetTagValue(vxmlRequest, "AMOUNTREQUESTED"))
    lngTotalLoanComponentAmount = CLng(m_objXMLAssist.GetTagValue(vxmlRequest, "LOANCOMPONENTAMOUNT"))
    If (lngLoanAmount <> lngTotalLoanComponentAmount) Then
    'BMIDS00656 }
    
        Err.Clear
        
        On Error Resume Next
        Set xmlLoanCompResp = objMortSubQuoteDO.GetLowestLoanComponent(xmlLoanCompRoot)
        Dim bRecordFound As Boolean
        
        bRecordFound = True
        ' Need to ignore record not found error
        m_objErrAssist.SaveErr
        On Error GoTo EstablishLenderErr
        
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
                m_objErrAssist.ReRaise
            Else
                bRecordFound = False
            End If
            
        End If
        
        Dim sLender As String
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Establish the Lender
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If bRecordFound = True Then
            Dim bPortedLoan As Boolean
            Dim xmlLenderResp As IXMLDOMNode
            
            bPortedLoan = m_objConvertAssist.CSafeBool(m_objXMLAssist.GetTagValue(xmlLoanCompResp, "PORTEDLOAN"))
        
            If bPortedLoan = False Then
                Dim xmlMortProdRoot As IXMLDOMNode
                Dim xmlMortProdItem As IXMLDOMNode
                
                Dim sProductCode As String
                Dim sStartDate As String
                    
                sProductCode = m_objXMLAssist.GetNodeValue(xmlLoanCompResp, "MORTGAGEPRODUCTCODE")
                sStartDate = m_objXMLAssist.GetNodeValue(xmlLoanCompResp, "STARTDATE")
                                    
                Set xmlMortProdRoot = domDoc.createElement("MORTGAGEPRODUCT")
                
                Set xmlMortProdItem = xmlMortProdRoot.appendChild(domDoc.createElement("MORTGAGEPRODUCTCODE"))
                xmlMortProdItem.Text = sProductCode
                
                Set xmlMortProdItem = xmlMortProdRoot.appendChild(domDoc.createElement("STARTDATE"))
                xmlMortProdItem.Text = sStartDate
                
                Set xmlLenderResp = objIMortgageProductDO.GetData(xmlMortProdRoot)
            Else
                Dim sLoanGUID As String
                Dim xmlPortedLenderRoot As IXMLDOMNode
                Dim xmlPortedLenderItem As IXMLDOMNode
                
                'MS 11/07/00 SYS0924 need to initialise xmlPortedLenderRoot
                
                Dim xmlDoc As FreeThreadedDOMDocument40
                Set xmlDoc = New FreeThreadedDOMDocument40
                xmlDoc.validateOnParse = False
                xmlDoc.setProperty "NewParser", True
                Set xmlPortedLenderRoot = xmlDoc.createElement("XMLDOCUMENT")
                
                sLoanGUID = m_objXMLAssist.GetNodeValue(xmlLoanCompResp, "MORTGAGELOANGUID")
                Set xmlPortedLenderItem = xmlPortedLenderRoot.appendChild(domDoc.createElement("MORTGAGELOANGUID"))
                xmlPortedLenderItem.Text = sLoanGUID
                
                Set xmlLenderResp = objMortSubQuoteDO.GetPortedLender(xmlPortedLenderRoot)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '' Comment 1
                '' MS 11/07/00 SYS1198. For Phase 2. See also comment 2
                '' We need to deal with record not found here. i.e. We need to do a search
                '' for the default non-panel lender product, so probably best to return an
                '' XML tag which can be checked for by the calling method.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
            End If
            
            sLender = m_objXMLAssist.GetNodeValue(xmlLenderResp, "ORGANISATIONID")
        End If
    End If      'BMIDS00656
    EstablishLender = sLender

EstablishLenderExit:

    '   clean-up section, release object references
    
    Set xmlDoc = Nothing
    
    Exit Function

EstablishLenderErr:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICostModelAssistBO_GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
  On Error GoTo GetDefaultsForNewLoanComponentErr
    
    Const strFunctionName As String = "ICostModelAssistBO_GetDefaultsForNewLoanComponent"

    
    Dim strNodeName As String
    Dim sQuotationType As String
    Dim strAppNo As String
    Dim strAppFactFind As String
    Dim strSubQuoteNo As String
    Dim strAmountRequested As String
    Dim strNoOfComponents As String
    Dim strAdditudeToBorrowing As String
    Dim strResponseNode As String
    Dim strAccountguid  As String
    Dim strOriginalTermInMonths As String
    Dim strOriginalTermInYears As String
    Dim strRepaymentType As String
    Dim strTypeofApplication As String
    Dim strRequest As String
    Dim strRepaymentTypeValueName  As String
    Dim strICBSOutStandingTerm As String
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlAppDataReq As IXMLDOMNode
    Dim xmlNodeReqElem As IXMLDOMNode
    Dim xmlAppDataResp As IXMLDOMNode
    Dim xmlNodeReqItem As IXMLDOMNode
    Dim xmlDOMDocument As FreeThreadedDOMDocument40
    Dim xmlDefaultsResp As IXMLDOMNode
    Dim xmlDefaultsReq As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlTempRequestElement  As IXMLDOMElement
    Dim xmlTempRespDomDoc As New FreeThreadedDOMDocument40
    xmlTempRespDomDoc.validateOnParse = False
    xmlTempRespDomDoc.setProperty "NewParser", True
    Dim xmlComboNode  As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempRequestElem As IXMLDOMElement

    Dim objApplicationBO As IApplicationBO
    Dim objMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objCustomerFinancialBO As ICustomerFinancialBO
    Dim objComboDO As ComboDO
    Dim objXmlAssist As XMLAssist
    
    Dim blnFurtherAdvanceTermInd  As Boolean
    Dim blnFurtherAdvanceRepTypeInd  As Boolean
    
    Set objXmlAssist = New XMLAssist
    
'Stage1:

    strNodeName = "LOANCOMPONENTDETAILS"

    ' Create default repsonse block
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objCustomerFinancialBO = New CustomerFinancialBO
        Set objComboDO = New ComboDO
        
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objCustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)

    ' Need to validate that we have the application number and application fact find number passed in
    ' GetNodeValue will raise an error if the values don't exist, or are empty
    strAppNo = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONNUMBER")
    strAppFactFind = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strNoOfComponents = m_objXMLAssist.GetTagValue(xmlRequestElem, "NUMBEROFCOMPONENTS")
    strAmountRequested = m_objXMLAssist.GetNodeValue(xmlRequestElem, "AMOUNTREQUESTED")
    strSubQuoteNo = m_objXMLAssist.GetNodeValue(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    
    Set xmlDefaultsReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlDOMDocument = xmlDefaultsReq.ownerDocument
    
'Stage2:

    'If Len(strNoOfComponents) = 0 Then
        ' Need to call get the application data by calling ApplicationBO.GetApplicationData
        'Removed above statement above for BM0561 as FURTHER ADV / SPL - Existing term / repayment method not transferred for 2nd loan component.
        
        Set xmlAppDataReq = xmlDefaultsReq.cloneNode(False)
        Set xmlNodeReqElem = xmlAppDataReq.appendChild(xmlDOMDocument.createElement("APPLICATION"))
        
        Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("APPLICATIONNUMBER"))
        xmlNodeReqItem.Text = strAppNo
        
        Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("APPLICATIONFACTFINDNUMBER"))
        xmlNodeReqItem.Text = strAppFactFind
        
        Set xmlAppDataResp = objApplicationBO.GetApplicationData(xmlAppDataReq)
        m_objErrAssist.CheckResponse xmlAppDataResp, xmlResponseRoot
        
        strAdditudeToBorrowing = m_objXMLAssist.GetTagValue(xmlAppDataResp, "ATTITUDETOBORROWINGSCORE")
        blnFurtherAdvanceTermInd = m_objConvertAssist.CSafeBool(m_objXMLAssist.GetTagValue(xmlAppDataResp, "FURTHERADVANCETERMIND"))
        blnFurtherAdvanceRepTypeInd = m_objConvertAssist.CSafeBool(m_objXMLAssist.GetTagValue(xmlAppDataResp, "FURTHERADVANCEREPTYPEIND"))
        strAccountguid = m_objXMLAssist.GetTagValue(xmlAppDataResp, "ACCOUNTGUID")
        strTypeofApplication = m_objXMLAssist.GetTagValue(xmlAppDataResp, "TYPEOFAPPLICATION")
        
    'End If
    
'Stage3:

    ' Establish Loan Component Defaults
    Set xmlNodeReqElem = xmlDefaultsReq.appendChild(xmlDOMDocument.createElement("QUOTATION"))
    
    ' Add the application number
    Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("APPLICATIONNUMBER"))
    xmlNodeReqItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlNodeReqItem.Text = strAppFactFind
    
    ' Add the amount requested
    Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("AMOUNTREQUESTED"))
    xmlNodeReqItem.Text = strAmountRequested
    
    ' Add Mortgage subquote number
    Set xmlNodeReqItem = xmlNodeReqElem.appendChild(xmlDOMDocument.createElement("MORTGAGESUBQUOTENUMBER"))
    xmlNodeReqItem.Text = strSubQuoteNo
    
    ' Get the defaults for the new loan component
    Set xmlDefaultsResp = objMortgageSubQuoteBO.GetDefaultsForNewLoanComponent(xmlDefaultsReq)
    m_objErrAssist.CheckResponse xmlDefaultsResp, xmlResponseRoot
    
    strResponseNode = "MORTGAGECALCULATORDEFAULT"
    m_objXMLAssist.AttachResponseData xmlResponseRoot, xmlDefaultsResp
    
    Set xmlResponseElem = m_objXMLAssist.GetNode(xmlResponseRoot, ".//MORTGAGECALCULATORDEFAULT")
    m_objXMLAssist.ChangeNodeName xmlResponseElem, strResponseNode, "NEWLOANCOMPONENT"
    
'Stage4:

    'MV - 06/06/2002 - DCWP4 - BM078 - BMIDS00032
    'Start
    
    'Check the TypeOfApplicaiton is "FURTHERADVANCE" type or not
    If objComboDO.IsFurtherAdvance(strTypeofApplication) = True Then
        
        'Preparing Request element
        Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        
        Set xmlTempRequestElement = xmlDOMDocument.createElement("MORTGAGELOAN")
        xmlTempRequest.appendChild xmlTempRequestElement
        
        Set xmlTempElement = xmlDOMDocument.createElement("ACCOUNTGUID")
        xmlTempElement.Text = strAccountguid
        xmlTempRequestElement.appendChild xmlTempElement
        
        'Call CustomerFinancialBO.FindMortgageLoanList
        Set xmlTempResponse = objCustomerFinancialBO.FindMortgageLoanList(xmlTempRequest)
        
        'Check Response Error
        m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
        
        'Select Mortgageloan Node
        Set xmlTempNode = xmlTempResponse.selectSingleNode(".//MORTGAGELOAN")
        
        'Extract the values
        strICBSOutStandingTerm = m_objXMLAssist.GetTagValue(xmlTempNode, "ICBSCALCULATEDOUTSTANDINGTERM")
        strOriginalTermInMonths = m_objXMLAssist.GetTagValue(xmlTempNode, "ORIGINALTERMMONTHS")
        strOriginalTermInYears = m_objXMLAssist.GetTagValue(xmlTempNode, "ORIGINALTERMYEARS")
        strRepaymentType = m_objXMLAssist.GetTagValue(xmlTempNode, "REPAYMENTTYPE")

                
'Stage5:

        If blnFurtherAdvanceTermInd Then
            
            If strICBSOutStandingTerm <> "" Then
                strOriginalTermInYears = CStr(Int(CInt(strICBSOutStandingTerm) / 12))
                strOriginalTermInMonths = CStr(CInt(strICBSOutStandingTerm) Mod 12)
            End If
            
            'Select DefaultTermYears Node
            Set xmlTempNode = xmlResponseRoot.selectSingleNode(".//DEFAULTTERMYEARS")
            
            'If DefaultTermYears Node Exists then replace the Node text
            If Not xmlTempNode Is Nothing Then
                xmlTempNode.Text = strOriginalTermInYears
            Else
                'Create a New Element
                Set xmlTempElement = xmlDOMDocument.createElement("DEFAULTTERMYEARS")
                xmlTempElement.Text = strOriginalTermInYears
                xmlResponseElem.appendChild xmlTempElement
            End If
            
            'Select DefaultTermMonths Node
            Set xmlTempNode = xmlResponseRoot.selectSingleNode(".//DEFAULTTERMMONTHS")
            
            'If DefaultTermMonths Node Exists then replace the Node text
            If Not xmlTempNode Is Nothing Then
                xmlTempNode.Text = strOriginalTermInMonths
            Else
                'Create a New Node
                Set xmlTempElement = xmlDOMDocument.createElement("DEFAULTTERMMONTHS")
                xmlTempElement.Text = strOriginalTermInMonths
                xmlResponseElem.appendChild xmlTempElement
            End If
        End If
        
'Stage6:

        If blnFurtherAdvanceRepTypeInd Then
        
            'Get RepaymentType ValueName from ComboValue Table for the above RepaymentType
            'Build Request Element
            Set xmlTempRequest = xmlDOMDocument.createElement("LIST")
            
            Set xmlTempRequestElem = xmlDOMDocument.createElement("GROUPNAME")
            xmlTempRequestElem.Text = "RepaymentType"
            xmlTempRequest.appendChild xmlTempRequestElem
            
            Set xmlTempRequestElem = xmlDOMDocument.createElement("VALUEID")
            xmlTempRequestElem.Text = strRepaymentType
            xmlTempRequest.appendChild xmlTempRequestElem
            
            'Call OmBase.ComboDO.GetcomboValue
            strRepaymentTypeValueName = objComboDO.GetComboValue(xmlTempRequest.xml)
            
            'Load Into Temp Dom Document
            Set xmlTempRespDomDoc = objXmlAssist.load(strRepaymentTypeValueName, TypeName(Me), strFunctionName)
            
            'Extract ValueName from the Response node
            strRepaymentTypeValueName = m_objXMLAssist.GetNodeValue(xmlTempRespDomDoc.documentElement, "//VALUENAME")

        
            'Select RepaymentType Node
            Set xmlTempElement = xmlResponseRoot.selectSingleNode(".//REPAYMENTTYPE")
            
            'If RepaymentType Node exists then replace the node text
            If Not xmlTempElement Is Nothing Then
            
                xmlTempElement.setAttribute "TEXT", objComboDO.GetComboValue(strRepaymentType)
                xmlTempNode.Text = strRepaymentType
            
            Else
                'Create a New Node
                Set xmlTempElement = xmlDOMDocument.createElement("REPAYMENTTYPE")
                xmlTempElement.setAttribute "TEXT", strRepaymentTypeValueName
                xmlTempElement.Text = strRepaymentType
                
                'Attach to the response element
                xmlResponseElem.appendChild xmlTempElement
                
            End If
        End If
    
    End If
    'End
    
'Stage7:
    
    'Append ICBSOUTSTANDINGTERM to the Response Element
    'Create a New Node
    Set xmlTempElement = xmlDOMDocument.createElement("ICBSOUTSTANDINGTERM")
    xmlTempElement.Text = strICBSOutStandingTerm
    'Attach to the response element
    xmlResponseElem.appendChild xmlTempElement
    
    'Append AttitudeToBorrowing to the Response Element
    If Len(strAdditudeToBorrowing) > 0 Then
        Set xmlResponseElem = xmlResponseElem.appendChild(xmlDOMDocument.createElement("ATTIDUDETOBORROWINGSCORE"))
        xmlResponseElem.Text = strAdditudeToBorrowing
    End If
    
    
'End:

    Set ICostModelAssistBO_GetDefaultsForNewLoanComponent = xmlResponseRoot
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDefaultsForNewLoanComponentExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseRoot = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlAppDataReq = Nothing
    Set xmlNodeReqElem = Nothing
    Set xmlAppDataResp = Nothing
    Set xmlNodeReqItem = Nothing
    Set xmlDOMDocument = Nothing
    Set xmlDefaultsResp = Nothing
    Set xmlDefaultsReq = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempElement = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempRequestElement = Nothing
    Set xmlTempRespDomDoc = Nothing
    Set xmlComboNode = Nothing
    Set xmlTempNode = Nothing
    
    Set objApplicationBO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set objCustomerFinancialBO = Nothing
    Set objComboDO = Nothing

    Exit Function

GetDefaultsForNewLoanComponentErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetDefaultsForNewLoanComponentExit
    
    
End Function

Private Function ICostModelAssistBO_CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement, _
                                                     Optional ByVal vstrContext As String = "QUICK QUOTE") As IXMLDOMNode
    
    On Error GoTo CalcCostModelLTVVbErr

    Const strFunctionName As String = "ICostModelAssistBO_CalcCostModelLTV"

    
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objLTVBO As LTVBO
    Dim objAppProcBO  As Object
    Dim objComboDO As ComboDO
    Dim objCustomerFinancialBO As CustomerFinancialBO
    Dim objHomeTrackBO As Object  'JD MAR41
    
    Dim xmlComboNode As IXMLDOMNode
    Dim xmlRequestElem As IXMLDOMNode
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlAppDataRespNode As IXMLDOMNode
    Dim xmlLoansNotRedeemedNode As IXMLDOMNode
    Dim xmlNewPropertyResp As IXMLDOMNode
    Dim xmlCalcLTVRoot As IXMLDOMNode
    Dim xmlCalcLTVNode As IXMLDOMNode
    Dim xmlCalcLTVValue As IXMLDOMNode
    Dim domDoc As FreeThreadedDOMDocument40
    Dim xmlDiscountAmount As IXMLDOMNode
    Dim xmlSharedOwnershipType As IXMLDOMNode
    Dim xmlSharedOwnershipAmount As IXMLDOMNode
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlValReportNode As IXMLDOMNode
    Dim domReq As FreeThreadedDOMDocument40
    Dim domResp As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlSharedPercentage As IXMLDOMElement
    
    Dim strSharedPercentage  As String
    Dim strTagName As String
    Dim strAppNo As String
    Dim strAppFactFind As String
    Dim strAmountRequested As String
    Dim strDiscountAmount As String
    Dim strSharedOwnershipAmount As String
    Dim strPurchasePrice As String
    Dim strTempResp As String
    Dim strRequest As String
    Dim strApplicationType As String
    Dim strCostModelResp As String
    
    Dim blnNotFoundOk As Boolean
    Dim blnValuationNotFound As Boolean
    Dim bPropertyDetailsNotFound As Boolean
    
    Dim dblPurchasePrice As Double
    Dim dblLatestValuation As Double
    Dim dblPostWorksValuation As Double
    Dim dblRetentionRoads As Double
    Dim dblRetentionWorks As Double
            
    Dim lngErrNumber As Long
    Dim blnChangeOfProperty As Boolean  'MAR1300 GHun
    Dim bDefaultLTV As Boolean 'MAR1640 JD
    Dim xmlLTVNode As IXMLDOMNode 'MAR1640
    
    bDefaultLTV = False  'MAR1640 initialise as false
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objLTVBO = New LTVBO
        Set objCustomerFinancialBO = New CustomerFinancialBO
        'Set objHomeTrackBO = New HomeTrackBO  'JD MAR41
    Else
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistdO")
        Set objLTVBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".LTVBO")
        Set objCustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
        'objHomeTrackBO = m_objContext.CreateInstance(gstrHOMETRACK_COMPONENT & ".HomeTrackBO") 'JD MAR41
    End If
    Set objHomeTrackBO = m_objContext.CreateInstance(gstrHOMETRACK_COMPONENT & ".HomeTrackBO") 'JD MAR41
'Stage1:
    
    ' Validate what's passed in
    
    strTagName = "LTV"
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, strTagName)
    
    strAppNo = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONNUMBER")
    strAppFactFind = m_objXMLAssist.GetNodeValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strAmountRequested = m_objXMLAssist.GetNodeValue(xmlRequestElem, "AMOUNTREQUESTED")
    
'Stage2:
    
    ' Get the application data
    
    Set xmlAppDataRespNode = objICostModelAssistDO.GetApplicationData(vxmlRequest)
    m_objXMLAssist.CheckNode xmlAppDataRespNode, TypeName(Me), strFunctionName
    
'Stage3:

    ' Get Loans Not Redeemed
    If vstrContext = "QUICK QUOTE" Then
        Set xmlLoansNotRedeemedNode = _
                        xmlAppDataRespNode.selectSingleNode(".//OUTSTANDINGLOANAMOUNT")
    Else
    
        Dim strResponse  As String
        Dim xmlMtgAcctDataDoc  As FreeThreadedDOMDocument40
        Dim xmlMtgAcctDataReqRoot As IXMLDOMElement
        Dim xmlMtgAcctNode As IXMLDOMNode
        Dim xmlTempReqElem As IXMLDOMElement
        
        Set xmlMtgAcctDataReqRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTempReqElem = xmlOut.createElement("MORTGAGEACCOUNT")
        xmlMtgAcctDataReqRoot.appendChild xmlTempReqElem
        
        Set xmlTempElem = xmlOut.createElement("APPLICATIONNUMBER")
        xmlTempElem.Text = strAppNo
        xmlTempReqElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlTempElem.Text = strAppFactFind
        xmlTempReqElem.appendChild xmlTempElem
        
        'BMIDS591 GHun 02/07/2003
        Set xmlTempElem = xmlOut.createElement("GETBMONLYNOTTOBEREDEEMEDLOANS")
        xmlTempElem.Text = "1"
        xmlTempReqElem.appendChild xmlTempElem
        'BMIDS591 End
        
        strResponse = objCustomerFinancialBO.GetMtgAcctDataForFurtherAdvance(xmlMtgAcctDataReqRoot.xml)
        Set xmlMtgAcctDataDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlMtgAcctDataDoc.documentElement, xmlResponseRoot
       
        Set xmlLoansNotRedeemedNode = xmlMtgAcctDataDoc.selectSingleNode(".//TOTALOUTSTANDINGBALANCE")
        
    End If
    
    ' Record not found is ok
    bPropertyDetailsNotFound = False
    
    On Error Resume Next
    Set xmlNewPropertyResp = objICostModelAssistDO.GetNewPropertyDetails(strAppNo, strAppFactFind)
    m_objErrAssist.SaveErr
    On Error GoTo CalcCostModelLTVVbErr
    
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
            m_objErrAssist.ReRaise
        Else
            bPropertyDetailsNotFound = True
        End If
    End If
    
'Stage4:

    ' Calc the Cost Model LTV
    
    Set xmlCalcLTVRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDOMDoc = xmlCalcLTVRoot.ownerDocument
    Set xmlCalcLTVNode = xmlCalcLTVRoot.appendChild(xmlDOMDoc.createElement("LTV"))
    
    strDiscountAmount = "0"
    strSharedPercentage = "0"
    
    Set xmlDiscountAmount = xmlCalcLTVNode.appendChild(xmlDOMDoc.createElement("DISCOUNTAMOUNT"))
    Set xmlSharedPercentage = xmlCalcLTVNode.appendChild(xmlDOMDoc.createElement("SHAREDPERCENTAGE"))
    
    If bPropertyDetailsNotFound = False Then
        strDiscountAmount = m_objXMLAssist.GetNodeValue(xmlNewPropertyResp, "//DISCOUNTAMOUNT", True, False)
        strSharedPercentage = m_objXMLAssist.GetNodeValue(xmlNewPropertyResp, "//SHAREDPERCENTAGE", True, False)
        xmlCalcLTVNode.appendChild m_objXMLAssist.GetNode(xmlNewPropertyResp, "//SHAREDOWNERSHIPTYPE")
    'MAR1300 GHun
        blnChangeOfProperty = xmlGetNodeAsBoolean(xmlNewPropertyResp, "//CHANGEOFPROPERTY")
    Else
        blnChangeOfProperty = False
    'MAR1300 End
    End If
    
    strPurchasePrice = m_objXMLAssist.GetNodeValue(xmlAppDataRespNode, "//PURCHASEPRICEORESTIMATEDVALUE")
    
    ' get the purchase price + discount amount for comparision
    dblPurchasePrice = m_objConvertAssist.CSafeDbl(strPurchasePrice) + _
                        m_objConvertAssist.CSafeDbl(strDiscountAmount)
                        
    strPurchasePrice = CStr(dblPurchasePrice) 'JD MAR41
    
'Stage5:
    
    If Not blnChangeOfProperty Then 'MAR1300 GHun
        ' Get the latest "present valuation"
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")   'BMIDS01024 MDC 20/11/2002
        
        'JD MAR41 Get the valuation from HomeTrack if applicationtype is remortgage
        If objComboDO.IsItemInValidation("TypeOfMortgage", _
            m_objXMLAssist.GetMandatoryElementText(xmlAppDataRespNode, "//TYPEOFAPPLICATION"), "R") _
            And _
            objComboDO.IsItemInValidation("ValuationType", _
            m_objXMLAssist.GetNodeValue(xmlNewPropertyResp, "//VALUATIONTYPE", True, False), "AU") Then
            
            'We are a remortgage and did a HomeTrack valuation so get the present valuation from there
            Dim xmlHomeTrackValDoc  As FreeThreadedDOMDocument40
            Dim xmlHomeTrackValReqRoot As IXMLDOMElement
            Dim xmlHomeTrackValNode As IXMLDOMNode
            
            Set xmlHomeTrackValReqRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
            Set xmlTempReqElem = xmlOut.createElement("HOMETRACKVALUATION")
            xmlHomeTrackValReqRoot.appendChild xmlTempReqElem
            
            'JD MAR506 hometrackBO expects attribute based xml
            xmlTempReqElem.setAttribute "APPLICATIONNUMBER", strAppNo
            xmlTempReqElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFind
            'Set xmlTempElem = xmlOut.createElement("APPLICATIONNUMBER")
            'xmlTempElem.Text = strAppNo
            'xmlTempReqElem.appendChild xmlTempElem
            
            'Set xmlTempElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
            'xmlTempElem.Text = strAppFactFind
            'xmlTempReqElem.appendChild xmlTempElem
            
            strResponse = objHomeTrackBO.GetPresentValuation(xmlHomeTrackValReqRoot.xml)
            Set xmlHomeTrackValDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlHomeTrackValDoc.documentElement, xmlResponseRoot
           
            dblLatestValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlHomeTrackValDoc.documentElement, "HOMETRACKVALUATION", "VALUATIONAMOUNT")) 'JD MAR506
            
            'JD MAR299 if valuation is 0 use the purchase price
            If (dblLatestValuation <> 0 And dblLatestValuation < m_objConvertAssist.CSafeDbl(strPurchasePrice)) Then
                strPurchasePrice = CStr(dblLatestValuation)
            End If
            
        Else
              
            ' Get the Valuation report
            Set domReq = New FreeThreadedDOMDocument40
            domReq.validateOnParse = False
            domReq.setProperty "NewParser", True
            Set domResp = New FreeThreadedDOMDocument40
            domResp.validateOnParse = False
            domResp.setProperty "NewParser", True
            
            Set xmlElem = domReq.appendChild(domReq.createElement("REQUEST"))
            xmlElem.setAttribute "OPERATION", "GetValuationReport"
            Set xmlTempElem = xmlElem.appendChild(domReq.createElement("VALUATION"))
            xmlTempElem.setAttribute "APPLICATIONNUMBER", strAppNo
            xmlTempElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFind
            
            ' Record not found is ok
            blnValuationNotFound = False
            Set objAppProcBO = m_objContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
            If objAppProcBO Is Nothing Then
                m_objErrAssist.ThrowError gstrAPPLICATIONPROCESSING & ".omAppProcBO", strFunctionName, oeObjectNotCreatable
            End If
            
            strTempResp = objAppProcBO.omAppProcRequest(xmlElem.xml)
            domResp.loadXML (strTempResp)
            Set xmlValReportNode = domResp.firstChild
            
            lngErrNumber = m_objErrAssist.CheckResponse(xmlValReportNode, xmlResponseRoot, False)
            
            'Skip the valuation if record Not found
            If lngErrNumber <> 0 Then
                lngErrNumber = m_objErrAssist.GetOmigaErrorNumber(lngErrNumber)
                If lngErrNumber = oeRecordNotFound Then
                   blnValuationNotFound = True
                Else
                   m_objErrAssist.ThrowError gstrAPPLICATIONQUOTE, strFunctionName, lngErrNumber
                End If
            End If
        
        'Stage6:
        
            If Not blnValuationNotFound Then
                
                dblLatestValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlValReportNode, "GETVALUATIONREPORT", "PRESENTVALUATION"))
                  
                dblPostWorksValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlValReportNode, "GETVALUATIONREPORT", "POSTWORKSVALUATION"))
                dblRetentionRoads = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlValReportNode, "GETVALUATIONREPORT", "RETENTIONROADS"))
                dblRetentionWorks = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlValReportNode, "GETVALUATIONREPORT", "RETENTIONWORKS"))
                
                Dim blnPostWorks As Boolean
                blnPostWorks = False
                
                If dblPostWorksValuation > 0 And (dblRetentionRoads > 0 Or dblRetentionWorks > 0) Then
                    dblLatestValuation = dblPostWorksValuation
                    blnPostWorks = True
                Else
                    If dblLatestValuation <= 0 Then
                       dblLatestValuation = 0
                    End If
                End If
                    
                'BMIDS00938 MDC 18/11/2002 - CC014
        '        ' Get the application (i.e. Mortgage) type
                ' JD MAR41 Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")   'BMIDS01024 MDC 20/11/2002
        '        strRequest = "<LIST>"
        '        strRequest = strRequest & "<GROUPNAME>" & "TypeOfMortgage" & "</GROUPNAME>"
        '        strRequest = strRequest & "<VALUEID>" & m_objXMLAssist.GetNodeValue(xmlAppDataRespNode, "//TYPEOFAPPLICATION") & "</VALUEID>"
        '        strRequest = strRequest & "</LIST>"
        '
        '        strApplicationType = objComboDO.GetComboValue(strRequest)
        '        domResp.loadXML (strApplicationType)
        '        Set xmlComboNode = domResp.firstChild
        '        strApplicationType = m_objXMLAssist.GetNodeValue(xmlComboNode, "//VALUENAME")
        
                ' Check for Remortgage
                If dblLatestValuation <> 0 Then
                    If blnPostWorks Then
                        strPurchasePrice = CStr(dblLatestValuation)
                    Else
                        'If (Mid(strApplicationType, 1, 1) = "R") And dblLatestValuation > 0 Then
                        If objComboDO.IsItemInValidation("TypeOfMortgage", _
                                    m_objXMLAssist.GetMandatoryElementText(xmlAppDataRespNode, "//TYPEOFAPPLICATION"), "N") Then
                            If dblLatestValuation < m_objConvertAssist.CSafeDbl(strPurchasePrice) Then
                                strPurchasePrice = CStr(dblLatestValuation)
                            End If
                        ElseIf dblLatestValuation > 0 Then
                            strPurchasePrice = CStr(dblLatestValuation)
                        End If
                    End If
                Else
                    bDefaultLTV = True
                End If
                'BMIDS00938 MDC 18/11/2002 - End
            End If
        End If
    End If 'MAR1300 GHun
    
    If bDefaultLTV = False Then
        Set xmlCalcLTVValue = xmlCalcLTVNode.appendChild(xmlDOMDoc.createElement("PURCHASEPRICE"))
        xmlCalcLTVValue.Text = strPurchasePrice
        
        ' Amount requested
        Set xmlCalcLTVValue = xmlCalcLTVNode.appendChild(xmlDOMDoc.createElement("AMOUNTREQUESTED"))
        xmlCalcLTVValue.Text = strAmountRequested
        
        xmlCalcLTVNode.appendChild m_objXMLAssist.GetNode(xmlAppDataRespNode, "//TYPEOFAPPLICATION")
        
        xmlDiscountAmount.Text = strDiscountAmount
        xmlSharedPercentage.Text = strSharedPercentage
        
        If Not xmlLoansNotRedeemedNode Is Nothing And Len(xmlLoansNotRedeemedNode.Text) > 0 Then
            Set xmlElement = xmlDOMDoc.createElement("LOANSNOTREDEEMED")
            xmlElement.Text = xmlLoansNotRedeemedNode.Text
            xmlCalcLTVNode.appendChild xmlElement
        End If
        
        strCostModelResp = objLTVBO.CalcCostModelLTV(xmlCalcLTVRoot.xml)
        Set domDoc = m_objXMLAssist.load(strCostModelResp, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse domDoc.documentElement, xmlResponseRoot, True
        m_objXMLAssist.AttachResponseData xmlResponseRoot, domDoc.documentElement
    Else
        'default LTV as 999
        Set xmlLTVNode = xmlResponseRoot.appendChild(xmlOut.createElement("LTV"))
        xmlLTVNode.Text = "999"
    End If
'Stage7:

    'm_objXMLAssist.AttachResponseData xmlResponseRoot, domDoc.documentElement
    
    Set ICostModelAssistBO_CalcCostModelLTV = xmlResponseRoot
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalcCostModelLTVExit:

    '   clean-up section, release object references
    Set objICostModelAssistDO = Nothing
    Set xmlRequestElem = Nothing
    Set xmlResponseRoot = Nothing
    Set objLTVBO = Nothing
    Set xmlOut = Nothing
    Set xmlAppDataRespNode = Nothing
    Set xmlLoansNotRedeemedNode = Nothing
    Set xmlNewPropertyResp = Nothing
    Set xmlCalcLTVRoot = Nothing
    Set xmlCalcLTVNode = Nothing
    Set xmlCalcLTVValue = Nothing
    Set domDoc = Nothing
    Set xmlDiscountAmount = Nothing
    Set xmlSharedOwnershipType = Nothing
    Set xmlSharedOwnershipAmount = Nothing
    Set xmlDOMDoc = Nothing
    Set xmlElement = Nothing
    Set xmlMtgAcctDataDoc = Nothing
    Set xmlMtgAcctDataReqRoot = Nothing
    Set xmlMtgAcctNode = Nothing
    Set xmlTempReqElem = Nothing
        
    Exit Function

CalcCostModelLTVVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_CalcCostModelLTV = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume CalcCostModelLTVExit

End Function


Private Function ICostModelAssistBO_CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                                           xmlLoanCompositionNode As IXMLDOMNode, _
                                                           xmlAppNo As IXMLDOMNode, _
                                                           xmlAppFactFindNo As IXMLDOMNode, _
                                                           xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                                           xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode
    
    
    
    Const strFunctionName As String = "ICostModelAssistBO_CalculateMortgageCosts"
    On Error GoTo CalculateMortgageCostsErr
    'AE MAR506
    Dim strAppNo As String
    Dim strAppFactFind As String
    Dim strNodeName As String, strResponse As String, strUserToCheck As String, strValidUser As String
    Dim strAccountguid  As String, strDeposit As String
    
    Dim objROTBO As Object
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlMortgageSubQuoteElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlContextElem As IXMLDOMNode
    Dim xmlLoansNotRedeemedNode As IXMLDOMElement
    Dim xmlCustomerList As IXMLDOMNode
    Dim xmlValidateUserReq As IXMLDOMNode
    Dim xmlValidateUserElem As IXMLDOMNode
    Dim xmlValidateUserItem As IXMLDOMNode
    Dim xmlValidateUserResp As IXMLDOMNode
    Dim xmlAmountRequestedElem As IXMLDOMNode
    Dim xmlLTVElem As IXMLDOMNode
    Dim xmlMortgageCostsReq As IXMLDOMNode
    Dim xmlMortgageCostsElem As IXMLDOMNode
    Dim xmlMortgageCostsItem As IXMLDOMNode
    Dim xmlMortgageCostsResp As IXMLDOMNode
    Dim xmlAppFactFindNode As IXMLDOMNode
    Dim xmlLoanCompListNode As IXMLDOMNode
    Dim xmlTypeOfAppNode As IXMLDOMNode
    'Dim xmlLifeCoverElem As IXMLDOMNode
    'Dim xmlLifeCoverDom As FreeThreadedDOMDocument40
    Dim xmlAppFFDoc As FreeThreadedDOMDocument40
    Dim xmlAppFFRespDoc As FreeThreadedDOMDocument40
    Dim xmlAppFFRoot As IXMLDOMElement
    Dim xmlAppFFTable As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlROTRequest As IXMLDOMElement
    Dim xmlROTRoot As IXMLDOMElement
    Dim xmlAttribute As IXMLDOMAttribute
    Dim xmlLoanCompList As IXMLDOMNodeList
    Dim xmlCompDate As IXMLDOMElement
    Dim xmlROTResponse As New FreeThreadedDOMDocument40
    xmlROTResponse.validateOnParse = False
    xmlROTResponse.setProperty "NewParser", True
    Dim xmlMortgageAcctDataReq As New FreeThreadedDOMDocument40
    xmlMortgageAcctDataReq.validateOnParse = False
    xmlMortgageAcctDataReq.setProperty "NewParser", True
    Dim xmlMortgageAcctDataReqElem As IXMLDOMNode
    Dim xmlMtgAcctDataElem As IXMLDOMNode, xmlMtgAcctDataElem1  As IXMLDOMNode
    Dim xmlMtgAcctDataRespDoc As New FreeThreadedDOMDocument40
    xmlMtgAcctDataRespDoc.validateOnParse = False
    xmlMtgAcctDataRespDoc.setProperty "NewParser", True
    Dim xmlOutStandingBalance As IXMLDOMNode, xmlOriginalMIGAmt As IXMLDOMNode, xmlAppStaffNode  As IXMLDOMNode
    Dim xmlAppStaff As New FreeThreadedDOMDocument40
    xmlAppStaff.validateOnParse = False
    xmlAppStaff.setProperty "NewParser", True
    Dim xmlAppStaffReqElem As IXMLDOMNode, xmlAppStaffElement As IXMLDOMNode
    Dim xmlAppStaffRespDoc As New FreeThreadedDOMDocument40, xmlROTPurchasePrice As IXMLDOMAttribute
    xmlAppStaffRespDoc.validateOnParse = False
    xmlAppStaffRespDoc.setProperty "NewParser", True
    Dim xmlLatestValAmtsReqDoc As New FreeThreadedDOMDocument40
    xmlLatestValAmtsReqDoc.validateOnParse = False
    xmlLatestValAmtsReqDoc.setProperty "NewParser", True
    Dim xmlLatestValAmtsReq As IXMLDOMElement, xmlLatestValAmtsReqElem As IXMLDOMElement
    Dim xmlLatestValAmtsRespDoc  As New FreeThreadedDOMDocument40
    xmlLatestValAmtsRespDoc.validateOnParse = False
    xmlLatestValAmtsRespDoc.setProperty "NewParser", True
    Dim xmlTempElement As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNewProperty As IXMLDOMElement
    Dim xmlSharePercentage As IXMLDOMNode
    '*=[MC]BMIDS763
    Dim xmlEleAppData As IXMLDOMNode
    'SECTION END
            
    Dim dblROTPurchasePrice  As Double, dblPostWorksValuation   As Double, dblPresentValuation As Double
    Dim dblAppFFPurchasePriceOrEstimatedValue As Double
    
    Dim lngRetentionRoads As Long, lngRetentionWorks As Long
    
    Dim objApplicationBO As ApplicationBO
    Dim objIOrganisationBO As IOrganisationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objCustomerFinancialBO As CustomerFinancialBO
    Dim objApplicationManagerBO As ApplicationManagerBO
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objAppProcBO As Object
    
    'BMIDS00938 MDC 19/11/2002 - CC014
    Dim objComboDO As ComboDO
    'BMIDS00938 MDC 19/11/2002 - End
    
    Dim strLocation As String, strValuationType As String, strSharedPercentage  As String
' BMIDS774
    Dim objIApplClassDef As IApplicationClassDef
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlAFFData As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlAFFElement As IXMLDOMElement
    Dim xmlCompDateEst As IXMLDOMElement
    Dim objIDOAssist As IDOAssist
    
    'MAR41
    Dim objHomeTrackBO As Object
    
    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
' BMIDS774
    
    strNodeName = "MORTGAGECOSTS"
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIOrganisationBO = New OrganisationBO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objApplicationBO = New ApplicationBO
        Set objCustomerFinancialBO = New CustomerFinancialBO
        Set objApplicationManagerBO = New ApplicationManagerBO
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objComboDO = New ComboDO    'BMIDS00938 MDC 19/11/2002 - CC014
    Else
        Set objIOrganisationBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objCustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
        Set objApplicationManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")   'BMIDS00938 MDC 19/11/2002 - CC014
    End If
    
    'BMIDS774
    Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    Set objROTBO = m_objContext.CreateInstance(gstrREPORTONTITLE_COMPONENT & ".omROTBO")

    ' Validate the nodes passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    m_objXMLAssist.CheckNode xmlLoanCompositionNode, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    Set xmlContextElem = m_objXMLAssist.GetNode(vxmlRequest, "//CONTEXT")
    '*=[MC]APPLICATIONDATE DATA RETRIEVAL FROM REQUEST
    Set xmlEleAppData = m_objXMLAssist.GetNode(vxmlRequest, "//APPLICATIONDATE")
    If xmlEleAppData Is Nothing Then
        Set xmlEleAppData = m_objXMLAssist.GetNode(xmlLoanCompositionNode, "//APPLICATIONDATE")
    End If
    'SECTION END
    
    ' And the Loan Composition
    strNodeName = "MORTGAGESUBQUOTE"
    Set xmlMortgageSubQuoteElem = m_objXMLAssist.GetNode(xmlLoanCompositionNode, "/RESPONSE/" & strNodeName, True)
    
    ' SR 19/12/2002 : BM0209 Start
    ' strAccountGUID = xmlLoanCompositionNode.selectSingleNode(".//APPLICATION/ACCOUNTGUID").Text
    dblAppFFPurchasePriceOrEstimatedValue = xmlLoanCompositionNode.selectSingleNode _
                                            (".//APPLICATIONFACTFIND/PURCHASEPRICEORESTIMATEDVALUE").Text  'JD BMIDS763
    ' SR 19/12/2002 : BM0209 - End
'Stage1:

    ' Get the loans that aren't redeemed
    
    Set xmlMortgageAcctDataReq = New FreeThreadedDOMDocument40
    xmlMortgageAcctDataReq.validateOnParse = False
    xmlMortgageAcctDataReq.setProperty "NewParser", True
    Set xmlMortgageAcctDataReqElem = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlMtgAcctDataElem = xmlMortgageAcctDataReq.createElement("MORTGAGEACCOUNT")
    xmlMortgageAcctDataReqElem.appendChild xmlMtgAcctDataElem
    
    xmlMtgAcctDataElem.appendChild xmlAppNo
    xmlMtgAcctDataElem.appendChild xmlAppFactFindNo
    '*=[MC]BMIDS763 - APPEND APPLICATIONDATE
    Call xmlMtgAcctDataElem.appendChild(xmlEleAppData.cloneNode(True))
    'SECTION END
    
    'BMIDS591 GHun 02/07/2003
    Set xmlTempElement = xmlMortgageAcctDataReq.createElement("GETBMONLYNOTTOBEREDEEMEDLOANS")
    xmlTempElement.Text = "1"
    xmlMtgAcctDataElem.appendChild xmlTempElement
    'BMIDS591 End
    
    strResponse = objCustomerFinancialBO.GetMtgAcctDataForFurtherAdvance(xmlMortgageAcctDataReqElem.xml)
    Set xmlMtgAcctDataRespDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    Dim lngErrorNo As Long

    lngErrorNo = m_objErrAssist.CheckResponse(xmlMtgAcctDataRespDoc.documentElement, xmlResponseElem, False)

    ' If not success or record not found re-raise the error
    If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
        m_objErrAssist.ReRaiseResponseError xmlROTResponse.documentElement
    End If
    
    Set xmlOutStandingBalance = xmlMtgAcctDataRespDoc.selectSingleNode(".//TOTALOUTSTANDINGBALANCE")
    If xmlOutStandingBalance Is Nothing Then
        Set xmlOutStandingBalance = xmlOut.createElement("TOTALOUTSTANDINGBALANCE")
        xmlOutStandingBalance.Text = "0"
    End If
    
    Set xmlOriginalMIGAmt = xmlMtgAcctDataRespDoc.selectSingleNode(".//TOTALORIGINALMIGAMOUNT")
    If xmlOriginalMIGAmt Is Nothing Then
        Set xmlOriginalMIGAmt = xmlOut.createElement("TOTALORIGINALMIGAMOUNT")
        xmlOriginalMIGAmt.Text = "0"
    End If
    
'Stage2:
    
    'Get New Property Details
    Set xmlNewProperty = xmlOut.createElement("NEWPROPERTY")
    xmlNewProperty.appendChild xmlAppNo
    xmlNewProperty.appendChild xmlAppFactFindNo
    '*=[MC]BMIDS763 - APPLICATIONDATE
    xmlNewProperty.appendChild xmlEleAppData
    'END SECTION
    
    Set xmlTempResponseNode = objICostModelAssistDO.GetValuationTypeAndLocation(xmlNewProperty)
    
    strLocation = m_objXMLAssist.GetTagValue(xmlTempResponseNode, "PROPERTYLOCATION")
    strValuationType = m_objXMLAssist.GetTagValue(xmlTempResponseNode, "VALUATIONTYPE")
    Set xmlSharePercentage = xmlTempResponseNode.selectSingleNode("SHAREDOWNERSHIPDETAILS/SHAREDPERCENTAGE")
    
'Stage3:
    
    'Is application staff
    Set xmlAppStaffReqElem = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlAppStaffElement = xmlAppStaff.createElement("APPLICATION")
    xmlAppStaffReqElem.appendChild xmlAppStaffElement
    
    xmlAppStaffElement.appendChild xmlAppNo
    xmlAppStaffElement.appendChild xmlAppFactFindNo
    '*=[MC]BMIDS763 - APPLICATION DATE
    xmlAppStaffElement.appendChild xmlEleAppData
    'END SECTION
    
    strResponse = objApplicationManagerBO.IsStaffApplication(xmlAppStaffReqElem.xml)
    Set xmlAppStaffRespDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlAppStaffRespDoc.documentElement, xmlResponseElem, True
    
    Set xmlAppStaffNode = xmlAppStaffRespDoc.selectSingleNode(".//MEMBEROFSTAFF")
    
'Stage4:
    
    ' Validate the user mandate level
    

    Set xmlValidateUserReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)

    Set xmlDOMDoc = xmlValidateUserReq.ownerDocument
    ' Need to add the userID, amount requested, and LTV
    Set xmlValidateUserElem = xmlValidateUserReq.appendChild(xmlDOMDoc.createElement("USER"))
    Set xmlAmountRequestedElem = m_objXMLAssist.GetNode(xmlMortgageSubQuoteElem, "AMOUNTREQUESTED")
    Set xmlLTVElem = m_objXMLAssist.GetNode(xmlMortgageSubQuoteElem, "LTV")

'-- BMIDS00818  SA 05/11/02 {
' The user mandate level should only be checked
' at underwriter review and approval stage

'    xmlValidateUserElem.appendChild xmlAmountRequestedElem.cloneNode(True)
'    xmlValidateUserElem.appendChild xmlLTVElem.cloneNode(True)
'
'    strUserToCheck = m_objXmlAssist.GetAttributeFromNode(vxmlRequest, "USERID")
'
'    If Len(strUserToCheck) = 0 Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "USERID must be entered"
'    End If
'
'    Set xmlValidateUserItem = xmlValidateUserElem.appendChild(xmlDOMDoc.createElement("USERID"))
'    xmlValidateUserItem.Text = strUserToCheck
'
'    Set xmlValidateUserResp = objIOrganisationBO.ValidateUserMandateLevel(xmlValidateUserReq)
'    m_objErrAssist.CheckResponse xmlValidateUserResp, xmlResponseElem, True
'
'    strValidUser = m_objXmlAssist.GetNodeValue(xmlValidateUserResp, "./VALIDUSER")
'
'    If m_objConvertAssist.CSafeBool(strValidUser) = False Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 180
'    End If
    '-- BMIDS00818  SA 05/11/02 }
'Stage5:

    ' Calculate the cost of the mortgage

    Dim strPurchasePrice As String
    
    Set xmlMortgageCostsReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDOMDoc = xmlMortgageCostsReq.ownerDocument
    
    ' Loan Details
    Set xmlLoanCompListNode = m_objXMLAssist.GetNode(xmlLoanCompositionNode, "MORTGAGESUBQUOTE")
    '*=[MC]BMIDS763
    If Not xmlLoanCompListNode Is Nothing Then
        If Not xmlEleAppData Is Nothing Then
            Call xmlLoanCompListNode.appendChild(xmlEleAppData)
        End If
    End If
    
    Set xmlMortgageCostsElem = xmlMortgageCostsReq.appendChild(xmlLoanCompListNode.cloneNode(True))
    
    m_objXMLAssist.CheckNode xmlAppNo, TypeName(Me), strFunctionName
    m_objXMLAssist.CheckNode xmlAppFactFindNo, TypeName(Me), strFunctionName
    m_objXMLAssist.CheckNode xmlMortgageSubQuoteNumber, TypeName(Me), strFunctionName

    ' Purchase price
    Set xmlAppFactFindNode = m_objXMLAssist.GetNode(xmlLoanCompositionNode, "//APPLICATIONFACTFIND")
    strPurchasePrice = m_objXMLAssist.GetNodeValue(xmlAppFactFindNode, "PURCHASEPRICEORESTIMATEDVALUE")
    Set xmlMortgageCostsItem = xmlMortgageCostsElem.appendChild(xmlDOMDoc.createElement("PURCHASEPRICE"))
    xmlMortgageCostsItem.Text = strPurchasePrice
    
    ' Type of application
    Set xmlTypeOfAppNode = m_objXMLAssist.GetNode(xmlAppFactFindNode, "TYPEOFAPPLICATION")
    xmlMortgageCostsElem.appendChild xmlTypeOfAppNode.cloneNode(True)
            
    ' Loans not redeemed
    Set xmlElement = xmlDOMDoc.createElement("LOANSNOTREDEEMED")
    xmlMortgageCostsElem.appendChild xmlElement
    xmlElement.Text = xmlOutStandingBalance.Text
    
    If xmlElement.Text = "" Then
        xmlElement.Text = 0
    End If
    
'Stage6:

    ' Call ROT to get completion date
    Set xmlROTRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    xmlROTRequest.setAttribute "OPERATION", "GetReportOnTitleData"
    Set xmlROTRoot = xmlDOMDoc.createElement("REPORTONTITLE")
    xmlROTRequest.appendChild xmlROTRoot
    xmlROTRoot.setAttribute "APPLICATIONNUMBER", xmlAppNo.Text
    xmlROTRoot.setAttribute "APPLICATIONFACTFINDNUMBER", xmlAppFactFindNo.Text
    
    strResponse = objROTBO.OmRotRequest(xmlROTRequest.xml)
    Set xmlROTResponse = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    lngErrorNo = m_objErrAssist.CheckResponse(xmlROTResponse.documentElement, xmlResponseElem, False)
    
    ' If not success or record not found re-raise the error
    If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
        m_objErrAssist.ReRaiseResponseError xmlROTResponse.documentElement
    End If
    
    'BMIDS774 Need to get the ApplicationFactFind EstimatedCompletionDate
    Set xmlInputElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlAFFElement.Text = xmlAppNo.Text
    xmlTableElem.appendChild xmlAFFElement
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlAFFElement.Text = xmlAppFactFindNo.Text
    xmlTableElem.appendChild xmlAFFElement
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData()
    Set xmlAFFData = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
    
    Set xmlCompDateEst = xmlAFFData.selectSingleNode("ESTIMATEDCOMPLETIONDATE")
    
'BMIDS774    Set xmlAttribute = xmlROTResponse.selectSingleNode("RESPONSE/REPORTONTITLE/@COMPLETIONDATE")
    If Not xmlCompDateEst Is Nothing Then
        'BMIDS774 If EstimatedCompletionDate is >= Sys Date, use it
        'otherwise don't create it so we get the alpha+ default.
        If DateDiff("d", Now, CSafeDate(xmlCompDateEst.Text)) >= 0 Then
            Set xmlLoanCompList = xmlMortgageCostsReq.selectNodes(".//LOANCOMPONENT")
             
            Set xmlCompDate = xmlDOMDoc.createElement("COMPLETIONDATE")
            xmlCompDate.Text = xmlCompDateEst.Text
    
            For Each xmlElement In xmlLoanCompList
                xmlElement.appendChild xmlCompDate.cloneNode(True)
            Next
        End If
    End If
        
    Set xmlROTPurchasePrice = xmlROTResponse.selectSingleNode("RESPONSE/REPORTONTITLE/@PURCHASEPRICE")
    If Not xmlROTPurchasePrice Is Nothing Then
        dblROTPurchasePrice = xmlROTPurchasePrice.Value
    Else
        dblROTPurchasePrice = 0
    End If
    
'Stage7:
    'Get Valuation Details
    If StrComp(xmlContextElem.Text, "COST MODELLING", vbTextCompare) = 0 Then
        
        'JD MAR41 get the hometrack valuation if hometrack was done
        If objComboDO.IsItemInValidation("TypeOfMortgage", xmlTypeOfAppNode.Text, "R") And _
           objComboDO.IsItemInValidation("ValuationType", strValuationType, "AU") Then
           
           'We are a remortgage and did a HomeTrack valuation so get the present valuation from there
             Dim xmlHomeTrackValDoc  As FreeThreadedDOMDocument40
             Dim xmlHomeTrackValReqRoot As IXMLDOMElement
             Dim xmlHomeTrackValNode As IXMLDOMNode
             Dim xmlTempReqElem As IXMLDOMElement
             
             Set xmlHomeTrackValReqRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
             Set xmlTempReqElem = xmlOut.createElement("HOMETRACKVALUATION")
             xmlHomeTrackValReqRoot.appendChild xmlTempReqElem
             
             'JD MAR506 hometrackBO expects attribute based xml
             'JD MAR299 get correct appno
             xmlTempReqElem.setAttribute "APPLICATIONNUMBER", xmlAppNo.Text
             xmlTempReqElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlAppFactFindNo.Text
'             Set xmlTempElement = xmlOut.createElement("APPLICATIONNUMBER")
'             xmlTempElement.Text = xmlAppNo.Text
'             xmlTempReqElem.appendChild xmlTempElement
'
'             Set xmlTempElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
'             xmlTempElement.Text = xmlAppFactFindNo.Text
'             xmlTempReqElem.appendChild xmlTempElement
             
             Set objHomeTrackBO = m_objContext.CreateInstance(gstrHOMETRACK_COMPONENT & ".HomeTrackBO") 'JD MAR41
             strResponse = objHomeTrackBO.GetPresentValuation(xmlHomeTrackValReqRoot.xml)
             Set xmlHomeTrackValDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
             m_objErrAssist.CheckResponse xmlHomeTrackValDoc.documentElement, xmlResponseElem
                
             'dblPresentValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetMandatoryElementText(xmlHomeTrackValDoc.documentElement, ".//VALUATIONAMOUNT"))
             dblPresentValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue(xmlHomeTrackValDoc.documentElement, "HOMETRACKVALUATION", "VALUATIONAMOUNT")) 'JD MAR506
        
             'JD MAR299 if presentvaluation is 0 don't use it.
             If (dblPresentValuation > 0) And (dblPresentValuation < dblAppFFPurchasePriceOrEstimatedValue) Then
                 xmlMortgageCostsItem.Text = CStr(dblPresentValuation)
             Else
                 xmlMortgageCostsItem.Text = CStr(dblAppFFPurchasePriceOrEstimatedValue)
             End If
        Else
            'get the latest valuation report
        
            Set xmlLatestValAmtsReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
            xmlLatestValAmtsReq.setAttribute "OPERATION", "GETLATESTVALUATIONAMOUNTS"
            xmlLatestValAmtsReq.setAttribute "ACTION", "GETLATESTVALUATIONAMOUNTS"
            
            Set xmlLatestValAmtsReqElem = xmlLatestValAmtsReqDoc.createElement("VALUATIONREPVALUATION")
            xmlLatestValAmtsReqElem.setAttribute "APPLICATIONNUMBER", xmlAppNo.Text
            xmlLatestValAmtsReqElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlAppFactFindNo.Text
            '*=[MC]BMIDS763 - APPLICATION DATE
            xmlLatestValAmtsReqElem.setAttribute "APPLICATIONDATE", xmlEleAppData.Text
            xmlLatestValAmtsReq.appendChild xmlLatestValAmtsReqElem
            
            Set objAppProcBO = m_objContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
            strResponse = objAppProcBO.omAppProcRequest(xmlLatestValAmtsReq.xml)
            Set xmlLatestValAmtsRespDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
            
            lngErrorNo = m_objErrAssist.CheckResponse(xmlLatestValAmtsRespDoc.documentElement, xmlResponseElem, False)
            ' If not success or record not found re-raise the error
            If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
                m_objErrAssist.ReRaiseResponseError xmlLatestValAmtsRespDoc.documentElement
            End If
           
            Set xmlTempElement = xmlLatestValAmtsRespDoc.selectSingleNode(".//VALUATIONREPVALUATION")
            If Not xmlTempElement Is Nothing Then
                
                dblPostWorksValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue( _
                                xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "POSTWORKSVALUATION", False))
                dblPresentValuation = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue( _
                                xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "PRESENTVALUATION", False))
                lngRetentionRoads = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue( _
                                xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "RETENTIONROADS", False))
                lngRetentionWorks = m_objConvertAssist.CSafeDbl(m_objXMLAssist.GetAttributeValue( _
                                xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "RETENTIONWORKS", False))
            End If
        
            'BMIDS00938 MDC 18/11/2002 - CC014
    '        'BMIDS00522 - Start
    '        Dim blnPostWorksValuation As Boolean
    '        blnPostWorksValuation = False
    '
    '        If lngRetentionRoads > 0 Or lngRetentionWorks > 0 Then
    '
    '            If dblPostWorksValuation > 0 Then
    '                xmlMortgageCostsItem.Text = dblPostWorksValuation
    '            Else
    '                blnPostWorksValuation = True
    '            End If
    '        Else
    '            blnPostWorksValuation = True
    '        End If
    '
    '        If blnPostWorksValuation Then
    '            If dblPresentValuation > 0 Then
    '                If dblAppFFPurchasePriceOrEstimatedValue <= dblPresentValuation Then
    '                    xmlMortgageCostsItem.Text = dblAppFFPurchasePriceOrEstimatedValue
    '                Else
    '                    xmlMortgageCostsItem.Text = dblPresentValuation
    '                End If
    '            Else
    '                xmlMortgageCostsItem.Text = dblAppFFPurchasePriceOrEstimatedValue
    '            End If
    '        End If
            If (lngRetentionRoads > 0 Or lngRetentionWorks > 0) And dblPostWorksValuation > 0 Then
                xmlMortgageCostsItem.Text = CStr(dblPostWorksValuation)
            Else
                If objComboDO.IsItemInValidation("TypeOfMortgage", xmlTypeOfAppNode.Text, "N") Then
                    'DPF BMIDS01030 - Fix
                    If dblPresentValuation > 0 And dblPresentValuation < dblAppFFPurchasePriceOrEstimatedValue Then
                        xmlMortgageCostsItem.Text = CStr(dblPresentValuation)
                    Else
                        xmlMortgageCostsItem.Text = CStr(dblAppFFPurchasePriceOrEstimatedValue)
                    End If
                Else
                    If dblPresentValuation > 0 Then
                        xmlMortgageCostsItem.Text = CStr(dblPresentValuation)
                    Else
                        xmlMortgageCostsItem.Text = CStr(dblAppFFPurchasePriceOrEstimatedValue)
                    End If
                End If
            End If
        End If 'if HomeTrack
    End If
    'BMIDS00522 - End

'Stage8:
          
    ' Deposit
    strDeposit = m_objXMLAssist.GetNodeValue(xmlMortgageSubQuoteElem, "DEPOSIT", False)
    
    If Len(strDeposit) > 0 Then
        Set xmlMortgageCostsItem = xmlMortgageCostsElem.appendChild(xmlDOMDoc.createElement("DEPOSIT"))
        xmlMortgageCostsItem.Text = strDeposit
    End If
    
    ' Amount requested,Context and LTV
    xmlMortgageCostsElem.appendChild xmlAmountRequestedElem.cloneNode(True)
    xmlMortgageCostsElem.appendChild xmlLTVElem.cloneNode(True)
    xmlMortgageCostsElem.appendChild xmlContextElem.cloneNode(True)
    
    xmlMortgageCostsElem.appendChild xmlOriginalMIGAmt.cloneNode(True)
    xmlMortgageCostsElem.appendChild xmlAppStaffNode.cloneNode(True)
    
    If xmlSharePercentage Is Nothing Then
        Set xmlSharePercentage = xmlOut.createElement("SHAREDPERCENTAGE")
        xmlSharePercentage.Text = "100"
    Else
        If xmlSharePercentage.Text = "" Then
            xmlSharePercentage.Text = "100"
        End If
    End If
    
    xmlMortgageCostsElem.appendChild xmlSharePercentage
    
    Set xmlMortgageCostsResp = objIMortgageSubQuoteBO.CalcAndSaveMortgageCosts(xmlMortgageCostsReq)
    m_objErrAssist.CheckResponse xmlMortgageCostsResp, xmlResponseElem, True
        
    Set ICostModelAssistBO_CalculateMortgageCosts = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateMortgageCostsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDOMDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objIOrganisationBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlMortgageSubQuoteElem = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlContextElem = Nothing
    Set xmlLoansNotRedeemedNode = Nothing
    Set xmlCustomerList = Nothing
    Set xmlValidateUserReq = Nothing
    Set xmlValidateUserElem = Nothing
    Set xmlValidateUserItem = Nothing
    Set xmlValidateUserResp = Nothing
    Set xmlAmountRequestedElem = Nothing
    Set xmlLTVElem = Nothing
    Set xmlMortgageCostsReq = Nothing
    Set xmlMortgageCostsElem = Nothing
    Set xmlMortgageCostsItem = Nothing
    Set xmlMortgageCostsResp = Nothing
    Set xmlAppFactFindNode = Nothing
    Set xmlLoanCompListNode = Nothing
    Set xmlTypeOfAppNode = Nothing
    'Set xmlLifeCoverElem = Nothing
    'Set xmlLifeCoverDom = Nothing
    Set objApplicationBO = Nothing
    Set xmlAppFFDoc = Nothing
    Set xmlAppFFRespDoc = Nothing
    Set xmlAppFFRoot = Nothing
    Set xmlAppFFTable = Nothing
    Set xmlElement = Nothing
    
    ' PSC 11/06/2002 BMIDS00043 - Start
    Set objROTBO = Nothing
    Set xmlROTRequest = Nothing
    Set xmlROTRoot = Nothing
    Set xmlAttribute = Nothing
    Set xmlLoanCompList = Nothing
    Set xmlCompDate = Nothing
    Set xmlROTResponse = Nothing
    ' PSC 11/06/2002 BMIDS00043 - End
    
    Set xmlMortgageAcctDataReq = Nothing
    Set xmlMortgageAcctDataReqElem = Nothing
    Set xmlMtgAcctDataElem = Nothing
    Set xmlMtgAcctDataElem1 = Nothing
    Set xmlMtgAcctDataRespDoc = Nothing
    Set xmlOutStandingBalance = Nothing
    Set xmlOriginalMIGAmt = Nothing
    Set xmlAppStaffNode = Nothing
    Set xmlAppStaff = Nothing
    Set xmlAppStaffReqElem = Nothing
    Set xmlAppStaffElement = Nothing
    Set xmlAppStaffRespDoc = Nothing
    Set xmlROTPurchasePrice = Nothing
    Set xmlLatestValAmtsReqDoc = Nothing
    Set xmlLatestValAmtsReq = Nothing
    Set xmlLatestValAmtsReqElem = Nothing
    Set xmlLatestValAmtsRespDoc = Nothing
    Set xmlTempElement = Nothing
    Set objICostModelAssistDO = Nothing
    Set objComboDO = Nothing    'BMIDS00938 MDC 19/11/2002 - CC014
    'BMIDS774
    Set xmlAFFData = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlAFFElement = Nothing
    Set xmlCompDateEst = Nothing
    Set objIApplClassDef = Nothing
    Set xmlApplClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    Exit Function

CalculateMortgageCostsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_CalculateMortgageCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateMortgageCostsExit
    
End Function



Private Function ICostModelAssistBO_GetLoansNotRedeemed(vxmlRequest As IXMLDOMNode) As IXMLDOMNode
    ' header ----------------------------------------------------------------------------------
' description:  Common method which returns the number of loans that aren't not redeemed
'
' pass:         vxmlRequest
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLoansNotRedeemedErr
    
    Const strFunctionName As String = "ICostModelAssistBO_GetLoansNotRedeemed"


    Dim objCustomerFinancialBO As CustomerFinancialBO
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlLoansNotRedeemedRoot As IXMLDOMElement
    Dim xmlLoansNotRedeemedDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
            
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objCustomerFinancialBO = New CustomerFinancialBO
    Else
        Set objCustomerFinancialBO = _
                m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
        
    Dim xmlCustomerList As IXMLDOMNode
    Set xmlCustomerList = m_objXMLAssist.GetNode(vxmlRequest, ".//CUSTOMERLIST")
    
    m_objXMLAssist.CheckNode xmlCustomerList, TypeName(Me), strFunctionName, "CustomerList is empty"
    Set xmlLoansNotRedeemedRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    xmlLoansNotRedeemedRoot.appendChild xmlCustomerList.cloneNode(True)
    
    Dim strLoansNotRedeemed As String
    
    ' PSC 30/06/00 SYS1012 - Start
    strLoansNotRedeemed = objCustomerFinancialBO.GetLoansNotRedeemedTotals(xmlLoansNotRedeemedRoot.xml)
    Set xmlLoansNotRedeemedDoc = m_objXMLAssist.load(strLoansNotRedeemed, _
                                                   TypeName(Me), _
                                                   strFunctionName)
    m_objErrAssist.CheckResponse xmlLoansNotRedeemedDoc.documentElement, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlLoansNotRedeemedDoc.documentElement
    ' PSC 30/06/00 SYS1012 - End


    Set ICostModelAssistBO_GetLoansNotRedeemed = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetLoansNotRedeemedExit:

    '   clean-up section, release object references
    Set objCustomerFinancialBO = Nothing
    Set xmlOut = Nothing
    Set xmlLoansNotRedeemedRoot = Nothing
    Set xmlLoansNotRedeemedDoc = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetLoansNotRedeemedErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_GetLoansNotRedeemed = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoansNotRedeemedExit

End Function

Private Function ICostModelAssistBO_BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' description:  Calculate the Term and Endowment Needs and Terms excluding non-ported loan
'               components
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>"
'   <LIFECOVER>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'       <LOANCOMPONENTLIST>
'           <LOANCOMPONENT>
'           </LOANCOMPONENT>
'       </LOANCOMPONENTLIST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER></CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LIFECOVER>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element

'MAR1893 GHun No longer used
'    Const strFunctionName As String = "ICostModelAssistBO_BuildLifeCoverCalcInputInterface"
'    Dim strNodeName As String
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim xmlResponseRoot As IXMLDOMElement
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlResponseItem As IXMLDOMElement
'    Dim xmlRequestElem As IXMLDOMElement
'    Dim xmlGetDefaultsResp As IXMLDOMNode
'    Dim objCustomerDO As ICustomerDO
'
'    On Error GoTo BuildLifeCoverCalcInputInterfaceErr
'
'    ' Create default repsonse block
'    Set xmlOut = New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
'    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
'
'    If m_objContext Is Nothing Then
'        Set objCustomerDO = New CustomerDO
'    Else
'        Set objCustomerDO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerDO")
'    End If
'
'    ' Validate what's passsed in
'    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
'    strNodeName = "LIFECOVERCOSTS"
'    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName)
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Get term and endowment needs
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim xmlLoanComponentElem As IXMLDOMElement
'    Dim lEndowmentNeed As Long
'    Dim lTermAssuranceNeed As Long
'    Dim lInterestOnlyTerm As Long
'    Dim lCapitalAndInterestTerm As Long
'
''   SR : 10/04/00 - Search for LOANCOMPONENTLIST rather than "MORTGAGESUBQUOTE/LOANCOMPONENTLIST"
''   Set xmlLoanComponentElem = m_objXMLAssist.GetNode(xmlRequestElem, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST")
'    Set xmlLoanComponentElem = m_objXMLAssist.GetNode(xmlRequestElem, ".//LOANCOMPONENTLIST")
'
'    GetTermAndEndowmentNeeds lEndowmentNeed, _
'                             lTermAssuranceNeed, _
'                             lInterestOnlyTerm, _
'                             lCapitalAndInterestTerm, _
'                             xmlLoanComponentElem
'
'    ' Build XML containing all above
'    ' Endowment Needs
'    Set xmlResponseItem = xmlResponseRoot.appendChild(xmlOut.createElement("ENDOWMENTNEED"))
'    xmlResponseItem.Text = CStr(lEndowmentNeed)
'
'    ' Term Assurance Needs
'    Set xmlResponseItem = xmlResponseRoot.appendChild(xmlOut.createElement("TERMASSURANCENEED"))
'    xmlResponseItem.Text = CStr(lTermAssuranceNeed)
'
'    ' Interest Only Term
'    Set xmlResponseItem = xmlResponseRoot.appendChild(xmlOut.createElement("INTERESTONLYTERM"))
'    xmlResponseItem.Text = CStr(lInterestOnlyTerm)
'
'    ' Capital and Interest Term
'    Set xmlResponseItem = xmlResponseRoot.appendChild(xmlOut.createElement("CAPITALANDINTERESTTERM"))
'    xmlResponseItem.Text = CStr(lCapitalAndInterestTerm)
'
'    Set ICostModelAssistBO_BuildLifeCoverCalcInputInterface = xmlResponseRoot
'
'    ' All done
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'
'
'BuildLifeCoverCalcInputInterfaceExit:
'
'    '   clean-up section, release object references
'
'    Exit Function
'
'BuildLifeCoverCalcInputInterfaceErr:
'
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseRoot
'        Resume Next
'    End If
'
'    m_objErrAssist.AddToErrSource strFunctionName
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    Set ICostModelAssistBO_BuildLifeCoverCalcInputInterface = m_objErrAssist.CreateErrorResponseEx
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' Append
'    ' go to clean-up section
'    Resume BuildLifeCoverCalcInputInterfaceExit

End Function

'MAR1893 GHun No longer used
'Private Sub GetTermAndEndowmentNeeds(lEndowmentNeed As Long, _
'                                     lTermAssuranceNeed As Long, _
'                                     lInterestOnlyTerm As Long, _
'                                     lCapitalAndInterestTerm As Long, _
'                                     xmlLoanComponentElem As IXMLDOMElement)
'    On Error GoTo GetTermAndEndowmentNeedsErr
'    Dim bPortedLoan As Boolean
'    Dim intTerm As Integer
'    Dim intTermInYears As Integer
'    Dim intTermInMonths As Integer
'    Dim lLoanAmount As Long
'    Const strFunctionName As String = "GetTermAndEndowmentNeeds"
'    Dim strRepaymentMethod As String
'    Dim xmlLoanComponentList As IXMLDOMNodeList
'    Dim xmlLoanComponentItem As IXMLDOMNode
'    Dim objComboDO As ComboDO
'
'    If m_objContext Is Nothing Then
'        Set objComboDO = New ComboDO
'    Else
'        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
'    End If
'
'    ' Get the node list
'    Set xmlLoanComponentList = xmlLoanComponentElem.getElementsByTagName("LOANCOMPONENT")
'    m_objXMLAssist.CheckNode xmlLoanComponentList, TypeName(Me), strFunctionName
'
'    lEndowmentNeed = 0
'    lInterestOnlyTerm = 0
'    lTermAssuranceNeed = 0
'    lCapitalAndInterestTerm = 0
'
'    For Each xmlLoanComponentItem In xmlLoanComponentList
'        bPortedLoan = m_objConvertAssist.CSafeBool(m_objXMLAssist.GetNodeValue(xmlLoanComponentItem, "PORTEDLOAN", False))
'        lLoanAmount = m_objConvertAssist.CSafeLng(m_objXMLAssist.GetNodeValue(xmlLoanComponentItem, "LOANAMOUNT", False))
'
'        If bPortedLoan = False Then
'            strRepaymentMethod = m_objXMLAssist.GetNodeValue(xmlLoanComponentItem, "REPAYMENTMETHOD")
'
'            ' Repayment comes in as a ValueID, so find the ValidationValue
'            strRepaymentMethod = objComboDO.GetFirstComboValidation("RepaymentType", strRepaymentMethod)
'
'            intTermInYears = m_objConvertAssist.CSafeInt(m_objXMLAssist.GetNodeValue(xmlLoanComponentItem, "TERMINYEARS"))
'            intTermInMonths = m_objConvertAssist.CSafeInt(m_objXMLAssist.GetNodeValue(xmlLoanComponentItem, "TERMINMONTHS"))
'
'            intTerm = (intTermInYears * 12) + intTermInMonths
'
'            Select Case strRepaymentMethod
'                Case cINTEREST_ONLY
'                    lEndowmentNeed = lEndowmentNeed + lLoanAmount
'
'                    If intTerm > lInterestOnlyTerm Then
'                        lInterestOnlyTerm = intTerm
'                    End If
'
'                Case cCAPITAL_AND_INTEREST
'                    lTermAssuranceNeed = lTermAssuranceNeed + lLoanAmount
'
'                    If intTerm > lCapitalAndInterestTerm Then
'                        lCapitalAndInterestTerm = intTerm
'                    End If
'
'                Case cPART_AND_PART
'                    lEndowmentNeed = lEndowmentNeed + lLoanAmount
'                    lTermAssuranceNeed = lTermAssuranceNeed + lLoanAmount
'
'                    If intTerm > lInterestOnlyTerm Then
'                        lInterestOnlyTerm = intTerm
'                    End If
'
'                    If intTerm > lCapitalAndInterestTerm Then
'                        lCapitalAndInterestTerm = intTerm
'                    End If
'            End Select
'        End If
'    Next
'
'    Exit Sub
'GetTermAndEndowmentNeedsErr:
'    m_objErrAssist.AddToErrSource strFunctionName
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Sub
'MAR1892 End

Private Function ICostModelAssistBO_BuildLifeCoverInterfaceXML(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

'MAR1893 GHun No longer used
'    On Error GoTo BuildLifeCoverInterfaceErr
'    Const strFunctionName As String = "ICostModelAssistBO_BuildLifeCoverInterfaceXML"
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlResponseItem As IXMLDOMElement
'    Dim xmlOut As FreeThreadedDOMDocument40
'    ' The Request should contain the customer version XML
'
'    Set xmlOut = New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Get the customer list
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim xmlCustomerList As IXMLDOMNodeList
'    Dim xmlCustomerNode As IXMLDOMNode
'    Dim xmlCustomerVersionNode As IXMLDOMNode
'    Dim strSmokerStatus As String
'    Dim strGoodhealth As String
'    Dim strAge As String
'    Dim intCustomerIndex As Integer
'    Dim strGender As String
'    Dim strXML As String
'
'    Set xmlCustomerList = vxmlRequest.getElementsByTagName("CUSTOMER")
'
'    ' Must be 1 or 2 applicants
'    If xmlCustomerList.length > 0 And xmlCustomerList.length < 3 Then
'        intCustomerIndex = 1
'
'        For Each xmlCustomerNode In xmlCustomerList
'            strAge = m_objXMLAssist.GetNodeValue(xmlCustomerNode, ".//AGE")
'            strGender = m_objXMLAssist.GetNodeValue(xmlCustomerNode, ".//GENDER")
'
'            '-----------------GD     02/07/2002  BMIDS00166 START------------------
'            'GD strGoodhealth = m_objXmlAssist.GetNodeValue(xmlCustomerNode, ".//GOODHEALTH")
'            'GD strSmokerStatus = m_objXmlAssist.GetNodeValue(xmlCustomerNode, ".//SMOKERSTATUS")
'
'            'GD strXML = "APPLICANT" & intCustomerIndex & "SMOKERSTATUS"
'            'GD Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement(strXML))
'            'GD xmlResponseItem.Text = strSmokerStatus
'            '-----------------GD     02/07/2002  BMIDS00166 END-------------------
'
'            strXML = "APPLICANT" & intCustomerIndex & "GENDER"
'            Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement(strXML))
'            xmlResponseItem.Text = strGender
'
'            '-----------------GD     02/07/2002  BMIDS00166 START----------------
'            'GD strXML = "APPLICANT" & intCustomerIndex & "GOODHEALTH"
'            'GD Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement(strXML))
'            'GD xmlResponseItem.Text = strGoodhealth
'            '-----------------GD     02/07/2002  BMIDS00166 END------------------
'
'            strXML = "APPLICANT" & intCustomerIndex & "AGE"
'            Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement(strXML))
'            xmlResponseItem.Text = strAge
'
'            intCustomerIndex = intCustomerIndex + 1
'        Next
'    Else
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 541
'    End If
'
'    Set ICostModelAssistBO_BuildLifeCoverInterfaceXML = xmlResponseElem
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'BuildLifeCoverInterfaceExit:
'
'    '   clean-up section, release object references
'
'    Exit Function
'
'BuildLifeCoverInterfaceErr:
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
'    Set ICostModelAssistBO_BuildLifeCoverInterfaceXML = m_objErrAssist.CreateErrorResponseEx
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' go to clean-up section
'    Resume BuildLifeCoverInterfaceExit

End Function
'MAR1893 End

Private Function ICostModelAssistBO_CalcAndSaveLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

'MAR1893 GHun No longer used
'    On Error GoTo CalcAndSaveLifeCoverErr
'    Const strFunctionName As String = "ICostModelAssistBO_CalcAndSaveLifeCover"
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlResponseItem As IXMLDOMElement
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim objLifeCoverBO As LifeCoverBO
'    'MS 19/07/00 Modified for performance enhancements
'    Dim objLifeSubQuoteBO As LifeSubQuoteBO
'    Dim xmlDOMDoc As FreeThreadedDOMDocument40
'
'    Set xmlOut = New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
'
'    If m_objContext Is Nothing Then
'        Set objLifeCoverBO = New LifeCoverBO
'        'MS 19/07/00 Modified for performance enhancements
'        Set objLifeSubQuoteBO = New LifeSubQuoteBO
'    Else
'        Set objLifeCoverBO = m_objContext.CreateInstance(gstrLIFECOVER_COMPONENT & ".LifeCoverBO")
'        'MS 19/07/00 Modified for performance enhancements
'        Set objLifeSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".LifeSubQuoteBO")
'    End If
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Calculate premiums
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim xmlPremiumsResp As IXMLDOMNode
'    Dim strCalcPremiumResp As String
'    Dim xmlDOMPremiumResp As FreeThreadedDOMDocument40
'
'
'    strCalcPremiumResp = objLifeCoverBO.CalculatePremiums(vxmlRequest.xml)
'    Set xmlDOMPremiumResp = m_objXMLAssist.load(strCalcPremiumResp, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlDOMPremiumResp.documentElement, xmlResponseElem, True
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Save Life Cover Details
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'MS 19/07/00 Modified for performance enhancements
'    Dim xmlSaveSubQuoteResp As FreeThreadedDOMDocument40
'    Dim xmlSaveSubQuoteReq As IXMLDOMNode
'    Dim xmlSaveSubQuoteElem As IXMLDOMNode
'    Dim xmlAppNo As IXMLDOMNode
'    Dim xmlAppFactFindNo As IXMLDOMNode
'    Dim xmlSubQuote As IXMLDOMNode
'    Dim xmlCustomers As IXMLDOMNode
'
'    Set xmlAppNo = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVER/APPLICATIONNUMBER")
'    Set xmlAppFactFindNo = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVER/APPLICATIONFACTFINDNUMBER")
'    Set xmlSubQuote = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVER/LIFESUBQUOTENUMBER")
'    Set xmlCustomers = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVER/CUSTOMERLIST")
'
'    Set xmlSaveSubQuoteReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlDOMDoc = xmlSaveSubQuoteReq.ownerDocument
'    Set xmlSaveSubQuoteElem = xmlSaveSubQuoteReq.appendChild(xmlDOMDoc.createElement("LIFESUBQUOTE"))
'
'    m_objXMLAssist.AttachResponseData xmlSaveSubQuoteElem, xmlDOMPremiumResp.documentElement
'
'    xmlSaveSubQuoteElem.appendChild xmlAppNo.cloneNode(True)
'    xmlSaveSubQuoteElem.appendChild xmlAppFactFindNo.cloneNode(True)
'    xmlSaveSubQuoteElem.appendChild xmlSubQuote.cloneNode(True)
'    xmlSaveSubQuoteElem.appendChild xmlCustomers.cloneNode(True)
'
'    'MS 19/07/00 Modified for performance enhancements
'    Dim strSaveSubQuoteResp As String
''    Set xmlSaveSubQuoteResp = objILifeSubQuoteBO.SaveRecostedBenefits(xmlSaveSubQuoteReq)
'    strSaveSubQuoteResp = objLifeSubQuoteBO.SaveRecostedBenefits(xmlSaveSubQuoteReq.xml)
'    Set xmlSaveSubQuoteResp = m_objXMLAssist.load(strSaveSubQuoteResp, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlSaveSubQuoteResp.documentElement, xmlResponseElem, True
'
'    Set ICostModelAssistBO_CalcAndSaveLifeCover = xmlResponseElem
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'CalcAndSaveLifeCoverExit:
'
'    '   clean-up section, release object references
'
'    Exit Function
'
'CalcAndSaveLifeCoverErr:
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
'    Set ICostModelAssistBO_CalcAndSaveLifeCover = m_objErrAssist.CreateErrorResponseEx
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' go to clean-up section
'    Resume CalcAndSaveLifeCoverExit

End Function
'MAR1893 End

Private Function ICostModelAssistBO_GetLTV(ByVal vxmlRequest As IXMLDOMElement, Optional ByVal strContext As Variant = "QUICK QUOTE") As IXMLDOMNode
    Const strFunctionName As String = "ICostModelAssistBO_GetLTV"
    Dim objQuotationBO As QuotationBO
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNodeCopyTo As IXMLDOMNode
    Dim xmlNodeCopyFrom As IXMLDOMNode
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlTemp As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String

    Dim strAmountRequested As String
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim objXmlAssist As New XMLAssist

On Error GoTo GetLTVErr
    
    ' Take a copy of the xml passed in
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    xmlIn.loadXML (vxmlRequest.xml)
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
  
    If m_objContext Is Nothing Then
        Set objQuotationBO = New QuotationBO
        Set objMortgageSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If

    Dim strResponse As String
    
    m_objXMLAssist.ChangeNodeName xmlIn, "ACTIVEQUOTENUMBER", "QUOTATIONNUMBER"
    
    ' Get quotation data
    strResponse = objQuotationBO.GetData(xmlIn.xml)
    Set xmlTemp = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTemp.documentElement, xmlResponseRoot, True
    
    ' use get node here
    Set xmlNodeCopyFrom = objXmlAssist.GetNode(xmlTemp, "//MORTGAGESUBQUOTENUMBER", True)
    Set xmlNodeCopyTo = objXmlAssist.GetNode(xmlIn, "//QUOTATION", True)
    
    objXmlAssist.CopyNode "MORTGAGESUBQUOTENUMBER", "MORTGAGESUBQUOTENUMBER", xmlNodeCopyFrom, xmlNodeCopyTo, False
      
    ' call MortgageSubQuoteDO.GetData
    ' Need to change tag name
    objXmlAssist.ChangeNodeName xmlIn, "QUOTATION", "MORTGAGESUBQUOTE"

    strResponse = objMortgageSubQuoteBO.GetData(xmlIn.xml)
    Set xmlTemp = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTemp.documentElement, xmlResponseRoot, True
    
    Dim xmlAmountRequestedNode As IXMLDOMNode
    Dim xmlMSubQuoteNode As IXMLDOMNode
    'SA BMIDS00437 Change search pattern.
    'Set xmlMSubQuoteNode = objXmlAssist.GetNode(xmlIn, "REQUEST/MORTGAGESUBQUOTE", True)
    Set xmlMSubQuoteNode = objXmlAssist.GetNode(xmlIn, ".//MORTGAGESUBQUOTE", True)
    
    objXmlAssist.CopyNode "AMOUNTREQUESTED", "AMOUNTREQUESTED", xmlTemp.firstChild, xmlMSubQuoteNode, True
    
    objXmlAssist.ChangeNodeName xmlMSubQuoteNode, "MORTGAGESUBQUOTE", "LTV"
    
    'copy the node over to the repsonse
    objXmlAssist.CopyNode "LTV", "LTV", xmlMSubQuoteNode, xmlResponseRoot, True
    
    Set ICostModelAssistBO_GetLTV = xmlResponseRoot

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetLTVExit:
    ' Clean up section
    Set objQuotationBO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set xmlOut = Nothing
    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNodeCopyTo = Nothing
    Set xmlNodeCopyFrom = Nothing
    Set xmlTempNode = Nothing
    Set xmlTemp = Nothing
    Set xmlResponseRoot = Nothing
    Set xmlAmountRequestedNode = Nothing
    Set xmlMSubQuoteNode = Nothing

Exit Function

GetLTVErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICostModelAssistBO_GetLTV = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    ' Go to clean up section
    Resume GetLTVExit
    
End Function

'MAR1702
Private Sub GetDOB(ByVal strAppNo As String, ByVal strAppFFNo As String, _
    ByRef strDOB() As String)
'----------------------------------------------------------------------------------
' Description:
'   Returns the date of birth of the applicants.
'   This can then be used when select a product by checking the maximum and minimum
'   application ages for a product
'----------------------------------------------------------------------------------

On Error GoTo GetMaxMinAgeVbErr

    Const strFunctionName As String = "GetMaxMinAge"

    Dim adoParameter As ADODB.Parameter
    Dim adoCommand As ADODB.Command
    Dim adoConnection As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    Dim objADOAssist As ADOAssist
    Dim intIndex As Integer
    
    If m_objContext Is Nothing Then
      Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = objADOAssist.GetConnStr
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
       
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = "{CALL USP_GETAPPLICATIONAGE(?,?)}"
    
    Set adoParameter = adoCommand.CreateParameter("p_APPID", adBSTR, adParamInput, Len(strAppNo), strAppNo)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("p_FFID", adInteger, adParamInput, , strAppFFNo)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    Set rstThisRecSet = New ADODB.Recordset
    
    Set rstThisRecSet.Source = adoCommand
    rstThisRecSet.CursorType = adOpenStatic
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.LockType = adLockReadOnly
    rstThisRecSet.Open
                
    If rstThisRecSet.State <> adStateOpen Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    If rstThisRecSet.RecordCount > 0 Then
        ReDim strDOB(rstThisRecSet.RecordCount)
        intIndex = 1
        rstThisRecSet.MoveFirst
        Do While Not rstThisRecSet.EOF
            strDOB(intIndex) = rstThisRecSet.Fields("DateOfBirth").Value
            intIndex = intIndex + 1
            rstThisRecSet.MoveNext
        Loop
    Else
        ReDim strDOB(0)
    End If
        
GetMaxMinAgeExit:

    If Not (adoParameter Is Nothing) Then Set adoParameter = Nothing
    If Not (adoCommand Is Nothing) Then Set adoCommand = Nothing
    If Not (adoConnection Is Nothing) Then Set adoConnection = Nothing
    If Not (rstThisRecSet Is Nothing) Then Set rstThisRecSet = Nothing
    If Not (objADOAssist Is Nothing) Then Set objADOAssist = Nothing
    
Exit Sub

GetMaxMinAgeVbErr:

    If Not (adoParameter Is Nothing) Then Set adoParameter = Nothing
    If Not (adoCommand Is Nothing) Then Set adoCommand = Nothing
    If Not (adoConnection Is Nothing) Then Set adoConnection = Nothing
    If Not (rstThisRecSet Is Nothing) Then Set rstThisRecSet = Nothing
    If Not (objADOAssist Is Nothing) Then Set objADOAssist = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


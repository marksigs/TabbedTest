VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "MortgageSubQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageSubQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Mortgage SubQuote Business Object which requires transactions
'
'Dependencies:  MortgageSubQuoteDO
'
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MCS    18/01/00    Created
'PSC    04/04/00    AQR SYS0593 - Reset LTV in ResetMortgageSubQuote
'JLD    05/04/00    Changes for mysterious 'Parameter is invalid' error on Calculate
'PSC    07/04/00    Solved mysterious 'Parameter is invalid' error on Calculate
'JLD    11/04/00    AQR SYS0603 - save TOTALNETMONTHLYCOST as double not long
'JLD    18/04/00    bug fix on call of DeleteAllMortgageIncentive
'JLD    19/04/00    in SaveLoanComponentDetails, deal correctly with the mortgage incentives.
'JLD    25/04/00    On call to SaveLoanComponentXML set the 'new' flag for creating a component.
'PSC    27/04/00    AQR SYS0593 - Reset LTV in ResetMortgageSubQuote to the value passed in
'JLD    28/04/00    Use correct tag to get one off cost type.
'JLD    02/05/00    Update TOTALLOANCOMPONENTAMOUNT not LOANAMOUNT in SaveOneOffCostDetails().
'                   CreateMortgageIncentive is a sub.
'                   In SaveLoanComponentDetails, don't DeleteAll mortgageincentives if we're a new loancomponent
'JLD    05/05/00    SYS0690 - change name of MORTGAGEINCENTIVESEQUENCENUMBE
'JLD    12/05/00    SYS0715 - update DEPOSIT when in ResetMortgageSubquote
'APS    30/05/00    SYS0767 - Change to CreateFirstSubQuote
'MS     12/06/00    SYS0829 - Make sure we save the charges on calculate
'MS     05/07/00    Not an error to have no Miras
'MS     11/07/00    SYS0924 error creating loan component with a ported loan
'CL     26/9/00     SYS0952 Total charges not being updated properly, also changed
'                   dblCostsToBeAdded to Double not Long
'PSC    23/03/01    SYS2053 Added UpdateLoanComponent
'PSC    23/03/01    SYS2166 Amend to use correct one off cost data and only append them to the
'                   main loan component
'SR     18/06/01    modified method 'IMortgageSubQuoteTxBO_SaveLoanComponentDetails'
'PSC    17/01/02    SYS3801 Amend CreateMortgageIncentiveXML it cater for new fields
'SG     28/05/02    SYS7467 MSMS to Core integration
'AShaw  20/12/06    EP2_56 Add param to CreateNewSubQuote.
'                   EP2_56 New methods for CreateComponentsFromExistingAcc.
'AShaw  25/01/2007  EP2_893 Alter code to NOT run CalcAndSaveProcFees if no Data
'AShaw  21/02/2007  EP2_8 IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts alteration.
'------------------------------------------------------------------------------------------
Option Explicit
Implements IMortgageSubQuoteTxBO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required

Private Const cstrTABLE_NAME = "MORTGAGESUBQUOTE"
Private Const cstrLOANCOMPONENT_NAME = "LOANCOMPONENT"


'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'PSC    03/07/2002  BMIDS00062  Amend UpdateLoanComponentXML for new fields
'                               Amend CalcAndSaveMortgageCosts to use composite rates
'                               Amend CreateNewSubQuote to set new fields to zero
'                               Amend SaveIllustrationResults
'                               Amend UpdateLoanComponentXML
'                               Amend SaveOneOffCostDetails
'MV     18/07/2002  BMIDS00179  Core Upgrade Rollback - Modified UpdateMortgageSubQuoteXML()
'DPF    19/07/2002  CMWP3       Added new method IMortgageSubQuoteTxBO_CalcAndSaveRedemptionFees()
'DPF    19/07/2002  CMWP3       Amended method CalcAndSaveMortgageCosts
'DPF    23/07/2002  CMWP3       Added two new fields to SaveLoanComponentXML method
'DPF    08/10/2002  CPWP1       BM037:  Added new step (9) to CalcAndSaveMortgageCosts method
'DPF    09/10/2002  CPWP1       BM020:  Removed code to create mortgage incentives in CreateNewSubQuote
'                               Added 3 new fields to CreateNewSubQuoteXML for Manual Incentives
'                               Also amended SaveOneOffCostDetails & CreateMortgageIncentiveXML
'DPF    11/11/2002  BMIDS00808  Have amended SaveOneOffCosts & UpdateLoanComponentXML to save MonthlyCostsLessDrawdown
'DPF    11/11/2002  BMIDS00335  Have amended CreateNewSubQuote to reset C&I amount to Net C&I amount
'DPF    18/11/2002  BMIDS00802  have amended IMortgageSubQuoteTXBO.CreateNewSubQuote to copy incentives if at offer stage.
'DPF    25/11/2002  BMIDS01066  Have amended CreateNewSubQuoteXML to keep the drawdown value of the last quote
'INR    06/08/2003  BMIDS624    ICC001 UpdateLoanComponentXML change to include RESOLVEDRATE
'PJO    20/10/2003  BMIDS653    Update product reservation
'GHun   20/04/2004  BMIDS736    CORE49 Alpha+ integration
'DRC    09/06/2004  BMIDS767    Reg027 - AdHoc Fee handing
'GHun   08/07/2004  BMIDS782    Amended SaveIllustrationResults to improve performance
'GHun   09/07/2004  BMIDS766    Amended CreateNewSubQuote, CreateNewSubQuoteXML, SaveOneOffCostDetails & CalcAndSaveMortgageCosts
'JD     21/07/04    BMIDS805    check for empty string in SaveOneOffCostDetails
'SR     11/08/2004  BMIDS815    Modified method 'SaveLoanComponentXML' - added 3 new columns of LoanComponent
'JD     24/08/2004  BMIDS854    ensure adhoc APR costs are included in the calc
'JD     13/09/2004  BMIDS854    in createNewSubquote, if we are calling from IssueOffer or remodelKFI we want to copy the addtoloan info for adhoc fees
'HMA    28/09/2004  BMIDS881    Correct SaveOneOffCostDetails
'HMA    07/10/2004  BMIDS881    Save Incentives when RemodelKFI. Save manual incentives for IssueOffer and RemodelKFI
'GHun   11/10/2004  BMIDS881    Changed CalcAndSaveMortgageCosts to include adhoc fees in total chrages
'HMA    21/04/2005  BMIDS977    Save Incentives and AddedToLoan indicator when Remodel.
'-----------------------------------------------------------------------------------------------------------------------------
'MARS Specific History:
'HMA    18/08/2005  MAR28       Add RefundAmount to OneOffCosts in CreateNewMortgageOneOffCostXML
'GHun   12/10/2005  MAR46       Changed SaveLoanComponentXML
'HMA    08/12/2005  MAR836      Correct error reporting in CreateOneOffCosts
'JD     08/03/2006  MAR1061     UpdateMortgageSubQuoteXML - added purchaseprice
'                               CalcAndSaveMortgageCosts, createFirstSubquoteXML - saved PurchasePrice on MSQ
'JD     05/05/2006  MAR1703     CalcAndSaveMortgageCosts - check global parameter to decide if the product is reserved or not
'------------------------------------------------------------------------------------------
'Epsom Specific History:
'
'Prog   Date        AQR         Description
'PSC    09/11/2006  EP2_41      Amend CalcAndSaveMortgageCosts to return one off costs
'PSC    01/12/2006  EP2_249     Amend CalcAndSaveMortgageCosts to calculate and save proc fees
'PSC    14/12/2006  EP2_507     Amend CalcAndSaveMortgageCosts to clone nodes when drawdown present
'MAH    03/01/2007  EP2_444     Added REPAYMENTVEHICLEMONTHLYCOST to SaveLoanComponentDetails
'IK     29/01/2007  EP2_1093    remove CALCFEESONLY test
'AShaw  29/01/2007  EP2_893     Alter code to NOT run CalcAndSaveProcFees if no Data - Revisited.
'PSC    05/02/2007  EP2_1219    Send RecipientId and RcipientType into SaveIntroducerFees
'GHun   12/02/2007  EP2_954     Changed CalcAndSaveMortgageCosts for additional borrowing
'PEdney 19/02/2007  EP2_1281    Do not calculate a valuation fee in one-off costs if no valuation is required.
'SR     20/20/2007  EP2_1546    Modified IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts.
'PEdney 26/02/2007  EP2_1228    Calculate total incentive amount
'PEdney 26/02/2007  EP2_1228    Calculate total incentive amount again
'SR     05/03/2007  EP2_1798    modified 'CalaAndSaeMortgageCosts'. use type double rather than currency
'SR     28/03/207   EP2_1998    modified SaveLoanComponentDetails - remove validation
'                               (now to allow subsequent loan components can be longer than first one).
'------------------------------------------------------------------------------------------

Private Function IMortgageSubQuoteTxBO_CalcAndSaveRedemptionFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSaveRedemptionFeesVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_CalcAndSaveRedemptionFees"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    'Variable Declaration
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim objXMLLoanComponent As IXMLDOMElement
    Dim objLoanCompNode As IXMLDOMNode
    Dim objXmlCalcResults As IXMLDOMElement
    Dim xmlCreateRedemptionReturn As IXMLDOMNode
    Dim objXmlFeeList As IXMLDOMNode
    Dim objXmlFee As IXMLDOMNode

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objCalcsBO As CalcsBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else throw an error
    If vXMLRequest.nodeName = "LOANCOMPONENTLIST" Then
        Set objXMLLoanComponent = vXMLRequest
    Else
        Set objXMLLoanComponent = vXMLRequest.getElementsByTagName("LOANCOMPONENTLIST").Item(0)
    End If
    
    If objXMLLoanComponent Is Nothing Then
         objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, "LOANCOMPONENTLIST tag not found"
    End If
   
   'set up connection to other objects
   If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objCalcsBO = New CalcsBO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objCalcsBO = objContext.CreateInstance(App.Title & ".CalcsBO")
    End If
        
    For Each objLoanCompNode In objXMLLoanComponent.selectNodes("LOANCOMPONENT")
    
        'run redemption fee calculations object
        Set objXmlCalcResults = objCalcsBO.CalculateRedemptionFees(objLoanCompNode)
            
        'If calculations were successful save the record to LOANCOMPONENTREDEMPTIONFEE table
        objErrAssist.CheckResponse objXmlCalcResults, xmlResponseElem, True
            
        'For each LoanComponentRedemptionFee record in returned XML create a record in the table
        Set objXmlFeeList = objXmlCalcResults.selectSingleNode("//LOANCOMPONENTREDEMPTIONFEELIST")
        If Not objXmlFeeList Is Nothing Then
            For Each objXmlFee In objXmlFeeList.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
                objIMortgageSubquoteDO.CreateLoanComponentRedemptionFee objXmlFee
            Next
        End If
    Next

'***  write results back to calling function
    Set IMortgageSubQuoteTxBO_CalcAndSaveRedemptionFees = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalcAndSaveRedemptionFeesExit:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set objCalcsBO = Nothing
    Set objXMLLoanComponent = Nothing
    Set objLoanCompNode = Nothing
    Set objXmlCalcResults = Nothing
    Set xmlCreateRedemptionReturn = Nothing
    Set objXmlFeeList = Nothing
    Set objXmlFee = Nothing
        
    Exit Function

CalcAndSaveRedemptionFeesVbErr:
    
        If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_CalcAndSaveRedemptionFees = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAndSaveRedemptionFeesExit
End Function

Private Function IMortgageSubQuoteTxBO_CreateComponentsFromExistingAcc(ByVal vxmlTableElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode

End Function

Private Function IMortgageSubQuoteTxBO_CreateFirstSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_CreateFirstSubQuote"

    Dim xmlCloneNode As IXMLDOMElement
    
    Set xmlCloneNode = vXMLRequest.cloneNode(True)
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim blnRaiseError As Boolean

    blnRaiseError = True
    
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String

    strInApplicationNumber = objXmlAssist.GetTagValue(xmlCloneNode, "APPLICATIONNUMBER")
    
    strInApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlCloneNode, "APPLICATIONFACTFINDNUMBER")
        
    If Len(strInApplicationNumber) = 0 Or _
       Len(strInApplicationFactFindNumber) = 0 Then
        Dim strDetails As String
        strDetails = "APPLICATIONNUMBER = " & strInApplicationNumber
        strDetails = strDetails & "APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If


    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim xmlTempElem As IXMLDOMElement
    
    Set xmlTempElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlResponseElem.appendChild xmlTempElem
        
    Dim xmlTempElem2 As IXMLDOMElement
    Set xmlTempElem2 = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlTempElem.appendChild xmlTempElem2
    'xmlTempElem2.Text = "1"
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If xmlCloneNode.nodeName = cstrTABLE_NAME Then
        Set xmlElement = xmlCloneNode
    Else
        Set xmlElement = xmlCloneNode.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    ' APS 30/05/2000 - SYS0767 Avoiding duplicate key errors
    ' when calling this method from the GUI
    
    Dim objIDOAssist As IDOAssist
    Dim objIomCMClassDef As IomCMClassDef
    
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objIDOAssist = New DOAssist
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
        
    Set objIomCMClassDef = New omCMClassDef
    
    objIDOAssist.GenerateSequenceNumber xmlElement, _
        objIomCMClassDef.LoadMortgageSubQuoteData, "MORTGAGESUBQUOTENUMBER"
    
    xmlTempElem2.Text = objXmlAssist.GetTagValue(xmlElement, "MORTGAGESUBQUOTENUMBER")
    
    'MS 11/07/00 SYS0924 error creating loan component with a ported loan
    
    Dim strMortgageSubQuoteNumber As String
    strMortgageSubQuoteNumber = xmlTempElem2.Text
    
    Dim xmlMortgageCalculatorDetails As IXMLDOMNode
    Dim lngAmountRequested As Long
    Dim objConvertAssist As ConvertAssist
    
'    If Not xmlElement Is Nothing Then
        
    objXmlAssist.ChangeNodeName xmlElement, "MORTGAGESUBQUOTE", "MORTGAGECALCULATORDETAILS"
    
    blnRaiseError = False
    
    'Record Not found is not an error
    Set xmlMortgageCalculatorDetails = objIMortgageSubquoteDO.GetMortgageCalculatorDetails(xmlElement)
        
    blnRaiseError = True
    
    If Not xmlMortgageCalculatorDetails Is Nothing Then
    
    lngAmountRequested = objConvertAssist.CSafeLng( _
                objXmlAssist.GetTagValue(xmlMortgageCalculatorDetails, "AMOUNTREQUESTED"))
    
    End If
    Dim xmlMortgageSubquote As IXMLDOMElement
    Set xmlMortgageSubquote = CreateFirstSubQuoteXML(xmlElement, lngAmountRequested)
   'thats created the mortgagesubquote record
    objIMortgageSubquoteDO.Create xmlMortgageSubquote
    'now create the loancomponent
    
    Dim xmlLoanComponentList As IXMLDOMNodeList
    
    Set xmlTempElem = xmlCloneNode.getElementsByTagName("LOANSTOBEPORTED").Item(0)
    
    If Not xmlTempElem Is Nothing Then
    'get alist of ported loans
    Set xmlLoanComponentList = xmlTempElem.selectNodes(".//MORTGAGELOAN")
    
    Dim xmlLoanComponent As IXMLDOMElement
    Dim intLoanComponentSequnceNumber As Integer

    'initialise sequncenumber to 1
    intLoanComponentSequnceNumber = 1
    
    For Each xmlElement In xmlLoanComponentList
        
        'MS 11/07/00 SYS0924 error creating loan component with a ported loan
        Set xmlLoanComponent = CreateLoanComponentXML(xmlElement, strInApplicationNumber, _
                                    strInApplicationFactFindNumber, intLoanComponentSequnceNumber, _
                                    strMortgageSubQuoteNumber)
    
            
        intLoanComponentSequnceNumber = intLoanComponentSequnceNumber + 1 'increment seq No.
    
        objIMortgageSubquoteDO.CreateLoanComponent xmlLoanComponent
        
    Next
    
   End If
   
    Set IMortgageSubQuoteTxBO_CreateFirstSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstSubQuoteExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateFirstSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_CreateFirstSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstSubQuoteExit
End Function

Private Function IMortgageSubQuoteTxBO_CreateNewSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_CreateNewSubQuote"

    Dim xmlCloneNode As IXMLDOMElement
    
    Set xmlCloneNode = vXMLRequest.cloneNode(True)
     
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim blnRaiseError As Boolean

    Dim objIDOAssist As IDOAssist
    Dim objIomCMClassDef As IomCMClassDef

    blnRaiseError = True
    
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    Dim strOperation As String                     ' BMIDS881
    Dim xmlTempManualIncentive As IXMLDOMElement   ' BMIDS881
    Dim strCreateSubQuotes As String               ' EP2_56
    
    strInApplicationNumber = m_objXmlAssist.GetTagValue(xmlCloneNode, "APPLICATIONNUMBER")
    
    strInApplicationFactFindNumber = m_objXmlAssist.GetTagValue(xmlCloneNode, "APPLICATIONFACTFINDNUMBER")
    
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(xmlCloneNode, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strInApplicationNumber) = 0 Or _
       Len(strMortgageSubQuoteNumber) = 0 Or _
       Len(strInApplicationFactFindNumber) = 0 Then
        Dim strDetails As String
        strDetails = "APPLICATIONNUMBER = " & strInApplicationNumber
        strDetails = strDetails & "APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber
        strDetails = strDetails & "MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlCopyElem As IXMLDOMElement
    
    Set xmlTempElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlResponseElem.appendChild xmlTempElem
    
    Dim xmlTempElem2 As IXMLDOMElement
    
    Set xmlTempElem2 = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")

    xmlTempElem.appendChild xmlTempElem2
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If xmlCloneNode.nodeName = cstrTABLE_NAME Then
        Set xmlElement = xmlCloneNode
    Else
        Set xmlElement = xmlCloneNode.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objIDOAssist = New DOAssist
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
       
    Dim xmlCurrentMortgageSubQuote As IXMLDOMNode
        
    Set xmlCurrentMortgageSubQuote = objIMortgageSubquoteDO.GetData(xmlElement)
                        
    Dim intNewMortgageSubQuoteNumber As Integer
    
    'BMIDS881 Start
    'Save the manual incentives for IssueOffer and RemodelKFI only. Clear for all other operations.
    'BMIDS977  Save for Remodel too.
    
    strOperation = m_objXmlAssist.GetAttributeValue(vXMLRequest, "REQUEST", "OPERATION")
    'If Operation is null OR
    '   Operation is something other than IssueOffer or RemodelKFI or Remodel then clear incentives.
    
    If (Len(Trim$(strOperation)) = 0) Or _
       ((StrComp(strOperation, "IssueOffer", vbTextCompare) <> 0) And _
        (StrComp(strOperation, "RemodelKFI", vbTextCompare) <> 0) And _
        (StrComp(strOperation, "Remodel", vbTextCompare) <> 0)) Then
      
        Set xmlTempManualIncentive = xmlCurrentMortgageSubQuote.selectSingleNode("//MANUALINCENTIVEAMOUNT")
        If Not xmlTempManualIncentive Is Nothing Then
            xmlTempManualIncentive.Text = "0"
        End If
        
        Set xmlTempManualIncentive = xmlCurrentMortgageSubQuote.selectSingleNode("//MANUALINCENTIVEUSERID")
        If Not xmlTempManualIncentive Is Nothing Then
            xmlTempManualIncentive.Text = ""
        End If
        
        Set xmlTempManualIncentive = xmlCurrentMortgageSubQuote.selectSingleNode("//MANUALINCENTIVEDATETIME")
        If Not xmlTempManualIncentive Is Nothing Then
            xmlTempManualIncentive.Text = ""
        End If
    End If
    'BMIDS881 End
    
    Set objIomCMClassDef = New omCMClassDef
    
    'get the next available mortsubquote number
    
    objIDOAssist.GenerateSequenceNumber xmlCurrentMortgageSubQuote, _
        objIomCMClassDef.LoadMortgageSubQuoteData, "MORTGAGESUBQUOTENUMBER"
    
    intNewMortgageSubQuoteNumber = (CInt(m_objXmlAssist.GetTagValue(xmlCurrentMortgageSubQuote, _
                                    "MORTGAGESUBQUOTENUMBER")))
                                                                                                
    CreateNewSubQuoteXML xmlCurrentMortgageSubQuote, intNewMortgageSubQuoteNumber
   'thats created the mortgagesubquote record
    objIMortgageSubquoteDO.Create xmlCurrentMortgageSubQuote
    
             

    
    'EP2_56 - Passing new param value if we Don't want to create supporting 'fluff'.
    strCreateSubQuotes = m_objXmlAssist.GetTagValue(xmlCloneNode, "CREATESUBQUOTES")
    If strCreateSubQuotes = "" Then    ' NO value passed in - Treat as before.
        
        'now get the loancomponent list
        Dim strTagName As String
       
        Dim xmlTempElement As IXMLDOMElement
       
        m_objXmlAssist.ChangeNodeName xmlCloneNode, "MORTGAGESUBQUOTE", "LOANCOMPONENT"
        
        Set xmlTempElement = xmlCloneNode.getElementsByTagName("LOANCOMPONENT").Item(0)
          
        Set xmlTempElem = objIMortgageSubquoteDO.FindLoanComponentList(xmlTempElement)
                   
        Dim xmlLoanComponentList As IXMLDOMNodeList
            
        Dim strTotalLoanAmount As String
        
        Set xmlLoanComponentList = xmlTempElem.selectNodes(".//LOANCOMPONENT")
            
        For Each xmlElement In xmlLoanComponentList
    
            strTagName = "MORTGAGESUBQUOTENUMBER"
            Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
            xmlTempElem.Text = intNewMortgageSubQuoteNumber
    
            If Not CSafeBool(m_objXmlAssist.GetTagValue(xmlElement, "PORTEDLOAN")) Then
            
                strTagName = "APR"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
        
                strTagName = "GROSSMONTHLYCOST"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
    'FIX ME no IPTAMOUNT on the loan component table
    'strTagName = "IPTAMOUNT"
    'Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
    'xmlTempElem.Text = "0"
        
                strTagName = "MIRASAMOUNT"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
        
                strTagName = "NETMONTHLYCOST"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
                
                strTagName = "LOANAMOUNT"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                strTotalLoanAmount = xmlTempElem.Text
                
                strTagName = "TOTALLOANCOMPONENTAMOUNT"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = strTotalLoanAmount
    
                ' PSC 03/07/2002 BMIDS00062 - Start
                strTagName = "FINALRATE"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
    
                strTagName = "FINALRATEMONTHLYCOST"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
                
                strTagName = "FINALRATEAPR"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = "0"
                ' PSC 03/07/2002 BMIDS00062 - End
                
                '**** DPF 09/10/2002 - CPWP1 (BM037) - Reset value for drawdown values
                strTagName = "MONTHLYCOSTLESSDRAWDOWN"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = ""
                '**** End of CPWP1 ****
                
                '**** DPF 11/11/2002 - BMIDS00335 - Reset C&I amount to Net C&I amount
                strTagName = "CAPITALANDINTERESTELEMENT"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                Set xmlCopyElem = xmlElement.getElementsByTagName("NETCAPANDINTELEMENT").Item(0).cloneNode(True)
                xmlTempElem.Text = xmlCopyElem.Text
                '**** END OF DPF (BMIDS00335)
                         
                'BMIDS766 GHun Clear fields added for Alpha+
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "AMOUNTPERUNITBORROWED")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "CAPINTMONTHLYCOST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "INTONLYMONTHLYCOST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "MINMONTHLYCOST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "MAXMONTHLYCOST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "UNROUNDEDAPR")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "ACCRUEDINTEREST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "CAPINTACCRUEDINTEREST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "INTONLYACCRUEDINTEREST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "INCREASEDMONTHLYCOST")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "INCREASEDMONTHLYCOSTDIFFERENCE")
                xmlTempElem.Text = vbNullString
                
                Set xmlTempElem = xmlGetMandatoryNode(xmlElement, "TOTALAMOUNTPAYABLE")
                xmlTempElem.Text = vbNullString
                'BMIDS766 End
                         
            End If
    
           objIMortgageSubquoteDO.CreateLoanComponent xmlElement
        
        Next
        
        m_objXmlAssist.ChangeNodeName xmlTempElement, "LOANCOMPONENT", "NONPANELMORTGAGEPRODUCT", False
        
        blnRaiseError = False
        'Not an error to have no records
        Set xmlTempElem = objIMortgageSubquoteDO.FindNonPanelMortgageProductList(xmlTempElement)
        blnRaiseError = True
        
        Dim xmlNonPanelMortgageProductList As IXMLDOMNodeList
                
        Set xmlNonPanelMortgageProductList = xmlTempElem.selectNodes(".//NONPANELMORTGAGEPRODUCT")
            
        For Each xmlElement In xmlNonPanelMortgageProductList
            
            strTagName = "MORTGAGESUBQUOTENUMBER"
            Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
            xmlTempElem.Text = intNewMortgageSubQuoteNumber
    
            objIMortgageSubquoteDO.CreateNonPanelMortgageProduct (xmlElement)
            
        Next
        
        '**** DPF 09/10/2002 - CPWP1 (BM020) remove code to set up mortgage incentives ****
        '**** DPF 18/11/2002 - BMIDS00802 - if at Issue Offer stage we want to copy the incentives
        If (StrComp(strOperation, "IssueOffer", vbTextCompare) = 0) Or _
           (StrComp(strOperation, "RemodelKFI", vbTextCompare) = 0) Or _
           (StrComp(strOperation, "Remodel", vbTextCompare) = 0) Then          ' BMIDS881 BMIDS977
     
            m_objXmlAssist.ChangeNodeName xmlTempElement, "NONPANELMORTGAGEPRODUCT", "MORTGAGEINCENTIVE"
            
            blnRaiseError = False
            'Not an error to have no records
            Set xmlTempElem = objIMortgageSubquoteDO.FindMortgageIncentiveList(xmlTempElement)
            'Not an error to have no mortgageincentives
            blnRaiseError = True
            
            Dim xmlMortgageIncentiveList As IXMLDOMNodeList
            Set xmlMortgageIncentiveList = xmlTempElem.selectNodes(".//MORTGAGEINCENTIVE")
                
            For Each xmlElement In xmlMortgageIncentiveList
                strTagName = "MORTGAGESUBQUOTENUMBER"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = intNewMortgageSubQuoteNumber
                objIMortgageSubquoteDO.CreateMortgageIncentive xmlElement
            Next
        
        End If
        
        '**** End of CPWP1 change ****
        
        xmlTempElem2.Text = intNewMortgageSubQuoteNumber
        'BMIDS767 DRC
        'Clone the AdHoc Oneooff Cost details from the old mortgegasubquote
        
        Dim XMLOneOffRequest  As IXMLDOMNode
        Dim xmlTempNode As IXMLDOMNode
        Dim xmlRootNode As IXMLDOMNode
        Dim xmlExistingOneOffCostsDoc As FreeThreadedDOMDocument40
        Dim xmlExistingOneOffCostList As IXMLDOMNodeList
        Dim objIMortgageSubQuoteBO As MortgageSubQuoteBO
        Dim objComboDO As ComboDO
        Dim xmlMortgageCost As IXMLDOMNode
        Dim strResponse As String
        Dim strCostType As String
        Dim strPattern As String
        Dim strValidationType As String
        Set xmlExistingOneOffCostsDoc = New FreeThreadedDOMDocument40
        xmlExistingOneOffCostsDoc.validateOnParse = False
        xmlExistingOneOffCostsDoc.setProperty "NewParser", True
        If m_objContext Is Nothing Then
            Set objComboDO = New ComboDO
            Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Else
            Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
            Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
        End If
        Set XMLOneOffRequest = vXMLRequest.cloneNode(False)
        Set xmlElement = xmlOut.createElement("ONEOFFCOST")
        Set xmlRootNode = xmlElement
        XMLOneOffRequest.appendChild xmlElement
        Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
        xmlElement.Text = strInApplicationNumber
        xmlRootNode.appendChild xmlElement
        Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlElement.Text = strInApplicationFactFindNumber
        xmlRootNode.appendChild xmlElement
        Set xmlElement = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
        xmlElement.Text = strMortgageSubQuoteNumber
        xmlRootNode.appendChild xmlElement
        strResponse = objIMortgageSubQuoteBO.GetOneOffCostsDetails(XMLOneOffRequest.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        
        Set xmlExistingOneOffCostsDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        'Only look for AdHoc Costs
        Set xmlExistingOneOffCostList = xmlExistingOneOffCostsDoc.selectNodes(".//MORTGAGEONEOFFCOST[ADHOCIND = '1']")
        If xmlExistingOneOffCostList.length > 0 Then
        ' Use CreateOneOffCosts
        
        'First need to get Validation types from the combo
            Dim xmlComboDoc As FreeThreadedDOMDocument40
            ' Get combo data for one off costs
            Set xmlRootNode = xmlOut.createElement("LIST")
            Set xmlElement = xmlOut.createElement("LISTNAME")
            xmlElement.Text = "OneOffCost"
            xmlRootNode.appendChild xmlElement
            Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
            strResponse = objComboDO.GetComboList(xmlRootNode.xml)
            Set xmlComboDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    '   Now go through the Existing list cloning & modifing
            Dim xmlMortgageSubquote As IXMLDOMNode
            Set xmlMortgageSubquote = m_objXmlAssist.GetNode(XMLOneOffRequest, ".//MORTGAGESUBQUOTENUMBER")
            xmlMortgageSubquote.Text = intNewMortgageSubQuoteNumber
            Set xmlRootNode = m_objXmlAssist.GetNode(XMLOneOffRequest, ".//ONEOFFCOST")
            m_objXmlAssist.ChangeNodeName xmlRootNode, "ONEOFFCOST", "MORTGAGESUBQUOTE"
            
            For Each xmlMortgageCost In xmlExistingOneOffCostList
              Set xmlElement = xmlOut.createElement("ONEOFFCOST")
              xmlRootNode.appendChild xmlElement
              Set xmlTempNode = m_objXmlAssist.GetNode(xmlMortgageCost, ".//ADDTOLOAN")
              'BMIDS854 BMIDS977 If we are calling from IssueOffer or remodelKFI or Remodel we want the addtoloan info
              If (StrComp(strOperation, "IssueOffer", vbTextCompare) = 0) Or _
                 (StrComp(strOperation, "RemodelKFI", vbTextCompare) = 0) Or _
                 (StrComp(strOperation, "Remodel", vbTextCompare) = 0) Then
                 xmlElement.appendChild xmlTempNode
              Else
                ' Don't add to loan in new subquote
                xmlTempNode.Text = "0"
                xmlElement.appendChild xmlTempNode
              End If
              xmlElement.appendChild m_objXmlAssist.GetNode(xmlMortgageCost, ".//AMOUNT")
              xmlElement.appendChild m_objXmlAssist.GetNode(xmlMortgageCost, ".//ADHOCIND")
              'Need to Use Identifier i.e. Validation type
              Set xmlTempNode = m_objXmlAssist.GetNode(xmlMortgageCost, ".//MORTGAGEONEOFFCOSTTYPE")
              strCostType = xmlTempNode.Text
              strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE[. != 'APR']"
              strValidationType = m_objXmlAssist.GetNodeValue(xmlComboDoc, strPattern)
              xmlTempNode.Text = strValidationType
              m_objXmlAssist.ChangeNodeName xmlTempNode, "MORTGAGEONEOFFCOSTTYPE", "IDENTIFIER"
              xmlElement.appendChild xmlTempNode
            
            Next
            ' Now create the one off costs for the new subquote
            strResponse = objIMortgageSubQuoteBO.CreateOneOffCosts(XMLOneOffRequest.xml)
            m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        End If
        
        
        Set IMortgageSubQuoteTxBO_CreateNewSubQuote = xmlResponseElem.cloneNode(True)
             
    ' EP2_56
    'End DON'T create Incentives, Non Panel data or One off costs if CREATESUBQUOTES = 0.
    Else
        xmlResponseElem.selectSingleNode("//MORTGAGESUBQUOTENUMBER").Text = intNewMortgageSubQuoteNumber
        Set IMortgageSubQuoteTxBO_CreateNewSubQuote = xmlResponseElem.cloneNode(True)
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewSubQuoteExit:

    Set objIMortgageSubquoteDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlCopyElem = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIomCMClassDef = Nothing
    Set objIDOAssist = Nothing
    
    Exit Function

CreateNewSubQuoteVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_CreateNewSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewSubQuoteExit
End Function


Private Function IMortgageSubQuoteTxBO_CreateOneOffCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOneOffCostsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_CreateOneOffCosts"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objIDOAssist As IDOAssist
    Dim objComboDO As ComboDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
        
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    Dim xmlCloneElement As IXMLDOMElement
    'we are going to change the node so clone it to preserve the original
    Set xmlCloneElement = xmlElement.cloneNode(True)
   
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objIDOAssist = New DOAssist
        Set objComboDO = New ComboDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    'get out the key values that we need
    '
    '
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlTempNode  As IXMLDOMNode
    Dim xmlSequenceNumberNode As IXMLDOMNode
    
    Set xmlSequenceNumberNode = xmlDoc.createElement("ONEOFFCOSTSEQUENCENUMBER")

    Dim intSequenceNumber As Integer
       
    Dim xmlCloneNode As IXMLDOMNode
    Dim blnFirst As Boolean
    blnFirst = True
    Set xmlNodeList = vXMLRequest.getElementsByTagName("ONEOFFCOST")
    ' For each relevant node call the data object create
    Dim xmlOneOffCostElement As IXMLDOMElement
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData

        'get the next sequence number and increment it each time
    
    For Each xmlOneOffCostElement In xmlNodeList
        'change the node name for each one off cost
        objXmlAssist.ChangeNodeName xmlOneOffCostElement, "ONEOFFCOST", "MORTGAGEONEOFFCOST"
        'append appno
        Set xmlTempNode = xmlCloneElement.getElementsByTagName("APPLICATIONNUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlOneOffCostElement.appendChild xmlCloneNode
        'app factfind no
        Set xmlTempNode = xmlCloneElement.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlOneOffCostElement.appendChild xmlCloneNode
        'mortgage subquote number
        Set xmlTempNode = xmlCloneElement.getElementsByTagName("MORTGAGESUBQUOTENUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlOneOffCostElement.appendChild xmlCloneNode
        'if first loop round get the next sequence number
        If blnFirst Then
            
            intSequenceNumber = objIDOAssist.GetNextSequenceNumber( _
                                xmlOneOffCostElement, xmlClassDefDoc, _
                                "MORTGAGEONEOFFCOST", "ONEOFFCOSTSEQUENCENUMBER")
            
            blnFirst = False
            
        End If
        'get the mortgageoneoffcosttype
        Set xmlTempNode = xmlOneOffCostElement.getElementsByTagName("IDENTIFIER").Item(0)
        'validate it
        xmlTempNode.Text = objComboDO.GetFirstComboValueId("OneOffCost", xmlTempNode.Text)
        ' if its valid
        If Len(xmlTempNode.Text) <> 0 Then
            'set th sequence number for the record to be created
            xmlSequenceNumberNode.Text = CStr(intSequenceNumber)
            'change it to the correct field name
            objXmlAssist.ChangeNodeName xmlOneOffCostElement, "IDENTIFIER", "MORTGAGEONEOFFCOSTTYPE"
            'add to this one off cost
            xmlOneOffCostElement.appendChild xmlSequenceNumberNode
            'and create it
            objIMortgageSubquoteDO.CreateOneOffCosts xmlOneOffCostElement
            
            'increment the sequnce number
            intSequenceNumber = intSequenceNumber + 1
        Else
        
            '??? is this an error
                        
        End If
    Next
            
    Set IMortgageSubQuoteTxBO_CreateOneOffCosts = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOneOffCostsExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateOneOffCostsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_CreateOneOffCosts = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOneOffCostsExit
End Function
Public Function IMortgageSubQuoteTxBO_DeleteLoanComponent(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_DeleteLoanComponent"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim blnRaiseError As Boolean
    blnRaiseError = True

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"

    End If
   
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    Dim strLoanComponentSequenceNumber As String
    
    strApplicationNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlElement, _
                                                  "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(xmlElement, _
                                                     "MORTGAGESUBQUOTENUMBER")
                                                     
    strLoanComponentSequenceNumber = objXmlAssist.GetTagValue(xmlElement, _
                                                     "LOANCOMPONENTSEQUENCENUMBER")
                                                     
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Or _
       Len(strMortgageSubQuoteNumber) = 0 Or Len(strLoanComponentSequenceNumber) = 0 Then
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & strApplicationNumber
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        strDetails = strDetails & "; LOANCOMPONENTSEQUENCENUMBER = " & strLoanComponentSequenceNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
   
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
           
    Dim xmlTotalLoanAmountNode As IXMLDOMNode
    'need to get the LOANAMOUNT ONLY from the loancomponent
    
    objXmlAssist.ChangeNodeName xmlElement, "MORTGAGESUBQUOTE", "LOANCOMPONENT"
    
    Set xmlTotalLoanAmountNode = objIMortgageSubquoteDO.GetLoanComponentData(xmlElement)
    
    Dim strLoanAmount As String
    Dim lngLoanAmount As Long
    
    strLoanAmount = objXmlAssist.GetTagValue(xmlTotalLoanAmountNode, "LOANAMOUNT")
    
    If Len(strLoanAmount) = 0 Then
    
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, "No Loan Amount for this loan component"
    
    Else
    
        lngLoanAmount = CLng(strLoanAmount)
    
    End If
    
    'Delete this loan component
    
    objIMortgageSubquoteDO.DeleteLoanComponent xmlElement
    blnRaiseError = False
    
    
    'now find non panel mortgage product list
    objXmlAssist.ChangeNodeName vXMLRequest, "LOANCOMPONENT", "NONPANELMORTGAGEPRODUCT"
        
    Set xmlTotalLoanAmountNode = objIMortgageSubquoteDO.FindNonPanelMortgageProductList(vXMLRequest)
    
    blnRaiseError = True
           
    Dim strTagName As String
    Dim xmlTempElement As IXMLDOMElement
       
       
    'update
    Dim xmlMortgageSubQuoteDoc As New FreeThreadedDOMDocument40
    xmlMortgageSubQuoteDoc.validateOnParse = False
    xmlMortgageSubQuoteDoc.setProperty "NewParser", True
    Dim xmlSubQuoteElem As IXMLDOMElement
    
    Set xmlSubQuoteElem = xmlMortgageSubQuoteDoc.createElement("MORTGAGESUBQUOTE")
    xmlMortgageSubQuoteDoc.appendChild xmlSubQuoteElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlMortgageSubQuoteDoc.createElement(strTagName)
    xmlTempElement.Text = strApplicationNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlMortgageSubQuoteDoc.createElement(strTagName)
    xmlTempElement.Text = strApplicationFactFindNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlMortgageSubQuoteDoc.createElement(strTagName)
    xmlTempElement.Text = strMortgageSubQuoteNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlMortgageSubQuoteDoc.createElement(strTagName)
    xmlTempElement.Text = strMortgageSubQuoteNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
    Set xmlTempElement = xmlMortgageSubQuoteDoc.createElement(strTagName)
    xmlTempElement.Text = strLoanComponentSequenceNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    If xmlTotalLoanAmountNode Is Nothing Then
                    
        objIMortgageSubquoteDO.UpdateTotalLoanAmount xmlMortgageSubQuoteDoc.documentElement, _
                                    lngLoanAmount, False
        
    Else
    
        objIMortgageSubquoteDO.UpdateTotalLoanAmount xmlMortgageSubQuoteDoc.documentElement, _
                                    lngLoanAmount
    
    End If
                    
    Set IMortgageSubQuoteTxBO_DeleteLoanComponent = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteLoanComponentExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlMortgageSubQuoteDoc = Nothing
        
    Exit Function

DeleteLoanComponentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_DeleteLoanComponent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLoanComponentExit


End Function

Private Function IMortgageSubQuoteTxBO_GetAPROneOffCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets all mortgageOneOffCosts needed to be included in APR calcs.
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAPROneOffCostsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_GetAPROneOffCosts"
    
    Dim xmlMSQNo As IXMLDOMNode
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFFNo As IXMLDOMNode
    Dim objMortSubQuoteBO As MortgageSubQuoteBO
    Dim objComboDO As ComboDO
    Dim objChargesBO As ChargesBO
    Dim xmlMtgeOneOffCostsDoc As FreeThreadedDOMDocument40
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlMtgeCostList As IXMLDOMNodeList
    Dim xmlMtgeCost As IXMLDOMNode
    Dim xmlCost As IXMLDOMNode
    Dim strCostType As String
    Dim strValidationType As String
    Dim xmlOneOffCosts As IXMLDOMNode
    Dim xmlOneOffCostsList As IXMLDOMNodeList
    Dim strResponse As String
    Dim strPattern As String
    Dim xmlTmpElement As IXMLDOMElement
    Dim blnRaiseError As Boolean
    Dim xmlAPROneOffCosts As IXMLDOMNodeList
    Dim xmlAPROneOffCostsElem As IXMLDOMElement
    Dim xmlOneOffCost As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMElement
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrLOANCOMPONENT_NAME & " tag not found"
    End If
    
    'Copy the node to use locally
    Dim xmlCopyNode As IXMLDOMNode
    Set xmlCopyNode = xmlElement.cloneNode(True)
    
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlCopyNode, "APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXmlAssist.GetNode(xmlCopyNode, "APPLICATIONFACTFINDNUMBER")
    Set xmlMSQNo = m_objXmlAssist.GetNode(xmlCopyNode, "MORTGAGESUBQUOTENUMBER")
    
    Dim strOneOffCosts As String
    Dim xmlOneOffCostsDoc As FreeThreadedDOMDocument40
    Set xmlOneOffCostsDoc = New FreeThreadedDOMDocument40
    xmlOneOffCostsDoc.validateOnParse = False
    xmlOneOffCostsDoc.setProperty "NewParser", True
    
    Set xmlOneOffCosts = m_objXmlAssist.GetNode(xmlCopyNode, ".//ONEOFFCOSTLIST")
    Set xmlOneOffCostsList = xmlOneOffCosts.selectNodes(".//ONEOFFCOST")
    
    Set xmlAPROneOffCostsElem = xmlOneOffCostsDoc.createElement("ONEOFFCOSTLIST")
    xmlResponseElem.appendChild xmlAPROneOffCostsElem
    
    
    'Get all one off costs including adhoc ones
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlRootElement = xmlOut.createElement("ONEOFFCOST")
    xmlTempRequest.appendChild xmlRootElement
    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
    xmlRootElement.appendChild xmlMSQNo.cloneNode(True)
    Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    blnRaiseError = False
    strResponse = objMortSubQuoteBO.GetOneOffCostsDetails(xmlTempRequest.xml)
    Set xmlMtgeOneOffCostsDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlMtgeOneOffCostsDoc.documentElement, xmlResponseElem, True
    blnRaiseError = True

    ' Get combo data for one off costs
    Set xmlRootElement = xmlOut.createElement("LIST")
    Set xmlTmpElement = xmlOut.createElement("LISTNAME")
    xmlTmpElement.Text = "OneOffCost"
    xmlRootElement.appendChild xmlTmpElement
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strResponse = objComboDO.GetComboList(xmlRootElement.xml)
    Set xmlComboDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

    'get a list of all mortgage costs that are to be included in the APR calcs

    strPattern = ".//MORTGAGEONEOFFCOST"

    Set xmlMtgeCostList = xmlMtgeOneOffCostsDoc.selectNodes(strPattern)
    If Not xmlMtgeCostList Is Nothing Then
    For Each xmlMtgeCost In xmlMtgeCostList

        ' Convert the cost type to the validation type
        strCostType = m_objXmlAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE[. != 'APR']"
        strValidationType = m_objXmlAssist.GetNodeValue(xmlComboDoc, strPattern)


        ' Get the calculated one off cost with the same validation type and copy the parameters over
            strPattern = ".//ONEOFFCOST[IDENTIFIER = '" & strValidationType & "']"
            Set xmlCost = m_objXmlAssist.GetNode(xmlOneOffCosts, strPattern, False)
            If Not (xmlCost Is Nothing) Then

            'if it's adhoc, we need to create a new OneOffCostElement
                If (m_objXmlAssist.GetNodeValue(xmlMtgeCost, "ADHOCIND", True, False) = "1") Then

                    Set xmlCost = xmlOneOffCostsDoc.createElement("ONEOFFCOST")
                    m_objXmlAssist.CopyNode "ADDTOLOAN", "ADDTOLOAN", xmlMtgeCost, xmlCost, False
                    m_objXmlAssist.CopyNode "ADHOCIND", "ADHOCIND", xmlMtgeCost, xmlCost, False
                    m_objXmlAssist.CopyNode "AMOUNT", "AMOUNT", xmlMtgeCost, xmlCost, False
                    Set xmlTempNode = xmlOneOffCostsDoc.createElement("IDENTIFIER")
                    xmlTempNode.Text = strValidationType
                    xmlCost.appendChild xmlTempNode
                    xmlAPROneOffCostsElem.appendChild xmlCost
                Else
                'if its a calculated cost, copy over the Add to Loan filed from the MortgageOneOff costs record
                    m_objXmlAssist.CopyNode "ADDTOLOAN", "ADDTOLOAN", xmlMtgeCost, xmlCost, False

                End If

            End If
    Next
    End If
   
    Set IMortgageSubQuoteTxBO_GetAPROneOffCosts = xmlResponseElem
    
    m_objContext.SetComplete
    
GetAPROneOffCostsExit:
    
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objMortSubQuoteBO = Nothing
    Set objComboDO = Nothing
    Set objChargesBO = Nothing
    Set xmlMtgeOneOffCostsDoc = Nothing
    Set xmlComboDoc = Nothing
    Set xmlOut = Nothing
    
    Exit Function

GetAPROneOffCostsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    'check for record not found
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_GetAPROneOffCosts = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume GetAPROneOffCostsExit
End Function

Private Function IMortgageSubQuoteTxBO_ReserveMortgageProduct(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ReserveMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_ReserveMortgageProduct"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
        End If
    End If
   
   
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageProductCode As String
    Dim strStartDate As String
    Dim strPortedLoan As String
    
    strApplicationNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
    'MAR1703 no need to check for product code now - only applicationnumber and factfind number are required.
    ' strMortgageProductCode = objXmlAssist.GetTagValue(xmlElement, _
    '                                    "MORTGAGEPRODUCTCODE")
    ' strStartDate = objXmlAssist.GetTagValue(xmlElement, _
    '                                                 "STARTDATE")
                                                     
    'strPortedLoan = objXmlAssist.GetTagValue(xmlElement, _
                                                     "PORTEDLOAN")
                                                     
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Then
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & strApplicationNumber
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        'strDetails = strDetails & "; STARTDATE = " & strStartDate
        'strDetails = strDetails & "; PORTEDLOAN = " & strPortedLoan
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, strDetails
    End If

   'If CBool(strPortedLoan) Then
   '
   '     objErrAssist.ThrowError TypeName(Me), strFunctionName, 189
   '
   'End If
   
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                    
    objXmlAssist.ChangeNodeName xmlElement, "MORTGAGESUBQUOTE", "RESERVEDPRODUCT"
                    
    Dim blnRaiseError As Boolean
                    
    blnRaiseError = False
                    
    objIMortgageSubquoteDO.ReserveMortgageProduct xmlElement
    
    Set IMortgageSubQuoteTxBO_ReserveMortgageProduct = xmlResponseElem.cloneNode(True)
    
    'A duplicate key is ok i.e. this product has already been reserved
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ReserveMortgageProductExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ReserveMortgageProductVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeDuplicateKey Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_ReserveMortgageProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReserveMortgageProductExit

End Function

Private Function IMortgageSubQuoteTxBO_SaveIncentives(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveIncentivesVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_SaveIncentives"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName(cstrTABLE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIMortgageSubquoteDO.SaveIncentiveAgainstLoanComponent xmlElement
    Else
        ' For each relevant node call the data object create
        For Each xmlElement In xmlNodeList
            objIMortgageSubquoteDO.SaveIncentiveAgainstLoanComponent xmlElement
        Next
    End If
    
    Set IMortgageSubQuoteTxBO_SaveIncentives = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveIncentivesExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveIncentivesVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_SaveIncentives = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveIncentivesExit
End Function

Private Function CreateFirstSubQuoteXML(ByVal xmlElement As IXMLDOMElement, _
                        ByVal lngAmountRequested As Long) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteXMLVbErr

    Const strFunctionName As String = "CreateFirstSubQuoteXML"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strDetails As String
    
    Dim strAppNo As String 'MAR1061
    Dim strAppFFNo As String 'MAR1061
    Dim xmlAppDoc As New FreeThreadedDOMDocument40 'MAR1061
    Dim xmlAppReqElem As IXMLDOMElement 'MAR1061
    Dim xmlAppAppElem As IXMLDOMElement 'MAR1061
    Dim xmlAppTmpNode As IXMLDOMNode 'MAR1061
    Dim xmlAppRespDoc As New FreeThreadedDOMDocument40 'MAR1061
    Dim xmlAppRespElem As IXMLDOMElement 'MAR1061
    Dim objApplicationBO As Object 'MAR1061
    
    If IsMissing(lngAmountRequested) Then
        
        strDetails = "lngAmountRequested = " & lngAmountRequested
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
        
    End If
    
    
    Dim strTagName As String
    Dim strText As String
        
    Set xmlSubQuoteElem = xmlOut.createElement(cstrTABLE_NAME)
    xmlOut.appendChild xmlSubQuoteElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    strAppNo = strText
    If Len(strText) = 0 Then
    
        strDetails = "APPLICATIONNUMBER = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
        
    End If
        
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement
    

    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    strAppFFNo = strText
    If Len(strText) = 0 Then

        strDetails = "APPLICATIONFACTFINDNUMBER = " & strText
    
        objErrAssist.ThrowError TypeName(Me), _
                            strFunctionName, _
                            oeInvalidParameter, _
                            strDetails
    
    End If

    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "AMOUNTREQUESTED"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = Str(lngAmountRequested)
    xmlSubQuoteElem.appendChild xmlTempElement
        
    strTagName = "DATEANDTIMEGENERATED"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    xmlSubQuoteElem.appendChild xmlTempElement
   
    strTagName = "DEPOSIT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
   
    strTagName = "LTV"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
   
    strTagName = "NONPANELLENDERSELECTD"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0" 'FALSE
    xmlSubQuoteElem.appendChild xmlTempElement
   
    strTagName = "QUOTATIONTYPE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then

    strDetails = "QUOTATIONTYPE = " & strText
    
    objErrAssist.ThrowError TypeName(Me), _
                            strFunctionName, _
                            oeInvalidParameter, _
                            strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TOTALCHARGES"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TOTALGROSSMONTHLYCOST"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TOTALINITIALCHARGES"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TOTALLOANAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "TOTALNETCHARGES"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "TOTALNETMONTHLYCOST"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "0"
    xmlSubQuoteElem.appendChild xmlTempElement
    
    'MAR1061 Get the purchaseprice from ApplicationFactFind
    Set xmlAppReqElem = xmlAppDoc.createElement("REQUEST")
    xmlAppDoc.appendChild xmlAppReqElem
    Set xmlAppAppElem = xmlAppReqElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFIND"))
    Set xmlAppTmpNode = xmlAppAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONNUMBER"))
    xmlAppTmpNode.Text = strAppNo
    Set xmlAppTmpNode = xmlAppAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlAppTmpNode.Text = strAppFFNo
        
    Set objApplicationBO = m_objContext.CreateInstance("omApp.ApplicationBO")
    xmlAppRespDoc.loadXML objApplicationBO.GetApplicationFactFindData(xmlAppDoc.documentElement.xml)
    Set xmlAppRespElem = xmlAppRespDoc.selectSingleNode("RESPONSE/APPLICATIONFACTFIND")
    If Not xmlAppRespElem Is Nothing Then
        strText = objXmlAssist.GetTagValue(xmlAppRespElem, "PURCHASEPRICEORESTIMATEDVALUE")
        strTagName = "PURCHASEPRICEORESTIMATEDVALUE"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = strText
        xmlSubQuoteElem.appendChild xmlTempElement
    End If
      
    Set CreateFirstSubQuoteXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing

    Exit Function

CreateFirstSubQuoteXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function CreateLoanComponentXML(ByVal xmlElement As IXMLDOMElement, _
                        ByVal strInApplicationNumber As String, _
                        ByVal strInApplicationFactFindNumber As String, _
                        ByVal intLoanComponentSequenceNumber As Integer, _
                        ByVal strMortgageSubQuoteNumber As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanComponentXMLVbErr

    Const strFunctionName As String = "CreateLoanComponentXML"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objComboDO As ComboDO

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
    Dim strText As String
    Dim strDetails As String
    
    If IsMissing(strInApplicationNumber) Or _
        IsMissing(strInApplicationFactFindNumber) Or _
        IsMissing(intLoanComponentSequenceNumber) Or _
        IsMissing(strMortgageSubQuoteNumber) Or strMortgageSubQuoteNumber = "0" Then
        
        strDetails = "strInApplicationNumber = " & strInApplicationNumber
        strDetails = strDetails & "strInApplicationFactFindNumber = " & strInApplicationFactFindNumber
        strDetails = strDetails & "intLoanComponentSequenceNumber = " & intLoanComponentSequenceNumber
        strDetails = strDetails & "strMortgageSubQuoteNumber = " & strMortgageSubQuoteNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails

        
    End If
    
    Set xmlSubQuoteElem = xmlOut.createElement("LOANCOMPONENT")
    xmlOut.appendChild xmlSubQuoteElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strInApplicationNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strInApplicationFactFindNumber
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    'MS 11/07/00 SYS0924 error creating loan component with a ported loan
    xmlTempElement.Text = strMortgageSubQuoteNumber
'    xmlTempElement.Text = "1"
    
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = intLoanComponentSequenceNumber
    xmlSubQuoteElem.appendChild xmlTempElement
        
    strTagName = "MORTGAGEPRODUCTCODE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "999999"
    xmlSubQuoteElem.appendChild xmlTempElement
        
    Dim strCapitalAndInterestElement As String
    Dim strInterestOnlyElement As String
    Dim strOutstandingBalance As String
    Dim strOriginalPartAndPartIntOnlyAmt As String
        
    strText = objXmlAssist.GetTagValue(xmlElement, "REPAYMENTTYPE")
    
    If Len(strText) = 0 Then
    
        strDetails = "REPAYMENTTYPE = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails

    End If
    
    'do combo validation here

    Dim blnPartAndPart As Boolean
            
    blnPartAndPart = objComboDO.IsItemInValidation("RepaymentType", strText, "P")
    
    If blnPartAndPart Then
        
        strOutstandingBalance = objXmlAssist.GetTagValue(xmlElement, "OUTSTANDINGBALANCE")
    
        strOriginalPartAndPartIntOnlyAmt = objXmlAssist.GetTagValue(xmlElement, "ORIGINALPARTANDPARTINTONLYAMT")
    
        If Len(strOutstandingBalance) = 0 Or _
            Len(strOriginalPartAndPartIntOnlyAmt) = 0 Then
        
            strDetails = "OUTSTANDINGBALANCE = " & strOutstandingBalance
            strDetails = strDetails & "ORIGINALPARTANDPARTINTONLYAMT = " & strOriginalPartAndPartIntOnlyAmt
            
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeInvalidParameter, _
                                    strDetails
        
        End If
        
        strInterestOnlyElement = strOriginalPartAndPartIntOnlyAmt
                
        strCapitalAndInterestElement = CStr(CLng(strOutstandingBalance) - CLng(strInterestOnlyElement))
    
    Else
        
        strCapitalAndInterestElement = "NULL"
        
        strInterestOnlyElement = "NULL"
    
    End If

    strTagName = "CAPITALANDINTERESTELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strCapitalAndInterestElement
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "INTERESTONLYELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strInterestOnlyElement
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "LOANAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strOutstandingBalance
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "MIRASAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
'MS 05/07/00 Not an error
'    If Len(strText) = 0 Then
'
'        strDetails = "MIRASAMOUNT = " & strText
'
'        objErrAssist.ThrowError TypeName(Me), _
'                                strFunctionName, _
'                                oeInvalidParameter, _
'                                strDetails
'
'    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "MIRASELIGIBILITY"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
'MS 05/07/00 Not an error
'    If Len(strText) = 0 Then
'
'        strDetails = "MIRASELIGIBILITY = " & strText
'
'        objErrAssist.ThrowError TypeName(Me), _
'                                strFunctionName, _
'                                oeInvalidParameter, _
'                                strDetails
'
'    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement
   
    strTagName = "PORTEDLOAN"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = "1" 'TRUE
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "PURPOSEOFLOAN"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "PURPOSEOFLOAN = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "REPAYMENTMETHOD"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strText = objXmlAssist.GetTagValue(xmlElement, "REPAYMENTTYPE")
    
    If Len(strText) = 0 Then
        
        strDetails = "REPAYMENTMETHOD = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails

    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement

    Dim strTermInMonths As String
    Dim strTermInYears As String
    Dim strOrigTermInMths As String
    Dim strOrigTermInYears As String
    Dim strStartDate As String
    Dim lngOustandingTermInMths As Long
    Dim intOrigTermInMths As Integer
    Dim dtEndDate As Date

    strOrigTermInYears = objXmlAssist.GetTagValue(xmlElement, "ORIGINALTERMYEARS")
    strOrigTermInMths = objXmlAssist.GetTagValue(xmlElement, "ORIGINALTERMMONTHS")
    strStartDate = objXmlAssist.GetTagValue(xmlElement, "STARTDATE")
    
    If Len(strOrigTermInYears) = 0 Or _
        Len(strOrigTermInMths) = 0 Or _
        Len(strStartDate) = 0 Then
    
        strDetails = "ORIGINALTERMYEARS = " & strOrigTermInYears
        strDetails = strDetails & "ORIGINALTERMMONTHS = " & strOrigTermInMths
        strDetails = strDetails & "STARTDATE = " & strStartDate

        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
       
    End If
    
    intOrigTermInMths = CInt(strOrigTermInYears * 12) + CInt(strOrigTermInMths)
    
    dtEndDate = DateAdd("m", intOrigTermInMths, strStartDate)
        
    lngOustandingTermInMths = DateDiff("m", Now, dtEndDate)
        
    strTermInYears = CStr(CInt(lngOustandingTermInMths / 12))
    
    strTermInMonths = CStr(lngOustandingTermInMths Mod 12)
    

    If (lngOustandingTermInMths) <= 0 Then
    'should never get here but just in case
    'ported loan expired - raise an error
        objErrAssist.ThrowError TypeName(Me), _
                        strFunctionName, _
                        oeInvalidParameter, _
                        "Ported Loan already expired"

    
    End If
    
    strTagName = "TERMINMONTHS"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strTermInMonths
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TERMINYEARS"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strTermInYears
    xmlSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "TOTALLOANCOMPONENTAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strOutstandingBalance
    xmlSubQuoteElem.appendChild xmlTempElement

    strTagName = "MORTGAGELOANGUID"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strText = objXmlAssist.GetTagValue(xmlElement, "MORTGAGELOANGUID")
    If Len(strText) = 0 Then
    
        strDetails = "MORTGAGELOANGUID = " & strText

        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails

    End If
    
    xmlTempElement.Text = strText
    xmlSubQuoteElem.appendChild xmlTempElement
      
    Set CreateLoanComponentXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set objComboDO = Nothing

    Exit Function

CreateLoanComponentXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set objComboDO = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub CreateNewSubQuoteXML(ByRef xmlElement As IXMLDOMElement, _
                                ByVal intNewMortgageSubQuoteNumber As Integer)
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteXMLVbErr

    Const strFunctionName As String = "CreateNewSubQuoteXML"

    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
                
    If IsMissing(intNewMortgageSubQuoteNumber) Then
                        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "intNewMortgageSubQuoteNumber = " & intNewMortgageSubQuoteNumber
        
    End If
                
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = intNewMortgageSubQuoteNumber

    strTagName = "DATEANDTIMEGENERATED"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = Format$(Now, "dd/mm/yyyy hh:mm:ss")
        
    strTagName = "TOTALCHARGES"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = "0"
    
    strTagName = "TOTALGROSSMONTHLYCOST"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = "0"
    
    strTagName = "TOTALINITIALCHARGES"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = "0"

    Dim strAmountRequested As String
    
    strTagName = "AMOUNTREQUESTED"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    strAmountRequested = xmlTempElement.Text

    strTagName = "TOTALLOANAMOUNT"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = strAmountRequested

    strTagName = "TOTALNETCHARGES"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = "0"

    strTagName = "TOTALNETMONTHLYCOST"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = "0"
    
    '**** DPF 09/10/2002 - CPWP1 (BM020) - Set Manual Incentive fields to null values ****
    'BMIDS881 Keep manual incentive values - they are set to null when required in CreateNewSubQuote
'    strTagName = "MANUALINCENTIVEAMOUNT"
'    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
'    xmlTempElement.Text = ""
'
'    strTagName = "MANUALINCENTIVEUSERID"
'    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
'    xmlTempElement.Text = ""
'
'    strTagName = "MANUALINCENTIVEDATETIME"
'    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
'    xmlTempElement.Text = ""
        
    '**** DPF 09/10/2002 - CPWP1 (BM037) - Reset value for drawdown values
    strTagName = "MONTHLYCOSTLESSDRAWDOWN"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    xmlTempElement.Text = ""
    
    strTagName = "DRAWDOWN"
    Set xmlTempElement = xmlElement.getElementsByTagName(strTagName).Item(0)
    'BMIDS01066 - DPF 25/11/2002 - Shouldn't be blanking this amount
    'xmlTempElement.Text = ""
    '**** End of CPWP1 Changes ****
    
    'BMIDS766 GHun Clear TotalAccruedInterest, AmountPerUnitBorrowed, APR,
    '               TotalAmountPayable and TotalMortgagePayments
    Set xmlTempElement = xmlGetMandatoryNode(xmlElement, "TOTALACCRUEDINTEREST")
    xmlTempElement.Text = vbNullString
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlElement, "AMOUNTPERUNITBORROWED")
    xmlTempElement.Text = vbNullString
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlElement, "APR")
    xmlTempElement.Text = vbNullString
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlElement, "TOTALAMOUNTPAYABLE")
    xmlTempElement.Text = vbNullString
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlElement, "TOTALMORTGAGEPAYMENTS")
    xmlTempElement.Text = vbNullString
    'BMIDS766 End
    
    Set xmlTempElement = Nothing

    Exit Sub

CreateNewSubQuoteXMLVbErr:
       
    Set xmlTempElement = Nothing
       
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
        
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function IMortgageSubQuoteTxBO_SaveLoanComponentDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveLoanComponentDetailsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_SaveLoanComponentDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim blnRaiseError As Boolean
    Dim blnNewLoanComponentDetails As Boolean

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objComboDO As ComboDO
    Dim objIMortgageProductDO As IMortgageProductDO
    Dim objIDOAssist As IDOAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then

        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"

    End If
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objComboDO = New ComboDO
        Set objIMortgageProductDO = New MortgageProductDO
        Set objIDOAssist = New DOAssist
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objGlobalParameterDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objIMortgageProductDO = objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductDO")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    blnRaiseError = True
    
    Dim objConvertAssist As New ConvertAssist
         
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlLoanComponentList As IXMLDOMNodeList

    'Validate loan component amount
    
    Dim lngAmountRequested As Long
    
    lngAmountRequested = objIMortgageSubquoteDO.GetAmountRequested(xmlElement)
    
    objXmlAssist.ChangeNodeName xmlElement, "MORTGAGESUBQUOTE", "LOANCOMPONENT"
        
    Dim intLoanComponentSequenceNumber As Integer
        
    intLoanComponentSequenceNumber = objConvertAssist.CSafeInt( _
            objXmlAssist.GetTagValue(xmlElement, "LOANCOMPONENTSEQUENCENUMBER"))

    blnNewLoanComponentDetails = False
    
    If (intLoanComponentSequenceNumber = 0) Then
        
        blnNewLoanComponentDetails = True
    
    End If
    
    blnRaiseError = False
    Set xmlLoanComponent = objIMortgageSubquoteDO.FindLoanComponentList(xmlElement)
    'NO ERROR ON RECORD NOT FOUND !!!
    blnRaiseError = True
        
    If xmlLoanComponent Is Nothing Then
            
        If blnNewLoanComponentDetails = False Then
        
                objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                        oeMissingPrimaryTag, _
                        " Sequence number supplied but record not found on database"

        
        End If
    
        blnNewLoanComponentDetails = True
            
    End If
    
    Dim lngTotalLoans As Long
        
    lngTotalLoans = 0
    
    Dim xmlLoanComponentElement As IXMLDOMElement
    
    If Not xmlLoanComponent Is Nothing Then
    
'        Dim intLoanComponentSequenceNumber As Integer
        If Not blnNewLoanComponentDetails Then
            'Get the loan component sequnce number out
            intLoanComponentSequenceNumber = objConvertAssist.CSafeInt( _
                                    objXmlAssist.GetTagValue(xmlElement, "LOANCOMPONENTSEQUENCENUMBER"))
        Else
            intLoanComponentSequenceNumber = 0
        End If
        
        Dim strPattern As String
        strPattern = ".//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER != " & CStr(intLoanComponentSequenceNumber) & "]"
        Set xmlLoanComponentList = xmlLoanComponent.selectNodes(strPattern)
                
        For Each xmlLoanComponentElement In xmlLoanComponentList
           lngTotalLoans = lngTotalLoans + objConvertAssist.CSafeLng( _
                     objXmlAssist.GetTagValue(xmlLoanComponentElement, "LOANAMOUNT"))
        Next
        
        If ((lngTotalLoans + objConvertAssist.CSafeLng( _
                                objXmlAssist.GetTagValue(xmlElement, "LOANAMOUNT"))) _
                                > lngAmountRequested) Then
        
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 196
    
        End If
    
    End If
        
    'Validate term
    Dim intMinimumTermMonths  As Integer
    Dim intMaximumTermMonths  As Integer
    Dim intTermInYears As Integer
    Dim intTermInMonths As Integer
    Dim intTerm As Integer

        
    intMinimumTermMonths = (objConvertAssist.CSafeInt(objGlobalParameterDO.GetCurrentParameterByType _
                                                ("MinimumTerm", "AMOUNT")) * 12)
    
    intMaximumTermMonths = (objConvertAssist.CSafeInt(objGlobalParameterDO.GetCurrentParameterByType _
                                                ("MaximumTerm", "AMOUNT")) * 12)
    
    intTermInYears = objConvertAssist.CSafeInt(objXmlAssist.GetTagValue(xmlElement, "TERMINYEARS"))
    
    intTermInMonths = objConvertAssist.CSafeInt(objXmlAssist.GetTagValue(xmlElement, "TERMINMONTHS"))
    
    intTerm = ((intTermInYears * 12) + intTermInMonths)
    
    If intTerm < intMinimumTermMonths Or intTerm > intMaximumTermMonths Then
    
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 130
    
    End If
'SR 28/03/207 : EP2_1998 - remove validation (now to allow subsequent loan components can be longer than first one).
'    If Not blnNewLoanComponentDetails Then
'        'Get all the loan components where portedloan = FALSE
'        Set xmlLoanComponentList = xmlLoanComponent.selectNodes(".//LOANCOMPONENT[PORTEDLOAN = 0 ]")
'        'and loop round all finding the FIRST Loan Component
'        Dim intLowestSequenceNumber As Integer
'        Dim intTempLowestSequenceNumber As Integer
'        Dim xmlFirstLoanComponent As IXMLDOMElement
'        Dim intFirstTerm As Integer
'
'        intLowestSequenceNumber = 0
'
'        For Each xmlLoanComponentElement In xmlLoanComponentList
'
'            intTempLowestSequenceNumber = objConvertAssist.CSafeInt( _
'                    objXmlAssist.GetTagValue(xmlLoanComponentElement, "LOANCOMPONENTSEQUENCENUMBER"))
'
'            If intTempLowestSequenceNumber < intLowestSequenceNumber Or intLowestSequenceNumber = 0 Then
'
'                intLowestSequenceNumber = intTempLowestSequenceNumber
'
'                Set xmlFirstLoanComponent = xmlLoanComponentElement
'
'            End If
'
'        Next
'
'        If Not xmlFirstLoanComponent Is Nothing And intLowestSequenceNumber <> intLoanComponentSequenceNumber Then
'
'            intFirstTerm = ((objConvertAssist.CSafeInt( _
'                            objXmlAssist.GetTagValue(xmlFirstLoanComponent, "TERMINYEARS") * 12)) _
'                            + (objConvertAssist.CSafeInt( _
'                            objXmlAssist.GetTagValue(xmlFirstLoanComponent, "TERMINMONTHS"))))
'
'            If intTerm > intFirstTerm Then
'
'                objErrAssist.ThrowError TypeName(Me), strFunctionName, 238
'
'            End If
'
'        End If
'
'    End If
'SR 28/03/207 : EP2_1998 - End
    'Validate Repayment Type and Part And Part Split
    
    If Not blnNewLoanComponentDetails Then
        
        Dim strRepaymentType As String
        Dim intNumberOfComponents As Integer
        Dim lCapitalAndInterestAmount As Long
        Dim lInterestOnlyAmount As Long
        Dim intInLoanComponentSequnceNumber As Integer
        
        strRepaymentType = objXmlAssist.GetTagValue(xmlElement, "REPAYMENTTYPE")
    
        If objComboDO.IsItemInValidation("RepaymentType", strRepaymentType, "P") Then
        
            Set xmlLoanComponentList = xmlLoanComponent.selectNodes(".//LOANCOMPONENT")
            
            intNumberOfComponents = xmlLoanComponentList.length
             
            intInLoanComponentSequnceNumber = objConvertAssist.CSafeInt( _
                    objXmlAssist.GetTagValue(xmlElement, "LOANCOMPONENTSEQUENCENUMBER"))
             
            
            If intNumberOfComponents > 1 Or (intNumberOfComponents = 1 And _
                                                intInLoanComponentSequnceNumber = 0) Then
        
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 197
            
            End If
            
            lCapitalAndInterestAmount = objConvertAssist.CSafeLng( _
                        objXmlAssist.GetTagValue(xmlElement, "CAPITALANDINTERESTELEMENT"))
             
            lInterestOnlyAmount = objConvertAssist.CSafeLng( _
                        objXmlAssist.GetTagValue(xmlElement, "INTERESTONLYELEMENT"))
            
            If lCapitalAndInterestAmount = 0 Or lInterestOnlyAmount = 0 Then
            
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 172
            
            End If
            
        End If
    
    End If
    
    'Validate Non-Panel Mortgage Product Details ( if appropriate)
    Dim xmlMortgageProduct As IXMLDOMElement
    
    objXmlAssist.ChangeNodeName xmlElement, "LOANCOMPONENT", "MORTGAGEPRODUCT"
    
    Set xmlMortgageProduct = objIMortgageProductDO.GetData(xmlElement)
    
    Dim blnNonPanelLenderOption As Boolean
    
    blnNonPanelLenderOption = False
    
    blnNonPanelLenderOption = objConvertAssist.CSafeBool(objXmlAssist.GetTagValue( _
                                    xmlMortgageProduct, "NONPANELLENDEROPTION"))
    
    If blnNonPanelLenderOption Then
    
        Dim strLenderName As String
        Dim strMortgageProductName As String
        Dim strMortgageProductType As String
        Dim strInterestRate1 As String
        Dim strInterestRate1Period As String
        
        strLenderName = objXmlAssist.GetTagValue(xmlElement, "LENDERNAME")
        strMortgageProductName = objXmlAssist.GetTagValue(xmlElement, "MORTGAGEPRODUCTNAME")
        strMortgageProductType = objXmlAssist.GetTagValue(xmlElement, "MORTGAGEPRODUCTTYPE")
        strInterestRate1 = objXmlAssist.GetTagValue(xmlElement, "INTERESTRATE1")
        strInterestRate1Period = objXmlAssist.GetTagValue(xmlElement, "INTERESTRATE1PERIOD")
                                                         
        If Len(strLenderName) = 0 Or Len(strMortgageProductName) = 0 Or _
           Len(strMortgageProductType) = 0 Or Len(strInterestRate1) = 0 Or _
           Len(strInterestRate1Period) = 0 Then
    
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 198
            
        End If
        
    End If
    
    'Validate Product Consistency
    Dim strMortgageProductCode As String
    Dim strIncentivesProductCode As String
    Dim strStartDate As String
    Dim strIncentivesProductStartDate As String
    
    strMortgageProductCode = objXmlAssist.GetTagValue(xmlElement, "MORTGAGEPRODUCTCODE")
    strStartDate = objXmlAssist.GetTagValue(xmlElement, "STARTDATE")
    strIncentivesProductCode = objXmlAssist.GetTagValue(xmlElement, "INCENTIVESPRODUCTCODE")
    strIncentivesProductStartDate = objXmlAssist.GetTagValue(xmlElement, "INCENTIVESPRODUCTSTARTDATE")
    
    If Not StrComp(strMortgageProductCode, strIncentivesProductCode, vbTextCompare) = 0 Or _
        Not StrComp(strStartDate, strIncentivesProductStartDate, vbTextCompare) = 0 Then
    
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 223
    
    End If
    Dim xmlLoanComponentUpdate As IXMLDOMElement
    'Save Data
    Dim xmlTempElement As IXMLDOMElement
    Dim strNewLoanComponentSequenceNumber As String
    If intLoanComponentSequenceNumber <> 0 Then
        'editing an existing loan component
        Set xmlLoanComponentUpdate = SaveLoanComponentXML(xmlElement)
    
        'need to return the loancomponentsequencenumber
    
        Set xmlTempElement = xmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER")
        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlLoanComponentUpdate, "LOANCOMPONENTSEQUENCENUMBER")
        xmlResponseElem.appendChild xmlTempElement

        objIMortgageSubquoteDO.UpdateLoanComponent xmlLoanComponentUpdate
    
    Else
        If Not blnNewLoanComponentDetails Then
            
            strNewLoanComponentSequenceNumber = "1"
        
        Else
            
            Dim objIomCMClassDef As IomCMClassDef
            Set objIomCMClassDef = New omCMClassDef
            'Calc the next sequence number

            objXmlAssist.ChangeNodeName xmlElement, "MORTGAGEPRODUCT", "LOANCOMPONENT"
            
            objIDOAssist.GenerateSequenceNumber _
                xmlElement, objIomCMClassDef.LoadLoanComponentData, "LOANCOMPONENTSEQUENCENUMBER"
            
            Set objIomCMClassDef = Nothing
            
        End If
    
        Set xmlLoanComponentUpdate = SaveLoanComponentXML(xmlElement, blnNewLoanComponentDetails)  'JLD set the 'new' flag
    
        Set xmlTempElement = xmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER")
        strNewLoanComponentSequenceNumber = objXmlAssist.GetTagValue(xmlLoanComponentUpdate, "LOANCOMPONENTSEQUENCENUMBER")
        xmlTempElement.Text = strNewLoanComponentSequenceNumber
        xmlResponseElem.appendChild xmlTempElement

    
        objIMortgageSubquoteDO.CreateLoanComponent xmlLoanComponentUpdate
        
    End If
    
    'Now do non-panel details
    
    If Not blnNewLoanComponentDetails Then 'New loan comp so Nothing to delete
        
        'Delete existing Non-panel mortgage products if we have any
        
        objXmlAssist.ChangeNodeName xmlElement, "MORTGAGEPRODUCT", "NONPANELMORTGAGEPRODUCT"
            
        blnRaiseError = False

        objIMortgageSubquoteDO.DeleteNonPanelMortgageProduct xmlElement
        
        blnRaiseError = True
        
    End If
    
    If blnNonPanelLenderOption Then
    
        Dim xmlNonPanelMortgageProduct As IXMLDOMElement
        
        xmlNonPanelMortgageProduct = CreateNonPanelMortgageProductXML(xmlElement)
                    
        objIMortgageSubquoteDO.CreateNonPanelMortgageProduct (xmlNonPanelMortgageProduct)
            
        
    End If
    
    Dim xmlIncentives As IXMLDOMElement
    
    Set xmlIncentives = xmlElement.getElementsByTagName("INCENTIVELIST").Item(0)
    
    'And the Incentives
    
    If Not xmlIncentives Is Nothing Then
                
        Dim intSequenceNumber As Integer
        intSequenceNumber = 1
        Dim xmlIncentiveList As IXMLDOMNodeList
        Dim xmlMortgageIncentive As IXMLDOMElement
        Set xmlMortgageIncentive = xmlElement.getElementsByTagName("MORTGAGEINCENTIVE").Item(0)
        If Not blnNewLoanComponentDetails Then 'New loan comp so Nothing to delete
            blnRaiseError = False
            objIMortgageSubquoteDO.DeleteAllMortgageIncentive xmlMortgageIncentive
            blnRaiseError = True
        End If
        
        Set xmlIncentiveList = xmlIncentives.selectNodes(".//MORTGAGEINCENTIVE")

        Dim xmlTmpElement As IXMLDOMElement
        Dim xmlCompletionIndicator As IXMLDOMElement
        Dim strNotInterfacedValue As String
        
        strNotInterfacedValue = objComboDO.GetFirstComboValueId("CompletionIndicator", "N")
        
        For Each xmlTmpElement In xmlIncentiveList
            If blnNewLoanComponentDetails Then
                'For a brand new loancomponent we need to set up the loancomponentsequencenumber in the mortgageincentive
                Dim xmlLoanCompSeqNumElem As IXMLDOMElement
                Set xmlLoanCompSeqNumElem = xmlTmpElement.getElementsByTagName("LOANCOMPONENTSEQUENCENUMBER").Item(0)
                xmlLoanCompSeqNumElem.Text = strNewLoanComponentSequenceNumber
                Set xmlCompletionIndicator = xmlOut.createElement("COMPLETIONINDICATOR")
                xmlCompletionIndicator.Text = strNotInterfacedValue
            End If
            Set xmlMortgageIncentive = CreateMortgageIncentiveXML(xmlTmpElement, Str(intSequenceNumber))
        
            objIMortgageSubquoteDO.CreateMortgageIncentive xmlMortgageIncentive
        
            intSequenceNumber = intSequenceNumber + 1
        Next
        
    End If
    
    Dim xmlMortgageSubQuoteUpdate As IXMLDOMElement
    
    ' Set the total loan amount
    Dim totalLoanAmount As Long
    totalLoanAmount = lngTotalLoans + objConvertAssist.CSafeLng( _
                        objXmlAssist.GetTagValue(xmlElement, "LOANAMOUNT"))
    Dim strTotalLoanAmount As String
    strTotalLoanAmount = CStr(totalLoanAmount)
    
    'And the mortgage subquote
    If blnNonPanelLenderOption Then
        
        Dim strNonPanelLenderName As String
        'Update Mortgage Sub quote here
        strNonPanelLenderName = objXmlAssist.GetTagValue(xmlElement, "NONPANELLENDERNAME")
        Set xmlMortgageSubQuoteUpdate = UpdateMortgageSubQuoteXML( _
            xmlElement, blnNonPanelLenderOption, strNonPanelLenderName, , strTotalLoanAmount)
        
    Else
        
        Set xmlMortgageSubQuoteUpdate = UpdateMortgageSubQuoteXML( _
            xmlElement, blnNonPanelLenderOption, , , strTotalLoanAmount)
        
    End If
    
    objIMortgageSubquoteDO.Update xmlMortgageSubQuoteUpdate

    Set IMortgageSubQuoteTxBO_SaveLoanComponentDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveLoanComponentDetailsExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objGlobalParameterDO = Nothing
    Set objComboDO = Nothing
    Set objIMortgageProductDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCompletionIndicator = Nothing
    Set objConvertAssist = Nothing
    
    Exit Function

SaveLoanComponentDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffected Or _
        objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Or _
        objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_SaveLoanComponentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLoanComponentDetailsExit

End Function

Private Function SaveLoanComponentXML(ByVal xmlElement As IXMLDOMElement, _
                                        Optional ByVal blnNew As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo SaveLoanComponentXMLVbErr

    Const strFunctionName As String = "SaveLoanComponentXML"

    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlLoanComponentElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
    Dim strText As String
    Dim strDetails As String
        
    Set xmlLoanComponentElem = xmlOut.createElement("LOANCOMPONENT")
    xmlOut.appendChild xmlLoanComponentElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    ' Ensure it is the LOANCOMPONENT table sequence number we get and not the one
    ' from the mortgageincentive table (if present)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName, , False)
    xmlLoanComponentElem.appendChild xmlTempElement
        
    strTagName = "MORTGAGEPRODUCTCODE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
        
    strTagName = "STARTDATE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement

    strTagName = "CAPITALANDINTERESTELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement

    strTagName = "INTERESTONLYELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement

    'SG 28/05/02 SYS4767 START
    strTagName = "NETCAPANDINTELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement

    strTagName = "NETINTONLYELEMENT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    'SG 28/05/02 SYS4767 END

    strTagName = "LOANAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement

    Set xmlTempElement = xmlOut.createElement("TOTALLOANCOMPONENTAMOUNT")
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
       
    'Not previously checked
    strTagName = "PURPOSEOFLOAN"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "PURPOSEOFLOAN = " & strText
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlLoanComponentElem.appendChild xmlTempElement
    
    'Not previously checked
    strTagName = "REPAYMENTMETHOD"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strText = m_objXmlAssist.GetTagValue(xmlElement, "REPAYMENTTYPE")
    
    If Len(strText) = 0 Then
        
        strDetails = "REPAYMENTMETHOD = " & strText
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails

    End If
        
    xmlTempElement.Text = strText
    xmlLoanComponentElem.appendChild xmlTempElement
    
'SW 20/06/2006 EP771 Start
    strTagName = "REPAYMENTVEHICLE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strText = m_objXmlAssist.GetTagValue(xmlElement, "REPAYMENTVEHICLE")

    xmlTempElement.Text = strText
    xmlLoanComponentElem.appendChild xmlTempElement
'SW 20/06/2006 EP771 End

'MAH    03/01/2007  EP2_444 start
    strTagName = "REPAYMENTVEHICLEMONTHLYCOST"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
'MAH    03/01/2007  EP2_444 end
        
    strTagName = "TERMINMONTHS"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "TERMINYEARS"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
          
    'SG 28/05/02 SYS4767 START
    strTagName = "PRODUCTCODESEARCHIND"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "RESOLVEDRATE"  'JLD MSMS0029
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "MANUALADJUSTMENTPERCENT"  'JLD MSMS0029
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    'SG 28/05/02 SYS4767 END
    
    'DPF 23/07/02 - CMWP3 - Added 2 new fields
    strTagName = "MANUALPORTEDLOANIND"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
        
    strTagName = "SUBPURPOSEOFLOAN"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    'DPF 23/07/02 - END
    
    If blnNew Then
        'only on create new
        strTagName = "PORTEDLOAN"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = "0" 'FALSE
        xmlLoanComponentElem.appendChild xmlTempElement
    
    End If
    
    'SR 11/08/2004 : BMIDS815 - add 3 new columns
    strTagName = "PRODUCTSTEP"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "REMAININGSTEPDURATION"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    
    strTagName = "PORTEDPRODUCTSTARTDATE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    'SR 11/08/2004 : BMIDS815 - End
    
    'MAR46 GHun
    strTagName = "PRODUCTSWITCHRETAINPRODUCTIND"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlLoanComponentElem.appendChild xmlTempElement
    'MAR46 End
    
    Set SaveLoanComponentXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set xmlLoanComponentElem = Nothing

    Exit Function

SaveLoanComponentXMLVbErr:
    
    Set xmlOut = Nothing
    Set xmlLoanComponentElem = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function


Private Function CreateMortgageIncentiveXML(ByVal xmlElement As IXMLDOMElement, _
                                            ByVal strMortgageIncentiveSequenceNumber As String, _
                                        Optional ByVal blnNew As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageIncentiveXMLVbErr

    Const strFunctionName As String = "CreateMortgageIncentiveXML"

        
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlMortgageIncentiveElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
    Dim strText As String
    Dim strDetails As String
    Dim objComboDO As ComboDO
    Dim strNotInterfaced As String
    
    ' PSC 17/01/02 SYS3801 - Start
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strNotInterfaced = objComboDO.GetFirstComboValueId("CompletionIndicator", "N")
    ' PSC 17/01/02 SYS3801 - End
    Set xmlMortgageIncentiveElem = xmlOut.createElement("MORTGAGEINCENTIVE")
    xmlOut.appendChild xmlMortgageIncentiveElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    
    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageIncentiveElem.appendChild xmlTempElement
        
    strTagName = "MORTGAGEINCENTIVESEQUENCENUM"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strMortgageIncentiveSequenceNumber
    xmlMortgageIncentiveElem.appendChild xmlTempElement
        
    'dpf 23/10/2002 - CPWP1
    strTagName = "INCENTIVEGUID"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageIncentiveElem.appendChild xmlTempElement

    'Not previously checked
    strTagName = "DESCRIPTION"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "DESCRIPTION = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlMortgageIncentiveElem.appendChild xmlTempElement

    'Not previously checked
    strTagName = "INCENTIVEAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "INCENTIVEAMOUNT = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    
    ' PSC 17/01/02 SYS3801 - Start
    strTagName = "INCENTIVEBENEFITTYPE"

    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "INCENTIVEBENFITTYPE = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    
    xmlTempElement.Text = strText
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    
    strTagName = "NOTIONALAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
        
    xmlTempElement.Text = strText
    xmlMortgageIncentiveElem.appendChild xmlTempElement
                                
    strTagName = "COMPLETIONINDICATOR"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strNotInterfaced
    xmlMortgageIncentiveElem.appendChild xmlTempElement
    ' PSC 17/01/02 SYS3801 - End
    
    Dim strIncentive As String
    strTagName = "TYPE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strIncentive = objXmlAssist.GetTagValue(xmlElement, strTagName)
    If StrComp(strIncentive, "Inclusive", vbTextCompare) = 0 Or _
        StrComp(strIncentive, "1", vbTextCompare) = 0 Then
    
        xmlTempElement.Text = "1"
    
    ElseIf StrComp(strIncentive, "Exclusive", vbTextCompare) = 0 Or _
            StrComp(strIncentive, "2", vbTextCompare) = 0 Then
    
        xmlTempElement.Text = "2"
    
    Else
    
        'Raise error ????
        
    End If
    
    xmlMortgageIncentiveElem.appendChild xmlTempElement
                    
    Set CreateMortgageIncentiveXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlMortgageIncentiveElem = Nothing
    Set objComboDO = Nothing


    Exit Function

CreateMortgageIncentiveXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlMortgageIncentiveElem = Nothing

    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function CreateNonPanelMortgageProductXML(ByVal xmlElement As IXMLDOMElement, _
                                        Optional ByVal blnNew As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNonPanelMortgageProductXMLVbErr

    Const strFunctionName As String = "CreateNonPanelMortgageProductXML"

        
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlNonPanelMortgageProductElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
    Dim strText As String
    Dim strDetails As String
        
    Set xmlNonPanelMortgageProductElem = xmlOut.createElement("LOANCOMPONENT")
    xmlOut.appendChild xmlNonPanelMortgageProductElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlNonPanelMortgageProductElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlNonPanelMortgageProductElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlNonPanelMortgageProductElem.appendChild xmlTempElement
    
    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlNonPanelMortgageProductElem.appendChild xmlTempElement
                
    'Not previously checked
    strTagName = "INTERESTRATEPERIOD1"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "INTERESTRATEPERIOD1 = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
        
    'Not previously checked
    strTagName = "INTERESTRATE1"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "INTERESTRATE1 = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If

    'Not previously checked
    strTagName = "INTERESTRATE2"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "INTERESTRATE2 = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If

    'Not previously checked
    strTagName = "MORTGAGEPRODUCTNAME"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "MORTGAGEPRODUCTNAME = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
    'Not previously checked
    strTagName = "MORTGAGEPRODUCTTYPE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    
    strText = objXmlAssist.GetTagValue(xmlElement, strTagName)
    
    If Len(strText) = 0 Then
        
        strDetails = "MORTGAGEPRODUCTTYPE = " & strText
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    
    End If
                     
    Set CreateNonPanelMortgageProductXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNonPanelMortgageProductElem = Nothing


    Exit Function

CreateNonPanelMortgageProductXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlNonPanelMortgageProductElem = Nothing

    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function UpdateMortgageSubQuoteXML(ByVal xmlElement As IXMLDOMElement, _
                                        ByVal blnNonPanelLender As Boolean, _
                                        Optional ByVal strNonPanelLenderName As String, _
                                        Optional ByVal blnResetMortgageSubQuote As Boolean = False, _
                                        Optional ByVal strTotalLoanAmount As String, _
                                        Optional ByVal strLTV As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteXMLVbErr

    Const strFunctionName As String = "UpdateMortgageSubQuoteXML"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlMortgageSubQuoteElem As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim strTagName As String
    Dim strText As String
    Dim strDetails As String
        
    Set xmlMortgageSubQuoteElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlOut.appendChild xmlMortgageSubQuoteElem
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageSubQuoteElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlMortgageSubQuoteElem.appendChild xmlTempElement
    
    If strTotalLoanAmount <> "" Then
        strTagName = "TOTALLOANAMOUNT"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = strTotalLoanAmount
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
    End If
    
    If blnNonPanelLender Then
    
        strTagName = "NONPANELLENDERSELECTD"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = "1" 'True
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
            
           
        strTagName = "NONPANELLENDERNAME"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = strNonPanelLenderName
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
    
    Else
        strTagName = "NONPANELLENDERSELECTD"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = "0" 'False
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
            
           
        strTagName = "NONPANELLENDERNAME"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = ""
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
        
    End If
          
    If blnResetMortgageSubQuote Then
          
        strTagName = "TOTALLOANAMOUNT"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = "0"
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
              
        strTagName = "AMOUNTREQUESTED"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
        
        'MAR1061 update purchasePriceOrEstimatedValue
        strTagName = "PURCHASEPRICEORESTIMATEDVALUE"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
        
        ' PSC 04/04/00 AQR SYS0593
        ' PSC 27/04/00 AQR SYS0593 - Update LTV with new value rather than null
        strTagName = "LTV"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = strLTV
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
        
        'JLD SYS0715 Update Deposit
        strTagName = "DEPOSIT"
        Set xmlTempElement = xmlOut.createElement(strTagName)
        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
        xmlMortgageSubQuoteElem.appendChild xmlTempElement
        
        'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Roll Back
        'SG 28/05/02 SYS4767 START
        'JLD MSMS0034 Update typeofmortgage and typeofbuyer too
        'strTagName = "TYPEOFAPPLICATION"
        'Set xmlTempElement = xmlOut.createElement(strTagName)
        'xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
        'xmlMortgageSubQuoteElem.appendChild xmlTempElement
        'strTagName = "TYPEOFBUYER"
        'Set xmlTempElement = xmlOut.createElement(strTagName)
        'xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
        'xmlMortgageSubQuoteElem.appendChild xmlTempElement
        'SG 28/05/02 SYS4767 END
          
    End If
    
    Set UpdateMortgageSubQuoteXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlMortgageSubQuoteElem = Nothing


    Exit Function

UpdateMortgageSubQuoteXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlMortgageSubQuoteElem = Nothing

    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteTxBO_ResetMortgageSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_ResetMortgageSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else

    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
    
    Dim xmlMortgageSubQuoteUpdate As IXMLDOMElement
    
    'Delete all loan components and related records
    objIMortgageSubquoteDO.ResetMortgageSubQuote xmlElement
    
    ' PSC 27/04/00 AQR SYS0593 - Start
    Dim strLTV As String
    strLTV = objXmlAssist.GetTagValue(xmlElement, "LTV")
    
    Set xmlMortgageSubQuoteUpdate = UpdateMortgageSubQuoteXML(xmlElement, _
                                                              False, , True, _
                                                              "", strLTV)
    ' PSC 27/04/00 AQR SYS0593 - End
    
    objIMortgageSubquoteDO.Update xmlMortgageSubQuoteUpdate

    Set IMortgageSubQuoteTxBO_ResetMortgageSubQuote = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ResetMortgageSubQuoteExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ResetMortgageSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_ResetMortgageSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit

End Function

Private Function IMortgageSubQuoteTxBO_SaveIllustrationResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveIllustrationResultsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_SaveIllustrationResults"
 
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    'Dim strContext As String
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    'Dim xmlLoanComponent As IXMLDOMElement
    'Dim xmlLoanComponentResults As IXMLDOMElement
    'Dim xmlTempRequest As IXMLDOMNode  'BMIDS736 GHun
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
     
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    'BMIDS736 GHun SaveIllustrationResults now saves all LoanComponents in the list in
    'one call, rather than being called multiple times, once for each component
    If vXMLRequest.nodeName = "LOANCOMPONENTLIST" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("LOANCOMPONENTLIST").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, "LOANCOMPONENTLIST" & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
    
    'BMIDS782 GHun Call a stored proc to do all the updates in one database call
    objIMortgageSubquoteDO.SaveIllustrationResults xmlElement
    
    ''BMIDS736 GHun
    'For Each xmlLoanComponent In xmlElement.selectNodes("LOANCOMPONENT")
    '    objIMortgageSubQuoteDO.UpdateLoanComponent xmlLoanComponent
    '
    '    'As the number of rate periods may have changed, it is safer to delete
    '    'the payment and balance schedules and recreate them
    '    Set xmlTempRequest = xmlOut.createElement("LOANCOMPONENTPAYMENTSCHEDULE")
    '    xmlTempRequest.appendChild xmlLoanComponent.selectSingleNode("APPLICATIONNUMBER").cloneNode(True)
    '    xmlTempRequest.appendChild xmlLoanComponent.selectSingleNode("APPLICATIONFACTFINDNUMBER").cloneNode(True)
    '    xmlTempRequest.appendChild xmlLoanComponent.selectSingleNode("MORTGAGESUBQUOTENUMBER").cloneNode(True)
    '    xmlTempRequest.appendChild xmlLoanComponent.selectSingleNode("LOANCOMPONENTSEQUENCENUMBER").cloneNode(True)
    '    objIMortgageSubQuoteDO.DeleteAllLoanComponentPaymentSchedules xmlTempRequest
    '
    '    xmlChangeNodeName xmlTempRequest, "LOANCOMPONENTPAYMENTSCHEDULE", "LOANCOMPONENTBALANCESCHEDULE"
    '    objIMortgageSubQuoteDO.DeleteAllLoanComponentBalanceSchedules xmlTempRequest
    '
    '    xmlChangeNodeName xmlTempRequest, "LOANCOMPONENTBALANCESCHEDULE", "LOANCOMPONENTREDEMPTIONFEE"
    '    objIMortgageSubQuoteDO.DeleteAllLoanComponentRedemptionFees xmlTempRequest
    'Next
    '
    'For Each xmlLoanComponent In xmlElement.selectNodes("LOANCOMPONENTPAYMENTSCHEDULE")
    '    objIMortgageSubQuoteDO.CreateLoanComponentPaymentSchedule xmlLoanComponent
    'Next
    '
    'For Each xmlLoanComponent In xmlElement.selectNodes("LOANCOMPONENTBALANCESCHEDULE")
    '    objIMortgageSubQuoteDO.CreateLoanComponentBalanceSchedule xmlLoanComponent
    'Next
    '
    'For Each xmlLoanComponent In xmlElement.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
    '    objIMortgageSubQuoteDO.CreateLoanComponentRedemptionFee xmlLoanComponent
    'Next
    '
    ''Set xmlLoanComponent = UpdateLoanComponentXML(xmlElement, xmlResponseElem)
    ''objIMortgageSubQuoteDO.UpdateLoanComponent xmlLoanComponent
    ''Set xmlLoanComponentResults = CreateLoanComponentResultsXML(xmlElement, xmlResponseElem)
    ''
    ''' PSC 03/07/2002 BMIDS00062 - Start
    ''strContext = m_objXmlAssist.GetTagValue(xmlElement, "CONTEXT")
    ''If UCase(strContext) = "UPDATE" Then
    ''    objIMortgageSubQuoteDO.UpdateLoanComponentResults xmlLoanComponentResults
    ''Else
    ''    objIMortgageSubQuoteDO.CreateLoanComponentResults xmlLoanComponentResults
    ''End If
    ''' PSC 03/07/2002 BMIDS00062 - End
    ''BMIDS736 End
    'BMIDS782 End
    
    Set IMortgageSubQuoteTxBO_SaveIllustrationResults = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveIllustrationResultsExit:
    Set objIMortgageSubquoteDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    'Set xmlLoanComponent = Nothing
    'Set xmlTempRequest = Nothing
    
    Exit Function

SaveIllustrationResultsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_SaveIllustrationResults = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveIllustrationResultsExit

End Function

Private Function IMortgageSubQuoteTxBO_SaveOneOffCostDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveOneOffCostDetailsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_SaveOneOffCostDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objComboDO As ComboDO
    
    ' PSC 03/07/2002 BMIDS00062
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objComboDO = New ComboDO
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' PSC 03/07/2002 BMIDS00062
    Set objMortgageSubQuoteBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
    
    Dim xmlCopyNode As IXMLDOMNode
    
    Set xmlCopyNode = xmlElement.cloneNode(True)
            
    m_objXmlAssist.ChangeNodeName xmlCopyNode, "MORTGAGESUBQUOTE", "MORTGAGEONEOFFCOST"
    '1 Delete existing one-off costs and save revised ones
    objIMortgageSubquoteDO.DeleteAllMortgageOneOffCost xmlCopyNode

    Dim xmlOneOffCostElement As IXMLDOMElement
    Dim xmlOneOffCostList As IXMLDOMNodeList
    
    Set xmlOneOffCostList = xmlElement.selectNodes(".//ONEOFFCOST")
                
    'loop round all the one off costs
    Dim intOneOffCostSequenceNumber As Integer
    intOneOffCostSequenceNumber = 1
    Dim strComboValueId As String
    Dim strComboValidationType As String     ' BMIDS881
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlMortgageOneOffCost As IXMLDOMElement
    
    Dim dblTotalCharges As Double
    Dim dblTotalInitialCharges As Double
    Dim dblCostsToBeAdded As Double         'CL 26/9/00 SYS0952
    Dim xmlOneOffCost As IXMLDOMElement
    Dim xmlMortgageSubquote As IXMLDOMElement
    
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlFirstLoanComponent As IXMLDOMNode
 
    For Each xmlOneOffCostElement In xmlOneOffCostList

        Set xmlTempNode = xmlOneOffCostElement.getElementsByTagName("COMBOVALIDATIONTYPE").Item(0)
        strComboValidationType = xmlTempNode.Text                              ' BMIDS881
        'We dont do TID not really a one off cost
        If StrComp(strComboValidationType, "TID", vbTextCompare) <> 0 Then     ' BMIDS881
            'validate it
            strComboValueId = objComboDO.GetFirstComboValueId("OneOffCost", xmlTempNode.Text)
    
            If Len(strComboValueId) = 0 Then
            
                    m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeInvalidParameter, _
                                    "No Combo Value Id for : " & xmlTempNode.Text
    
            End If
                                
            Set xmlMortgageOneOffCost = CreateNewMortgageOneOffCostXML(xmlElement, _
                                            intOneOffCostSequenceNumber, _
                                            xmlOneOffCostElement, strComboValueId)
    
            objIMortgageSubquoteDO.CreateMortgageOneOffCost xmlMortgageOneOffCost
    
        End If
        
        dblTotalCharges = dblTotalCharges + _
            CSafeDbl(m_objXmlAssist.GetTagValue(xmlOneOffCostElement, "AMOUNT"))

        Dim strPattern As String
        ' Calculate the total final charges based on Sealing Fee and Deeds Release Fee
        
        'BMIDS881  Correct test for Sealing Fee or Deeds Release Fee
'        strPattern = ".//IDENTIFIER[. != ""DEE"" or IDENTIFIER != ""SEA""]"
'        Set xmlOneOffCost = xmlOneOffCostElement.selectSingleNode(strPattern)
'
'        If Not xmlOneOffCost Is Nothing Then
        If (StrComp(strComboValidationType, "DEE", vbTextCompare) <> 0) And _
           (StrComp(strComboValidationType, "SEA", vbTextCompare) <> 0) Then

            dblTotalInitialCharges = dblTotalInitialCharges + _
                CSafeDbl(m_objXmlAssist.GetTagValue(xmlOneOffCostElement, "AMOUNT"))
        End If

        If CSafeBool(m_objXmlAssist.GetTagValue(xmlOneOffCostElement, "ADDEDTOLOAN")) Then
        
            dblCostsToBeAdded = dblCostsToBeAdded + _
                CSafeDbl(m_objXmlAssist.GetTagValue(xmlOneOffCostElement, "AMOUNT"))
        
        End If   'Get Total Charges and costs to be added
    
        intOneOffCostSequenceNumber = intOneOffCostSequenceNumber + 1
    
    Next
 
    Set xmlMortgageSubquote = objIMortgageSubquoteDO.GetData(xmlElement)
    
    ' PSC 03/07/2002 BMIDS00062 - Start
'    objXmlAssist.ChangeNodeName xmlCopyNode, "MORTGAGEONEOFFCOST", "LOANCOMPONENT"

'    Dim xmlLoanComponent As IXMLDOMElement
'    Dim dblExistingTotalGrossMonthlyCost As Double
'    Dim dblLoanComponentGrossMonthlyCost As Double
'    Dim dblInComponentGrossMonthlyCost As Double
'
'    Set xmlLoanComponent = objIMortgageSubQuoteDO.GetLoanComponentData(xmlCopyNode)
'
    'Update Mortgage subquote
    Dim xmlTempElement As IXMLDOMElement

    'Total Charges
    
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALCHARGES").Item(0)
        xmlTempElement.Text = CStr(dblTotalCharges)
    
    'TotalGrossMonthlyCost
        
'        dblLoanComponentGrossMonthlyCost = objConvAssist.CSafeDbl( _
'                                            objXmlAssist.GetTagValue(xmlLoanComponent, _
'                                            "GROSSMONTHLYCOST"))
'
'        dblInComponentGrossMonthlyCost = objConvAssist.CSafeDbl( _
'                                        objXmlAssist.GetTagValue(xmlElement, _
'                                            "COMPONENTGROSSMONTHLYCOST"))

        
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALGROSSMONTHLYCOST").Item(0)
        Set xmlNode = vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/TOTALGROSSMONTHLYCOST")
        xmlTempElement.Text = xmlNode.Text

'        dblExistingTotalGrossMonthlyCost = objConvAssist.CSafeDbl(xmlTempElement.Text)
'
'        xmlTempElement.Text = CStr(dblExistingTotalGrossMonthlyCost - dblLoanComponentGrossMonthlyCost + _
                                dblInComponentGrossMonthlyCost)
        ' PSC 03/07/2002 BMIDS00062 - End
    
    'Initial Charges
    
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALINITIALCHARGES").Item(0)
        xmlTempElement.Text = CStr(dblTotalInitialCharges)

    'Total Loan Amount
    
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALLOANAMOUNT").Item(0)
        xmlTempElement.Text = CSafeDbl( _
                                m_objXmlAssist.GetTagValue(xmlElement, "TOTALLOANAMOUNT"))

    'Total Net Charges

        Dim dblDeposit As Double
        Dim dblTotalIncentives As Double
        Dim dblManualIncentives As Double  'DPF 10/10/2002 - CPWP1 (BM020)
        Dim dblIncentives As Double  'DPF 10/10/2002 - CPWP1 (BM020)
        Dim xmlRequestElement As IXMLDOMNode 'DPF 10/10/2002 - CPWP1 (BM020)

        dblDeposit = CSafeDbl(m_objXmlAssist.GetTagValue(xmlMortgageSubquote, "DEPOSIT"))
    
        dblTotalIncentives = CSafeDbl(m_objXmlAssist.GetTagValue(xmlElement, "TOTALINCENTIVES"))
        'DPF 10/10/2002 - CPWP1 (BM020)
        dblManualIncentives = CSafeDbl(m_objXmlAssist.GetTagValue(xmlElement, "MANUALINCENTIVES"))
        
        If IsNumeric(dblManualIncentives) And dblManualIncentives > 0 Then
            dblIncentives = dblManualIncentives
            Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("MANUALINCENTIVEAMOUNT").Item(0)
            xmlTempElement.Text = dblManualIncentives
            Set xmlRequestElement = xmlGetRequestNode(vXMLRequest)
            Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("MANUALINCENTIVEUSERID").Item(0)
            xmlTempElement.Text = xmlGetMandatoryAttribute(xmlRequestElement, "USERID").Text
            Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("MANUALINCENTIVEDATETIME").Item(0)
            xmlTempElement.Text = Format$(Now, "DD/MM/YYYY HH:NN:SS")
        Else
            dblIncentives = dblTotalIncentives
        End If
        'END:  DPF 10/10/2002 - CPWP1 (BM020)
    
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALNETCHARGES").Item(0)
        'CL 26/9/00 SYS0952
        xmlTempElement.Text = CStr(dblTotalCharges + dblDeposit - dblIncentives - dblCostsToBeAdded)

    'Total Net Monthly Cost
    
         ' PSC 03/07/2002 BMIDS00062 - Start
'        Dim dblLoanComponentNetMonthlyCost As Double
'        Dim dblInComponentNetMonthlyCost As Double
'        Dim dblTotalNetMonthlyCost As Double
'
'        dblLoanComponentNetMonthlyCost = objConvAssist.CSafeDbl(objXmlAssist.GetTagValue( _
'                                        xmlLoanComponent, "NETMONTHLYCOST"))
'
'        dblInComponentNetMonthlyCost = objConvAssist.CSafeDbl(objXmlAssist.GetTagValue( _
'                                        xmlElement, "COMPONENTNETMONTHLYCOST"))
        
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALNETMONTHLYCOST").Item(0)

'        dblTotalNetMonthlyCost = objConvAssist.CSafeDbl(xmlTempElement.Text)
'
'        xmlTempElement.Text = CStr(dblTotalNetMonthlyCost - dblLoanComponentNetMonthlyCost + _
'                                dblInComponentNetMonthlyCost)
        
        Set xmlNode = vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/TOTALNETMONTHLYCOST")
        xmlTempElement.Text = xmlNode.Text
        
    'Total Monthly Cost Less Drawdown
    
        If Not vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/MONTHLYCOSTLESSDRAWDOWN") Is Nothing Then
            Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("MONTHLYCOSTLESSDRAWDOWN").Item(0)
            Set xmlNode = vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/MONTHLYCOSTLESSDRAWDOWN")
            xmlTempElement.Text = xmlNode.Text
        End If
        
    ' Accrued Interest
        Set xmlTempElement = xmlMortgageSubquote.getElementsByTagName("TOTALACCRUEDINTEREST").Item(0)
        Set xmlNode = vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/TOTALACCRUEDINTEREST")
        xmlTempElement.Text = xmlNode.Text
        
        ' PSC 03/07/2002 BMIDS00062 - End
        
    'BMIDS766 GHun
    ' AmountPerUnitBorrowed, APR TotalAmountPayable, TotalMortgagePayments
    Set xmlTempElement = xmlGetMandatoryNode(xmlMortgageSubquote, "AMOUNTPERUNITBORROWED")
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/AMOUNTPERUNITBORROWED")
    xmlTempElement.Text = xmlNode.Text
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlMortgageSubquote, "APR")
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/APR")
    xmlTempElement.Text = xmlNode.Text
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlMortgageSubquote, "TOTALAMOUNTPAYABLE")
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/TOTALAMOUNTPAYABLE")
    xmlTempElement.Text = xmlNode.Text
    
    Set xmlTempElement = xmlGetMandatoryNode(xmlMortgageSubquote, "TOTALMORTGAGEPAYMENTS")
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/TOTALMORTGAGEPAYMENTS")
    xmlTempElement.Text = xmlNode.Text
    'BMIDS766 End
        
    objIMortgageSubquoteDO.Update xmlMortgageSubquote
    
    ' PSC 03/07/2002 BMIDS00062 - Start
    'Update LoanComponent
        
    'Apr
        
'        Set xmlTempElement = xmlLoanComponent.getElementsByTagName("APR").Item(0)
'        xmlTempElement.Text = objConvAssist.CSafeDbl( _
                                objXmlAssist.GetTagValue(xmlElement, "APR"))
    'GrossMonthlyCost

'        Set xmlTempElement = xmlLoanComponent.getElementsByTagName("GROSSMONTHLYCOST").Item(0)
'        xmlTempElement.Text = objConvAssist.CSafeDbl( _
'                                objXmlAssist.GetTagValue(xmlElement, "COMPONENTGROSSMONTHLYCOST"))

    'NetMonthlyCost

'        Set xmlTempElement = xmlLoanComponent.getElementsByTagName("NETMONTHLYCOST").Item(0)
'        xmlTempElement.Text = objConvAssist.CSafeDbl( _
'                                objXmlAssist.GetTagValue(xmlElement, "COMPONENTNETMONTHLYCOST"))

    'InterestOnlyAmount
    Dim strTemp As String
    strTemp = m_objXmlAssist.GetTagValue(xmlElement, "INTERESTONLYAMOUNT")
    'empty string or 0 ??
    If StrComp(strTemp, "0", vbTextCompare) <> 0 And strTemp <> "" Then 'JLD BMIDS805 check for empty string too
    
        ' Find first loan component
        Dim intMinSeqNo As Integer
        Dim intSeqNo As Integer
        intMinSeqNo = 9999

        Set xmlLoanComponentList = xmlElement.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT[PORTEDLOAN !='1']")
    
        For Each xmlLoanComponent In xmlLoanComponentList
            intSeqNo = CDbl(m_objXmlAssist.GetMandatoryElementText(xmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER"))
            
            If intSeqNo < intMinSeqNo Then
                intMinSeqNo = intSeqNo
                Set xmlFirstLoanComponent = xmlLoanComponent
            End If
        Next
    
        Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlFirstLoanComponent, "CAPITALANDINTERESTELEMENT")
        xmlTempElement.Text = CSafeDbl(m_objXmlAssist.GetTagValue(xmlElement, "CAPITALANDINTERESTAMOUNT"))

        Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlFirstLoanComponent, "INTERESTONLYELEMENT")
        xmlTempElement.Text = CSafeDbl(m_objXmlAssist.GetTagValue(xmlElement, "INTERESTONLYAMOUNT"))

    End If
'
'        Dim dblLoanAmount As Double
'
'        Set xmlTempElement = xmlLoanComponent.getElementsByTagName("LOANAMOUNT").Item(0)
'        dblLoanAmount = objConvAssist.CSafeDbl(xmlTempElement.Text)
'        dblLoanAmount = dblLoanAmount + dblCostsToBeAdded
'        Set xmlTempElement = xmlLoanComponent.getElementsByTagName("TOTALLOANCOMPONENTAMOUNT").Item(0)
'        xmlTempElement.Text = CStr(dblLoanAmount)
'
'
'    objIMortgageSubQuoteDO.UpdateLoanComponent xmlLoanComponent
'
'    Dim xmlLoanComponentResults As IXMLDOMElement
'
'    Set xmlLoanComponentResults = CreateLoanComponentResultsXML(xmlElement)
'
'    objIMortgageSubQuoteDO.UpdateLoanComponentResults xmlLoanComponentResults
'    objIMortgageSubQuoteDO.SaveOneOffCostDetails xmlElement

    Dim strResponse As String
    strResponse = objMortgageSubQuoteBO.SaveCompositeRateResults(vXMLRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
    ' PSC 03/07/2002 BMIDS00062 - End
    
    Set IMortgageSubQuoteTxBO_SaveOneOffCostDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveOneOffCostDetailsExit:

    Set objIMortgageSubquoteDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlFirstLoanComponent = Nothing
    
    Exit Function

SaveOneOffCostDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_SaveOneOffCostDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOneOffCostDetailsExit

End Function

'BMIDS782 GHun
'BMIDS736 GHun No longer used, but cannot be totally removed due to OOSS
'Private Function UpdateLoanComponentXML(ByVal xmlElement As IXMLDOMElement, _
'                            ByVal xmlResponse As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:  Creates the xml structure for the first sub quote
'
'' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
''                               instance of the persistant data to be retrieved
''
'' return:       First sub quote xml
''------------------------------------------------------------------------------------------
'On Error GoTo UpdateLoanComponentXMLVbErr
'
'    Const strFunctionName As String = "UpdateLoanComponentXML"
'
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim xmlTempElement As IXMLDOMElement
'    Dim xmlLoanComponentOut As IXMLDOMElement
'    Dim strTagName As String
'    Dim strText As String
'    Dim strNetMonthlyCost As String
'    Dim strMonthlyCostLessDrawDown As String
'    Dim strGrossMonthlyCost As String
'
'    Set xmlOut = New FreeThreadedDOMDocument40
'    Set xmlLoanComponentOut = xmlOut.createElement("LOANCOMPONENT")
'    xmlOut.appendChild xmlLoanComponentOut
'
'    strTagName = "APPLICATIONNUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "APPLICATIONFACTFINDNUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "MORTGAGESUBQUOTENUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "APR"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlElement, "OUT030_APR")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETMONTHLYCOST"
'    strNetMonthlyCost = m_objXmlAssist.GetTagValue(xmlElement, "OUT060_TOTALNETMORTGAGEPAYMENT1")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strNetMonthlyCost
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    'DPF BMIDS00808 - 8/11/2002
'    If Not xmlElement.selectSingleNode("//MONTHLYCOSTLESSDRAWDOWN") Is Nothing Then
'        strTagName = "MONTHLYCOSTLESSDRAWDOWN"
'        strMonthlyCostLessDrawDown = m_objXmlAssist.GetTagValue(xmlElement, "MONTHLYCOSTLESSDRAWDOWN")
'        Set xmlTempElement = xmlOut.createElement(strTagName)
'        xmlTempElement.Text = strMonthlyCostLessDrawDown
'        xmlLoanComponentOut.appendChild xmlTempElement
'    End If
'    'END OF DPF BMIDS00808
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strNetMonthlyCost
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "GROSSMONTHLYCOST"
'
'    strGrossMonthlyCost = m_objXmlAssist.GetTagValue(xmlElement, "OUT130_TOTALGROSSMORTGAGEPAYMENT1")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strGrossMonthlyCost
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strGrossMonthlyCost
'    xmlResponse.appendChild xmlTempElement
'
'    ' PSC 03/07/2002 BMIDS00062 - Start
'    strTagName = "FINALRATE"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, "OUT220_FINALRATE")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "FINALRATEAPR"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, "OUT230_FINALRATEAPR")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "FINALRATEMONTHLYCOST"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, "OUT240_FINALRATEMONTHLYCOST")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "TOTALLOANCOMPONENTAMOUNT"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "CAPITALANDINTERESTELEMENT"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'
'    strTagName = "INTERESTONLYELEMENT"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'    ' PSC 03/07/2002 BMIDS00062 - End
'
'    'INR BMIDS624/ICC001
'    strTagName = "RESOLVEDRATE"
'    strText = m_objXmlAssist.GetTagValue(xmlElement, strTagName)
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'    xmlTempElement.Text = strText
'    xmlResponse.appendChild xmlTempElement
'    'INR ICC001 - End
'
'    Set UpdateLoanComponentXML = xmlOut.documentElement
'
'    Set xmlOut = Nothing
'    Set xmlLoanComponentOut = Nothing
'    Set xmlTempElement = Nothing
'
'    Exit Function
'
'UpdateLoanComponentXMLVbErr:
'
'    Set xmlOut = Nothing
'    Set xmlLoanComponentOut = Nothing
'    Set xmlTempElement = Nothing
'
'    m_objErrAssist.AddToErrSource strFunctionName
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    '   re-raise error to calling function
'    Err.Raise Err.number, Err.Source, Err.Description
'
'End Function
'BMIDS736 End
'BMIDS782 End

'BMIDS782 GHun
'BMIDS736 GHun No longer used, but cannot be totally removed due to OOSS
'Private Function CreateLoanComponentResultsXML(ByVal xmlElement As IXMLDOMElement, _
'                                            Optional ByVal xmlResponse As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:  Creates the xml structure for the first sub quote
'
'' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
''                               instance of the persistant data to be retrieved
''
'' return:       First sub quote xml
''------------------------------------------------------------------------------------------
'On Error GoTo CreateLoanComponentResultsXMLVbErr
'
'    Dim strFunctionName As String
'    strFunctionName = "CreateLoanComponentResultsXML"
'
'    #If PROFILING Then
'        Dim objLogAssist As New LogAssist
'        objLogAssist.StartTimerEx TypeName(Me) & "." & strFunctionName, False, otBO
'    #End If
'
'
'    Dim xmlOut As New FreeThreadedDOMDocument40
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
'    Dim xmlTempElement As IXMLDOMElement
'    Dim xmlLoanComponentOut As IXMLDOMElement
'    Dim strTagName As String
'    Dim strText As String
'    Dim strDetails As String
'
'    Set xmlLoanComponentOut = xmlOut.createElement("LOANCOMPONENTRESULTS")
'    xmlOut.appendChild xmlLoanComponentOut
'
'    strTagName = "APPLICATIONNUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "APPLICATIONFACTFINDNUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "MORTGAGESUBQUOTENUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "LOANCOMPONENTSEQUENCENUMBER"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Dim strAccruedInterest As String
'
'    strTagName = "ACCRUEDINTEREST"
'
'    strAccruedInterest = objXmlAssist.GetTagValue(xmlElement, "OUT020_ACCRUEDINTEREST")
'
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = strAccruedInterest
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    If Not xmlResponse Is Nothing Then
'
'        Set xmlTempElement = xmlResponse.ownerDocument.createElement(strTagName)
'        xmlTempElement.Text = strAccruedInterest
'        xmlResponse.appendChild xmlTempElement
'
'    End If
'
'    strTagName = "TOTALAMOUNTPAYABLE"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT040_TOTALAMOUNTPAYABLE")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "TOTALMORTGAGEPAYMENTS"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT050_TOTALMORTGAGEPAYMENTS")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT2"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT070_TOTALNETMORTGAGEPAYMENT2")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT3"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT080_TOTALNETMORTGAGEPAYMENT3")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT4"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT090_TOTALNETMORTGAGEPAYMENT4")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT5"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT100_TOTALNETMORTGAGEPAYMENT5")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT6"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT110_TOTALNETMORTGAGEPAYMENT6")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "NETPAYMENT7"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT120_TOTALNETMORTGAGEPAYMENT7")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT2"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT140_TOTALGROSSMORTGAGEPAYMENT2")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT3"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT150_TOTALGROSSMORTGAGEPAYMENT3")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT4"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT160_TOTALGROSSMORTGAGEPAYMENT4")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT5"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT170_TOTALGROSSMORTGAGEPAYMENT5")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT6"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT180_TOTALGROSSMORTGAGEPAYMENT6")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "GROSSPAYMENT7"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT190_TOTALGROSSMORTGAGEPAYMENT7")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "UNROUNDEDAPR"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT200_UNROUNDEDAPR")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    strTagName = "FINALPAYMENT"
'    Set xmlTempElement = xmlOut.createElement(strTagName)
'    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "OUT210_FINALGROSSPAYMENT")
'    xmlLoanComponentOut.appendChild xmlTempElement
'
'    Set CreateLoanComponentResultsXML = xmlOut.documentElement
'
'    Set xmlOut = Nothing
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
'    Set xmlLoanComponentOut = Nothing
'
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx otBO
'        Set objLogAssist = Nothing
'    #End If
'
'    Exit Function
'
'CreateLoanComponentResultsXMLVbErr:
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx otBO
'        Set objLogAssist = Nothing
'    #End If
'
'    Set xmlOut = Nothing
'    Set objXmlAssist = Nothing
'    Set xmlLoanComponentOut = Nothing
'
'
'    objErrAssist.AddToErrSource strFunctionName
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    Set objErrAssist = Nothing
'
'    '   re-raise error to calling function
'    Err.Raise Err.number, Err.Source, Err.Description
'End Function
'BMIDS736 End
'BMIDS782 End

Private Function IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSaveMortgageCostsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts"

' PJO 25/09/2003 BMIDS631
' Ignore duplicate key errors on ReserveMortgageProduce
    Dim blnReserve As Boolean
    blnReserve = False

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlDrawDown As IXMLDOMElement 'CPWP1 - DPF 4/10/2002
    Dim xmlReserveProd As IXMLDOMElement ' PJO 20/10/2003
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlOneOffCostList As IXMLDOMElement
    Dim xmlCompRateRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlOldNode As IXMLDOMNode

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objChargesBO As ChargesBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objGlobalParameterDO As GlobalParameterDO 'MAR1703
    
    Dim xmlCalcLoanCompList As IXMLDOMNode
    Dim xmlCalcResponse As IXMLDOMNode
    Dim xmlDrawDownResponse As IXMLDOMNode 'CPWP1 - DPF 8/10/2002
    
    'JD BMIDS854
    Dim xmlMSQNo As IXMLDOMNode
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFFNo As IXMLDOMNode
    Dim objMortSubQuoteBO As MortgageSubQuoteBO
    Dim xmlOneOffCosts As IXMLDOMNode
    Dim strResponse As String
    Dim xmlAPROneOffCosts As IXMLDOMNodeList
    Dim xmlOneOffCost As IXMLDOMNode
    Dim xmlAPROneOffCostsDoc As FreeThreadedDOMDocument40
    Dim blnRaiseError As Boolean
    Dim xmlAPRReqNode As IXMLDOMNode
    Dim xmlAPRMSQNode As IXMLDOMNode
    Dim blnGlobalReserve As Boolean 'MAR1703
    
    Dim xmlMortgageLender As IXMLDOMNode    ' PSC 09/11/2006 EP2_41
    Dim blnCalcFeesOnly As Boolean          ' PSC 09/11/2006 EP2_41
    Dim dblAddedCharges As Double           ' PSC 09/11/2006 EP2_41
    Dim blnCalcAndSaveProcFeesRequired As Boolean     ' EP2_893
    Dim lngWarningMsg As Long               ' EP2_893

    Dim xmlAddBorrow    As IXMLDOMElement   'EP2_954 GHun
    Dim xmlValuationType As IXMLDOMElement  'EP2_1281
    Dim blnIncentivePercentExclDrawdown  As Boolean ' EP2_1228
    Dim dblDrawdown As Double                       ' EP2_1228  'EP2_1798 - double rather than currency
    Dim xmlComponent As IXMLDOMNode                 ' EP2_1228
    Dim intComponent As Integer                     ' EP2_1228
    Dim xmlLoanAmount As IXMLDOMNode                ' EP2_1228
    Dim dblLoanAmount As Double                     ' EP2_1228  'EP2_1798 - double rather than currency
    Dim objCRUD As omCRUDBO                         ' EP2_1228
    Dim xmlTempRequest As IXMLDOMNode               ' EP2_1228
    Dim xmlCRUDMortgageIncentive As DOMDocument     ' EP2_1228
    Dim xmlRoot As IXMLDOMNode                      ' EP2_1228
    Dim objComboDO As ComboDO                       ' EP2_1228
    Dim strFinancialBenefitID As String             ' EP2_1228
    Dim xmlIncentive As IXMLDOMNode                 ' EP2_1228
    Dim curIncentive As Currency                    ' EP2_1228
    Dim dblIncentivePercent As Double               ' EP2_1228
    Dim xmlCRUDIncentive As DOMDocument             ' EP2_1228
    Dim dblIncentiveAmount As Double                ' EP2_1228  'EP2_1798 - double rather than currency
    Dim xmlIncentiveAmount As IXMLDOMNode           ' EP2_1228
    
    'MAR836  Initialise Raise Error flag so that correct error is reported
    blnRaiseError = True

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objChargesBO = New ChargesBO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objGlobalParameterDO = New GlobalParameterDO 'MAR1703
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objChargesBO = m_objContext.CreateInstance(App.Title & ".ChargesBO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO") 'MAR1703
    End If
    
    
    'Copy the node to use locally
    Dim xmlCopyNode As IXMLDOMNode
    Set xmlCopyNode = xmlElement.cloneNode(True)
    
    'JD BMIDS854
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlCopyNode, "APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXmlAssist.GetNode(xmlCopyNode, "APPLICATIONFACTFINDNUMBER")
    Set xmlMSQNo = m_objXmlAssist.GetNode(xmlCopyNode, "MORTGAGESUBQUOTENUMBER")
    
    '1 Calculate total incentive amount
    m_objXmlAssist.ChangeNodeName xmlCopyNode, "MORTGAGESUBQUOTE", "MORTGAGEINCENTIVE"
    
    Dim dblTotalIncentiveAmount As Double
    
    'EP2_1228
    blnIncentivePercentExclDrawdown = objGlobalParameterDO.GetCurrentParameterByType("IncentivePercentExclDrawdown", "BOOLEAN")
    If blnIncentivePercentExclDrawdown Then
                
        dblDrawdown = CSafeDbl(m_objXmlAssist.GetTagValue(xmlCopyNode, "DRAWDOWN"))  'SR EP2_1798
        If dblDrawdown > 0 Then
        
            For intComponent = (xmlCopyNode.selectNodes("LOANCOMPONENTLIST/LOANCOMPONENT").length - 1) To 0 Step -1
                Set xmlComponent = xmlCopyNode.selectSingleNode("LOANCOMPONENTLIST").childNodes(intComponent)
                Set xmlLoanAmount = m_objXmlAssist.GetNode(xmlComponent, "LOANAMOUNT")
                dblLoanAmount = CSafeDbl(xmlLoanAmount.Text) 'SR EP2_1798
                
                If dblDrawdown > 0 Then
                    If dblDrawdown >= dblLoanAmount Then
                        xmlLoanAmount.Text = "0"
                        dblDrawdown = dblDrawdown - dblLoanAmount
                    Else
                        xmlLoanAmount.Text = CStr(dblLoanAmount - dblDrawdown)
                        dblDrawdown = 0
                    End If
                End If
                
            Next
                                  
            Set objComboDO = New ComboDO
            strFinancialBenefitID = objComboDO.GetFirstComboValueId("IncentiveBenefitType", "F")
                                  
            'Get the mortgage incentives from the database
            Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
            xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "READ"
            xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCRUD"
            xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "MORTGAGEINCENTIVE"
            Set xmlRoot = xmlTempRequest.ownerDocument.createElement("MORTGAGEINCENTIVE")
            xmlSetAttributeValue xmlRoot, "APPLICATIONNUMBER", m_objXmlAssist.GetNodeValue(xmlCopyNode, "APPLICATIONNUMBER")
            xmlSetAttributeValue xmlRoot, "APPLICATIONFACTFINDNUMBER", m_objXmlAssist.GetNodeValue(xmlCopyNode, "APPLICATIONFACTFINDNUMBER")
            xmlSetAttributeValue xmlRoot, "MORTGAGESUBQUOTENUMBER", m_objXmlAssist.GetNodeValue(xmlCopyNode, "MORTGAGESUBQUOTENUMBER")
            xmlSetAttributeValue xmlRoot, "INCENTIVEBENEFITTYPE", strFinancialBenefitID
            xmlTempRequest.appendChild xmlRoot
            
            Set objCRUD = New omCRUD.omCRUDBO
            strResponse = objCRUD.OmRequest(xmlTempRequest.xml)
            Set xmlCRUDMortgageIncentive = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlCRUDMortgageIncentive.documentElement, , True
            
            dblTotalIncentiveAmount = 0
            For Each xmlIncentive In xmlCRUDMortgageIncentive.selectNodes("/RESPONSE/MORTGAGEINCENTIVE")
                Set xmlComponent = xmlCopyNode.selectSingleNode("LOANCOMPONENTLIST/LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER=" & m_objXmlAssist.GetAttributeFromNode(xmlIncentive, "LOANCOMPONENTSEQUENCENUMBER") & "]")
                                                
                'Get the related incentive from the database
                Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
                xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "READ"
                xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCRUD"
                xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "INCENTIVE"
                Set xmlRoot = xmlTempRequest.ownerDocument.createElement("INCENTIVE")
                xmlSetAttributeValue xmlRoot, "INCENTIVEGUID", m_objXmlAssist.GetAttributeFromNode(xmlIncentive, "INCENTIVEGUID")
                xmlTempRequest.appendChild xmlRoot
                
                strResponse = objCRUD.OmRequest(xmlTempRequest.xml)
                Set xmlCRUDIncentive = xmlLoad(strResponse, strFunctionName)
                errCheckXMLResponseNode xmlCRUDIncentive.documentElement, , True
                
                curIncentive = 0
                dblIncentivePercent = 0
                On Error Resume Next
                curIncentive = m_objXmlAssist.GetAttributeValue(xmlCRUDIncentive.firstChild, "INCENTIVE", "AMOUNT")
                dblIncentivePercent = m_objXmlAssist.GetAttributeValue(xmlCRUDIncentive.firstChild, "INCENTIVE", "PERCENTAGE")
                On Error GoTo CalcAndSaveMortgageCostsVbErr
                
                If curIncentive > 0 Then
                    dblTotalIncentiveAmount = dblTotalIncentiveAmount + curIncentive
                Else
                    dblIncentiveAmount = CSafeDbl(m_objXmlAssist.GetNodeValue(xmlComponent, "LOANAMOUNT")) * (dblIncentivePercent / 100) 'SR EP2_1798
                    dblTotalIncentiveAmount = dblTotalIncentiveAmount + dblIncentiveAmount
                    xmlSetAttributeValue xmlIncentive, "INCENTIVEAMOUNT", CStr(dblIncentiveAmount)
                End If
                
            Next
                        
            'Update the mortgage incentives in database
            Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
            xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "UPDATE"
            xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCRUD"
            xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "MORTGAGEINCENTIVE"
            For Each xmlIncentive In xmlCRUDMortgageIncentive.selectNodes("/RESPONSE/MORTGAGEINCENTIVE")
                xmlTempRequest.appendChild xmlIncentive.cloneNode(True)
            Next
            strResponse = objCRUD.OmRequest(xmlTempRequest.xml)
            Set xmlCRUDIncentive = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlCRUDIncentive.documentElement, , True
            
        Else
            dblTotalIncentiveAmount = objIMortgageSubquoteDO.TotalIncentivesForAllComponents(xmlCopyNode)
        End If
        
    Else
        dblTotalIncentiveAmount = objIMortgageSubquoteDO.TotalIncentivesForAllComponents(xmlCopyNode)
    End If


    '2 Calculate one-off costs
    
    Dim strOneOffCosts As String
    Dim xmlOneOffCostsDoc As FreeThreadedDOMDocument40
    
    Set xmlOneOffCosts = xmlGetNode(xmlElement, "ONEOFFCOSTLIST")

    ' PSC 09/11/2006 EP2_41 - Start
    If xmlOneOffCosts Is Nothing Then
        m_objXmlAssist.ChangeNodeName xmlCopyNode, "MORTGAGEINCENTIVE", "CHARGES"
            
        'EP2_1281 - Peter Edney
        'SR 20/20/2007 : EP2_1546 - clone node only if it exists
        Set xmlValuationType = vXMLRequest.selectSingleNode("//VALUATIONTYPE")
        If Not xmlValuationType Is Nothing Then
            xmlCopyNode.appendChild xmlValuationType.cloneNode(True)
        End If
        
        strOneOffCosts = objChargesBO.CalcCharges(xmlCopyNode.xml)
        m_objErrAssist.CheckXMLResponse strOneOffCosts, True, xmlResponseElem
        
        Set xmlOneOffCostsDoc = m_objXmlAssist.load(strOneOffCosts, TypeName(Me), strFunctionName)
        Set xmlTempNode = xmlOneOffCostsDoc.selectSingleNode(".//ONEOFFCOSTLIST")
        Set xmlOneOffCosts = xmlTempNode.cloneNode(True)
    Else
        Set xmlOneOffCostsDoc = New FreeThreadedDOMDocument40
        xmlOneOffCostsDoc.validateOnParse = False
        xmlOneOffCostsDoc.setProperty "NewParser", True
        xmlOneOffCostsDoc.appendChild xmlOneOffCosts.cloneNode(True)
    End If
    ' PSC 09/11/2006 EP2_41 - End
    
    blnCalcFeesOnly = xmlGetNodeAsBoolean(xmlElement, "CALCULATEFEESONLY")
    
    ' PSC 09/11/2006 EP2_41 - End
    ' IK EP2_1093
    ' If blnCalcFeesOnly = False Then
    
        'build request for GetAPROneOffCosts  JD bmids854
        Set xmlAPROneOffCostsDoc = New FreeThreadedDOMDocument40
        xmlAPROneOffCostsDoc.validateOnParse = False
        xmlAPROneOffCostsDoc.setProperty "NewParser", True
        Set xmlAPRReqNode = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
        Set xmlAPRMSQNode = xmlAPROneOffCostsDoc.createElement("MORTGAGESUBQUOTE")
        xmlAPRReqNode.appendChild xmlAPRMSQNode
        xmlAPRMSQNode.appendChild xmlAppNo
        xmlAPRMSQNode.appendChild xmlAppFFNo
        xmlAPRMSQNode.appendChild xmlMSQNo
        xmlAPRMSQNode.appendChild xmlOneOffCosts
        xmlAPROneOffCostsDoc.appendChild xmlAPRReqNode
        
        Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        blnRaiseError = False
        strResponse = objMortSubQuoteBO.GetAPROneOffCosts(xmlAPROneOffCostsDoc.documentElement.xml)
        Set xmlAPROneOffCostsDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlAPROneOffCostsDoc.documentElement, xmlResponseElem, True
        blnRaiseError = True
    
        '''' BMIDS854 END
        
        '3 Calculate Total Charges
        ' If one off costs are present then calculate charges
        Dim xmlFeeList As IXMLDOMNodeList
        Dim xmlFeeElement As IXMLDOMElement
        Dim dblTotalInitialCharges As Double
        Dim dblTotalCharges As Double
        Dim strPattern As String
                
        strPattern = ".//ONEOFFCOST[IDENTIFIER != ""CBC"" "
        strPattern = strPattern & "and IDENTIFIER != ""CPP"" "
        strPattern = strPattern & "and IDENTIFIER != ""TID""]"
                
        dblTotalCharges = 0
                
        If Not xmlOneOffCostsDoc Is Nothing Then
            
            Set xmlFeeList = xmlOneOffCostsDoc.selectNodes(strPattern)
                    
            For Each xmlFeeElement In xmlFeeList
                dblTotalCharges = dblTotalCharges + _
                    CSafeDbl(m_objXmlAssist.GetTagValue(xmlFeeElement, "AMOUNT"))
            Next
        
        End If
        
        'BMIDS881 GHun Add Adhoc fees to total charges
        If Not xmlAPROneOffCostsDoc Is Nothing Then
            For Each xmlFeeElement In xmlAPROneOffCostsDoc.selectNodes(strPattern)
                dblTotalCharges = dblTotalCharges + CSafeDbl(m_objXmlAssist.GetTagValue(xmlFeeElement, "AMOUNT"))
            Next
        End If
        
        dblTotalInitialCharges = 0
        strPattern = ".//ONEOFFCOST[IDENTIFIER = ""DEE"" or IDENTIFIER = ""SEA""]"
        'BMIDS881 End
    
        If Not xmlOneOffCostsDoc Is Nothing Then
            
            ' Calculate the total final charges based on Sealing Fee and Deeds Release Fee
            Set xmlFeeList = xmlOneOffCostsDoc.selectNodes(strPattern)
            
            For Each xmlFeeElement In xmlFeeList
                dblTotalInitialCharges = dblTotalInitialCharges + _
                    CSafeDbl(m_objXmlAssist.GetTagValue(xmlFeeElement, "AMOUNT"))
            Next
        End If
        
        'BMIDS881 GHun
        If Not xmlAPROneOffCostsDoc Is Nothing Then
            For Each xmlFeeElement In xmlAPROneOffCostsDoc.selectNodes(strPattern)
                dblTotalInitialCharges = dblTotalInitialCharges + CSafeDbl(m_objXmlAssist.GetTagValue(xmlFeeElement, "AMOUNT"))
            Next
        End If
        
        dblTotalInitialCharges = dblTotalCharges - dblTotalInitialCharges
        'BMIDS881 End
        
        ' PSC 09/11/2006 EP2_41 - Start
        dblAddedCharges = 0

        If Not xmlOneOffCostsDoc Is Nothing Then
            strPattern = ".//ONEOFFCOST[ADDTOLOAN='1']"

            ' Calculate added one off costs
            Set xmlFeeList = xmlOneOffCostsDoc.selectNodes(strPattern)
            
            For Each xmlFeeElement In xmlFeeList
                dblAddedCharges = dblAddedCharges + _
                    CSafeDbl(m_objXmlAssist.GetTagValue(xmlFeeElement, "AMOUNT"))
            Next
        End If
        ' PSC 09/11/2006 EP2_41 - End

        
        Dim dblTotalNetCharges As Double
        
        ' PSC 09/11/2006 EP2_41
        dblTotalNetCharges = dblTotalCharges + _
                            CSafeDbl(m_objXmlAssist.GetTagValue(xmlElement, "DEPOSIT")) _
                            - dblTotalIncentiveAmount - dblAddedCharges
        
        '4 Save One off costs to mortgageoneoffcosts in database
        Dim xmlOneOffCostsReturn As IXMLDOMElement
        'pull off the data we need to enable us to create mortgageoneoffcosts records
        Dim xmlCloneNode As IXMLDOMNode
        
        Dim xmlRootNode As IXMLDOMNode
        
        Set xmlRootNode = xmlOneOffCostsDoc.getElementsByTagName("ONEOFFCOSTLIST").Item(0)
        
        Set xmlOneOffCostList = xmlRootNode.cloneNode(True)    ' PSC 23/03/01 SYS2166
        
        Set xmlTempNode = xmlElement.getElementsByTagName("APPLICATIONNUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlRootNode.appendChild xmlCloneNode
        
        Set xmlTempNode = xmlElement.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlRootNode.appendChild xmlCloneNode
        
        Set xmlTempNode = xmlElement.getElementsByTagName("MORTGAGESUBQUOTENUMBER").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlRootNode.appendChild xmlCloneNode
            
        Set xmlTempNode = xmlElement.getElementsByTagName("APPLICATIONDATE").Item(0)
        
        If Not xmlTempNode Is Nothing Then
            Set xmlCloneNode = xmlTempNode.cloneNode(True)
            xmlRootNode.appendChild xmlCloneNode
        End If
        
        Set xmlTempNode = xmlElement.getElementsByTagName("CONTEXT").Item(0)
        Set xmlCloneNode = xmlTempNode.cloneNode(True)
        xmlRootNode.appendChild xmlCloneNode
                    
        m_objXmlAssist.ChangeNodeName xmlRootNode, "ONEOFFCOSTLIST", "MORTGAGESUBQUOTE"
            
        Set xmlOneOffCostsReturn = objIMortgageSubQuoteBO.CreateOneOffCosts(xmlRootNode)
        m_objErrAssist.CheckResponse xmlOneOffCostsReturn, xmlResponseElem, True
    
            
        '5 initialise mortgsgesubquote totals
        
        Set xmlTempNode = xmlElement.getElementsByTagName("TOTALNETMONTHLYCOST").Item(0)
        xmlTempNode.Text = "0"
        
        Set xmlTempNode = xmlElement.getElementsByTagName("TOTALGROSSMONTHLYCOST").Item(0)
        xmlTempNode.Text = "0"
            
        '6 call calculations engine
       
       'get the ApplicationDate from the ApplicationFactFind
       'the Mortgageproduct data is contained within the loan component
       'we also need the oneoffcosts this is all contained in xmlRootNode
       
    '    Dim xmlIllustrationResults As IXMLDOMElement
    '
    '    Dim xmlCalcsEngine As IXMLDOMElement
        Dim xmlLoanComponentList As IXMLDOMNodeList
        Dim xmlLoanComponentElement As IXMLDOMElement
    '    Dim xmlCalcsEngineInputInterface As IXMLDOMElement
    '
        Set xmlLoanComponentList = xmlElement.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
    '
        ' PSC 23/03/01 SYS2166 - Start
        ' Find main loan component
        Dim intMainLoanNo As Integer
        Dim intCurrentNo As Integer
        Dim strMainLoanNo As String
        Dim strValue As String
    
        intMainLoanNo = 9999
    
        For Each xmlLoanComponentElement In xmlLoanComponentList
            strValue = m_objXmlAssist.GetNodeValue(xmlLoanComponentElement, ".//LOANCOMPONENTSEQUENCENUMBER")
            intCurrentNo = CSafeInt(strValue)
    
            If intCurrentNo < intMainLoanNo Then
                intMainLoanNo = intCurrentNo
            End If
    
        Next
    
        strMainLoanNo = CStr(intMainLoanNo)
        ' PSC 23/03/01 SYS2166 - End
                
        Set xmlCompRateRequest = vXMLRequest.cloneNode(False)
        Set xmlCalcLoanCompList = xmlOut.createElement("LOANCOMPONENTLIST")
        xmlCompRateRequest.appendChild xmlCalcLoanCompList
    
        'EP2_954 GHun
        Set xmlAddBorrow = vXMLRequest.selectSingleNode("MORTGAGESUBQUOTE/MORTGAGELOANLIST")
        If Not xmlAddBorrow Is Nothing Then
            xmlCompRateRequest.appendChild xmlAddBorrow.cloneNode(True)
        End If
        'EP2_954 End
    
        For Each xmlLoanComponentElement In xmlLoanComponentList
    
            ' PSC 23/03/01 SYS2166 - Start
            Set xmlNode = xmlLoanComponentElement.cloneNode(True)
            
            strValue = m_objXmlAssist.GetNodeValue(xmlNode, _
                                                   ".//LOANCOMPONENTSEQUENCENUMBER")
    
    
            
            If strValue = strMainLoanNo Then
                ' JD BMIDS854 add the apr costs
                Set xmlAPROneOffCosts = xmlAPROneOffCostsDoc.selectNodes(".//ONEOFFCOST")
                For Each xmlOneOffCost In xmlAPROneOffCosts
                    xmlOneOffCostList.appendChild xmlOneOffCost
                Next
                xmlNode.appendChild xmlOneOffCostList
            End If
            ' PSC 23/03/01 SYS2166 - End
            
            xmlCalcLoanCompList.appendChild xmlNode
            
    '
    '        Set xmlCalcsEngineInputInterface = _
    '                        objIMortgageSubQuoteBO.BuildCalcsEngineInputInterface(xmlLoanComponentElement)
    '        'check return here
    '        objErrAssist.CheckResponse xmlCalcsEngineInputInterface, xmlResponseElem, True
    '
    '        Set xmlCalcsEngine = objIMortgageSubQuoteBO.RunCalcsEngine(xmlCalcsEngineInputInterface)
    '        'check return here
    '        objErrAssist.CheckResponse xmlCalcsEngine, xmlResponseElem, True
    '
    '        Set xmlTempNode = xmlLoanComponentElement.getElementsByTagName("MORTGAGEPRODUCTDETAILS").Item(0)
    '        'remove the mortgageproduct details
    '        xmlLoanComponentElement.removeChild xmlTempNode
    '
    '        xmlLoanComponentElement.appendChild xmlCalcsEngine
    '
    '        Set xmlIllustrationResults = objIMortgageSubQuoteBO.SaveIllustrationResults(xmlLoanComponentElement)
    '        objErrAssist.CheckResponse xmlIllustrationResults, xmlResponseElem, True
    '
    '        Dim objConvertAssist As ConvertAssist
    '        Set objConvertAssist = New ConvertAssist
    '        Dim dblTotalGrossMonthlyCost As Double
    '        Dim dblTotalNetMonthlyCost As Double
    '        Dim lngTotalAccruedInterest As Long
    '
    '        'add the returned grossmonthly cost to the existing totalgrossmonthly cost
    '        dblTotalGrossMonthlyCost = objConvertAssist.CSafeDbl(objXmlAssist.GetTagValue( _
    '                                xmlElement, "TOTALGROSSMONTHLYCOST")) _
    '                                + _
    '                                objConvertAssist.CSafeDbl(objXmlAssist.GetTagValue( _
    '                                xmlIllustrationResults, "GROSSMONTHLYCOST"))
    '
    '        'and set the totalgrossmonthly cost
    '        Set xmlTempNode = xmlElement.getElementsByTagName("TOTALGROSSMONTHLYCOST").Item(0)
    '        xmlTempNode.Text = dblTotalGrossMonthlyCost
    '
    '        'add the returned netmonthly cost to the existing totalnetmonthly cost
    '        dblTotalNetMonthlyCost = objConvertAssist.CSafeDbl( _
    '                                objXmlAssist.GetTagValue(xmlElement, "TOTALNETMONTHLYCOST")) _
    '                                + _
    '                                objConvertAssist.CSafeDbl(objXmlAssist.GetTagValue( _
    '                                xmlIllustrationResults, "NETMONTHLYCOST"))
    '
    '        'and set the totalnetmonthly cost
    '        Set xmlTempNode = xmlElement.getElementsByTagName("TOTALNETMONTHLYCOST").Item(0)
    '        xmlTempNode.Text = dblTotalNetMonthlyCost
    '
    '        'add the returned accrued interest to the existing totalaccrued interest
    '        lngTotalAccruedInterest = objConvertAssist.CSafeLng( _
    '                                objXmlAssist.GetTagValue(xmlElement, "TOTALACCRUEDINTEREST")) _
    '                                + _
    '                                objConvertAssist.CSafeLng(objXmlAssist.GetTagValue( _
    '                                xmlIllustrationResults, "ACCRUEDINTEREST"))
    '        'and set the total accrued interest
    '        Set xmlTempNode = xmlElement.getElementsByTagName("TOTALACCRUEDINTEREST").Item(0)
    '        xmlTempNode.Text = lngTotalAccruedInterest
    '
        Next
        
        'DPF 08/10/2002 - copy the request before calcs are done to use to calculate drawdown costs
        Set xmlDrawDown = xmlCompRateRequest.cloneNode(True)
        
        Set xmlCalcResponse = objIMortgageSubQuoteBO.RunCompositeRateCalcAndSaveResults(xmlCompRateRequest)
    
        m_objErrAssist.CheckResponse xmlCalcResponse, xmlResponseElem, True
    
        Dim xmlMortgageSubquote As IXMLDOMNode
            
        Set xmlMortgageSubquote = xmlElement.getElementsByTagName("MORTGAGESUBQUOTE").Item(0)
        
    '*****************************************************************************************
    'CMWP3 - DPF 12/07/02 - New call to be added here  to a new procedure called
    'CalcAndSaveRedemptionFees (Still within MortgageSubQuoteBO)
    'This will, as the name suggests calculate & then save the redemption fees.
    
        '7 Calculate and Save the Redemption Fees
        m_objXmlAssist.ChangeNodeName xmlCopyNode, "CHARGES", "MORTGAGESUBQUOTE"
        
        'BMIDS736 RedemptionFees are now saved within RunCompositeRateCalcAndSaveResults
        'Dim CalcAndSaveRedemptionFees As IXMLDOMElement
        'Set CalcAndSaveRedemptionFees = objIMortgageSubQuoteBO.CalcAndSaveRedemptionFees(xmlCopyNode)
        'objErrAssist.CheckResponse CalcAndSaveRedemptionFees, xmlResponseElem, True
        'BMIDS736 End
        
    '*****************************************************************************************
        
        '8 update mortgage sub quote record
        
        'MS 12/06/00 SYS0829 - Make sure we save the charges on calculate
        
        Set xmlTempNode = xmlElement.selectSingleNode(".//TOTALCHARGES")
        xmlTempNode.Text = dblTotalCharges
        
        Set xmlTempNode = xmlElement.selectSingleNode(".//TOTALINITIALCHARGES")
        xmlTempNode.Text = dblTotalInitialCharges
        
        Set xmlTempNode = xmlElement.selectSingleNode(".//TOTALNETCHARGES")
        xmlTempNode.Text = dblTotalNetCharges
        
        ' PSC 03/07/2002 BMIDS00062 - Start
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "TOTALNETMONTHLYCOST")
        Set xmlTempNode = xmlElement.selectSingleNode("TOTALNETMONTHLYCOST")
        xmlTempNode.Text = strValue
       
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "TOTALGROSSMONTHLYCOST")
        Set xmlTempNode = xmlElement.selectSingleNode("TOTALGROSSMONTHLYCOST")
        xmlTempNode.Text = strValue
        
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "TOTALACCRUEDINTEREST")
        Set xmlTempNode = xmlElement.selectSingleNode("TOTALACCRUEDINTEREST")
        xmlTempNode.Text = strValue
        ' PSC 03/07/2002 BMIDS00062 - End
        
        'BMIDS766 GHun
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "AMOUNTPERUNITBORROWED")
        Set xmlTempNode = xmlElement.selectSingleNode("AMOUNTPERUNITBORROWED")
        xmlTempNode.Text = strValue
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "APR")
        Set xmlTempNode = xmlElement.selectSingleNode("APR")
        xmlTempNode.Text = strValue
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "TOTALAMOUNTPAYABLE")
        Set xmlTempNode = xmlElement.selectSingleNode("TOTALAMOUNTPAYABLE")
        xmlTempNode.Text = strValue
        strValue = m_objXmlAssist.GetMandatoryElementText(xmlCalcResponse, "TOTALMORTGAGEPAYMENTS")
        Set xmlTempNode = xmlElement.selectSingleNode("TOTALMORTGAGEPAYMENTS")
        xmlTempNode.Text = strValue
        'BMIDS766 End
        
        'MAR1061 update the purchaseprice.
        m_objXmlAssist.ChangeNodeName xmlElement, "PURCHASEPRICE", "PURCHASEPRICEORESTIMATEDVALUE"
        
        objIMortgageSubquoteDO.Update xmlElement
        
    '*********************************************************************************************
    'CPWP1 - DPF 03/10/2002 - BM037 (Drawdowns)
    
        ' PJO 20/10/2003 - element needed for Product reservation
        Set xmlReserveProd = xmlDrawDown.cloneNode(True)
        '9 Calculate & Save Monthly Costs less Drawdown - for Flexible Products
        
        'First check if a drawdown actually exists and if so assign it to a variable
        'Dim dblDrawdown As Double  ' SR EP2_1798
        dblDrawdown = 0
        If IsNumeric(xmlElement.selectSingleNode("//DRAWDOWN").Text) Then
            dblDrawdown = xmlElement.selectSingleNode("//DRAWDOWN").Text
        End If
        
        'If a drawdown exists start doing calculations for monthly cost less drawdown
        If dblDrawdown > 0 Then
        
            Dim LoanComponentList As IXMLDOMNodeList
            Dim IntCompCount As Integer
            Dim dblInitialLAVal As Double
            Dim dblInitialTLCAVal As Double
            Dim dblLADrawDown As Double
            Dim dblTLCADrawDown As Double
                    
            IntCompCount = 1
            Set LoanComponentList = xmlDrawDown.selectNodes("//LOANCOMPONENT")
            IntCompCount = LoanComponentList.length
            
            dblLADrawDown = dblDrawdown
            dblTLCADrawDown = dblDrawdown
            
            'loop through loan components, starting with highest seq no and subtract the drawdown amount
            Do While IntCompCount > 0
                dblInitialLAVal = LoanComponentList.Item(IntCompCount - 1).selectSingleNode("LOANAMOUNT").Text
                dblInitialTLCAVal = LoanComponentList.Item(IntCompCount - 1).selectSingleNode("TOTALLOANCOMPONENTAMOUNT").Text
                
                'Adjust [Loan Amount] value
                If dblLADrawDown > dblInitialLAVal Then
                    dblLADrawDown = dblLADrawDown - dblInitialLAVal
                    dblInitialLAVal = 0
                Else
                    dblInitialLAVal = dblInitialLAVal - dblLADrawDown
                    dblLADrawDown = 0
                End If
                
                'Adjust [Total Loan Component Amount] value
                If dblTLCADrawDown > dblInitialTLCAVal Then
                    dblTLCADrawDown = dblTLCADrawDown - dblInitialTLCAVal
                    dblInitialTLCAVal = 0
                Else
                    dblInitialTLCAVal = dblInitialTLCAVal - dblTLCADrawDown
                    dblTLCADrawDown = 0
                End If
                
                LoanComponentList.Item(IntCompCount - 1).selectSingleNode("LOANAMOUNT").Text = dblInitialLAVal
                LoanComponentList.Item(IntCompCount - 1).selectSingleNode("TOTALLOANCOMPONENTAMOUNT").Text = dblInitialTLCAVal
                            
                IntCompCount = IntCompCount - 1
            Loop
            
            Dim xmlComponentElem As IXMLDOMElement
            
            For Each xmlComponentElem In xmlDrawDown.selectNodes("//LOANCOMPONENT")
                If xmlComponentElem.selectSingleNode("LOANAMOUNT").Text = "0" Then
                    xmlDrawDown.selectSingleNode("//LOANCOMPONENTLIST").removeChild xmlComponentElem
                End If
            Next
            
            ' EP2_9 - Check whether Any rows left to action.
            ' If so, action as before.
            If xmlDrawDown.selectNodes("//LOANCOMPONENT").length > 0 Then
                'Calculate monthly costs (less drawdown) and check response
                Set xmlDrawDownResponse = objIMortgageSubQuoteBO.RunCompositeRateCalc(xmlDrawDown)
                m_objErrAssist.CheckResponse xmlDrawDownResponse, xmlResponseElem, True
            
                'Update the MortgageSubQuote record with Monthly Costs - Less Drawdown
                Set xmlOldNode = m_objXmlAssist.GetNode(xmlElement, "//MONTHLYCOSTLESSDRAWDOWN")
                Set xmlNode = m_objXmlAssist.GetNode(xmlDrawDownResponse, "//TOTALNETMONTHLYCOST")
                xmlChangeNodeName xmlNode, "TOTALNETMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
                xmlElement.replaceChild xmlNode, xmlOldNode
                objIMortgageSubquoteDO.Update xmlElement
                
            Else ' If not then the Drawdown Amount = Total amount and we just need to set to 0.
                Set xmlOldNode = m_objXmlAssist.GetNode(xmlElement, "//MONTHLYCOSTLESSDRAWDOWN")
                xmlOldNode.Text = "0"
                objIMortgageSubquoteDO.Update xmlElement
                
                ' Dummy up the xmlDrawDownResponse node.
                ' Exact contents are irrelevant as long as it DOESNT contain LoanComponents.
                Set xmlDrawDownResponse = xmlElement.selectSingleNode("//DRAWDOWN")
                
            End If  ' End EP2_9 - Check whether Any rows left to action.
                        
                'Update each loan component record with monthly cost less drawdown amount
                Dim xmlResponseComp As IXMLDOMNode
                Dim xmlUpdateRequest As IXMLDOMNode
                Dim xmlLoanComponentNode As IXMLDOMNode
                Set xmlLoanComponentList = xmlElement.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
                
                Set xmlUpdateRequest = vXMLRequest.cloneNode(False)
               
                For Each xmlLoanComponentElement In xmlLoanComponentList
                    'build up request
                    Set xmlLoanComponentNode = xmlLoanComponentElement.cloneNode(False)
                    
                    ' PSC 14/12/2006 EP2_507 - Start
                    Set xmlNode = m_objXmlAssist.GetNode(xmlLoanComponentElement, ".//APPLICATIONNUMBER")
                    xmlLoanComponentNode.appendChild xmlNode.cloneNode(True)
                    
                    Set xmlNode = m_objXmlAssist.GetNode(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
                    xmlLoanComponentNode.appendChild xmlNode.cloneNode(True)
                    
                    Set xmlNode = m_objXmlAssist.GetNode(xmlElement, ".//MORTGAGESUBQUOTENUMBER")
                    xmlLoanComponentNode.appendChild xmlNode.cloneNode(True)
                    ' PSC 14/12/2006 EP2_507 - End
                    
                    strValue = m_objXmlAssist.GetNodeValue(xmlLoanComponentElement, "LOANCOMPONENTSEQUENCENUMBER")
                    Set xmlNode = m_objXmlAssist.GetNode(xmlLoanComponentElement, "LOANCOMPONENTSEQUENCENUMBER")
                    xmlLoanComponentNode.appendChild xmlNode
                    
                    ' PSC 14/11/2006 EP2_507 - Start
                    Set xmlResponseComp = xmlGetNode(xmlDrawDownResponse, "//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER = " & strValue & "]")
                    
                    If Not xmlResponseComp Is Nothing Then
                        Set xmlNode = m_objXmlAssist.GetNode(xmlResponseComp, "//NETMONTHLYCOST")
                        xmlChangeNodeName xmlNode, "NETMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
                    Else
                        Set xmlNode = xmlOut.createElement("MONTHLYCOSTLESSDRAWDOWN")
                        xmlNode.Text = "0.0"
                    End If
                    
                    xmlLoanComponentNode.appendChild xmlNode
    
                    xmlUpdateRequest.appendChild xmlLoanComponentNode
    
                    ' Update Loan Component
                    objIMortgageSubquoteDO.UpdateLoanComponent xmlUpdateRequest
    
                    ' remove loan comp from XML
                    xmlUpdateRequest.removeChild xmlLoanComponentNode
                    
                    
                    'check we have a MonthlyCostLessDrawDown amount to assign before we update the loan component
    '                For Each xmlComponentElem In xmlDrawDown.selectNodes("//LOANCOMPONENT")
    '                    If xmlComponentElem.selectSingleNode("LOANCOMPONENTSEQUENCENUMBER").Text = strValue Then
    '                        Set xmlResponseComp = m_objXmlAssist.GetNode(xmlDrawDownResponse, "//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER = " & strValue & "]")
    '
    '                        'BMIDS736 GHun
    '                        'Set xmlNode = m_objXmlAssist.GetNode(xmlResponseComp, "//OUT060_TOTALNETMORTGAGEPAYMENT1")
    '                        'xmlChangeNodeName xmlNode, "OUT060_TOTALNETMORTGAGEPAYMENT1", "MONTHLYCOSTLESSDRAWDOWN"
    '                        Set xmlNode = m_objXmlAssist.GetNode(xmlResponseComp, "//NETMONTHLYCOST")
    '                        xmlChangeNodeName xmlNode, "NETMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
    '                        'BMIDS736 End
    '
    '                        xmlLoanComponentNode.appendChild xmlNode
    '
    '                        xmlUpdateRequest.appendChild xmlLoanComponentNode
    '
    '                        'Update Loan Component
    '                        objIMortgageSubQuoteDO.UpdateLoanComponent xmlUpdateRequest
    '
    '                        'remove loan comp from XML
    '                        xmlUpdateRequest.removeChild xmlLoanComponentNode
    '                    End If
    '                Next
                    ' PSC 14/11/2006 EP2_507 - End
                    'empty out request
                    Set xmlLoanComponentNode = Nothing
                                      
                Next
                          
        End If   ' dblDrawdown > 0
    
    '** END of Change for CPWP1 ****************************************************************
    
    ' PJO 25/09/2003 BMIDS631 and BMIDS653
    ' (10) Reserve each loan element
        'MAR1703 check global parameter to see if we reserver products here.
        blnGlobalReserve = objGlobalParameterDO.GetCurrentParameterByType("ReserveMtgProductAtQuote", "BOOLEAN")
        If blnGlobalReserve = True Then
            blnReserve = True
            Set xmlLoanComponentList = xmlReserveProd.selectNodes("//LOANCOMPONENT")
            For Each xmlLoanComponentElement In xmlLoanComponentList
        
                Set xmlNode = xmlLoanComponentElement.cloneNode(True)
                m_objXmlAssist.ChangeNodeName xmlNode, "LOANCOMPONENT", "RESERVEDPRODUCT"
                            
                objIMortgageSubquoteDO.ReserveMortgageProduct xmlNode
            Next
        End If
        
        ' EP2_893 - New condition.
        blnCalcAndSaveProcFeesRequired = CalcAndSaveProcFeesRequired(vXMLRequest)
        If blnCalcAndSaveProcFeesRequired = True Then   'Continue with processing.
            ' PSC 04/12/2006 EP2_249
            CalcAndSaveProcFees vXMLRequest
        Else ' Throw a warning error message.
            lngWarningMsg = 8702
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngWarningMsg
        End If
        
'   IK EP2_1093
     If blnCalcFeesOnly Then
'    Else
        Set xmlMortgageLender = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGELENDER")
        xmlResponseElem.appendChild xmlOneOffCosts
        xmlResponseElem.appendChild xmlMortgageLender
    End If
    ' PSC 09/11/2006 EP2_41 - End
'   IK EP2_1093_ends
         
    Set IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalcAndSaveMortgageCostsExit:

    Set objIMortgageSubquoteDO = Nothing
    Set xmlNodeList = Nothing
    Set LoanComponentList = Nothing
    Set xmlElement = Nothing
    Set xmlDrawDown = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set objChargesBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlOneOffCostList = Nothing
    Set xmlMortgageLender = Nothing     ' PSC 09/11/2006 EP2_41
    Set xmlAddBorrow = Nothing  'EP2_954 GHun
    
    Exit Function

CalcAndSaveMortgageCostsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

' PJO 25/09/2003 BMIDS631
' Ignore duplicate keys on ReserveMortgageProduct
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeDuplicateKey Then
        If blnReserve Then
            Resume Next
        End If
    End If
    
    'JD BMIDS854 check for record not found
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_CalcAndSaveMortgageCosts = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAndSaveMortgageCostsExit

End Function


Private Function CreateNewMortgageOneOffCostXML(ByRef xmlElement As IXMLDOMElement, _
                                ByVal intNewOneOffCostSequenceNumber As Integer, _
                                ByRef xmlOneOffCost As IXMLDOMElement, _
                                ByVal strComboValueId As String) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewMortgageOneOffCostXMLVbErr

    Const strFunctionName As String = "CreateNewMortgageOneOffCostXML"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlOneOffCostElem As IXMLDOMElement
    
    Dim strTagName As String
                
    If IsMissing(intNewOneOffCostSequenceNumber) Or _
        IsMissing(strComboValueId) Then
                        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "intNewOneOffCostSequenceNumber = " & intNewOneOffCostSequenceNumber & _
                                "strComboValueId = " & strComboValueId
    End If
                
    Set xmlOneOffCostElem = xmlOut.createElement("MORTGAGEONEOFFCOST")
    xmlOut.appendChild xmlOneOffCostElem
                
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlOneOffCostElem.appendChild xmlTempElement
        
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlOneOffCostElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, strTagName)
    xmlOneOffCostElem.appendChild xmlTempElement
    
    strTagName = "MORTGAGEONEOFFCOSTTYPE"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = strComboValueId
    xmlOneOffCostElem.appendChild xmlTempElement
    
    strTagName = "ONEOFFCOSTSEQUENCENUMBER"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = CStr(intNewOneOffCostSequenceNumber)
    xmlOneOffCostElem.appendChild xmlTempElement
        
    strTagName = "ADDTOLOAN"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlOneOffCost, "ADDEDTOLOAN")
    xmlOneOffCostElem.appendChild xmlTempElement
        
    strTagName = "AMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlOneOffCost, strTagName)
    xmlOneOffCostElem.appendChild xmlTempElement
    
 'DRC BMIDS767 Rev - Start
    strTagName = "ADHOCIND"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = m_objXmlAssist.GetTagValue(xmlOneOffCost, strTagName, False, True)
    xmlOneOffCostElem.appendChild xmlTempElement
    
'DRC BMIDS767 Rev - End

    'MAR28 Add Refund Amount
    strTagName = "REFUNDAMOUNT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    xmlTempElement.Text = objXmlAssist.GetTagValue(xmlOneOffCost, strTagName)
    xmlOneOffCostElem.appendChild xmlTempElement
         
    Dim strIdentifier As String
    strTagName = "IPT"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strIdentifier = objXmlAssist.GetTagValue(xmlOneOffCost, "IDENTIFIER")
        
    If StrComp(strIdentifier, "MIG", vbTextCompare) = 0 Then

        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlElement, "MIGIPT")

    Else
    
        xmlTempElement.Text = "0"
    
    End If

    xmlOneOffCostElem.appendChild xmlTempElement
        
    strTagName = "DESCRIPTION"
    Set xmlTempElement = xmlOut.createElement(strTagName)
    strIdentifier = objXmlAssist.GetTagValue(xmlOneOffCost, "IDENTIFIER")
        
    If StrComp(strIdentifier, "OTH", vbTextCompare) = 0 Then

        xmlTempElement.Text = objXmlAssist.GetTagValue(xmlOneOffCost, "NAME")
    
    Else
    
        xmlTempElement.Text = "NULL"
        
    End If

    xmlOneOffCostElem.appendChild xmlTempElement
        
    Set CreateNewMortgageOneOffCostXML = xmlOut.documentElement
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CreateNewMortgageOneOffCostXMLVbErr:
    
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteTxBO_Update(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else

    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
    
    Dim xmlMortgageSubQuoteUpdate As IXMLDOMElement
    
    objIMortgageSubquoteDO.Update xmlElement

    Set IMortgageSubQuoteTxBO_Update = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageSubQuoteExit:

    Set objIMortgageSubquoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateMortgageSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_Update = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit

End Function

Private Function IMortgageSubQuoteTxBO_UpdateLoanComponent(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteTxBO_UpdateLoanComponent"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrLOANCOMPONENT_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrLOANCOMPONENT_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrLOANCOMPONENT_NAME & " tag not found"
    End If
   
    Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    
    objIMortgageSubquoteDO.UpdateLoanComponent xmlElement
    
    Set IMortgageSubQuoteTxBO_UpdateLoanComponent = xmlResponseElem
    
    m_objContext.SetComplete
    
UpdateLoanComponentExit:
    
    Set objIMortgageSubquoteDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    
    Exit Function

UpdateLoanComponentVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteTxBO_UpdateLoanComponent = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateLoanComponentExit

End Function

Private Function CalcAndSaveProcFeesRequired(ByVal vXMLRequest As IXMLDOMNode) As Boolean
' EP2_893 - Return boolean as to whether CalcAndSaveProcFees is required to run.
On Error GoTo CalcAndSaveProcFeesRequiredVbErr


    Const cstrFunctionName As String = "CalcAndSaveProcFeesRequired"
        
    'Declare our Variables.
    Dim xmlIntroducerProcFees As IXMLDOMNode
    Dim strSubmissionRoute As String
    Dim objErrAssist As New ErrAssist
  
    
    ' Set the Node and get the Submission Route value.
    Set xmlIntroducerProcFees = xmlGetNode(vXMLRequest, "MORTGAGESUBQUOTE/INTRODUCERPROCFEEDETAILS")
    If Not xmlIntroducerProcFees Is Nothing Then
        strSubmissionRoute = xmlGetNodeText(xmlIntroducerProcFees, "SUBMISSIONROUTE")
    End If
    
    ' Check whether Value found.
    If strSubmissionRoute = "" Then
        CalcAndSaveProcFeesRequired = False
    Else
        CalcAndSaveProcFeesRequired = True
    End If
   
    
    Set xmlIntroducerProcFees = Nothing

    Exit Function

CalcAndSaveProcFeesRequiredVbErr:
    
     objErrAssist.AddToErrSource cstrFunctionName
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Function
        
        
' PSC 04/12/2006 EP2_249 - Start
Private Sub CalcAndSaveProcFees(ByVal vXMLRequest As IXMLDOMNode)
On Error GoTo CalcAndSaveProcFeesVbErr

    Const cstrFunctionName As String = "CalcAndSaveProcFees"
        
    Dim xmlIntroducerProcFees As IXMLDOMNode
    Dim xmlProcFees As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlNatureOfLoan As IXMLDOMNode
    Dim xmlComponentList As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlProcFee As IXMLDOMNode
    Dim xmlAssociationFee As IXMLDOMNode
    Dim xmlPackagingFee As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlLoanComponent As IXMLDOMNode

    Dim xmlLoanComponents As IXMLDOMNodeList
    
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    
    Dim objCalcsBO As CalcsBO
    Dim objIMortgageSubQuote As IMortgageSubQuoteBO
    Dim objCRUD As omCRUDBO
    
    Dim strSubmissionRouteId As String
    Dim strResponse As String
    Dim strExclusivityLoading As String
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    
    Set xmlNatureOfLoan = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/NATUREOFLOAN").cloneNode(True)
    Set xmlIntroducerProcFees = xmlGetNode(vXMLRequest, "MORTGAGESUBQUOTE/INTRODUCERPROCFEEDETAILS")
    Set xmlProcFees = xmlIntroducerProcFees.cloneNode(True)
    xmlChangeNodeName xmlProcFees, "INTRODUCERPROCFEEDETAILS", "PROCFEES"
    
    strSubmissionRouteId = xmlGetMandatoryNodeText(xmlProcFees, "SUBMISSIONROUTEID")
    
    xmlChangeNodeName xmlNatureOfLoan, "NATUREOFLOAN", "PRODUCTCATEGORY"
    
    xmlProcFees.appendChild xmlNatureOfLoan
    
    Set xmlComponentList = xmlResponseDoc.createElement("LOANCOMPONENTLIST")
    xmlProcFees.appendChild xmlComponentList
    
    Set xmlLoanComponents = vXMLRequest.selectNodes("MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT[MANUALPORTEDLOANIND!='1' and PRODUCTSWITCHRETAINPRODUCTIND!='1']")
    
    ' Setup crud request to get loadings
    Set objCRUD = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCrudBO")

    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "READ"
    xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCM"
    xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "MORTGAGEPRODUCTEXCLUSIVITY"
    Set xmlRoot = xmlResponseDoc.createElement("MORTGAGEPRODUCTEXCLUSIVITY")
    xmlSetAttributeValue xmlRoot, "FIRMID", strSubmissionRouteId
    xmlTempRequest.appendChild xmlRoot

    For Each xmlLoanComponent In xmlLoanComponents
    
        strExclusivityLoading = "0"
        
        Set xmlComponent = xmlResponseDoc.createElement("LOANCOMPONENT")
        Set xmlNode = xmlResponseDoc.createElement("LOANAMOUNT")
        xmlNode.Text = xmlGetNodeText(xmlLoanComponent, "LOANAMOUNT")
        xmlComponent.appendChild xmlNode
        
        xmlSetAttributeValue xmlRoot, "MORTGAGEPRODUCTCODE", xmlGetMandatoryNodeText(xmlLoanComponent, "MORTGAGEPRODUCTCODE")
        strResponse = objCRUD.OmRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True

        Set xmlNode = xmlGetNode(xmlResponseDoc, "RESPONSE/MORTGAGEPRODUCTEXCLUSIVITY")
                
        If Not xmlNode Is Nothing Then
            strExclusivityLoading = xmlGetAttributeText(xmlNode, "PROCFEELOADING", "0")
        End If

        Set xmlNode = xmlResponseDoc.createElement("EXCLUSIVELOADING")
        xmlNode.Text = strExclusivityLoading
        xmlComponent.appendChild xmlNode
        
        xmlComponentList.appendChild xmlComponent
    Next
    
    Set xmlNode = xmlResponseDoc.createElement("PRODUCTSCHEME")
    xmlNode.Text = xmlGetNodeText(xmlLoanComponents(0), ".//SPECIALGROUP/GROUPTYPESEQUENCENUMBER")
    xmlProcFees.appendChild xmlNode
    
    Set xmlNode = xmlResponseDoc.createElement("LOANAMOUNT")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, ".//AMOUNTREQUESTED")
    xmlProcFees.appendChild xmlNode
    
    Set xmlNode = xmlResponseDoc.createElement("LTV")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, ".//LTV")
    xmlProcFees.appendChild xmlNode
    
    Set xmlNode = xmlResponseDoc.createElement("TYPEOFAPPLICATION")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, ".//TYPEOFAPPLICATION")
    xmlProcFees.appendChild xmlNode
   
    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    xmlTempRequest.appendChild xmlProcFees
    
    Set objCalcsBO = m_objContext.CreateInstance(App.Title & ".CalcsBO")
    
    strResponse = objCalcsBO.CalculateProcFees(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    
    Set xmlProcFee = xmlGetNode(xmlResponseDoc, "RESPONSE/PROCFEE")
    Set xmlAssociationFee = xmlGetNode(xmlResponseDoc, "RESPONSE/ASSOCIATIONFEE")
    Set xmlPackagingFee = xmlGetNode(xmlResponseDoc, "RESPONSE/PACKAGINGFEE")
    
    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    Set xmlRoot = xmlResponseDoc.createElement("MORTGAGEINTRODUCERFEE")
    xmlTempRequest.appendChild xmlRoot
    
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/APPLICATIONNUMBER")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/APPLICATIONFACTFINDNUMBER")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(vXMLRequest, "MORTGAGESUBQUOTE/MORTGAGESUBQUOTENUMBER")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "SUBMISSIONROUTE")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "SUBMISSIONROUTEID")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "PACKAGERASSOCIATIONID")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "MORTGAGECLUBID")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "PACKAGERID")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    
    ' PSC 05/02/2007 EP2_1219 - Start
    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "RECIPIENTID")
    xmlRoot.appendChild xmlNode.cloneNode(True)

    Set xmlNode = xmlGetMandatoryNode(xmlIntroducerProcFees, "RECIPIENTTYPE")
    xmlRoot.appendChild xmlNode.cloneNode(True)
    ' PSC 05/02/2007 EP2_1219 - End

    xmlRoot.appendChild xmlProcFee
    xmlRoot.appendChild xmlAssociationFee
    xmlRoot.appendChild xmlPackagingFee
    
    Set objIMortgageSubQuote = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
    Set xmlTempResponse = objIMortgageSubQuote.SaveIntroducerFees(xmlTempRequest)
    errCheckXMLResponseNode xmlTempResponse, , True
     
    Set xmlIntroducerProcFees = Nothing
    Set xmlProcFees = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNatureOfLoan = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    Set xmlProcFee = Nothing
    Set xmlAssociationFee = Nothing
    Set xmlPackagingFee = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlResponseDoc = Nothing
    Set objCalcsBO = Nothing
    Set objIMortgageSubQuote = Nothing
    Set objCRUD = Nothing

    Exit Sub

CalcAndSaveProcFeesVbErr:
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    Set xmlIntroducerProcFees = Nothing
    Set xmlProcFees = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNatureOfLoan = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    Set xmlProcFee = Nothing
    Set xmlAssociationFee = Nothing
    Set xmlPackagingFee = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlResponseDoc = Nothing
    Set objCalcsBO = Nothing
    Set objIMortgageSubQuote = Nothing
    Set objCRUD = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.number, Err.Source, Err.Description

End Sub
        
        

Private Sub CreateComponentsFromExistingAcc(ByVal vXMLRequest As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure to create Components

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
End Sub
        
' PSC 04/12/2006 EP2_249 - End

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


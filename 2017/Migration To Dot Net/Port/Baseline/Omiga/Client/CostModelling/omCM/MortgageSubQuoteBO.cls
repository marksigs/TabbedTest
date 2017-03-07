VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MortgageSubQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageSubQuoteBO.cls
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
'               e.g. Omiga4TxBO, MortgageSubQuoteDO
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
'PSC    21/10/99    Amend warning processing
'MCS    04/01/2000  Added interfacing and validate sub quote
'MCS    13/01/2000  Calcs stuff moved from MortgageCostsBO
'JLD    14/03/00    Added Update method to BO
'JLD    20/03/00    skip record not found error for loancomponents in ValidateSubQuote
'DJP    28/03/00    Fixed variable not set error in IMortgageSubQuoteBO_ProcessAddedOneOffCost
'SR     06/04/00   Modified method 'RecostComponent' based on the new spec (01.01.05)
'JLD    07/04/00    Error in ValidateSubQuote method
'SR     26/04/00    New method 'GetData'
'JLD    26/04/00    If Mig premium is 0 then there will not be a MIG type one off cost. (ProcessAddedOneOffCost)
'JLD    02/05/00    use TOTALLOANCOMPONENTAMOUNT as INP020_LOANAMOUNT in BuildCalcEngineInputInterface.
'                   LTV and TOTALLOANAMOUNT required for CalcMIG method.
'JLD    08/06/00    Added string interface method for CreateFirstSubquote for use with homezone.
'APS    08/06/00    SYS0833 - You cannot have a valid Mortgage subquote without loan components
'DLM    10/07/00    AQR SYS0949 It is not an error to have a mortgage subquote without loan components.
'MS     11/07/00    SYS0924 add Return tag for ported loan
' MS    21/07/00    performance rework.
'DLM    17/08/00    AQR SYS0924 add response tag for ported loan xml
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'BG     30/11/00    AQR SYS0924 - added code to IMortgageSubQuoteBO.ValidateSubQuote to check for
'                   "Nothing" to stop "object variable or with block not set" error.
'PSC    19/03/01    SYS2053 Add string based method for CreateOneOffCosts,
'                   BuildCalcsEngineInputInterface, SaveIllustrationResults and RunCalcsEngine
'                   Also Add UpdateLoanComponent
'AS     09/07/01    SYS2463 XML Parser 3 work around
'SR     06/09/01    SYS2412 - New function 'FindMortgageIncentiveList'
'MC     14/09/01    SYS2250 - New Base Rate Table Structure
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'MC     12/10/01    SYS2250 - Temporarily remove new Base Rate functionality for Build
'MC     06/12/01    SYS2250 - Reinstate Base Rate functionality
'MDC    10/12/01    SYS3440 - Rate Change amendments
'STB    22/03/02    SYS3557 - DevonRulingInd altered so repayment figures returned for multiple interest rate periods.
'SG     28/05/02    SYS7467 - MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit

'BMIDS736 GHun No longer used
'Private Const cintMAX_INTEREST_RATES = 5
'Private Const cintMAX_INTEREST_RATE_BANDS = 5
'Private Const cstrSPM_GROUP = "MortCalcDebugOptions"
'Private Const cstrSPM_PROPERTY = "DebugFileName"
'BMIDS736 End
Private Const cstrROOT_NODE_NAME As String = "MORTGAGESUBQUOTE"
Private Const cstrLOANCOMP_NODE_NAME As String = "LOANCOMPONENT"
Private Const cstrMORTGAGEINCENTIVE_NODE_NAME As String = "MORTGAGEINCENTIVE"
Private Const cstrMTGINTRODUCERFEE_NODE_NAME As String = "MORTGAGEINTRODUCERFEE"    ' PSC 30/11/2006 EP2_249


Implements IMortgageSubQuoteBO
Implements ObjectControl
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

' PSC 03/07/2002 BMIDS00062
Private Const cstrLOANCOMPLIST_NODE_NAME = "LOANCOMPONENTLIST"

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'PSC    11/06/2002  BMIDS00043  Change BuildCalcsEnginInputInterface and SetupCharges
'                               for BMids specific calculation
'MDC    17/06/2002  BMIDS00074  CMWP6 BM016 - Move BaseRateSet from MortgageProduct
'                               to InterestRateType
'MV     20/06/2002  BMIDS00077  Core Upgrade - Modified IMortgageSubQuoteBO_BuildCalcsEngineInputinterface
'                               SetUpPanelInterestRateDetails
'PSC    03/07/2002  BMIDS00062  Add CalculateInterestRate
'MV     18/07/2002  BMIDS00179  Core Upgrade Rollback - Modified IMortgageSubQuoteBO_ValidateSubQuote
'DPF    19/07/2002  CMWP3       Created method IMortgageSubQuoteBO_CalcAndSaveRedemptionFees()
'PSC    29/07/2002  BMIDS00261  Amend BuildCalcsEngineInputInterface to setup INP190_INTERESTRATE2BAND4 correctly
'MV     16/09/2002  BMIDS00458  Modified IMortgagesubQuoteBO_ProcessOneOffCosts()
'MV     07/10/2002  BMIDS00588  Modified IMortgageSubQuoteBO_ProcessAddedOneOffCost()
'MV     11/10/2002  BMIDS00590  Amended IMortgageSubQuoteBO_ProcessAddedOneOffCost()
'DPF    23/10/2002  CPWP1       Added FindAvailableIncentives & IMortgageSubQuoteBO_FindAvailableIncentives()
'MV     23/10/2002  BMIDS00667  Removed Stage6 in IMortgageSubQuoteBO_ProcessAddedOneOffCost
'SA     24/10/2002  BMIDS00515  ValidateSubQuote altered to cater for LTV
'MV     28/10/2002  BMIDS00724  Amended IMortgageSubQuoteBO_ProcessAddedOneOffCost
'DPF    30/10/2002  BMIDS00786  Amended IMortgageSubQuoteBO_FindAvailableIncentives() to fix bugs
'MO     31/10/2002  BMIDS00765  Amended ValidateSubQuote to check for the existance of the LTV node.
'DPF    04/11/2002  BMIDS00770  Amended IMortgageSubQuoteBO_FindAvailableIncentives - calculation
'                               of amount remaining on incentive for percentage based incentives.
'DPF    11/11/2002  BMIDS00808  Added two new functions.  CalculateDrawDownCosts & IMortgageSubQuoteBO_CalculateDrawDownCosts
'                               to calculate monthly costs where a drawdown exists on a flexible mortgage
'PSC    13/11/2002  BMIDS00931  Amend BuildCompositeRateInterface as Part and Parts do not always have Interest Only Elements
'                               and Capital And Interest Elements at interface time as it is based on disbursement amount
'PSC    14/11/2002  BMIDS00842  Amend RunCompositeRateCalc to take into account rounding differences when apportioning
'                               last loan component values
'SA     15/11/2002  BMIDS00952  Amend ValidateSubQuote so it no longer takes IsSelfEmployed indicator into account
'DPF    15/11/2002  BMIDS00805  Amended IMortgageSubQuoteBO_FindAvailableIncentives to check the product for the active
'                               loan component hasn't been altered since the last call
'PSC    19/11/2002  BMIDS00989  Amend RunCompositeRateCalc to take into account calcs returning zero in output
'SR     02/01/2003  BM0209      Modified IMortgageSubQuoteBO_GetLoanCompositionDetails
'MDC    02/01/2003  BM0211      Rate Change Fix in IMortgageSubQuoteBO_GetLoanCompositionDetails
'MDC    08/01/2003  BM0230      Failed Rate Change batches
'INR    06/08/2003  BMIDS624    ICC001 Update LoanComponent to include RESOLVEDRATE
'GHun   28/10/2003  BMIDS624    ICC001 Moved setting ResolvedRate from BuildCompositeRateInterface to RunCompositeRateCalc
'GHun   03/11/2003  BMIDS622    Clean up of some methods used by RateChange
'GHun   20/04/2004  BMIDS736    CORE49 Alpha+ integration
'MC     12/06/2004  BMIDS763    APPLICATIONDATE IS PASSING TO LoancompositionDetails() Method.
'DRC    22/06/2004  BMIDS767    Pass only One Off Costs with APR validation type in BuildCalcsEngineInputInterface
'MC     06/07/2004  BMIDS767    ApplicationDate passed to GetFullProductDetails() and OneOffCostCalculation methods
'INR    08/07/2004  BMIDS774    May not be any OneOffCosts added to the loan, check OneOffCostList exists before trying to use it.
'GHun   08/07/2004  BMIDS782    Cost Modelling performance
'GHun   09/07/2004  BMIDS766    Amended RunCompositeRateCalcAndSaveResults
'GHun   29/07/2004  BMIDS821    Amended FindAvailableIncentives
'JD     09/08/2004  BMIDS826    Changed error message text for 222 in BuildCalcsEngineInputInterface
'JD     24/08/2004  BMIDS854    Added method GetAPROneOffCosts
'GHun   07/09/2004  BMIDS815    Amended GetLoanCompositionDetails
'GHun   07/09/2004  BMIDS815    Amended ProcessAddedOneOffCost & GetLoanCompositionDetails
'GHun   08/09/2004  BMIDS815    Amended BuildCalcsEngineInputInterface
'GHun   08/09/2004  BMIDS815    Amended ProcessAddedOneOffCost & GetLoanCompositionDetails
'GHun   25/10/2004  BMIDS929    Changed BuildCalcsEngineInputInterface and added MonthDiff
'GHun   28/10/2004  BMIDS936    Changed BuildCalcsEngineInputInterface
'KRW    09/11/2004  BMIDS942    Changed AdjustForCompositeRate
'------------------------------------------------------------------------------------------
'ING Specific History:
'
'Prog   Date        AQR         Description
'GHun   07/07/2005  MAR9        Removed code for CompositeRate adjustment
'GHun   05/10/2005  MAR46       Changed GetLoanCompositionDetails for porting
'GHun   15/11/2005  MAR355      Changed GetLoanCompositionDetails
'JD     08/03/2006  MAR1061     GetLoanCompositionDetails - use purchaseprice from MSQ
'GHun   15/05/2006  MAR1777     Changed BuildCalcsEngineInputInterface and added
'------------------------------------------------------------------------------------------
'Epsom Specific History:
'
'Prog   Date        AQR         Description
'PSC    09/11/2006  EP2_41      Amend CalcAndSaveMortgageCosts to return one off costs
'PSC    30/11/2006  EP2_249     Added SaveIntroducerFees
'PSC    05/02/2007  EP2_1219    Amend SaveIntroducerFees to save recipientid and recipienttype
'                               differently
'GHun   12/02/2007  EP2_954     Changes for additional borrowing
'PSC    13/02/2007  EP2_1288    Amend SaveIntroducerFees to pay packaging fee to packager
'PSC    19/02/2007  EP2_1488    Amend GetLoanCompositionDetails to pass in ORIGINALLTV
'                               to GetFullProductDetails
'SR     27/02/2007  EP2_1272    Modified method 'IMortgageSubQuoteBO_GetloanCompositionDetails'
'SR     06/03/2007  EP2_1613    modified ProcessAddedOneOffCosts. Add total of OneOffCosts to
'                               appropriate loanComponent incase of ProductSwitch or ProterdLoan
'PSC    15/03/2007  EP2_1956    Added CopyIntroducerFees
'PSC    30/03/2007  EP2_2149    Amend GetLoanCompositionDetails to use existing ORIGINALLTV if present
'------------------------------------------------------------------------------------------

Public Function ValidateCompulsoryProducts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCompulsoryProductsVbErr

    Const strFunctionName As String = "ValidateCompulsoryProducts"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_ValidateCompulsoryProducts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateCompulsoryProducts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateCompulsoryProductsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateCompulsoryProductsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ValidateCompulsoryProducts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateCompulsoryProductsExit

End Function
Public Function GetAPROneOffCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  returns a list of mortgageoneoffcosts which need to be included in any APR calculations
' pass:         vstrXmlRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetAPROneOffCostsVbErr

    Const strFunctionName As String = "GetAPROneOffCosts"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetAPROneOffCosts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAPROneOffCosts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetAPROneOffCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAPROneOffCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetAPROneOffCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAPROneOffCostsExit

End Function
Public Function CreateComponentsFromExistingAcc(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateComponentsFromExistingAcc
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

End Function

Public Function CalcAndSaveMortgageCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CalcAndSaveMortgageCosts a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSaveMortgageCostsVbErr

    Const strFunctionName As String = "CalcAndSaveMortgageCosts"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomCalcAndSaveMortgageCosts)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_CalcAndSaveMortgageCosts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalcAndSaveMortgageCosts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalcAndSaveMortgageCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalcAndSaveMortgageCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CalcAndSaveMortgageCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAndSaveMortgageCostsExit

End Function


Public Function RecostComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'<REQUEST>
'   <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'   <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'   <CONTEXT>..</CONTEXT>
'   <ONEOFFCOST>
'       ....
'       ....
'   </ONEOFFCOST>
'   <LOANCOMPONENT>
'       ....
'       ....
'   </LOANCOMPONENT>
'   <APPLICATIONDATE>..</APPLICATIONDATE>
'   <PURCHASEPRICE>...</PURCHASEPRICE>
'   <TYPEOFAPPLICATION>...</TYPEOFAPPLICATION>
'   <LOCATION>...</LOCATION>
'   <TYPEOFVALUATION>..</TYPEOFVALUATION>
'   <FEETYPE>...</FEETYPE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo RecostComponentVbErr

    Const strFunctionName As String = "RecostComponent"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_RecostComponent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    RecostComponent = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecostComponentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RecostComponentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    RecostComponent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostComponentExit

End Function

Public Function ProcessAddedOneOffCost(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ProcessAddedOneOffCostVbErr

    Const strFunctionName As String = "ProcessAddedOneOffCost"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_ProcessAddedOneOffCost(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ProcessAddedOneOffCost = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ProcessAddedOneOffCostExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ProcessAddedOneOffCostVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ProcessAddedOneOffCost = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostExit

End Function

Public Function ValidateSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateSubQuoteVbErr

    Const strFunctionName As String = "ValidateSubQuote"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_ValidateSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateSubQuote = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ValidateSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateSubQuoteExit

End Function

Public Function GetOneOffCostsDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetOneOffCostsDetailsVbErr

    Const strFunctionName As String = "GetOneOffCostsDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetOneOffCostDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetOneOffCostsDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetOneOffCostsDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetOneOffCostsDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetOneOffCostsDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetOneOffCostsDetailsExit

End Function

Public Function GetMortgageSubQuoteData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageSubQuoteDataVbErr

    Const strFunctionName As String = "GetMortgageSubQuoteData"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetMortgageSubQuoteData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageSubQuoteData = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetMortgageSubQuoteDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetMortgageSubQuoteDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetMortgageSubQuoteData = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageSubQuoteDataExit

End Function

Public Function GetLoanCompositionDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanCompositionDetailsVbErr

    Const strFunctionName As String = "GetLoanCompositionDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetLoanCompositionDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLoanCompositionDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetLoanCompositionDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLoanCompositionDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoanCompositionDetailsExit

End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDefaultsForNewLoanComponentVbErr

    Const strFunctionName As String = "GetDefaultsForNewLoanComponent"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetDefaultsForNewLoanComponent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetDefaultsForNewLoanComponent = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDefaultsForNewLoanComponentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetDefaultsForNewLoanComponentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDefaultsForNewLoanComponentExit

End Function

Public Function FindActiveLoanComponentsList(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindActiveLoanComponentsListVbErr

    Const strFunctionName As String = "FindActiveLoanComponentsList"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_FindActiveLoanComponentsList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindActiveLoanComponentsList = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindActiveLoanComponentsListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindActiveLoanComponentsListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindActiveLoanComponentsList = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindActiveLoanComponentsListExit

End Function

Public Function GetLoanComponentData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanComponentDataVbErr

    Const strFunctionName As String = "GetLoanComponentData"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetLoanComponentData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLoanComponentData = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetLoanComponentDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLoanComponentDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetLoanComponentData = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoanComponentDataExit

End Function

Public Function ReserveMortgageProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  ReserveMortgageProduct a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ReserveMortgageProductVbErr

    Const strFunctionName As String = "ReserveMortgageProduct"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomReserveMortgageProduct)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_ReserveMortgageProduct(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ReserveMortgageProduct = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReserveMortgageProductExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ReserveMortgageProductVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ReserveMortgageProduct = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReserveMortgageProductExit

End Function

Public Function SaveLoanComponentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveLoanComponentDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveLoanComponentDetailsVbErr

    Const strFunctionName As String = "SaveLoanComponentDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomSaveLoanComponentDetails)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_SaveLoanComponentDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveLoanComponentDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveLoanComponentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveLoanComponentDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    SaveLoanComponentDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLoanComponentDetailsExit

End Function

Public Function CreateFirstSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateFirstSubQuote"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_CreateFirstSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateFirstSubQuote = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateFirstSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function DeleteLoanComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteLoanComponent a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLoanComponentVbErr

    Const strFunctionName As String = "DeleteLoanComponent"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_DeleteLoanComponent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteLoanComponent = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteLoanComponentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLoanComponentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    DeleteLoanComponent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLoanComponentExit

End Function

Public Function ResetMortgageSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  ResetMortgageSubQuote a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "ResetMortgageSubQuote"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomResetMortgageSubQuote)

    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_ResetMortgageSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ResetMortgageSubQuote = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ResetMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ResetMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit
    
End Function

Public Function Update(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "Update"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomResetMortgageSubQuote)

    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Update(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Update = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Update = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
    
End Function

Private Function IMortgageSubQuoteBO_CalcAndSaveMortgageCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSaveMortgageCostsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CalcAndSaveMortgageCosts"

    Dim blnMortgageCalcFeesOnly As Boolean
    Dim xmlOneOffCosts As IXMLDOMNode
    Dim xmlMortgageLender As IXMLDOMNode
    Dim blnCalcFeesOnly As Boolean
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.CalcAndSaveMortgageCosts(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    ' PSC 09/11/2006 EP2_41 - Start
    blnCalcFeesOnly = xmlGetNodeAsBoolean(vXMLRequest, cstrROOT_NODE_NAME & "/CALCULATEFEESONLY")
    
    If blnCalcFeesOnly Then
        Set xmlOneOffCosts = xmlGetMandatoryNode(xmlTempResponseNode, "ONEOFFCOSTLIST")
        Set xmlMortgageLender = xmlGetMandatoryNode(xmlTempResponseNode, "MORTGAGELENDER")
        xmlResponseElem.appendChild xmlOneOffCosts
        xmlResponseElem.appendChild xmlMortgageLender
    End If
    ' PSC 09/11/2006 EP2_41 - End
    
    Set IMortgageSubQuoteBO_CalcAndSaveMortgageCosts = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalcAndSaveMortgageCostsExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOneOffCosts = Nothing   ' PSC 09/11/2006 EP2_41
    Set xmlMortgageLender = Nothing   ' PSC 09/11/2006 EP2_41

    Exit Function

CalcAndSaveMortgageCostsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_CalcAndSaveMortgageCosts = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAndSaveMortgageCostsExit

End Function

Private Function IMortgageSubQuoteBO_CalcAndSaveRedemptionFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSaveRedemptionFeesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CalcAndSaveRedemptionFees"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.CalcAndSaveRedemptionFees(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IMortgageSubQuoteBO_CalcAndSaveRedemptionFees = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalcAndSaveRedemptionFeesExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IMortgageSubQuoteBO_CalcAndSaveRedemptionFees = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAndSaveRedemptionFeesExit

End Function

Private Function IMortgageSubQuoteBO_CalculateDrawDownCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
'This function takes in the MortgageSub Quote Data as a request and calculates the monthly costs for
'the loans components and the mortgage sub quote with a drawdown.

On Error GoTo CalculateDrawDownCostsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CalculateDrawDownCosts"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseData As IXMLDOMElement
    Dim xmlResponseComp As IXMLDOMNode
    Dim xmlUpdateRequest As IXMLDOMNode
    Dim xmlLoanComponentNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlOldNode As IXMLDOMNode
    Dim xmlRemoveNode As IXMLDOMNode
    Dim LoanComponentList As IXMLDOMNodeList
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlComponentElem As IXMLDOMElement
    Dim xmlLoanComponentElement As IXMLDOMElement
    Dim xmlDrawDownResponse As IXMLDOMNode
    Dim IntCompCount As Integer
    Dim dblInitialLAVal As Double
    Dim dblInitialTLCAVal As Double
    Dim dblLADrawDown As Double
    Dim dblTLCADrawDown As Double
    Dim strValue As String
    Dim dblDrawdown As Double
                   
    Set xmlElement = vXMLRequest.cloneNode(True)
    Set xmlResponseData = vXMLRequest.cloneNode(True)
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                   
    dblDrawdown = CSafeDbl(xmlElement.selectSingleNode("//DRAWDOWN").Text)
                   
    IntCompCount = 1
    Set LoanComponentList = xmlElement.selectNodes("//LOANCOMPONENT")
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
        
    For Each xmlComponentElem In xmlElement.selectNodes("//LOANCOMPONENT")
        If xmlComponentElem.selectSingleNode("LOANAMOUNT").Text = "0" Then
            xmlElement.selectSingleNode("//LOANCOMPONENTLIST").removeChild xmlComponentElem
        End If
    Next
    
    'Calculate monthly costs (less drawdown) and check response
    Set xmlDrawDownResponse = IMortgageSubQuoteBO_RunCompositeRateCalc(xmlElement)
    m_objErrAssist.CheckResponse xmlDrawDownResponse, xmlResponseElem, True
    
    'Update the MortgageSubQuote data with the new monthly cost less drawdown values
    Set xmlOldNode = m_objXmlAssist.GetNode(xmlResponseData, "MORTGAGESUBQUOTE/MONTHLYCOSTLESSDRAWDOWN", False)
    Set xmlNode = m_objXmlAssist.GetNode(xmlDrawDownResponse, "//TOTALNETMONTHLYCOST")
    xmlChangeNodeName xmlNode, "TOTALNETMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
    
    If Not xmlOldNode Is Nothing Then
        xmlResponseData.replaceChild xmlNode, xmlOldNode
    Else
        xmlChangeNodeName xmlNode, "TOTALMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
        xmlResponseData.appendChild xmlNode
    End If
    
    'Update each loan component element with monthly cost less drawdown amount

    Set xmlLoanComponentList = xmlResponseData.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
    
    For Each xmlLoanComponentElement In xmlLoanComponentList
        strValue = xmlLoanComponentElement.selectSingleNode("LOANCOMPONENTSEQUENCENUMBER").Text
        Set xmlResponseComp = m_objXmlAssist.GetNode(xmlDrawDownResponse, "//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER = " & strValue & "]", False)
        If Not xmlResponseComp Is Nothing Then
            'BMIDS736 GHun
            'Set xmlNode = m_objXmlAssist.GetNode(xmlResponseComp, "//OUT060_TOTALNETMORTGAGEPAYMENT1")
            'xmlChangeNodeName xmlNode, "OUT060_TOTALNETMORTGAGEPAYMENT1", "MONTHLYCOSTLESSDRAWDOWN"
            Set xmlNode = m_objXmlAssist.GetNode(xmlResponseComp, "//NETMONTHLYCOST")
            xmlChangeNodeName xmlNode, "NETMONTHLYCOST", "MONTHLYCOSTLESSDRAWDOWN"
            'BMIDS736 End
        
            If Not xmlLoanComponentElement.selectSingleNode("MONTHLYCOSTLESSDRAWDOWN") Is Nothing Then
                Set xmlRemoveNode = xmlLoanComponentElement.selectSingleNode("MONTHLYCOSTLESSDRAWDOWN")
                xmlLoanComponentElement.removeChild xmlRemoveNode
            End If
            xmlLoanComponentElement.appendChild xmlNode
        End If
    Next
            
    xmlResponseElem.appendChild xmlResponseData
    
    Set IMortgageSubQuoteBO_CalculateDrawDownCosts = xmlResponseElem
    
    m_objContext.SetComplete
    
CalculateDrawDownCostsExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlResponseComp = Nothing
    Set xmlUpdateRequest = Nothing
    Set xmlLoanComponentNode = Nothing
    Set xmlDrawDownResponse = Nothing
    Set LoanComponentList = Nothing
    Set xmlComponentElem = Nothing
    Set xmlNode = Nothing
    Set xmlOldNode = Nothing
    Set xmlRemoveNode = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlLoanComponentElement = Nothing
    
    Exit Function

CalculateDrawDownCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_CalculateDrawDownCosts = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CalculateDrawDownCostsExit
End Function

Public Function CalculateDrawDownCosts(ByVal vstrXMLRequest As String) As String
'This function takes in the MortgageSub Quote Data as a request and calculates the monthly vosts for
'the loans components and the mortgage sub quote with a drawdown.

On Error GoTo CalculateDrawDownCostsVbErr

    Const strFunctionName As String = "CalculateDrawDownCosts"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_CalculateDrawDownCosts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CalculateDrawDownCosts = xmlResponseElem.xml
    
    m_objContext.SetComplete

CalculateDrawDownCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

CalculateDrawDownCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CalculateDrawDownCosts = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CalculateDrawDownCostsExit

End Function

Private Function IMortgageSubQuoteBO_CalculateTotalIncentiveAmount(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Private Function IMortgageSubQuoteBO_CreateComponentsFromExistingAcc(ByVal vxmlTableElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------

End Function

Private Function IMortgageSubQuoteBO_CreateFirstSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CreateFirstSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objXmlAssist As New XMLAssist
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strLoansToBePorted As String

    strApplicationNo = objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONNUMBER")
    
    strApplicationFFNo = objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONFACTFINDNUMBER")
    
    
    If Len(strApplicationNo) = 0 Or _
       Len(strApplicationFFNo) = 0 Then
        Dim strDetails As String
        strDetails = " APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & " APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If



    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.CreateFirstSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_CreateFirstSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateFirstSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_CreateFirstSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstSubQuoteExit

End Function

Private Function IMortgageSubQuoteBO_CreateNewSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CreateNewSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objXmlAssist As New XMLAssist
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strMortgageSubQuoteNumber As String

    strApplicationNo = objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONNUMBER")
    
    strApplicationFFNo = objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONFACTFINDNUMBER")
    
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(vXMLRequest, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNo) = 0 Or _
       Len(strMortgageSubQuoteNumber) = 0 Or _
       Len(strApplicationFFNo) = 0 Then
        Dim strDetails As String
        strDetails = " APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & " APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
        strDetails = strDetails & " MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If



    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.CreateNewSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_CreateNewSubQuote = xmlResponseElem.cloneNode(True)
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_CreateNewSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewSubQuoteExit

End Function

Private Function IMortgageSubQuoteBO_CreateOneOffCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOneOffCostsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_CreateOneOffCosts"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    Dim objXmlAssist As New XMLAssist
    
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.CreateOneOffCosts(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_CreateOneOffCosts = xmlResponseElem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOneOffCostsExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IMortgageSubQuoteBO_CreateOneOffCosts = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOneOffCostsExit
End Function

Private Function IMortgageSubQuoteBO_DeleteLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLoanComponentVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_DeleteLoanComponent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
                                           
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.DeleteLoanComponent(vxmlTableElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_DeleteLoanComponent = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLoanComponentExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLoanComponentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_DeleteLoanComponent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLoanComponentExit

End Function

Private Function IMortgageSubQuoteBO_FindActiveLoanComponentsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindActiveLoanComponentsListVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_FindActiveLoanComponentsList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vxmlTableElement.nodeName = "FINDACTIVELOANCOMPONENTS" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("FINDACTIVELOANCOMPONENTS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FINDACTIVELOANCOMPONENTS" & " tag not found"
    End If
        
    'objXmlAssist.ChangeNodeName xmlRequestNode, "FINDACTIVELOANCOMPONENTS", "LOANCOMPONENT"
        
    Set xmlDataNode = objIMortgageSubquoteDO.FindActiveLoanComponentsList(xmlRequestNode)
    objXmlAssist.ChangeNodeName xmlDataNode, "FINDACTIVELOANCOMPONENTS", "LOANCOMPONENT"
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageSubQuoteBO_FindActiveLoanComponentsList = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
FindActiveLoanComponentsListExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindActiveLoanComponentsListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_FindActiveLoanComponentsList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindActiveLoanComponentsListExit

End Function


Private Function IMortgageSubQuoteBO_FindAvailableincentives(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
'This procedure takes in as a request data regarding the loan component and mortgage sub quote,
'checks for incentives still available for this product and returns them in the response.
'------------------------------------------------------------------------------------------
On Error GoTo FindAvailableIncentivesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_FindAvailableincentives"

    
    'declare variables
    Dim xmlExclusiveIncList As IXMLDOMNode
    Dim xmlInclusiveIncList As IXMLDOMNode
    Dim xmlExistingIncentive As IXMLDOMNode
    Dim xmlLoanCompNodeList As IXMLDOMNodeList
    Dim xmlIncentiveNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlLoanComp As IXMLDOMElement
    Dim xmlIncentive As IXMLDOMElement
    Dim xmlIncentive2 As IXMLDOMElement
    Dim xmlLoanCompIncGUID As IXMLDOMElement
    Dim xmlIncentiveGUID As IXMLDOMElement
    Dim xmlType As IXMLDOMElement
    Dim xmlExclusiveElem As IXMLDOMElement
    Dim xmlInclusiveElem As IXMLDOMElement
    Dim xmlIncHeaderElem As IXMLDOMElement
    Dim xmlInclHeaderElem As IXMLDOMElement
    Dim xmlExclHeaderElem As IXMLDOMElement
    Dim xmlBenefitType As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlList As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlLoanCompListNode As IXMLDOMNode
    Dim xmlActiveCompNode As IXMLDOMNode
    Dim xmlSeqNoNode As IXMLDOMNode
    Dim xmlGetIncentiveRequestNode As IXMLDOMNode
    Dim xmlGetIncentiveParentNode As IXMLDOMNode
    Dim xmlLoanAmountNode As IXMLDOMNode
    'DPF 14/11/2002 - BMIDS00805
    Dim xmlProductNode As IXMLDOMNode
    Dim xmlStartDateNode As IXMLDOMNode
    Dim xmlMortgageProductNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlLoanCompDoc As FreeThreadedDOMDocument40
    Dim xmlIncentivesDoc As FreeThreadedDOMDocument40
    Dim xmlComboRequest As FreeThreadedDOMDocument40
    Dim xmlComboResponse As FreeThreadedDOMDocument40
    Dim strLoanCompDetails As String
    Dim strGetIncentives As String
    Dim strComboList As String
    Dim lngProductLoanAmount As Long
    Dim intBenefitType As Integer
    Dim blnFinancial As Boolean
    Dim dblTotalIncentiveAmount As Double
    Dim dblIncentiveAmount As Double
    Dim objMortgageProductBO As MortgageProductBO
    Dim objConvertAssist As New ConvertAssist
    Dim objComboBO As New ComboBO
    
    'Set up DOM documents
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlLoanCompDoc = New FreeThreadedDOMDocument40
    xmlLoanCompDoc.validateOnParse = False
    xmlLoanCompDoc.setProperty "NewParser", True
    Set xmlIncentivesDoc = New FreeThreadedDOMDocument40
    xmlIncentivesDoc.validateOnParse = False
    xmlIncentivesDoc.setProperty "NewParser", True
    Set xmlComboRequest = New FreeThreadedDOMDocument40
    xmlComboRequest.validateOnParse = False
    xmlComboRequest.setProperty "NewParser", True
    Set xmlComboResponse = New FreeThreadedDOMDocument40
    xmlComboResponse.validateOnParse = False
    xmlComboResponse.setProperty "NewParser", True
    
    'Build response
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Set up connections to other components
    If m_objContext Is Nothing Then
        Set objMortgageProductBO = New MortgageProductBO
    Else
        Set objMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
    
    'Take in <MORTAGESUBQUOTE> section of request & Check we have data
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    'Get incentive benefit type combo data
    xmlComboRequest.async = False
    
    Set xmlList = xmlComboRequest.createElement("LIST")
    xmlComboRequest.appendChild xmlList
    Set xmlNode = xmlComboRequest.createElement("LISTNAME")
    xmlNode.Text = "IncentiveBenefitType"
    xmlList.appendChild xmlNode
    
    strComboList = objComboBO.GetComboList(xmlList.xml)
    Set xmlComboResponse = xmlLoad(strComboList, "FindAvailableIncentives")
    
    lngProductLoanAmount = CSafeLng(xmlRequestNode.selectSingleNode("//LOANAMOUNT").Text)
    dblTotalIncentiveAmount = 0
    dblIncentiveAmount = 0
    
    '------------------------------------------------------------------------------------------
    'STEP 1 :  Get Loan Composition Details and pull out component we are working with
    
    'Pull back loan composition details & feed them into a DOM document
    strLoanCompDetails = GetLoanCompositionDetails(vXMLRequest.xml)
    
    Set xmlLoanCompDoc = xmlLoad(strLoanCompDetails, "FindAvailableincentives")
    'BMIDS821 GHun Handle errors returned by GetLoanCompositionDetails
    errCheckXMLResponseNode xmlLoanCompDoc.documentElement, xmlResponseElem, True
    'BMIDS821 End
    
    Set xmlSeqNoNode = xmlGetNode(xmlRequestNode, "LOANCOMPONENTSEQUENCENUMBER")
    Set xmlProductNode = xmlGetNode(xmlRequestNode, "MORTGAGEPRODUCTCODE")
    Set xmlStartDateNode = xmlGetNode(xmlRequestNode, "STARTDATE")
    Set xmlLoanCompListNode = xmlLoanCompDoc.selectSingleNode("//LOANCOMPONENTLIST")
    
    'Retrieve the loan component we are dealing with (if it exists)
    If Not xmlSeqNoNode Is Nothing Then
        Set xmlActiveCompNode = xmlLoanCompListNode.selectSingleNode("LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER=" & xmlSeqNoNode.Text & "]")
        
        'DPF BMIDS00805 - check if the product hasn't been edited, otherwise we treat it as a new loan component
        If xmlActiveCompNode.selectSingleNode("MORTGAGEPRODUCTCODE").Text = xmlProductNode.Text _
        And xmlActiveCompNode.selectSingleNode("STARTDATE").Text = xmlStartDateNode.Text Then
            xmlLoanCompListNode.removeChild xmlActiveCompNode
        Else
            Set xmlActiveCompNode = xmlRequestNode.cloneNode(True)
            xmlChangeNodeName xmlActiveCompNode, "MORTAGESUBQUOTE", "LOANCOMPONENT"
        End If
    Else
        'Loan Component has not yet been saved to the sub quote so we'll grab details fed into procedure
        Set xmlActiveCompNode = xmlRequestNode.cloneNode(True)
        xmlChangeNodeName xmlActiveCompNode, "MORTAGESUBQUOTE", "LOANCOMPONENT"
    End If
    
    Set xmlMortgageProductNode = xmlActiveCompNode.selectSingleNode("MORTGAGEPRODUCTCODE")
    
    '------------------------------------------------------------------------------------------
    'STEP 2 :  Run through other Loan Components on application, creating a node list of components
    '          using the same mortgage product and adding up loan amounts and incentive amounts
        
    If Not xmlLoanCompListNode Is Nothing Then
    
        Set xmlLoanCompNodeList = xmlLoanCompListNode.selectNodes("LOANCOMPONENT[MORTGAGEPRODUCTCODE='" & xmlMortgageProductNode.Text & "']")
        
        For Each xmlElement In xmlLoanCompNodeList
            lngProductLoanAmount = lngProductLoanAmount + objConvertAssist.CSafeLng( _
                                        m_objXmlAssist.GetTagValue(xmlElement, "LOANAMOUNT"))
        Next
    End If
    
    '------------------------------------------------------------------------------------------
    'STEP 3 :  Build XML for GetIncentivesForProduct and make call to get Incentives back
    
    Set xmlGetIncentiveRequestNode = vXMLRequest.cloneNode(False)
    Set xmlGetIncentiveParentNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0).cloneNode(False)
    xmlChangeNodeName xmlGetIncentiveParentNode, "MORTGAGESUBQUOTE", "MORTGAGEPRODUCT"
    Set xmlLoanAmountNode = xmlActiveCompNode.selectSingleNode("LOANAMOUNT")
    xmlLoanAmountNode.Text = lngProductLoanAmount
    Set xmlStartDateNode = xmlActiveCompNode.selectSingleNode("STARTDATE")
        
    xmlGetIncentiveParentNode.appendChild xmlLoanAmountNode
    xmlGetIncentiveParentNode.appendChild xmlMortgageProductNode
    xmlGetIncentiveParentNode.appendChild xmlStartDateNode
    xmlGetIncentiveRequestNode.appendChild xmlGetIncentiveParentNode
    
    'Get incentives available for this product & feed it into a DOM for data manipulation
    strGetIncentives = objMortgageProductBO.GetIncentivesForProduct(xmlGetIncentiveRequestNode.xml)
    Set xmlIncentivesDoc = xmlLoad(strGetIncentives, "FindAvailableincentives")
    
    'If no incentives have been returned for either list we create a header node for that list
    If xmlIncentivesDoc.selectSingleNode("//INCLUSIVEINCENTIVELIST") Is Nothing Then
        Set xmlInclusiveElem = xmlIncentivesDoc.createElement("INCLUSIVEINCENTIVELIST")
        xmlIncentivesDoc.documentElement.appendChild xmlInclusiveElem
    End If
    
    If xmlIncentivesDoc.selectSingleNode("//EXCLUSIVEINCENTIVELIST") Is Nothing Then
        Set xmlExclusiveElem = xmlIncentivesDoc.createElement("EXCLUSIVEINCENTIVELIST")
        xmlIncentivesDoc.documentElement.appendChild xmlExclusiveElem
    End If
   
    Set xmlInclusiveIncList = xmlIncentivesDoc.selectSingleNode("//INCLUSIVEINCENTIVELIST")
    Set xmlExclusiveIncList = xmlIncentivesDoc.selectSingleNode("//EXCLUSIVEINCENTIVELIST")
    
    '------------------------------------------------------------------------------------------
    'STEP 4 :  Run through incentives already used on this application (for matching product)
    '          and calculate what incentives can still be offered to customer
    
    'check we have some incentives, if so check we're not already upto our limit.
    
    If Not xmlLoanCompListNode Is Nothing Then
        If xmlLoanCompNodeList.length <> 0 Then
            For Each xmlLoanComp In xmlLoanCompNodeList
                For Each xmlIncentive In xmlLoanComp.selectNodes("MORTGAGEINCENTIVELIST/MORTGAGEINCENTIVE")
                    Set xmlLoanCompIncGUID = xmlIncentive.selectSingleNode("INCENTIVEGUID")
                    Set xmlType = xmlIncentive.selectSingleNode("TYPE")
                    intBenefitType = xmlIncentive.selectSingleNode("INCENTIVEBENEFITTYPE").Text
                    
                    'check it's financial
                    Set xmlBenefitType = xmlComboResponse.selectSingleNode("//LIST/LISTNAME/LISTENTRY[VALUEID = " & intBenefitType & "]")
                    If xmlBenefitType.selectSingleNode("VALIDATIONTYPELIST/VALIDATIONTYPE").Text = "F" Then
                        blnFinancial = True
                    Else
                        blnFinancial = False
                    End If
                    
                    If xmlType.Text = "2" Then
                    'Exclusive Incentive
                        If xmlExclusiveIncList.hasChildNodes <> False Then
                            For Each xmlExclusiveElem In xmlExclusiveIncList.selectNodes("INCENTIVE")
                                If xmlExclusiveElem.selectSingleNode("INCENTIVEGUID").Text = xmlLoanCompIncGUID.Text Then
                                                                        
                                    If blnFinancial = True Then
                                        'If the incentive has a fixed amount and has already been applied remove it
                                        If xmlExclusiveElem.selectSingleNode("INCENTIVEFIXEDINDICATOR").Text = "1" Then
                                            xmlExclusiveIncList.removeChild xmlExclusiveElem
                                        End If
                                        
                                        'if the incentive is a percentage calculate remaining amount
                                        'DPF 04/11/02 - removed AND condition to do with Percentage Maximum amount
                                        If xmlExclusiveElem.selectSingleNode("INCENTIVEFIXEDINDICATOR").Text = "0" Then
                                            dblTotalIncentiveAmount = xmlExclusiveElem.selectSingleNode("AMOUNT").Text
                                            dblIncentiveAmount = xmlIncentive.selectSingleNode("INCENTIVEAMOUNT").Text
                                            dblIncentiveAmount = dblTotalIncentiveAmount - dblIncentiveAmount
                                                                                
                                            If dblIncentiveAmount < 0 Then
                                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                                oeInvalidIncentiveAmount, "Maximum Incentive Amount already exceeded"
                                            ElseIf dblIncentiveAmount = 0 Then
                                            'Remove Incentive from response
                                                xmlExclusiveIncList.removeChild xmlExclusiveElem
                                            Else
                                            'Amend Incentive amount
                                                xmlExclusiveElem.selectSingleNode("AMOUNT").Text = dblIncentiveAmount
                                            End If
                                        End If
                                    Else
                                        xmlExclusiveIncList.removeChild xmlExclusiveElem
                                    End If
                                Else
                                    xmlExclusiveIncList.removeChild xmlExclusiveElem
                                End If
                            Next
                        End If
                    Else
                    'Inclusive Incentive
                        If xmlInclusiveIncList.hasChildNodes <> False Then
                            For Each xmlInclusiveElem In xmlInclusiveIncList.selectNodes("INCENTIVE")
                                If xmlInclusiveElem.selectSingleNode("INCENTIVEGUID").Text = xmlLoanCompIncGUID.Text Then
                                    If blnFinancial = True Then
                                        'If the incentive has a fixed amount and has already been applied remove it
                                        If xmlInclusiveElem.selectSingleNode("INCENTIVEFIXEDINDICATOR").Text = "1" Then
                                            xmlInclusiveIncList.removeChild xmlInclusiveElem
                                        End If
                                        
                                        'if the incentive is a percentage calculate remaining amount
                                        'DPF 04/11/02 - removed AND condition to do with Percentage Maximum amount
                                        If xmlInclusiveElem.selectSingleNode("INCENTIVEFIXEDINDICATOR").Text = "0" Then
                                            dblTotalIncentiveAmount = xmlInclusiveElem.selectSingleNode("AMOUNT").Text
                                            dblIncentiveAmount = xmlIncentive.selectSingleNode("INCENTIVEAMOUNT").Text
                                            dblIncentiveAmount = dblTotalIncentiveAmount - dblIncentiveAmount
                                                                                
                                            If dblIncentiveAmount < 0 Then
                                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                                oeInvalidIncentiveAmount, "Maximum Incentive Amount already exceeded"
                                            ElseIf dblIncentiveAmount = 0 Then
                                            'Remove Incentive from response
                                                xmlInclusiveIncList.removeChild xmlInclusiveElem
                                            Else
                                            'Amend Incentive amount
                                                xmlInclusiveElem.selectSingleNode("AMOUNT").Text = dblIncentiveAmount
                                            End If
                                        End If
                                    Else
                                        xmlInclusiveIncList.removeChild xmlInclusiveElem
                                    End If
                                End If
                            Next
                        End If
                    End If
                    
                Next
            Next
        Else
                 
        End If
    End If
           
    'loop through incentives and check if we already have a match for the current loan component
    'if so overwite the incentive amount.
    For Each xmlExclusiveElem In xmlExclusiveIncList.selectNodes("INCENTIVE")
        If Not xmlActiveCompNode.selectSingleNode("MORTGAGEINCENTIVELIST/MORTGAGEINCENTIVE[INCENTIVEGUID = '" & xmlExclusiveElem.selectSingleNode("INCENTIVEGUID").Text & "']") Is Nothing Then
            Set xmlExistingIncentive = xmlActiveCompNode.selectSingleNode("MORTGAGEINCENTIVELIST/MORTGAGEINCENTIVE[INCENTIVEGUID = '" & xmlExclusiveElem.selectSingleNode("INCENTIVEGUID").Text & "']")
            dblIncentiveAmount = xmlExistingIncentive.selectSingleNode("INCENTIVEAMOUNT").Text
            xmlExclusiveElem.selectSingleNode("AMOUNT").Text = dblIncentiveAmount
        End If
    Next

    For Each xmlInclusiveElem In xmlInclusiveIncList.selectNodes("INCENTIVE")
        If Not xmlActiveCompNode.selectSingleNode("MORTGAGEINCENTIVELIST/MORTGAGEINCENTIVE[INCENTIVEGUID = '" & xmlInclusiveElem.selectSingleNode("INCENTIVEGUID").Text & "']") Is Nothing Then
            Set xmlExistingIncentive = xmlActiveCompNode.selectSingleNode("MORTGAGEINCENTIVELIST/MORTGAGEINCENTIVE[INCENTIVEGUID = '" & xmlInclusiveElem.selectSingleNode("INCENTIVEGUID").Text & "']")
            dblIncentiveAmount = xmlExistingIncentive.selectSingleNode("INCENTIVEAMOUNT").Text
            xmlInclusiveElem.selectSingleNode("AMOUNT").Text = dblIncentiveAmount
        End If
    Next
           
    'append incentive lists to the response
    xmlResponseElem.appendChild xmlExclusiveIncList
    xmlResponseElem.appendChild xmlInclusiveIncList

    
    m_objContext.SetComplete
    
    Set IMortgageSubQuoteBO_FindAvailableincentives = xmlResponseElem
    

FindAvailableIncentivesExit:

    Set xmlExclusiveIncList = Nothing
    Set xmlInclusiveIncList = Nothing
    Set xmlLoanCompNodeList = Nothing
    Set xmlIncentiveNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlLoanComp = Nothing
    Set xmlIncentive = Nothing
    Set xmlIncentive2 = Nothing
    Set xmlLoanCompIncGUID = Nothing
    Set xmlIncentiveGUID = Nothing
    Set xmlType = Nothing
    Set xmlExclusiveElem = Nothing
    Set xmlInclusiveElem = Nothing
    Set xmlIncHeaderElem = Nothing
    Set xmlInclHeaderElem = Nothing
    Set xmlExclHeaderElem = Nothing
    Set xmlBenefitType = Nothing
    Set xmlNode = Nothing
    Set xmlList = Nothing
    Set xmlRequestNode = Nothing
    Set xmlLoanCompListNode = Nothing
    Set xmlActiveCompNode = Nothing
    Set xmlSeqNoNode = Nothing
    Set xmlGetIncentiveRequestNode = Nothing
    Set xmlGetIncentiveParentNode = Nothing
    Set xmlLoanAmountNode = Nothing
    Set xmlMortgageProductNode = Nothing
    Set xmlStartDateNode = Nothing
    Set xmlProductNode = Nothing
    Set xmlExistingIncentive = Nothing
    Set xmlOut = Nothing
    Set xmlLoanCompDoc = Nothing
    Set xmlIncentivesDoc = Nothing
    Set xmlComboRequest = Nothing
    Set xmlComboResponse = Nothing
    Set objMortgageProductBO = Nothing
    Set objConvertAssist = Nothing
    Set objComboBO = Nothing
        

    Exit Function

FindAvailableIncentivesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_FindAvailableincentives = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindAvailableIncentivesExit

End Function

Private Function IMortgageSubQuoteBO_FindMortgageIncentiveList( _
                         ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageIncentiveListVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_FindMortgageIncentiveList"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
                
    If vXMLRequest.nodeName = cstrMORTGAGEINCENTIVE_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrMORTGAGEINCENTIVE_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrMORTGAGEINCENTIVE_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageSubquoteDO.FindMortgageIncentiveList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set IMortgageSubQuoteBO_FindMortgageIncentiveList = xmlResponseElem

FindMortgageIncentiveListExit:

    Set xmlRequestNode = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindMortgageIncentiveListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_FindMortgageIncentiveList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindMortgageIncentiveListExit

End Function

Private Function IMortgageSubQuoteBO_GetAPROneOffCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets all mortgageoneoffcosts to be included in APR calcs.
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAPROneOffCostsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetAPROneOffCosts"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIMortgageSubQuoteTxBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.GetAPROneOffCosts(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    Set IMortgageSubQuoteBO_GetAPROneOffCosts = xmlResponseElem.cloneNode(True)

    m_objContext.SetComplete
    
GetAPROneOffCostsExit:

    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

GetAPROneOffCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_GetAPROneOffCosts = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume GetAPROneOffCostsExit

End Function

Private Function IMortgageSubQuoteBO_GetData( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageSubquoteDO.GetData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageSubQuoteBO_GetData = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetDataExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_GetData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function IMortgageSubQuoteBO_GetDefaultsForNewLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDefaultsForNewLoanComponentVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetDefaultsForNewLoanComponent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim ERRNUMBER As Long
    Dim blnRaiseError As Boolean
        
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                    
    If vxmlTableElement.nodeName = "QUOTATION" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("QUOTATION").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "QUOTATION" & " tag not found"
    End If
        
    Dim objSQLAssist As New SQLAssist
    
    Dim strAmountRequested As String
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strAmountRequested = objXmlAssist.GetTagValue(xmlRequestNode, "AMOUNTREQUESTED")
    strApplicationNumber = objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(xmlRequestNode, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Or _
       Len(strMortgageSubQuoteNumber) = 0 Or Len(strAmountRequested) = 0 Then
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        strDetails = strDetails & "; AMOUNTREQUESTED = " & strAmountRequested
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If

    Dim objConvertAssist As New ConvertAssist
    
    Dim lngInAmountRequested As Long
    
    lngInAmountRequested = objConvertAssist.CSafeLng(strAmountRequested)
            
    Dim lngTotalLoanAmount As Long
    Dim blnMortCalcDefault As Boolean
    
    blnMortCalcDefault = True
    
    lngTotalLoanAmount = 0
    
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    
    blnRaiseError = False
    
    objXmlAssist.ChangeNodeName xmlRequestNode, "QUOTATION", "LOANCOMPONENT"
    
    Set xmlTempElem = objIMortgageSubquoteDO.FindLoanComponentList(xmlRequestNode)
    blnRaiseError = True
    'record not found is not an error here
    If Not xmlTempElem Is Nothing Then
        Set xmlLoanComponentList = xmlTempElem.selectNodes(".//LOANCOMPONENT")
               
        For Each xmlElement In xmlLoanComponentList
            
            lngTotalLoanAmount = lngTotalLoanAmount + objConvertAssist.CSafeLng( _
                                    objXmlAssist.GetTagValue(xmlElement, "LOANAMOUNT"))
            
            'Check if this value has already been set
            If blnMortCalcDefault Then
                'check the portedloan value
                If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlElement, "PORTEDLOAN")) = False Then
            
                    blnMortCalcDefault = False
                
                End If
            
            End If
        
        Next
    End If
    
    Dim xmlDefaultElem As IXMLDOMElement
    Set xmlDefaultElem = xmlOut.createElement("MORTGAGECALCULATORDEFAULT")

    xmlResponseElem.appendChild xmlDefaultElem
       
    Dim strTagName As String
    strTagName = "DEFAULTLOANAMOUNT"
    
    Set xmlTempElem = xmlOut.createElement(strTagName)
    
    xmlTempElem.Text = lngInAmountRequested - lngTotalLoanAmount

    xmlDefaultElem.appendChild xmlTempElem
    
    If blnMortCalcDefault Then
    
        Dim objIomCMClassDef As IomCMClassDef
        Set objIomCMClassDef = New omCMClassDef
    
        Dim xmlMortgageCalculatorElem As IXMLDOMElement
        blnRaiseError = False
        'record not found is NOT!!! an error here
        objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENT", "MORTGAGECALCULATORDETAILS"
        Set xmlMortgageCalculatorElem = objIMortgageSubquoteDO.GetMortgageCalculatorDetails(vxmlTableElement)
                                    
        strTagName = "DEFAULTMORTGAGEPRODUCTCODE"
        Set xmlTempElem = xmlOut.createElement(strTagName)
        
        If Not xmlMortgageCalculatorElem Is Nothing Then
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlMortgageCalculatorElem, "MORTGAGEPRODUCTCODE")
        End If
        
        xmlDefaultElem.appendChild xmlTempElem
    
        strTagName = "DEFAULTSTARTDATE"
        Set xmlTempElem = xmlOut.createElement(strTagName)
        
        If Not xmlMortgageCalculatorElem Is Nothing Then
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlMortgageCalculatorElem, "STARTDATE")
        End If
        
        xmlDefaultElem.appendChild xmlTempElem
        
        strTagName = "DEFAULTTERMYEARS"
        Set xmlTempElem = xmlOut.createElement(strTagName)
        
        If Not xmlMortgageCalculatorElem Is Nothing Then
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlMortgageCalculatorElem, "TERMYEARS")
        End If
        
        xmlDefaultElem.appendChild xmlTempElem
        
        strTagName = "DEFAULTTERMMONTHS"
        Set xmlTempElem = xmlOut.createElement(strTagName)
        
        If Not xmlMortgageCalculatorElem Is Nothing Then
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlMortgageCalculatorElem, "TERMMONTHS")
        End If
    
        xmlDefaultElem.appendChild xmlTempElem
        
    End If
                
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IMortgageSubQuoteBO_GetDefaultsForNewLoanComponent = xmlResponseElem.cloneNode(True)

GetDefaultsForNewLoanComponentExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objConvertAssist = Nothing
    Set objSQLAssist = Nothing
    Set objIomCMClassDef = Nothing

    Exit Function

GetDefaultsForNewLoanComponentVbErr:

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

    Set IMortgageSubQuoteBO_GetDefaultsForNewLoanComponent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDefaultsForNewLoanComponentExit


End Function

Private Function IMortgageSubQuoteBO_GetLoanComponentData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanComponentDataVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetLoanComponentData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vxmlTableElement.nodeName = "MORTGAGESUBQUOTE" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("MORTGAGESUBQUOTE").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGESUBQUOTE" & " tag not found"
    End If
        
    objXmlAssist.ChangeNodeName xmlRequestNode, "MORTGAGESUBQUOTE", "LOANCOMPONENT"
        
    Set xmlDataNode = objIMortgageSubquoteDO.GetLoanComponentData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageSubQuoteBO_GetLoanComponentData = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetLoanComponentDataExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetLoanComponentDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_GetLoanComponentData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoanComponentDataExit

End Function

Private Function IMortgageSubQuoteBO_GetLoanCompositionDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetLoanCompositionDetailsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetLoanCompositionDetails"

    Dim xmlRequestNode                      As IXMLDOMNode
    Dim xmlOut                              As FreeThreadedDOMDocument40
    Dim xmlMPSearch                         As FreeThreadedDOMDocument40
    Dim xmlResponseElem                     As IXMLDOMElement
    Dim objSQLAssist                        As SQLAssist
    Dim objComboDO                          As ComboDO
    Dim blnRaiseError                       As Boolean
    Dim objIMortgageSubquoteDO              As IMortgageSubQuoteDO
    Dim objMortgageProductBO                As MortgageProductBO
    Dim xmlTempElem                         As IXMLDOMElement
    Dim strApplicationNumber                As String
    Dim strApplicationFactFindNumber        As String
    '*=[MC]BMIDS763 - APPLICATION DATE
    Dim xmlEleApplicationDate               As IXMLDOMElement
    Dim strApplicationDate                  As String
    '*=SECTION END
    Dim strMortgageSubQuoteNumber           As String
    Dim strLocation                         As String
    Dim strValuationType                    As String
    Dim strPurchasePriceOrEstimatedValue    As String
    Dim strTypeOfApplication                As String
    Dim strDetails                          As String
    Dim xmlProductSearchNode                As IXMLDOMElement
    Dim xmlMortgageSubQuoteData             As IXMLDOMElement
    Dim xmlLoanComponentList                As IXMLDOMNodeList
    Dim xmlLoanComponentNode                As IXMLDOMElement
    Dim xmlMortgageIncentiveList            As IXMLDOMElement
    Dim strTemp                             As String
    Dim strTagName                          As String
    Dim blnValue                            As Boolean
    Dim strLegalFeeType                     As String
    Dim xmlElement                          As IXMLDOMElement
    Dim strText                             As String
    Dim objConvertAssist                    As ConvertAssist
    Dim strMortgageProduct                  As String
    Dim xmlMortgageProductDoc               As FreeThreadedDOMDocument40
    'MAR46 GHun
    Dim blnIsTOE                            As Boolean
    Dim blnIsPorted                         As Boolean
    Dim blnIsUpdateRequired                 As Boolean
    Dim xmlLCUpdateRequest                  As IXMLDOMNode
    Dim xmlLCNode                           As IXMLDOMNode
    Dim xmlIntRate                          As IXMLDOMNode
    Dim strValue                            As String
    'MAR46 End
    'SR 27/02/2007 : EP2_1272
    Dim xmlTempRequest                      As IXMLDOMNode
    Dim objCrudBO                           As omCRUDBO
    Dim xmlIntroducerFeeResp                As IXMLDOMNode
    'SR 27/02/2007 : EP2_1272 - End
    
    
    blnRaiseError = True
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                  
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objMortgageProductBO = New MortgageProductBO
        Set objComboDO = New ComboDO
        Set objCrudBO = New omCRUDBO
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objCrudBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDBO")
    End If
                
    If vxmlTableElement.nodeName = "MORTGAGESUBQUOTE" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("MORTGAGESUBQUOTE").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGESUBQUOTE" & " tag not found"
    End If
        
    strApplicationNumber = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "APPLICATIONFACTFINDNUMBER")
    '*=[MC]BMIDS763
    strApplicationDate = m_objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONDATE")
    'END SECTION
    
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "MORTGAGESUBQUOTENUMBER")
    strLocation = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "LOCATION")
    strValuationType = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "TYPEOFVALUATION")
    'MAR1061 get purchaseprice from MortgageSubquote
    'strPurchasePriceOrEstimatedValue = m_objXmlAssist.GetTagValue(xmlRequestNode, _
    '                                                "PURCHASEPRICE")
    strTypeOfApplication = m_objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "TYPEOFAPPLICATION")
                                                     
    'Len(strPurchasePriceOrEstimatedValue) = 0 Or MAR1061 removed from if below
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strLocation) = 0 Or _
        Len(strValuationType) = 0 Or _
        Len(strTypeOfApplication) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
        
        Set objSQLAssist = New SQLAssist
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; LOCATION = " & strLocation
        strDetails = strDetails & "; TYPEOFVALUATION = " & strValuationType
        'MAR1061 strDetails = strDetails & "; PURCHASEPRICE = " & strPurchasePriceOrEstimatedValue
        strDetails = strDetails & "; TYPEOFAPPLICATION = " & strTypeOfApplication
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
    
    'XML structure to be used in Product search later

    Set xmlMPSearch = New FreeThreadedDOMDocument40
    xmlMPSearch.validateOnParse = False
    xmlMPSearch.setProperty "NewParser", True
    Set xmlTempElem = xmlMPSearch.createElement("MORTGAGEPRODUCT")
    Set xmlProductSearchNode = xmlMPSearch.appendChild(xmlTempElem)
    
    ' SR 19/12/2002 : BM0209 - get data required for CostModelling only
    Set xmlTempElem = xmlMPSearch.createElement("CALCSDATAONLY")
    xmlTempElem.Text = "1"
    xmlProductSearchNode.appendChild xmlTempElem
    ' SR 19/12/2002 : BM0209 - End
    '*=[MC]BMIDS763
    Set xmlTempElem = xmlMPSearch.createElement("APPLICATIONNUMBER")
    xmlProductSearchNode.appendChild xmlTempElem
    
    Set xmlTempElem = xmlMPSearch.createElement("APPLICATIONFACTFINDNUMBER")
    xmlProductSearchNode.appendChild xmlTempElem
    
    Set xmlTempElem = xmlMPSearch.createElement("APPLICATIONDATE")
    xmlProductSearchNode.appendChild xmlTempElem
    
    '*=END SECTION
    
    Set xmlTempElem = xmlMPSearch.createElement("MORTGAGEPRODUCTCODE")
    xmlProductSearchNode.appendChild xmlTempElem
    
    Set xmlTempElem = xmlMPSearch.createElement("STARTDATE")
    xmlProductSearchNode.appendChild xmlTempElem
        
    'MAR1061 get purchaseprice from MortgageSubquote
    'Set xmlTempElem = xmlMPSearch.createElement("PURCHASEPRICE")
    'xmlProductSearchNode.appendChild xmlTempElem
    'xmlTempElem.Text = strPurchasePriceOrEstimatedValue ' passed in
            
    Set xmlTempElem = xmlMPSearch.createElement("TYPEOFAPPLICATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strTypeOfApplication 'passed in
    
    Set xmlTempElem = xmlMPSearch.createElement("LOCATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strLocation 'passed in
    
    Set xmlTempElem = xmlMPSearch.createElement("TYPEOFVALUATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strValuationType ' passed in
    
    Set xmlTempElem = xmlMPSearch.createElement("FEETYPE")
    xmlProductSearchNode.appendChild xmlTempElem
    
    'Get the mortgagesubquote record
    Set xmlMortgageSubQuoteData = objIMortgageSubquoteDO.GetData(xmlRequestNode)
        
    'append returned "AMOUNTREQUESTED" value to use in MP search later
    
    strTagName = "AMOUNTREQUESTED"
    Set xmlTempElem = xmlMortgageSubQuoteData.getElementsByTagName(strTagName).Item(0)
    strTemp = xmlTempElem.Text
   
    Set xmlTempElem = xmlMPSearch.createElement(strTagName)
    xmlTempElem.Text = strTemp
    xmlProductSearchNode.appendChild xmlTempElem

    'append returned "LTV" value to use in MP search later
    
    strTagName = "LTV"
    Set xmlTempElem = xmlMortgageSubQuoteData.getElementsByTagName(strTagName).Item(0)
    strTemp = xmlTempElem.Text
   
    Set xmlTempElem = xmlMPSearch.createElement(strTagName)
    xmlTempElem.Text = strTemp
    xmlProductSearchNode.appendChild xmlTempElem
    
    'MAR1061 append purchasepriceorestimatedvalue
    strTagName = "PURCHASEPRICEORESTIMATEDVALUE"
    Set xmlTempElem = xmlMortgageSubQuoteData.getElementsByTagName(strTagName).Item(0)
    strTemp = xmlTempElem.Text
    
    Set xmlTempElem = xmlMPSearch.createElement("PURCHASEPRICE")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strTemp
    
    'and append the mortgagesubquotedata to the return node
    Set xmlMortgageSubQuoteData = xmlResponseElem.appendChild(xmlMortgageSubQuoteData)
    
    'Now find the list of loan components for this mortgagesubquote
    
    m_objXmlAssist.ChangeNodeName xmlRequestNode, "MORTGAGESUBQUOTE", "LOANCOMPONENT", False    'BMIDS782 GHun
    
    blnRaiseError = False 'dont raise record not found as an error
    Set xmlLoanComponentNode = objIMortgageSubquoteDO.FindLoanComponentList(xmlRequestNode)
    blnRaiseError = True
    If Not xmlLoanComponentNode Is Nothing Then
        'append the list to the mortgagesubquote
        xmlMortgageSubQuoteData.appendChild xmlLoanComponentNode
        Set xmlLoanComponentList = xmlLoanComponentNode.selectNodes(".//LOANCOMPONENT")
    End If
    
        
    'Do the legal fee stuff
    blnValue = True
    
    If Not xmlLoanComponentList Is Nothing Then
        For Each xmlElement In xmlLoanComponentList
            
            strText = m_objXmlAssist.GetTagValue(xmlElement, "REPAYMENTMETHOD")
    
            If Len(strText) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, "Missing RepaymentMethod Value"
            End If
                    
            blnValue = objComboDO.IsItemInValidation("RepaymentType", strText, "C")
            
            If Not blnValue Then 'false
                strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "D")
                Exit For
            End If
            'get the mortgage product if applicable
        Next
    End If
    
    If Len(strLegalFeeType) = 0 Then
        strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "R")
    End If
    
    Set xmlTempElem = xmlProductSearchNode.getElementsByTagName("FEETYPE").Item(0)
    xmlTempElem.Text = strLegalFeeType
    
' MS    21/07/00    performance rework.
    Set objConvertAssist = New ConvertAssist
    
    'AS 09/07/01 SYS2463 XML Parser 3 work around
    'Need to reset xmlLoanComponentList back to beginning.
    If Not xmlLoanComponentNode Is Nothing Then
        Set xmlLoanComponentList = xmlLoanComponentNode.selectNodes(".//LOANCOMPONENT")
    End If

    'BMIDS815 GHun Application nodes only need to be set once outside of the loop
    '*=[MC]BMIDS763
    Dim xmlTempDOMDocument As FreeThreadedDOMDocument40
    Set xmlTempDOMDocument = New FreeThreadedDOMDocument40
    xmlTempDOMDocument.validateOnParse = False
    xmlTempDOMDocument.setProperty "NewParser", True
    
    Set xmlEleApplicationDate = xmlTempDOMDocument.createElement("APPLICATIONDATE")
    xmlEleApplicationDate.Text = strApplicationDate
    Call xmlProductSearchNode.appendChild(xmlEleApplicationDate)
    
    xmlProductSearchNode.selectSingleNode("APPLICATIONNUMBER").Text = strApplicationNumber
    xmlProductSearchNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text = strApplicationFactFindNumber
    xmlProductSearchNode.selectSingleNode("APPLICATIONDATE").Text = strApplicationDate
    '*=END SECTION
    'BMIDS815 End

    'MAR46 GHun Check if the application is a transfer of equity
    blnIsTOE = IsValidationType("TypeOfMortgage", strTypeOfApplication, "TOE")  'MAR355 GHun
    'MAR46 End

    If Not xmlLoanComponentList Is Nothing Then
        For Each xmlElement In xmlLoanComponentList
                   
            If Not objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue(xmlElement, "PORTEDLOAN")) Then
                    
                'get the mortgageproductcode for this loancomponent
                strTagName = "MORTGAGEPRODUCTCODE"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                strTemp = xmlTempElem.Text
                
                Set xmlTempElem = xmlProductSearchNode.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = strTemp
                
                'get the startdate for this loancomponent
                strTagName = "STARTDATE"
                Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
                strTemp = xmlTempElem.Text
                
                Set xmlTempElem = xmlProductSearchNode.getElementsByTagName(strTagName).Item(0)
                xmlTempElem.Text = strTemp
                                        
                'BM0211 MDC 31/12/2002
                xmlCopyAttribute vxmlTableElement, xmlProductSearchNode, "RATECHANGE"
                'BM0211 MDC 31/12/2002 - End
                                        
                'BM0230 MDC 08/01/2003
                xmlCopyAttribute vxmlTableElement, xmlProductSearchNode, "BASERATEAPPLIED"
                'BM0230 MDC 08/01/2003 - End
                
                'BMIDS815 GHun Only create the porting nodes if necessary
                Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PORTINGINDICATOR")
                If xmlTempElem Is Nothing Then
                    Set xmlTempElem = xmlOut.createElement("PORTINGINDICATOR")
                    xmlProductSearchNode.appendChild xmlTempElem
                End If
                'MAR46 GHun
                blnIsPorted = (blnIsTOE Or xmlGetNodeText(xmlElement, "MANUALPORTEDLOANIND") = "1" Or xmlGetNodeText(xmlElement, "PRODUCTSWITCHRETAINPRODUCTIND") = "1")
      
                If blnIsPorted Then
                    xmlTempElem.Text = "1"   'PortingIndicator = True
                'MAR46 End
                    Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PRODUCTSTEP")
                    If xmlTempElem Is Nothing Then
                        Set xmlTempElem = xmlOut.createElement("PRODUCTSTEP")
                        xmlProductSearchNode.appendChild xmlTempElem
                    End If
                    xmlTempElem.Text = xmlGetNodeText(xmlElement, "PRODUCTSTEP")
                    
                    Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "REMAININGSTEPDURATION")
                    If xmlTempElem Is Nothing Then
                        Set xmlTempElem = xmlOut.createElement("REMAININGSTEPDURATION")
                        xmlProductSearchNode.appendChild xmlTempElem
                    End If
                    xmlTempElem.Text = xmlGetNodeText(xmlElement, "REMAININGSTEPDURATION")
                    
                    Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PORTEDPRODUCTSTARTDATE")
                    If xmlTempElem Is Nothing Then
                        Set xmlTempElem = xmlOut.createElement("PORTEDPRODUCTSTARTDATE")
                        xmlProductSearchNode.appendChild xmlTempElem
                    End If
                    xmlTempElem.Text = xmlGetNodeText(xmlElement, "PORTEDPRODUCTSTARTDATE")
                'BMIDS815 End
                'MAR46 GHun
                Else
                    xmlTempElem.Text = "0"  'PortingIndicator = False
                'MAR46 End
                End If
                
                ' PSC 19/02/2007 EP2_1488 - Start
                ' PSC 30/03/2007 EP2_2149 - Start
                Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "ORIGINALLTV")
                
                If xmlTempElem Is Nothing Then
                    Set xmlTempElem = xmlOut.createElement("ORIGINALLTV")
                    xmlProductSearchNode.appendChild xmlTempElem
                End If

                xmlTempElem.Text = xmlGetNodeText(xmlElement, "ORIGINALLTV")
                ' PSC 30/03/2007 EP2_2149 - Start
                ' PSC 19/02/2007 EP2_1488 - End
                
                strMortgageProduct = objMortgageProductBO.GetFullProductDetails(xmlProductSearchNode.xml)
            
            Else
            
                'add a dummy mortgageproduct record
                'DM 17/08/00 AQR0924 Added the RESPONSE tag around the xml string below
                strMortgageProduct = "<RESPONSE TYPE=""SUCCESS""><MORTGAGEPRODUCT><MORTGAGEPRODUCTCODE>999999</MORTGAGEPRODUCTCODE></MORTGAGEPRODUCT></RESPONSE>"

                'What Format is needed for this MP ?
            End If
            ' MS    21/07/00    performance rework.
            Set xmlMortgageProductDoc = m_objXmlAssist.load(strMortgageProduct, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlMortgageProductDoc.documentElement, xmlElement, True
            m_objXmlAssist.AttachResponseData xmlElement, xmlMortgageProductDoc.documentElement
                                                                                                                     
            'MAR46 GHun Update PRODUCTSTEP and REMAININGSTEPDURATION on LoanComponent if required
            If blnIsPorted Then
                blnIsUpdateRequired = False
                Set xmlLCUpdateRequest = vxmlTableElement.cloneNode(False)
                Set xmlLCNode = xmlOut.createElement("LOANCOMPONENT")
                xmlLCUpdateRequest.appendChild xmlLCNode
                
                Set xmlIntRate = m_objXmlAssist.GetMandatoryNode(xmlMortgageProductDoc, "/RESPONSE/MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[1]")
                
                If Not xmlIntRate Is Nothing Then
                    strValue = m_objXmlAssist.GetNodeValue(xmlIntRate, "INTERESTRATETYPESEQUENCENUMBER", True, True)
                    If Len(strValue) > 0 Then
                        blnIsUpdateRequired = True
                        Set xmlTempElem = xmlOut.createElement("PRODUCTSTEP")
                        xmlTempElem.Text = strValue
                        xmlLCNode.appendChild xmlTempElem
                    End If
                    
                    strValue = m_objXmlAssist.GetNodeValue(xmlIntRate, "INTERESTRATEPERIOD", True, True)
                    If Len(strValue) > 0 Then
                        blnIsUpdateRequired = True
                        Set xmlTempElem = xmlOut.createElement("REMAININGSTEPDURATION")
                        xmlTempElem.Text = strValue
                        xmlLCNode.appendChild xmlTempElem
                    End If
                End If
                
                If blnIsUpdateRequired Then
                    xmlLCNode.appendChild xmlElement.selectSingleNode("APPLICATIONNUMBER").cloneNode(True)
                    xmlLCNode.appendChild xmlElement.selectSingleNode("APPLICATIONFACTFINDNUMBER").cloneNode(True)
                    xmlLCNode.appendChild xmlElement.selectSingleNode("MORTGAGESUBQUOTENUMBER").cloneNode(True)
                    xmlLCNode.appendChild xmlElement.selectSingleNode("LOANCOMPONENTSEQUENCENUMBER").cloneNode(True)
                    
                    Set xmlTempElem = IMortgageSubQuoteBO_UpdateLoanComponent(xmlLCUpdateRequest)
                    m_objErrAssist.CheckResponse xmlTempElem, xmlElement, True
                End If
            End If
            'MAR46 End
                                                                                                                     
            m_objXmlAssist.ChangeNodeName xmlElement, "LOANCOMPONENT", "MORTGAGEINCENTIVE", False   'BMIDS782 GHun
            
            blnRaiseError = False 'dont raise record not found as an error
            
            Set xmlMortgageIncentiveList = objIMortgageSubquoteDO.FindMortgageIncentiveList(xmlElement)
            
            m_objXmlAssist.ChangeNodeName xmlElement, "MORTGAGEINCENTIVE", "LOANCOMPONENT", False   'BMIDS782 GHun
            
            If Not xmlMortgageIncentiveList Is Nothing Then
                xmlElement.appendChild xmlMortgageIncentiveList
                Set xmlMortgageIncentiveList = Nothing
            End If
            
        Next
    End If
    
    'BMIDS782 GHun
    Set xmlTempElem = m_objXmlAssist.GetNode(vxmlTableElement, "LOANCOMPONENT")
    m_objXmlAssist.ChangeNodeName xmlTempElem, "LOANCOMPONENT", "NONPANELMORTGAGEPRODUCT", False
    'BMIDS782 End
    
    Set xmlTempElem = Nothing
    blnRaiseError = False
    Set xmlTempElem = objIMortgageSubquoteDO.FindNonPanelMortgageProductList(vxmlTableElement)
    
    If Not xmlTempElem Is Nothing Then
        xmlMortgageSubQuoteData.appendChild xmlTempElem
    End If
    
    'SR 27/02/2007 : EP2_1272 - Add IntroducerFeeList to Response
    Set xmlTempRequest = vxmlTableElement.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "INTRODUCERFEE"
    xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "READ"
    
    Set xmlElement = xmlOut.createElement("INTRODUCERFEE")
    xmlElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlTempRequest.appendChild xmlElement
    strTemp = objCrudBO.OmRequest(xmlTempRequest.xml)
    
    Set xmlTempDOMDocument = m_objXmlAssist.load(strTemp, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempDOMDocument.documentElement, xmlResponseElem, True
    Set xmlIntroducerFeeResp = xmlMakeNodeElementBased(xmlTempDOMDocument.documentElement, True, "RESPONSE")
    
    If xmlIntroducerFeeResp.selectNodes(".//INTRODUCERFEE").length > 0 Then
        Set xmlElement = xmlOut.createElement("INTRODUCERFEELIST")
        For Each xmlTempElem In xmlIntroducerFeeResp.selectNodes("//INTRODUCERFEE")
            xmlElement.appendChild xmlTempElem.cloneNode(True)
        Next xmlTempElem
        xmlMortgageSubQuoteData.appendChild xmlElement
    End If
    'SR 27/02/2007 : EP2_1272 - End
    
    Set IMortgageSubQuoteBO_GetLoanCompositionDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetLoanCompositionDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objSQLAssist = Nothing
    Set xmlMPSearch = Nothing
    Set objComboDO = Nothing
    Set objMortgageProductBO = Nothing
    Set objConvertAssist = Nothing
    Set xmlMortgageProductDoc = Nothing
    Set xmlTempElem = Nothing
    Set xmlProductSearchNode = Nothing
    Set xmlMortgageSubQuoteData = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlLoanComponentNode = Nothing
    Set xmlMortgageIncentiveList = Nothing
    Set xmlElement = Nothing
    Set objConvertAssist = Nothing
    'MAR46 GHun
    Set xmlLCUpdateRequest = Nothing
    Set xmlLCNode = Nothing
    Set xmlIntRate = Nothing
    'MAR46 End
    'SR 27/02/2007 : EP2_1272
    Set xmlTempRequest = Nothing
    Set objCrudBO = Nothing
    Set xmlIntroducerFeeResp = Nothing
    'SR 27/02/2007 : EP2_1272 - End
    Exit Function

GetLoanCompositionDetailsVbErr:

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

    Set IMortgageSubQuoteBO_GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoanCompositionDetailsExit
End Function


Private Function IMortgageSubQuoteBO_GetMortgageSubQuoteData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageSubQuoteDataVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetMortgageSubQuoteData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objSQLAssist As New SQLAssist
        
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
                  
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vxmlTableElement.nodeName = "MORTGAGESUBQUOTE" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("MORTGAGESUBQUOTE").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGESUBQUOTE" & " tag not found"
    End If
        
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(xmlRequestNode, _
                                                    "MORTGAGESUBQUOTENUMBER")
                                                        
                                                     
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
           Len(strMortgageSubQuoteNumber) = 0 Then
        
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
    Dim xmlMortgageSubQuoteData As IXMLDOMElement
    'Get the mortgagesubquote record
    Set xmlMortgageSubQuoteData = objIMortgageSubquoteDO.GetData(xmlRequestNode)
       
    'append the mortgagesubquotedata to the return node
    Set xmlMortgageSubQuoteData = xmlResponseElem.appendChild(xmlMortgageSubQuoteData)
  
 
    Set IMortgageSubQuoteBO_GetMortgageSubQuoteData = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetMortgageSubQuoteDataExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objSQLAssist = Nothing
   
    
    Exit Function

GetMortgageSubQuoteDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

  

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_GetMortgageSubQuoteData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
'        objContext.setcomplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageSubQuoteDataExit
End Function

Private Function IMortgageSubQuoteBO_GetOneOffCostDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetOneOffCostDetailsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_GetOneOffCostDetails"

    Dim xmlRequestNode                  As IXMLDOMNode
    Dim xmlDataNode                     As IXMLDOMNode
    Dim xmlOut                          As FreeThreadedDOMDocument40
    Dim xmlResponseElem                 As IXMLDOMElement
    Dim objSQLAssist                    As SQLAssist
    Dim objIMortgageSubquoteDO          As IMortgageSubQuoteDO
    Dim strApplicationNumber            As String
    Dim strApplicationFactFindNumber    As String
    Dim strMortgageSubQuoteNumber       As String
    Dim strDetails                      As String
    Dim xmlOneOffCostsData              As IXMLDOMElement
    Dim xmlTempElem                     As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vxmlTableElement.nodeName = "ONEOFFCOST" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("ONEOFFCOST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ONEOFFCOSTS" & " tag not found"
    End If
            
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, _
                                                  "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, _
                                                     "MORTGAGESUBQUOTENUMBER")
                                                     
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
         
        Set objSQLAssist = New SQLAssist
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
    
    Set xmlOneOffCostsData = objIMortgageSubquoteDO.GetDataForOneOffCostsDetails(vxmlTableElement)
    'add the one off costs to the response element
    xmlResponseElem.appendChild xmlOneOffCostsData
    'then get the total incentives as a separate call
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "ONEOFFCOST", "MORTGAGEINCENTIVE"
    'create element
    Set xmlTempElem = xmlOut.createElement("TOTALINCENTIVES")
    ' and append
    xmlResponseElem.appendChild xmlTempElem
    xmlTempElem.Text = objIMortgageSubquoteDO.TotalIncentivesForAllComponents(vxmlTableElement)
    
    Set IMortgageSubQuoteBO_GetOneOffCostDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetOneOffCostDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objSQLAssist = Nothing
    Set xmlOneOffCostsData = Nothing
    Set xmlTempElem = Nothing
    
    Exit Function

GetOneOffCostDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_GetOneOffCostDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetOneOffCostDetailsExit

End Function

Private Function IMortgageSubQuoteBO_ProcessAddedOneOffCost(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo ProcessAddedOneOffCostVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_ProcessAddedOneOffCost"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlMigDoc As New FreeThreadedDOMDocument40
    xmlMigDoc.validateOnParse = False
    xmlMigDoc.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMPSearch As New FreeThreadedDOMDocument40
    xmlMPSearch.validateOnParse = False
    xmlMPSearch.setProperty "NewParser", True
    Dim xmlCalcMIG As New FreeThreadedDOMDocument40
    xmlCalcMIG.validateOnParse = False
    xmlCalcMIG.setProperty "NewParser", True
    Dim xmlAddedOneOffCosts As New FreeThreadedDOMDocument40
    xmlAddedOneOffCosts.validateOnParse = False
    xmlAddedOneOffCosts.setProperty "NewParser", True
    Dim xmlOneOffCostListElem As IXMLDOMElement
    Dim xmlOneOffCostList As IXMLDOMNodeList
    Dim xmlOneOffCost As IXMLDOMElement
    Dim xmlMortgageSubQuoteData As IXMLDOMElement
    Dim xmlLoanComponentListElem As IXMLDOMElement
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlMortgageProductDoc As New FreeThreadedDOMDocument40
    xmlMortgageProductDoc.validateOnParse = False
    xmlMortgageProductDoc.setProperty "NewParser", True
    Dim xmlElement As IXMLDOMElement
    Dim xmlLoanComponentElement As IXMLDOMElement
    Dim xmlAddedOneOffCostsNode As IXMLDOMElement
    Dim xmlFirstLoanComponent As IXMLDOMElement
    Dim xmlMigNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlCalcMIGNode As IXMLDOMElement
    Dim xmlProductSearchNode As IXMLDOMElement
    
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objComboDO As ComboDO
    Dim objSQLAssist As New SQLAssist
    Dim objMortgageProductBO As MortgageProductBO
    Dim objMigBO As MIGBO
    Dim objConvertAssist As New ConvertAssist
    
    Dim blnRaiseError As Boolean
    Dim blnPartAndPart As Boolean
    Dim bAddedToLoan As Boolean
    Dim blnValue As Boolean
    
    Dim lngCostsToBeAdded As Long
    Dim lngTotalLoanAmount As Long
    
    Dim strOneOffCostType As String
    Dim strTagName As String
    Dim strTemp As String
    Dim strMortgageProduct As String
    Dim strMigValues As String
    Dim strLoanComponentSequenceNumber As String
    Dim strIPT As String
    Dim strText As String
    Dim strLegalFeeType As String
    Dim strMigPremium As String
    Dim strTotalIndemnity As String
    
    Dim intLowestSequenceNumber As Integer
    Dim intTempLowestSequenceNumber As Integer
    Dim intFirstTerm As Integer
    
    Set xmlResponseElem = xmlMigDoc.createElement("RESPONSE")
    xmlMigDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objComboDO = New ComboDO
        Set objMortgageProductBO = New MortgageProductBO
        Set objMigBO = New MIGBO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objMortgageProductBO = objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
        Set objMigBO = objContext.CreateInstance(App.Title & ".MIGBO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
'Stage1:

    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    Dim strLocation As String
    Dim strValuationType As String
    Dim strPurchasePriceOrEstimatedValue As String
    Dim strTypeOfApplication As String
    Dim strLoansNotRedeemed As String
    Dim strMemberOfStaff  As String
    Dim strSharedPercentage  As String
    Dim strApplicationDate As String
    
    strApplicationNumber = objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    '*=[MC]BMIDS763 APPLICATION DATE
    strApplicationDate = objXmlAssist.GetTagValue(xmlRequestNode, "APPLICATIONDATE")
    'SECTION END
    
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(xmlRequestNode, "MORTGAGESUBQUOTENUMBER")
    strLocation = objXmlAssist.GetTagValue(xmlRequestNode, "LOCATION")
    strValuationType = objXmlAssist.GetTagValue(xmlRequestNode, "TYPEOFVALUATION")
    strPurchasePriceOrEstimatedValue = objXmlAssist.GetTagValue(xmlRequestNode, "PURCHASEPRICE")
    strTypeOfApplication = objXmlAssist.GetTagValue(xmlRequestNode, "TYPEOFAPPLICATION")
    strLoansNotRedeemed = objXmlAssist.GetTagValue(xmlRequestNode, "LOANSNOTREDEEMED")
    strMemberOfStaff = objXmlAssist.GetTagValue(xmlRequestNode, "MEMBEROFSTAFF")
    strSharedPercentage = objXmlAssist.GetTagValue(xmlRequestNode, "SHAREDPERCENTAGE")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strLocation) = 0 Or _
        Len(strValuationType) = 0 Or _
        Len(strPurchasePriceOrEstimatedValue) = 0 Or _
        Len(strTypeOfApplication) = 0 Or _
        Len(strLoansNotRedeemed) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Or _
        Len(strMemberOfStaff) = 0 Then
        
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; LOCATION = " & strLocation
        strDetails = strDetails & "; TYPEOFVALUATION = " & strValuationType
        strDetails = strDetails & "; PURCHASEPRICE = " & strPurchasePriceOrEstimatedValue
        strDetails = strDetails & "; TYPEOFAPPLICATION = " & strTypeOfApplication
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        strDetails = strDetails & "; LOANSNOTREDEEMED = " & strLoansNotRedeemed
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, strDetails
        
    End If
    
'Stage2:

    'Preparing Request XML structure to be used in Product search later
    Set xmlTempElem = xmlMPSearch.createElement("MORTGAGEPRODUCT")
    Set xmlProductSearchNode = xmlMPSearch.appendChild(xmlTempElem)

    Set xmlTempElem = xmlMPSearch.createElement("MORTGAGEPRODUCTCODE")
    xmlProductSearchNode.appendChild xmlTempElem
    
    Set xmlTempElem = xmlMPSearch.createElement("STARTDATE")
    xmlProductSearchNode.appendChild xmlTempElem
        
    Set xmlTempElem = xmlMPSearch.createElement("PURCHASEPRICE")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strPurchasePriceOrEstimatedValue
    
    Set xmlTempElem = xmlMPSearch.createElement("TYPEOFAPPLICATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strTypeOfApplication
    
    Set xmlTempElem = xmlMPSearch.createElement("LOCATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strLocation 'passed in
    
    Set xmlTempElem = xmlMPSearch.createElement("TYPEOFVALUATION")
    xmlProductSearchNode.appendChild xmlTempElem
    xmlTempElem.Text = strValuationType ' passed in
    
    Set xmlTempElem = xmlMPSearch.createElement("FEETYPE")
    xmlProductSearchNode.appendChild xmlTempElem
                               
    'Prepare Request XML structure to be used in CalcMIG later
    Set xmlTempElem = xmlCalcMIG.createElement("MIG")
    Set xmlCalcMIGNode = xmlCalcMIG.appendChild(xmlTempElem)
               
    Set xmlTempElem = xmlCalcMIG.createElement("LOANSNOTREDEEMED")
    xmlCalcMIGNode.appendChild xmlTempElem
    xmlTempElem.Text = strLoansNotRedeemed
    
    Set xmlTempElem = xmlCalcMIG.createElement("PURCHASEPRICE")
    xmlCalcMIGNode.appendChild xmlTempElem
    xmlTempElem.Text = strPurchasePriceOrEstimatedValue
                        
    Set xmlTempElem = xmlCalcMIG.createElement("TYPEOFAPPLICATION")
    xmlCalcMIGNode.appendChild xmlTempElem
    xmlTempElem.Text = strTypeOfApplication
    
    Set xmlTempElem = xmlCalcMIG.createElement("MEMBEROFSTAFF")
    xmlCalcMIGNode.appendChild xmlTempElem
    xmlTempElem.Text = strMemberOfStaff
    
    Set xmlTempElem = xmlCalcMIG.createElement("SHAREDPERCENTAGE")
    xmlCalcMIGNode.appendChild xmlTempElem
    xmlTempElem.Text = strSharedPercentage
    
    
'Stage3:
    
    'Identify One-off costs to be added to loan
    
    lngTotalLoanAmount = 0
    lngCostsToBeAdded = 0
    
    Set xmlOneOffCostListElem = xmlRequestNode.getElementsByTagName("ONEOFFCOSTLIST").Item(0)
        
    If xmlOneOffCostListElem Is Nothing Then
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, "ONEOFFCOSTLIST" & " tag not found"

    End If
        
    Set xmlOneOffCostList = xmlOneOffCostListElem.selectNodes(".//ONEOFFCOST")
            
    For Each xmlOneOffCost In xmlOneOffCostList
        
        strOneOffCostType = objXmlAssist.GetTagValue(xmlOneOffCost, "COMBOVALIDATIONTYPE")
        
        'If StrComp(strOneOffCostType, "MIG", vbTextCompare) <> 0 And objConvertAssist.CSafeBool( _
        '        objXmlAssist.GetTagValue(xmlOneOffCost, "ADDEDTOLOAN")) Then
        
        If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlOneOffCost, "ADDEDTOLOAN")) Then
            
            lngCostsToBeAdded = lngCostsToBeAdded + _
                objConvertAssist.CSafeLng(objXmlAssist.GetTagValue(xmlOneOffCost, "AMOUNT"))
        
        End If
                    
    Next
            
'Stage4:

    Set xmlMortgageSubQuoteData = objIMortgageSubquoteDO.GetData(xmlRequestNode)
        
    'Append returned "AMOUNTREQUESTED" value to use in MP search later
    strTagName = "AMOUNTREQUESTED"
    Set xmlTempElem = xmlMortgageSubQuoteData.getElementsByTagName(strTagName).Item(0)
    strTemp = xmlTempElem.Text
   
    Set xmlTempElem = xmlMPSearch.createElement(strTagName)
    xmlTempElem.Text = strTemp
    xmlProductSearchNode.appendChild xmlTempElem

    'Add this to the CalcMIG
    'lngTotalLoanAmount = CLng(strTemp)
    lngTotalLoanAmount = lngCostsToBeAdded + CLng(strTemp)
    
    Set xmlTempElem = xmlCalcMIG.createElement("TOTALLOANAMOUNT")
    xmlTempElem.Text = CStr(lngTotalLoanAmount)
    xmlCalcMIGNode.appendChild xmlTempElem
   
    'append returned "LTV" value to use in MP search later
    
    strTagName = "LTV"
    Set xmlTempElem = xmlMortgageSubQuoteData.getElementsByTagName(strTagName).Item(0)
    strTemp = xmlTempElem.Text
   
    Set xmlTempElem = xmlMPSearch.createElement(strTagName)
    xmlTempElem.Text = strTemp
    xmlProductSearchNode.appendChild xmlTempElem
            
    'append LTV to CalcMIG
    Set xmlTempElem = xmlCalcMIG.createElement("LTV")
    xmlTempElem.Text = strTemp
    xmlCalcMIGNode.appendChild xmlTempElem
    
    'Retrieve Mortgage Product Details
    objXmlAssist.ChangeNodeName xmlRequestNode, "MORTGAGESUBQUOTE", "LOANCOMPONENT"
     
    blnRaiseError = True
     
    Set xmlLoanComponentListElem = objIMortgageSubquoteDO.FindLoanComponentList(xmlRequestNode)
            
    Set xmlLoanComponentList = xmlLoanComponentListElem.selectNodes(".//LOANCOMPONENT")
            
'Stage5:

    'Add this to the CalcMIG

    xmlCalcMIGNode.appendChild xmlLoanComponentListElem
        
    blnValue = True
    
    For Each xmlLoanComponent In xmlLoanComponentList
        strText = objXmlAssist.GetTagValue(xmlLoanComponent, "REPAYMENTMETHOD")
            If Len(strText) = 0 Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, "Missing RepaymentMethod Value"
            End If
        blnValue = objComboDO.IsItemInValidation("RepaymentType", strText, "C")
        If Not blnValue Then 'false
            strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "D")
            Exit For
        End If
    Next
            
    If Len(strLegalFeeType) = 0 Then
        strLegalFeeType = objComboDO.GetFirstComboValueId("LegalFeeType", "R")
    End If
    
    Set xmlTempElem = xmlProductSearchNode.getElementsByTagName("FEETYPE").Item(0)
    xmlTempElem.Text = strLegalFeeType
    
    '*=[MC]Add ApplicationNumber,Date TO THE PRODUCT SEARCH REQUEST
    'BMIDS815 GHun Only need to add application nodes once outside of the loop
    xmlProductSearchNode.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlProductSearchNode.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    xmlProductSearchNode.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONDATE").cloneNode(True)
    'Section end
    
    For Each xmlElement In xmlLoanComponentList
    
        If Not objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlElement, "PORTEDLOAN")) Then
                
            'get the mortgageproductcode for this loancomponent
            strTagName = "MORTGAGEPRODUCTCODE"
            Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
            strTemp = xmlTempElem.Text
            
            Set xmlTempElem = xmlProductSearchNode.getElementsByTagName(strTagName).Item(0)
            xmlTempElem.Text = strTemp
            
            'get the startdate for this loancomponent
            strTagName = "STARTDATE"
            Set xmlTempElem = xmlElement.getElementsByTagName(strTagName).Item(0)
            strTemp = xmlTempElem.Text
            
            Set xmlTempElem = xmlProductSearchNode.getElementsByTagName(strTagName).Item(0)
            xmlTempElem.Text = strTemp
                        
            'BMIDS815 GHun
            Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PORTINGINDICATOR")
            If xmlTempElem Is Nothing Then
                Set xmlTempElem = xmlMigDoc.createElement("PORTINGINDICATOR")
                xmlProductSearchNode.appendChild xmlTempElem
            End If
            strText = xmlGetNodeText(xmlElement, "MANUALPORTEDLOANIND")
            If Len(strText) = 0 Then
                strText = "0"
            End If
            xmlTempElem.Text = strText
            
            If CInt(xmlTempElem.Text) = 1 Then
                Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PRODUCTSTEP")
                If xmlTempElem Is Nothing Then
                    Set xmlTempElem = xmlMigDoc.createElement("PRODUCTSTEP")
                    xmlProductSearchNode.appendChild xmlTempElem
                End If
                xmlTempElem.Text = xmlGetNodeText(xmlElement, "PRODUCTSTEP")
                
                Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "REMAININGSTEPDURATION")
                If xmlTempElem Is Nothing Then
                    Set xmlTempElem = xmlMigDoc.createElement("REMAININGSTEPDURATION")
                    xmlProductSearchNode.appendChild xmlTempElem
                End If
                xmlTempElem.Text = xmlGetNodeText(xmlElement, "REMAININGSTEPDURATION")
                
                Set xmlTempElem = xmlGetNode(xmlProductSearchNode, "PORTEDPRODUCTSTARTDATE")
                If xmlTempElem Is Nothing Then
                    Set xmlTempElem = xmlMigDoc.createElement("PORTEDPRODUCTSTARTDATE")
                    xmlProductSearchNode.appendChild xmlTempElem
                End If
                xmlTempElem.Text = xmlGetNodeText(xmlElement, "PORTEDPRODUCTSTARTDATE")
            End If
            'BMIDS815 End
          
            strMortgageProduct = objMortgageProductBO.GetFullProductDetails(xmlProductSearchNode.xml)
                                                
            Set xmlMortgageProductDoc = objXmlAssist.load(strMortgageProduct, TypeName(Me), strFunctionName)
            
            objErrAssist.CheckResponse xmlMortgageProductDoc.documentElement, xmlResponseElem, True
            objXmlAssist.AttachResponseData xmlElement, xmlMortgageProductDoc.documentElement
        
        End If
                                
        'Dont raise record not found as an error
        blnRaiseError = False
                        
    Next
            
 'Stage6:
 
    'Recalcalculate MIG
                
'    strMigValues = objMigBO.CalcMIG(xmlCalcMIG.xml)
'
'    Set xmlMigDoc = objXmlAssist.load(strMigValues, TypeName(Me), strFunctionName)
'
'    strMigPremium = objXmlAssist.GetTagValue(xmlMigDoc.documentElement, "MIGPREMIUM")
'    strTotalIndemnity = objXmlAssist.GetTagValue(xmlMigDoc.documentElement, "TOTALINDEMNITY")
'    strIPT = objXmlAssist.GetTagValue(xmlMigDoc.documentElement, "IPT")
'
'    If strMigPremium <> "0" And strMigPremium <> "" Then
'        Set xmlOneOffCost = xmlOneOffCostListElem.selectSingleNode( _
'                        ".//ONEOFFCOST[COMBOVALIDATIONTYPE = ""MIG"" ]/AMOUNT")
'        xmlOneOffCost.Text = strMigPremium
'    End If
'
'    bAddedToLoan = False
'
'    Set xmlMigNode = xmlOneOffCostListElem.selectSingleNode( _
'                ".//ONEOFFCOST[COMBOVALIDATIONTYPE = ""MIG"" ]/ADDEDTOLOAN")
'
'    If Not xmlMigNode Is Nothing Then
'        If objConvertAssist.CSafeBool(xmlMigNode.Text) Then
'            lngTotalLoanAmount = lngTotalLoanAmount + objConvertAssist.CSafeLng(strMigPremium)
'            lngCostsToBeAdded = lngCostsToBeAdded + objConvertAssist.CSafeLng(strMigPremium)
'        End If
'    End If
'
'    Set xmlOneOffCost = xmlOneOffCostListElem.selectSingleNode( _
'                        ".//ONEOFFCOST[COMBOVALIDATIONTYPE = ""TID"" ]/AMOUNT")
'
'    If Not xmlOneOffCost Is Nothing Then
'        xmlOneOffCost.Text = strTotalIndemnity
'    End If
    
 'Stage7:
    'Allocate one-off costs added to loan to first loan component
    'SR 06/03/2007 : EP2_1613
    If objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApplication, "PSW") Then
        Set xmlLoanComponentList = xmlLoanComponentListElem.selectNodes(".//LOANCOMPONENT[PRODUCTSWITCHRETAINPRODUCTIND=0]")
    ElseIf objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApplication, "NP") Then
        Set xmlLoanComponentList = xmlLoanComponentListElem.selectNodes(".//LOANCOMPONENT[MANUALPORTEDLOANIND=0]")
    Else  'Get all the loan components where portedloan = FALSE
        xmlLoanComponentListElem.selectNodes (".//LOANCOMPONENT[PORTEDLOAN = 0 ]")
    End If
    'SR 06/03/2007 : EP2_1613 - End
    intLowestSequenceNumber = 0
    
    'Loop round all finding the FIRST Loan Component
    For Each xmlLoanComponentElement In xmlLoanComponentList
    
        intTempLowestSequenceNumber = objConvertAssist.CSafeInt( _
                objXmlAssist.GetTagValue(xmlLoanComponentElement, "LOANCOMPONENTSEQUENCENUMBER"))
        
        If intTempLowestSequenceNumber < intLowestSequenceNumber Or intLowestSequenceNumber = 0 Then
            intLowestSequenceNumber = intTempLowestSequenceNumber
            Set xmlFirstLoanComponent = xmlLoanComponentElement
        End If
            
    Next

'Stage8:

    Set xmlTempElem = xmlFirstLoanComponent.selectSingleNode(".//TOTALLOANCOMPONENTAMOUNT")
                  
    xmlTempElem.Text = CStr(objConvertAssist.CSafeLng( _
                            objXmlAssist.GetTagValue( _
                            xmlFirstLoanComponent, "LOANAMOUNT")) + lngCostsToBeAdded)
     
    Set xmlTempElem = xmlFirstLoanComponent.selectSingleNode(".//REPAYMENTMETHOD")
                    
    blnPartAndPart = objComboDO.IsItemInValidation("RepaymentType", xmlTempElem.Text, "P")
    strLoanComponentSequenceNumber = objXmlAssist.GetTagValue(xmlFirstLoanComponent, "LOANCOMPONENTSEQUENCENUMBER")
    
'Stage9:

    Set xmlTempElem = xmlAddedOneOffCosts.createElement("ADDEDONEOFFCOSTS")
    Set xmlAddedOneOffCostsNode = xmlAddedOneOffCosts.appendChild(xmlTempElem)
 
    'Add the one off cost
    If Not xmlOneOffCostListElem Is Nothing Then
        xmlAddedOneOffCostsNode.appendChild xmlOneOffCostListElem
    End If
    
    'MV - BMIDS00667 - 23/10/2002
    Set xmlTempElem = xmlAddedOneOffCosts.createElement("MIGIPT")
    xmlAddedOneOffCostsNode.appendChild xmlTempElem
    xmlTempElem.Text = "0"

    Set xmlTempElem = xmlAddedOneOffCosts.createElement("TOTALLOANAMOUNT")
    xmlAddedOneOffCostsNode.appendChild xmlTempElem
    xmlTempElem.Text = lngTotalLoanAmount

    Set xmlTempElem = xmlAddedOneOffCosts.createElement("LOANCOMPONENTSEQUENCENUMBER")
    xmlAddedOneOffCostsNode.appendChild xmlTempElem
    xmlTempElem.Text = strLoanComponentSequenceNumber

    Set xmlTempElem = xmlAddedOneOffCosts.createElement("PARTANDPART")
    xmlAddedOneOffCostsNode.appendChild xmlTempElem
    xmlTempElem.Text = CStr(blnPartAndPart)
    
    If Not xmlFirstLoanComponent Is Nothing Then
        'Add single loan component first, clone it so it is not removed from the loanComponentList
        xmlAddedOneOffCostsNode.appendChild xmlFirstLoanComponent.cloneNode(True)
    End If
    
    'Add loan component list
    xmlAddedOneOffCostsNode.appendChild xmlLoanComponentListElem

    xmlResponseElem.appendChild xmlAddedOneOffCostsNode
 
 'Stage10:
 
    Set IMortgageSubQuoteBO_ProcessAddedOneOffCost = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ProcessAddedOneOffCostExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlMigDoc = Nothing
    Set xmlResponseElem = Nothing
    Set objComboDO = Nothing

    Exit Function

ProcessAddedOneOffCostVbErr:

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

    Set IMortgageSubQuoteBO_ProcessAddedOneOffCost = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostExit


End Function

'BMIDS736 GHun This method is only called by commented out code, so I have commented it out as well
'It will need to be updated if it needs to be used
Private Function IMortgageSubQuoteBO_RecostComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:  Updates one-off cost and related info for a mortgage sub-quote
'' pass:         vxmlRequest  xml Request data stream containing data to be persisted
'' return:       xml Response Node
'' Input Format :
''<REQUEST>
''   <MORTGAGESUBQUOTE>
''       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
''       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
''       <CONTEXT>..</CONTEXT>
''       <ONEOFFCOST>
''           ....
''           ....
''       </ONEOFFCOST>
''       <LOANCOMPONENT>
''           ....
''           ....
''       </LOANCOMPONENT>
''       <APPLICATIONDATE>..</APPLICATIONDATE>
''       <PURCHASEPRICE>...</PURCHASEPRICE>
''       <TYPEOFAPPLICATION>...</TYPEOFAPPLICATION>
''       <LOCATION>...</LOCATION>
''       <TYPEOFVALUATION>..</TYPEOFVALUATION>
''       <FEETYPE>...</FEETYPE>
''   </MORTGAGESUBQUOTE>
''</REQUEST>
''------------------------------------------------------------------------------------------
'On Error GoTo RecostComponentVbErr
'
'    Dim strFunctionName As String
'    strFunctionName = "RecostComponent"
'
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
'    Dim xmlRequestNode As IXMLDOMNode
'    Dim xmlNewRequest As IXMLDOMNode
'    Dim xmlTempNode As IXMLDOMNode
'    Dim xmlOut As New FreeThreadedDOMDocument40
'    Dim xmlDoc As FreeThreadedDOMDocument40
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlTableElem As IXMLDOMElement
'    Dim xmlMortgageSubQuote As IXMLDOMElement
'
'    Dim objMortgageProductBO As MortgageProductBO
'    Dim objIMortgageSubQuoteDO As IMortgageSubQuoteDO
'
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    If objContext Is Nothing Then
'        Set objMortgageProductBO = New MortgageProductBO
'        Set objIMortgageSubQuoteDO = New MortgageSubQuoteDO
'    Else
'        Set objMortgageProductBO = objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
'        Set objIMortgageSubQuoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
'    End If
'
'    If vxmlTableElement.nodeName = cstrROOT_NODE_NAME Then
'        Set xmlRequestNode = vxmlTableElement
'    Else
'        Set xmlRequestNode = vxmlTableElement.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
'    End If
'
'    If xmlRequestNode Is Nothing Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
'            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
'    End If
'
'    '----------------------------------------------------------------------------
'    ' Get MortgageSubQuote data
'    '----------------------------------------------------------------------------
'    Set xmlTableElem = xmlOut.createElement(cstrROOT_NODE_NAME)
'    Set xmlTempNode = xmlOut.createElement("APPLICATIONNUMBER")
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode("APPLICATIONNUMBER").Text
'    xmlTableElem.appendChild xmlTempNode
'
'    Set xmlTempNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
'    xmlTableElem.appendChild xmlTempNode
'
'    Set xmlTempNode = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode("LOANCOMPONENT/MORTGAGESUBQUOTENUMBER").Text
'    xmlTableElem.appendChild xmlTempNode
'
'    Set xmlMortgageSubQuote = objIMortgageSubQuoteDO.GetData(xmlTableElem)
'
'    '-----------------------------------------------------------------------
'    'Find full details of MortgageProduct
'    '-----------------------------------------------------------------------
'    Dim strTag As String, strResponse As String
'
'
'    Set xmlTableElem = xmlOut.createElement("MORTGAGEPRODUCT")
'    strTag = "MORTGAGEPRODUCTCODE"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode("LOANCOMPONENT/" & strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "STARTDATE"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode("LOANCOMPONENT/" & strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "AMOUNTREQUESTED"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlMortgageSubQuote.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "PURCHASEPRICE"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "LTV"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlMortgageSubQuote.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "TYPEOFAPPLICATION"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "LOCATION"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "TYPEOFVALUATION"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strTag = "FEETYPE"
'    Set xmlTempNode = xmlOut.createElement(strTag)
'    xmlTempNode.Text = xmlRequestNode.selectSingleNode(strTag).Text
'    xmlTableElem.appendChild xmlTempNode
'
'    strResponse = objMortgageProductBO.GetFullProductDetails(xmlTableElem.xml)
'    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'
' 'MS    21/07/00    performance rework.
'    objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'
'
'    '-------------------------------------------------------------------------
'    'Build the Input Interface to the Calcs engine
'    '-------------------------------------------------------------------------
'    Set xmlNewRequest = vxmlTableElement.cloneNode(True)
'    objXmlAssist.AttachResponseData xmlNewRequest, xmlDoc.documentElement
'
'
'    Set xmlTempNode = xmlOut.createElement("LOANCOMPONENTNUM")
'    xmlTempNode.Text = 1
'    xmlNewRequest.appendChild xmlTempNode
'
'    Set xmlTempNode = IMortgageSubQuoteBO_BuildCalcsEngineInputInterface(xmlNewRequest)
'    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
'
'    '------------------------------------------------------------------------------
'    'Run the calcs engine with the returned data
'    '-----------------------------------------------------------------------------
'    Set xmlTempNode = IMortgageSubQuoteBO_RunCalcsEngine(xmlTempNode)
'    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
'
'    ' Add to the response element and return it
'    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempNode
'    Set IMortgageSubQuoteBO_RecostComponent = xmlResponseElem
'
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
'
'RecostComponentExit:
'
'    Set objXmlAssist = Nothing
'    Set xmlRequestNode = Nothing
'    Set objErrAssist = Nothing
'    Set objMortgageProductBO = Nothing
'    Set objIMortgageSubQuoteDO = Nothing
'    Set xmlTempNode = Nothing
'    Set xmlNewRequest = Nothing
'    Set xmlOut = Nothing
'    Set xmlDoc = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlTableElem = Nothing
'
'    Exit Function
'
'RecostComponentVbErr:
'
'    If objErrAssist.IsWarning = True Then
'        objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'
'    objErrAssist.AddToErrSource strFunctionName
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    Set IMortgageSubQuoteBO_RecostComponent = objErrAssist.CreateErrorResponseEx()
'
'    If Not objContext Is Nothing Then
'        objContext.SetAbort
'    End If
'
'    ' go to clean-up section
'    Resume RecostComponentExit
End Function
'BMIDS736 End

Private Function IMortgageSubQuoteBO_ReserveMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ReserveMortgageProductVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_ReserveMortgageProduct"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        

    'validate the ported loan status here

    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.ReserveMortgageProduct(vxmlTableElement)
   
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_ReserveMortgageProduct = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ReserveMortgageProductExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    
    Exit Function

ReserveMortgageProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_ReserveMortgageProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReserveMortgageProductExit

End Function

Private Function IMortgageSubQuoteBO_ResetMortgageSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_ResetMortgageSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.ResetMortgageSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IMortgageSubQuoteBO_ResetMortgageSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ResetMortgageSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IMortgageSubQuoteBO_ResetMortgageSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit

End Function

Private Function IMortgageSubQuoteBO_SaveIllustrationResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveIllustrationResultsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_SaveIllustrationResults"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.SaveIllustrationResults(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_SaveIllustrationResults = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveIllustrationResultsExit:
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    Set IMortgageSubQuoteBO_SaveIllustrationResults = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveIllustrationResultsExit
End Function

Private Function IMortgageSubQuoteBO_SaveIncentives(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveIncentivesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_SaveIncentives"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.SaveIncentives(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_SaveIncentives = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveIncentivesExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    
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

    Set IMortgageSubQuoteBO_SaveIncentives = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveIncentivesExit

End Function

Private Function IMortgageSubQuoteBO_SaveLoanComponentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveLoanComponentDetailsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_SaveLoanComponentDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'THIS NEEDS TO BE MOVED TO HARD CODED VALIDATION
    'Validate missing data
    
    Dim strPurposeOfLoan As String
    Dim strLoanAmount As String
    Dim strRepaymentType As String
    Dim strTermInYears As String
    Dim strTermInMonths As String
    Dim strMortgageProductCode As String
    Dim strStartDate As String
        
    strPurposeOfLoan = objXmlAssist.GetTagValue(vxmlTableElement, "PURPOSEOFLOAN")
    strLoanAmount = objXmlAssist.GetTagValue(vxmlTableElement, "LOANAMOUNT")
    strRepaymentType = objXmlAssist.GetTagValue(vxmlTableElement, "REPAYMENTTYPE")
    strTermInYears = objXmlAssist.GetTagValue(vxmlTableElement, "TERMINYEARS")
    strTermInMonths = objXmlAssist.GetTagValue(vxmlTableElement, "TERMINMONTHS")
    strMortgageProductCode = objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGEPRODUCTCODE")
    strStartDate = objXmlAssist.GetTagValue(vxmlTableElement, "STARTDATE")
                                                     
    If Len(strPurposeOfLoan) = 0 Or Len(strLoanAmount) = 0 Or _
       Len(strRepaymentType) = 0 Or Len(strTermInYears) = 0 Or _
       Len(strTermInMonths) = 0 Or Len(strMortgageProductCode) = 0 Or _
       Len(strStartDate) = 0 Then
        Dim strDetails As String
        
        strDetails = "PURPOSEOFLOAN = " & objSQLAssist.FormatString(strPurposeOfLoan)
        strDetails = strDetails & "; LOANAMOUNT = " & strLoanAmount
        strDetails = strDetails & "; REPAYMENTTYPE = " & strRepaymentType
        strDetails = strDetails & "; TERMINYEARS = " & strTermInYears
        strDetails = strDetails & "; TERMINMONTHS = " & strTermInMonths
        strDetails = strDetails & "; MORTGAGEPRODUCTCODE = " & strMortgageProductCode
        strDetails = strDetails & "; STARTDATE = " & strStartDate
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                195, strDetails
    End If


    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.SaveLoanComponentDetails(vxmlTableElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IMortgageSubQuoteBO_SaveLoanComponentDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveLoanComponentDetailsExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function

SaveLoanComponentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_SaveLoanComponentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLoanComponentDetailsExit

End Function

Private Function IMortgageSubQuoteBO_SaveOneOffCostDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveOneOffCostDetailsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_SaveOneOffCostDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.SaveOneOffCostDetails(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IMortgageSubQuoteBO_SaveOneOffCostDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveOneOffCostDetailsExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveOneOffCostDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_SaveOneOffCostDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOneOffCostDetailsExit

End Function

Private Function IMortgageSubQuoteBO_Update(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_Update"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIMortgageSubQuoteTxBO = New MortgageSubQuoteTxBO
    Else
        Set objIMortgageSubQuoteTxBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.Update(vxmlTableElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IMortgageSubQuoteBO_Update = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IMortgageSubQuoteBO_Update = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit

End Function

Private Function IMortgageSubQuoteBO_ValidateCompulsoryProducts(ByVal vXMLRequest _
                                        As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCompulsoryProductsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_ValidateCompulsoryProducts"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlCompulsoryBC As IXMLDOMElement
    Dim xmlCompulsoryPP As IXMLDOMElement
'    strTagName = "ONEOFFCOSTSEQUENCENUMBER"
    Set xmlCompulsoryBC = xmlOut.createElement("COMPULSORYBCFLAG")
    xmlCompulsoryBC.Text = "0" 'FALSE
    xmlResponseElem.appendChild xmlCompulsoryBC
    
    Set xmlCompulsoryPP = xmlOut.createElement("COMPULSORYPPFLAG")
    xmlCompulsoryPP.Text = "0" 'FALSE
    xmlResponseElem.appendChild xmlCompulsoryPP
    
        
    Dim xmlCloneNode As IXMLDOMElement
    
    Set xmlCloneNode = vXMLRequest.cloneNode(True)
            
    objXmlAssist.ChangeNodeName xmlCloneNode, cstrROOT_NODE_NAME, "LOANCOMPONENT"
    
    Set xmlTempElem = objIMortgageSubquoteDO.GetLoanComponentProductParameters(xmlCloneNode)
            
    Dim xmlMPParametersList As IXMLDOMNodeList
    Dim xmlMPParameters As IXMLDOMElement
    Dim objConvertAssist As New ConvertAssist
                
    Set xmlMPParametersList = xmlTempElem.selectNodes(".//MORTGAGEPRODUCTPARAMETERS")
        
    For Each xmlMPParameters In xmlMPParametersList
    
    
        'Still false
        If StrComp(xmlCompulsoryBC.Text, "0", vbTextCompare) = 0 Then
        
            If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue( _
                                            xmlMPParameters, "COMPULSORYBC")) Then
                
                'Set The COMPULSORYBC Flag
                xmlCompulsoryBC.Text = "1"
                
                If StrComp(xmlCompulsoryPP.Text, "0", vbTextCompare) = 0 Then
                    'Both flags set to true so exit loop
                    Exit For
                    
                End If
            
            End If
        
        End If
        
        'Still false
        If StrComp(xmlCompulsoryPP.Text, "0", vbTextCompare) = 0 Then
        
            If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue( _
                                            xmlMPParameters, "COMPULSORYPP")) Then
        
                'Set The COMPULSORYPP Flag
                xmlCompulsoryPP.Text = "1"
                
                If StrComp(xmlCompulsoryBC.Text, "0", vbTextCompare) = 0 Then
                    'Both flags set to true so exit loop
                    Exit For
                    
                End If
                
            End If
    
        End If
    
    Next
                
    Set IMortgageSubQuoteBO_ValidateCompulsoryProducts = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ValidateCompulsoryProductsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objXmlAssist = Nothing

    Exit Function

ValidateCompulsoryProductsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_ValidateCompulsoryProducts = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateCompulsoryProductsExit

End Function

Private Function IMortgageSubQuoteBO_ValidateSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateSubQuoteVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_ValidateSubQuote"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim inumPortedLoans As Integer

    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim objIMortgageProductDO As IMortgageProductDO
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767
    'Dim objApplicationBO As Object
   
    If objContext Is Nothing Then
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objIMortgageProductDO = objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductDO")
    End If
        
    
    If vxmlTableElement.nodeName = "QUOTATION" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("QUOTATION").Item(0)
    End If
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "QUOTATION" & " tag not found"
    End If
    
    Dim xmlListNode As IXMLDOMNodeList
    Dim xmlReturnElement As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    
    ' Create default response block
    Dim xmlResponseElem As IXMLDOMElement
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim xmlReturnElem As IXMLDOMElement
    objXmlAssist.ChangeNodeName xmlRequestNode, "QUOTATION", "LOANCOMPONENT"
    
    ' Subquote is still valid if there are no loancomponents
    ' APS 08/06/00 SYS0833 - You cannot have a valid Mortgage subquote without loan components
    On Error GoTo RecordNotFoundErr
    Set xmlReturnElem = objIMortgageSubquoteDO.FindLoanComponentList(xmlRequestNode)
    On Error GoTo ValidateSubQuoteVbErr
    ' DLM 10/07/00 SYS0949 Default to True because it is not an error to have a Mortgage
    ' sub quote without loan components
    Dim blnValidProduct As Boolean
    blnValidProduct = True
    If Not xmlReturnElem Is Nothing Then
        Set xmlListNode = xmlReturnElem.getElementsByTagName("LOANCOMPONENT")
    
        Dim xmlPortedLoanComponentsDoc As New FreeThreadedDOMDocument40
        xmlPortedLoanComponentsDoc.validateOnParse = False
        xmlPortedLoanComponentsDoc.setProperty "NewParser", True
        Dim xmlTempElem As IXMLDOMElement
        Dim xmlTempPortedLoanElem As IXMLDOMElement
    
        Set xmlTempElem = xmlPortedLoanComponentsDoc.createElement("PORTEDLOANCOMPONENTLIST")
        Set xmlTempPortedLoanElem = xmlPortedLoanComponentsDoc.appendChild(xmlTempElem)
    
        Dim objConvertAssist As New ConvertAssist
        Dim blnPortedLoan As Boolean
        Dim nLoop As Integer

        Dim xmlTempElement2 As IXMLDOMElement
            
        For nLoop = 0 To (xmlListNode.length - 1)
        
                
            Set xmlTempElement = xmlListNode.Item(nLoop)
            Dim strPattern As String

            'SG 28/05/02 SYS4767 START
            'STB: SYS4219 - If A specific product was selected for this loan
            'component then we don't need to validate it.
            If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlTempElement, "PRODUCTCODESEARCHIND")) = False Then
                'we will only need to check this in cost modelling as qq doesn't have portedloan xml
                If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlTempElement, "PORTEDLOAN")) = False Then
                
              
                    Set xmlTempElement2 = vxmlTableElement.getElementsByTagName("ISMEMBEROFSTAFF").Item(0)
                    'Fix me here
                    Set xmlTempElement2 = xmlTempElement.appendChild(xmlTempElement2.cloneNode(True))
                
                    'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
                    'Set xmlTempElement2 = vxmlTableElement.getElementsByTagName("ISSELFEMPLOYED").Item(0)
                    'Set xmlTempElement2 = xmlTempElement.appendChild(xmlTempElement2.cloneNode(True))
                
                    'BMIDS00515 SA 24/10/02 {
                    Set xmlTempElement2 = vxmlTableElement.getElementsByTagName("LTV").Item(0)
                    
                    'BMIDS00765 MO 31/10/2002 {
                    If Not xmlTempElement2 Is Nothing Then
                
                        Set xmlTempElement2 = xmlTempElement.appendChild(xmlTempElement2.cloneNode(True))
                        'BMIDS00515 }
                    End If
                    'BMIDS00765 }
                                    
                    blnValidProduct = objIMortgageProductDO.GetValidActiveProduct(xmlTempElement)
                
                    If Not blnValidProduct Then
                    '
                        Exit For
                
                    End If
            'SG 28/05/02 SYS4767 END
            
                End If
        
            End If
        
        Next
    End If
    
    'Reset the name
    objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENT", "QUOTATION"
    
    Dim xmlTempElement3 As IXMLDOMElement
    
    Set xmlTempElement3 = vxmlTableElement.getElementsByTagName("QUOTATION").Item(0)
    
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlPortedLoanList As IXMLDOMNodeList
    
    'create element to return
    Set xmlReturnNode = xmlResponseElem.ownerDocument.createElement("VALIDMORTGAGESUBQUOTE")
   
    'select nodes that are ported
    strPattern = ".//LOANCOMPONENT[PORTEDLOAN = 1 ]"
    If Not xmlReturnElem Is Nothing Then
        Set xmlPortedLoanList = xmlReturnElem.selectNodes(strPattern)
    End If
        
    'Need to validate the LOANCOMPONENT XML passed in as well
    Dim xmlTempElemIn As IXMLDOMElement
    Dim xmlPortedLoanInList As IXMLDOMNodeList
    
    Set xmlTempElemIn = vxmlTableElement.getElementsByTagName("LOANSTOBEPORTED").Item(0)

    If xmlTempElemIn Is Nothing Then 'No XML PORTED LOANS passed in
        'if ported loan component list > 0
        If Not xmlPortedLoanList Is Nothing Then
            If xmlPortedLoanList.length > 0 Then
                blnValidProduct = False
            End If
        End If
    Else
        'if we get here we have xml containing portedloancomponents passed in

        
        Set xmlPortedLoanInList = xmlTempElemIn.getElementsByTagName("MORTGAGELOAN")
        'BG SYS0924 30/11/00
        If xmlPortedLoanList Is Nothing Then
            inumPortedLoans = 0
        Else
            inumPortedLoans = xmlPortedLoanList.length
        End If
                    
        If xmlPortedLoanInList.length <> inumPortedLoans Then
        'BG SYS0924 30/11/00 END
            blnValidProduct = False
        Else
        
            Dim nLoop1 As Integer
            Dim strMortgageLoanGuid As String
            Dim strMonthlyRepayment As String
            Dim strOutstandingBalance As String
            Dim strRepaymentType As String
            Dim blnPortedLoanDiscrepancy As Boolean
            blnPortedLoanDiscrepancy = True
            'Loop round this list comparing against the other one
            For nLoop1 = 0 To (xmlPortedLoanList.length - 1)
                                
                'get the first one out
                Set xmlTempElement = xmlPortedLoanList.Item(nLoop1)
                
                strMortgageLoanGuid = objXmlAssist.GetTagValue(xmlTempElement, "MORTGAGELOANGUID")
                
                strMonthlyRepayment = objXmlAssist.GetTagValue(xmlTempElement, "MONTHLYREPAYMENT")
                strOutstandingBalance = objXmlAssist.GetTagValue(xmlTempElement, "OUTSTANDINGBALANCE")
                strRepaymentType = objXmlAssist.GetTagValue(xmlTempElement, "REPAYMENTTYPE")
                
                Dim strMortgageLoanGuidIn As String
                Dim strMonthlyRepaymentIn As String
                Dim strLoanAmountIn As String
                Dim strTotalLoanAmountIn As String
                Dim strRepaymentMethodIn As String
                
                Dim nLoop2 As Integer
                
                For nLoop2 = 0 To (xmlPortedLoanInList.length - 1)
                
                    Set xmlTempElement2 = xmlPortedLoanInList.Item(nLoop2)
                    
                    strMortgageLoanGuidIn = objXmlAssist.GetTagValue(xmlTempElement2, "MORTGAGELOANGUID")
                    
                    strMonthlyRepaymentIn = objXmlAssist.GetTagValue(xmlTempElement2, "MONTHLYREPAYMENT")
                    strLoanAmountIn = objXmlAssist.GetTagValue(xmlTempElement2, "LOANAMOUNT")
                    strTotalLoanAmountIn = objXmlAssist.GetTagValue(xmlTempElement2, "TOTALLOANCOMPONENTAMOUNT")
                    strRepaymentMethodIn = objXmlAssist.GetTagValue(xmlTempElement2, "REPAYMENTMETHOD")
                    
                    If StrComp(strMortgageLoanGuid, strMortgageLoanGuidIn, vbTextCompare) = 0 Then
                    
                        If StrComp(strMonthlyRepayment, strMonthlyRepaymentIn, vbTextCompare) = 0 And _
                            StrComp(strOutstandingBalance, strLoanAmountIn, vbTextCompare) = 0 And _
                            StrComp(strOutstandingBalance, strTotalLoanAmountIn, vbTextCompare) = 0 And _
                            StrComp(strRepaymentType, strRepaymentMethodIn, vbTextCompare) = 0 Then
                        
                        blnPortedLoanDiscrepancy = False
                        
                        End If
                    
                    End If
                Next
                
                If blnPortedLoanDiscrepancy = True Then
                    
                    blnValidProduct = False
                    
                    Exit For
                    
                End If
            Next
        
        
        End If
    
    End If
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767 START
    'JLD MSMS0034 check that the TypeOfApplication and TypeOfBuyer on the
    '  mortgagesubquote match those on the applicationFactFind if we are not in
    '  Quick Quote mode
    'Dim xmlAppDoc As New FreeThreadedDOMDocument40
    'Dim xmlAppReqElem As IXMLDOMElement
    'Dim xmlAppAppElem As IXMLDOMElement
    'Dim xmlAppTmpNode As IXMLDOMNode
    'Dim xmlAppRespDoc As New FreeThreadedDOMDocument40
    'Dim xmlMSQRespNode As IXMLDOMNode
    'Dim xmlAppRespElem As IXMLDOMElement
    'Dim xmlMSQTypeOfAppNode As IXMLDOMNode
    'Dim xmlMSQTypeOfBuyerNode As IXMLDOMNode
    'Dim xmlIsQQNode As IXMLDOMNode
    
    'Set xmlIsQQNode = vxmlTableElement.selectSingleNode(".//ISQUICKQUOTE")
    'If xmlIsQQNode Is Nothing Then
    
        ' get the application data
    '    Set xmlAppReqElem = xmlAppDoc.createElement("REQUEST")
    '    xmlAppDoc.appendChild xmlAppReqElem
    '    Set xmlAppAppElem = xmlAppReqElem.appendChild(xmlAppDoc.createElement("APPLICATION"))
    '    Set xmlAppTmpNode = xmlAppAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONNUMBER"))
    '    xmlAppTmpNode.Text = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER", True, True)
    '    Set xmlAppTmpNode = xmlAppAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFINDNUMBER"))
    '    xmlAppTmpNode.Text = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER", True, True)
        
    '    Set objApplicationBO = objContext.CreateInstance("omApp.ApplicationBO")
    '    xmlAppRespDoc.loadXML objApplicationBO.GetApplicationData(xmlAppDoc.documentElement.xml)
    '    Set xmlAppRespElem = xmlAppRespDoc.selectSingleNode("RESPONSE")
    '
    '    ' get the mortgageSubQuote data (modify the AppDoc request)
    '    objXmlAssist.ChangeNodeName xmlAppAppElem, "APPLICATION", "MORTGAGESUBQUOTE"
    '   Set xmlAppTmpNode = xmlAppAppElem.appendChild(xmlAppDoc.createElement("MORTGAGESUBQUOTENUMBER"))
    '    xmlAppTmpNode.Text = objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER", True, True)
    '    Set xmlMSQRespNode = objIMortgageSubQuoteDO.GetData(xmlAppAppElem)
    '
    '    Set xmlMSQTypeOfAppNode = xmlMSQRespNode.selectSingleNode(".//TYPEOFAPPLICATION")
    '    Set xmlMSQTypeOfBuyerNode = xmlMSQRespNode.selectSingleNode(".//TYPEOFBUYER")
    '    If Not xmlMSQTypeOfAppNode Is Nothing Then
    '        If Len(xmlMSQTypeOfAppNode.Text) > 0 And xmlMSQTypeOfAppNode.Text <> objXmlAssist.GetTagValue(xmlAppRespElem, "TYPEOFAPPLICATION", True, True) Then
    '            blnValidProduct = False
    '        End If
    '    End If
    '    If Not xmlMSQTypeOfBuyerNode Is Nothing Then
    '        If Len(xmlMSQTypeOfBuyerNode.Text) > 0 And xmlMSQTypeOfBuyerNode.Text <> objXmlAssist.GetTagValue(xmlAppRespElem, "TYPEOFBUYER", True, True) Then
    '            blnValidProduct = False
    '        End If
    '    End If
    'End If
    ' END MSMS0034 JLD
    'SG 28/05/02 SYS4767 END

    xmlReturnNode.Text = Abs(blnValidProduct)
    'xmlTempElement3.appendChild xmlReturnNode
    'xmlResponseElem.appendChild xmlTempElement3
    xmlResponseElem.appendChild xmlReturnNode
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set IMortgageSubQuoteBO_ValidateSubQuote = xmlOut.documentElement
    

ValidateSubQuoteExit:
    
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubquoteDO = Nothing
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767 START
    'Set objApplicationBO = Nothing
    'Set xmlAppDoc = Nothing
    'Set xmlAppReqElem = Nothing
    'Set xmlAppAppElem = Nothing
    'Set xmlAppTmpNode = Nothing
    'Set xmlAppRespDoc = Nothing
    'Set xmlMSQRespNode = Nothing
    'Set xmlAppRespElem = Nothing
    'Set xmlMSQTypeOfAppNode = Nothing
    'Set xmlMSQTypeOfBuyerNode = Nothing
    'Set xmlIsQQNode = Nothing
    'SG 28/05/02 SYS4767 END
    
    Exit Function

RecordNotFoundErr:

    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
        ' DLM 10/07/00 AQR SYS0949 It is not an error to have no records
        Resume Next
        'm_objErrAssist.ThrowError TypeName(Me), strFunctionName, 290
    End If

ValidateSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_ValidateSubQuote = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ValidateSubQuoteExit

End Function

Private Sub SetUpCharges(ByVal vobjOneOffChargesElem As IXMLDOMElement, _
                         ByVal vobjXmlAppendToElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Builds the the input parameters concerned with charges
'
' pass:         vobjOneOffChargesElem     xml containing the One Off Charges Details
'
'               vobjXmlAppendToElem       The output element that the Charges
'                                         parameters are to be appended to
'
' return:       n/a
'
' Raise Errors:
'
'------------------------------------------------------------------------------------------
On Error GoTo SetUpChargesVbErr
    
    Const strFunctionName As String = "SetUpCharges"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objElement As IXMLDOMElement
    Dim objFeeElement As IXMLDOMElement
    Dim objFeeList As IXMLDOMNodeList
    
    Dim strTagName As String
    Dim strTagValue As String
    Dim strFeeType As String
    
    Dim dblDeedsReleaseFee As Double
    Dim dblSealingFee As Double
    Dim dblTotalFinalCharges As Double
    Dim dblTotalNonMirasInitCharges As Double   ' PSC 11/06/2002 BMIDS00043
    
    Dim strPattern As String
    
    ' If one off costs are present then calculate charges
    If Not vobjOneOffChargesElem Is Nothing Then
        
        ' Calculate the total final charges based on Sealing Fee and Deeds Release Fee
        strPattern = ".//ONEOFFCOST[IDENTIFIER = ""DEE"" or IDENTIFIER = ""SEA""]"
        Set objFeeList = vobjOneOffChargesElem.selectNodes(strPattern)
        
        For Each objFeeElement In objFeeList
            strTagValue = objXmlAssist.GetTagValue(objFeeElement, "AMOUNT")
            dblTotalFinalCharges = dblTotalFinalCharges + objConvAssist.CSafeDbl(strTagValue)
        Next
    
        ' Calculate the total non MIRAS initial charges by totalling charges that are not
        ' Sealing Fee, Deeds Release Fee, Stamp Duty or Total Indemnity
        strPattern = ".//ONEOFFCOST[IDENTIFIER != ""DEE"" "
        strPattern = strPattern & "and IDENTIFIER != ""SEA"" "
        strPattern = strPattern & "and IDENTIFIER != ""STA"" "
        strPattern = strPattern & "and IDENTIFIER != ""TID""]"
        
        Set objFeeList = vobjOneOffChargesElem.selectNodes(strPattern)
                
        For Each objFeeElement In objFeeList
            strTagValue = objXmlAssist.GetTagValue(objFeeElement, "AMOUNT")
            
            ' PSC 11/06/2002 BMIDS00043 - Start
            dblTotalNonMirasInitCharges = _
                    dblTotalNonMirasInitCharges + objConvAssist.CSafeDbl(strTagValue)
            ' PSC 11/06/2002 BMIDS00043 - End
        Next
    End If
    
    ' PSC 11/06/2002 BMIDS00043 - Set INP440_TOTALMIRASINITIALCHARGES to zero
    strTagName = "INP440_TOTALMIRASINITIALCHARGES"
    Set objElement = vobjXmlAppendToElem.ownerDocument.createElement(strTagName)
    objElement.Text = "0"
    vobjXmlAppendToElem.appendChild objElement

    ' PSC 11/06/2002 BMIDS00043 - Set INP450_TOTALNONMIRASINITIALCHARGES to dblTotalNonMirasInitCharges
    strTagName = "INP450_TOTALNONMIRASINITIALCHARGES"
    Set objElement = vobjXmlAppendToElem.ownerDocument.createElement(strTagName)
    objElement.Text = CStr(dblTotalNonMirasInitCharges)
    vobjXmlAppendToElem.appendChild objElement
    
    strTagName = "INP460_TOTALFINALCHARGES"
    Set objElement = vobjXmlAppendToElem.ownerDocument.createElement(strTagName)
    objElement.Text = CStr(dblTotalFinalCharges)
    vobjXmlAppendToElem.appendChild objElement

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objElement = Nothing
    Set objFeeElement = Nothing
    Set objFeeList = Nothing
    
    Exit Sub

SetUpChargesVbErr:

    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objElement = Nothing
    Set objFeeElement = Nothing
    Set objFeeList = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function GetDebugFileName() As String
'BMIDS736 GHun
' header ----------------------------------------------------------------------------------
' description:  Gets the path to save debugging info to from the registry.
'               Creates the folder if it does not exist.
' pass:         none
' return:       A string containing a valid path name. Returns "" if the path does not exist.
'------------------------------------------------------------------------------------------
    Const cstrRegKey    As String = "HKLM\SOFTWARE\Omiga4\System Configuration\AlphaPlusDebugPath"
    Dim objWshShell     As Object
    Dim strDebugPath    As String
    Dim strRandom       As String
       
' Ignore any errors that occur
On Error GoTo GetDebugFileNameVbErr
    Set objWshShell = CreateObject("WScript.Shell")
    ' Read the path to save debugging info to from the registry
    strDebugPath = Trim$(objWshShell.RegRead(cstrRegKey))
    If Len(strDebugPath) > 0 Then
        If Right$(strDebugPath, 1) <> "\" Then
            strDebugPath = strDebugPath & "\"
        End If
        If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
            'Create the folder if it does not exist
            MkDir strDebugPath
            If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
                ' The output folder does not exist and cannot be created
                strDebugPath = vbNullString
            End If
        End If
    End If
    
    If Len(strDebugPath) > 0 Then
        'Add a random number so the filename will hopefully still be unique if multiple
        'requests are made in the same second
        Randomize
        strRandom = Format$(Rnd * 10000, "0000")
    
        strDebugPath = strDebugPath & Format$(Now(), "YYYYMMDD_HHNNSS") & "_" & strRandom & ".txt"
    End If
        
    GetDebugFileName = strDebugPath
    Set objWshShell = Nothing
    Exit Function

GetDebugFileNameVbErr:
    GetDebugFileName = vbNullString
    Set objWshShell = Nothing
    'Clear any errors that may have occurred as they can be ignored
    Err.Clear
End Function
'BMIDS736 End

'BMIDS736 GHun 21/04/2004 Save debugging info to a file
Private Sub SaveDebugInfo(ByVal vstrFileName As String, ByVal vstrMessage As String, ByVal vstrXML As String)
    Dim objFSO          As Object
    Dim objTextStream   As Object
    
On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextStream = objFSO.OpenTextFile(vstrFileName, 8, True)
    objTextStream.WriteLine "------------------------------------------------------------------------------------------"
    objTextStream.WriteLine Format$(Now(), "DD/MM/YYYY HH:NN:SS") & ": " & vstrMessage
    objTextStream.WriteLine "------------------------------------------------------------------------------------------"
    objTextStream.WriteLine vstrXML & vbCrLf
    objTextStream.Close
    
    Set objFSO = Nothing
    Set objTextStream = Nothing
    If Err.number <> 0 Then
        Err.Clear
    End If
End Sub
'BMIDS736 End

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

Private Function IMortgageSubQuoteBO_Validate(ByVal vXMLRequest As IXMLDOMElement, _
                                    ByVal veboMethod As IOmiga4BOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IMortgageSubQuoteBO_Validate"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
     
    Select Case veboMethod
    
    Case bomResetMortgageSubQuote, _
         bomReserveMortgageProduct, _
         bomSaveLoanComponentDetails, _
         bomCalcAndSaveMortgageCosts, _
         bomCreateOneOffCosts
         
            If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
                Set xmlRequestNode = vXMLRequest
            Else
                Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
            End If
            
            If xmlRequestNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
            End If
            
    Case bomSaveIllustrationResults, _
         bomUpdateLoanComponent
            If vXMLRequest.nodeName = cstrLOANCOMP_NODE_NAME Then
                Set xmlRequestNode = vXMLRequest
            Else
                Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrLOANCOMP_NODE_NAME).Item(0)
            End If
            
            If xmlRequestNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, cstrLOANCOMP_NODE_NAME & " tag not found"
            End If
    Case Else
    
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Validation method not implemented"
    End Select
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageSubQuoteBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetData"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_GetData(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetData = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function CreateOneOffCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateOneOffCostsVbErr

    Const strFunctionName As String = "CreateOneOffCosts"
    
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomCreateOneOffCosts)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_CreateOneOffCosts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateOneOffCosts = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateOneOffCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

CreateOneOffCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateOneOffCosts = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateOneOffCostsExit

End Function

'BMIDS736 GHun
Private Function BuildCalcsEngineInputInterface(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
    Const cstrFunctionName      As String = "BuildCalcsEngineInputInterface"
    Const cstrAlphaDateFormat   As String = "YYYY-MM-DD"
    Dim objComboDO              As ComboDO
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc       As FreeThreadedDOMDocument40
    Dim xslDoc                  As FreeThreadedDOMDocument40
    Dim xmlElem                 As IXMLDOMElement
    Dim xmlResponseElem         As IXMLDOMElement
    Dim xmlRepayment            As IXMLDOMElement
    Dim xmlInterestRate         As IXMLDOMElement
    Dim xmlControl              As IXMLDOMElement
    Dim xmlElemGroup            As IXMLDOMElement
    Dim xmlLoanComponent        As IXMLDOMNode
    Dim xmlNode                 As IXMLDOMNode
    Dim xmlValitionTypes        As IXMLDOMNode
    Dim xmlPeriod               As IXMLDOMNode
    Dim xmlLoanCompList         As IXMLDOMNodeList
    Dim strRepayment            As String
    Dim strResponse             As String
    Dim strValue                As String
    Dim strXSLPath              As String
    Dim dblManAdjust            As Double
    Dim dblBandAdjust           As Double
    Dim dblRate                 As Double
    Dim dblIncreasedRate        As Double
    Dim dteStartDate            As Date
    Dim dteRateEndDate          As Date
    Dim strRateEndDate          As String
    Dim intPeriod               As Integer
    Dim blnRateIncreased        As Boolean
    'BMIDS767 Start
    Dim xmlOneOffCostList As IXMLDOMNodeList
    Dim xmlOneOffCostListNode As IXMLDOMNode
    Dim xmlOneOffCost As IXMLDOMNode
    Dim xmlValuationListNode As IXMLDOMNode
    Dim strIdentifier As String
    Dim xmlCostTypeNode As IXMLDOMNode
    Dim xmlTestNode As IXMLDOMNode
    'BMIDS767 End
    Dim dteCompletionDate       As Date 'MAR1777 GHun
    Dim xmlAddBorrowing         As IXMLDOMNode  'EP2_954 GHun
    
On Error GoTo BuildCalcsEngineInputInterfaceVBErr
       
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
               
    'Set the current date in the format expected by Alpha+
    vXMLRequest.setAttribute "DATE", Format$(Date, cstrAlphaDateFormat)

    'Get the validation types for RepaymentType
    Set xmlNode = xmlDoc.createElement("LIST")
    Set xmlElem = xmlDoc.createElement("LISTNAME")
    xmlElem.Text = "RepaymentType"
    xmlNode.appendChild xmlElem
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strResponse = objComboDO.GetComboList(xmlNode.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    Set xmlLoanCompList = vXMLRequest.selectNodes("//LOANCOMPONENTLIST/LOANCOMPONENT")
       
    'Interest Rate band adjustments are no longer used, so always use 0
    dblBandAdjust = 0

    'MAR1777 GHun
    dteCompletionDate = CalcExpectedCompletionDate(Date)
    vXMLRequest.setAttribute "COMPLETIONDATE", Format$(dteCompletionDate, cstrAlphaDateFormat)
    'MAR1777 End

    'Loop through each loan component
    For Each xmlLoanComponent In xmlLoanCompList
        
        'If the completion date exists, convert it to the Alpha+ expected format
        Set xmlElem = xmlLoanComponent.selectSingleNode("COMPLETIONDATE")
        If Not xmlElem Is Nothing Then
            xmlElem.Text = Format$(CSafeDate(xmlElem.Text), cstrAlphaDateFormat)
        End If
        
        'Add the appropriate RepaymentType validation type to each RepaymentMethod
        Set xmlRepayment = xmlGetMandatoryNode(xmlLoanComponent, "REPAYMENTMETHOD")
        strRepayment = xmlRepayment.Text
        Set xmlValitionTypes = xmlDoc.selectSingleNode("LIST/LISTNAME/LISTENTRY[GROUPNAME='RepaymentType' and VALUEID='" & strRepayment & "']/VALIDATIONTYPELIST")
        If Not xmlValitionTypes Is Nothing Then
            xmlRepayment.appendChild xmlValitionTypes.cloneNode(True)
        End If
        
        strValue = m_objXmlAssist.GetTagValue(xmlLoanComponent, "MANUALADJUSTMENTPERCENT")
        dblManAdjust = CSafeDbl(strValue)
        
        blnRateIncreased = False
        dteStartDate = dteCompletionDate    'MAR1777 GHun
                
        For Each xmlInterestRate In xmlLoanComponent.selectNodes("MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE")
            dblRate = CalculateInterestRate(xmlInterestRate, dblManAdjust, dblBandAdjust)
            Set xmlElem = xmlDoc.createElement("ALPHAPLUSRATE")
            xmlElem.Text = CStr(dblRate)
            xmlInterestRate.appendChild xmlElem
            
            'The 1% increase only applies to the first variable rate where a full 1% increase can be applied
            If Not (blnRateIncreased) Then
                dblIncreasedRate = CalculateIncreasedInterestRate(dblRate, xmlInterestRate)
                If dblIncreasedRate > dblRate Then
                    blnRateIncreased = True
                    Set xmlElem = xmlDoc.createElement("ALPHAPLUSINCREASEDRATE")
                    xmlElem.Text = CStr(dblIncreasedRate)
                    xmlInterestRate.appendChild xmlElem
                End If
            End If
                        
            'Convert InterestRateEndDates to InterestRatePeriods
            Set xmlPeriod = xmlGetNode(xmlInterestRate, "INTERESTRATEPERIOD")
            If xmlPeriod Is Nothing Then
                'If there isn't a InterestRatePeriod node already then create one
                xmlPeriod = xmlDoc.createElement("INTERESTRATEPERIOD")
                xmlInterestRate.appendChild xmlPeriod
                intPeriod = 0
            Else
                intPeriod = CSafeInt(xmlPeriod.Text)
            End If
            
            'Rate period is to be calculated using end dates
            If intPeriod = 0 Then
                'Must have an end date
                strRateEndDate = xmlGetNodeText(xmlInterestRate, "INTERESTRATEENDDATE")
                If IsDate(strRateEndDate) = False Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                           "Invalid Interest Rate End Date: " & strRateEndDate
                End If
    
                dteRateEndDate = CSafeDate(strRateEndDate)
    
                'End date must be after start date
                If dteRateEndDate <= dteStartDate Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                            "Interest Rate End Date must be in the future"  'JD BMIDS826 changed error message
                End If
    
                'BMIDS929 GHun
                'xmlPeriod.Text = DateDiff("m", dteStartDate, dteRateEndDate)
                xmlPeriod.Text = CStr(MonthDiff(dteStartDate, dteRateEndDate))
                'BMIDS929 End
                dteStartDate = dteRateEndDate
            Else
                If intPeriod <> -1 Then ' -1 denotes rest of term
                    dteStartDate = DateAdd("m", intPeriod, dteStartDate)
                End If
            End If
        Next
        
        dteStartDate = Date
        
        'Convert RedemptionFee PeriodEndDates to RedemptionFee Periods
        For Each xmlElem In xmlLoanComponent.selectNodes("MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND")
            
            Set xmlPeriod = xmlGetNode(xmlElem, "PERIOD")
            If xmlPeriod Is Nothing Then
                'If there isn't a RedemptionFee Period node already then create one
                xmlPeriod = xmlDoc.createElement("PERIOD")
                xmlElem.appendChild xmlPeriod
                intPeriod = 0
            Else
                intPeriod = CSafeInt(xmlPeriod.Text)
            End If
            
            'RedemptionFee period is to be calculated using end dates
            If intPeriod = 0 Then
                'Must have an end date
                strRateEndDate = xmlGetNodeText(xmlElem, "PERIODENDDATE")
                If IsDate(strRateEndDate) = False Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                           "Invalid RedemptionFee Period End Date: " & strRateEndDate
                End If
    
                dteRateEndDate = CSafeDate(strRateEndDate)
    
                'End date must be after start date
                If dteRateEndDate <= dteStartDate Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                            "RedemptionFee Period End Date must be in the future" 'JD BMIDS826 changed error message
                End If
    
                'BMIDS929 GHun
                'xmlPeriod.Text = DateDiff("m", dteStartDate, dteRateEndDate)
                xmlPeriod.Text = CStr(MonthDiff(dteStartDate, dteRateEndDate))
                'BMIDS929 End
                dteStartDate = dteRateEndDate
            Else
                If intPeriod <> -1 Then ' -1 denotes rest of term
                    dteStartDate = DateAdd("m", intPeriod, dteStartDate)
                End If
            End If
        Next
    Next
    
    'EP2_954 GHun Additional Borrowing
    'Loop through each mortgage loan
    For Each xmlLoanComponent In vXMLRequest.selectNodes("//MORTGAGELOANLIST/MORTGAGELOAN")
        
        'Add the appropriate RepaymentType validation type to each RepaymentMethod
        Set xmlRepayment = xmlGetMandatoryNode(xmlLoanComponent, "REPAYMENTTYPE")
        strRepayment = xmlRepayment.Text
        Set xmlValitionTypes = xmlDoc.selectSingleNode("LIST/LISTNAME/LISTENTRY[GROUPNAME='RepaymentType' and VALUEID='" & strRepayment & "']/VALIDATIONTYPELIST")
        If Not xmlValitionTypes Is Nothing Then
            xmlRepayment.appendChild xmlValitionTypes.cloneNode(True)
        End If
        
        'Manual Adjustment will always be 0 for additional borrowing
        dblManAdjust = 0
        
        blnRateIncreased = False
        dteStartDate = dteCompletionDate
                
        For Each xmlInterestRate In xmlLoanComponent.selectNodes("INTERESTRATETYPE")
            dblRate = CalculateInterestRate(xmlInterestRate, dblManAdjust, dblBandAdjust)
            Set xmlElem = xmlDoc.createElement("ALPHAPLUSRATE")
            xmlElem.Text = CStr(dblRate)
            xmlInterestRate.appendChild xmlElem
                                    
            'Convert InterestRateEndDates to InterestRatePeriods
            Set xmlPeriod = xmlGetNode(xmlInterestRate, "INTERESTRATEPERIOD")
            If xmlPeriod Is Nothing Then
                'If there isn't a InterestRatePeriod node already then create one
                xmlPeriod = xmlDoc.createElement("INTERESTRATEPERIOD")
                xmlInterestRate.appendChild xmlPeriod
                intPeriod = 0
            Else
                intPeriod = CSafeInt(xmlPeriod.Text)
            End If
            
            'Rate period is to be calculated using end dates
            If intPeriod = 0 Then
                'Must have an end date
                strRateEndDate = xmlGetNodeText(xmlInterestRate, "INTERESTRATEENDDATE")
                If IsDate(strRateEndDate) = False Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                           "Invalid Interest Rate End Date: " & strRateEndDate
                End If
    
                dteRateEndDate = CSafeDate(strRateEndDate)
    
                'End date must be after start date
                If dteRateEndDate <= dteStartDate Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 222, _
                                            "Interest Rate End Date must be in the future"  'JD BMIDS826 changed error message
                End If
    
                'BMIDS929 GHun
                'xmlPeriod.Text = DateDiff("m", dteStartDate, dteRateEndDate)
                xmlPeriod.Text = CStr(MonthDiff(dteStartDate, dteRateEndDate))
                'BMIDS929 End
                dteStartDate = dteRateEndDate
            Else
                If intPeriod <> -1 Then ' -1 denotes rest of term
                    dteStartDate = DateAdd("m", intPeriod, dteStartDate)
                End If
            End If
        Next
    Next
    'EP2_954 End
    
    'BMIDS 767 Start
    'Remove Non APR One OffCosts from request
    'Get the validation types for OneOffCost

    Set xmlNode = xmlDoc.createElement("LIST")
    Set xmlElem = xmlDoc.createElement("LISTNAME")
    xmlElem.Text = "OneOffCost"
    xmlNode.appendChild xmlElem
    strResponse = objComboDO.GetComboList(xmlNode.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    
    Set xmlOneOffCostListNode = vXMLRequest.selectSingleNode("//LOANCOMPONENTLIST/LOANCOMPONENT/ONEOFFCOSTLIST")
    'Set xmlOneOffCostList = vXMLRequest.selectNodes("//LOANCOMPONENTLIST/LOANCOMPONENT/ONEOFFCOSTLIST/ONEOFFCOST")
    'BMIDS774 May not be any OneOffCosts added to the loan, check OneOffCostList exists before trying to use it.
    If Not xmlOneOffCostListNode Is Nothing Then
        Set xmlOneOffCostList = xmlOneOffCostListNode.selectNodes("ONEOFFCOST")
        If Not xmlOneOffCostList Is Nothing Then
            For Each xmlOneOffCost In xmlOneOffCostList
                'Check for IDENTIFIER NODE
                Set xmlCostTypeNode = m_objXmlAssist.GetNode(xmlOneOffCost, "IDENTIFIER", False)
                If xmlCostTypeNode Is Nothing Then
                    'No IDENTIFIER so Check for COMBOVALIDATIONTYPE NODE
                    Set xmlCostTypeNode = m_objXmlAssist.GetNode(xmlOneOffCost, "COMBOVALIDATIONTYPE", True)
                End If
                strIdentifier = xmlCostTypeNode.Text
                Set xmlValuationListNode = xmlDoc.selectSingleNode("LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE = '" & strIdentifier & "']")
                ' Check that there is also a validation type of APR for this One Off Cost
                Set xmlTestNode = m_objXmlAssist.GetNode(xmlValuationListNode, ".//VALIDATIONTYPE[.='APR']", False)
                If xmlTestNode Is Nothing Then
                    'Take it off the list
                    xmlOneOffCostListNode.removeChild xmlOneOffCost
                Else
                    'BMIDS936 GHun
                    Set xmlTestNode = xmlGetNode(xmlValuationListNode, ".//VALIDATIONTYPE[.='END']")
                    If Not xmlTestNode Is Nothing Then
                        Set xmlTestNode = xmlDoc.createElement("PAIDATEND")
                        xmlTestNode.Text = "1"
                        xmlOneOffCost.appendChild xmlTestNode
                    End If
                    'BMIDS936 End
                End If
            Next
        End If
    End If
    
    'BMIDS 767 End
    Set xslDoc = New FreeThreadedDOMDocument40
    xslDoc.validateOnParse = False
    xslDoc.setProperty "NewParser", True
    strXSLPath = Replace(App.Path, "DLL", "XML")
    
    'Transform the Omiga request XML to AlphaPlus format
    If Not xslDoc.load(strXSLPath & "\AlphaPlusRequest.xsl") Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in loading XSL File"
    End If
    
    Set xmlTransformedDoc = New FreeThreadedDOMDocument40
    xmlTransformedDoc.validateOnParse = False
    xmlTransformedDoc.setProperty "NewParser", True
    Set xmlNode = vXMLRequest.cloneNode(True)
    xmlNode.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
            
    If xmlTransformedDoc.parseError <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaParseError, xmlTransformedDoc.parseError.reason
    End If
        
    'Set RepaymentChargeTerms by merging and deduplicating all the EarlyRepaymentCharge durations
    Set xmlControl = xmlGetMandatoryNode(xmlTransformedDoc, "Request/Inputs/Control")
    If Not xslDoc.load(strXSLPath & "\RepaymentChargeTerms.xsl") Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in loading XSL File"
    End If
    strResponse = xmlControl.transformNode(xslDoc.documentElement)
    
    If Not xmlDoc.loadXML(strResponse) Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaParseError, "Error after XSL tranformation"
    End If
    
    strRepayment = xmlGetAttributeText(xmlDoc.documentElement, "RepaymentChargeTerms")
    If Len(Trim$(strRepayment)) > 0 Then
        xmlControl.setAttribute "RepaymentChargeTerms", Trim$(strRepayment)
    End If
    
    For Each xmlElem In xmlTransformedDoc.selectNodes("Request/Inputs/Control/Mortgage/ElementGroup")
        Set xmlElemGroup = xmlGetNode(xmlDoc, "RESPONSE/ElementGroup[@Id='" & xmlElem.getAttribute("Id") & "']")
        If Not xmlElemGroup Is Nothing Then
            strValue = xmlGetAttributeText(xmlElemGroup, "ParameterOverrideIds", vbNullString)
            If Len(Trim$(strValue)) > 0 Then
                xmlElem.setAttribute "ParameterOverrideIds", strValue
            End If
        End If
    Next
        
    Set BuildCalcsEngineInputInterface = xmlTransformedDoc.documentElement
    
BuildCalcsEngineInputInterfaceExit:
    Set objComboDO = Nothing
    Set xmlDoc = Nothing
    Set xmlTransformedDoc = Nothing
    Set xslDoc = Nothing
    Set xmlElem = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRepayment = Nothing
    Set xmlInterestRate = Nothing
    Set xmlControl = Nothing
    Set xmlElemGroup = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlNode = Nothing
    Set xmlValitionTypes = Nothing
    Set xmlPeriod = Nothing
    Set xmlLoanCompList = Nothing
    Set xmlAddBorrowing = Nothing   'EP2_954 GHun
    
    Exit Function

BuildCalcsEngineInputInterfaceVBErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    Set BuildCalcsEngineInputInterface = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume BuildCalcsEngineInputInterfaceExit
End Function
'BMIDS736 End

Public Function SaveIllustrationResults(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveIllustrationResultsVbErr

    Const strFunctionName As String = "SaveIllustrationResults"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomSaveIllustrationResults)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_SaveIllustrationResults(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveIllustrationResults = xmlResponseElem.xml
    
    m_objContext.SetComplete

SaveIllustrationResultsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    SaveIllustrationResults = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume SaveIllustrationResultsExit

End Function

Public Function FindLoanComponentList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLoanComponentListVbErr

    Const strFunctionName As String = "FindLoanComponentList"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_FindLoanComponentList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindLoanComponentList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindLoanComponentListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindLoanComponentListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindLoanComponentList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindLoanComponentListExit

End Function

Public Function FindMortgageIncentiveList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'MortgageIncentive'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindMortgageIncentiveListVbErr

    Const strFunctionName As String = "FindMortgageIncentiveList"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_FindMortgageIncentiveList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindMortgageIncentiveList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindMortgageIncentiveListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindMortgageIncentiveListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindMortgageIncentiveList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindMortgageIncentiveListExit

End Function

Public Function FindAvailableincentives(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
'This procedure takes in as a request data regarding the loan component and mortgage sub quote,
'as a string, converts it into XML and passes it to IMortgageSubQuoteBO_FindAvailableINcentives
'the response is then returned to calling GUI screen as string
'------------------------------------------------------------------------------------------
 On Error GoTo FindAvailableIncentivesVbErr

    Const strFunctionName As String = "FindAvailableincentives"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_FindAvailableincentives(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindAvailableincentives = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindAvailableIncentivesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindAvailableIncentivesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindAvailableincentives = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindAvailableIncentivesExit

End Function

Private Function IMortgageSubQuoteBO_FindLoanComponentList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLoanComponentListVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_FindLoanComponentList"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIMortgageSubquoteDO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageSubquoteDO.FindLoanComponentList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set IMortgageSubQuoteBO_FindLoanComponentList = xmlResponseElem

FindLoanComponentListExit:

    Set xmlRequestNode = Nothing
    Set objIMortgageSubquoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindLoanComponentListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_FindLoanComponentList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindLoanComponentListExit

End Function

Public Function UpdateLoanComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanComponentVbErr

    Const strFunctionName As String = "UpdateLoanComponent"
    
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IMortgageSubQuoteBO_Validate(xmlIn.documentElement, bomUpdateLoanComponent)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IMortgageSubQuoteBO_UpdateLoanComponent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLoanComponent = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateLoanComponentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
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

    UpdateLoanComponent = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateLoanComponentExit
    
End Function

Private Function IMortgageSubQuoteBO_UpdateLoanComponent(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanComponentVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_UpdateLoanComponent"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIMortgageSubQuoteTxBO As IMortgageSubQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIMortgageSubQuoteTxBO = m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteTxBO")
    
    Set xmlTempResponseNode = objIMortgageSubQuoteTxBO.UpdateLoanComponent(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IMortgageSubQuoteBO_UpdateLoanComponent = xmlResponseElem

    m_objContext.SetComplete
    
UpdateLoanComponentExit:

    Set objIMortgageSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
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

    Set IMortgageSubQuoteBO_UpdateLoanComponent = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateLoanComponentExit

End Function

Private Function GetBaseRate(ByVal vXMLRequest As IXMLDOMNode, _
                Optional ByVal blnRateChange As Boolean = False) As Double
' header ----------------------------------------------------------------------------------
' description:  Returns the current Base Interest Rate related to a Base Rate Set
'
' pass:         vxmlRequest  xml containing a Base Rate Set value
'               blnRateChange   Indicates if performing calculation as part of
'                               a Rate Change batch.
'
' return:       The current interest rate associated with the BaseRateSet/BaseRateBand
'
' notes:        The rate returned is based on BaseRate.BaseInterestRate
'                                                       + BaseRateBand.RateDifference
'------------------------------------------------------------------------------------------
On Error GoTo GetBaseRateVbErr
    
Const strFunctionName As String = "GetBaseRate"

Dim objConvAssist As New ConvertAssist
Dim objMortgageProductBO As MortgageProductBO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlNode As IXMLDOMNode

Dim strResponse As String
Dim strRate As String
Dim strAdjust As String

    'Get BaseRateSet node
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(vXMLRequest, ".//BASERATESET")
    
    If m_objContext Is Nothing Then
        Set objMortgageProductBO = New MortgageProductBO
    Else
        Set objMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
        
    'Create the Request
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlRequest.appendChild xmlNode
    'SYS3440 MDC 10/12/2001.
    Set xmlNode = xmlDoc.createElement("RATECHANGE")
    xmlNode.Text = CStr(Abs(CInt(blnRateChange)))
    xmlRequest.appendChild xmlNode
    'SYS3440 End
    
    'Get Base Rate data
    strResponse = objMortgageProductBO.GetRateForBaseRateSet(xmlRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlTempResponse, True
            
    'Return Base Interest Rate (with adjustment)
    strRate = m_objXmlAssist.GetMandatoryElementText(xmlDoc.documentElement, "//BASEINTERESTRATE")
    strAdjust = m_objXmlAssist.GetElementText(vXMLRequest, ".//RATEDIFFERENCE")
    GetBaseRate = objConvAssist.CSafeDbl(strRate) + objConvAssist.CSafeDbl(strAdjust)
    
    'Dereference objects
    Set objConvAssist = Nothing
    Set objMortgageProductBO = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    
    Exit Function

GetBaseRateVbErr:
    'Dereference objects
    Set objConvAssist = Nothing
    Set objMortgageProductBO = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function CalculateInterestRate(ByVal vxmlInterestRateType As IXMLDOMNode, _
                                       ByVal vdblManAdjustment As Double, _
                                       ByVal vdblBandAdjustment As Double) As Double
'-----------------------------------------------------------------------------------------
' description:  Calculate the actual interest rate given an InterestRateType node and
'               the appropriate adjustments
'
' pass:         vxmlInterestRateType    XML containing a interest rate type data
'               vdblManAdjustment       Manual adjustment percentage
'               vdblBandAdjustment      Interest rate band adjustment percentage
'
'
' return:       The calculated interest rate
'
' notes:        The rate returned is based on BaseRate.BaseInterestRate
'                                                       + BaseRateBand.RateDifference
'------------------------------------------------------------------------------------------
On Error GoTo CalculateInterestRateErr

    Const cstrFunctionName As String = "CalculateInterestRate"
    
    Dim dblBaseRate As Double
    Dim dblDiscRate As Double
    Dim dblFlooredRate As Double
    Dim dblCappedRate As Double
    Dim dblIntRate As Double
    Dim dblFixedRate As Double
    
    Dim strRateType As String
    Dim strTagValue As String
    
    Dim objConvAssist As ConvertAssist
    Set objConvAssist = New ConvertAssist
            
    dblBaseRate = objConvAssist.CSafeDbl(m_objXmlAssist.GetMandatoryElementText _
                                        (vxmlInterestRateType, "BASERATEBAND/RATE"))
    dblBaseRate = dblBaseRate + vdblManAdjustment
    
    strRateType = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "RATETYPE")
        
    Select Case strRateType
        Case "B"        ' Base Rate
            dblIntRate = dblBaseRate + vdblBandAdjustment
                
        Case "D", "C"   ' Discounted or Capped/Floored
            strTagValue = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "RATE")
            dblDiscRate = objConvAssist.CSafeDbl(strTagValue)
            dblIntRate = dblBaseRate - dblDiscRate + vdblBandAdjustment
                                
            ' If Capped/Floored set the interest rate to the capped/floored
            ' rate if the calculated rate is above/below these
            If strRateType = "C" Then
                'BMIDS736 A blank FlooredRate should be treated as no floor rather than 0 (as a result of CSafeDbl)
                strTagValue = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "FLOOREDRATE")
                If Len(Trim$(strTagValue)) > 0 Then
                    dblFlooredRate = objConvAssist.CSafeDbl(strTagValue)
                    If dblIntRate < dblFlooredRate Then
                        dblIntRate = dblFlooredRate
                    End If
                End If
                
                'BMIDS736 A blank CeilingRate should be treated as no cap rather than 0 (as a result of CSafeDbl)
                strTagValue = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "CEILINGRATE")
                If Len(Trim$(strTagValue)) > 0 Then
                    dblCappedRate = objConvAssist.CSafeDbl(strTagValue)
                    If dblIntRate > dblCappedRate Then
                        dblIntRate = dblCappedRate
                    End If
                End If
            End If
                
        Case "F"        ' Fixed
            strTagValue = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "RATE")
            dblFixedRate = objConvAssist.CSafeDbl(strTagValue)

            If dblFixedRate = 0 Then
                dblIntRate = dblBaseRate
            Else
                dblIntRate = dblFixedRate + vdblBandAdjustment + vdblManAdjustment
            End If
    End Select

    Set objConvAssist = Nothing
    
    CalculateInterestRate = dblIntRate
    
    Exit Function

CalculateInterestRateErr:
    
    Set objConvAssist = Nothing

    m_objErrAssist.AddToErrSource (cstrFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description
                                                                      
End Function
                                       
Public Function RunCompositeRateCalc(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo RunCompositeRateCalcVbErr

    Const strFunctionName As String = "RunCompositeRateCalc"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '*=[MC] if Object is nothing create an instance.
    
    If m_objXmlAssist Is Nothing Then
        Set m_objXmlAssist = New XMLAssist
    End If
    
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_RunCompositeRateCalc(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    RunCompositeRateCalc = xmlResponseElem.xml
    
    m_objContext.SetComplete

RunCompositeRateCalcExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RunCompositeRateCalcVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    RunCompositeRateCalc = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume RunCompositeRateCalcExit

End Function

Private Function IMortgageSubQuoteBO_RunCompositeRateCalc(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RunCompositeRateCalcVbErr
    
    Const cstrFunctionName As String = "IMortgageSubQuoteBO_RunCompositeRateCalc"

    Dim objAlpha            As ALPHACOMPLUSLib.Alpha
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlError            As IXMLDOMElement
    Dim strResponse         As String
    Dim blnDebug            As Boolean
    Dim strDebugFileName    As String
    'EP2_954 GHun
    Dim xmlAddBorrowing     As IXMLDOMNode
    Dim xmlElem             As IXMLDOMElement
    Dim xmlMortgage         As IXMLDOMNode
    Dim xmlNextSibling      As IXMLDOMNode
    Dim xmlABDoc            As FreeThreadedDOMDocument40
    'EP2_954 End
    
    Set xmlDoc = xmlCreateDOMObject
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Check if XML should be saved for debugging purposes
    strDebugFileName = GetDebugFileName
    If Len(strDebugFileName) > 0 Then
        blnDebug = True
        SaveDebugInfo strDebugFileName, "Original Request:", vXMLRequest.xml
    Else
        blnDebug = False
    End If
    
    Set xmlRequest = BuildCalcsEngineInputInterface(vXMLRequest)
    
    'EP2_954 GHun Extract additional borrowing info, if present
    Set xmlAddBorrowing = xmlRequest.selectSingleNode("/Request/Inputs/Control/Mortgage/ADDITIONALBORROWING")
    If Not xmlAddBorrowing Is Nothing Then
        Set xmlElem = xmlDoc.createElement("TEMP")
        xmlElem.appendChild xmlAddBorrowing
    End If
    'EP2_954
    
    If blnDebug Then
        SaveDebugInfo strDebugFileName, "Request (after BuildCalcsEngineInputInterface):", vXMLRequest.xml
        SaveDebugInfo strDebugFileName, "Alpha+ Request:", xmlRequest.xml
    End If
    
    'GHun Only check the output if it is in Omiga format
    If xmlRequest.baseName = "RESPONSE" Then
        errCheckXMLResponseNode xmlRequest, xmlResponseElem, True
    End If
    
    'Call Alpha Plus
    If Not m_objContext Is Nothing Then
        Set objAlpha = m_objContext.CreateInstance("ALPHACOMPLUS.Alpha")
    Else
        Set objAlpha = New ALPHACOMPLUSLib.Alpha
    End If
    
    strResponse = objAlpha.aceRequest(xmlRequest.xml)
    xmlDoc.loadXML strResponse
    
    If blnDebug Then
        SaveDebugInfo strDebugFileName, "Response from Alpha+:", xmlDoc.xml
    End If

    'Check if any errors occurred
    Set xmlError = xmlDoc.selectSingleNode("Response/Errors/Error")
    If Not xmlError Is Nothing Then
        'BMIDS782 GHun Display a more detailed error message
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeAlphaPlusError, xmlError.getAttribute("Description") & vbCrLf & xmlError.getAttribute("MessageDetail")
    End If

    'EP2_954 GHun Call alpha+ again for Additional Borrowing
    If Not xmlAddBorrowing Is Nothing Then
        'Strip out the non relevant parts of the previous request
        Set xmlMortgage = xmlRequest.selectSingleNode("/Request/Inputs/Control/Mortgage")
        'Remove all the nodes from the first ElementGroup onwards
        Set xmlElem = xmlMortgage.selectSingleNode("ElementGroup")
        Do
            Set xmlNextSibling = xmlElem.nextSibling
            xmlMortgage.removeChild xmlElem
            Set xmlElem = xmlNextSibling
        Loop Until xmlNextSibling Is Nothing
        
        Set xmlElem = xmlMortgage.parentNode
        If Not xmlElem.Attributes.getNamedItem("RepaymentChargeTerms") Is Nothing Then
            xmlElem.removeAttribute "RepaymentChargeTerms"
        End If
        
        'Add the additional borrowing nodes to the request
        For Each xmlElem In xmlAddBorrowing.childNodes
            xmlMortgage.appendChild xmlElem
        Next
    
        If blnDebug Then
            SaveDebugInfo strDebugFileName, "Additional Borrowing: Alpha+ Request:", xmlRequest.xml
        End If
    
        strResponse = objAlpha.aceRequest(xmlRequest.xml)
        Set xmlABDoc = xmlCreateDOMObject
        xmlABDoc.loadXML strResponse
        
        If blnDebug Then
            SaveDebugInfo strDebugFileName, "Additional Borrowing: Response from Alpha+:", xmlABDoc.xml
        End If
    
        'Check if any errors occurred
        Set xmlError = xmlABDoc.selectSingleNode("Response/Errors/Error")
        If Not xmlError Is Nothing Then
            'BMIDS782 GHun Display a more detailed error message
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeAlphaPlusError, xmlError.getAttribute("Description") & vbCrLf & xmlError.getAttribute("MessageDetail")
        End If
    End If
    'EP2_954 End

    'Convert the XML output from Alpha+ to a format usable by Omiga
    'EP2_954 GHun
    If xmlABDoc Is Nothing Then
        Set xmlResponseElem = ConvertAlphaPlusOutput(vXMLRequest, xmlDoc.documentElement)
    Else
        Set xmlResponseElem = ConvertAlphaPlusOutput(vXMLRequest, xmlDoc.documentElement, xmlABDoc.documentElement)
    End If
    'EP2_954 End
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If blnDebug Then
        SaveDebugInfo strDebugFileName, "Transformed Alpha+ response:", xmlResponseElem.xml
    End If
        
    Set IMortgageSubQuoteBO_RunCompositeRateCalc = xmlResponseElem
   
RunCompositeRateCalcExit:
    Set objAlpha = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequest = Nothing
    Set xmlError = Nothing
    Exit Function
    
RunCompositeRateCalcVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    Set IMortgageSubQuoteBO_RunCompositeRateCalc = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume RunCompositeRateCalcExit
End Function

Public Function SaveCompositeRateResults(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveCompositeRateResultsVbErr

    Const strFunctionName As String = "SaveCompositeRateResults"
                
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_SaveCompositeRateResults(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    SaveCompositeRateResults = xmlResponseElem.xml
    
    m_objContext.SetComplete

SaveCompositeRateResultsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
   
    Exit Function

SaveCompositeRateResultsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    SaveCompositeRateResults = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume SaveCompositeRateResultsExit

End Function

Private Function IMortgageSubQuoteBO_SaveCompositeRateResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the results for each loan component passed in
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveCompositeRateResultsVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteBO_SaveCompositeRateResults"
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlSaveRequest As IXMLDOMNode
    'Dim xmlCurrentComp As IXMLDOMNode
    'Dim xmlLoanComp As IXMLDOMNode
    'Dim xmlLoanCompList As IXMLDOMNodeList
    Dim xmlSaveResponse As IXMLDOMNode
    'Dim xmlContext As IXMLDOMNode
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                      
    If vXMLRequest.nodeName = cstrLOANCOMPLIST_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrLOANCOMPLIST_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrLOANCOMPLIST_NODE_NAME & " tag not found"
    End If
    
    Set xmlSaveRequest = vXMLRequest.cloneNode(True)
    
    'BMIDS736 GHun
    'Set xmlLoanCompList = xmlRequestNode.selectNodes("LOANCOMPONENT")
    '
    'Set xmlContext = m_objXmlAssist.GetNode(vXMLRequest, ".//CONTEXT", False)
    '
    '' Save each loan component
    'For Each xmlLoanComp In xmlLoanCompList
    '    Set xmlCurrentComp = xmlLoanComp.cloneNode(True)
    '    xmlSaveRequest.appendChild xmlCurrentComp
    '
    '    If Not xmlContext Is Nothing Then
    '        xmlCurrentComp.appendChild xmlContext.cloneNode(True)
    '    End If
    '
    '    Set xmlSaveResponse = IMortgageSubQuoteBO_SaveIllustrationResults(xmlSaveRequest)
    '    m_objErrAssist.CheckResponse xmlSaveResponse, xmlResponseElem, True
    '
    '    xmlSaveRequest.removeChild xmlCurrentComp
    'Next
            
    Set xmlSaveResponse = IMortgageSubQuoteBO_SaveIllustrationResults(xmlSaveRequest)
    m_objErrAssist.CheckResponse xmlSaveResponse, xmlResponseElem, True
    'BMIDS736 End
            
    Set IMortgageSubQuoteBO_SaveCompositeRateResults = xmlResponseElem
    
    m_objContext.SetComplete

SaveCompositeRateResultsExit:
    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlSaveRequest = Nothing
    'Set xmlCurrentComp = Nothing
    'Set xmlLoanComp = Nothing
    'Set xmlLoanCompList = Nothing
    Set xmlSaveResponse = Nothing
    'Set xmlContext = Nothing
    
    Exit Function

SaveCompositeRateResultsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_SaveCompositeRateResults = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume SaveCompositeRateResultsExit

End Function

Public Function RunCompositeRateCalcAndSaveResults(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo RunCompositeRateCalcAndSaveResultsVbErr

    Const strFunctionName As String = "RunCompositeRateCalcAndSaveResults"
               
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageSubQuoteBO_RunCompositeRateCalcAndSaveResults(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    RunCompositeRateCalcAndSaveResults = xmlResponseElem.xml
    
    m_objContext.SetComplete

RunCompositeRateCalcAndSaveResultsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RunCompositeRateCalcAndSaveResultsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    RunCompositeRateCalcAndSaveResults = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume RunCompositeRateCalcAndSaveResultsExit

End Function

Private Function IMortgageSubQuoteBO_RunCompositeRateCalcAndSaveResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RunCompositeRateCalcAndSaveResultsVbErr
    
    Const strFunctionName   As String = "IMortgageSubQuoteBO_RunCompositeRateCalcAndSaveResults"
   
    Dim xmlRequestNode      As IXMLDOMNode
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlCalcsResponse    As IXMLDOMNode
    'Dim xmlSaveRequest     As IXMLDOMNode
    Dim xmlSaveResponse     As IXMLDOMNode
    'Dim xmlLoanCompList    As IXMLDOMNode
    Dim xmlNode             As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                      
    Set xmlCalcsResponse = IMortgageSubQuoteBO_RunCompositeRateCalc(vXMLRequest)
    m_objErrAssist.CheckResponse xmlCalcsResponse, xmlResponseElem, True
    
    'BMIDS736 GHun
    'Set xmlSaveRequest = vXMLRequest.cloneNode(False)
    'Set xmlLoanCompList = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "LOANCOMPONENTLIST")
    'xmlSaveRequest.appendChild xmlLoanCompList
    'Set xmlSaveResponse = IMortgageSubQuoteBO_SaveCompositeRateResults(xmlSaveRequest)
    Set xmlSaveResponse = IMortgageSubQuoteBO_SaveCompositeRateResults(xmlCalcsResponse)
    'BMIDS736 End
       
    m_objErrAssist.CheckResponse xmlSaveResponse, xmlResponseElem, True
    
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "TOTALNETMONTHLYCOST")
    xmlResponseElem.appendChild xmlNode
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "TOTALGROSSMONTHLYCOST")
    xmlResponseElem.appendChild xmlNode
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "TOTALACCRUEDINTEREST")
    xmlResponseElem.appendChild xmlNode
      
    'BMIDS766 GHun
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "AMOUNTPERUNITBORROWED")
    xmlResponseElem.appendChild xmlNode
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "APR")
    xmlResponseElem.appendChild xmlNode
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "TOTALAMOUNTPAYABLE")
    xmlResponseElem.appendChild xmlNode
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCalcsResponse, "TOTALMORTGAGEPAYMENTS")
    xmlResponseElem.appendChild xmlNode
    'BMIDS766 End
    
    Set IMortgageSubQuoteBO_RunCompositeRateCalcAndSaveResults = xmlResponseElem
    m_objContext.SetComplete
    
RunCompositeRateCalcAndSaveResultsExit:
    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCalcsResponse = Nothing
    Set xmlSaveResponse = Nothing
    Set xmlNode = Nothing
    Exit Function

RunCompositeRateCalcAndSaveResultsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_RunCompositeRateCalcAndSaveResults = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume RunCompositeRateCalcAndSaveResultsExit
End Function

'BMIDS736 GHun
Private Function ConvertAlphaPlusOutput(ByVal vXMLRequest As IXMLDOMElement, ByVal vXMLResponse As IXMLDOMElement, Optional ByVal vxmlAddBorrowResponse As IXMLDOMElement = Nothing) As IXMLDOMNode

    Const cstrFunctionName  As String = "ConvertAlphaPlusOutput"
    Dim xmlNode             As IXMLDOMNode
    Dim xmlTemp             As IXMLDOMNode
    Dim xslDoc              As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc   As FreeThreadedDOMDocument40
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim strXSLPath          As String
    
On Error GoTo ConvertAlphaPlusOutputErr

    Set xmlDoc = xmlCreateDOMObject
    Set xmlNode = xmlDoc.createElement("XML")
    xmlDoc.appendChild xmlNode
    xmlNode.appendChild vXMLRequest
    xmlNode.appendChild vXMLResponse
    'EP2_954 GHun
    If Not vxmlAddBorrowResponse Is Nothing Then
        Set xmlTemp = xmlDoc.createElement("ADDITIONALBORROWING")
        xmlTemp.appendChild vxmlAddBorrowResponse
        xmlNode.appendChild xmlTemp
    End If
    'EP2_954 End
    
    Set xslDoc = xmlCreateDOMObject
    strXSLPath = Replace(App.Path, "DLL", "XML")
    
    'Transform the Alpha+ response XML to Omiga format
    If Not xslDoc.load(strXSLPath & "\AlphaPlusResponse.xsl") Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSL File"
    End If
    
    Set xmlTransformedDoc = xmlCreateDOMObject
    xmlDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc

    If xmlTransformedDoc.parseError.errorCode <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaParseError, xmlTransformedDoc.parseError.reason
    End If

    Set ConvertAlphaPlusOutput = xmlTransformedDoc.documentElement
   
ConvertAlphaPlusOutputExit:
    Set xmlNode = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    Set xmlDoc = Nothing
    Exit Function
    
ConvertAlphaPlusOutputErr:
    m_objErrAssist.AddToErrSource (cstrFunctionName)

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    Err.Raise Err.number, Err.Source, Err.Description

End Function
'BMIDS736 End

'BMIDS736 GHun
'This method is based on CalculateInterestRate. It calculates the interest rate and
'adds 1% to the rate if possible. Fixed rates are not changed, and if the increase is
'above a cap, then it is ignored (only interested in full 1% increases).
Private Function CalculateIncreasedInterestRate(ByVal vdblIntRate As Double, ByVal vxmlInterestRateType As IXMLDOMNode) As Double
'-----------------------------------------------------------------------------------------
' description:  Add 1% to the passed in interest rate, if appropriate.
'
' pass:         vdblIntRate             The current interest rate
'               vxmlInterestRateType    XML containing a interest rate type data
'
' return:       The calculated interest rate
'------------------------------------------------------------------------------------------
On Error GoTo CalculateIncreasedInterestRateErr
    Const cstrFunctionName As String = "CalculateIncreasedInterestRate"
    Dim dblCappedRate As Double
    Dim dblIntRate As Double
    Dim strRateType As String
    Dim strTagValue As String
    Dim objConvAssist As ConvertAssist
    Set objConvAssist = New ConvertAssist
            
    strRateType = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "RATETYPE")
        
    Select Case strRateType
        Case "F"   'Fixed rates are not changed
            dblIntRate = vdblIntRate
    
        Case "B", "D"   'Base Rate and discounted
            dblIntRate = vdblIntRate + 1
                
        Case "C"   'Capped/Floored
            dblIntRate = vdblIntRate + 1
            
            strTagValue = m_objXmlAssist.GetTagValue(vxmlInterestRateType, "CEILINGRATE")
            If Len(Trim$(strTagValue)) > 0 Then
                dblCappedRate = objConvAssist.CSafeDbl(strTagValue)
                'If the 1% increase is above the cap then a full 1% cannot be applied, so reset the value
                If dblIntRate > dblCappedRate Then
                    dblIntRate = vdblIntRate
                End If
            End If
    End Select
    
    CalculateIncreasedInterestRate = dblIntRate
    
    Set objConvAssist = Nothing
    Exit Function
    
CalculateIncreasedInterestRateErr:
    Set objConvAssist = Nothing
    m_objErrAssist.AddToErrSource (cstrFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    Err.Raise Err.number, Err.Source, Err.Description
End Function
'BMIDS736 End

'BMIDS929 GHun
'Calculate the number of whole months difference between 2 dates (i.e. rounding down)
'The VB DataDiff function ignores the day of the month in the dates, so it sometimes
'rounds this value up, and other times rounds it down.
Private Function MonthDiff(ByVal vdteDate1 As Date, ByVal vdteDate2 As Date) As Long
    Dim lngMonths As Long
    
    lngMonths = DateDiff("m", vdteDate1, vdteDate2)
    'If the difference has been rounded up then decrement the number of months
    If DateDiff("d", DateAdd("m", lngMonths, vdteDate1), vdteDate2) < 0 Then
        lngMonths = lngMonths - 1
    End If
    
    MonthDiff = lngMonths
End Function
'BMIDS929 End

'MAR1777 GHun
Private Function CalcExpectedCompletionDate(ByVal vdteDate As Date) As Date
    Dim dteCompDate As Date
    Dim intDayOfMonth As Integer
    
    'Add one month
    dteCompDate = DateAdd("m", 1, vdteDate)
    
    'Set the day of the month to AssumedCompletionDate
    intDayOfMonth = GetGlobalParamAmount("AssumedCompletionDayOfMonth")
    If intDayOfMonth = 0 Then   'Default the value in case the globalParameter does not exist
        intDayOfMonth = 1
    End If
    
    dteCompDate = CSafeDate(Format$(intDayOfMonth, "00") & "/" & Format$(Month(dteCompDate), "00") & "/" & CStr(Year(dteCompDate)))
    
    CalcExpectedCompletionDate = dteCompDate
    
End Function
'MAR1777 End
' PSC 30/11/2006 EP2_249 - Start
Public Function SaveIntroducerFees(ByVal vstrXMLRequest As String) As String
On Error GoTo SaveIntroducerFeesVbErr

    Const cstrFunctionName As String = "SaveIntroducerFees"
          
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponse
    
    Set xmlIn = xmlLoad(vstrXMLRequest, cstrFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponse = IMortgageSubQuoteBO_SaveIntroducerFees(xmlIn.documentElement)
    errCheckXMLResponseNode xmlTempResponse, xmlResponse, True
    
    SaveIntroducerFees = xmlResponse.xml
    
    m_objContext.SetComplete

SaveIntroducerFeesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlTempResponse = Nothing

    Exit Function

SaveIntroducerFeesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponse
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    SaveIntroducerFees = m_objErrAssist.CreateErrorResponse()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume SaveIntroducerFeesExit
    
End Function

Private Function IMortgageSubQuoteBO_SaveIntroducerFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo SaveIntroducerFeesVbErr
    
    Const cstrFunctionName As String = "IMortgageSubQuoteBO_SaveIntroducerFees"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlOperation As IXMLDOMNode
    Dim xmlMortgageIntroducerFee As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim objCrudTxBO As omCRUDTxBO

    Dim strApplicationNumber As String
    Dim strMtgSubQuoteNumber As String
    Dim strSubmissionRoute As String
    Dim strAssociationFee As String
    Dim strPackagingFee As String
    Dim strResponse As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponse
    
    Set xmlMortgageIntroducerFee = xmlGetMandatoryNode(vXMLRequest, cstrMTGINTRODUCERFEE_NODE_NAME)
    strApplicationNumber = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "APPLICATIONNUMBER")
    strMtgSubQuoteNumber = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "MORTGAGESUBQUOTENUMBER")
    strSubmissionRoute = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "SUBMISSIONROUTE")
    
    Set objCrudTxBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDTxBO")

    ' Set up crud request
    Set xmlCrudRequest = vXMLRequest.cloneNode(False)
    Set xmlOperation = xmlOut.createElement("OPERATION")
    xmlSetAttributeValue xmlOperation, "CRUD_OP", "CREATE"
    xmlSetAttributeValue xmlOperation, "ENTITY_REF", "MORTGAGEINTRODUCERFEE"
    
    ' Add insert for proc fee
    Set xmlRoot = xmlOut.createElement("MORTGAGEINTRODUCERFEE")
    xmlSetAttributeValue xmlRoot, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlRoot, "MORTGAGESUBQUOTENUMBER", strMtgSubQuoteNumber
    xmlSetAttributeValue xmlRoot, "INTRODUCERFEETYPE", GetFirstComboValueId("IntroducerFeeType", "PR")
    xmlSetAttributeValue xmlRoot, "FEEAMOUNT", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "PROCFEE")
    ' PSC 05/02/2007 EP2_1219 - Start
    xmlSetAttributeValue xmlRoot, "RECIPIENTTYPE", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "RECIPIENTTYPE")
    xmlSetAttributeValue xmlRoot, "RECIPIENTID", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "RECIPIENTID")
    ' PSC 05/02/2007 EP2_1219 - End
    xmlOperation.appendChild xmlRoot
    xmlCrudRequest.appendChild xmlOperation

    ' Add insert for association fee if present
    strAssociationFee = xmlGetNodeText(xmlMortgageIntroducerFee, "ASSOCIATIONFEE")
    
    If CSafeLng(strAssociationFee) > 0 Then
    
        Set xmlOperation = xmlOut.createElement("OPERATION")
        xmlSetAttributeValue xmlOperation, "CRUD_OP", "CREATE"
        xmlSetAttributeValue xmlOperation, "ENTITY_REF", "MORTGAGEINTRODUCERFEE"
        
        Set xmlRoot = xmlOut.createElement("MORTGAGEINTRODUCERFEE")
        xmlSetAttributeValue xmlRoot, "APPLICATIONNUMBER", strApplicationNumber
        xmlSetAttributeValue xmlRoot, "MORTGAGESUBQUOTENUMBER", strMtgSubQuoteNumber
        xmlSetAttributeValue xmlRoot, "INTRODUCERFEETYPE", GetFirstComboValueId("IntroducerFeeType", "A")
        xmlSetAttributeValue xmlRoot, "FEEAMOUNT", strAssociationFee
        xmlSetAttributeValue xmlRoot, "RECIPIENTTYPE", strSubmissionRoute
        
        If IsValidationType("ProcFeeSubmissionRoute", CSafeInt(strSubmissionRoute), "PA") Then
            xmlSetAttributeValue xmlRoot, "RECIPIENTID", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "PACKAGERASSOCIATIONID")
        ElseIf IsValidationType("ProcFeeSubmissionRoute", CSafeInt(strSubmissionRoute), "MC") Then
            xmlSetAttributeValue xmlRoot, "RECIPIENTID", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "MORTGAGECLUBID")
        Else
            errThrowError cstrFunctionName, oeInvalidParameter, "Submission Route " & strSubmissionRoute & " is invalid for association fee"
        End If
        
        xmlOperation.appendChild xmlRoot
        xmlCrudRequest.appendChild xmlOperation
    
    End If
    
    ' Add insert for packaging fee if present
    strPackagingFee = xmlGetNodeText(xmlMortgageIntroducerFee, "PACKAGINGFEE")

    If CSafeLng(strPackagingFee) > 0 Then
        
        Set xmlOperation = xmlOut.createElement("OPERATION")
        xmlSetAttributeValue xmlOperation, "CRUD_OP", "CREATE"
        xmlSetAttributeValue xmlOperation, "ENTITY_REF", "MORTGAGEINTRODUCERFEE"
        
        Set xmlRoot = xmlOut.createElement("MORTGAGEINTRODUCERFEE")
        xmlSetAttributeValue xmlRoot, "APPLICATIONNUMBER", strApplicationNumber
        xmlSetAttributeValue xmlRoot, "MORTGAGESUBQUOTENUMBER", strMtgSubQuoteNumber
        xmlSetAttributeValue xmlRoot, "INTRODUCERFEETYPE", GetFirstComboValueId("IntroducerFeeType", "PA")
        xmlSetAttributeValue xmlRoot, "FEEAMOUNT", strPackagingFee
        
        ' PSC 13/02/2007 EP2_1288 - Start
        xmlSetAttributeValue xmlRoot, "RECIPIENTTYPE", GetFirstComboValueId("ProcFeeSubmissionRoute", "P")
        xmlSetAttributeValue xmlRoot, "RECIPIENTID", xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "PACKAGERID")
        ' PSC 13/02/2007 EP2_1288 - End
        
        xmlOperation.appendChild xmlRoot
        xmlCrudRequest.appendChild xmlOperation
    
    End If
    
    strResponse = objCrudTxBO.OmRequest(xmlCrudRequest.xml)
    errCheckXMLResponse strResponse, True, xmlResponse
    
    Set IMortgageSubQuoteBO_SaveIntroducerFees = xmlResponse
    
    m_objContext.SetComplete
    
SaveIntroducerFeesExit:

    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlOperation = Nothing
    Set xmlMortgageIntroducerFee = Nothing
    Set xmlTempNode = Nothing
    Set objCrudTxBO = Nothing

    Exit Function

SaveIntroducerFeesVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponse
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If errIsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_SaveIntroducerFees = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume SaveIntroducerFeesExit

End Function
' PSC 30/11/2006 EP2_249 - End

' PSC 15/03/2007 EP2_1956 - Start
Public Function CopyIntroducerFees(ByVal vstrXMLRequest As String) As String
On Error GoTo CopyIntroducerFeesVbErr

    Const cstrFunctionName As String = "CopyIntroducerFees"
          
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponse
    
    Set xmlIn = xmlLoad(vstrXMLRequest, cstrFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponse = IMortgageSubQuoteBO_CopyIntroducerFees(xmlIn.documentElement)
    errCheckXMLResponseNode xmlTempResponse, xmlResponse, True
    
    CopyIntroducerFees = xmlResponse.xml
    
    m_objContext.SetComplete

CopyIntroducerFeesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlTempResponse = Nothing

    Exit Function

CopyIntroducerFeesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponse
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    CopyIntroducerFees = m_objErrAssist.CreateErrorResponse()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CopyIntroducerFeesExit
    
End Function

Private Function IMortgageSubQuoteBO_CopyIntroducerFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo CopyIntroducerFeesVbErr
    
    Const cstrFunctionName As String = "IMortgageSubQuoteBO_CopyIntroducerFees"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCrudResponse As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
    Dim xmlOperation As IXMLDOMNode
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlMortgageIntroducerFee As IXMLDOMNode
    Dim xmlMortgageIntroducerFees As IXMLDOMNodeList
    Dim objCrudTxBO As omCRUDTxBO
    Dim objCrudBO As omCRUDBO

    Dim strResponse As String
    Dim strApplicationNumber As String
    Dim strMtgSubQuoteNumber As String
    Dim strNewMtgSubQuoteNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponse
    
    Set xmlMortgageIntroducerFee = xmlGetMandatoryNode(vXMLRequest, cstrMTGINTRODUCERFEE_NODE_NAME)
    strApplicationNumber = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "APPLICATIONNUMBER")
    strMtgSubQuoteNumber = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "MORTGAGESUBQUOTENUMBER")
    strNewMtgSubQuoteNumber = xmlGetMandatoryNodeText(xmlMortgageIntroducerFee, "NEWMORTGAGESUBQUOTENUMBER")
    
    Set objCrudBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDBO")
    Set objCrudTxBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDTxBO")

    ' Set up crud read request
    Set xmlCrudRequest = vXMLRequest.cloneNode(False)
    xmlSetAttributeValue xmlCrudRequest, "CRUD_OP", "READ"
    xmlSetAttributeValue xmlCrudRequest, "ENTITY_REF", "MORTGAGEINTRODUCERFEE"
    Set xmlRoot = xmlOut.createElement("MORTGAGEINTRODUCERFEE")
    xmlSetAttributeValue xmlRoot, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlRoot, "MORTGAGESUBQUOTENUMBER", strMtgSubQuoteNumber
    xmlCrudRequest.appendChild xmlRoot
    
    strResponse = objCrudBO.OmRequest(xmlCrudRequest.xml)
    Set xmlCrudResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlCrudResponse.documentElement, xmlResponse, True
    
    Set xmlMortgageIntroducerFees = xmlCrudResponse.selectNodes("RESPONSE/MORTGAGEINTRODUCERFEE")
    
    If xmlMortgageIntroducerFees.length > 0 Then
    
        ' Set up crud create request to copy fees
        Set xmlCrudRequest = vXMLRequest.cloneNode(False)
        
        For Each xmlMortgageIntroducerFee In xmlMortgageIntroducerFees
            Set xmlOperation = xmlOut.createElement("OPERATION")
            xmlSetAttributeValue xmlOperation, "CRUD_OP", "CREATE"
            xmlSetAttributeValue xmlOperation, "ENTITY_REF", "MORTGAGEINTRODUCERFEE"
            xmlCrudRequest.appendChild xmlOperation

            xmlOperation.appendChild xmlMortgageIntroducerFee
            xmlSetAttributeValue xmlMortgageIntroducerFee, "MORTGAGESUBQUOTENUMBER", strNewMtgSubQuoteNumber
        Next
        
        strResponse = objCrudTxBO.OmRequest(xmlCrudRequest.xml)
        errCheckXMLResponse strResponse, True, xmlResponse
    End If

    Set IMortgageSubQuoteBO_CopyIntroducerFees = xmlResponse
    
    m_objContext.SetComplete
    
CopyIntroducerFeesExit:

    Set xmlOut = Nothing
    Set xmlCrudResponse = Nothing
    Set xmlResponse = Nothing
    Set xmlOperation = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlMortgageIntroducerFee = Nothing
    Set xmlMortgageIntroducerFees = Nothing
    Set objCrudTxBO = Nothing
    Set objCrudBO = Nothing

    Exit Function

CopyIntroducerFeesVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponse
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If errIsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    Set IMortgageSubQuoteBO_CopyIntroducerFees = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CopyIntroducerFeesExit

End Function
' PSC 15/03/2007 EP2_1956 - End



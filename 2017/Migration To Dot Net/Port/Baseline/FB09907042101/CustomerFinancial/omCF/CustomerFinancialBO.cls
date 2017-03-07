VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerFinancialBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerFinancialBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for CustomerFinancial Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (CustomerFinancialTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. CustomerFinancialTxBO, CustomerFinancialDO
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
'SR     18/02/00    Modified method - Validate Mortgage Loan - AQR - SYS0235
'SR     28/04/00    New method 'FindOutgoingsTotalList' with string interface. One with
'                   DOM interace already exists.
'SR     28/04/00    Passing the entire request to DO method so that it gets data for
'                   for multiple input records in the method
'                   ICustomerFinancialBO_FindMortgageLoanList
'IW     08/05/00    SYS0145 - Fixed anomolous error message.
'APS    01/06/00    SYS0800 - Handle Record Not Found processing from FindRegularOutgoingsListDO call
'MS     07/06/00    SYS0811 - Automation error Record not found is NOT an error
'MC     08/06/00    SYS0688 - Validate Ported Loan when Adding a loan.
'BG     14/06/00    Added SaveMortgageAccount Method which calls either CustomerFinancialBO.Create/Delete/UpdateMortgageAccount
'SR     23/06/00    Modified method ValidateLiability - Call SetComplete on success
'                   New method SaveLiability
'PSC    29/06/00    SYS1012 - Add new method GetLoansNotRedeemedTotals
'BG     SYS1626     Changed FindMortgageListForArrears code to look for Tag MORTGAGEACCOUNTLIST rather than MORTGAGEACCOUNT
'                   as previously.
'PSC   29/01/01     SYS1888 SetComplete on Find and Get Methods rather than SetAbort
'GD     14/03/01    SYS2047 Added GetMortgageAccount
'DRC    22/03/01    SYS1932 Added GetAccountDetails method
'DRC    14/06/01    SYS0999 Added ValidatePortedLoan method & finished off ValidateMortgageLoan method
'DRC    04/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'DS     09/04/02    Fix to create object in context.
'-------------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        AQR         Description
'PSC    17/07/02    BMIDS00006  Added CreateAccountRelationships
'                                    DeleteAccountRelationships
'                                    UpdateAccountRelationships
'                                    CreateSpecialFeature
'                                    UpdateSpecialFeature
'                                    DeleteSpecialFeature
'                                    FindSpecialFeatureList
'                                    GetMortgageAccountDetails
'MV     14/08/2002  BMIDS00311  Core Ref: SYS1705 ; Modified ICustomerFinancialBO_SaveMortgageAccountAndLoan
'GHun   22/08/2002  BMIDS00190  DCWP3 BM076 Added CalculateCreditCardRepayment,
'                                                 FindLoansListForArrears
'MV     02/09/2002  BMIDS00353  CMWP4 - BM087 ; Added GetMtgAcctDataForFurtherAdv
'MDC    10/09/2002  BMIDS00336  CCWP1 BM062 Credit Check & Bureau Download
'MDC    08/10/2002  BMIDS00561  Return AccountGuid generated in CreateLiability
'GHun   14/11/2002  BMIDS00444  Add FindRemortgageAccountAddress
'SR     25/05/2004  BMIDS772    New methods GetCreditHistorySummaryView and ICustomerFinancialBO_GetCreditHistorySummaryView
'                               GetFinancialSummaryView and ICustomerFinancialBO_GetFinancialSummaryView
'SR     26/05/2004  BMIDS772    Modified methods ICustomerFinancialBO_GetFinancialSummaryView,
'                                                ICustomerFinancialBO_GetCreditHistorySummaryView
'GHun   08/09/2004  BMIDS815    Changed GetPortedStepAndPeriodFromMortgageAccount
'------------------------------------------------------------------------------------------------------------------
'Epsom History:
'
'Prog   Date        Description
'GHun   05/02/2007  EP2_954 Added new function GetAdditionalBorrowingDataForProjectionCalc
'------------------------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerFinancialBO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Private Const cstrFEATUREROOT_NODE_NAME As String = "MORTGAGEACCOUNTSPECIALFEATURE"
Private Const cstrACCOUNT_ROOT_NODE_NAME As String = "ACCOUNT"
Private Const cstrMORTACCOUNT_ROOT_NODE_NAME As String = "MORTGAGEACCOUNT"
Private Const cstrCUSTOMER_ROOT_NODE_NAME As String = "CUSTOMER"

Public Function GetMtgAcctDataForFurtherAdvance(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo GetMtgAcctDataForFurtherAdvanceVbErr

    Const strFunctionName As String = "GetMtgAcctDataForFurtherAdvance"
        
        
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
    Set xmlTempResponseNode = ICustomerFinancialBO_GetMtgAcctDataForFurtherAdv(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMtgAcctDataForFurtherAdvance = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetMtgAcctDataForFurtherAdvanceExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetMtgAcctDataForFurtherAdvanceVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMtgAcctDataForFurtherAdvance = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetMtgAcctDataForFurtherAdvanceExit

End Function

' SR 25/05/2004 : BMIDS772 - new method
Private Function ICustomerFinancialBO_GetCreditHistorySummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetCreditHistorySummaryViewVbErr
    
    Const cstrFunctionName As String = "ICustomerFinancialBO_GetCreditHistorySummaryView"
   
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode
    
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName <> cstrCUSTOMER_ROOT_NODE_NAME Then
        Set xmlNodeList = vxmlRequest.selectNodes("//" & cstrCUSTOMER_ROOT_NODE_NAME)
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
                oeMissingPrimaryTag, cstrCUSTOMER_ROOT_NODE_NAME & " tag not found"
        Else
            For Each xmlNode In xmlNodeList
                If m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER") = "" Or _
                   m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERVERSIONNUMBER") = "" Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 256
                End If
            Next xmlNode
        End If
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetCreditHistorySummaryView(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetCreditHistorySummaryView = xmlResponseElem
    
    m_objContext.SetComplete
    
GetCreditHistorySummaryViewExit:
    
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    Exit Function

GetCreditHistorySummaryViewVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetCreditHistorySummaryView = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetCreditHistorySummaryViewExit


End Function

'' SR 25/05/2004 : BMIDS772 - new method
Private Function ICustomerFinancialBO_GetFinancialSummaryView(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetFinancialSummaryViewVbErr
    
    Const cstrFunctionName As String = "ICustomerFinancialBO_GetFinancialSummaryView"
   
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode
    
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName <> cstrCUSTOMER_ROOT_NODE_NAME Then
        Set xmlNodeList = vxmlRequest.selectNodes("//" & cstrCUSTOMER_ROOT_NODE_NAME)
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
                oeMissingPrimaryTag, cstrCUSTOMER_ROOT_NODE_NAME & " tag not found"
        Else
            For Each xmlNode In xmlNodeList
                If m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER") = "" Or _
                   m_objXmlAssist.GetTagValue(xmlNode, "CUSTOMERVERSIONNUMBER") = "" Then
                    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 256
                End If
            Next xmlNode
        End If
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetFinancialSummaryView(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetFinancialSummaryView = xmlResponseElem
    
    m_objContext.SetComplete
    
GetFinancialSummaryViewExit:
    
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    Exit Function

GetFinancialSummaryViewVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetFinancialSummaryView = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetFinancialSummaryViewExit

End Function

'BMIDS00336 MDC 10/09/2002
Private Function ICustomerFinancialBO_GetMortgageAccountLoanTotals(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetMortgageAccountLoanTotalsVbErr
    
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim objICustomerFinancialDO As ICustomerFinancialDO

Dim strNodeName As String

Const strFunctionName As String = "ICustomerFinancialBO_GetMortgageAccountLoanTotals"


    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    strNodeName = "MORTGAGEACCOUNTLOANTOTALS"

    If vxmlTableElement.nodeName = strNodeName Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName(strNodeName).Item(0)
    End If
        
    If vxmlTableElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strNodeName & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetMortgageAccountLoanTotals(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetMortgageAccountLoanTotals = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetMortgageAccountLoanTotalsExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetMortgageAccountLoanTotalsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetMortgageAccountLoanTotals = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageAccountLoanTotalsExit

End Function
'BMIDS00336 MDC 10/09/2002 - End

Private Function ICustomerFinancialBO_GetMtgAcctDataForFurtherAdv(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetMtgAcctDataForFurtherAdvVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetMtgAcctDataForFurtherAdv"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName = cstrMORTACCOUNT_ROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrMORTACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrMORTACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetMtgAcctDataForFurtherAdv(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetMtgAcctDataForFurtherAdv = xmlResponseElem
    
    m_objContext.SetComplete
    
GetMtgAcctDataForFurtherAdvExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetMtgAcctDataForFurtherAdvVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetMtgAcctDataForFurtherAdv = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetMtgAcctDataForFurtherAdvExit
End Function


Private Function ICustomerFinancialBO_GetMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Gets the total of outstanding balances and the total or repayments being
'               made on loans that are not to be redeemed
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetMortgageAccount"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    Dim strNodeName As String
'    strNodeName = "CUSTOMERLIST"
'
'    If vxmlRequest.nodeName = strNodeName Then
        Set xmlRequestNode = vxmlRequest
'    Else
'        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strNodeName).Item(0)
'    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strNodeName & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetMortgageAccount(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetMortgageAccount = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetMortgageAccountExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetMortgageAccount = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageAccountExit

End Function

Private Function ICustomerFinancialBO_CreateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateArrearsHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateArrearsHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateArrearsHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateArrearsHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateArrearsHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateArrearsHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateArrearsHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArrearsHistoryExit

End Function

Private Function ICustomerFinancialBO_CreateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankCard a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankCardVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateBankCard"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateBankCard function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateBankCard(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateBankCard = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankCardExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankCardExit

End Function

Private Function ICustomerFinancialBO_CreateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankruptcyHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankruptcyHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateBankruptcyHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateBankruptcyHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateBankruptcyHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateBankruptcyHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankruptcyHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialBO_CreateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateCCJHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateCCJHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateCCJHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateCCJHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateCCJHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateCCJHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCCJHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateCCJHistoryExit

End Function

Private Function ICustomerFinancialBO_CreateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateDeclinedMortgage a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclinedMortgageVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateDeclinedMortgage"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateDeclinedMortgage function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateDeclinedMortgage(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateDeclinedMortgage = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateDeclinedMortgageExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateDeclinedMortgageExit

End Function

Private Function ICustomerFinancialBO_CreateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateGroupConnection a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateGroupConnectionVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateGroupConnection"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateGroupConnection function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateGroupConnection(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateGroupConnection = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateGroupConnectionExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateGroupConnectionExit

End Function

Private Function ICustomerFinancialBO_CreateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLiability a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLiabilityVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateLiability"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateLiability function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateLiability(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    'BMIDS00561 MDC 08/10/2002
'    Set ICustomerFinancialBO_CreateLiability = xmlResponseElem.cloneNode(True)
    Set ICustomerFinancialBO_CreateLiability = xmlTempResponseNode.cloneNode(True)
    'BMIDS00561 MDC 08/10/2002 - End

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLiabilityExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLiabilityExit

End Function

Private Function ICustomerFinancialBO_CreateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLifeProduct a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLifeProductVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateLifeProduct"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateLifeProduct function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateLifeProduct(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateLifeProduct = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLifeProductExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLifeProductExit

End Function

Private Function ICustomerFinancialBO_CreateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageAccount a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateMortgageAccount"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlKeysNode = xmlOut.createElement("MORTGAGEACCOUNT")
    xmlResponseElem.appendChild xmlKeysNode
       
    '   call Business Transaction Object CreateMortgageAccount function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateMortgageAccount(vxmlRequest, xmlKeysNode)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    
    Set ICustomerFinancialBO_CreateMortgageAccount = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set ICustomerFinancialBO_CreateMortgageAccount = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerFinancialBO_CreateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLoan a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLoanVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateMortgageLoan"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
       
    Set xmlKeysNode = xmlOut.createElement("MORTGAGELOAN")
    xmlResponseElem.appendChild xmlKeysNode
    
    '   call Business Transaction Object CreateMortgageLoan function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateMortgageLoan(vxmlRequest, xmlKeysNode)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateMortgageLoan = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateMortgageLoanExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageLoanExit

End Function

Private Function ICustomerFinancialBO_CreateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateRegularOutgoings a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateRegularOutgoingsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateRegularOutgoings"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    '   call Business Transaction Object CreateRegularOutgoings function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateRegularOutgoings(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    Set ICustomerFinancialBO_CreateRegularOutgoings = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set ICustomerFinancialBO_CreateRegularOutgoings = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerFinancialBO_DeleteArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteArrearsHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArrearsHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteArrearsHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteArrearsHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteArrearsHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteArrearsHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteArrearsHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArrearsHistoryExit

End Function

Private Function ICustomerFinancialBO_DeleteBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankCard a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankCardVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteBankCard"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteBankCard function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteBankCard(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteBankCard = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankCardExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankCardExit

End Function

Private Function ICustomerFinancialBO_DeleteBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankruptcyHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankruptcyHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteBankruptcyHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteBankruptcyHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteBankruptcyHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteBankruptcyHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankruptcyHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialBO_DeleteCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteCCJHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCCJHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteCCJHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteCCJHistory function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteCCJHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteCCJHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteCCJHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCCJHistoryExit

End Function

Private Function ICustomerFinancialBO_DeleteDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteDeclinedMortgage a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDeclinedMortgageVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteDeclinedMortgage"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteDeclinedMortgage function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteDeclinedMortgage(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteDeclinedMortgage = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteDeclinedMortgageExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteDeclinedMortgageExit

End Function

Private Function ICustomerFinancialBO_DeleteGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteGroupConnection a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteGroupConnectionVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteGroupConnection"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteGroupConnection function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteGroupConnection(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteGroupConnection = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteGroupConnectionExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteGroupConnectionExit

End Function

Private Function ICustomerFinancialBO_DeleteLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLiability a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLiabilityVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteLiability"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteLiability function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteLiability(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteLiability = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLiabilityExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLiabilityExit

End Function

Private Function ICustomerFinancialBO_DeleteLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLifeProduct a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeProductVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteLifeProduct"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteLifeProduct function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteLifeProduct(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteLifeProduct = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLifeProductExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLifeProductExit

End Function

Private Function ICustomerFinancialBO_DeleteMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageAccount a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteMortgageAccount"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteMortgageAccount function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteMortgageAccount(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteMortgageAccount = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteMortgageAccountExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageAccountExit

End Function

Private Function ICustomerFinancialBO_DeleteMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageLoan a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageLoanVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteMortgageLoan"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteMortgageLoan function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteMortgageLoan(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteMortgageLoan = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteMortgageLoanExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageLoanExit

End Function

Private Function ICustomerFinancialBO_DeleteRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageLoan a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteRegularOutgoingsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteRegularOutgoings"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
           
    '   call Business Transaction Object DeleteRegularOutgoings function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteRegularOutgoings(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteRegularOutgoings = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteRegularOutgoingsExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteRegularOutgoingsExit

End Function

Private Function ICustomerFinancialBO_FindArrearsHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindArrearsHistoryListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindArrearsHistoryList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "ARREARSHISTORYLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("ARREARSHISTORYLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ARREARSHISTORYLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindArrearsHistoryList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindArrearsHistoryList = xmlResponseElem.cloneNode(True)

FindArrearsHistoryListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindArrearsHistoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindArrearsHistoryList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindArrearsHistoryListExit

End Function

Private Function ICustomerFinancialBO_FindBankCardList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindBankCardListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindBankCardList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "BANKCREDITCARDLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("BANKCREDITCARDLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BANKCREDITCARDLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindBankCardList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindBankCardList = xmlResponseElem.cloneNode(True)

FindBankCardListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindBankCardListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindBankCardList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankCardListExit

End Function

Private Function ICustomerFinancialBO_FindBankruptcyHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindBankruptcyHistoryListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindBankruptcyHistoryList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    'MS SYS0811 Automation error
    Dim blnResumeNext As Boolean
    blnResumeNext = False
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "BANKRUPTCYHISTORYLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("BANKRUPTCYHISTORYLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BANKRUPTCYHISTORYLIST tag not found"
    End If
    
    'MS SYS0811 Automation error
    'Record Not found is NOT an error
    blnResumeNext = True
    Set xmlDataNode = objICustomerFinancialDO.FindBankruptcyHistoryList(xmlRequestNode)
    blnResumeNext = False
    'Record Not found is NOT an error
    'No data - just continue
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindBankruptcyHistoryList = xmlResponseElem.cloneNode(True)

FindBankruptcyHistoryListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindBankruptcyHistoryListVbErr:

    'MS SYS0811 Automation error
    'Record not found is NOT an error
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        If blnResumeNext Then
            Resume Next
        End If
    End If
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindBankruptcyHistoryList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankruptcyHistoryListExit

End Function

Private Function ICustomerFinancialBO_FindCCJHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindCCJHistoryListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindCCJHistoryList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "CCJHISTORYLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CCJHISTORYLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CCJHISTORYLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindCCJHistoryList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindCCJHistoryList = xmlResponseElem.cloneNode(True)

FindCCJHistoryListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindCCJHistoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindCCJHistoryList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindCCJHistoryListExit

End Function

Private Function ICustomerFinancialBO_FindDeclinedMortgageList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindDeclinedMortgageListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindDeclinedMortgageList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "DECLINEDMORTGAGELIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("DECLINEDMORTGAGELIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "DECLINEDMORTGAGELIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindDeclinedMortgageList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindDeclinedMortgageList = xmlResponseElem.cloneNode(True)

FindDeclinedMortgageListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindDeclinedMortgageListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindDeclinedMortgageList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindDeclinedMortgageListExit

End Function

Private Function ICustomerFinancialBO_FindGroupConnectionList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindGroupConnectionListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindGroupConnectionList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "GROUPCONNECTIONLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("GROUPCONNECTIONLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "GROUPCONNECTIONLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindGroupConnectionList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindGroupConnectionList = xmlResponseElem.cloneNode(True)

FindGroupConnectionListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindGroupConnectionListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindGroupConnectionList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindGroupConnectionListExit

End Function

Private Function ICustomerFinancialBO_FindLiabilityList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLiabilityListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindLiabilityList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "LOANSLIABILITIESLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("LOANSLIABILITIESLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "LOANSLIABILITIESLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindLiabilityList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindLiabilityList = xmlResponseElem.cloneNode(True)

FindLiabilityListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindLiabilityListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindLiabilityList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLiabilityListExit

End Function

Private Function ICustomerFinancialBO_FindLifeProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLifeProductListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindLifeProductList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "LIFEPRODUCTLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("LIFEPRODUCTLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "LIFEPRODUCTLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindLifeProductList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindLifeProductList = xmlResponseElem.cloneNode(True)

FindLifeProductListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindLifeProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindLifeProductList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLifeProductListExit

End Function

Private Function ICustomerFinancialBO_FindLoansToBePorted(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLoansToBePortedVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindLoansToBePorted"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "CUSTOMERLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CUSTOMERLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindLoansToBePorted(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindLoansToBePorted = xmlResponseElem.cloneNode(True)

FindLoansToBePortedExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindLoansToBePortedVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindLoansToBePorted = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLoansToBePortedExit

End Function

Private Function ICustomerFinancialBO_FindMortgageAccountAndLoanList( _
                        ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindMortgageAccountAndLoanList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
       
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlTableElement.nodeName = "MORTGAGEACCOUNTANDLOANLIST" Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("MORTGAGEACCOUNTANDLOANLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEACCOUNTANDLOANLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindMortgageAccountAndLoanList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerFinancialBO_FindMortgageAccountAndLoanList = xmlResponseElem.cloneNode(True)

FindListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindMortgageAccountAndLoanList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function ICustomerFinancialBO_FindMortgageAccountList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageAccountListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindMortgageAccountList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "MORTGAGEACCOUNTLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNTLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEACCOUNTLIST tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindMortgageAccountList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindMortgageAccountList = xmlResponseElem.cloneNode(True)

FindMortgageAccountListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindMortgageAccountListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindMortgageAccountList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete          ' PSC 29/01/01 SYS1888
    End If
    
    ' go to clean-up section
    Resume FindMortgageAccountListExit

End Function

Private Function ICustomerFinancialBO_FindRegularOutgoingsList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindRegularOutgoingsListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindRegularOutgoingsList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "REGULAROUTGOINGSLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("REGULAROUTGOINGSLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "REGULAROUTGOINGSLIST tag not found"
    End If
        
    ' APS 01/06/00 SYS0800 - Handle Record Not Found processing from FindRegularOutgoingsListDO call
    Set xmlDataNode = objICustomerFinancialDO.FindRegularOutgoingsList(xmlRequestNode)
    
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindRegularOutgoingsList = xmlResponseElem.cloneNode(True)

FindRegularOutgoingsListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindRegularOutgoingsListVbErr:

    ' APS 01/06/00 SYS0800 - Handle Record Not Found processing from FindRegularOutgoingsListDO call
    If (objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound) Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindRegularOutgoingsList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindRegularOutgoingsListExit

End Function

Private Function ICustomerFinancialBO_FindMortgageListForArrears(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageListForArrearsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindMortgageListForArrears"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    'BG SYS1626 Changed code to look for Tag MORTGAGEACCOUNTLIST rather than MORTGAGEACCOUNT
    'as previously.
    If vxmlRequest.nodeName = "MORTGAGEACCOUNTLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNTLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEACCOUNTLIST" & " tag not found"
    End If
    'BG SYS1626 End
        
    Set xmlDataNode = objICustomerFinancialDO.FindMortgageListForArrears(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindMortgageListForArrears = xmlResponseElem.cloneNode(True)

FindMortgageListForArrearsExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindMortgageListForArrearsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindMortgageListForArrears = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindMortgageListForArrearsExit

End Function

Private Function ICustomerFinancialBO_FindMortgageLoanList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageLoanListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindMortgageLoanList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "MORTGAGELOAN" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGELOAN").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGELOAN" & " tag not found"
    End If
    'SR 27/04/00 - Pass vxmlRequest to DO rather than xmlRequestNode
    Set xmlDataNode = objICustomerFinancialDO.FindMortgageLoanList(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindMortgageLoanList = xmlResponseElem.cloneNode(True)

FindMortgageLoanListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindMortgageLoanListVbErr:
    objErrAssist.SaveErr
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindMortgageLoanList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindMortgageLoanListExit

End Function

Private Function ICustomerFinancialBO_FindOutgoingsTotalList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindOutgoingsTotalListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindOutgoingsTotalList"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "CUSTOMERLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CUSTOMERLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindOutgoingsTotalList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerFinancialBO_FindOutgoingsTotalList = xmlResponseElem.cloneNode(True)

FindOutgoingsTotalListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindOutgoingsTotalListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindOutgoingsTotalList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindOutgoingsTotalListExit
End Function

Private Function ICustomerFinancialBO_GetAccountDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountDetailsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetAccountDetails"


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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
                
    If vxmlRequest.nodeName = "ACCOUNT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("ACCOUNT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ACCOUNT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetAccountDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    Set ICustomerFinancialBO_GetAccountDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetAccountDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetAccountDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetAccountDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAccountDetailsExit

End Function

Private Function ICustomerFinancialBO_GetLoansNotRedeemedTotals(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the total of outstanding balances and the total or repayments being
'               made on loans that are not to be redeemed
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLoansNotRedeemedTotalsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetLoansNotRedeemedTotals"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    Dim strNodeName As String
    strNodeName = "CUSTOMERLIST"
                
    If vxmlRequest.nodeName = strNodeName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strNodeName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strNodeName & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetLoansNotRedeemedTotals(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetLoansNotRedeemedTotals = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetLoansNotRedeemedTotalsExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetLoansNotRedeemedTotalsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetLoansNotRedeemedTotals = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoansNotRedeemedTotalsExit

End Function

Private Function ICustomerFinancialBO_GetPortedStepAndPeriodFromMortgageAccount( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetPortedStepAndPeriodFromMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetPortedStepAndPeriodFromMortgageAccount"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlChild As IXMLDOMNode
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    Const strNodeName As String = "APPLICATION"

    If vxmlRequest.nodeName = strNodeName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strNodeName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strNodeName & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetPortedStepAndPeriodFromMortgageAccount(xmlRequestNode)
    'BMIDS815 GHun Append each child node to the response
    For Each xmlChild In xmlDataNode.childNodes
        xmlResponseElem.appendChild xmlChild
    Next
        
    Set ICustomerFinancialBO_GetPortedStepAndPeriodFromMortgageAccount = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetPortedStepAndPeriodFromMortgageAccountExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlChild = Nothing

    Exit Function

GetPortedStepAndPeriodFromMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetPortedStepAndPeriodFromMortgageAccount = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPortedStepAndPeriodFromMortgageAccountExit
End Function

Private Function ICustomerFinancialBO_SaveGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveGroupConnection a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveGroupConnectionVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_SaveGroupConnection"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIClassDef As ICFClassDef
    Dim botOperation As BOOPERATIONTYPE

    Set objIClassDef = New CFClassDef

    botOperation = objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadGroupConnectionData)

    Select Case botOperation
        Case booCreate: Set ICustomerFinancialBO_SaveGroupConnection = ICustomerFinancialBO_CreateGroupConnection(vxmlRequest)
        Case booUpdate: Set ICustomerFinancialBO_SaveGroupConnection = ICustomerFinancialBO_UpdateGroupConnection(vxmlRequest)
        Case booDelete: Set ICustomerFinancialBO_SaveGroupConnection = ICustomerFinancialBO_DeleteGroupConnection(vxmlRequest)
        Case Else
            ' N/A
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    
    Exit Function

SaveGroupConnectionVbErr:
    Set objIClassDef = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set ICustomerFinancialBO_SaveGroupConnection = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerFinancialBO_SaveLiability( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves multiple instances of the persistant data associated with the
'               table 'LoansLiabilities'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveLiabilityVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_SaveLiability"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objICFTxBo As ICustomerFinancialTxBO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICFTxBo = New CustomerFinancialTxBO
    Else
        Set objICFTxBo = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBo")
    End If
    
    Set xmlTempResponseNode = objICFTxBo.SaveLiability(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set ICustomerFinancialBO_SaveLiability = xmlResponseElem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveLiabilityExit:
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOut = Nothing
    
    Exit Function

SaveLiabilityVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_SaveLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLiabilityExit
    
End Function

Private Function ICustomerFinancialBO_SaveMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:   a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo SaveMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_SaveMortgageAccount"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    Dim xmlGeneratedKeys As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
        
    Set xmlGeneratedKeys = xmlOut.createElement("GENERATEDKEYS")
    xmlResponseElem.appendChild xmlGeneratedKeys
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.SaveMortgageAccount(vxmlRequest, xmlGeneratedKeys)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_SaveMortgageAccount = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveMortgageAccountExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_SaveMortgageAccount = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountExit

End Function


Private Function ICustomerFinancialBO_SaveMortgageAccountAndLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
    On Error GoTo SaveMortgageAccountAndLoanVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_SaveMortgageAccountAndLoan"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    Dim xmlGeneratedKeys As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlGeneratedKeys = xmlOut.createElement("MORTGAGEACCOUNTANDLOAN")
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.SaveMortgageAccountAndLoan(vxmlRequest, xmlGeneratedKeys)
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
    If xmlGeneratedKeys.hasChildNodes() Then
        xmlResponseElem.appendChild xmlGeneratedKeys
    End If
    
    Set ICustomerFinancialBO_SaveMortgageAccountAndLoan = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveMortgageAccountAndLoanExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveMortgageAccountAndLoanVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_SaveMortgageAccountAndLoan = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountAndLoanExit

End Function

Private Function ICustomerFinancialBO_UpdateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateArrearsHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArrearsHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateArrearsHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateArrearsHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateArrearsHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateArrearsHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArrearsHistoryExit

End Function

Private Function ICustomerFinancialBO_UpdateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankCard a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankCardVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateBankCard"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateBankCard(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateBankCard = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankCardExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankCardExit

End Function

Private Function ICustomerFinancialBO_UpdateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankruptcyHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankruptcyHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateBankruptcyHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateBankruptcyHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateBankruptcyHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankruptcyHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialBO_UpdateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateCCJHistory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCCJHistoryVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateCCJHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateCCJHistory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateCCJHistory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateCCJHistoryExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCCJHistoryExit

End Function

Private Function ICustomerFinancialBO_UpdateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateDeclinedMortgage a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDeclinedMortgageVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateDeclinedMortgage"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateDeclinedMortgage(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateDeclinedMortgage = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateDeclinedMortgageExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateDeclinedMortgageExit

End Function

Private Function ICustomerFinancialBO_UpdateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateGroupConnection a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateGroupConnectionVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateGroupConnection"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateGroupConnection(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateGroupConnection = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateGroupConnectionExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateGroupConnectionExit

End Function

Private Function ICustomerFinancialBO_UpdateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLiability a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLiabilityVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateLiability"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateLiability(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateLiability = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateLiabilityExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLiabilityExit

End Function

Private Function ICustomerFinancialBO_UpdateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLifeProduct a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeProductVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateLifeProduct"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateLifeProduct(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateLifeProduct = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdatelifeproductExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatelifeproductExit

End Function

Private Function ICustomerFinancialBO_UpdateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageAccount a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageAccountVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateMortgageAccount"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateMortgageAccount(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateMortgageAccount = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageAccountExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageAccountExit

End Function

Private Function ICustomerFinancialBO_UpdateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateRegularOutgoings a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRegularOutgoingsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateRegularOutgoings"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateRegularOutgoings(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateRegularOutgoings = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateRegularOutgoingsExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateRegularOutgoingsExit

End Function

Private Function ICustomerFinancialBO_UpdateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageLoan a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageLoanVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateMortgageLoan"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerFinancialTxBO = New CustomerFinancialTxBO
    Else
        Set objICustomerFinancialTxBO = objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateMortgageLoan(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateMortgageLoan = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageLoanExit:

    Set objErrAssist = Nothing
    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageLoanExit

End Function

Private Function ICustomerFinancialBO_ValidateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateArrearsHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' DESCRIPTIONOFLOAN cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DESCRIPTIONOFLOAN")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 150
    End If
    
    ' DATECLEARED must not be after today's date
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATECLEARED")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If CDate(strThisValue) >= Date Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 274
        End If
    End If

    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 158
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateArrearsHistory = xmlResponseElem.cloneNode(True)

    
ValidateArrearsHistoryExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateArrearsHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateArrearsHistoryExit

End Function

Private Function ICustomerFinancialBO_ValidateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateBankCard"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' CARDTYPE cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//CARDTYPE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 147
    End If
    
    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 158
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateBankCard = xmlResponseElem.cloneNode(True)

    
ValidateBankCardExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateBankCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateBankCardExit

End Function

Private Function ICustomerFinancialBO_ValidateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateBankruptcyHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String
    Dim dtDateDeclared As Date
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' MONTHLYREPAYMENT cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//MONTHLYREPAYMENT")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 153
    End If
    
    ' DATEDECLARED must not be after today's date
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATEDECLARED")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
        If IsDate(strThisValue) Then
            dtDateDeclared = CDate(strThisValue)
        End If
    End If
    If dtDateDeclared >= Date Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 162
    End If

    ' DATEDISCHARGED must not be before the DATEDECLARED
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATEOFDISCHARGE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If CDate(strThisValue) < dtDateDeclared Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 140
        End If
    End If

    Set ICustomerFinancialBO_ValidateBankruptcyHistory = xmlResponseElem.cloneNode(True)

    
ValidateBankruptcyHistoryExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateBankruptcyHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialBO_ValidateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateCCJHistory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
    Dim dtDateOfJudgement As Date
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' MONTHLYREPAYMENT cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//MONTHLYREPAYMENT")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 152
    End If
    
    ' DATEOFJUDGEMENT must not be greater than today's date
    strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATEOFJUDGEMENT")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
        If IsDate(strThisValue) Then
            dtDateOfJudgement = CDate(strThisValue)
        End If
    End If
    If dtDateOfJudgement > Date Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 160
    End If

    ' DATEOFJUDGEMENT must not be greater than DATECLEARED
    strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATECLEARED")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If dtDateOfJudgement > CDate(strThisValue) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 214
        End If
    End If

    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 159
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateCCJHistory = xmlResponseElem.cloneNode(True)

    
ValidateCCJHistoryExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateCCJHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateCCJHistoryExit

End Function

Private Function ICustomerFinancialBO_ValidateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateDeclinedMortgage"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' DECLINEDDETAILS cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DECLINEDDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 149
    End If
    
    ' DATEDECLINED must not be after today's date
    strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//DATEDECLINED")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If CDate(strThisValue) >= Date Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 162
        End If
    End If

    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 157
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateDeclinedMortgage = xmlResponseElem.cloneNode(True)

    
ValidateDeclinedMortgageExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateDeclinedMortgageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateDeclinedMortgageExit

End Function

Private Function ICustomerFinancialBO_ValidateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateGroupConnection"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' ACCOUNTTYPE cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ACCOUNTTYPE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 263
    End If

    Set ICustomerFinancialBO_ValidateGroupConnection = xmlResponseElem.cloneNode(True)

    
ValidateGroupConnectionExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateGroupConnectionVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateGroupConnectionExit

End Function

Private Function ICustomerFinancialBO_ValidateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateLiability"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' MONTHLYREPAYMENT cannot be null
    strThisValue = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//MONTHLYREPAYMENT")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If strThisValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 151
    End If
    
    ' ENDDATE must not be before today's date
    strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ENDDATE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If CDate(strThisValue) < Date Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 161
        End If
    End If

    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 155
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerFinancialBO_ValidateLiability = xmlResponseElem.cloneNode(True)

    
ValidateLiabilityExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateLiabilityExit

End Function

Private Function ICustomerFinancialBO_ValidateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateLifeProduct"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 158
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateLifeProduct = xmlResponseElem.cloneNode(True)

    
ValidateLifeProductExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateLifeProductVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateLifeProductExit

End Function

Private Function ICustomerFinancialBO_ValidateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateMortgageAccount"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' INDEMNITYDATEPAID must not be after today's date
    'strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//INDEMNITYDATEPAID")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
        If IsDate(strThisValue) Then
            If CDate(strThisValue) > Date Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 219
            End If
        End If
    End If

    ' LASTVALUATIONDATE must not be after today's date
    'strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//LASTVALUATIONDATE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
        If IsDate(strThisValue) Then
            If CDate(strThisValue) > Date Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 219
            End If
        End If
    End If
    
    ' ADDITIONALINDICATOR and ADDITIONALDETAILS must correlate
    strThisValue = ""
    strThisValue2 = ""
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALINDICATOR")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//ADDITIONALDETAILS")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue2 = Trim$(xmlThisNode.Text)
    End If

    If (strThisValue = "1") And (strThisValue2 = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 221
    End If
    If (strThisValue <> "1") And (strThisValue2 <> "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 262
    End If

    Set ICustomerFinancialBO_ValidateMortgageAccount = xmlResponseElem.cloneNode(True)

    
ValidateMortgageAccountExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateMortgageAccountVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateMortgageAccountExit

End Function

Private Function ICustomerFinancialBO_ValidateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateRegularOutgoings"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set ICustomerFinancialBO_ValidateRegularOutgoings = xmlResponseElem.cloneNode(True)

    
ValidateRegularOutgoingsExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    
    Exit Function

ValidateRegularOutgoingsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateRegularOutgoingsExit

End Function

Private Function ICustomerFinancialBO_ValidateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidateMortgageLoan"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objComoboDo As ComboDO
    Dim objConvertAssist As ConvertAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String, _
        strThisValue2 As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' STARTDATE must not be after today's date
    strThisValue = "0"
    Set xmlThisNode = vxmlRequest.selectSingleNode(".//STARTDATE")
    If Not (xmlThisNode Is Nothing) Then
        strThisValue = Trim$(xmlThisNode.Text)
    End If
    If IsDate(strThisValue) Then
        If CDate(strThisValue) >= Date Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 219
        End If
    End If
    
    ' Original Part&Part IntOnly Amount cannot be greater than the Original Loan
    Set objConvertAssist = New ConvertAssist
    strThisValue = objXmlAssist.GetTagValue(vxmlRequest, "ORIGINALLOANAMOUNT")
    strThisValue2 = objXmlAssist.GetTagValue(vxmlRequest, "ORIGINALPARTANDPARTINTONLYAMT")
    
    If strThisValue <> "" And (Not IsNull(strThisValue)) And strThisValue2 <> "" And (Not IsNull(strThisValue2)) Then
        If objConvertAssist.CSafeDbl(strThisValue2) > objConvertAssist.CSafeDbl(strThisValue) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 141
        End If
    End If
    
    ' Redemption Date cannot be greater than currentdate
    Dim strRedemptionStatus As String, blnIsDeemedLoan As Boolean
    Dim blnLoanTobePorted As Boolean
    
    strRedemptionStatus = objXmlAssist.GetTagValue(vxmlRequest, "REDEMPTIONSTATUS")
    Set objComoboDo = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    If Len(strRedemptionStatus) > 0 Then  ' SR 06/09/2004 : BMIDS815
        blnIsDeemedLoan = objComoboDo.IsItemInValidation("RedemptionStatus", _
                                                         strRedemptionStatus, _
                                                         "A")
    End If  ' SR 06/09/2004 : BMIDS815
    If blnIsDeemedLoan Then
        strThisValue = objXmlAssist.GetTagValue(vxmlRequest, "REDEMPTIONDATE")
        If strThisValue <> "" Then
            If CDate(strThisValue) >= Date Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 552
            End If
        End If
    Else
        '    Validate for Ported Loan and existing lenders
        If Len(strRedemptionStatus) > 0 Then  ' SR 06/09/2004 : BMIDS815
            blnLoanTobePorted = objComoboDo.IsItemInValidation("RedemptionStatus", _
                                                         strRedemptionStatus, _
                                                         "P")
         End If  ' SR 06/09/2004 : BMIDS815
        '-----------------------------------------------------------------------
        'If the loan is to be ported in this transaction and, was not in the
        'ported previously then do the following validation.
        '-----------------------------------------------------------------------
        If blnLoanTobePorted Then
            ' Validate for Ported Loan Outstanding Balance
            Dim strInitialRedemptionStatus As String, strDirectoryGUID As String
            Dim blnPortedLoan As Boolean
            strInitialRedemptionStatus = objXmlAssist.GetTagValue(vxmlRequest, _
                                                   "INITIALREDEMPTIONSTATUS")
            If strInitialRedemptionStatus <> "" Then
                ' Check whether this was a ported loan before this transaction
                blnPortedLoan = objComoboDo.IsItemInValidation("RedemptionStatus", _
                                                     strInitialRedemptionStatus, _
                                                     "P")
            End If
            strDirectoryGUID = objXmlAssist.GetTagValue(vxmlRequest, "DIRECTORYGUID")
            If strDirectoryGUID = "" Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 538
            Else
               Dim xmlTempDoc  As FreeThreadedDOMDocument40
               Dim xmlTableElem  As IXMLDOMElement
               Dim xmlDirectoryIDNode  As IXMLDOMNode
               Dim xmlTempNode As IXMLDOMNode
               Dim objICustomerFinancialDO As ICustomerFinancialDO
               Dim xmlLenderOrganisationNode As IXMLDOMNode
                '   Get the OrganizationId for loan to be ported
                      '''''''''''''''''''''''''''''''
                Set xmlTempDoc = New FreeThreadedDOMDocument40
                xmlTempDoc.validateOnParse = False
                xmlTempDoc.setProperty "NewParser", True
                Set xmlTableElem = xmlTempDoc.createElement("MORTGAGELENDERDIRECTORY")
                Set xmlDirectoryIDNode = xmlTempDoc.createElement("DIRECTORYGUID")
                xmlDirectoryIDNode.Text = strDirectoryGUID
                xmlTableElem.appendChild xmlDirectoryIDNode.cloneNode(True)
                If objContext Is Nothing Then
                 Set objICustomerFinancialDO = New CustomerFinancialDO
                Else
                 Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                End If
                
                Set xmlLenderOrganisationNode = objXmlAssist.GetNode(objICustomerFinancialDO.GetMortgageLenderDirectoryData(xmlTableElem), "./ORGANISATIONID")
                      
                ' Build the request string, and call validatePortedLoan
                
                Set xmlTempNode = vxmlRequest.appendChild(xmlLenderOrganisationNode)
                Set xmlTempNode = ICustomerFinancialBO_ValidatePortedLoan(vxmlRequest)
                objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
            End If
        End If
    End If
    
    
    Set ICustomerFinancialBO_ValidateMortgageLoan = xmlResponseElem.cloneNode(True)
    
ValidateMortgageLoanExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objComoboDo = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlThisNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlLenderOrganisationNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDirectoryIDNode = Nothing
    Set xmlTableElem = Nothing
    Set xmlTempDoc = Nothing
    
    Exit Function

ValidateMortgageLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidateMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateMortgageLoanExit

End Function

Public Function CreateArrearsHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateArrearsHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateArrearsHistoryVbErr

    Const strFunctionName As String = "CreateArrearsHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateArrearsHistory default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateArrearsHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateArrearsHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateArrearsHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateArrearsHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateArrearsHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArrearsHistoryExit

End Function

Public Function CreateBankCard(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateBankCard a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankCardVbErr

    Const strFunctionName As String = "CreateBankCard"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateBankCard default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateBankCard(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateBankCard(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateBankCard = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateBankCardExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateBankCard = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankCardExit

End Function

Public Function CreateBankruptcyHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateBankruptcyHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankruptcyHistoryVbErr

    Const strFunctionName As String = "CreateBankruptcyHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateBankruptcyHistory default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateBankruptcyHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateBankruptcyHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateBankruptcyHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateBankruptcyHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateBankruptcyHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankruptcyHistoryExit

End Function

Public Function CreateCCJHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateCCJHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateCCJHistoryVbErr

    Const strFunctionName As String = "CreateCCJHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateCCJHistory default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateCCJHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateCCJHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateCCJHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateCCJHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateCCJHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateCCJHistoryExit

End Function

Public Function CreateDeclinedMortgage(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateDeclinedMortgage a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclinedMortgageVbErr

    Const strFunctionName As String = "CreateDeclinedMortgage"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateDeclinedMortgage default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateDeclinedMortgage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateDeclinedMortgage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateDeclinedMortgage = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateDeclinedMortgageExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateDeclinedMortgage = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateDeclinedMortgageExit

End Function

Public Function CreateGroupConnection(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateGroupConnection a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateGroupConnectionVbErr

    Const strFunctionName As String = "CreateGroupConnection"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateGroupConnection default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateGroupConnection = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateGroupConnectionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateGroupConnection = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateGroupConnectionExit

End Function

Public Function CreateLiability(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateLiability a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateLiabilityVbErr

    Const strFunctionName As String = "CreateLiability"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateLiability default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//LOANSLIABILITIES/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    'BMIDS00561 MDC 08/10/2002
    'CreateLiability = xmlResponseElem.xml
    CreateLiability = xmlTempResponseNode.xml
    'BMIDS00561 MDC 08/10/2002 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateLiabilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateLiability = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLiabilityExit

End Function

Public Function CreateLifeProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateLifeProduct a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateLifeProductVbErr

    Const strFunctionName As String = "CreateLifeProduct"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateLifeProduct default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//ACCOUNT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateLifeProduct(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateLifeProduct(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateLifeProduct = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateLifeProductExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateLifeProduct = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLifeProductExit

End Function

Public Function CreateMortgageAccount(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageAccount a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageAccountVbErr

    Const strFunctionName As String = "CreateMortgageAccount"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateMortgageAccount default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//ACCOUNT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateMortgageAccount(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                          
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateMortgageAccount(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
 
    CreateMortgageAccount = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateMortgageAccount = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageAccountExit

End Function

Public Function CreateRegularOutgoings(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateRegularOutgoings a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateRegularOutgoingsVbErr

    Const strFunctionName As String = "CreateRegularOutgoings"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateRegularOutgoings default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateRegularOutgoings(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                          
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateRegularOutgoings(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
 
    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    CreateRegularOutgoings = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateRegularOutgoingsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateRegularOutgoings = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateRegularOutgoingsExit

End Function

Public Function CreateMortgageLoan(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLoan a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLoanVbErr

    Const strFunctionName As String = "CreateMortgageLoan"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateMortgageLoan default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateMortgageLoan(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateMortgageLoan(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateMortgageLoan = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateMortgageLoanExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateMortgageLoan = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageLoanExit

End Function

Public Function DeleteArrearsHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteArrearsHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArrearsHistoryVbErr

    Const strFunctionName As String = "DeleteArrearsHistory"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteArrearsHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteArrearsHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteArrearsHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteArrearsHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArrearsHistoryExit

End Function

Public Function DeleteBankCard(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteBankCard a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankCardVbErr

    Const strFunctionName As String = "DeleteBankCard"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteBankCard(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteBankCard = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankCardExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteBankCard = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankCardExit

End Function

Public Function DeleteBankruptcyHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteBankruptcyHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankruptcyHistoryVbErr

    Const strFunctionName As String = "DeleteBankruptcyHistory"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteBankruptcyHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteBankruptcyHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankruptcyHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteBankruptcyHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankruptcyHistoryExit

End Function

Public Function DeleteCCJHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteCCJHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCCJHistoryVbErr

    Const strFunctionName As String = "DeleteCCJHistory"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteCCJHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteCCJHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteCCJHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteCCJHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCCJHistoryExit

End Function

Public Function DeleteDeclinedMortgage(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteDeclinedMortgage a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDeclinedMortgageVbErr

    Const strFunctionName As String = "DeleteDeclinedMortgage"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteDeclinedMortgage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteDeclinedMortgage = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteDeclinedMortgageExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteDeclinedMortgage = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteDeclinedMortgageExit

End Function

Public Function DeleteGroupConnection(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteGroupConnection a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteGroupConnectionVbErr

    Const strFunctionName As String = "DeleteGroupConnection"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteGroupConnection = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteGroupConnectionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteGroupConnection = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteGroupConnectionExit

End Function

Public Function DeleteLiability(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteLiability a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLiabilityVbErr

    Const strFunctionName As String = "DeleteLiability"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteLiability = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLiabilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteLiability = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLiabilityExit

End Function

Public Function DeleteLifeProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteLifeProduct a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeProductVbErr

    Const strFunctionName As String = "DeleteLifeProduct"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteLifeProduct(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteLifeProduct = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLifeProductExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteLifeProduct = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLifeProductExit

End Function

Public Function DeleteMortgageAccount(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageAccount a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageAccountVbErr

    Const strFunctionName As String = "DeleteMortgageAccount"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteMortgageAccount(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteMortgageAccount = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteMortgageAccount = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageAccountExit

End Function

Public Function DeleteMortgageLoan(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageLoan a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageLoanVbErr

    Const strFunctionName As String = "DeleteMortgageLoan"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteMortgageLoan(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteMortgageLoan = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteMortgageLoanExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteMortgageLoan = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageLoanExit

End Function

Public Function DeleteRegularOutgoings(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteRegularOutgoings a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteRegularOutgoingsVbErr

    Const strFunctionName As String = "DeleteRegularOutgoings"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteRegularOutgoings(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteRegularOutgoings = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteRegularOutgoingsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteRegularOutgoings = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteRegularOutgoingsExit

End Function

Public Function FindArrearsHistoryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindArrearsHistoryListVbErr

    Const strFunctionName As String = "FindArrearsHistoryList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindArrearsHistoryList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindArrearsHistoryList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindArrearsHistoryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindArrearsHistoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindArrearsHistoryList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindArrearsHistoryListExit

End Function

Public Function FindBankCardList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindBankCardListVbErr

    Const strFunctionName As String = "FindBankCardList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindBankCardList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindBankCardList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindBankCardListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindBankCardListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBankCardList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankCardListExit

End Function

Public Function FindBankruptcyHistoryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindBankruptcyHistoryListVbErr

    Const strFunctionName As String = "FindBankruptcyHistoryList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindBankruptcyHistoryList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindBankruptcyHistoryList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindBankruptcyHistoryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindBankruptcyHistoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBankruptcyHistoryList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankruptcyHistoryListExit

End Function

Public Function FindCCJHistoryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindCCJHistoryListVbErr

    Const strFunctionName As String = "FindCCJHistoryList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindCCJHistoryList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindCCJHistoryList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindCCJHistoryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindCCJHistoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCCJHistoryList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindCCJHistoryListExit

End Function

Public Function FindDeclinedMortgageList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindDeclinedMortgageListVbErr

    Const strFunctionName As String = "FindDeclinedMortgageList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindDeclinedMortgageList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindDeclinedMortgageList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindDeclinedMortgageListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindDeclinedMortgageListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindDeclinedMortgageList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindDeclinedMortgageListExit

End Function

Public Function FindGroupConnectionList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindGroupConnectionListVbErr

    Const strFunctionName As String = "FindGroupConnectionList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindGroupConnectionList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindGroupConnectionList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindGroupConnectionListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindGroupConnectionListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindGroupConnectionList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindGroupConnectionListExit

End Function

Public Function FindLiabilityList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLiabilityListVbErr

    Const strFunctionName As String = "FindLiabilityList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindLiabilityList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindLiabilityList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindLiabilityListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindLiabilityListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLiabilityList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLiabilityListExit

End Function

Public Function FindLifeProductList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLifeProductListVbErr

    Const strFunctionName As String = "FindLifeProductList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindLifeProductList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindLifeProductList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindLifeProductListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindLifeProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLifeProductList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLifeProductListExit

End Function

Public Function FindLoansToBePorted(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLoansToBePortedVbErr

    Const strFunctionName As String = "FindLoansToBePorted"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindLoansToBePorted(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindLoansToBePorted = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindLoansToBePortedExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindLoansToBePortedVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLoansToBePorted = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLoansToBePortedExit

End Function

Public Function FindMortgageAccountList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindMortgageAccountListVbErr

    Const strFunctionName As String = "FindMortgageAccountList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindMortgageAccountList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindMortgageAccountList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindMortgageAccountListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindMortgageAccountListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageAccountList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete          ' PSC 29/01/01 SYS1888
    End If
    
    ' go to clean-up section
    Resume FindMortgageAccountListExit

End Function

Public Function FindMortgageAccountAndLoanList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

'************** BG 26/07/01  Note I tried to use this method -
'************** it isn't finished in the DO yet            ***
 
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindMortgageAccountAndLoanList"
    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindMortgageAccountAndLoanList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindMortgageAccountAndLoanList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageAccountAndLoanList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function FindOutgoingsTotalList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindOutgoingsTotalListVbErr

    Const strFunctionName As String = "FindOutgoingsTotalList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindOutgoingsTotalList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindOutgoingsTotalList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindOutgoingsTotalListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindOutgoingsTotalListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindOutgoingsTotalList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindOutgoingsTotalListExit

End Function

Public Function FindRegularOutgoingsList(ByVal vstrXMLRequest As String) As String
Attribute FindRegularOutgoingsList.VB_Description = "Business Process to return relevant data from the ""regularOutgoings"" table"
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' Note:         There is a version of this function within AgreementInprinciple which
'               also looks for Quick quote Outgoings if there are no Regular outgoings
'               (QQ Outgoings used as defaults)
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindRegularOutgoingsListVbErr

    Const strFunctionName As String = "FindRegularOutgoingsList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindRegularOutgoingsList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindRegularOutgoingsList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindRegularOutgoingsListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindRegularOutgoingsListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindRegularOutgoingsList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindRegularOutgoingsListExit

End Function

Public Function FindMortgageListForArrears(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindMortgageListForArrearsVbErr

    Const strFunctionName As String = "FindMortgageListForArrears"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindMortgageListForArrears(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindMortgageListForArrears = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindMortgageListForArrearsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindMortgageListForArrearsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageListForArrears = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindMortgageListForArrearsExit

End Function

Public Function FindMortgageLoanList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindMortgageLoanListVbErr

    Const strFunctionName As String = "FindMortgageLoanList"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindMortgageLoanList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindMortgageLoanList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindMortgageLoanListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindMortgageLoanListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageLoanList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindMortgageLoanListExit

End Function

Public Function SaveGroupConnection(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveGroupConnection a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveGroupConnectionVbErr

    Const strFunctionName As String = "SaveGroupConnection"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn  As FreeThreadedDOMDocument40, _
        xmlOut As New FreeThreadedDOMDocument40
        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_SaveGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveGroupConnection = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveGroupConnectionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing

    Exit Function

SaveGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveGroupConnection = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveGroupConnectionExit

End Function

Public Function SaveLiability(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Saves multiple instances of the persistant data associated with the
'               the table 'LoansLiabilities'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveLiabilityVbErr

    Const strFunctionName As String = "SaveLiability"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn  As FreeThreadedDOMDocument40, _
        xmlOut As New FreeThreadedDOMDocument40
        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_SaveLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveLiability = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveLiabilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing

    Exit Function

SaveLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveLiability = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLiabilityExit

End Function

Public Function UpdateArrearsHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateArrearsHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArrearsHistoryVbErr

    Const strFunctionName As String = "UpdateArrearsHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateArrearsHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateArrearsHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateArrearsHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateArrearsHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateArrearsHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateArrearsHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArrearsHistoryExit
    
End Function

Public Function UpdateBankCard(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateBankCard a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankCardVbErr

    Const strFunctionName As String = "UpdateBankCard"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateBankCard(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateBankCard(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateBankCard = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateBankCardExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateBankCardVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateBankCard = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankCardExit
    
End Function

Public Function UpdateBankruptcyHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateBankruptcyHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankruptcyHistoryVbErr

    Const strFunctionName As String = "UpdateBankruptcyHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateBankruptcyHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateBankruptcyHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateBankruptcyHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateBankruptcyHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateBankruptcyHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateBankruptcyHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankruptcyHistoryExit
    
End Function

Public Function UpdateCCJHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateCCJHistory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCCJHistoryVbErr

    Const strFunctionName As String = "UpdateCCJHistory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateCCJHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateCCJHistory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateCCJHistory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateCCJHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateCCJHistoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCCJHistory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCCJHistoryExit
    
End Function

Public Function UpdateDeclinedMortgage(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateDeclinedMortgage a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDeclinedMortgageVbErr

    Const strFunctionName As String = "UpdateDeclinedMortgage"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateDeclinedMortgage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateDeclinedMortgage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateDeclinedMortgage = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateDeclinedMortgageExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateDeclinedMortgageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateDeclinedMortgage = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateDeclinedMortgageExit
    
End Function

Public Function UpdateGroupConnection(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateGroupConnection a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateGroupConnectionVbErr

    Const strFunctionName As String = "UpdateGroupConnection"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateGroupConnection(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateGroupConnection = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateGroupConnectionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateGroupConnectionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateGroupConnection = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateGroupConnectionExit
    
End Function

Public Function UpdateLiability(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateLiability a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLiabilityVbErr

    Const strFunctionName As String = "UpdateLiability"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//LOANSLIABILITIES/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateLiability(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLiability = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateLiabilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateLiabilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateLiability = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLiabilityExit
    
End Function

Public Function UpdateLifeProduct(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateLifeProduct a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeProductVbErr

    Const strFunctionName As String = "UpdateLifeProduct"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//ACCOUNT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateLifeProduct(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateLifeProduct(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLifeProduct = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdatelifeproductExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateLifeProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateLifeProduct = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatelifeproductExit
    
End Function

Public Function UpdateMortgageAccount(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageAccount a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageAccountVbErr

    Const strFunctionName As String = "UpdateMortgageAccount"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//ACCOUNT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateMortgageAccount(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateMortgageAccount(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateMortgageAccount = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateMortgageAccountVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateMortgageAccount = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageAccountExit
    
End Function

Public Function UpdateRegularOutgoings(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateRegularOutgoings a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRegularOutgoingsVbErr

    Const strFunctionName As String = "UpdateRegularOutgoings"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateRegularOutgoings(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateRegularOutgoings(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateRegularOutgoings = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateRegularOutgoingsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateRegularOutgoingsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateRegularOutgoings = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateRegularOutgoingsExit
    
End Function

Public Function UpdateMortgageLoan(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageLoan a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageLoanVbErr

    Const strFunctionName As String = "UpdateMortgageLoan"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateMortgageLoan(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateMortgageLoan(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateMortgageLoan = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateMortgageLoanExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageLoanVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateMortgageLoan = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageLoanExit
    
End Function

Private Function ICustomerFinancialBO_ValidatePortedLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidatePortedLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialBO_ValidatePortedLoan"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvertAssist As New ConvertAssist
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim objMortgageLender As IMortgageLenderBO
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlAccountDoc As New FreeThreadedDOMDocument40
    xmlAccountDoc.validateOnParse = False
    xmlAccountDoc.setProperty "NewParser", True
    Dim xmlAccountListNode As IXMLDOMNode
    Dim xmlAccountNode As IXMLDOMNode
    Dim xmlCustomerNumberNode As IXMLDOMNode
    Dim xmlCustomerVersionNumberNode As IXMLDOMNode
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode

    Dim xmlThisNode As IXMLDOMNode
    Dim strThisValue  As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
''''''''''''''''''''''''''''''''''''''''''''''
' Build Mortgage Account List XML from Request
''''''''''''''''''''''''''''''''''''''''''''''
    Set xmlAccountDoc = New FreeThreadedDOMDocument40
    xmlAccountDoc.validateOnParse = False
    xmlAccountDoc.setProperty "NewParser", True
          xmlAccountDoc.async = False
    
          Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
          xmlAccountDoc.appendChild xmlRequestElem
    
          
          
     Set xmlAccountListNode = xmlAccountDoc.createElement("MORTGAGEACCOUNTLIST")
     xmlRequestElem.appendChild xmlAccountListNode
     Set xmlAccountNode = xmlAccountDoc.createElement("MORTGAGEACCOUNT")
     xmlAccountListNode.appendChild xmlAccountNode
     Set xmlCustomerNumberNode = objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERNUMBER")
     xmlAccountNode.appendChild xmlCustomerNumberNode.cloneNode(True)
     Set xmlCustomerVersionNumberNode = objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
     xmlAccountNode.appendChild xmlCustomerVersionNumberNode.cloneNode(True)

    Set xmlResponseNode = objICustomerFinancialDO.FindMortgageAccountList(xmlAccountListNode)

''''''''''''''''''''''''''''''''''''''''''
' get the account guid from the response
''''''''''''''''''''''''''''''''''''''''''
       Dim xmlAccountList As IXMLDOMNodeList
 
       Set xmlAccountDoc = New FreeThreadedDOMDocument40
       xmlAccountDoc.validateOnParse = False
       xmlAccountDoc.setProperty "NewParser", True
       xmlAccountDoc.appendChild xmlResponseNode
       Set xmlAccountList = xmlAccountDoc.selectNodes("MORTGAGEACCOUNTLIST/MORTGAGEACCOUNT/ACCOUNT")
       

    If xmlAccountList Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "Missing Account GUID"
        
    Else
          Dim xmlTempDoc As FreeThreadedDOMDocument40
          Dim xmlTableElem As IXMLDOMElement
          Dim xmlLoanList As IXMLDOMNodeList
          Dim xmlLoanNode As IXMLDOMNode
          Dim xmlAccountGuid As IXMLDOMNode
    
          Dim xmlAccount As IXMLDOMNode
          Dim strResponse As String
          Dim strRedemptionStatus As String
          Dim xmlRedemptionNode As IXMLDOMNode
          Dim xmlDirectoryIDNode As IXMLDOMNode
          Dim objComboDO As ComboDO
          Dim xmlLenderOrganisationNode As IXMLDOMNode
    
          Dim intPortedLoanCount As Integer
          Dim intMaxNoLoans As Integer
          intPortedLoanCount = 0
    
    '
          Set objComboDO = New ComboDO
          For Each xmlAccount In xmlAccountList
              Set xmlAccountGuid = objXmlAssist.GetNode(xmlAccount, "ACCOUNTGUID")
              Set xmlDirectoryIDNode = objXmlAssist.GetNode(xmlAccount, "DIRECTORYGUID")
              Set xmlTempDoc = New FreeThreadedDOMDocument40
              xmlTempDoc.validateOnParse = False
              xmlTempDoc.setProperty "NewParser", True
              xmlTempDoc.async = False
    
              Set xmlRequestElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
              xmlTempDoc.appendChild xmlRequestElem
    
              Set xmlTableElem = xmlTempDoc.createElement("MORTGAGELOAN")
              xmlRequestElem.appendChild xmlTableElem
              xmlTableElem.appendChild xmlAccountGuid.cloneNode(True)
        '
              strResponse = FindMortgageLoanList(xmlTempDoc.xml)
              Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
              objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
        '
              Set xmlLoanList = xmlTempDoc.selectNodes(".//MORTGAGELOAN")
        '''''''''''''''''''''''''''''''''
        '        ' Go through each  loan.
        '''''''''''''''''''''''''''''''''
              For Each xmlLoanNode In xmlLoanList
    
                  strRedemptionStatus = objXmlAssist.GetNodeValue(xmlLoanNode, "REDEMPTIONSTATUS")
                  If objComboDO.IsItemInValidation("RedemptionStatus", _
                                                     strRedemptionStatus, _
                                                     "P") Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'check Ported Loan is not current one being validated
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If objXmlAssist.GetNodeValue(xmlLoanNode, "ACCOUNTGUID") <> _
                       objXmlAssist.GetNodeValue(vxmlRequest, ".//ACCOUNTGUID") Then
                       intPortedLoanCount = intPortedLoanCount + 1
                       '''''''''''''''''''''''''''''''
                       ' Check Lender for Ported Loan
                       '''''''''''''''''''''''''''''''
                       Set xmlTempDoc = New FreeThreadedDOMDocument40
                       xmlTempDoc.validateOnParse = False
                       xmlTempDoc.setProperty "NewParser", True
                       Set xmlTableElem = xmlTempDoc.createElement("MORTGAGELENDERDIRECTORY")
                       xmlTableElem.appendChild xmlDirectoryIDNode.cloneNode(True)
                       Set xmlResponseNode = objICustomerFinancialDO.GetMortgageLenderDirectoryData(xmlTableElem)
                       Set xmlLenderOrganisationNode = objXmlAssist.GetNode(xmlResponseNode, "./ORGANISATIONID")
    '
                       If xmlLenderOrganisationNode.Text <> _
                          objXmlAssist.GetNodeValue(vxmlRequest, ".//ORGANISATIONID") Then
                              objErrAssist.ThrowError TypeName(Me), strFunctionName, 255
                       End If
    
                    End If
    
                  End If
              Next
    
        Next
        '''''''''''''''''''''''''''''''''''''''''''''
        'Get the Maximum No of loans for this lender
        '''''''''''''''''''''''''''''''''''''''''''''
        Set objMortgageLender = objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        ' make sure we get the details for the requested lender
        Set xmlLenderOrganisationNode = objXmlAssist.GetNode(vxmlRequest, ".//ORGANISATIONID")
        Set xmlTableElem = xmlTempDoc.createElement("MORTGAGELENDER")
        xmlRequestElem.appendChild xmlTableElem
        xmlTableElem.appendChild xmlLenderOrganisationNode.cloneNode(True)
        Set xmlResponseNode = objMortgageLender.GetLenderDetails(xmlTableElem)
        objErrAssist.CheckResponse xmlResponseNode, xmlResponseElem, True
    
        intMaxNoLoans = objConvertAssist.CSafeInt(objXmlAssist.GetNodeValue(xmlResponseNode.firstChild, "MAXNOLOANS"))
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' don't do check if Max number of loans has not been set
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If ((intMaxNoLoans > 0) And (intPortedLoanCount >= intMaxNoLoans)) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 142
        End If
'
    End If
  
    
    Set ICustomerFinancialBO_ValidatePortedLoan = xmlResponseElem.cloneNode(True)
    
ValidatePortedLoanExit:

   
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objContext = Nothing
    Set xmlResponseElem = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlAccountDoc = Nothing
    Set xmlAccountListNode = Nothing
    Set xmlAccountNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNumberNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlAccountDoc = Nothing
    Set xmlAccountList = Nothing
    Set objComboDO = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlDirectoryIDNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlLoanList = Nothing
    Set xmlResponseNode = Nothing
    Set xmlLenderOrganisationNode = Nothing
    Set objMortgageLender = Nothing
    Set xmlLenderOrganisationNode = Nothing

    
    Exit Function

ValidatePortedLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_ValidatePortedLoan = objErrAssist.CreateErrorResponseEx()

    
    ' go to clean-up section
    Resume ValidatePortedLoanExit
    
End Function

Public Function GetMortgageAccount(ByVal vstrXMLRequest As String) As String


' header ----------------------------------------------------------------------------------
' description:  Gets the total of outstanding balances and the total or repayments being
'               made on loans that are not to be redeemed
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountVbErr

    Const strFunctionName As String = "GetMortgageAccount"

        
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
    Set xmlTempResponseNode = ICustomerFinancialBO_GetMortgageAccount(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageAccount = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageAccount = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageAccountExit

End Function

' SR 25/05/2004 : BMIDS772 - new method
Public Function GetFinancialSummaryView(ByVal vstrXMLRequest As String) As String

On Error GoTo GetFinancialSummaryViewVbErr

    Const cstrFunctionName As String = "GetFinancialSummaryView"

        
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
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerFinancialBO_GetFinancialSummaryView(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetFinancialSummaryView = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetFinancialSummaryViewExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetFinancialSummaryViewVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetFinancialSummaryView = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetFinancialSummaryViewExit

End Function


' SR 25/05/2004 : BMIDS772 - new method
Public Function GetCreditHistorySummaryView(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCreditHistorySummaryViewVbErr

    Const cstrFunctionName As String = "GetCreditHistorySummaryView"

        
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
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerFinancialBO_GetCreditHistorySummaryView(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCreditHistorySummaryView = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCreditHistorySummaryViewExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetCreditHistorySummaryViewVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetCreditHistorySummaryView = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetCreditHistorySummaryViewExit

End Function



Public Function SaveMortgageAccount(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calls either CustomerFinancialBO.Create/Delete/UpdateMortgageAccount
'               Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveMortgageAccountVbErr

    Const strFunctionName As String = "SaveMortgageAccount"

    
        
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
    Set xmlTempResponseNode = ICustomerFinancialBO_ValidateMortgageAccount(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_SaveMortgageAccount(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    SaveMortgageAccount = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
  

    Exit Function

SaveMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveMortgageAccount = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountExit
    

End Function

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

Public Function SaveMortgageAccountAndLoan(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveMortgageAccountAndLoanVbErr

    Const strFunctionName As String = "SaveMortgageAccountAndLoan"

    
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
    ' SR 13/07/00 - No validation required (this method is written for egg Home zone)
    
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                         
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerFinancialBO_SaveMortgageAccountAndLoan(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    SaveMortgageAccountAndLoan = xmlResponseElem.xml
   
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveMortgageAccountAndLoanExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveMortgageAccountAndLoanVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveMortgageAccountAndLoan = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountAndLoanExit
    
End Function

Public Function GetLoansNotRedeemedTotals(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the total of outstanding balances and the total or repayments being
'               made on loans that are not to be redeemed
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLoansNotRedeemedTotalsVbErr

    Const strFunctionName As String = "GetLoansNotRedeemedTotals"

        
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
    Set xmlTempResponseNode = ICustomerFinancialBO_GetLoansNotRedeemedTotals(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLoansNotRedeemedTotals = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetLoansNotRedeemedTotalsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLoansNotRedeemedTotalsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLoansNotRedeemedTotals = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetLoansNotRedeemedTotalsExit

End Function

Public Function GetAccountDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountDetailsVbErr

    Const strFunctionName As String = "GetAccountDetails"

    
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
    Set xmlTempResponseNode = ICustomerFinancialBO_GetAccountDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetAccountDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetAccountDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAccountDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAccountDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAccountDetailsExit

End Function

Public Function CreateAccountRelationships(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create instances of Account Relationship
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountRelationshipsVbErr

    Const strFunctionName As String = "CreateAccountRelationships"
    
         
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
        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateAccountRelationships(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateAccountRelationships = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateAccountRelationshipsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

CreateAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateAccountRelationships = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateAccountRelationshipsExit

End Function


Private Function ICustomerFinancialBO_CreateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create instances of Account Relationship
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountRelationshipsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateAccountRelationships"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
        
    '   call Business Transaction Object CreateAccountRelationships function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateAccountRelationships(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_CreateAccountRelationships = xmlResponseElem

    m_objContext.SetComplete
    
CreateAccountRelationshipsExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    

    Exit Function

CreateAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateAccountRelationshipsExit

End Function

Public Function DeleteAccountRelationships(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Account Relationship
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAccountRelationshipsVbErr

    Const strFunctionName As String = "DeleteAccountRelationships"
    
         
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
        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteAccountRelationships(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DeleteAccountRelationships = xmlResponseElem.xml
    
    m_objContext.SetComplete

DeleteAccountRelationshipsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

DeleteAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAccountRelationships = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteAccountRelationshipsExit

End Function


Private Function ICustomerFinancialBO_DeleteAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Account Relationship
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAccountRelationshipsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteAccountRelationships"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
        
    '   call Business Transaction Object DeleteAccountRelationships function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteAccountRelationships(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_DeleteAccountRelationships = xmlResponseElem

    m_objContext.SetComplete
    
DeleteAccountRelationshipsExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    

    Exit Function

DeleteAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteAccountRelationshipsExit

End Function

Public Function UpdateAccountRelationships(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update instances of Account Relationship
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountRelationshipsVbErr

    Const strFunctionName As String = "UpdateAccountRelationships"
    
         
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
        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateAccountRelationships(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateAccountRelationships = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateAccountRelationshipsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

UpdateAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateAccountRelationships = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateAccountRelationshipsExit

End Function


Private Function ICustomerFinancialBO_UpdateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instances of Account Relationship
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountRelationshipsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateAccountRelationships"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
        
    '   call Business Transaction Object UpdateAccountRelationships function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateAccountRelationships(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerFinancialBO_UpdateAccountRelationships = xmlResponseElem

    m_objContext.SetComplete
    
UpdateAccountRelationshipsExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    

    Exit Function

UpdateAccountRelationshipsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateAccountRelationshipsExit

End Function

Public Function FindAccountRelationshipList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Find instances of Account Relationship
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo FindAccountRelationshipListVbErr

    Const strFunctionName As String = "FindAccountRelationshipList"
    
       
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindAccountRelationshipList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindAccountRelationshipList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindAccountRelationshipListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindAccountRelationshipListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindAccountRelationshipList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindAccountRelationshipListExit

End Function

Private Function ICustomerFinancialBO_FindAccountRelationshipList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Find instances of Account Relationship
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindAccountRelationshipListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindAccountRelationshipList"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlCopy As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName = cstrACCOUNT_ROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
    
    Set xmlCopy = xmlRequestNode.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlCopy, "ACCOUNT", "ACCOUNTRELATIONSHIP"
        
    Set xmlDataNode = objICustomerFinancialDO.FindAccountRelationshipList(xmlCopy)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICustomerFinancialBO_FindAccountRelationshipList = xmlResponseElem

FindAccountRelationshipListExit:

    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCopy = Nothing
    

    Exit Function

FindAccountRelationshipListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindAccountRelationshipList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindAccountRelationshipListExit

End Function










Public Function CreateSpecialFeature(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create instances of Mortgage Account Special Feature
'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateSpecialFeatureVbErr

    Const strFunctionName As String = "CreateSpecialFeature"
    
        
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
        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CreateSpecialFeature(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateSpecialFeature = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateSpecialFeatureExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

CreateSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateSpecialFeature = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateSpecialFeatureExit

End Function

Public Function UpdateSpecialFeature(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update instances of Mortgage Account Special Feature
'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSpecialFeatureVbErr

    Const strFunctionName As String = "UpdateSpecialFeature"
    
        
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
    
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_UpdateSpecialFeature(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateSpecialFeature = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateSpecialFeatureExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

UpdateSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateSpecialFeature = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateSpecialFeatureExit
    
End Function


Public Function DeleteSpecialFeature(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Mortgage Account Special Feature
'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteSpecialFeatureVbErr

    Const strFunctionName As String = "DeleteSpecialFeature"
    
        
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
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_DeleteSpecialFeature(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteSpecialFeature = xmlResponseElem.xml
    
    m_objContext.SetComplete

DeleteSpecialFeatureExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

DeleteSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteSpecialFeature = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteSpecialFeatureExit

End Function

Public Function FindSpecialFeatureList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Find instances of Mortgage Account Special Feature
'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindSpecialFeatureListVbErr

    Const strFunctionName As String = "FindSpecialFeatureList"
    
       
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindSpecialFeatureList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindSpecialFeatureList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindSpecialFeatureListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindSpecialFeatureListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindSpecialFeatureList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindSpecialFeatureListExit

End Function

Private Function ICustomerFinancialBO_CreateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create instances of Mortgage Account Special Feature
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateSpecialFeatureVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CreateSpecialFeature"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.CreateSpecialFeature(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set ICustomerFinancialBO_CreateSpecialFeature = xmlResponseElem

    m_objContext.SetComplete
    
CreateSpecialFeatureExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    

    Exit Function

CreateSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CreateSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateSpecialFeatureExit

End Function

Private Function ICustomerFinancialBO_DeleteSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Mortgage Account Special Feature
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteSpecialFeatureVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_DeleteSpecialFeature"

            
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objICustomerFinancialTxBO.DeleteSpecialFeature(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_DeleteSpecialFeature = xmlResponseElem

    m_objContext.SetComplete

DeleteSpecialFeatureExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    

    Exit Function

DeleteSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_DeleteSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteSpecialFeatureExit

End Function

Private Function ICustomerFinancialBO_FindSpecialFeatureList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Find instances of Mortgage Account Special Feature
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindSpecialFeatureListVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindSpecialFeatureList"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName = cstrFEATUREROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrFEATUREROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFEATUREROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindSpecialFeatureList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICustomerFinancialBO_FindSpecialFeatureList = xmlResponseElem

FindSpecialFeatureListExit:

    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindSpecialFeatureListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindSpecialFeatureList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindSpecialFeatureListExit

End Function


Private Function ICustomerFinancialBO_UpdateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instances of Mortgage Account Special Feature
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSpecialFeatureVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_UpdateSpecialFeature"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerFinancialTxBO As ICustomerFinancialTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerFinancialTxBO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialTxBO")
    
    Set xmlTempResponseNode = objICustomerFinancialTxBO.UpdateSpecialFeature(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerFinancialBO_UpdateSpecialFeature = xmlResponseElem

    m_objContext.SetComplete
    
UpdateSpecialFeatureExit:

    Set objICustomerFinancialTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

UpdateSpecialFeatureVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_UpdateSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateSpecialFeatureExit

End Function

Public Function GetMortgageAccountDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get details of the mortgage account
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountDetailsVbErr

    Const strFunctionName As String = "GetMortgageAccountDetails"
        
        
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
    Set xmlTempResponseNode = ICustomerFinancialBO_GetMortgageAccountDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageAccountDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetMortgageAccountDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetMortgageAccountDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageAccountDetails = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetMortgageAccountDetailsExit

End Function

Private Function ICustomerFinancialBO_GetMortgageAccountDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get details of the mortgage account
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountDetailsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_GetMortgageAccountDetails"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
                
    If vxmlRequest.nodeName = cstrMORTACCOUNT_ROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrMORTACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrMORTACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetMortgageAccountDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerFinancialBO_GetMortgageAccountDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetMortgageAccountDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetMortgageAccountDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetMortgageAccountDetails = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetMortgageAccountDetailsExit

End Function

'BMIDS00190
Public Function CalculateCreditCardRepayment(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the monthly repayment for a Credit Card
' pass:         vstrXmlRequest  xml Request data stream containing outstanding balance
' return:                       xml Response data stream containing montly repaymant amount
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateCreditCardRepaymentVbErr

    Const strFunctionName As String = "CalculateCreditCardRepayment"

    
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem     As IXMLDOMElement
    Dim objIScriptInterface As IScriptInterface
    
    ' CalculateCreditCardRepayment default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerFinancialBO_CalculateCreditCardRepayment(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateCreditCardRepayment = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateCreditCardRepaymentExit:

    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

CalculateCreditCardRepaymentVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateCreditCardRepayment = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateCreditCardRepaymentExit
End Function

Private Function ICustomerFinancialBO_CalculateCreditCardRepayment(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Calculates the monthly repayment for a Credit Card
' pass:         vstrXmlRequest  xml Request data stream containing outstanding balance
' return:                       xml Response data stream containing montly repaymant amount
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateCreditCardRepaymentVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_CalculateCreditCardRepayment"

    
    Dim xmlResponseElem         As IXMLDOMElement
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlRepayment            As IXMLDOMNode
    Dim objConvAssist           As ConvertAssist
    Dim objGlobalParameterDO    As GlobalParameterDO
    Dim dblTotal                As Double
    Dim dblRepayment            As Double
    Dim dblPercent              As Double
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objConvAssist = New ConvertAssist
    dblTotal = objConvAssist.CSafeDbl(m_objXmlAssist.GetMandatoryElementText(vxmlRequest, "TOTALOUTSTANDINGBALANCE"))
    
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    dblPercent = objConvAssist.CSafeDbl(objGlobalParameterDO.GetCurrentParameterByType("CreditCardPercentage", "PERCENTAGE"))

    dblRepayment = dblTotal * dblPercent / 100

    Set xmlRepayment = xmlOut.createElement("MONTHLYREPAYMENT")
    xmlRepayment.Text = Format(dblRepayment, "0.00")
    xmlResponseElem.appendChild xmlRepayment

    Set ICustomerFinancialBO_CalculateCreditCardRepayment = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateCreditCardRepaymentExit:
    
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParameterDO = Nothing
    
    Exit Function
    
CalculateCreditCardRepaymentVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_CalculateCreditCardRepayment = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateCreditCardRepaymentExit

End Function

Public Function FindLoansListForArrears(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLoansListForArrearsVbErr

    Const strFunctionName As String = "FindLoansListForArrears"

       
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
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerFinancialBO_FindLoansListForArrears(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindLoansListForArrears = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindLoansListForArrearsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindLoansListForArrearsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLoansListForArrears = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLoansListForArrearsExit

End Function

Private Function ICustomerFinancialBO_FindLoansListForArrears(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLoansListForArrearsVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindLoansListForArrears"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If vxmlRequest.nodeName = "LOANSLIABILITIESLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("LOANSLIABILITIESLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "LOANSLIABILITIESLIST" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindLoansListForArrears(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set ICustomerFinancialBO_FindLoansListForArrears = xmlResponseElem.cloneNode(True)

FindLoansListForArrearsExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindLoansListForArrearsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindLoansListForArrears = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLoansListForArrearsExit

End Function

'BMIDS00190 End

'BMIDS00336 MDC 10/09/2002
Public Function GetMortgageAccountLoanTotals(ByVal vstrXMLRequest As String) As String

On Error GoTo GetMortgageAccountLoanTotalsVbErr

Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

Const strFunctionName As String = "GetMortgageAccountLoanTotals"
        
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerFinancialBO_GetMortgageAccountLoanTotals(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageAccountLoanTotals = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetMortgageAccountLoanTotalsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetMortgageAccountLoanTotalsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageAccountLoanTotals = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetMortgageAccountLoanTotalsExit

End Function
'BMIDS00336 MDC 10/09/2002 - End

'BMIDS00444
Public Function FindRemortgageAccountAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindRemortgageAccountAddressVbErr

    Const strFunctionName As String = "FindRemortgageAccountAddress"

       
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
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
    Set xmlTempResponseNode = ICustomerFinancialBO_FindRemortgageAccountAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindRemortgageAccountAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindRemortgageAccountAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindRemortgageAccountAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindRemortgageAccountAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindRemortgageAccountAddressExit

End Function


Private Function ICustomerFinancialBO_FindRemortgageAccountAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Find the address related to each mortgage account that has its remortgage
'               indicator set to true
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindRemortgageAccountAddressVbErr
    
    Const strFunctionName As String = "ICustomerFinancialBO_FindRemortgageAccountAddress"


    Dim xmlRequestNode          As IXMLDOMNode
    Dim xmlDataNode             As IXMLDOMNode
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlResponseElem         As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If vxmlRequest.nodeName = "MORTGAGEACCOUNT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEACCOUNT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.FindRemortgageAccountAddress(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set ICustomerFinancialBO_FindRemortgageAccountAddress = xmlResponseElem.cloneNode(True)

FindRemortgageAccountAddressExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindRemortgageAccountAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_FindRemortgageAccountAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindRemortgageAccountAddressExit

End Function
'BMIDS00444 End


' SR 25/05/2004 : BMIDS815 - new method
Public Function GetPortedStepAndPeriodFromMortgageAccount(ByVal vstrXMLRequest As String) As String

On Error GoTo GetPortedStepAndPeriodFromMortgageAccountVbErr

    Const cstrFunctionName As String = "GetPortedStepAndPeriodFromMortgageAccount"

        
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
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerFinancialBO_GetPortedStepAndPeriodFromMortgageAccount(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPortedStepAndPeriodFromMortgageAccount = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetPortedStepAndPeriodFromMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPortedStepAndPeriodFromMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetPortedStepAndPeriodFromMortgageAccount = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPortedStepAndPeriodFromMortgageAccountExit

End Function

'EP2_954 GHun
Private Function ICustomerFinancialBO_GetAdditionalBorrowingDataForProjectionCalc(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    Const cstrFunctionName As String = "ICustomerFinancialBO_GetAdditionalBorrowingDataForProjectionCalc"
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlResponseElem         As IXMLDOMElement
    Dim xmlRequestNode          As IXMLDOMNode
    Dim xmlDataNode             As IXMLDOMNode

On Error GoTo GetAdditionalBorrowingDataForProjectionCalcErr

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If vxmlRequest.nodeName = "MORTGAGEACCOUNT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEACCOUNT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerFinancialDO.GetAdditionalBorrowingDataForProjectionCalc(xmlRequestNode)
    SetRatePeriods xmlDataNode
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set ICustomerFinancialBO_GetAdditionalBorrowingDataForProjectionCalc = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetAdditionalBorrowingDataForProjectionCalcExit:
    
    Set objICustomerFinancialDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Exit Function

GetAdditionalBorrowingDataForProjectionCalcErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialBO_GetAdditionalBorrowingDataForProjectionCalc = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAdditionalBorrowingDataForProjectionCalcExit
End Function
'EP2_954 End

'EP2_954 GHun
Private Sub SetRatePeriods(ByVal vxmlData As IXMLDOMElement)
    Const cstrFunctionName  As String = "SetRatePeriods"
    Dim xmlMortgageLoan         As IXMLDOMElement
    Dim dteStartDate            As Date
    Dim dteEndDate              As Date
    Dim lngPeriodDifference     As Long
    Dim lngExpiredPeriod        As Long
    Dim lngTotalExpiredPeriod   As Long
    Dim lngStartingSequence     As Long
    Dim lngEndDateDifference    As Long
    Dim xmlInterestRateType     As IXMLDOMElement
    Dim xmlNode                 As IXMLDOMNode
    Dim strDate                 As String
    
    For Each xmlMortgageLoan In vxmlData.selectNodes("MORTGAGELOAN")
        lngPeriodDifference = 0
        lngTotalExpiredPeriod = 0
        lngExpiredPeriod = 0
        lngStartingSequence = 0
        strDate = m_objXmlAssist.GetMandatoryElementText(xmlMortgageLoan, "STARTDATE")
        If IsDate(strDate) Then
            dteStartDate = CDate(strDate)
            
            'This relies on the InterestRateTypes being sorted by InterestRateTypeSequenceNumber (ascending)
            For Each xmlInterestRateType In xmlMortgageLoan.selectNodes("INTERESTRATETYPE")
                
                If (lngStartingSequence = 0) Then
                    strDate = m_objXmlAssist.GetElementText(xmlInterestRateType, "INTERESTRATEENDDATE")
                    If IsDate(strDate) Then
                        dteEndDate = CDate(strDate)
                        If Date < dteEndDate Then
                            lngStartingSequence = CLng(m_objXmlAssist.GetMandatoryElementText(xmlInterestRateType, "INTERESTRATETYPESEQUENCENUMBER"))
                        Else
                            xmlMortgageLoan.removeChild xmlInterestRateType
                            lngEndDateDifference = MonthDiffRoundUp(dteEndDate, dteStartDate)
                            lngTotalExpiredPeriod = lngTotalExpiredPeriod + lngEndDateDifference
                            dteStartDate = dteEndDate
                        End If
                    Else
                        strDate = m_objXmlAssist.GetMandatoryElementText(xmlMortgageLoan, "PRODUCTSTARTDATE")
                        If IsDate(strDate) Then
                            lngPeriodDifference = MonthDiffRoundUp(CDate(strDate), Date)
                        Else
                            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMissingElement, "MORTGAGELOAN/PRODUCTSTARTDATE"
                        End If
                    
                        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlInterestRateType, "INTERESTRATEPERIOD")
                        lngExpiredPeriod = CLng(xmlNode.Text)
                        
                        If lngExpiredPeriod = -1 Then
                            lngStartingSequence = CLng(m_objXmlAssist.GetMandatoryElementText(xmlInterestRateType, "INTERESTRATETYPESEQUENCENUMBER"))
                        Else
                            lngTotalExpiredPeriod = lngTotalExpiredPeriod + lngExpiredPeriod
                            
                            If lngTotalExpiredPeriod > lngPeriodDifference Then
                                lngStartingSequence = CLng(m_objXmlAssist.GetMandatoryElementText(xmlInterestRateType, "INTERESTRATETYPESEQUENCENUMBER"))
    
                                'Remaining period
                                xmlNode.Text = lngTotalExpiredPeriod - lngPeriodDifference
                            Else
                                xmlMortgageLoan.removeChild xmlInterestRateType
                                dteStartDate = DateAdd("m", lngExpiredPeriod, dteStartDate)
                            End If
                        End If
                    End If
                End If
                
            Next
        End If
    Next
End Sub
'EP2_954 End

'EP2_954 GHun
'Calculate the number of months difference between 2 dates rounded up to the nearest whole month
Private Function MonthDiffRoundUp(ByVal vdteDate1 As Date, ByVal vdteDate2 As Date) As Long
    Dim lngMonths As Long
    
    lngMonths = DateDiff("m", vdteDate1, vdteDate2)
    'If the difference has been rounded up then decrement the number of months
    If DateDiff("d", DateAdd("m", lngMonths, vdteDate1), vdteDate2) > 0 Then
        lngMonths = lngMonths + 1
    End If
    
    MonthDiffRoundUp = lngMonths
End Function
'EP2_954 End

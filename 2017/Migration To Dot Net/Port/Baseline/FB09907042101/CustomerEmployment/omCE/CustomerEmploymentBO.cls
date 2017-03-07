VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerEmploymentBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4BO' with actual class name
' *****************************************************************************

'Workfile:      CustomerEmploymentBO.cls
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
'               e.g. CustomerEmploymentTxBO, CustomerEmploymentDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'----------------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MCS    09/11/1999  Created
'IVW    08/05/1999  SYS0508 - need to zeroize totals otherwise it can
'SR     26/05/00    Modified method - SaveEmployed details. Check whether a record exists
'                   before creating it, rather than handling it in error condition (which
'                   was causing automation error)
'JLD    09/06/00    Modified SaveOtherLoan to cope with a list of otherincomes. Also enabled the delete option
'JLD    12/06/00    Added method GetCurrentEmploymentAndIncome for homezone use.
'SR     24/07/00    New method 'FindEmploymentList'
'SR     26/07/00    Modified method 'SaveAccountant', if Employment tag found in the Request, update that table
'MC     31/07/00    New method GetEmploymentAndSelfEmployedDetails
'BG     22/08/00    SYS0501 CalculateEarnedTotals  - cleared variables for each iteration.
'SR     31/08/00    SYS1498 - Modified method SaveOtherIncome; validate the request only in case of create or update
'SR     13/10/00    SYS1601 - Modified private sub CalculateEarnedTotals. Move the statements clearing
'                   variables (SYS0501)to the immediate higher-level For loop (for each customer)
'DRC    22/03/01    SYS1932 Added GetEmploymentccountant & GetAccountantGUID methods
'SA     24/05/01    SYS2084 ValidateEarnedIncome. First item in the strAdditionalOptions argument of Throwerror
'                    is additional text not replacement text so added empty string to arguments.
'DRC    03/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Check Methods
'MEVA   23/04/02    SYS2742 Amend SaveTaxDetails to see if record exists instead of trying to insert and see if it fails
'AM     5/09/2005   Added new method GetincomeSummaryDetails
'-------------------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    14/05/2002  BMIDS00004 - Add CopyEmploymentDetails method
'AW     15/11/02    BMIDS00953 - Overridden CalculateEarnedTotals
'-------------------------------------------------------------------------------------------------------------------
'Mars History:
'
'Prog   Date        Description
'GHun   10/07/2006  MAR1891 Changed SaveTaxDetails
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerEmploymentBO

Private Function CalculateAverageNetProfit(ByVal vxmlData As IXMLDOMElement) As Long
' header ----------------------------------------------------------------------------------
' description:  Calculates the averagenetprofit from the xml "NETPROFIT" passed
'
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       AverageNetProfit As Long rounded UP
'
' errors:       Raise error oeMissingPrimaryTag if tag "NETPROFIT" not found
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAverageNetProfitVbErr
    
    Const strFunctionName As String = "CalculateAverageNetProfit"


    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    If vxmlData.nodeName = "NETPROFIT" Then
        Set xmlNode = vxmlData
    Else
        Set xmlNode = vxmlData.getElementsByTagName("NETPROFIT").Item(0)
    End If
    
    If xmlNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "NETPROFIT" & " tag not found"
    End If
    
    Dim xmlElement As IXMLDOMElement
    Dim lngAverageNetProfit As Long
    Dim lngTotalNetProfit As Long
    Dim intNetProfitCount As Integer
    Dim lngTempAmount As Long
    
    intNetProfitCount = 0
    lngAverageNetProfit = 0
    lngTotalNetProfit = 0
    lngTempAmount = 0

    Set xmlNode = vxmlData.selectSingleNode("/NETPROFIT/YEAR1AMOUNT")
    lngTempAmount = Val(xmlNode.Text)
    
    If lngTempAmount > 0 Then
        
        intNetProfitCount = intNetProfitCount + 1
        lngTotalNetProfit = lngTotalNetProfit + lngTempAmount
        lngTempAmount = 0
        
    End If
    
    Set xmlNode = vxmlData.selectSingleNode("/NETPROFIT/YEAR2AMOUNT")
    lngTempAmount = Val(xmlNode.Text)

    If lngTempAmount > 0 Then
        
        intNetProfitCount = intNetProfitCount + 1
        lngTotalNetProfit = lngTotalNetProfit + lngTempAmount
        lngTempAmount = 0
        
    End If
    
    Set xmlNode = vxmlData.selectSingleNode("/NETPROFIT/YEAR3AMOUNT")
    lngTempAmount = Val(xmlNode.Text)

    
    If lngTempAmount > 0 Then
        
        intNetProfitCount = intNetProfitCount + 1
        lngTotalNetProfit = lngTotalNetProfit + lngTempAmount
        lngTempAmount = 0
        
    End If

    If intNetProfitCount > 0 Then
    
        lngAverageNetProfit = lngTotalNetProfit / intNetProfitCount
        
    Else
    
        lngAverageNetProfit = 0
        
    End If
            
    CalculateAverageNetProfit = Fix(lngAverageNetProfit) 'round UP
    

    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing

    Exit Function

CalculateAverageNetProfitVbErr:
    
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlElement = Nothing
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
            
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Sub CalculateEarnedTotals(ByVal vxmlData As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Calculates the earned income totals for employment(s) passed in appends
'               the TOTALGROSSEARNEDINCOME and the TOTALNETMONTHLYINCOME to the incomesummary
'               also appends the YEARLYTOTALAMOUNT (employed) or AVERAGENETPROFIT
'               (contract, self-employed) to each employment
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateEarnedTotalsVbErr
    
    Const strFunctionName As String = "CalculateEarnedTotals"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As ConvertAssist
    
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Set objICustomerEmploymentDO = New CustomerEmploymentDO

    Dim lngTotalNetMonthlyIncome As Long
    Dim lngTotalGrossEarnedIncome As Long
    
    Dim xmlCustomerNodelist As IXMLDOMNodeList
    Set xmlCustomerNodelist = vxmlData.getElementsByTagName("CUSTOMER")

    Dim xmlCustomerNode As IXMLDOMElement
    
    Dim nCustomerloop As Integer
    Dim xmlElement As IXMLDOMElement

    Dim xmlEmploymentNodeList As IXMLDOMNodeList
    Dim lngNetMonthlyIncome As Long
    Dim lngTotalEarnedIncome As Long
    Dim strEmploymentStatus As String
    Dim strNetMonthlyIncome As String
    Dim nEmploymentloop As Integer

    Dim xmlIncomeSummaryNode As IXMLDOMNode
    Dim xmlAppendNode As IXMLDOMNode

    Dim lngYearlyAmount As Long
    Dim lngYearlyTotalAmount As Long
    Dim lngEarnedIncomeAmount As Long
    Dim nPaymentFrequency As Integer
    Dim xmlEarnedIncome As IXMLDOMElement
    Dim xmlEarnedIncomeNodeList As IXMLDOMNodeList
    Dim xmlEarnedIncomeElement As IXMLDOMElement
    Dim nEarnedIncomeLoop As Integer
    Dim xmlResponse As IXMLDOMNode

    Dim strFromDate As String, _
        strToDate As String
    Dim blnSkipTotals As Boolean
    
    Set objConvAssist = New ConvertAssist
    
    For nCustomerloop = 0 To (xmlCustomerNodelist.length - 1)
        'loop round all the customers and calculate the earnedincome for each
        Set xmlCustomerNode = xmlCustomerNodelist.Item(nCustomerloop)
        Set xmlEmploymentNodeList = xmlCustomerNode.getElementsByTagName("EMPLOYMENT")
                
        'SR 13/10/00 SYS1601
        lngTotalNetMonthlyIncome = 0
        lngTotalEarnedIncome = 0
        lngTotalGrossEarnedIncome = 0
                
        For nEmploymentloop = 0 To (xmlEmploymentNodeList.length - 1)
            'BG 22/08/00 SYS0501 clear variables for each iteration.
            'SR 13/10/00 SYS1601 - clear variable in the higher-level For loop (for each customer)
            'lngTotalNetMonthlyIncome = 0
            'lngTotalEarnedIncome = 0
            'lngTotalGrossEarnedIncome = 0
            
            'loop round all the customers and calculate the earnedincome for each
            'Get the net monthly income value here to be used in the calculation
            Set xmlElement = xmlEmploymentNodeList.Item(nEmploymentloop)

            strFromDate = objXmlAssist.GetTagValue(xmlElement, "DATESTARTEDORESTABLISHED")
            strToDate = objXmlAssist.GetTagValue(xmlElement, "DATELEFTORCEASEDTRADING")

            ' Only include current employments in totals
            blnSkipTotals = False
            If IsDate(strFromDate) Then
                blnSkipTotals = Not (CDate(strFromDate) < Now)
            End If
            If IsDate(strToDate) Then
                blnSkipTotals = Not (CDate(strToDate) > Now)
            End If
                
            strNetMonthlyIncome = objXmlAssist.GetTagValue(xmlElement, "NETMONTHLYINCOME")
            lngNetMonthlyIncome = objConvAssist.CSafeLng(strNetMonthlyIncome)
            'add this on
            If Not blnSkipTotals Then
                lngTotalNetMonthlyIncome = lngTotalNetMonthlyIncome + lngNetMonthlyIncome
            End If
            
            strEmploymentStatus = objXmlAssist.GetTagValue(xmlElement, "EMPLOYMENTSTATUS")
            
            If objICustomerEmploymentDO.IsEmploymentStatus(Employed, strEmploymentStatus) Then
                objXmlAssist.ChangeNodeName xmlElement, "EMPLOYMENT", "EARNEDINCOME"
                lngYearlyTotalAmount = 0

                On Error GoTo GetFinancialDetailsVbErr
                'AW 15/11/02    BMIDS00953
                Set xmlEarnedIncome = Nothing
                'AW 15/11/02    BMIDS00953 - End
                Set xmlEarnedIncome = objICustomerEmploymentDO.FindEarnedIncomeList(xmlElement)
                On Error GoTo CalculateEarnedTotalsVbErr

                If Not (xmlEarnedIncome Is Nothing) Then
                    Set xmlEarnedIncomeNodeList = xmlEarnedIncome.getElementsByTagName("EARNEDINCOME")
    
                    'loop round all the earned incomes for this employment
                    For nEarnedIncomeLoop = 0 To (xmlEarnedIncomeNodeList.length - 1)
                        Set xmlEarnedIncomeElement = xmlEarnedIncomeNodeList.Item(nEarnedIncomeLoop)
                        ' loop round all earnedincome records for this employment
                        'for each earnedincomesequencenumber get the EarnedIncomeAmount
                        lngEarnedIncomeAmount = Val(objXmlAssist.GetTagValue(xmlEarnedIncomeElement, "EARNEDINCOMEAMOUNT"))
                        nPaymentFrequency = Val(objXmlAssist.GetTagValue(xmlEarnedIncomeElement, "PAYMENTFREQUENCYTYPE"))
                        
                        lngYearlyAmount = (lngEarnedIncomeAmount * CInt(nPaymentFrequency))
                        lngYearlyTotalAmount = lngYearlyTotalAmount + lngYearlyAmount
                    Next
                End If

                'Append the lngYearlyTotalAmount amount to the employment for this customer
                Set xmlAppendNode = vxmlData.ownerDocument.createElement("YEARLYTOTALAMOUNT")
                xmlAppendNode.Text = lngYearlyTotalAmount
                'and add it to this 'employment'
                xmlElement.appendChild xmlAppendNode

                If Not blnSkipTotals Then
                    lngTotalEarnedIncome = lngTotalEarnedIncome + lngYearlyTotalAmount
                End If
            
                objXmlAssist.ChangeNodeName xmlElement, "EARNEDINCOME", "EMPLOYMENT"
            ElseIf objICustomerEmploymentDO.IsEmploymentStatus(SelfEmployed, strEmploymentStatus) _
                    Or objICustomerEmploymentDO.IsEmploymentStatus(Contract, strEmploymentStatus) Then

                'NetProfit is 1:1 so now a getdata
                Dim lngAverageNetProfit As Long
                
                objXmlAssist.ChangeNodeName xmlElement, "EMPLOYMENT", "NETPROFIT"
                
                'IVW - SYS0508 - need to zeroize totals otherwise it can
                'give an incorrect average
                
                lngAverageNetProfit = 0
                Set xmlResponse = Nothing
                
                On Error GoTo GetFinancialDetailsVbErr
                Set xmlResponse = objICustomerEmploymentDO.GetNetProfitDetails(xmlElement)
                On Error GoTo CalculateEarnedTotalsVbErr

                If Not (xmlResponse Is Nothing) Then
                    lngAverageNetProfit = CalculateAverageNetProfit(xmlResponse)
                End If

                If Not blnSkipTotals Then
                    lngTotalEarnedIncome = lngTotalEarnedIncome + lngAverageNetProfit
                End If
                            
                'Append the lngAverageNetProfit amount to the employment for this customer
                Set xmlAppendNode = vxmlData.ownerDocument.createElement("AVERAGENETPROFIT")
                xmlAppendNode.Text = lngAverageNetProfit
                'and add it to this 'employment'
                xmlElement.appendChild xmlAppendNode

                objXmlAssist.ChangeNodeName xmlElement, "NETPROFIT", "EMPLOYMENT"
            End If
        Next nEmploymentloop

        lngTotalGrossEarnedIncome = lngTotalGrossEarnedIncome + lngTotalEarnedIncome

        Set xmlIncomeSummaryNode = xmlCustomerNode.getElementsByTagName("INCOMESUMMARY").Item(0)
        If xmlIncomeSummaryNode Is Nothing Then
            ' No INCOMESUMMARY has yet been created for this customer
            Set xmlIncomeSummaryNode = xmlCustomerNode.appendChild(xmlCustomerNode.ownerDocument.createElement("INCOMESUMMARY"))
        End If
        
        'Append the lngTotalGrossEarnedIncome amount to the incomesummary for this customer
        If Fix(lngTotalGrossEarnedIncome) > 0 Then  'round UP
            Set xmlAppendNode = vxmlData.ownerDocument.createElement("TOTALGROSSEARNEDINCOME")
            
            xmlAppendNode.Text = lngTotalGrossEarnedIncome
            xmlIncomeSummaryNode.appendChild xmlAppendNode
        End If
        
        If Fix(lngTotalNetMonthlyIncome) > 0 Then  'round UP
            Set xmlAppendNode = vxmlData.ownerDocument.createElement("TOTALNETMONTHLYINCOME")
            
            xmlAppendNode.Text = lngTotalNetMonthlyIncome
            xmlIncomeSummaryNode.appendChild xmlAppendNode
        End If
    Next nCustomerloop
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlAppendNode = Nothing
    Set xmlAppendNode = Nothing
    Set xmlIncomeSummaryNode = Nothing
    Set xmlEarnedIncome = Nothing
    Set xmlResponse = Nothing
    Set objConvAssist = Nothing

    Exit Sub

GetFinancialDetailsVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

CalculateEarnedTotalsVbErr:

    Set objXmlAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlAppendNode = Nothing
    Set xmlAppendNode = Nothing
    Set xmlIncomeSummaryNode = Nothing
    Set xmlEarnedIncome = Nothing
    Set xmlResponse = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function CalculateGrossEarnedIncome(ByVal vstrCustomerNumber As String, _
                                            ByVal vstrCustomerVersionNumber As String) As Double
' header ----------------------------------------------------------------------------------
' description:  Calculates the Total Gross Income for the passed EMPLOYMENTLIST
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       Long
'------------------------------------------------------------------------------------------
    On Error GoTo CalculateGrossEarnedIncomeVbErr

Const strFunctionName As String = "CalculateGrossEarnedIncome"


Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
Dim objXmlAssist As New XMLAssist
Dim objErrAssist As New ErrAssist

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlCustomerNode As IXMLDOMNode
Dim xmlResponse As IXMLDOMNode

Dim strXML As String
Dim dblGrossEarnedIncome As Double

    '
    ' Build the XML to pass to FindEmploymentAndIncomeList
    '
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.appendChild xmlDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerNode = xmlDoc.firstChild.appendChild(xmlDoc.createElement("CUSTOMER"))

    ' Add Customer/Version Number to XML
    xmlCustomerNode.appendChild(xmlDoc.createElement("CUSTOMERNUMBER")).Text = vstrCustomerNumber
    xmlCustomerNode.appendChild(xmlDoc.createElement("CUSTOMERVERSIONNUMBER")).Text = vstrCustomerVersionNumber

    '
    ' Get the Employment+Income XML
    '
    Set xmlResponse = ICustomerEmploymentBO_FindEmploymentAndIncomeDetails(xmlDoc.documentElement)
    objErrAssist.CheckXMLResponse xmlResponse.xml, vblnRaiseError:=True

    dblGrossEarnedIncome = Val(objXmlAssist.GetTagValue(xmlResponse, "TOTALGROSSEARNEDINCOME"))

    CalculateGrossEarnedIncome = dblGrossEarnedIncome

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlResponse = Nothing

    Exit Function

CalculateGrossEarnedIncomeVbErr:
    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlResponse = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


'BMIDS00004 MDC 15/05/2002
Private Function ICustomerEmploymentBO_CopyEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CopyEmploymentDetails is used to copy employment data from one customer version
'               to another.
' pass:         vxmlRequest  xml Request data stream containing key data idenitifying
'               records to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CopyEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_CopyEmploymentDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As ErrAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set objErrAssist = New ErrAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
        
    '   call Business Transaction Object CopyEmploymentDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CopyEmploymentDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    Set ICustomerEmploymentBO_CopyEmploymentDetails = xmlResponseElem.cloneNode(True)

    objContext.SetComplete
    
CopyEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CopyEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_CopyEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    objContext.SetAbort
    
    ' go to clean-up section
    Resume CopyEmploymentDetailsExit

End Function
'BMIDS00004 MDC 15/05/2002 - End

Private Function ICustomerEmploymentBO_CreateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateAccountantDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountantDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_CreateAccountantDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
        
    '   call Business Transaction Object CreateAccountantDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateAccountantDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    Set ICustomerEmploymentBO_CreateAccountantDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateAccountantDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateAccountantDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_CreateAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateAccountantDetailsExit

End Function

Private Function ICustomerEmploymentBO_CreateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateEmploymentDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_CreateEmploymentDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
        
    '   call Business Transaction Object CreateEmploymentDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateEmploymentDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If

    Set ICustomerEmploymentBO_CreateEmploymentDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_CreateEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentBO_CreateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateOtherIncome a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherIncomeVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_CreateOtherIncome"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
        
    '   call Business Transaction Object CreateOtherIncome function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateOtherIncome(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerEmploymentBO_CreateOtherIncome = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOtherIncomeExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateOtherIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_CreateOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOtherIncomeExit

End Function

Private Function ICustomerEmploymentBO_DeleteEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteEmploymentDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_DeleteEmploymentDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
           
    '   call Business Transaction Object DeleteEmploymentDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.DeleteEmploymentDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerEmploymentBO_DeleteEmploymentDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_DeleteEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentBO_DeleteOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteOtherIncome a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOtherIncomeVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_DeleteOtherIncome"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
           
    '   call Business Transaction Object DeleteOtherIncome function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.DeleteOtherIncome(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerEmploymentBO_DeleteOtherIncome = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteOtherIncomeExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteOtherIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_DeleteOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteOtherIncomeExit

End Function

Private Function ICustomerEmploymentBO_FindEarnedIncomeList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindEarnedIncomeListVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_FindEarnedIncomeList"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EARNEDINCOME" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EARNEDINCOME").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EARNEDINCOME" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.FindEarnedIncomeList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerEmploymentBO_FindEarnedIncomeList = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml

FindEarnedIncomeListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindEarnedIncomeListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_FindEarnedIncomeList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindEarnedIncomeListExit

End Function

Private Function ICustomerEmploymentBO_FindEmploymentAndIncomeDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindEmploymentAndIncomeDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_FindEmploymentAndIncomeDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
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
        
    Set xmlDataNode = objICustomerEmploymentDO.FindEmploymentAndIncomeDetails(xmlRequestNode)
    CalculateEarnedTotals xmlDataNode

    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerEmploymentBO_FindEmploymentAndIncomeDetails = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml

FindEmploymentAndIncomeDetailsExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindEmploymentAndIncomeDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_FindEmploymentAndIncomeDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindEmploymentAndIncomeDetailsExit

End Function

Private Function ICustomerEmploymentBO_FindEmploymentList( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindEmploymentListVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_FindEmploymentList"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EMPLOYMENTLIST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EMPLOYMENTLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EMPLOYMENTLIST" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.FindEmploymentList(xmlRequestNode)

    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerEmploymentBO_FindEmploymentList = xmlResponseElem.cloneNode(True)

FindEmploymentListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindEmploymentListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_FindEmploymentList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindEmploymentListExit

End Function

Private Function ICustomerEmploymentBO_FindOtherIncomeList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindOtherIncomeListVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_FindOtherIncomeList"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO

    Dim xmlUnearnedIncomeList As IXMLDOMNode
    Dim xmlUnearnedIncomeNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim lngTotalOtherIncome       As Long, _
        lngTotalGrossIncome       As Long, _
        lngTotalGrossEarnedIncome As Long

    Dim strCustomerNumber        As String, _
        strCustomerVersionNumber As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "UNEARNEDINCOME" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("UNEARNEDINCOME").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "UNEARNEDINCOME" & " tag not found"
    End If

    On Error GoTo RecordNotFoundErr
    Set xmlDataNode = objICustomerEmploymentDO.FindOtherIncomeList(xmlRequestNode)
    On Error GoTo FindOtherIncomeListVbErr

    If xmlDataNode Is Nothing Then
        ' No UNEARNEDINCOME records found, so create an empty tag for putting the totals under
        Set xmlUnearnedIncomeList = xmlRequestNode.ownerDocument.createElement("UNEARNEDINCOMELIST")
    Else
        Set xmlUnearnedIncomeList = xmlDataNode
    End If

    '
    ' Process the returned data
    '
    strCustomerNumber = objXmlAssist.GetTagValue(xmlRequestNode, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlRequestNode, "CUSTOMERVERSIONNUMBER")

    '
    ' Loop through each UNEARNEDINCOME record in the XML and calculate the TotalOtherIncomeAmount
    '
    lngTotalOtherIncome = 0
    lngTotalGrossIncome = 0

    Set xmlNodeList = xmlUnearnedIncomeList.selectNodes("UNEARNEDINCOME")
    For Each xmlUnearnedIncomeNode In xmlNodeList
        ' Total Other Income Amount = Total Other Income Amount + (Unearned Income Amount * Payment Frequency)
        lngTotalOtherIncome = lngTotalOtherIncome + _
            (Val(objXmlAssist.GetTagValue(xmlUnearnedIncomeNode, "UNEARNEDINCOMEAMOUNT")) * _
             Val(objXmlAssist.GetTagValue(xmlUnearnedIncomeNode, "PAYMENTFREQUENCY")))
    Next xmlUnearnedIncomeNode

    '
    ' Calculate the TotalGrossIncome from TotalGrossEarnedIncome and TotalOtherIncome
    '
    ' Get Total Gross Earned Income
    lngTotalGrossEarnedIncome = CalculateGrossEarnedIncome(strCustomerNumber, strCustomerVersionNumber)
    lngTotalGrossIncome = lngTotalGrossEarnedIncome + lngTotalOtherIncome

    '
    ' Insert the results of the calculation into the XML
    '
    Set xmlNode = xmlUnearnedIncomeList.appendChild(xmlUnearnedIncomeList.ownerDocument.createElement("TOTALOTHERINCOMEAMOUNT"))
    xmlNode.Text = lngTotalOtherIncome

    Set xmlNode = xmlUnearnedIncomeList.appendChild(xmlUnearnedIncomeList.ownerDocument.createElement("TOTALGROSSINCOME"))
    xmlNode.Text = lngTotalGrossIncome

    xmlResponseElem.appendChild xmlUnearnedIncomeList
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_FindOtherIncomeList = xmlResponseElem.cloneNode(True)

FindOtherIncomeListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlUnearnedIncomeList = Nothing
    Set xmlUnearnedIncomeNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing

    Exit Function

RecordNotFoundErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

FindOtherIncomeListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_FindOtherIncomeList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindOtherIncomeListExit

End Function

Private Function ICustomerEmploymentBO_GetAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountantDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetAccountantDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "ACCOUNTANT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("ACCOUNTANT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ACCOUNTANT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetAccountantDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetAccountantDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetAccountantDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetAccountantDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAccountantDetailsExit

End Function

Private Function ICustomerEmploymentBO_GetAccountantGUID(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountantGUIDVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetAccountantGUID"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlTempRespNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim strAccountantGUID As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
            

    
    Set xmlTempRespNode = objICustomerEmploymentDO.GetAccountantGUID(vxmlRequest)
'
' Error Check
'
 
    xmlResponseElem.appendChild xmlTempRespNode
    objErrAssist.CheckResponse xmlResponseElem, xmlResponseElem, True
    

    Set ICustomerEmploymentBO_GetAccountantGUID = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetAccountantGUIDExit:
    

    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetAccountantGUIDVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set ICustomerEmploymentBO_GetAccountantGUID = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAccountantGUIDExit

End Function

Private Function ICustomerEmploymentBO_GetContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetContractDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetContractDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "CONTRACTDETAILS" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CONTRACTDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CONTRACTDETAILS" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetContractDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetContractDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetContractDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetContractDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetContractDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetContractDetailsExit

End Function

Private Function ICustomerEmploymentBO_GetEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetEmployedDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetEmployedDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EMPLOYEDDETAILS" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EMPLOYEDDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EMPLOYEDDETAILS" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetEmployedDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetEmployedDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmployedDetailsExit

End Function

Private Function ICustomerEmploymentBO_GetEmploymentAndSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentAndSelfEmployedDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetEmploymentAndSelfEmployedDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EMPLOYMENT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EMPLOYMENT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetEmploymentAndSelfEmployedDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetEmploymentAndSelfEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetEmploymentAndSelfEmployedDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetEmploymentAndSelfEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetEmploymentAndSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentAndSelfEmployedDetailsExit


End Function

Private Function ICustomerEmploymentBO_GetEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetEmploymentDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EMPLOYMENT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EMPLOYMENT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetEmploymentDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetEmploymentDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetEmploymentDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentBO_GetCurrentEmploymentAndIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetCurrentEmploymentAndIncome"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("EMPLOYMENT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "EMPLOYMENT" & " tag not found"
    End If
    
    'record not found is not an error
    On Error GoTo RecordNotFoundVbErr
    Set xmlDataNode = objICustomerEmploymentDO.GetEmploymentDetails(xmlRequestNode)
    On Error GoTo GetEmploymentVbErr
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    'change EMPLOYMENT tag to EMPLOYEDDETAILS
    objXmlAssist.ChangeNodeName xmlRequestNode, "EMPLOYMENT", "EMPLOYEDDETAILS"
    On Error GoTo RecordNotFoundVbErr
    Set xmlDataNode = objICustomerEmploymentDO.GetEmployedDetails(xmlRequestNode)
    On Error GoTo GetEmploymentVbErr
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    'change EMPLOYEDDETAILS to UNEARNEDINCOME
    objXmlAssist.ChangeNodeName xmlRequestNode, "EMPLOYEDDETAILS", "UNEARNEDINCOME"
    On Error GoTo RecordNotFoundVbErr
    Set xmlDataNode = objICustomerEmploymentDO.FindOtherIncomeList(xmlRequestNode)
    On Error GoTo GetEmploymentVbErr
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetCurrentEmploymentAndIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetEmploymentExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

RecordNotFoundVbErr:

    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

GetEmploymentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetCurrentEmploymentAndIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentExit

End Function

Private Function ICustomerEmploymentBO_GetIncomeSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
      
' header ----------------------------------------------------------------------------------
' AM  5/09/05 : New method to get the Employment Income Summary details
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetIncomeSummaryVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetIncomeSummary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlTempRespNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim strAccountantGUID As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
            
   
    Set xmlTempRespNode = objICustomerEmploymentDO.GetTaxDetails(vxmlRequest)
'
' Error Check
'
 
    xmlResponseElem.appendChild xmlTempRespNode
    objErrAssist.CheckResponse xmlResponseElem, xmlResponseElem, True
    

    Set ICustomerEmploymentBO_GetIncomeSummary = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetIncomeSummaryExit:
    

    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetIncomeSummaryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set ICustomerEmploymentBO_GetIncomeSummary = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetIncomeSummaryExit

End Function

Private Function ICustomerEmploymentBO_GetNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetNetProfitDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetNetProfitDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "NETPROFIT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("NETPROFIT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "NETPROFIT" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetNetProfitDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetNetProfitDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetNetProfitDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNetProfitDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetNetProfitDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetNetProfitDetailsExit

End Function

Private Function ICustomerEmploymentBO_GetSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetSelfEmployedDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_GetSelfEmployedDetails"


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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
                
    If vxmlRequest.nodeName = "SELFEMPLOYEDDETAILS" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("SELFEMPLOYEDDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "SELFEMPLOYEDDETAILS" & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerEmploymentDO.GetSelfEmployedDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerEmploymentBO_GetSelfEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetSelfEmployedDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetSelfEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_GetSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetSelfEmployedDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveAccountantDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                             Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveAccountantDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveAccountantDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveAccountantDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIScriptInterface As IScriptInterface
    Dim xmlKeysNode As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode, xmlNode As IXMLDOMNode

    Dim objIClassDef As IomCEClassDef
    Set objIClassDef = New omCEClassDef

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateAccountantDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                                                strFunctionName, vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    Select Case objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadAccountantData)
        Case booCreate:
            Set xmlTempResponseNode = ICustomerEmploymentBO_CreateAccountantDetails(vxmlRequest)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

        Case booUpdate:
            Set xmlTempResponseNode = ICustomerEmploymentBO_UpdateAccountantDetails(vxmlRequest)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case booDelete: ' N/A
        Case Else
            ' N/A
    End Select

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
    
    ' SR - 26/07/00-------------------------------------------------------------------------------
    ' If 'EMPLOYMENT' tag is mentioned in the Request, update this table.
    ' If the operation performed on accountant was CREATE, then add AccountantGUID to the request
    '---------------------------------------------------------------------------------------------
    Set xmlEmployment = vxmlRequest.selectSingleNode(".//EMPLOYMENT")
    If Not xmlEmployment Is Nothing Then
        If Not xmlKeysNode Is Nothing Then
            Set xmlNode = xmlEmployment.selectSingleNode(".//ACCOUNTANTGUID")
            If xmlNode Is Nothing Then
                Set xmlNode = xmlOut.createElement("ACCOUNTANTGUID")
            End If
            xmlNode.Text = xmlKeysNode.selectSingleNode("ACCOUNTANTGUID").Text
        End If
                
        Set xmlTempResponseNode = ICustomerEmploymentBO_UpdateEmploymentDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    Set ICustomerEmploymentBO_SaveAccountantDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveAccountantDetailsExit:

    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set objIScriptInterface = Nothing
    Set xmlKeysNode = Nothing
    Set xmlEmployment = Nothing
    Set xmlNode = Nothing
    
    Exit Function

SaveAccountantDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveAccountantDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveContractDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                           Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveContractDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveContractDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveContractDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface

    Dim blnUpdate As Boolean
    blnUpdate = False

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If

    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateContractDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    '   call Business Transaction Object SaveContractDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateContractDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    If blnUpdate Then
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateContractDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
   
    Set ICustomerEmploymentBO_SaveContractDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveContractDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveContractDetailsVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
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

    Set ICustomerEmploymentBO_SaveContractDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveContractDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveCurrentEmploymentAndIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveCurrentEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveCurrentEmploymentAndIncome"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If

    '   call Business Transaction Object SaveCurrentEmploymentAndIncome function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.SaveCurrentEmploymentAndIncome(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set ICustomerEmploymentBO_SaveCurrentEmploymentAndIncome = xmlTempResponseNode

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveCurrentEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveCurrentEmploymentDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveCurrentEmploymentAndIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveCurrentEmploymentDetailsExit


End Function

Private Function ICustomerEmploymentBO_SaveEarnedIncome(ByVal vxmlRequest As IXMLDOMElement, _
                                                        Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveEarnedIncome a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveEarnedIncomeVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveEarnedIncome"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If

    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateEarnedIncome(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
        
    '   call Business Transaction Object SaveEarnedIncome function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.SaveEarnedIncome(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set ICustomerEmploymentBO_SaveEarnedIncome = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveEarnedIncomeExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveEarnedIncomeVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveEarnedIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEarnedIncomeExit

End Function

Private Function ICustomerEmploymentBO_SaveEmployedDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                           Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveEmployedDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveEmployedDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveEmployedDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface

    Dim blnUpdate As Boolean
    blnUpdate = False

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
        
    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateEmployedDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    '   call Business Transaction Object SaveEmployedDetails function
    '=========================================================================================
    ' SR : 26-05-00 : Check whether a record already exists, if so update it else create on
    ' Reason : This method gives automation error before updating the record, when called from
    '          as a part of big transaction
    '==========================================================================================
    If Not CheckEmployedDetailsExists(vxmlRequest) Then
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateEmployedDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    Else
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateEmployedDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
       
    Set ICustomerEmploymentBO_SaveEmployedDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveEmployedDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEmployedDetailsExit

End Function

Private Function CheckEmployedDetailsExists(ByVal vxmlRequest As IXMLDOMNode) As Boolean
'----------------------------------------------------------------------------
'Called from : SaveEmployedDetails
'----------------------------------------------------------------------------
On Error GoTo VbErrHandler
    
    Const strFunctionName As String = "CheckEmployedDetailsExists"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext
    
    Dim strCondition As String
    Dim strCustNumber As String, strCustVersionNumber As String, strEmpSeqNumber As String
    Dim objIAdoAssist As IADOAssist
    Dim objSqlAssist As New SQLAssist
    
    If objContext Is Nothing Then
        Set objIAdoAssist = New ADOAssist
    Else
        Set objIAdoAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    ' Build the condition
    strCustNumber = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustVersionNumber = vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    strEmpSeqNumber = vxmlRequest.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
    
    strCondition = " CUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustNumber) _
                   & " AND CUSTOMERVERSIONNUMBER = " & objSqlAssist.FormatString(strCustVersionNumber) _
                   & " AND EMPLOYMENTSEQUENCENUMBER = " & objSqlAssist.FormatString(strEmpSeqNumber)
    
    CheckEmployedDetailsExists = objIAdoAssist.CheckRecordExists("EMPLOYEDDETAILS", strCondition)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objIAdoAssist = Nothing
    Set objSqlAssist = Nothing
    Set objContext = Nothing
    
    Exit Function
VbErrHandler:
    Dim objErrAssist As New ErrAssist
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Set objErrAssist = Nothing
    Set objIAdoAssist = Nothing
    Set objSqlAssist = Nothing
    Set objContext = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function ICustomerEmploymentBO_SaveEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                             Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveEmploymentDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveEmploymentDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIScriptInterface As IScriptInterface
    Dim botOperationType As BOOPERATIONTYPE
    Dim xmlKeysNode As IXMLDOMNode

    Dim objIClassDef As IomCEClassDef
    Set objIClassDef = New omCEClassDef

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    botOperationType = objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadEmploymentData)

    If vblnDoValidation And ((botOperationType = booCreate) Or (botOperationType = booUpdate)) Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateGeneralEmployment(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    Select Case botOperationType
        Case booCreate:
            Set xmlTempResponseNode = ICustomerEmploymentBO_CreateEmploymentDetails(vxmlRequest)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

            Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
            If Not (xmlKeysNode Is Nothing) Then
                xmlResponseElem.appendChild xmlKeysNode
            End If
        Case booUpdate:
            Set xmlTempResponseNode = ICustomerEmploymentBO_UpdateEmploymentDetails(vxmlRequest)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case booDelete:
            Set xmlTempResponseNode = ICustomerEmploymentBO_DeleteEmploymentDetails(vxmlRequest)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case Else
            ' N/A
    End Select
   
    Set ICustomerEmploymentBO_SaveEmploymentDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set objIScriptInterface = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

SaveEmploymentDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                            Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveNetProfitDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNetProfitDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveNetProfitDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface

    Dim blnUpdate As Boolean
    blnUpdate = False

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If

    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateNetProfitDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
        
    '   call Business Transaction Object SaveNetProfitDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateNetProfitDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    If blnUpdate Then
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateNetProfitDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
   
    Set ICustomerEmploymentBO_SaveNetProfitDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveNetProfitDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveNetProfitDetailsVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
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

    Set ICustomerEmploymentBO_SaveNetProfitDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNetProfitDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveOtherIncome(ByVal vxmlRequest As IXMLDOMElement, _
                                                       Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveOtherIncome a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveOtherIncomeVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveOtherIncome"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIScriptInterface As IScriptInterface
    Dim xmlNodeList As IXMLDOMNodeList
    Dim domDoc As FreeThreadedDOMDocument40

    Dim objIClassDef As IomCEClassDef
    Set objIClassDef = New omCEClassDef

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If vblnDoValidation Then
        ' SR 31/08/00 : Perform validations only for create or update, so call validate
        '               method before calling create/update method
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    'loop around a list of entries
    Set xmlNodeList = vxmlRequest.ownerDocument.getElementsByTagName("UNEARNEDINCOME")
    Set domDoc = objIClassDef.LoadUnearnedIncomeData
    Dim intNodeIndex As Integer
    Dim xmlUnearnedIncomeElem As IXMLDOMElement
    
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set xmlUnearnedIncomeElem = xmlNodeList.Item(intNodeIndex)
        
        'Select Case objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadUnearnedIncomeData)
        Select Case objXmlAssist.DetermineOperationEx(xmlUnearnedIncomeElem, domDoc)
            Case booCreate:
                ' Run validation checks
                Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateOtherIncome(vxmlRequest)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

                Set xmlTempResponseNode = ICustomerEmploymentBO_CreateOtherIncome(xmlUnearnedIncomeElem)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
            Case booUpdate:
                ' Run validation checks
                Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateOtherIncome(vxmlRequest)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
                Set xmlTempResponseNode = ICustomerEmploymentBO_UpdateOtherIncome(xmlUnearnedIncomeElem)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            Case booDelete:
                Set xmlTempResponseNode = ICustomerEmploymentBO_DeleteOtherIncome(xmlUnearnedIncomeElem)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            Case Else
                ' N/A
        End Select
    Next
   
    Set ICustomerEmploymentBO_SaveOtherIncome = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveOtherIncomeExit:

    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveOtherIncomeVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_SaveOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOtherIncomeExit

End Function

Private Function ICustomerEmploymentBO_SaveSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                               Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveSelfEmployedDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveSelfEmployedDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveSelfEmployedDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface

    Dim blnUpdate As Boolean
    blnUpdate = False

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
        
    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateSelfEmployedDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    '   call Business Transaction Object SaveSelfEmployedDetails function
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateSelfEmployedDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    If blnUpdate Then
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateSelfEmployedDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
   
    Set ICustomerEmploymentBO_SaveSelfEmployedDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveSelfEmployedDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveSelfEmployedDetailsVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
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

    Set ICustomerEmploymentBO_SaveSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveSelfEmployedDetailsExit

End Function

Private Function ICustomerEmploymentBO_SaveTaxDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                                      Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveTaxDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveTaxDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_SaveTaxDetails"

    
    Dim objContext As ObjectContext

    
    Dim objErrAssist As ErrAssist

    Dim xmlOut As FreeThreadedDOMDocument40

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objIScriptInterface As IScriptInterface
    
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim xmlGetResponse As IXMLDOMNode
    Dim xmlGetRequest As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Dim blnUpdate As Boolean
    
    'MAR1891 GHun
    Dim lngCount As Integer
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlIncomeSummary As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    'MAR1891 End
    
    Set objErrAssist = New ErrAssist
    Set objContext = GetObjectContext()
    Set objXmlAssist = New XMLAssist
    
    blnUpdate = True
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If

    If vblnDoValidation Then
        ' Run validation checks
        Set xmlTempResponseNode = ICustomerEmploymentBO_ValidateContractDetails(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Run scripting interface
        Set objIScriptInterface = New ScriptInterface
        Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                                TypeName(Me), _
                                                                strFunctionName, _
                                                                vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If

    'MAR1891 GHun When called from CalcAllowableIncome, the existence of the IncomeSummary record is
    ' already known, and multiple IncomeSummary records may be passed in in one call
    If objXmlAssist.GetAttributeValue(vxmlRequest, "RESPONSE", "EXISTENCECHECKED", False) = 1 Then
        Set xmlNodeList = vxmlRequest.selectNodes("INCOMESUMMARY")
        
        lngCount = xmlNodeList.length
        If vxmlRequest.selectNodes("INCOMESUMMARY[@EXISTS='1']").length = lngCount _
            Or vxmlRequest.selectNodes("INCOMESUMMARY[not(@EXISTS)]").length = lngCount Then
            If vxmlRequest.selectSingleNode("INCOMESUMMARY[@EXISTS='1']") Is Nothing Then
                Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateTaxDetails(vxmlRequest)
            Else
                Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateTaxDetails(vxmlRequest)
            End If
        Else
            For Each xmlIncomeSummary In xmlNodeList
                Set xmlTempRequest = vxmlRequest.cloneNode(False)
                xmlTempRequest.appendChild xmlIncomeSummary.cloneNode(True)
                If xmlIncomeSummary.Attributes.getNamedItem("EXISTS") Is Nothing Then
                    Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateTaxDetails(xmlTempRequest)
                Else
                    Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateTaxDetails(xmlTempRequest)
                End If
            Next
        End If
    Else
    'MAR1891 End
        ' See if the record already exists
        Set xmlGetRequest = objXmlAssist.GetNode(vxmlRequest, "INCOMESUMMARY")
        Set xmlGetResponse = objICustomerEmploymentDO.GetTaxDetails(xmlGetRequest)
        
        If blnUpdate = False Then
            ' Call Business Transaction Object SaveTaxDetails function
            Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateTaxDetails(vxmlRequest)
        Else
            Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateTaxDetails(vxmlRequest)
        End If
    End If
    
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerEmploymentBO_SaveTaxDetails = xmlResponseElem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveTaxDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set xmlGetResponse = Nothing
    Set xmlGetRequest = Nothing
    Set objXmlAssist = Nothing
    'MAR1891 GHun
    Set xmlNodeList = Nothing
    Set xmlIncomeSummary = Nothing
    Set xmlTempRequest = Nothing
    'MAR1891 End
    
    Exit Function

SaveTaxDetailsVbErr:

    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        blnUpdate = False
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

    Set ICustomerEmploymentBO_SaveTaxDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveTaxDetailsExit

End Function

Private Function ICustomerEmploymentBO_UpdateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateAccountantDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountantDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_UpdateAccountantDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateAccountantDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerEmploymentBO_UpdateAccountantDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateAccountantDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateAccountantDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_UpdateAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateAccountantDetailsExit

End Function

Private Function ICustomerEmploymentBO_UpdateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateEmploymentDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_UpdateEmploymentDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateEmploymentDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerEmploymentBO_UpdateEmploymentDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_UpdateEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentBO_UpdateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateOtherIncome a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherIncomeVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentBO_UpdateOtherIncome"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateOtherIncome(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerEmploymentBO_UpdateOtherIncome = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateOtherIncomeExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateOtherIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_UpdateOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateOtherIncomeExit

End Function

Private Function ICustomerEmploymentBO_ValidateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateAccountantDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateAccountantDetails = xmlResponseElem.cloneNode(True)
    
ValidateAccountantDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateAccountantDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateAccountantDetailsExit

End Function

Private Function ICustomerEmploymentBO_ValidateContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateContractDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNetProfitXML As IXMLDOMNode
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Validate associated NETPROFIT record
    '
    Set xmlNetProfitXML = vxmlRequest.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        Set xmlResponseElem = ICustomerEmploymentBO_ValidateNetProfitDetails(xmlNetProfitXML)
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateContractDetails = xmlResponseElem.cloneNode(True)
    
ValidateContractDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNetProfitXML = Nothing
    
    Exit Function

ValidateContractDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateContractDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateContractDetailsExit

End Function

Private Function ICustomerEmploymentBO_ValidateEarnedIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateEarnedIncome"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objComboDO As New ComboDO
    Dim strIncomeTypeID As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strIncomeTypeID = objXmlAssist.GetTagValue(vxmlRequest, "EARNEDINCOMETYPE")

    '
    ' Earned income amount must be specified if payment frequency is specified
    ' SYS2084 SA First item in the strAdditionalOptions is additional text not
    ' replacement text, according to the Throwerror function. Added empty string
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlRequest, "PAYMENTFREQUENCYTYPE")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlRequest, "EARNEDINCOMEAMOUNT")) = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 244, _
            "", objComboDO.GetComboText("IncomeType", strIncomeTypeID)
    End If

    '
    ' Payment frequency must be specified if earned income amount is specified
    ' SYS2084 SA First item in the strAdditionalOptions is additional text not
    ' replacement text, according to the Throwerror function.
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlRequest, "EARNEDINCOMEAMOUNTTYPE")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlRequest, "PAYMENTFREQUENCYTYPE")) = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 245, _
            "", objComboDO.GetComboText("IncomeType", strIncomeTypeID)
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateEarnedIncome = xmlResponseElem.cloneNode(True)
    
ValidateEarnedIncomeExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateEarnedIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateEarnedIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateEarnedIncomeExit

End Function

Private Function ICustomerEmploymentBO_ValidateEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateEmployedDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlEarnedIncomeXML As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Job title must not be null
    'If Trim$(objXmlAssist.GetTagValue(vxmlRequest, "JOBTITLE")) = "" Then
    '    objErrAssist.ThrowError TypeName(Me), strFunctionName, 243
    'End If

    '
    ' Validate associated EARNEDINCOME records
    '
    Set xmlEarnedIncomeXML = vxmlRequest.selectSingleNode(".//EARNEDINCOMELIST")
    If Not (xmlEarnedIncomeXML Is Nothing) Then
        For Each xmlNode In xmlEarnedIncomeXML.childNodes
            Set xmlResponseElem = ICustomerEmploymentBO_ValidateEarnedIncome(xmlNode)
        Next xmlNode
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateEmployedDetails = xmlResponseElem.cloneNode(True)
    
ValidateEmployedDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlEarnedIncomeXML = Nothing
    Set xmlNode = Nothing
    
    Exit Function

ValidateEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateEmployedDetailsExit

End Function

Private Function ICustomerEmploymentBO_ValidateGeneralEmployment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateGeneralEmploymentVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateGeneralEmployment"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDetailsXML As IXMLDOMNode
    Dim strEmploymentStatus As String
    Dim objComboDO As New ComboDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '--------------------------------------------------------------------------------
    ' Mandatory field removed as part of HomeZone changes (ref. Home Zone Matrix.xls)
    '================================================================================
    ' Employment type must not be null
    'If Trim$(objXmlAssist.GetTagValue(vxmlRequest, "EMPLOYMENTTYPE")) = "" Then
    '    objErrAssist.ThrowError TypeName(Me), strFunctionName, 241
    'End If
    '--------------------------------------------------------------------------------

    ' Employment status must not be null
    strEmploymentStatus = Trim$(objXmlAssist.GetTagValue(vxmlRequest, "EMPLOYMENTSTATUS"))
    If strEmploymentStatus = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 242
    End If

    '
    ' Validate the associated EMPLOYEDDETAILS, CONTRACTDETAILS or SELFEMPLOYEDDETAILS record
    '
    If strEmploymentStatus <> "" Then
        strEmploymentStatus = objComboDO.GetFirstComboValidation("EmploymentStatus", strEmploymentStatus)
    End If

    Select Case strEmploymentStatus
        Case "E" ' EMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlRequest.selectSingleNode(".//EMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                Set xmlResponseElem = ICustomerEmploymentBO_ValidateEmployedDetails(xmlDetailsXML)
            End If
        Case "S" ' SELFEMPLOYEDDETAILS
            Set xmlDetailsXML = vxmlRequest.selectSingleNode(".//SELFEMPLOYEDDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                Set xmlResponseElem = ICustomerEmploymentBO_ValidateSelfEmployedDetails(xmlDetailsXML)
            End If
        Case "C" ' CONTRACTDETAILS
            Set xmlDetailsXML = vxmlRequest.selectSingleNode(".//CONTRACTDETAILS")
            If Not (xmlDetailsXML Is Nothing) Then
                Set xmlResponseElem = ICustomerEmploymentBO_ValidateContractDetails(xmlDetailsXML)
            End If
        Case Else
    End Select

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateGeneralEmployment = xmlResponseElem.cloneNode(True)
    
ValidateGeneralEmploymentExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlDetailsXML = Nothing
    Set objComboDO = Nothing
    
    Exit Function

ValidateGeneralEmploymentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateGeneralEmployment = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateGeneralEmploymentExit

End Function

Private Function ICustomerEmploymentBO_ValidateNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateNetProfitDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim nYear1 As Integer, _
        nYear2 As Integer, _
        nYear3 As Integer
    Dim nYear1Amount As Long, _
        nYear2Amount As Long, _
        nYear3Amount As Long
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    nYear1 = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR1"))
    nYear2 = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR2"))
    nYear3 = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR3"))
    nYear1Amount = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR1AMOUNT"))
    nYear2Amount = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR2AMOUNT"))
    nYear3Amount = Val(objXmlAssist.GetTagValue(vxmlRequest, "YEAR3AMOUNT"))

    ' If a year is specified, the amount for that year must not be null
    If ((nYear1 > 0) And (nYear1Amount = 0)) Or _
       ((nYear2 > 0) And (nYear2Amount = 0)) Or _
       ((nYear3 > 0) And (nYear3Amount = 0)) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 247
    End If

    ' If a year amount is specified, the corresponding year must not be null
    If ((nYear1 = 0) And (nYear1Amount > 0)) Or _
       ((nYear2 = 0) And (nYear2Amount > 0)) Or _
       ((nYear3 = 0) And (nYear3Amount > 0)) Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 248
    End If

    If (nYear1 > Year(Now)) Or (nYear2 > Year(Now)) Or (nYear3 > Year(Now)) Then
        ' No specified years cannot be in the future
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 279
    End If

    If ((nYear1 = nYear2) And (nYear1 > 0)) Or _
       ((nYear2 = nYear3) And (nYear2 > 0)) Or _
       ((nYear1 = nYear3) And (nYear3 > 0)) Then
        ' No specified years can be equal
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 278
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateNetProfitDetails = xmlResponseElem.cloneNode(True)
    
ValidateNetProfitDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateNetProfitDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateNetProfitDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateNetProfitDetailsExit

End Function

Private Function ICustomerEmploymentBO_ValidateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateOtherIncome"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Description must not be null
    If Trim$(objXmlAssist.GetTagValue(vxmlRequest, "UNEARNEDINCOMETYPE")) = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 249
    End If

    ' Amount must not be null
    If Trim$(objXmlAssist.GetTagValue(vxmlRequest, "UNEARNEDINCOMEAMOUNT")) = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 252
    End If

    ' Payment frequency must not be null
    If Trim$(objXmlAssist.GetTagValue(vxmlRequest, "PAYMENTFREQUENCY")) = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 253
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateOtherIncome = xmlResponseElem.cloneNode(True)
    
ValidateOtherIncomeExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateOtherIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateOtherIncomeExit

End Function

Private Function ICustomerEmploymentBO_ValidateSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateSelfEmployedDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNetProfitXML As IXMLDOMNode
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Validate associated NETPROFIT record
    '
    Set xmlNetProfitXML = vxmlRequest.selectSingleNode(".//NETPROFIT")
    If Not (xmlNetProfitXML Is Nothing) Then
        Set xmlResponseElem = ICustomerEmploymentBO_ValidateNetProfitDetails(xmlNetProfitXML)
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateSelfEmployedDetails = xmlResponseElem.cloneNode(True)
    
ValidateSelfEmployedDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNetProfitXML = Nothing
    
    Exit Function

ValidateSelfEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateSelfEmployedDetailsExit

End Function

Private Function ICustomerEmploymentBO_ValidateTaxDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentBO_ValidateTaxDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set ICustomerEmploymentBO_ValidateTaxDetails = xmlResponseElem.cloneNode(True)
    
ValidateTaxDetailsExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateTaxDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentBO_ValidateTaxDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateTaxDetailsExit

End Function

Public Function DeleteEmploymentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteEmploymentDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEmploymentDetailsVbErr

    Const strFunctionName As String = "DeleteEmploymentDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_DeleteEmploymentDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteEmploymentDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteEmploymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteEmploymentDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEmploymentDetailsExit

End Function

Public Function DeleteOtherIncome(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteOtherIncome a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOtherIncomeVbErr

    Const strFunctionName As String = "DeleteOtherIncome"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_DeleteOtherIncome(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteOtherIncome = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteOtherIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteOtherIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteOtherIncome = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteOtherIncomeExit

End Function

Public Function FindEarnedIncomeList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindEarnedIncomeListVbErr

    Const strFunctionName As String = "FindEarnedIncomeList"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_FindEarnedIncomeList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindEarnedIncomeList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindEarnedIncomeListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindEarnedIncomeListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindEarnedIncomeList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindEarnedIncomeListExit

End Function

Public Function FindEmploymentAndIncomeDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindEmploymentAndIncomeDetailsVbErr

    Const strFunctionName As String = "FindEmploymentAndIncomeDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_FindEmploymentAndIncomeDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindEmploymentAndIncomeDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindEmploymentAndIncomeDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindEmploymentAndIncomeDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindEmploymentAndIncomeDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindEmploymentAndIncomeDetailsExit

End Function

Public Function FindOtherIncomeList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindOtherIncomeListVbErr

    Const strFunctionName As String = "FindOtherIncomeList"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_FindOtherIncomeList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindOtherIncomeList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindOtherIncomeListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindOtherIncomeListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindOtherIncomeList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindOtherIncomeListExit

End Function

Public Function GetAccountantDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountantDetailsVbErr

    Const strFunctionName As String = "GetAccountantDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetAccountantDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetAccountantDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetAccountantDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAccountantDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAccountantDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetAccountantDetailsExit

End Function

Public Function GetEmploymentAccountant(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentAccountantVbErr

    Const strFunctionName As String = "GetEmploymentAccountant"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    Dim strTmpResponse As String
    
  
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
                                    
    ' Delegate to DomDocument based method GetAccountantGUID
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetAccountantGUID(xmlIn.documentElement)
    objErrAssist.CheckXMLResponse xmlTempResponseNode.xml, True
    
    ' now must build an Accountant node
    '
          
    Set xmlParentNode = xmlIn.createElement("ACCOUNTANT")
    xmlParentNode.appendChild xmlTempResponseNode.selectSingleNode("ACCOUNTANTGUID")
    '
    ' Calling CustomerEmploymentBO.GetAccountantDetails
    '
    strTmpResponse = GetAccountantDetails(xmlParentNode.xml)
    objErrAssist.CheckXMLResponse strTmpResponse, True
            
    GetEmploymentAccountant = strTmpResponse
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEmploymentAccountantExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlParentNode = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEmploymentAccountantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetEmploymentAccountant = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentAccountantExit

End Function

Public Function GetContractDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetContractDetailsVbErr

    Const strFunctionName As String = "GetContractDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetContractDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetContractDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetContractDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetContractDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetContractDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetContractDetailsExit

End Function

Public Function GetEmployedDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEmployedDetailsVbErr

    Const strFunctionName As String = "GetEmployedDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetEmployedDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetEmployedDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEmployedDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetEmployedDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmployedDetailsExit

End Function

Public Function GetEmploymentAndSelfEmployedDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentAndSelfEmployedDetailsVbErr

    Const strFunctionName As String = "GetEmploymentAndSelfEmployedDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetEmploymentAndSelfEmployedDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetEmploymentAndSelfEmployedDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEmploymentAndSelfEmployedDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEmploymentAndSelfEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetEmploymentAndSelfEmployedDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentAndSelfEmployedDetailsExit

End Function

Public Function GetEmploymentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentDetailsVbErr

    Const strFunctionName As String = "GetEmploymentDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetEmploymentDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetEmploymentDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEmploymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetEmploymentDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentDetailsExit

End Function

Public Function GetCurrentEmploymentAndIncome(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEmploymentVbErr

    Const strFunctionName As String = "GetCurrentEmploymentAndIncome"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetCurrentEmploymentAndIncome(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetCurrentEmploymentAndIncome = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEmploymentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEmploymentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentEmploymentAndIncome = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEmploymentExit

End Function

Public Function GetNetProfitDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetNetProfitDetailsVbErr

    Const strFunctionName As String = "GetNetProfitDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetNetProfitDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetNetProfitDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetNetProfitDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNetProfitDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNetProfitDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetNetProfitDetailsExit

End Function

Public Function GetSelfEmployedDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetSelfEmployedDetailsVbErr

    Const strFunctionName As String = "GetSelfEmployedDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetSelfEmployedDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetSelfEmployedDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetSelfEmployedDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetSelfEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetSelfEmployedDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetSelfEmployedDetailsExit

End Function

Public Function SaveAccountantDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveAccountantDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveAccountantDetailsVbErr

    Const strFunctionName As String = "SaveAccountantDetails"

    
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
    Dim xmlKeysNode As IXMLDOMNode
    
    Dim strDirectoryAddressIndicator As String, _
        xmlDirectoryGUIDNode As IXMLDOMNode
    
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
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//EMPLOYMENT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveAccountantDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
    
    SaveAccountantDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveAccountantDetailsExit:
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set xmlKeysNode = Nothing

    Exit Function

SaveAccountantDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveAccountantDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveAccountantDetailsExit
    
End Function

Public Function SaveContractDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveContractDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveContractDetailsVbErr

    Const strFunctionName As String = "SaveContractDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveContractDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveContractDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveContractDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveContractDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveContractDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveContractDetailsExit
    
End Function

Public Function SaveEarnedIncome(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveEarnedIncome a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveEarnedIncomeVbErr

    Const strFunctionName As String = "SaveEarnedIncome"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveEarnedIncome(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveEarnedIncome = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveEarnedIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveEarnedIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveEarnedIncome = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEarnedIncomeExit
    
End Function

Public Function SaveEmployedDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveEmployedDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveEmployedDetailsVbErr

    Const strFunctionName As String = "SaveEmployedDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveEmployedDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveEmployedDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveEmployedDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveEmployedDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEmployedDetailsExit
    
End Function

Public Function SaveCurrentEmploymentAndIncome(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveEmploymentVbErr

    Const strFunctionName As String = "SaveCurrentEmploymentAndIncome"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveCurrentEmploymentAndIncome(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveCurrentEmploymentAndIncome = xmlTempResponseNode.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveEmploymentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveEmploymentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveCurrentEmploymentAndIncome = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEmploymentExit
    
End Function

Public Function SaveEmploymentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveEmploymentDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveEmploymentDetailsVbErr

    Const strFunctionName As String = "SaveEmploymentDetails"

    
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
    Dim xmlKeysNode As IXMLDOMNode
    Dim strDirectoryAddressIndicator As String, _
        xmlDirectoryGUIDNode As IXMLDOMNode
    
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
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//EMPLOYMENT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And (strDirectoryAddressIndicator = "0") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveEmploymentDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
    
    SaveEmploymentDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveEmploymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

SaveEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveEmploymentDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEmploymentDetailsExit
    
End Function

Public Function SaveNetProfitDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveNetProfitDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveNetProfitDetailsVbErr

    Const strFunctionName As String = "SaveNetProfitDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveNetProfitDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveNetProfitDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveNetProfitDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveNetProfitDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveNetProfitDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNetProfitDetailsExit
    
End Function

Public Function SaveOtherIncome(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveOtherIncome a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveOtherIncomeVbErr

    Const strFunctionName As String = "SaveOtherIncome"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveOtherIncome(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveOtherIncome = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveOtherIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveOtherIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveOtherIncome = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOtherIncomeExit
    
End Function

Public Function SaveSelfEmployedDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveSelfEmployedDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveSelfEmployedDetailsVbErr

    Const strFunctionName As String = "SaveSelfEmployedDetails"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveSelfEmployedDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveSelfEmployedDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveSelfEmployedDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveSelfEmployedDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveSelfEmployedDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveSelfEmployedDetailsExit
    
End Function

Public Function SaveTaxDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveTaxDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveTaxDetailsVbErr

    Const strFunctionName As String = "SaveTaxDetails"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set objErrAssist = New ErrAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objXmlAssist = New XMLAssist
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ICustomerEmploymentBO_SaveTaxDetails(xmlIn.documentElement, vblnDoValidation:=True)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveTaxDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveTaxDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveTaxDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveTaxDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveTaxDetailsExit
    
End Function

Public Function FindEmploymentList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table 'Employment'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindEmploymentList"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerEmploymentBO_FindEmploymentList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindEmploymentList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

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

    FindEmploymentList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function CopyEmploymentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a new customer version based on an existing one
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       CreateNewCustomerVersion  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of customer key data
'               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------

On Error GoTo CopyEmploymentDetailsVbErr
    
Const strFunctionName As String = "CopyEmploymentDetails"

Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
Dim objErrAssist As New ErrAssist


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    'Initialise response
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    xmlIn.loadXML vstrXMLRequest
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerEmploymentBO_CopyEmploymentDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CopyEmploymentDetails = xmlTempResponseNode.xml
    objContext.SetComplete

CopyEmploymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    
    Exit Function

CopyEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CopyEmploymentDetails = objErrAssist.CreateErrorResponse

    objContext.SetAbort

    ' go to clean-up section
    Resume CopyEmploymentDetailsExit

End Function

Public Function GetIncomeSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetIncomeSummaryVbErr

    Const strFunctionName As String = "GetIncomeSummary"

    
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
    Set xmlTempResponseNode = ICustomerEmploymentBO_GetIncomeSummary(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetIncomeSummary = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetIncomeSummaryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetIncomeSummaryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetIncomeSummary = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetIncomeSummaryExit

End Function

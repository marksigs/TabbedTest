VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omBACSInterfaceBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omBACSInterfaceBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   BACS Interface Business Object

'Dependencies:

'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SG     11/03/02    SYS4202 Created
'SG     15/05/02    MSMS0084 Amended ConstructPaymentRecord to check output string lengths
'SG     30/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------------------

Option Explicit

'To hold the BACS file record
Dim mstrDestSortCode As String
Dim mstrDestAccountNo As String
Dim mstrAccountType As String
Dim mstrTransactionCode As String
Dim mstrOriginatorSortCode As String
Dim mstrOriginatorAccNo As String
Dim mstrStandingOrderRef As String
Dim mstrAmount As String
Dim mstrOriginatorName As String
Dim mstrOriginatorRef As String
Dim mstrDestAccName As String

Implements ObjectControl

Private gobjContext As ObjectContext

'Private gobjTxBO As IOm4Ph2TxBO
'Private gobjDO As IOm4Ph2DO


'-------------------------------------------------------------------------------------------------------
'                      BMids History:
'
'Prog   Date        AQR         Description
'PSC    08/10/2002  BMIDS00595  Amend to process incentive releases
'PSC    18/11/2002  BMIDS00984  Amend Validation to return errors if it fails
'INR    14/01/2003  BMIDS646 Call Compensating Resource Manager to write to text files within a
'                   transaction, allowing roll back if the transaction fails after the text write.
'
'------------------------------------------------------------------------------------------------------


Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function OmRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete
    
    ' create REQUEST & RESPONSE DOMDocuments
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.loadXML vstrXMLRequest
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                ' copy attributes from REQUEST node to OPERATION node
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If

    ' if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    OmRequest = xmlOutDoc.xml
    Debug.Print OmRequest
    
OmRequestExit:
    
    ' release XML DOM objects
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmRequestVbErr:
    
    ' create ERROR response block from Err object
    OmRequest = errCreateErrorResponse()
    Debug.Print OmRequest
    
    ' perform clean up
    GoTo OmRequestExit

End Function

Private Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Request broker, private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       raises oeNotImplemented if OPERATION value not in list of known functions.
'------------------------------------------------------------------------------------------

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "BUILDBACSINTERFACE"
            BuildBACSInterface vxmlRequestNode, vxmlResponseNode
         
        Case Else
            errThrowError _
                "DoRequest", _
                oeNotImplemented, _
                vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub BuildBACSInterface( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'       Writes relevant payment data to a BACS file
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

    On Error GoTo BuildBACSInterfaceExit
    
    Const strFunctionName As String = "BuildBACSInterface"
        
    'Some objects to play with
    Dim xmlPaymentRecordNode As IXMLDOMNode
    Dim xmlLoanCompPayNode As IXMLDOMNode
    Dim xmlLoanCompNode As IXMLDOMNode
    Dim xmlMortProdNode As IXMLDOMNode
    Dim xmlBACSNode As IXMLDOMNode
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Dim xmlDisb As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objMortSQ As MortgageSubQuoteBO
    Dim xmlPayment As IXMLDOMNode
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim strIRValueId As String
    Dim strPaymentType As String
    Dim strResponse As String
    Dim strAppFFNo As String
    Dim strAppNo As String
    ' PSC 08/10/2002 BMIDS00595 - End
    
    Dim strReason As String         ' PSC 18/11/2002 BMIDS00984
      
    '1 Get Payment and Payee Details
    '---------------------------------
    Set xmlPaymentRecordNode = GetPaymentAndPayeeDetails(vxmlRequestNode)
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    ' Get the value id for an incentive release
    strIRValueId = GetFirstComboValueId("PaymentType", "C")
    Set xmlDisb = xmlGetMandatoryNode(xmlPaymentRecordNode, "DISBURSEMENTPAYMENT")
    strPaymentType = xmlGetMandatoryAttributeText(xmlDisb, "PAYMENTTYPE")
    ' PSC 08/10/2002 BMIDS00595 - End

    
    'Isolate the Loan Component
    Set xmlLoanCompPayNode = xmlPaymentRecordNode.selectSingleNode(".//LOANCOMPONENTPAYMENT")
    
    ' PSC 08/10/2002 BMIDS00595 - Incentive Release will not have a loan component payment
    If xmlLoanCompPayNode Is Nothing And strPaymentType <> strIRValueId Then
        'Node not present, no need to write to continue
        GoTo BuildBACSInterfaceExit:
    Else
        'Node present - validate it
        ' PSC 18/11/2002 BMIDS00984 - Start
        If ValidatePaymentRecordData(xmlPaymentRecordNode, strReason) = False Then
            'Data not valid
            errThrowError strFunctionName, 7010, strReason
            ' PSC 18/11/2002 BMIDS00984 - End
        End If
    End If
                                                    
    '2. Get Loan Component Data
    '---------------------------
    ' PSC 08/10/2002 BMIDS00595 - Start
    If strPaymentType = strIRValueId Then
        
        strAppNo = xmlGetMandatoryAttributeText(xmlDisb, "APPLICATIONNUMBER")
        Set xmlPayment = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
        strAppFFNo = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONFACTFINDNUMBER")
    
        ' Get the accepted list of components
        Set xmlTempRequest = xmlPayment.ownerDocument.createElement("REQUEST")
        Set xmlRoot = xmlPayment.ownerDocument.createElement("FINDACTIVELOANCOMPONENTS")
        xmlTempRequest.appendChild xmlRoot
        Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONNUMBER")
        xmlRoot.appendChild xmlNode
        xmlNode.Text = strAppNo
        Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
        xmlRoot.appendChild xmlNode
        xmlNode.Text = strAppFFNo

        Set objMortSQ = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
        
        strResponse = objMortSQ.FindActiveLoanComponentsList(xmlTempRequest.xml)
        Set xmlResponse = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponseNode xmlResponse.documentElement, , True
        
        ' Set the loan component to the first one as lender details will be the same
        ' fo all of them
        Set xmlLoanCompNode = xmlGetMandatoryNode(xmlResponse, ".//LOANCOMPONENT")
    Else
        Set xmlLoanCompNode = GetLoanComponentData(xmlLoanCompPayNode)
    End If
    ' PSC 08/10/2002 BMIDS00595 - End
    
    If xmlLoanCompNode Is Nothing Then
        'Node not present, no need to write to continue
        GoTo BuildBACSInterfaceExit:
    End If
    
    '3. Get Mortgage Product Data
    '-----------------------------
    Set xmlMortProdNode = GetMortgageProductData(xmlLoanCompNode)
    If xmlMortProdNode Is Nothing Then
        'Node not present, no need to write to continue
        GoTo BuildBACSInterfaceExit:
    End If
    
    '4. Get BACS Data
    '-----------------
    Set xmlBACSNode = GetBACSData(xmlMortProdNode)
    If xmlBACSNode Is Nothing Then
        'Node not present, no need to write to continue
        GoTo BuildBACSInterfaceExit:
    Else
        'SG 30/05/02 SYS4767
        'SG 15/05/02 MSMS0083 - validate node
        If ValidateBACSData(xmlBACSNode) = False Then
            'Data not valid
            GoTo BuildBACSInterfaceExit:
        End If
    End If
           
    '5. Collate payment record fields
    '---------------------------------
    'From Payment Record data...
    mstrDestSortCode = xmlPaymentRecordNode.selectSingleNode(".//PAYEEHISTORY").Attributes.getNamedItem("BANKSORTCODE").Text
    mstrDestAccountNo = xmlPaymentRecordNode.selectSingleNode(".//PAYEEHISTORY").Attributes.getNamedItem("ACCOUNTNUMBER").Text
    mstrOriginatorRef = xmlPaymentRecordNode.Attributes.getNamedItem("APPLICATIONNUMBER").Text
    mstrDestAccName = xmlPaymentRecordNode.selectSingleNode(".//THIRDPARTY").Attributes.getNamedItem("COMPANYNAME").Text
    mstrAmount = xmlPaymentRecordNode.Attributes.getNamedItem("AMOUNT").Text
    
    'From BACS Data...
    mstrOriginatorAccNo = xmlBACSNode.selectSingleNode(".//BACSACCOUNTNUMBER").Text
    mstrOriginatorName = xmlBACSNode.selectSingleNode(".//LENDERNAME").Text
    mstrOriginatorSortCode = xmlBACSNode.selectSingleNode(".//BACSSORTCODE").Text
    
    'Strings...
    mstrAccountType = "0"
    mstrTransactionCode = "99"
    mstrStandingOrderRef = Space(4)
                    
    '6. Output to file
    '------------------
    OutputToFile ConstructPaymentRecord()
             
BuildBACSInterfaceExit:
    
    If Not xmlPaymentRecordNode Is Nothing Then Set xmlPaymentRecordNode = Nothing
    If Not xmlLoanCompPayNode Is Nothing Then Set xmlLoanCompPayNode = Nothing
    If Not xmlLoanCompNode Is Nothing Then Set xmlLoanCompNode = Nothing
    If Not xmlMortProdNode Is Nothing Then Set xmlMortProdNode = Nothing
    If Not xmlBACSNode Is Nothing Then Set xmlBACSNode = Nothing
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Set xmlDisb = Nothing
    Set xmlTempRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set objMortSQ = Nothing
    Set xmlResponse = Nothing
    Set xmlPayment = Nothing
    ' PSC 08/10/2002 BMIDS00595 - End
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Function ValidatePaymentRecordData(ByVal xmlPaymentRecordNode As IXMLDOMNode, rstrReason) As Boolean

    On Error GoTo ValidatePaymentRecordDataExit
    
    Const strFunctionName As String = "ValidatePaymentRecordData"
    
    Dim xmlNode As IXMLDOMNode          ' PSC 18/11/2002 BMIDS00984

    ValidatePaymentRecordData = True
    
    'Sort codes
    ' PSC 18/11/2002 BMIDS00984 - Start
    Set xmlNode = xmlPaymentRecordNode.selectSingleNode(".//PAYEEHISTORY/@BANKSORTCODE")
    
    If xmlNode Is Nothing Then
        ValidatePaymentRecordData = False
        rstrReason = "Missing Bank Sort Code"
        GoTo ValidatePaymentRecordDataExit:
    ElseIf ValidateSortCode(xmlNode.Text) = False Then
        ValidatePaymentRecordData = False
        rstrReason = "Invalid Bank Sort Code"
        GoTo ValidatePaymentRecordDataExit:
    End If
       
    'Account number
    'SG 30/05/02 SYS4767
    'SG 15/05/02 MSMS0083 - validation code moved to ValidateAccountNo.
    Set xmlNode = xmlPaymentRecordNode.selectSingleNode(".//PAYEEHISTORY/@ACCOUNTNUMBER")
    
    If xmlNode Is Nothing Then
        ValidatePaymentRecordData = False
        rstrReason = "Missing Account Number"
        GoTo ValidatePaymentRecordDataExit:
    ElseIf ValidateAccountNo(xmlNode.Text) = False Then
        'Account number is to big
        ValidatePaymentRecordData = False
        rstrReason = "Invalid Account Number"
        GoTo ValidatePaymentRecordDataExit:
    End If

    'Amount
    Set xmlNode = xmlGetAttributeNode(xmlPaymentRecordNode, "AMOUNT")
    
    If xmlNode Is Nothing Then
        ValidatePaymentRecordData = False
        rstrReason = "Missing Amount"
        GoTo ValidatePaymentRecordDataExit:
    ElseIf Len(xmlNode.Text) > 9 Then
        'Amount is too big
        ValidatePaymentRecordData = False
        rstrReason = "Amount is too big"
        GoTo ValidatePaymentRecordDataExit:
    ElseIf Val(xmlNode.Text) = 0 Then
        'Amount is  zero
        ValidatePaymentRecordData = False
        rstrReason = "Amount is zero"
        GoTo ValidatePaymentRecordDataExit:
    End If
    ' PSC 18/11/2002 BMIDS00984 - End
    
ValidatePaymentRecordDataExit:
    
    Set xmlNode = Nothing               ' PSC 18/11/2002 BMIDS00984
    
    errCheckError strFunctionName, TypeName(Me)
    
End Function


Private Function ValidateBACSData(ByVal xmlBACSNode As IXMLDOMNode) As Boolean
'SG 15/05/02 MSMS0083 Function added

    On Error GoTo ValidateBACSDataExit
    
    Const strFunctionName As String = "ValidateBACSData"

    ValidateBACSData = True

    'Sort codes
    If ValidateSortCode(xmlBACSNode.selectSingleNode(".//BACSSORTCODE").Text) = False Then
        ValidateBACSData = True
        GoTo ValidateBACSDataExit:
    End If
    
    'Account number
    If ValidateAccountNo(xmlBACSNode.selectSingleNode(".//BACSACCOUNTNUMBER").Text) = False Then
        ValidateBACSData = False
        GoTo ValidateBACSDataExit:
    End If

ValidateBACSDataExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function ValidateAccountNo(ByVal strAccountNo As String) As Boolean
'SG 15/05/02 MSMS0083 Function added

    On Error GoTo ValidateAccountNoExit:
    
    Const strFunctionName As String = "ValidateAccountNo"

    ValidateAccountNo = True

    If Len(Trim(strAccountNo)) <> 8 Then
        'Account number not 8 characters
        ValidateAccountNo = False
        GoTo ValidateAccountNoExit:
    End If

    If Not IsNumeric(strAccountNo) Then
        'Account number contains invalid characters
        ValidateAccountNo = False
        GoTo ValidateAccountNoExit:
    End If

ValidateAccountNoExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function ValidateSortCode(ByVal strSortCode As String) As Boolean
'Purpose:   SortCode can either be "123456" or "12-34-56"
'           Check that sortcode is in correct format and contains no invalid characters

    On Error GoTo ValidateSortCodeExit
    
    Const strFunctionName As String = "ValidateSortCode"

    Dim lngPos1 As Long
    Dim lngPos2 As Long
    
    ValidateSortCode = True
    
    'SG 30/05/02 SYS4767
    'SG 15/05/02 MSMS0083 - check for empty string
    If Len(Trim(strSortCode)) = 0 Then
        'Invalid
        ValidateSortCode = False
        GoTo ValidateSortCodeExit
    End If
    
    If Len(strSortCode) <> 6 And Len(strSortCode) <> 8 Then
        'Invalid
        ValidateSortCode = False
        GoTo ValidateSortCodeExit
    End If
    
    If Len(strSortCode) = 6 Then
        'Must be of format "123456"
        If IsNumeric(strSortCode) = False Then
            'One of the characters is non-numeric - Invalid
            ValidateSortCode = False
            GoTo ValidateSortCodeExit
        End If
    End If
    
    If Len(strSortCode) = 8 Then
        'Has 8 characters, must be of format "12-34-56"
        If Mid(strSortCode, 3, 1) <> "-" Then
            'First "-" is in the wrong place - Invalid
            ValidateSortCode = False
        Else
            If Mid(strSortCode, 6, 1) <> "-" Then
                'First "-" is in the wrong place - Invalid
                ValidateSortCode = False
            Else
                'Check constituent parts
                If IsNumeric(Mid(strSortCode, 1, 2)) = False Then
                    ValidateSortCode = False
                Else
                    If IsNumeric(Mid(strSortCode, 4, 2)) = False Then
                        ValidateSortCode = False
                    Else
                        If IsNumeric(Mid(strSortCode, 6, 2)) = False Then
                            ValidateSortCode = False
                        End If
                    End If
                End If
            End If
        End If
    End If

ValidateSortCodeExit:
    
    errCheckError strFunctionName, TypeName(Me)


End Function

Private Function GetPaymentAndPayeeDetails(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo GetPaymentAndPayeeDetailsErr
    
    Const strFunctionName As String = "GetPaymentAndPayeeDetails"
        
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
        
    Dim strTempResponse As String
    
    Dim objPaymentProcessingBO As Object
    
    'Create a temp request doc
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    'Create a temp response doc
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.async = False
       
    'Copy the Request node passed in to the temp request, changing the operation attribute
    Set xmlTempRequestNode = xmlTempRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlTempRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetFullPaymentDetails"
      
    'Create a PAYMENTRECORD child in temp request and give it some attributes
    Set xmlNode = xmlTempRequestDoc.createElement("PAYMENTRECORD")
    xmlTempRequestDoc.selectSingleNode("REQUEST").appendChild xmlNode
    
    'SG 30/05/02 SYS4767 START
    'xmlCopyMandatoryAttribute vxmlRequestNode, xmlNode, "APPLICATIONNUMBER"
    'xmlCopyMandatoryAttribute vxmlRequestNode, xmlNode, "PAYMENTSEQUENCENUMBER"
        
    If vxmlRequestNode.nodeName = "REQUEST" Then
        xmlCopyMandatoryAttribute vxmlRequestNode.selectSingleNode("PAYMENTRECORD"), xmlNode, "APPLICATIONNUMBER"
        xmlCopyMandatoryAttribute vxmlRequestNode.selectSingleNode("PAYMENTRECORD"), xmlNode, "PAYMENTSEQUENCENUMBER"
    Else
        'Assume it's just PAYMENTRECORD node.
        xmlCopyMandatoryAttribute vxmlRequestNode, xmlNode, "APPLICATIONNUMBER"
        xmlCopyMandatoryAttribute vxmlRequestNode, xmlNode, "PAYMENTSEQUENCENUMBER"
    End If
    'SG 30/05/02 SYS4767 END
    
    'Call Payment Processing component
    Set objPaymentProcessingBO = gobjContext.CreateInstance("omPayProc.PaymentProcessingBO")
    
    strTempResponse = objPaymentProcessingBO.omPayProcRequest(xmlTempRequestNode.xml)
    xmlTempResponseDoc.loadXML strTempResponse
    
    errCheckXMLResponseNode xmlTempResponseDoc.selectSingleNode("RESPONSE"), , True

    Set GetPaymentAndPayeeDetails = xmlTempResponseDoc.selectSingleNode(".//PAYMENTRECORD")

    gobjContext.SetComplete

GetPaymentAndPayeeDetailsExit:

    If Not xmlTempRequestDoc Is Nothing Then Set xmlTempRequestDoc = Nothing
    If Not xmlTempResponseDoc Is Nothing Then Set xmlTempResponseDoc = Nothing
    If Not xmlTempRequestNode Is Nothing Then Set xmlTempRequestNode = Nothing
    If Not xmlTempResponseNode Is Nothing Then Set xmlTempResponseNode = Nothing
    If Not xmlNode Is Nothing Then Set xmlNode = Nothing
    If Not objPaymentProcessingBO Is Nothing Then Set objPaymentProcessingBO = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Function
    
GetPaymentAndPayeeDetailsErr:
    gobjContext.SetAbort
    GoTo GetPaymentAndPayeeDetailsExit

End Function

Private Function GetLoanComponentData(ByVal xmlLoanComponentPaymentNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo GetLoanComponentDataErr

    Const strFunctionName As String = "GetLoanComponentData"
        
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
        
    Dim strTempResponse As String
    
    Dim objMortSubQuoteBO As Object
    
    'Create a temp request doc
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    'Create a temp response doc
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.async = False
       
    'Set up the TempRequest node
    Set xmlTempRequestNode = xmlTempRequestDoc.createElement("REQUEST")
    xmlTempRequestDoc.appendChild xmlTempRequestNode
    Set xmlNode = xmlTempRequestDoc.createElement("MORTGAGESUBQUOTE")
    xmlTempRequestDoc.selectSingleNode("REQUEST").appendChild xmlNode
        
    'Append required fields
    Set xmlNode = xmlTempRequestDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetAttributeText(xmlLoanComponentPaymentNode, "APPLICATIONNUMBER")
    xmlTempRequestDoc.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlNode
            
    Set xmlNode = xmlTempRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = xmlGetAttributeText(xmlLoanComponentPaymentNode, "APPLICATIONFACTFINDNUMBER")
    xmlTempRequestDoc.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlNode
          
    Set xmlNode = xmlTempRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlNode.Text = xmlGetAttributeText(xmlLoanComponentPaymentNode, "MORTGAGESUBQUOTENUMBER")
    xmlTempRequestDoc.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlNode
      
    Set xmlNode = xmlTempRequestDoc.createElement("LOANCOMPONENTSEQUENCENUMBER")
    xmlNode.Text = xmlGetAttributeText(xmlLoanComponentPaymentNode, "LOANCOMPONENTSEQUENCENUMBER")
    xmlTempRequestDoc.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlNode
          
    'Call Cost Modelling component
    Set objMortSubQuoteBO = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
    strTempResponse = objMortSubQuoteBO.GetLoanComponentData(xmlTempRequestNode.xml)
    xmlTempResponseDoc.loadXML strTempResponse
                
    errCheckXMLResponseNode xmlTempResponseDoc.selectSingleNode("RESPONSE"), , True
                
    Set GetLoanComponentData = xmlTempResponseDoc.selectSingleNode(".//LOANCOMPONENT")
        
    gobjContext.SetComplete
        
GetLoanComponentDataExit:

    If Not xmlTempRequestDoc Is Nothing Then Set xmlTempRequestDoc = Nothing
    If Not xmlTempResponseDoc Is Nothing Then Set xmlTempResponseDoc = Nothing
    If Not xmlTempRequestNode Is Nothing Then Set xmlTempRequestNode = Nothing
    If Not xmlTempResponseNode Is Nothing Then Set xmlTempResponseNode = Nothing
    If Not xmlNode Is Nothing Then Set xmlNode = Nothing
    If Not objMortSubQuoteBO Is Nothing Then Set objMortSubQuoteBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

    Exit Function
    
GetLoanComponentDataErr:
    gobjContext.SetAbort
    GoTo GetLoanComponentDataExit

End Function

Private Function GetMortgageProductData(ByVal xmlLoanComponentNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo GetMortgageProductDataErr

    Const strFunctionName As String = "GetMortgageProductData"
        
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
        
    Dim strTempResponse As String
    
    Dim objMortgageProductBO As Object
    
    'Create a temp request doc
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    'Create a temp response doc
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.async = False
       
    'Set up the TempRequest node
    Set xmlTempRequestNode = xmlTempRequestDoc.createElement("REQUEST")
    xmlTempRequestDoc.appendChild xmlTempRequestNode
    Set xmlNode = xmlTempRequestDoc.createElement("MORTGAGEPRODUCT")
    xmlTempRequestDoc.selectSingleNode("REQUEST").appendChild xmlNode
    
    'Append required fields
    Set xmlNode = xmlTempRequestDoc.createElement("MORTGAGEPRODUCTCODE")
    xmlNode.Text = xmlLoanComponentNode.selectSingleNode("//MORTGAGEPRODUCTCODE").Text
    xmlTempRequestDoc.selectSingleNode("//MORTGAGEPRODUCT").appendChild xmlNode
            
    Set xmlNode = xmlTempRequestDoc.createElement("STARTDATE")
    xmlNode.Text = xmlLoanComponentNode.selectSingleNode("//STARTDATE").Text
    xmlTempRequestDoc.selectSingleNode("//MORTGAGEPRODUCT").appendChild xmlNode
                    
    'Call Mortgage Product component
    Set objMortgageProductBO = gobjContext.CreateInstance("omMP.MortgageProductBO")
    strTempResponse = objMortgageProductBO.FindProductList(xmlTempRequestNode.xml)
    xmlTempResponseDoc.loadXML strTempResponse
    
    errCheckXMLResponseNode xmlTempResponseDoc.selectSingleNode("RESPONSE"), , True
        
    Set GetMortgageProductData = xmlTempResponseDoc.selectSingleNode(".//MORTGAGEPRODUCT")
        
    gobjContext.SetComplete
        
GetMortgageProductDataExit:

    If Not xmlTempRequestDoc Is Nothing Then Set xmlTempRequestDoc = Nothing
    If Not xmlTempResponseDoc Is Nothing Then Set xmlTempResponseDoc = Nothing
    If Not xmlTempRequestNode Is Nothing Then Set xmlTempRequestNode = Nothing
    If Not xmlTempResponseNode Is Nothing Then Set xmlTempResponseNode = Nothing
    If Not xmlNode Is Nothing Then Set xmlNode = Nothing
    If Not objMortgageProductBO Is Nothing Then Set objMortgageProductBO = Nothing

    errCheckError strFunctionName, TypeName(Me)

    Exit Function
    
GetMortgageProductDataErr:
    gobjContext.SetAbort
    GoTo GetMortgageProductDataExit

End Function

Private Function GetBACSData(xmlMortgageProductNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo GetBACSDataErr

    Const strFunctionName As String = "GetBACSData"
        
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
        
    Dim strTempResponse As String
    
    Dim objMortgageLenderBO As Object
    
    'Create a temp request doc
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    xmlTempRequestDoc.async = False
    
    'Create a temp response doc
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.async = False
       
    'Set up the TempRequest node
    Set xmlTempRequestNode = xmlTempRequestDoc.createElement("REQUEST")
    xmlTempRequestDoc.appendChild xmlTempRequestNode
    Set xmlNode = xmlTempRequestDoc.createElement("MORTGAGELENDER")
    xmlTempRequestDoc.selectSingleNode("REQUEST").appendChild xmlNode
    
    'Append required fields
    Set xmlNode = xmlTempRequestDoc.createElement("ORGANISATIONID")
    xmlNode.Text = xmlMortgageProductNode.selectSingleNode("//ORGANISATIONID").Text
    xmlTempRequestDoc.selectSingleNode("//MORTGAGELENDER").appendChild xmlNode
                                
    'Call Mortgage Product component
    Set objMortgageLenderBO = gobjContext.CreateInstance("omOrg.MortgageLenderBO")
    strTempResponse = objMortgageLenderBO.GetLenderDetails(xmlTempRequestNode.xml)
    xmlTempResponseDoc.loadXML strTempResponse
    
    errCheckXMLResponseNode xmlTempResponseDoc.selectSingleNode("RESPONSE"), , True
        
    Set GetBACSData = xmlTempResponseDoc.selectSingleNode(".//MORTGAGELENDER")
        
    gobjContext.SetComplete
        
GetBACSDataExit:

    If Not xmlTempRequestDoc Is Nothing Then Set xmlTempRequestDoc = Nothing
    If Not xmlTempResponseDoc Is Nothing Then Set xmlTempResponseDoc = Nothing
    If Not xmlTempRequestNode Is Nothing Then Set xmlTempRequestNode = Nothing
    If Not xmlTempResponseNode Is Nothing Then Set xmlTempResponseNode = Nothing
    If Not xmlNode Is Nothing Then Set xmlNode = Nothing
    If Not objMortgageLenderBO Is Nothing Then Set objMortgageLenderBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

    Exit Function

GetBACSDataErr:
    gobjContext.SetAbort
    GoTo GetBACSDataExit

End Function

Private Sub OutputToFile(ByVal strPaymentRecord As String)
    
    On Error GoTo OutputToFileErr

    Const strFunctionName As String = "OutputToFile"
        
    'Some objects to play with
    Dim strFilePathAndName As String
    'BMIDS646
'    Dim strPath As String
'    Dim fsoBACS As FileSystemObject
'    Dim tsBACS As TextStream
'    Dim intPos As Integer
    Dim objCRM As omCRMWrapper
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strResponse As String
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Get the file path and location
    strFilePathAndName = globalAssist.GetGlobalParamString("BACSPathandFileName")
    
    'Check it's a text file
    If Right(strFilePathAndName, 4) <> ".txt" Then
        strFilePathAndName = strFilePathAndName & ".txt"
    End If

    'BMIDS646 omCRM will create the folder/file if required
'    'Open or create the file
'    Set fsoBACS = New FileSystemObject
'    If fsoBACS.FileExists(strFilePathAndName) Then
'        'Open an existing file...
'        Set tsBACS = fsoBACS.OpenTextFile(strFilePathAndName, ForAppending)
'    Else
'        'Create the file...
'        'Does the path exist?
'        intPos = 5
'        Do While intPos < Len(strFilePathAndName)
'            If Left(Right(strFilePathAndName, intPos), 1) = "\" Then
'                'Can now capture the path
'                strPath = Left(strFilePathAndName, (Len(strFilePathAndName) - intPos))
'                'Break the loop
'                intPos = Len(strFilePathAndName)
'            End If
'            intPos = intPos + 1
'        Loop
'
'        If fsoBACS.FolderExists(strPath) = False Then
'            fsoBACS.CreateFolder strPath
'        End If
'
'        Set tsBACS = fsoBACS.CreateTextFile(strFilePathAndName, True)
'    End If
'
'    tsBACS.WriteLine strPaymentRecord
'
'    'Close
'    tsBACS.Close

    'BMIDS646
    nMaxRetries = GetMandatoryGlobalParamAmount("LockTextFileRetries")
    nSleepMilliSecs = GetMandatoryGlobalParamAmount("LockTextFileSleep")
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("TEXTOUT")
    xmlElement.setAttribute "PATHANDFILENAME", strFilePathAndName
    xmlElement.setAttribute "MAXRETRIES", nMaxRetries
    xmlElement.setAttribute "SLEEPMILLISECS", nSleepMilliSecs
    'WRITEMODE is either "Append" or OverWrite"
    xmlElement.setAttribute "WRITEMODE", "Append"
    xmlRequest.appendChild xmlElement
    
    'add carriage return/line feed
    strPaymentRecord = strPaymentRecord & Chr$(13) & Chr$(10)
    
    Set objCRM = gobjContext.CreateInstance(gstrCRM_COMPONENT & ".omCRMWrapper")
    strResponse = objCRM.AddText(xmlRequest.xml, strPaymentRecord)
                                                
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True


    gobjContext.SetComplete

OutputToFileExit:

    'BMIDS646
    Set objCRM = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlRequest = Nothing
'    If Not fsoBACS Is Nothing Then Set fsoBACS = Nothing
'    If Not tsBACS Is Nothing Then Set tsBACS = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
OutputToFileErr:
    gobjContext.SetAbort
    GoTo OutputToFileExit

End Sub

Private Function ConstructPaymentRecord() As String

    On Error GoTo ConstructPaymentRecordExit

    Const strFunctionName As String = "ConstructPaymentRecord"
           
    Dim lngDecimalPos As Long
           
    'Sort Codes - remove "-"
    mstrOriginatorSortCode = Replace(mstrOriginatorSortCode, "-", "")
    mstrDestSortCode = Replace(mstrDestSortCode, "-", "")
    
    'Account numbers - pad left with 0
    If Len(mstrDestAccountNo) < 8 Then
         mstrDestAccountNo = Replace((Space(8 - Len(mstrDestAccountNo)) & mstrDestAccountNo), " ", "0")
    End If
    
    If Len(mstrOriginatorAccNo) < 8 Then
        mstrOriginatorAccNo = Replace((Space(8 - Len(mstrOriginatorAccNo)) & mstrOriginatorAccNo), " ", "0")
    End If
        
    'Amount - Make sure 2dp
    lngDecimalPos = InStr(1, mstrAmount, ".", vbTextCompare)
    If lngDecimalPos > 0 Then
        'Work out where the decimal is.
        If Len(mstrAmount) - lngDecimalPos = 2 Then
            '2 decimal places
        ElseIf Len(mstrAmount) - lngDecimalPos = 1 Then
            '1 decimal place
            mstrAmount = mstrAmount & "0"
        Else
            'More than 2 decimal places
            mstrAmount = Round(Val(mstrAmount), 2)
        End If
    Else
        'No decimals
        mstrAmount = mstrAmount & ".00"
    End If
        
    'Amount - pad left with 0
    If Len(mstrAmount) < 11 Then
        mstrAmount = Replace((Space(11 - Len(mstrAmount)) & mstrAmount), " ", "0")
    End If
    
    'Strings - Pad right with spaces
    
    'SG 15/05/02 MSMS0082 added elseif
    If Len(mstrOriginatorName) < 18 Then
        mstrOriginatorName = mstrOriginatorName & Space(18 - Len(mstrOriginatorName))
    ElseIf Len(mstrOriginatorName) > 18 Then
        mstrOriginatorName = Left(mstrOriginatorName, 18)
    End If
    
    If Len(mstrOriginatorRef) < 18 Then
        mstrOriginatorRef = mstrOriginatorRef & Space(18 - Len(mstrOriginatorRef))
    End If
    
    'SG 15/05/02 MSMS0082 added elseif
    If Len(mstrDestAccName) < 18 Then
        mstrDestAccName = mstrDestAccName & Space(18 - Len(mstrDestAccName))
    ElseIf Len(mstrDestAccName) > 18 Then
        mstrDestAccName = Left(mstrDestAccName, 18)
    End If

    'Output the final string
    ConstructPaymentRecord = mstrDestSortCode & _
                             mstrDestAccountNo & _
                             mstrAccountType & _
                             mstrTransactionCode & _
                             mstrOriginatorSortCode & _
                             mstrOriginatorAccNo & _
                             mstrStandingOrderRef & _
                             mstrAmount & _
                             mstrOriginatorName & _
                             mstrOriginatorRef & _
                             mstrDestAccName
                             
ConstructPaymentRecordExit:

    errCheckError strFunctionName, TypeName(Me)

End Function



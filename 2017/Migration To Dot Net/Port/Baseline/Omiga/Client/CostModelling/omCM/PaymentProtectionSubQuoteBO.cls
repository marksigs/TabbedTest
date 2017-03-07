VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProtectionSubQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProtectionSubQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Payment Protection Sub Quote Business Object which
'               'supports transactions' only. Any transactioned requests will
'               be delegated to PaymentProtectionSubQuoteTxBO

'Dependencies:  PaymentProtectionSubQuoteTxBO, PaymentProtectionSubQuoteDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    03/11/99    Created
'SR     30/11/99    Added method ValidateSubQuote
'MCS    05/01/2000  interface and modified ValidateSubQuote
'PSC    31/01/2000  Amend implementation of ValidateSubQuote
'JLD    29/03/00    Added string interface for ValidateSubQuote (used in gui CM075)
'APS    30/05/2000  SYS0647 - Recommended Cover Amount should have no decimal places
'MS     08/07/00    SYS0824 - UK residency NOT a criterion for PP Eligibility
'MS     12/06/00    SYS0882 - Be sure we have a second applicant other wise we return
'                             blnApp2Eligible = True even if we havn't got one.
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
Implements IPaymentProtectionSubQuoteBO
Private Const cstrSUBQUOTE_NAME = "PAYMENTPROTECTIONSUBQUOTE"
Public Function ValidateSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Call validate function
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateSubQuote"

    
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
    Set xmlTempResponseNode = IPaymentProtectionSubQuoteBO_ValidateSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ValidateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function CreateFirstSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  creates the first payment protection sub quote for an application
'               fact find
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr
    
    Const strFunctionName As String = "CreateFirstSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objPPSubQuoteTxBO As PaymentProtectionSubQuoteTxBO
    Dim xmlDoc As FreeThreadedDOMDocument40

    If objContext Is Nothing Then
        Set objPPSubQuoteTxBO = New PaymentProtectionSubQuoteTxBO
    Else
        Set objPPSubQuoteTxBO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strResponse As String
    
    strResponse = objPPSubQuoteTxBO.CreateFirstSubQuote(vstrXMLRequest)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    xmlResponseElem.appendChild xmlDoc.documentElement
        
    CreateFirstSubQuote = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objPPSubQuoteTxBO = Nothing
    Set xmlDoc = Nothing

    Exit Function

CreateFirstSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateFirstSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstSubQuoteExit

End Function

Public Function CreateNewSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  creates a new payment protection sub quote for an application
'               fact find
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr
    
    Const strFunctionName As String = "CreateNewSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objPPSubQuoteTxBO As PaymentProtectionSubQuoteTxBO
    Dim xmlDoc As FreeThreadedDOMDocument40

    If objContext Is Nothing Then
        Set objPPSubQuoteTxBO = New PaymentProtectionSubQuoteTxBO
    Else
        Set objPPSubQuoteTxBO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strResponse As String
    
    strResponse = objPPSubQuoteTxBO.CreateNewSubQuote(vstrXMLRequest)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    xmlResponseElem.appendChild xmlDoc.documentElement
        
    CreateNewSubQuote = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objPPSubQuoteTxBO = Nothing
    Set xmlDoc = Nothing

    Exit Function

CreateNewSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewSubQuoteExit

End Function


Public Function SaveDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the payment protection sub quote and details
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveDetailsVbErr
    
    Const strFunctionName As String = "SaveDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objPPSubQuoteTxBO As PaymentProtectionSubQuoteTxBO
    
    If objContext Is Nothing Then
        Set objPPSubQuoteTxBO = New PaymentProtectionSubQuoteTxBO
    Else
        Set objPPSubQuoteTxBO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    objPPSubQuoteTxBO.SaveDetails vstrXMLRequest
        
    SaveDetails = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveDetailsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objPPSubQuoteTxBO = Nothing
    
    Exit Function

SaveDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveDetails = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume SaveDetailsExit
    
End Function


Public Function GetSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the payment protection sub quote and details
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetSubQuoteDetailsVbErr
    
    Const strFunctionName As String = "GetSubQuoteDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
    
    '   call Data Object GetData function
    Dim strXMLData As String
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrSUBQUOTE_NAME & " tag not found"
    End If
        
    strXMLData = objPPSubQuoteDO.GetSubQuoteDetails(xmlRequestNode.xml)
    Set xmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
    
    GetSubQuoteDetails = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetSubQuoteDetailsExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objPPSubQuoteDO = Nothing

    Exit Function

GetSubQuoteDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetSubQuoteDetails = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetSubQuoteDetailsExit

End Function
    
Public Function FindList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "FindList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
        
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
    
    Dim strXMLData As String
    
    '   call Data Object FindList function
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    strXMLData = objPPSubQuoteDO.FindList(xmlRequestNode.xml)
    Set xmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
        
    FindList = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objPPSubQuoteDO = Nothing

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

    FindList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindListExit

End Function

Public Function CheckApplicantEligibility(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks whether applicants are eligible for payment protection
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CheckApplicantEligibilityVbErr
    
    Const strFunctionName As String = "CheckApplicantEligibility"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrSUBQUOTE_NAME & " tag not found"
    End If
    
'    Dim blnApp1UKResident As Boolean
'    Dim blnApp2UKResident As Boolean
    Dim blnApp1EmpProblem As Boolean
    Dim blnApp2EmpProblem As Boolean
    Dim strApp1EmpStatus As String
    Dim strApp2EmpStatus As String
    Dim strTagName As String
    Dim strTagValue As String
        
'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
    
'    strTagName = "APPLICANT1UK"
'    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
'    blnApp1UKResident = objConvAssist.CSafeBool(strTagValue)
        
'    strTagName = "APPLICANT2UK"
'    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
'    blnApp2UKResident = objConvAssist.CSafeBool(strTagValue)
       
    strTagName = "APPLICANT1EMPLOYMENTSTATUS"
    strApp1EmpStatus = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
       
    strTagName = "APPLICANT2EMPLOYMENTSTATUS"
    strApp2EmpStatus = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
        
    strTagName = "APPLICANT1EMPPROBLEM"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    blnApp1EmpProblem = objConvAssist.CSafeBool(strTagValue)
        
    strTagName = "APPLICANT2EMPPROBLEM"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    blnApp2EmpProblem = objConvAssist.CSafeBool(strTagValue)
       
    Dim blnApp1Eligible As Boolean
    Dim blnApp2Eligible As Boolean
       
    blnApp1Eligible = False
    blnApp2Eligible = False
       
    ' Check applicant 1 for eligibility
    ' Applicant must reside in UK, be employed and not have an employment problem
    ' APS 08/03/00 - Change combo validation match from 'U' to 'N'
    
    ' MS SYS0824 UK residency NOT a criterion for PP Eligibility
    'If blnApp1UKResident = True And
    If StrComp(strApp1EmpStatus, "N", vbTextCompare) <> 0 And _
       blnApp1EmpProblem = False Then
        blnApp1Eligible = True
    End If
    
    ' Check applicant 2 for eligibility
    ' Applicant must reside in UK, be employed and not have an employment problem
    ' APS 08/03/00 - Change combo validation match from 'U' to 'N'
    
    ' MS SYS0824 UK residency NOT a criterion for PP Eligibility
    
    'If blnApp2UKResident = True And
    
    'MS SYS0882 Be sure we have a second applicant other wise we return
    'blnApp2Eligible = True even if we havn't got one.
    
    If Len(strApp2EmpStatus) > 0 Then
        If StrComp(strApp2EmpStatus, "N", vbTextCompare) <> 0 And _
           blnApp2EmpProblem = False Then
            blnApp2Eligible = True
        End If
    End If
    
    Set xmlElement = xmlOut.createElement("APPLICANT1ELIGIBLE")
    xmlElement.Text = Abs(CInt(blnApp1Eligible))
    xmlResponseElem.appendChild xmlElement
       
    Set xmlElement = xmlOut.createElement("APPLICANT2ELIGIBLE")
    xmlElement.Text = Abs(CInt(blnApp2Eligible))
    xmlResponseElem.appendChild xmlElement
      
    CheckApplicantEligibility = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CheckApplicantEligibilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    
    Exit Function

CheckApplicantEligibilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CheckApplicantEligibility = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CheckApplicantEligibilityExit

End Function

Public Function CalculateRecommendedCover(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the recommended payment protection cover
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateRecommendedCoverVbErr
    
    Const strFunctionName As String = "CalculateRecommendedCover"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    Dim objGlobalParamDO As GlobalParameterDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim dblNewRepayments As Double
    Dim dblExistingRepayments As Double
    Dim strTagName As String
    Dim strTagValue As String
    
    strTagName = "NEWREPAYMENTS"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblNewRepayments = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "EXISTINGREPAYMENTS"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblExistingRepayments = objConvAssist.CSafeDbl(strTagValue)
   
    Dim dblPPRoundingFactor As Long
    Dim dblPPMaxCover As Long
    Dim dblRecommendedCover As Double
    Dim lngRemainder As Long
    
    dblPPRoundingFactor = objGlobalParamDO.GetCurrentParameterByType("PPRoundingFactor", _
                                                                     "AMOUNT")
      
    dblPPMaxCover = objGlobalParamDO.GetCurrentParameterByType("PPMaximumCover", "AMOUNT")
          
    dblRecommendedCover = dblNewRepayments + dblExistingRepayments
    
    'APS 30/05/2000 SYS0647 - Recommended Cover Amount should have no decimal places,
    ' rounded up the nearest pound
    dblRecommendedCover = Round(dblRecommendedCover + 0.5, 0)
    lngRemainder = dblRecommendedCover Mod dblPPRoundingFactor
    
    If lngRemainder <> 0 Then
        dblRecommendedCover = dblRecommendedCover + dblPPRoundingFactor - lngRemainder
    End If
    
    If dblRecommendedCover > dblPPMaxCover Then
        dblRecommendedCover = dblPPMaxCover
    End If
    
    strTagName = "RECOMMENDEDCOVER"
    Set xmlElement = xmlOut.createElement(strTagName)

    xmlElement.Text = CStr(dblRecommendedCover)
    xmlResponseElem.appendChild xmlElement
    
    CalculateRecommendedCover = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalculateRecommendedCoverExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestElem = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    
    Exit Function

CalculateRecommendedCoverVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateRecommendedCover = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateRecommendedCoverExit

End Function

'Public Function ValidateSubQuote(ByVal vstrXMLRequest As String) As String
'' header ----------------------------------------------------------------------------------
'' description:
'' pass:         vstrXMLRequest  xml Request data stream of the format
''
'' return:      xml Response data stream containing results of operation
''              either: TYPE="SUCCESS" and 1 (True) if Quote is Valid Else 0 (False)
''              or: TYPE="SYSERR" and <ERROR> element
''NOTE ** CURRENTLY THIS METHOD JUST RETURNS TRUE FOR ALL QUOTES (NOT FULLY IMPLEMENTED)
''------------------------------------------------------------------------------------------
'On Error GoTo ValidateSubQuoteVbErr
'
'    Dim strFunctionName As String
'    strFunctionName = "ValidateSubQuote"

'    #If USING_VSA Then
'        Dim VSA As vsa_shared
'        Set VSA = New vsa_shared
'        VSA.Initialize (strFunctionName)
'    #End If
'
'    Dim xmlOut As New DOMDocument
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlResponseNode As IXMLDOMNode
'    Dim xmlElement As IXMLDOMElement
'
'    Dim objxmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
'
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    Set xmlElement = xmlOut.createElement("SUBQUOTEVALIDITY")
'    xmlElement.Text = 1
'    xmlResponseElem.appendChild xmlElement
'
'    ValidateSubQuote = xmlOut.xml
'
'ValidateSubQuoteExit:
'
'    '   clean-up section, release object references
'    Set xmlOut = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlResponseNode = Nothing
'    Set xmlElement = Nothing
'
'    Set objxmlAssist = Nothing
'    Set objErrAssist = Nothing
'
'    Exit Function
'
'ValidateSubQuoteVbErr:
'
'    If objErrAssist.IsWarning = True Then
'        objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'
'    objErrAssist.AddToErrSource strFunctionName
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    ValidateSubQuote = objErrAssist.CreateErrorResponse
'
'    ' go to clean-up section
'    Resume ValidateSubQuoteExit
'
'End Function
'
Private Function IPaymentProtectionSubQuoteBO_ValidateSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:      xml Response data stream containing results of operation
'              either: TYPE="SUCCESS" and 1 (True) if Quote is Valid Else 0 (False)
'              or: TYPE="SYSERR" and <ERROR> element
'NOTE ** CURRENTLY THIS METHOD JUST RETURNS TRUE FOR ALL QUOTES (NOT FULLY IMPLEMENTED)
'------------------------------------------------------------------------------------------
On Error GoTo ValidateSubQuoteVbErr
    
    Const strFunctionName As String = "IPaymentProtectionSubQuoteBO_ValidateSubQuote"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRootElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objErrAssist As New ErrAssist
    Dim strTagName As String
    
    strTagName = "PAYMENTPROTECTIONSUBQUOTE"
    
    If vxmlTableElement.nodeName = strTagName Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = vxmlTableElement.getElementsByTagName("PAYMENTPROTECTIONSUBQUOTE").Item(0)
    End If
        
    Dim blnValidProduct As Boolean
    blnValidProduct = True
    
    Set xmlRootElem = xmlOut.createElement("PAYMENTPROTECTIONSUBQUOTE")
    xmlResponseElem.appendChild xmlRootElem
    
    Set xmlElement = xmlOut.createElement("VALIDPPSUBQUOTE")
    xmlRootElem.appendChild xmlElement
    xmlElement.Text = Abs(blnValidProduct)
          
    Set IPaymentProtectionSubQuoteBO_ValidateSubQuote = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ValidateSubQuoteExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set objErrAssist = Nothing

    Exit Function

ValidateSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    IPaymentProtectionSubQuoteBO_ValidateSubQuote = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ValidateSubQuoteExit

End Function


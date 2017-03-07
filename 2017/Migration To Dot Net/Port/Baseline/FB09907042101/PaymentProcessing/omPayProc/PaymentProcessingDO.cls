VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProcessingDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProcessingDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   PaymentProcessing Data Object.

'Dependencies:
'-------------------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'SR     01/08/01    SYS2552 - Modified methods FindLoanComponentPaymentList, FindSanctioningList
'SR     06/09/01    SYS2412 - New method 'DeletePayeeHistory'
'DRC    15/10/01    SYS2797 - Modified PaymentMethod handling in FindSanctioningList
'MDC    05/12/2001  SYS2997 - Generic SQL change
'SR     12/12/2001  SYS3391 - new method DeleteFeePayments
'MDC    20/12/2001  SYS2997 - Further generic SQL change
'------------------------------------------------------------------------------------
'   BMIDS
'
'Prog   Date        AQR         Description
'AW     23/08/2002  BM029       Added CreateChequeAudit
'PSC    18/09/2002  SYS5181     Applied this core change to BMids version
'PSC    18/09/2002  SYS4863     Applied this core change to BMids version
'PSC    20/09/2002  BMIDS00480  Amend FindPaymentsForBatch to return PaymentType and
'                               add FindInterfacedInitialAdvances
'PSC    08/10/2002  BMIDS00595  Change FindPaymentsForBatch to pick up incentive releases
'PSC    13/11/2002  BMIDS00915  Change FindPaymentsForBatch to only select disbursements that
'                               have not been returned
'PSC    15/11/2002  BMIDS00968  Change FindPaymentsForBatch to get all value ids for payment types
'KRW    27/01/2004  BMIDS679    Change to way  results are displayed,Now Ordered by Application Number and then Payment Type.
'MC     12/06/2004  BMIDS763    GetProductSwitchFee() AND GetTTFeeAmount() added
'MC     23/06/2004  BMIDS763 executeGetRecordSetEx() ADDED TO handle nextrecordset() functionality from
'                   stored Procedure
'SR     06/08/2004  BMIDS813    Modified methods GetProductSwitchFee, GetTTFeeAmount
'KRW    04/11/2004  BMIDS665    Changed search filter to not include payments with cancellationdates
'------------------------------------------------------------------------------------
' MARS
' HM    23/09/2005  WP13 MAR49  ADDED ValidateCompletionDate, ValidateSolicitor, SetAdvanceDate,  SetPaymentMethod, CreateDelayedCompletionTask
' JD    04/11/2005  MAR414      changed SetAdvanceDate to use global TTPaymentDays
' JD    08/11/2005  MAR383      Changed logic in SetPaymentMethod to match spec.
' PSC   15/11/2005  MAR140      Remove GetPostCompletionCaseTasksList
' HMA   30/11/2005  MAR736      Changed ResetPostCompletionTasks
' JD    08/12/2005  MAR803      UPDATEPROPERTY attribute for ResetPostCompletionTasks is TaskStatus not TaskStatusId
' JD    14/12/2005  MAR845      Pass requestnode not response node to IsDateWorkingDay in SetAdvanceDate and SetPaymentMethod. IsDateWorkingDay uses REQUEST attributes in call to omORG
' JD    09/02/2006  MAR1220     use globals in SetPaymentMethod
' JD    10/02/2006  MAR1245     Check advance date <= today in FindPaymentsForBatch
'JD     15/02/2006  MAR1273     SetPaymentMethod adding too many business days
' DRC   02/03/2006  MAR1334     Added FindPaymentsToCancel
' JD    15/03/2006  MAR1364     SetAdvanceDate - take off business days from completion date.
'GHun   19/04/2006  MAR1632     Changed IsDateWorkingDay, SetPaymentMethod and ValidateCompletionDate
'GHun   20/04/2005  MAR1547     Changed ResetPostCompletionTasks to set the time on the TaskDueDateAndTime
'************************************************************************************
Option Explicit

Implements IPAymentProcessingDO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub IPAymentProcessingDO_FindPaymentsToCancel(ByVal vxmlRequestNode As MSXML2.IXMLDOMNode, ByVal vxmlResponseNode As MSXML2.IXMLDOMNode)
On Error GoTo FindPaymentsToCancelExit
'
' Gets all outstanding disbursement payment records for a single application
'
    Const cstrFunctionName As String = "IPAymentProcessingDO_FindPaymentsToCancel"
    
    Dim objParam As ADODB.Parameter
    Dim objCommand As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim nodeHeader As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlDisbursementElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strSQL As String
    Dim strValue As String
    Dim strPaymentStatus As String
    Dim strPayMethod As String
    Dim strPayTypeBCP As String
    Dim blnSelected As Boolean
    Dim strResponse As String

    Set objCommand = New ADODB.Command
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    

    strSQL = "select pr.applicationnumber, pr.paymentsequencenumber, dp.paymenttype "
    strSQL = strSQL & "From paymentrecord pr, disbursementpayment dp "
    strSQL = strSQL & "Where pr.applicationnumber = dp.applicationnumber "
    strSQL = strSQL & "and pr.paymentsequencenumber = dp.paymentsequencenumber"
    
     
    Set nodeHeader = xmlGetMandatoryNode(vxmlRequestNode, "HEADER")
    
    'APPLICATIONNUMBER
    strValue = xmlGetAttributeText(nodeHeader, "APPLICATIONNUMBER")
    If Len(Trim$(strValue)) > 0 Then

        #If GENERIC_SQL Then
            strSQL = strSQL & " and pr.applicationnumber = '" & strValue & "'"
        #Else
            Set objParam = New ADODB.Parameter
            strSQL = strSQL & " and pr.applicationnumber = ?"
            objParam.Type = adBSTR
            objParam.Direction = adParamInput
            objParam.Value = strValue
            objCommand.Parameters.Append objParam
            Set objParam = Nothing
        #End If

    End If
   
    
    'PAYMENTTYPE - All Outstanding
            'Pay Sanctioned Payment
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "S")   'Sanctioned
        strPayTypeBCP = GetValueIdForValidationType("PaymentType", "NCB")   'Balance Cancellation Payment
        strSQL = strSQL & " and (( dp.paymentstatus = " & strPaymentStatus
        
        strSQL = strSQL & " and dp.paymenttype != " & strPayTypeBCP & ")"
           'Interface Paid Payment
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "P")   'Paid
        strSQL = strSQL & " or dp.paymentstatus = " & strPaymentStatus
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "IF")   'Failed
        strSQL = strSQL & " or dp.paymentstatus = " & strPaymentStatus
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "U")   'Failed
        strSQL = strSQL & " or dp.paymentstatus = " & strPaymentStatus & ")"
        
        'PSC 18/09/2002 SYS5181 - End
        
        ' PSC 13/11/2002 BMIDS00915
        strSQL = strSQL & " AND NOT EXISTS (SELECT 'X' FROM PAYMENTRECORD " & _
                                           "WHERE APPLICATIONNUMBER = DP.APPLICATIONNUMBER " & _
                                           "AND ASSOCPAYSEQNUMBER = DP.PAYMENTSEQUENCENUMBER)"

    
    'Execute the Command SQL
    objCommand.CommandType = adCmdText
    objCommand.CommandText = strSQL
    Debug.Print strSQL
    Set objRs = executeGetRecordSet(objCommand)

    'Convert Recordset to XML Response
    If Not objRs Is Nothing Then
        If Not (objRs.BOF Or objRs.EOF) Then
            objRs.MoveFirst
            Do Until objRs.EOF
                Set xmlElement = xmlDoc.createElement("PAYMENTRECORD")
                xmlElement.setAttribute "APPLICATIONNUMBER", objRs.Fields("APPLICATIONNUMBER")
                xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", objRs.Fields("PAYMENTSEQUENCENUMBER")
                
                xmlElement.setAttribute "PAYMENTTYPE", objRs.Fields("PAYMENTTYPE")
                Set xmlDisbursementElement = xmlDoc.createElement("DISBURSEMENTPAYMENT")
                xmlDisbursementElement.setAttribute "APPLICATIONNUMBER", objRs.Fields("APPLICATIONNUMBER")
                xmlDisbursementElement.setAttribute "PAYMENTSEQUENCENUMBER", objRs.Fields("PAYMENTSEQUENCENUMBER")
                xmlElement.appendChild xmlDisbursementElement
                vxmlResponseNode.appendChild xmlElement
                objRs.MoveNext
            Loop
        Else
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
    Else
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindPaymentsToCancelExit:

    Set objParam = Nothing
    Set objCommand = Nothing
    Set objRs = Nothing
    Set nodeHeader = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    
    
    errCheckError cstrFunctionName, TypeName(Me)


    

End Sub



Private Sub IPAymentProcessingDO_ValidateCompletionDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Const cstrFunctionName = "IPAymentProcessingDO_ValidateCompletionDate"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlChild As IXMLDOMNode
    
    'request params
    Dim strAppNumber As String
    Dim strAppFactFindNumber As String
    Dim strAppCompletionDate As String
    'global params
    Dim objGlobalParamDO As GlobalParameterDO
    Dim lbPPValidateCompletionDate As Boolean
    Dim intPlusDays As Integer
    'comp
    Dim objAppBO As ApplicationBO
    Dim objOrgBO As OrganisationBO
    Dim strAppOfferExpireDate As String
    Dim dtExpireDatePlusGlobalValue As Date
        
    Dim strResponse As String
    
On Error GoTo ValidateCompletionDateExit

    'Make a call to find whether the validation completion date flag is switched on
    'read global param "number of days after expire date"
    
    Set objGlobalParamDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    lbPPValidateCompletionDate = objGlobalParamDO.GetCurrentParameterByType("PPValidateCompletionDate", "BOOLEAN")
    intPlusDays = objGlobalParamDO.GetCurrentParameterByType("PPNoOfDaysAftOfferExpiryDate", "AMOUNT")
    
    If Not lbPPValidateCompletionDate Then
        'validation is switched off so return True to allow processing to continue
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    Else
        'validation is switched on so check the date
        'get params from request
        Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/VALIDATECOMPLETIONDATE")
        strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        strAppFactFindNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
        strAppCompletionDate = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")

        'form a new request
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
       
        ' check whether the completion date in the past
        If CDate(strAppCompletionDate) < Date Then
            errThrowError cstrFunctionName, 8518    '"Completion date is in the past"
        End If
        
        'check whether the date is a working day
        If Not IsDateWorkingDay(strAppCompletionDate, vxmlRequestNode) Then
            errThrowError cstrFunctionName, 8519    '"Completion date falls on non-business day"
        End If
    End If
    
    'get the offer expire date
    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")

    Set xmlRequest = xmlDoc.createElement("APPLICATIONOFFER")
    xmlSetAttributeValue xmlRequest, "OPERATION", "ApplicationOfferData"

    Set xmlNode = xmlDoc.createElement("APPLICATION")
    xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNumber
    xmlSetAttributeValue xmlNode, "APPLICATIONFACTFINDNUMBER", strAppFactFindNumber
    Set xmlChild = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChild.Text = strAppNumber
    xmlNode.appendChild xmlChild.cloneNode(True)
    Set xmlChild = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChild.Text = strAppFactFindNumber
    xmlNode.appendChild xmlChild.cloneNode(True)
    Set xmlChild = xmlDoc.createElement("APPLICATIONOFFERDATA")
    xmlNode.appendChild xmlChild.cloneNode(True)
    xmlRequest.appendChild xmlNode.cloneNode(True)

    'call GetApplicationOfferData
    strResponse = objAppBO.GetApplicationOfferData(xmlRequest.xml)

    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True  'MAR1632 GHun

    strAppOfferExpireDate = xmlGetNodeText(xmlDoc.documentElement, "RESPONSE/APPLICATIONOFFER/OFFEREXPIRYDATE")

    If IsDate(strAppOfferExpireDate) Then
        If CDate(strAppCompletionDate) > CDate(strAppOfferExpireDate) Then
            dtExpireDatePlusGlobalValue = DateAdd("d", intPlusDays, CDate(strAppOfferExpireDate))
            If CDate(strAppCompletionDate) > dtExpireDatePlusGlobalValue Then
                errThrowError cstrFunctionName, 8520    ' "Completion date is after the offer has expired. The offer should be reissured"
            Else
                errThrowError cstrFunctionName, 8521    ' "Completion date is after the offer has expired. The offer should be reissured and the completion date checked"
            End If
        End If
    Else
        errThrowError cstrFunctionName, 8525            '"Offer expiry date is invalid"
    End If
    
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"

    gobjContext.SetComplete
    
ValidateCompletionDateExit:

    Set objGlobalParamDO = Nothing
    Set objAppBO = Nothing
    Set objOrgBO = Nothing

    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlNode = Nothing
    Set xmlChild = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode

End Sub

Private Sub IPAymentProcessingDO_ValidateSolicitor(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Const cstrFunctionName = "IPAymentProcessingDO_ValidateSolicitor"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlChild As IXMLDOMNode
    Dim xmlTPRequest As FreeThreadedDOMDocument40
    Dim xmlTPElement As IXMLDOMElement
    Dim xmlTPNode As IXMLDOMElement

    Dim xmlGuidNode As IXMLDOMNode

    'request params
    Dim strAppNumber As String
    Dim strAppFactFindNumber As String
    'comp
    Dim objAppTPBO As ApplicationThirdPartyBO
    Dim objThirdPartyBO As ThirdPartyBO
    Dim strDirGUID As String

    Dim strResponse As String
    Dim lngErrNo As Long
    Dim strErrSource As String
    Dim strErrDescr As String

On Error GoTo ValidateSolicitorExit

    'get params from request
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/SOLICITOR")
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFactFindNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")

    'form a new request
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    Set xmlRequest = xmlDoc.createElement("REQUEST")
    
    Set xmlNode = xmlDoc.createElement("APPLICATIONLEGALREP")
    Set xmlChild = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChild.Text = strAppNumber
    xmlNode.appendChild xmlChild.cloneNode(True)
    Set xmlChild = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChild.Text = strAppFactFindNumber
    xmlNode.appendChild xmlChild.cloneNode(True)
    xmlRequest.appendChild xmlNode.cloneNode(True)

    'call GetLergalRep
    On Error Resume Next
    Set objAppTPBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
    strResponse = objAppTPBO.GetLegalRep(xmlRequest.xml)
    'errCheckXMLResponse strResponse, True
    
    lngErrNo = Err.Number
    strErrSource = Err.Source
    strErrDescr = Err.Description

    On Error GoTo ValidateSolicitorExit

    If Err.Number <> 0 Then
        errThrowError strErrSource, lngErrNo, strErrDescr
    End If

    xmlDoc.loadXML strResponse

    strDirGUID = xmlGetNodeText(xmlDoc.documentElement, "/RESPONSE/APPLICATIONLEGALREP/DIRECTORYGUID")

    If Len(strDirGUID) = 0 Then
        'error, not exist
        errThrowError cstrFunctionName, 8522 '"Solicitor is not on the panel"
    End If

    'get name and address directory info for this solicitor
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Create a temporary Request for the Third Party method call
    Set xmlTPRequest = New FreeThreadedDOMDocument40
    xmlTPRequest.validateOnParse = False
    xmlTPRequest.setProperty "NewParser", True
    Set xmlTPElement = xmlTPRequest.createElement("REQUEST")
    xmlTPRequest.appendChild xmlTPElement
    Set xmlTPNode = xmlTPRequest.createElement("NAMEANDADDRESSDIRECTORY")
    xmlTPElement.appendChild xmlTPNode

    Set xmlGuidNode = xmlTPRequest.selectSingleNode(".//THIRDPARTYGUID")
    If xmlGuidNode Is Nothing Then
        Set xmlGuidNode = xmlTPRequest.createElement("DIRECTORYGUID")
    End If
    xmlGuidNode.Text = strDirGUID
    xmlTPNode.appendChild xmlGuidNode

    'Find directory details
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    On Error Resume Next

    strResponse = objThirdPartyBO.GetDirectoryDetails(xmlTPRequest.xml)
        
    lngErrNo = Err.Number
    strErrSource = Err.Source
    strErrDescr = Err.Description

    On Error GoTo ValidateSolicitorExit

    If Err.Number <> 0 Then
        errThrowError strErrSource, lngErrNo, strErrDescr
    End If

    If xmlDoc.loadXML(strResponse) Then

        Dim strDateFrom As String
        Dim strDateTo As String
        strDateFrom = xmlGetNodeText(xmlDoc.documentElement, ".//NAMEANDADDRESSACTIVEFROM")
        strDateTo = xmlGetNodeText(xmlDoc.documentElement, ".//NAMEANDADDRESSACTIVETO")

        If IsDate(strDateFrom) Then
            If Date < CDate(strDateFrom) Then
                errThrowError cstrFunctionName, 8523 '"Solicitor is inactive"
            End If
        End If
        
        If IsDate(strDateTo) Then
            If Date > CDate(strDateTo) Then
                errThrowError cstrFunctionName, 8523 ' "Solicitor is inactive"
            End If
        End If
    Else
        errThrowError cstrFunctionName, oeXMLParserError, "Invalid Response from ThirdPartyBO.GetThirdParty"
    End If

    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"

    gobjContext.SetComplete

ValidateSolicitorExit:
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlNode = Nothing
    
    Set xmlTPRequest = Nothing
    Set xmlTPElement = Nothing
    Set xmlTPNode = Nothing
    Set xmlGuidNode = Nothing
    Set objAppTPBO = Nothing
    Set objThirdPartyBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_SetAdvanceDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName = "IPAymentProcessingDO_SetAdvanceDate"
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeResp As IXMLDOMNode

    'params
    Dim strAppCompletionDate As String
    Dim strAppPaymentMethod As String
    Dim strRtnAdvanceDate As String
    Dim strPMethodParamName As String
    'global params
    Dim objGlobalParamDO As GlobalParameterDO
    Dim lbPPDeriveAdvanceDate As Boolean
    Dim lngPaymentDays As Long
    
    Dim lbIsDateWorkingDay As Boolean
    
On Error GoTo SetAdvanceDateExit
    
    Set objGlobalParamDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    lbPPDeriveAdvanceDate = objGlobalParamDO.GetCurrentParameterByType("PPDeriveAdvanceDate", "BOOLEAN")
    
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/SETADVANCEDATE")
    strAppPaymentMethod = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTMETHOD")
    strAppCompletionDate = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")

    If Not lbPPDeriveAdvanceDate Then
        'the advance date will not be set by the system
        strRtnAdvanceDate = strAppCompletionDate
    Else
        Select Case strAppPaymentMethod
            
        Case "B", "YC", "CH"            '"BACS", "CHAPS/TT", "Cheque"
            Select Case strAppPaymentMethod
            Case "B"
                strPMethodParamName = "BACSPaymentDays"
            Case "YC"
                strPMethodParamName = "TTPaymentDays" 'JD MAR414 changed from CHAPSPaymentDays
            Case "CH"
                strPMethodParamName = "CHQPaymentDays"
            End Select
            
            lngPaymentDays = objGlobalParamDO.GetCurrentParameterByType(strPMethodParamName, "AMOUNT")
            'MAR1364 check they are business days
            'strRtnAdvanceDate = Format$(DateAdd("d", -lngPaymentDays, CDate(strAppCompletionDate)), "dd/mm/yyyy")
            strRtnAdvanceDate = strAppCompletionDate
            Dim i As Integer
            For i = 1 To lngPaymentDays
                lbIsDateWorkingDay = False
                Do While Not lbIsDateWorkingDay
                    strRtnAdvanceDate = Format$(DateAdd("d", -1, CDate(strRtnAdvanceDate)), "dd/mm/yyyy")
                    lbIsDateWorkingDay = IsDateWorkingDay(strRtnAdvanceDate, vxmlRequestNode)
                Loop
            Next i
            
        Case Else
            'payment method is not set so don't set an advance date
            strRtnAdvanceDate = vbNullString
        End Select
    End If
    
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    Set xmlNodeResp = xmlNode.cloneNode(False)
    xmlChangeNodeName xmlNodeResp, "SETADVANCEDATE", "ADVANCEDATERETURNED"
    xmlNodeResp.Attributes.removeNamedItem ("COMPLETIONDATE")
    xmlNodeResp.Attributes.removeNamedItem ("PAYMENTMETHOD")
    xmlSetAttributeValue xmlNodeResp, "ADVANCEDATE", strRtnAdvanceDate
    vxmlResponseNode.appendChild xmlNodeResp.cloneNode(False)

SetAdvanceDateExit:
    Set xmlNode = Nothing
    Set xmlNodeResp = Nothing
    Set objGlobalParamDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub IPAymentProcessingDO_SetPaymentMethod(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Const cstrFunctionName = "IPAymentProcessingDO_SetPaymentMethod"
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeResp As IXMLDOMNode
    'request params
    'Dim strAppNumber As String
    'Dim strAppFactFindNumber As String
    Dim strAppCompletionDate As String
    Dim strRtnPaymentMethod As String
    'global param
    Dim strPPCutOffTime As String
    Dim objGlobalParamDO As GlobalParameterDO
    
    Dim strToday As String
    Dim lbIsDateWorkingDay As Boolean
    'Dim lbPlus3WorkingDays As Boolean
    Dim lbPlusDays As Boolean
    Dim nBusinessDaysToAdd As Integer 'JD MAR1220
    
On Error GoTo SetPaymentMethodExit
    
    Set objGlobalParamDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")

    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/SETPAYMENTMETHOD")
    'MAR1632 GHun Not used
    'strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    'strAppFactFindNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    
    'MAR1632 GHun Strip any time from the completion date
    strAppCompletionDate = Format$(xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE"), "dd/mm/yyyy")
    strToday = Format$(Now, "dd/mm/yyyy")

    If strAppCompletionDate = strToday Then
        If IsDateWorkingDay(strToday, vxmlRequestNode) Then
            'check if we are within the CHAPS cutoff time
            strPPCutOffTime = objGlobalParamDO.GetCurrentParameterByType("PPCHAPSCutOffTime", "STRING")
            
            If Format$(Now, "hh:mm:ss") < strPPCutOffTime Then
                strRtnPaymentMethod = "YC"  '"CHAPS"
            Else
                errThrowError cstrFunctionName, 8526 '"CHAPS cut off time exceeded"
            End If
        Else
            errThrowError cstrFunctionName, 8527 '"Today is not a working day. Payment method not set"
        End If
    Else
        If IsDateWorkingDay(strToday, vxmlRequestNode) Then
            'check if we are within the BACS cutoff time
            strPPCutOffTime = objGlobalParamDO.GetCurrentParameterByType("PPBACSCutOffTime", "STRING")
            If Format$(Now, "hh:mm:ss") < strPPCutOffTime Then
                lbPlusDays = True
                nBusinessDaysToAdd = objGlobalParamDO.GetCurrentParameterByType("PPNoOfDayInFutureBefBACSCutOff", "AMOUNT") 'JD MAR1220
            Else
                lbPlusDays = True
                nBusinessDaysToAdd = objGlobalParamDO.GetCurrentParameterByType("PPNoOfDayInFutureAftBACSCutOff", "AMOUNT") 'JD MAR1220
            End If
        Else
            'not a working date
            lbPlusDays = True
            nBusinessDaysToAdd = objGlobalParamDO.GetCurrentParameterByType("PPNoOfDaysAftNonBusinessDay", "AMOUNT") 'JD MAR1220
        End If
    End If
    
    'If lbPlus3Days Or lbPlus3WorkingDays Then
    If lbPlusDays Then
        Dim i As Integer
        For i = 1 To nBusinessDaysToAdd  'JD MAR1273
            lbIsDateWorkingDay = False
            Do While Not lbIsDateWorkingDay
                strToday = Format$(DateAdd("d", 1, CDate(strToday)), "dd/mm/yyyy")
                lbIsDateWorkingDay = IsDateWorkingDay(strToday, vxmlRequestNode) 'JD MAR845 send request node not response node
            Loop
        Next i
        If DateDiff("d", CDate(strToday), CSafeDate(strAppCompletionDate)) >= 0 Then  'JD MAR383
            strRtnPaymentMethod = "B"   '"BACS"
        Else
            strRtnPaymentMethod = "YC"  '"CHAPS"
        End If
    End If

    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    Set xmlNodeResp = xmlNode.cloneNode(False)
    xmlChangeNodeName xmlNodeResp, "SETPAYMENTMETHOD", "PAYMENTMETHODRETURNED"
    xmlNodeResp.Attributes.removeNamedItem ("COMPLETIONDATE")
    xmlNodeResp.Attributes.removeNamedItem ("APPLICATIONNUMBER")
    xmlNodeResp.Attributes.removeNamedItem ("APPLICATIONFACTFINDNUMBER")
    xmlSetAttributeValue xmlNodeResp, "PAYMENTMETHOD", strRtnPaymentMethod
    vxmlResponseNode.appendChild xmlNodeResp.cloneNode(False)

SetPaymentMethodExit:
    Set xmlNode = Nothing
    Set xmlNodeResp = Nothing
    Set objGlobalParamDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub IPAymentProcessingDO_ResetPostCompletionTasks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "IPAymentProcessingDO_ResetPostCompletionTasks"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlList As IXMLDOMNodeList

    'global param
    Dim objGlobalParamDO As GlobalParameterDO
    'comp
    Dim objMsgTmBO As MsgTmBO
    Dim strResponse As String
       
    Dim sUpdateProperty As String
    Dim sCompletionDate As String
    Dim sNewCompletionDate As String
    Dim sNewTaskStatusID As String
    
    Dim sStageID As String
    Dim sTaskStatus As String
    Dim sTaskType As String
    Dim nDays As Integer
    Dim sAuthorityLevel As String
    
    Dim intTaskStatus As Integer    ' MAR736
    
On Error GoTo ResetPostCompletionTasksExit

    'form a new request
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlSetAttributeValue xmlRequest, "OPERATION", "FindCaseTaskList"
    
    Set xmlElem = xmlGetMandatoryNode(vxmlRequestNode, "FindCaseTaskList/CASETASK")
    ' get completion date etc
    sAuthorityLevel = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    
    'extract property depending on UPDATEPROPERTY attrib value
    sUpdateProperty = xmlGetMandatoryAttributeText(vxmlRequestNode, "UPDATEPROPERTY")
    Select Case UCase$(sUpdateProperty)
        Case "DUEDATE"
            sCompletionDate = xmlElem.getAttribute("COMPLETIONDATE")
            xmlElem.removeAttribute ("COMPLETIONDATE")
        Case "TASKSTATUS"
            sNewTaskStatusID = xmlGetMandatoryAttributeText(xmlElem, "TASKSTATUSID")
            xmlElem.removeAttribute ("TASKSTATUSID")
    End Select
    
    sStageID = xmlElem.getAttribute("STAGEID")
    sTaskStatus = xmlElem.getAttribute("TASKSTATUS")
    sTaskType = xmlElem.getAttribute("TASKTYPE")
    
    xmlRequest.appendChild xmlElem.cloneNode(True)
    
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    
    strResponse = objMsgTmBO.TmRequest(xmlRequest.xml)

    Dim lngRet As Long
    lngRet = errCheckXMLResponse(strResponse, False)
    
    If lngRet = 0 Then
    
        Set xmlRequest = xmlDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlRequest, "OPERATION", "UpdateCaseTask"
        xmlSetAttributeValue xmlRequest, "USERAUTHORITYLEVEL", sAuthorityLevel
        xmlRequest.Attributes.removeNamedItem ("UPDATEPROPERTY")
        
        xmlDoc.loadXML strResponse
        ' filter by task.status=incomplete and tasktype ='PC' post completed is set in request
        Set xmlList = xmlDoc.selectNodes("RESPONSE/CASETASK")
        
        For Each xmlElem In xmlList
            'update only task.status=incomplete and task.tasktype="PC" and stage=interface
            intTaskStatus = xmlGetAttributeAsInteger(xmlElem, "TASKSTATUS")
                   
            'MAR736  Correct inspection of status and type
            If ((xmlElem.getAttribute("STAGEID") = sStageID) And _
                (xmlElem.getAttribute("TASKTYPE") = sTaskType) And _
                (IsValidationType("TaskStatus", intTaskStatus, sTaskStatus))) Then
    
                Select Case UCase$(sUpdateProperty)
                    Case "DUEDATE"
                        'find a new completion date
                        If (xmlElem.getAttribute("CHASINGPERIODDAYS") <> "") Then
                            nDays = CInt(xmlElem.getAttribute("CHASINGPERIODDAYS"))
                        End If
                        If (nDays <> 0) Then
                            sNewCompletionDate = Format$(DateAdd("d", nDays, CDate(sCompletionDate)), "dd/mm/yyyy")
                        Else
                            sNewCompletionDate = sCompletionDate
                        End If
                        'MAR1547 GHun
                        If Len(sNewCompletionDate) <= 10 Then
                            sNewCompletionDate = sNewCompletionDate & " " & GetGlobalParamString("TMPostCompTaskDueTime")
                        End If
                        'MAR1547 End
                        xmlElem.setAttribute "TASKDUEDATEANDTIME", sNewCompletionDate
                    Case "TASKSTATUS"  'JD MAR803
                        'update status to "30" = Not Aplicable
                        xmlElem.setAttribute "TASKSTATUS", sNewTaskStatusID
                End Select
                
                xmlRequest.appendChild xmlElem
            End If
        Next
        
        'call update task passing in a list
        If xmlRequest.hasChildNodes Then
            strResponse = objMsgTmBO.TmRequest(xmlRequest.xml)
        End If
        
        lngRet = errCheckXMLResponse(strResponse, False)
    
    End If

ResetPostCompletionTasksExit:
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlElem = Nothing
    Set xmlList = Nothing
    
    Set objMsgTmBO = Nothing
    Set objGlobalParamDO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub IPAymentProcessingDO_CreateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationFeeTypeExit

Const strFunctionName As String = "IPAymentProcessingDO_CreateApplicationFeeType"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("APPLICATIONFEETYPE")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateApplicationFeeTypeExit:
    Set xmlSchemaNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPaymentProcessingDO_CreateApplicationFeeTypes(ByVal vxmlApplicationFeesTypeNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationFeeTypesExit

Const strFunctionName As String = "IPaymentProcessingDO_CreateApplicationFeeTypes"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("APPLICATIONFEETYPE")
    adoCreate vxmlApplicationFeesTypeNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateApplicationFeeTypesExit:
    Set xmlSchemaNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreateChequeAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateChequeAuditExit

Const strFunctionName As String = "IPAymentProcessingDO_CreateChequeAudit"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("CHEQUEAUDIT")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateChequeAuditExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateDisbursementExit

Const strFunctionName As String = "IPAymentProcessingDO_CreateDisbursement"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("DISBURSEMENTPAYMENT")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateDisbursementExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

Const strFunctionName As String = "IPAymentProcessingDO_CreateFeeTypePayment"
On Error GoTo CreateFeeTypePaymentExit

Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("FEEPAYMENT")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateFeeTypePaymentExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreateLoanComponentPaymentDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

Const strFunctionName As String = "IPAymentProcessingDO_CreateLoanComponentPaymentDetails"
On Error GoTo CreateLoanComponentPaymentDetailsExit

Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("LOANCOMPONENTPAYMENT")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateLoanComponentPaymentDetailsExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreatePayeeHistoryDetailsExit

Const strFunctionName As String = "IPAymentProcessingDO_CreatePayeeHistoryDetails"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("PAYEEHISTORY")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreatePayeeHistoryDetailsExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreatePaymentAuditExit

Const strFunctionName As String = "IPAymentProcessingDO_CreatePaymentAudit"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("PAYPROCBATCHAUDIT")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreatePaymentAuditExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_CreatePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

Const strFunctionName As String = "IPAymentProcessingDO_CreatePaymentRecord"
On Error GoTo CreatePaymentRecordExit

Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("PAYMENTRECORD")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreatePaymentRecordExit:
    Set xmlSchemaNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_DeleteApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteApplicationFeeTypesExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeleteApplicationFeeTypes"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "APPLICATIONFEETYPE", False)
    gobjContext.SetComplete
    
DeleteApplicationFeeTypesExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_DeleteFeePayments(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo DeleteFeePaymentsExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeleteFeePayments"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "FEEPAYMENT", False)
    gobjContext.SetComplete
    
DeleteFeePaymentsExit:
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IPAymentProcessingDO_DeleteFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteFeeTypePaymentExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeleteFeeTypePayment"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "FEEPAYMENT")
    gobjContext.SetComplete
    
DeleteFeeTypePaymentExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_DeleteLoanComponentPaymentDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteLoanComponentPaymentDetailsExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeleteLoanComponentPaymentDetails"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "LOANCOMPONENTPAYMENT", False)
    gobjContext.SetComplete
    
DeleteLoanComponentPaymentDetailsExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_DeletePayeeHistory( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeletePayeeHistoryExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeletePayeeHistory"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "PAYEEHISTORY", False)
    gobjContext.SetComplete
    
DeletePayeeHistoryExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_DeletePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeletePaymentRecordExit
    
Const strFunctionName As String = "IPAymentProcessingDO_DeletePaymentRecord"
Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "PAYMENTRECORD")
    gobjContext.SetComplete
    
DeletePaymentRecordExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindApplicationThirdPartyList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindApplicationThirdPartyListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindApplicationThirdPartyList"
 
Dim xmlTPNode As IXMLDOMNode

    'Extract data node
    Set xmlTPNode = vxmlRequestNode.selectSingleNode("APPLICATIONTHIRDPARTY")
    adoGetAsXML xmlTPNode, vxmlResponseNode, "FINDAPPLICATIONTHIRDPARTY"
    
    'Check if a record has been found...
    Set xmlTPNode = vxmlResponseNode.selectSingleNode("THIRDPARTY")
    If xmlTPNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
    gobjContext.SetComplete
    
FindApplicationThirdPartyListExit:
    Set xmlTPNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindBatchPaymentList( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo IPAymentProcessingDO_FindBatchPaymentList_Exit

    Const strFunctionName As String = "IPAymentProcessingDO_FindBatchPaymentList"
    
    Dim xmlRequest As IXMLDOMNode

    If vxmlRequestNode.nodeName = "BATCHPAYMENT" Then
        Set xmlRequest = vxmlRequestNode
    Else
        Set xmlRequest = vxmlRequestNode.selectSingleNode("BATCHPAYMENT")
    End If
    
    adoGetAsXML xmlRequest, vxmlResponseNode, "BATCHPAYMENT"
    gobjContext.SetComplete

IPAymentProcessingDO_FindBatchPaymentList_Exit:
    Set xmlRequest = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindDisbursementPaymentSummary(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindDisbursementPaymentSummaryExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindDisbursementPaymentSummary"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "FINDDISBURSEMENTPAYMENTSUMMARY"
    gobjContext.SetComplete
    
FindDisbursementPaymentSummaryExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindPaymentAuditList( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo IPAymentProcessingDO_FindPaymentAuditList_Exit

    Const strFunctionName As String = "IPAymentProcessingDO_FindPaymentAuditList"
    
    Dim xmlRequest As IXMLDOMNode

    If vxmlRequestNode.nodeName = "PAYPROCBATCHAUDIT" Then
        Set xmlRequest = vxmlRequestNode
    Else
        Set xmlRequest = vxmlRequestNode.selectSingleNode("PAYPROCBATCHAUDIT")
    End If
    
    adoGetAsXML xmlRequest, vxmlResponseNode, "PAYPROCBATCHAUDIT"
    gobjContext.SetComplete

IPAymentProcessingDO_FindPaymentAuditList_Exit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindDisbursementPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindDisbursementPaymentListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindDisbursementPaymentList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "DISBURSEMENTPAYMENT"
    gobjContext.SetComplete
    
FindDisbursementPaymentListExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindFeePaymentList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindFeePaymentListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindFeePaymentList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "FEEPAYMENT"
    gobjContext.SetComplete
    
FindFeePaymentListExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindFeeTypeList(ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindFeeTypeListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindFeeTypeList"
Dim xmlAppFeeNode As IXMLDOMNode
    
    'Extract Payment node
    Set xmlAppFeeNode = vxmlRequestNode.selectSingleNode("APPLICATIONFEETYPE")
    adoGetAsXML xmlAppFeeNode, vxmlResponseNode, "APPLICATIONFEETYPE"
    
    'Check if a record has been found...
    Set xmlAppFeeNode = vxmlResponseNode.selectSingleNode("APPLICATIONFEETYPE")
    If xmlAppFeeNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
    gobjContext.SetComplete
    
FindFeeTypeListExit:
    Set xmlAppFeeNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindLoanComponentAllocatedAmount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindLoanComponentAllocatedAmountExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindLoanComponentAllocatedAmount"
    
Dim xmlLCNode As IXMLDOMNode

    'Extract data node
    Set xmlLCNode = vxmlRequestNode.selectSingleNode("GETLCALLOCATEDAMOUNT")
    adoGetAsXML xmlLCNode, vxmlResponseNode, "GETLCALLOCATEDAMOUNT"
    gobjContext.SetComplete
    
FindLoanComponentAllocatedAmountExit:
    Set xmlLCNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPaymentProcessingDO_FindLoanComponentPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindLoanComponentPaymentListExit
    
Const strFunctionName As String = "IPaymentProcessingDO_FindLoanComponentPaymentList"
    
Dim xmlNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode

Dim strValue As String
Dim intDay As Integer

    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "FINDLOANCOMPONENTPAYMENTLIST"
    Set xmlNode = vxmlResponseNode.selectSingleNode("LOANCOMPONENTPAYMENT")
    If xmlNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
        
    strValue = xmlGetAttributeText(xmlNode, "FIRSTPAYMENTDATE")
    If strValue <> "" Then
        intDay = DatePart("d", strValue)
        xmlSetAttributeValue xmlNode, "FIRSTPAYMENTDAY", intDay
    Else
        xmlSetAttributeValue xmlNode, "FIRSTPAYMENTDAY", ""
    End If
        
    gobjContext.SetComplete
    
FindLoanComponentPaymentListExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindPayeeHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode, Optional ByVal blnErrorIfNotFound As Boolean = True)

On Error GoTo FindPayeeHistoryListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindPayeeHistoryList"
Dim xmlPayeeHistNode As IXMLDOMNode
    
    'Extract PAYEEHISTORY node
    If vxmlRequestNode.nodeName = "PAYEEHISTORY" Then
        Set xmlPayeeHistNode = vxmlRequestNode
    Else
        Set xmlPayeeHistNode = vxmlRequestNode.selectSingleNode("PAYEEHISTORY")
    End If
    adoGetAsXML xmlPayeeHistNode, vxmlResponseNode, "FINDPAYEEHISTORY"
    
    'Check if a record has been found...
    If blnErrorIfNotFound Then
        Set xmlPayeeHistNode = vxmlResponseNode.selectSingleNode("PAYEEHISTORY")
        If xmlPayeeHistNode Is Nothing Then
            '... and if not raise an error
            errThrowError strFunctionName, oeRecordNotFound
        End If
    End If
    
    gobjContext.SetComplete
    
FindPayeeHistoryListExit:
    Set xmlPayeeHistNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindPaymentAllocatedAmount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindPaymentAllocatedAmountExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindPaymentAllocatedAmount"
 
Dim xmlPRNode As IXMLDOMNode

    'Extract data node
    Set xmlPRNode = vxmlRequestNode.selectSingleNode("GETPRALLOCATEDAMOUNT")
    adoGetAsXML xmlPRNode, vxmlResponseNode, "GETPRALLOCATEDAMOUNT"
    gobjContext.SetComplete
    
FindPaymentAllocatedAmountExit:
    Set xmlPRNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindPaymentHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindFeeTypeListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindPaymentHistoryList"
Dim xmlPaymentNode As IXMLDOMNode
    
    'Extract Payment node
    If vxmlRequestNode.nodeName = "PAYMENTRECORD" Then
        Set xmlPaymentNode = vxmlRequestNode
    Else
        Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    End If
    adoGetAsXML xmlPaymentNode, vxmlResponseNode, "PAYMENTRECORD"
    
    'Check if a record has been found...
    Set xmlPaymentNode = vxmlResponseNode.selectSingleNode("PAYMENTRECORD")
    If xmlPaymentNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
    gobjContext.SetComplete
    
FindFeeTypeListExit:
    Set xmlPaymentNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub IPAymentProcessingDO_FindPaymentRecordList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindPaymentRecordListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindPaymentRecordList"
    
Dim strFilter As String
Dim strValue As String
Dim strValidType As String
Dim strValueId As String
Dim intStart As Integer
Dim intPos As Integer

Dim xmlElement As IXMLDOMElement
Dim colValueIds As Collection           ' PSC 18/11/2002 BMIDS00968
Dim varValueId As Variant               ' PSC 18/11/2002 BMIDS00968

    Set xmlElement = vxmlRequestNode.cloneNode(True)
    
    'If PAYMENTMETHOD specified...
    strValue = xmlGetAttributeText(vxmlRequestNode, "PAYMENTMETHOD")
    If Len(Trim$(strValue)) > 0 Then
        intStart = 1
        intPos = 1
        Do Until intPos = 0
            ' PSC 18/11/2002 BMIDS00968
            Set colValueIds = New Collection
        
            intPos = InStr(intStart, strValue, ",")
            If intPos > 0 Then
                strValidType = Mid$(strValue, intStart, intPos - intStart)
            Else
                strValidType = Right$(strValue, Len(strValue) - intStart + 1)
            End If
            
            ' PSC 18/11/2002 BMIDS00968 - Start
            GetValueIdsForValidationType "PaymentMethod", strValidType, colValueIds
        
            For Each varValueId In colValueIds
                If Len(strFilter) > 0 Then
                    strFilter = strFilter & ","
                Else
                    strFilter = strFilter & " (PAYMENTMETHOD IN("
                End If
                strFilter = strFilter & CStr(varValueId)
            Next
            ' PSC 18/11/2002 BMIDS00968 - End
    
            If intStart < Len(strValue) Then
                intStart = intPos + 1
            Else
                intPos = 0
            End If
            
            ' PSC 18/11/2002 BMIDS00968
            Set colValueIds = Nothing
        Loop
        strFilter = strFilter & "))"
        'Remove the PAYMENTMETHOD attribute
        xmlElement.removeAttribute "PAYMENTMETHOD"
    End If
    
    adoGetAsXML xmlElement, vxmlResponseNode, "PAYMENTRECORD", strFilter
    
    gobjContext.SetComplete
    
FindPaymentRecordListExit:
    Set xmlElement = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Function GetValueIdForValidationType(ByVal strComboGroup As String, ByVal strValidationType As String) As String

On Error GoTo GetValueIdForValidationTypeExit

Const strFunctionName As String = "GetValueIdForValidationType"

Dim colValueId As Collection

    Set colValueId = New Collection
    GetValueIdsForValidationType strComboGroup, strValidationType, colValueId
    If colValueId.Count > 0 Then
        GetValueIdForValidationType = CStr(colValueId.Item(1))
    Else
        errThrowError strFunctionName, oePPComboValueIdNotFound, "PaymentMethod"
    End If

GetValueIdForValidationTypeExit:
    Set colValueId = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function IsDateWorkingDay(ByVal strDate As String, ByVal vxmlRequestNode As IXMLDOMNode) As Boolean

    Const cstrFunctionName As String = "IsDateWorkingDay"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlChild As IXMLDOMNode
    Dim objOrgBO As OrganisationBO
    
    Dim strResponse  As String
    Dim strIsWorkingDate As String

    On Error GoTo IsDateWorkingDayExit
    
    IsDateWorkingDay = False
    
    'form a new request
    Set xmlDoc = xmlCreateDOMObject
    
    Set objOrgBO = gobjContext.CreateInstance("omOrg.OrganisationBO")
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlSetAttributeValue xmlRequest, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    
    Set xmlNode = xmlDoc.createElement("SEARCH")
    Set xmlChild = xmlDoc.createElement("CHANNELID")
    xmlChild.Text = xmlGetMandatoryAttributeText(vxmlRequestNode, "CHANNELID")
    xmlNode.appendChild xmlChild.cloneNode(True)
    Set xmlChild = xmlDoc.createElement("DATE")
    xmlChild.Text = strDate
    xmlNode.appendChild xmlChild.cloneNode(True)
    xmlRequest.appendChild xmlNode.cloneNode(True)
    
    'call IsWorkingDate
    strResponse = objOrgBO.IsDateWorkingDay(xmlRequest.xml)

    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    strIsWorkingDate = xmlGetNodeText(xmlDoc.documentElement, "STATUS")  '<RESPONSE TYPE="SUCCESS"><STATUS>0</STATUS></RESPONSE>
    
    If strIsWorkingDate <> "0" Then
        IsDateWorkingDay = True
    End If
    
IsDateWorkingDayExit:

    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlNode = Nothing
    Set xmlChild = Nothing
    Set objOrgBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Sub IPAymentProcessingDO_FindPaymentsForBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPaymentsForBatchExit

    Const cstrFunctionName As String = "IPAymentProcessingDO_FindPaymentsForBatch"
    
    Dim objParam As ADODB.Parameter
    Dim objCommand As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim nodeBatch As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objBatch As BatchScheduleBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim colValueIds As Collection               ' PSC 15/11/2002 BMIDS00968
    
    Dim strSQL As String
    Dim strValue As String
    Dim strPaymentStatus As String
    Dim strJobTypePSP As String
    Dim strPayMethod As String
    Dim strJobTypeIPP As String
    Dim strPayTypeBCP As String
    Dim blnSelected As Boolean
    Dim strResponse As String
    Dim varValueId As Variant                   ' PSC 15/11/2002 BMIDS00968
    Dim lngColIndex As Long                     ' PSC 15/11/2002 BMIDS00968

    Set objCommand = New ADODB.Command
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set colValueIds = New Collection            ' PSC 15/11/2002 BMIDS00968
    
    
    ' PSC 20/09/2002 BMIDS00480 - Added payment type
    strSQL = "select pr.applicationnumber, pr.paymentsequencenumber, dp.paymenttype "
    strSQL = strSQL & "From paymentrecord pr, disbursementpayment dp "
    strSQL = strSQL & "Where pr.applicationnumber = dp.applicationnumber "
    strSQL = strSQL & "and pr.paymentsequencenumber = dp.paymentsequencenumber"
    
    ' RF 13/02/02 Removed obsolete, commented-out code
    
    Set nodeBatch = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    
    'APPLICATIONNUMBERFROM
    strValue = xmlGetAttributeText(nodeBatch, "APPNUMRANGESTART")
    If Len(Trim$(strValue)) > 0 Then
        'SYS2997 MDC 20/12/2001. Generic SQL change
        #If GENERIC_SQL Then
            strSQL = strSQL & " and pr.applicationnumber >= '" & strValue & "'"
        #Else
            Set objParam = New ADODB.Parameter
            strSQL = strSQL & " and pr.applicationnumber >= ?"
            objParam.Type = adBSTR
            objParam.Direction = adParamInput
            objParam.Value = strValue
            objCommand.Parameters.Append objParam
            Set objParam = Nothing
        #End If
        'SYS2997 End
    End If
    
    'APPLICATIONNUMBERTO
    strValue = xmlGetAttributeText(nodeBatch, "APPNUMRANGEEND")
    If Len(Trim$(strValue)) > 0 Then
        'SYS2997 MDC 05/12/2001. Generic SQL change
        #If GENERIC_SQL Then
            strSQL = strSQL & " and pr.applicationnumber <= '" & strValue & "'"
        #Else
            Set objParam = New ADODB.Parameter
            strSQL = strSQL & " and pr.applicationnumber <= ?"
            objParam.Type = adBSTR
            objParam.Direction = adParamInput
            objParam.Value = strValue
            objCommand.Parameters.Append objParam
            Set objParam = Nothing
        #End If
        'SYS2997 End
    End If
    
    '------------------------------------------------------------------------------------------
    'PAYMENTMETHOD
    '------------------------------------------------------------------------------------------
    
    'BACS
    blnSelected = xmlGetMandatoryAttributeAsBoolean(nodeBatch, "BACS")
    If blnSelected Then
        
        ' PSC 15/11/2002 BMIDS00968 - Start
        ' Get all BACS value ids
        GetValueIdsForValidationType "PaymentMethod", "B", colValueIds
        
        ' Add each one into the list
        For Each varValueId In colValueIds
        
            If Len(strPayMethod) > 0 Then
                strPayMethod = strPayMethod & ", "
            End If

            strPayMethod = strPayMethod & varValueId
        
        Next
        
        ' Clear down the collection
        For lngColIndex = 1 To colValueIds.Count
            colValueIds.Remove (1)
        Next
        ' PSC 15/11/2002 BMIDS00968 - Start

    End If
    
    'ChapsOrTT
    blnSelected = xmlGetMandatoryAttributeAsBoolean(nodeBatch, "CHAPSORTT")
    If blnSelected Then
    
        ' PSC 15/11/2002 BMIDS00968 - Start
        ' Get all CHAPS value ids
        GetValueIdsForValidationType "PaymentMethod", "YC", colValueIds
        
        ' Add each one into the list
        For Each varValueId In colValueIds
        
            If Len(strPayMethod) > 0 Then
                strPayMethod = strPayMethod & ", "
            End If

            strPayMethod = strPayMethod & varValueId
        
        Next
        
        ' Clear down the collection
        For lngColIndex = 1 To colValueIds.Count
            colValueIds.Remove (1)
        Next
        ' PSC 15/11/2002 BMIDS00968 - End
    
    End If
    
    'Cheque
    blnSelected = xmlGetMandatoryAttributeAsBoolean(nodeBatch, "CHEQUE")
    If blnSelected Then
    
        ' PSC 15/11/2002 BMIDS00968 - Start
        ' Get all Cheque value ids
        GetValueIdsForValidationType "PaymentMethod", "CH", colValueIds
        
        ' Add each one into the list
        For Each varValueId In colValueIds
        
            If Len(strPayMethod) > 0 Then
                strPayMethod = strPayMethod & ", "
            End If

            strPayMethod = strPayMethod & varValueId
        
        Next
        
        ' Clear down the collection
        For lngColIndex = 1 To colValueIds.Count
            colValueIds.Remove (1)
        Next
        ' PSC 15/11/2002 BMIDS00968 - End
        
    End If
    
    'Internal Transfer
    blnSelected = xmlGetMandatoryAttributeAsBoolean(nodeBatch, "INTERNALTRANSFER")
    If blnSelected Then
    
        ' PSC 15/11/2002 BMIDS00968 - Start
        ' Get all Cheque value ids
        GetValueIdsForValidationType "PaymentMethod", "TR", colValueIds
        
        ' Add each one into the list
        For Each varValueId In colValueIds
        
            If Len(strPayMethod) > 0 Then
                strPayMethod = strPayMethod & ", "
            End If

            strPayMethod = strPayMethod & varValueId
        
        Next
        
        ' Clear down the collection
        For lngColIndex = 1 To colValueIds.Count
            colValueIds.Remove (1)
        Next
        ' PSC 15/11/2002 BMIDS00968 - End
    
    End If
    
    If Len(strPayMethod) > 0 Then
        strSQL = strSQL & " and pr.paymentmethod in (" & strPayMethod & ")"
    End If
    
    'ISSUEDATE
    strValue = xmlGetAttributeText(nodeBatch, "ISSUEDATE")
    If Len(Trim$(strValue)) > 0 Then
        'SYS2997 MDC 20/12/2001. Generic SQL change
        #If GENERIC_SQL Then
            strSQL = strSQL & " and dp.issuedate <= '" & Format$(strValue, "dd-mmm-yyyy") & "'"
        #Else
            Set objParam = New ADODB.Parameter
            strSQL = strSQL & " and dp.issuedate <= ?"
            objParam.Type = adDBTimeStamp
            objParam.Direction = adParamInput
            objParam.Value = strValue
            objCommand.Parameters.Append objParam
            Set objParam = Nothing
        #End If
        'SYS2997 End
    End If

    'JD MAR1245 check advance date <= today
    #If GENERIC_SQL Then
        strSQL = strSQL & " and dp.completiondate <= '" & Format$(Now, "dd-mmm-yyyy") & "'"
    #Else
        Set objParam = New ADODB.Parameter
        strSQL = strSQL & " and dp.completiondate <= ?"
        objParam.Type = adDBTimeStamp
        objParam.Direction = adParamInput
        objParam.Value = Format$(Now, "dd-mmm-yyyy")
        objCommand.Parameters.Append objParam
        Set objParam = Nothing
    #End If
    
'    Set nodeBatch = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTPROCESSING")
    
    'Get Value Id's from Validation Types for PaymentJobTypes
    strJobTypePSP = GetValueIdForValidationType("PaymentJobType", "P")  'Pay Sanctioned Payment
    strJobTypeIPP = GetValueIdForValidationType("PaymentJobType", "I")  'Interface paid Payment
    
    'PAYMENTTYPE
    strValue = xmlGetAttributeText(nodeBatch, "PAYMENTJOBTYPE")
    If CSafeInt(strValue) = strJobTypePSP Then
        'Pay Sanctioned Payment
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "S")   'Sanctioned
        strPayTypeBCP = GetValueIdForValidationType("PaymentType", "NCB")   'Balance Cancellation Payment
        strSQL = strSQL & " and dp.paymentstatus = " & strPaymentStatus
        
        ' PSC 08/10/2002 BMIDS00595
        strSQL = strSQL & " and dp.paymenttype != " & strPayTypeBCP
    ElseIf CSafeInt(strValue) = strJobTypeIPP Then
        'Interface Paid Payment
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "P")   'Paid
        
        'PSC 18/09/2002 SYS5181 - Start
        strSQL = strSQL & " and (dp.paymentstatus = " & strPaymentStatus
        strPaymentStatus = GetValueIdForValidationType("PaymentStatus", "IF")   'Failed
        strSQL = strSQL & " or dp.paymentstatus = " & strPaymentStatus & ")"
        'PSC 18/09/2002 SYS5181 - End
        
        ' PSC 13/11/2002 BMIDS00915
        strSQL = strSQL & " AND NOT EXISTS (SELECT 'X' FROM PAYMENTRECORD " & _
                                           "WHERE APPLICATIONNUMBER = DP.APPLICATIONNUMBER " & _
                                           "AND ASSOCPAYSEQNUMBER = DP.PAYMENTSEQUENCENUMBER)"

    End If
    
    'Execute the Command SQL
    objCommand.CommandType = adCmdText
    objCommand.CommandText = strSQL
    Debug.Print strSQL
    Set objRs = executeGetRecordSet(objCommand)

    'Convert Recordset to XML Response
    If Not objRs Is Nothing Then
        If Not (objRs.BOF Or objRs.EOF) Then
            objRs.MoveFirst
            Do Until objRs.EOF
                Set xmlElement = xmlDoc.createElement("PAYMENTRECORD")
                xmlElement.setAttribute "APPLICATIONNUMBER", objRs.Fields("APPLICATIONNUMBER")
                xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", objRs.Fields("PAYMENTSEQUENCENUMBER")
                
                ' PSC 20/09/2002 BMIDS00480
                xmlElement.setAttribute "PAYMENTTYPE", objRs.Fields("PAYMENTTYPE")

                vxmlResponseNode.appendChild xmlElement
                objRs.MoveNext
            Loop
        Else
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
    Else
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindPaymentsForBatchExit:

    Set objParam = Nothing
    Set objCommand = Nothing
    Set objRs = Nothing
    Set nodeBatch = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    Set colValueIds = Nothing                   ' PSC 15/11/2002 BMIDS00968
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub IPAymentProcessingDO_FindPaymentsForPayeeList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindPaymentsForPayeeListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindPaymentsForPayeeList"
Dim xmlPayeeHistoryNode As IXMLDOMNode
    
    'Extract PAYEEHISTORY node
    Set xmlPayeeHistoryNode = vxmlRequestNode.selectSingleNode("PAYEEHISTORY")
    adoGetAsXML xmlPayeeHistoryNode, vxmlResponseNode, "FINDPAYMENTSFORPAYEE"
    
    'Check if a record has been found...
    Set xmlPayeeHistoryNode = vxmlResponseNode.selectSingleNode("DISBURSEMENTPAYMENT")
    If xmlPayeeHistoryNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
    gobjContext.SetComplete
    
FindPaymentsForPayeeListExit:
    Set xmlPayeeHistoryNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_FindSanctioningList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo IPAymentProcessingDO_FindSanctioningList_Exit

Const strFunctionName As String = "IPAymentProcessingDO_FindSanctioningList"

Dim xmlElement As IXMLDOMElement
Dim xmlPaymentNode As IXMLDOMNode

Dim strValue As String
Dim strFilter As String
Dim strValidType As String
Dim strValueId As String
Dim intStart As Integer
Dim intPos As Integer


    Set xmlElement = vxmlRequestNode.cloneNode(True)

    'If PAYMENTMETHOD specified...
    strValue = xmlGetAttributeText(xmlElement, "PAYMENTMETHOD")
    If Len(Trim$(strValue)) > 0 Then
' AQR SYS2797 DRC
'        intStart = 1
'        intPos = 1
'        Do Until intPos = 0
'            intPos = InStr(intStart, strValue, ",")
'            If intPos > 0 Then
'                strValidType = Mid$(strValue, intStart, intPos - intStart)
'            Else
'                strValidType = Right$(strValue, Len(strValue) - intStart + 1)
'            End If
'            strValueId = GetValueIdForValidationType("PaymentMethod", strValidType)
'            If Len(strFilter) > 0 Then
'                strFilter = strFilter & " OR "
'            Else
'                strFilter = "("
'            End If
'            strFilter = strFilter & "PaymentMethod = " & strValueId
'            If intStart < Len(strValue) Then
'                intStart = intPos + 1
'            Else
'                intPos = 0
'            End If
'        Loop
'        strFilter = strFilter & ")"
' END - AQR SYS2797 DRC
        strFilter = "(PaymentMethod = " & strValue & ")"
        'Remove the PAYMENTMETHOD attribute
        xmlElement.removeAttribute "PAYMENTMETHOD"
    End If
    
    'If ISSUEDATE is specified...
    strValue = xmlGetAttributeText(xmlElement, "ISSUEDATE")
    If Len(Trim$(strValue)) > 0 And IsDate(strValue) Then
        strValue = Format$(strValue, "dd-MMM-yyyy")
        If Len(strFilter) > 0 Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & "(IssueDate <= '" & strValue & "')"
        'Remove the ISSUEDATE attribute
        xmlElement.removeAttribute "ISSUEDATE"
    End If
    
    'SYS2279 & SYS2280 CL 26/04/2001
    'If PAYMENTSTATUS is specified...
    strValue = xmlGetAttributeText(xmlElement, "PAYMENTSTATUS")
    
    'SR 06/07/01 : SYS2412
    If Len(strFilter) > 0 Then
        strFilter = strFilter & " AND "
    End If
    'SR 06/07/01 End
    
    If Len(Trim$(strValue)) > 0 Then
        strFilter = strFilter & " (PaymentStatus = " & strValue & ")"
    Else
         strFilter = strFilter & " (PaymentStatus = 10 or PaymentStatus = 20) and (CancellationDate IS NULL)" 'BMIDS665 KRW 04/11/04
     End If
    'BMIDS679 Call Below now orders on Applicationno and paymenttype KRW 27/01/04'
    
    adoGetAsXML xmlElement, vxmlResponseNode, "SANCTIONINGLIST", strFilter, "APPLICATIONNUMBER,PAYMENTTYPE"
    
    Set xmlPaymentNode = vxmlResponseNode.selectSingleNode("PAYMENT")
    
    If xmlPaymentNode Is Nothing Then
        '... and if not raise an error
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    gobjContext.SetComplete


IPAymentProcessingDO_FindSanctioningList_Exit:
    Set xmlElement = Nothing
    Set xmlPaymentNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub IPAymentProcessingDO_GetBatchParameters(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo IPAymentProcessingDO_GetBatchParameters_Exit

Const strFunctionName As String = "IPAymentProcessingDO_GetBatchParameters"

Dim xmlParamNode As IXMLDOMNode

    'Extract BATCHPARAMETERS node
    If vxmlRequestNode.nodeName = "BATCHPARAMETERS" Then
        Set xmlParamNode = vxmlRequestNode
    Else
        Set xmlParamNode = vxmlRequestNode.selectSingleNode("BATCHPARAMETERS")
    End If
    
    adoGetAsXML xmlParamNode, vxmlResponseNode, "PAYPROCBATCHPARAMS"
    gobjContext.SetComplete
    
IPAymentProcessingDO_GetBatchParameters_Exit:
    Set xmlParamNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_GetDisbursementPayments(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetDisbursementPaymentsExit

Const strFunctionName As String = "IPAymentProcessingDO_GetDisbursementPayments"
Dim xmlDPNode As IXMLDOMNode

    'Extract GETDISBURSEMENTPAYMENTS node
    If vxmlRequestNode.nodeName = "GETDISBURSEMENTPAYMENTS" Then
        Set xmlDPNode = vxmlRequestNode
    Else
        Set xmlDPNode = vxmlRequestNode.selectSingleNode("GETDISBURSEMENTPAYMENTS")
    End If
    
    adoGetAsXML xmlDPNode, vxmlResponseNode, "GETDISBURSEMENTPAYMENTS"
    gobjContext.SetComplete
    
GetDisbursementPaymentsExit:
    Set xmlDPNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_GetProductSwitchFee(ByVal vxmlRequestNode As MSXML2.IXMLDOMNode, ByVal vxmlResponseNode As MSXML2.IXMLDOMNode)
    
On Error GoTo GetProductSwitchFee_Exit

    Const strFunctionName As String = "IPAymentProcessingDO_GetProductSwitchFee"
    
    Dim objCommand          As Command
    Dim objParam            As Parameter
    Dim sApplicationNumber  As String
    Dim lAppFactFindNumber  As Long
    Dim rstData             As Recordset
    Dim TempXML             As FreeThreadedDOMDocument40
    Dim sTemp               As String
    Dim xmlEleTemp          As IXMLDOMNode
    Dim dAmount             As Double
    
    Set TempXML = New FreeThreadedDOMDocument40
    TempXML.validateOnParse = False
    TempXML.setProperty "NewParser", True
    Set objCommand = New Command

    sApplicationNumber = xmlGetNodeText(vxmlRequestNode, "APPLICATIONNUMBER")
    lAppFactFindNumber = Val(xmlGetNodeText(vxmlRequestNode, "APPLICATIONFACTFINDNUMBER"))
    
    If xmlGetNode(vxmlResponseNode, "PRODUCTSWITCHFEE") Is Nothing Then
        Set xmlEleTemp = vxmlResponseNode.appendChild(TempXML.createElement("PRODUCTSWITCHFEE"))
    Else
        Set xmlEleTemp = vxmlResponseNode.selectSingleNode("PRODUCTSWITCHFEE")
    End If
    
    If xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
        Call xmlEleTemp.appendChild(TempXML.createElement("AMOUNT"))
    End If
    
    With objCommand
        .CommandText = "USP_GetProductSwitchFeeAmount"
        .CommandType = adCmdStoredProc
        
        Set objParam = objCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(sApplicationNumber), sApplicationNumber)
        Call .Parameters.Append(objParam)
        
        Set objParam = objCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, 5, lAppFactFindNumber)
        Call .Parameters.Append(objParam)
    End With
    '*=[MC]Execute Stored Procedure
    'Set rstData = executeGetRecordSet(objCommand)
    '*=Extended method call, due to nextRecordset functionality.
    Set rstData = executeGetRecordSetEx(objCommand)
    If Not rstData Is Nothing Then
        If Not rstData.EOF And Not rstData.BOF Then
            If Not xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
                dAmount = Val(rstData.Fields("Amount").Value & "")
                'xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount
            End If
        End If
    End If
    '*=Set the Amount
    xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount
    
    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            rstData.Close
        End If
    End If

GetProductSwitchFee_Exit:
    
    Set rstData = Nothing
    Set objParam = Nothing
    Set objCommand = Nothing
      
   errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IPAymentProcessingDO_GetTTFeeAmount(ByVal vxmlRequestNode As MSXML2.IXMLDOMNode, ByVal vxmlResponseNode As MSXML2.IXMLDOMNode)
    
On Error GoTo GetTTFeeAmount_Exit
    
    Const strFunctionName As String = "IPAymentProcessingDO_GetTTFeeAmount"
    
    Dim objCommand          As Command
    Dim objDBConn           As Connection
    Dim objParam            As Parameter
    Dim sApplicationNumber  As String
    Dim lAppFactFindNumber  As Long
    Dim rstData             As Recordset
    Dim TempXML             As FreeThreadedDOMDocument40
    Dim sTemp               As String
    Dim xmlEleTemp          As IXMLDOMNode
    Dim dAmount             As Double
    
    Set TempXML = New FreeThreadedDOMDocument40
    TempXML.validateOnParse = False
    TempXML.setProperty "NewParser", True
    Set objCommand = New Command
    
    If xmlGetNode(vxmlResponseNode, "TTFEESET") Is Nothing Then
        Set xmlEleTemp = vxmlResponseNode.appendChild(TempXML.createElement("TTFEESET"))
    Else
        Set xmlEleTemp = vxmlResponseNode.selectSingleNode("TTFEESET")
    End If
    
    If xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
        Call xmlEleTemp.appendChild(TempXML.createElement("AMOUNT"))
    End If

    sApplicationNumber = xmlGetNodeText(vxmlRequestNode, "APPLICATIONNUMBER")
    lAppFactFindNumber = Val(xmlGetNodeText(vxmlRequestNode, "APPLICATIONFACTFINDNUMBER"))
    
        
    With objCommand
        .CommandText = "USP_GetTTFeeAmount"
        .CommandType = adCmdStoredProc
        Set objParam = objCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(sApplicationNumber), sApplicationNumber)
        Call .Parameters.Append(objParam)
        
        Set objParam = objCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, 5, lAppFactFindNumber)
        Call .Parameters.Append(objParam)
        
    End With
    
    '*=[MC]Extended method called, result may returned in rst.nextRecordset
    Set rstData = executeGetRecordSetEx(objCommand)
    
    If Not rstData Is Nothing Then
        If Not rstData.EOF And Not rstData.BOF Then
            
            If Not xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
                dAmount = Val(rstData.Fields("Amount").Value & "")
                xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount 'rstData.Fields("Amount").Value
            End If
        End If
    Else
        xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount
    End If
    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            rstData.Close
        End If
    End If
    
GetTTFeeAmount_Exit:
    
    Set rstData = Nothing
    Set objParam = Nothing
    Set objCommand = Nothing
    Set objDBConn = Nothing
      
   errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo UpdateFeeTypePaymentExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdateApplicationFeeType"
    
    adoUpdateFromNode vxmlRequestNode, "APPLICATIONFEETYPE"
    gobjContext.SetComplete
    
UpdateFeeTypePaymentExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo UpdateDisbursementExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdateDisbursement"
    
Dim xmlDPNode As IXMLDOMNode

    If vxmlRequestNode.nodeName = "DISBURSEMENTPAYMENT" Then
        Set xmlDPNode = vxmlRequestNode
    Else
        Set xmlDPNode = vxmlRequestNode.selectSingleNode("DISBURSEMENTPAYMENT")
    End If
    
    adoUpdateFromNode xmlDPNode, "DISBURSEMENTPAYMENT"
    gobjContext.SetComplete
    
UpdateDisbursementExit:
    Set xmlDPNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo UpdateFeeTypePaymentExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdateFeeTypePayment"
    
    adoUpdateFromNode vxmlRequestNode, "FEEPAYMENT"
    gobjContext.SetComplete
    
UpdateFeeTypePaymentExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdateLoanComponentPayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateLoanComponentPayment_Exit

Const strFunctionName As String = "IPAymentProcessingDO_UpdateLoanComponentPayment"

    adoUpdateFromNode vxmlRequestNode, "LOANCOMPONENTPAYMENT"
    gobjContext.SetComplete

UpdateLoanComponentPayment_Exit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdatePayeeHistoryDetailsExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdatePayeeHistoryDetails"
    
    adoUpdateFromNode vxmlRequestNode, "PAYEEHISTORY"
    gobjContext.SetComplete
    
UpdatePayeeHistoryDetailsExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo IPAymentProcessingDO_UpdatePaymentAudit_Exit

Const strFunctionName As String = "IPAymentProcessingDO_UpdatePaymentAudit"

    adoUpdateFromNode vxmlRequestNode, "PAYPROCBATCHAUDIT"
    gobjContext.SetComplete

IPAymentProcessingDO_UpdatePaymentAudit_Exit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPAymentProcessingDO_UpdatePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo UpdatePaymentRecordExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdatePaymentRecord"
    
    adoUpdateFromNode vxmlRequestNode, "PAYMENTRECORD"
    gobjContext.SetComplete
    
UpdatePaymentRecordExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

' PSC 19/09/2002 SYS4863 New function
Private Sub IPAymentProcessingDO_UpdateFeePaymentList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdateFeePaymentListExit
    
Const strFunctionName As String = "IPAymentProcessingDO_UpdateFeePaymentList"
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Set xmlNodeList = vxmlRequestNode.selectNodes("FEEPAYMENT")
    
    adoUpdateFromNodeList xmlNodeList, "FEEPAYMENT"
    gobjContext.SetComplete
    
UpdateFeePaymentListExit:
    errCheckError strFunctionName, TypeName(Me)

End Sub

' PSC 20/09/2002 BMIDS00480 - Start
Private Sub IPAymentProcessingDO_FindInterfacedInitialAdvances(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindInterfacedInitialAdvancesExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindInterfacedInitialAdvances"
    
    Dim xmlNode As IXMLDOMNode
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "INTERFACEDINITIALADVANCE"
    
    Set xmlNode = xmlGetNode(vxmlResponseNode, "INTERFACEDINITIALADVANCE")
    
    If xmlNode Is Nothing Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    gobjContext.SetComplete
    
FindInterfacedInitialAdvancesExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub
' PSC 20/09/2002 BMIDS00480 - End

' PSC 20/19/2002 BMIDS01013 - Start
Private Sub IPAymentProcessingDO_FindROFTotals(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindROFTotalsExit
    
Const strFunctionName As String = "IPAymentProcessingDO_FindROFTotals"
    
    Dim xmlNode As IXMLDOMNode
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ROFTOTAL"
    
    Set xmlNode = xmlGetNode(vxmlResponseNode, "ROFTOTAL")
    
    If xmlNode Is Nothing Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    gobjContext.SetComplete
    
FindROFTotalsExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub
' PSC 20/11/2002 BMIDS01013 - End

'MAR28 New function to get Valuation Fee
Private Sub IPAymentProcessingDO_GetValuationFee(ByVal vxmlRequestNode As MSXML2.IXMLDOMNode, _
        ByVal vxmlResponseNode As MSXML2.IXMLDOMNode, Optional ByVal blnRefund As Boolean = False)

' This function uses a stored procedure to get the Valuation Fee
' If the Refund parameter is set to TRUE, the Refund amount is returned (ie the Valuation amount for Scheme 1)

On Error GoTo GetValuationFee_Exit

    Const strFunctionName As String = "IPAymentProcessingDO_GetValuationFee"
    
    Dim objComboDO As ComboDO

    Dim objCommand             As Command
    Dim objParam               As Parameter
    Dim sApplicationNumber     As String
    Dim lAppFactFindNumber     As Long
    Dim lMSQNumber             As Long
    Dim rstData                As Recordset
    Dim TempXML                As FreeThreadedDOMDocument40
    Dim sTemp                  As String
    Dim xmlEleTemp             As IXMLDOMNode
    Dim dAmount                As Double
    Dim dRefundAmount          As Double
    Dim strCommandText         As String
    Dim strRefundValuationType As String
    
    Set TempXML = New FreeThreadedDOMDocument40
    TempXML.validateOnParse = False
    TempXML.setProperty "NewParser", True
    Set objCommand = New Command

    If gobjContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    dAmount = 0
    dRefundAmount = 0

    sApplicationNumber = xmlGetNodeText(vxmlRequestNode, "APPLICATIONNUMBER")
    lAppFactFindNumber = Val(xmlGetNodeText(vxmlRequestNode, "APPLICATIONFACTFINDNUMBER"))
    lMSQNumber = Val(xmlGetNodeText(vxmlRequestNode, "MORTGAGESUBQUOTENUMBER"))
    
    strRefundValuationType = objComboDO.GetFirstComboValueId("ValuationType", "SC1")

    If xmlGetNode(vxmlResponseNode, "VALUATIONFEE") Is Nothing Then
        Set xmlEleTemp = vxmlResponseNode.appendChild(TempXML.createElement("VALUATIONFEE"))
    Else
        Set xmlEleTemp = vxmlResponseNode.selectSingleNode("VALUATIONFEE")
    End If
    
    If xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
        Call xmlEleTemp.appendChild(TempXML.createElement("AMOUNT"))
    End If
        
    If xmlGetNode(xmlEleTemp, "REFUNDAMOUNT") Is Nothing Then
        Call xmlEleTemp.appendChild(TempXML.createElement("REFUNDAMOUNT"))
    End If
    
    With objCommand
        .CommandText = "USP_GetValuationFeeAmount"
        .CommandType = adCmdStoredProc
        
        Set objParam = objCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(sApplicationNumber), sApplicationNumber)
        Call .Parameters.Append(objParam)
        
        Set objParam = objCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, 5, lAppFactFindNumber)
        Call .Parameters.Append(objParam)
    
        Set objParam = objCommand.CreateParameter("MORTGAGESUBQUOTENUMBER", adInteger, adParamInput, 5, lMSQNumber)
        Call .Parameters.Append(objParam)
    
    End With
    
    'Execute Stored Procedure
    Set rstData = executeGetRecordSetEx(objCommand)
    If Not rstData Is Nothing Then
        If Not rstData.EOF And Not rstData.BOF Then
            If Not xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
                dAmount = Val(rstData.Fields("Amount").Value & "")
            End If
        End If
    End If
    
    'Set the Amount
    xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount
    
    
    'Now get the Refund Amount
    With objCommand
        .CommandText = "USP_GetValuationRefundAmount"
        .CommandType = adCmdStoredProc
        
        Set objParam = objCommand.CreateParameter("REFUNDVALUATIONTYPE", adBSTR, adParamInput, Len(strRefundValuationType), strRefundValuationType)
        Call .Parameters.Append(objParam)
    
    End With
    
    'Execute Stored Procedure
    Set rstData = executeGetRecordSetEx(objCommand)
    If Not rstData Is Nothing Then
        If Not rstData.EOF And Not rstData.BOF Then
            If Not xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
                dRefundAmount = Val(rstData.Fields("Amount").Value & "")
            End If
        End If
    End If
    
    'Set the RefundAmount
    xmlGetNode(xmlEleTemp, "REFUNDAMOUNT").Text = dRefundAmount
    
    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            rstData.Close
        End If
    End If

GetValuationFee_Exit:
    
    Set rstData = Nothing
    Set objParam = Nothing
    Set objCommand = Nothing
      
   errCheckError strFunctionName, TypeName(Me)
    
End Sub

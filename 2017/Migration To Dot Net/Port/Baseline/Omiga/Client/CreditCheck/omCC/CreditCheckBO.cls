VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CreditCheckBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4BO' with actual class name
' *****************************************************************************

'Workfile:      Omiga4BO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions'
'               only
'               i.e. any code that does not require transaction support will reside in
'               this first level module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level)
'               module (Omiga4TxBO in this example)
'
'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MDC    13/03/2000  Created
'MDC    18/05/2000  Add RunCreditCheck and DemoOptions methods. Updated interface
'MDC    04/07/2000  Add CheckCreditCheckResponse method
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'SR     08/12/00    included methods for Full Bureau Download
'SR     13/12/00    included methods for Full Bureau Download
'DRC    04/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BMID History:
'
'Prog   Date        Description
'MDC    14/08/2002  BMIDS00336 CCWP1 BM062 - Credit Check and Bureau Download
'MDC    08/10/2002  BMIDS00561 Correct header block id's in deduplication routine
'MDC    19/11/2002  BMIDS00985 Matching customers to Voters Roll data in GetAndDeduplicateBureauRecords
'MDC    22/11/2002  BMIDS01055 Handle apostrophes in customer names in GetAndDeduplicateBureauRecords
'MDC    16/01/2003  BM0251 Duplicated Mortgage Account records
'MDC    17/01/2003  BM0258
'GD     03/03/2003  BM0373 ENHANCE THE DE-DUPLICATE EXPERIAN RECORD PROCESS
'GD     27/03/2003  BM0486 ARREARS HISTORY - REPOSSESSIONS
'GD     30/06/2003  BM0592 Check for 'FBJOINTACCOUNT' instaed of 'FBPUBLICINFOJOINTACCOUNTFLAG' in GetAndDeduplicateBureauRecords()
'INR    29/08/2003  BMIDS630 should be checking the first Initial of Forename and Surname for a match, if
'                   Full Forename and Surname don't match, regardless of the length of the Customer Forename.
'INR    09/09/2003  BMIDS630 Also check FBPUBLICINFOREGISTRATIONDATE when determining joint records
'GHun   28/11/2003  BMIDS658 Minor performance improvements and clean up
'KRW    03/02/2004  BMIDS530 CC035 Do not save unassigned experian records
'DRC    24/02/2004  BMIDS693 CC056  Credit Search in Alias records
'INR    16/02/2004  BMIDS682 Address Targeting Processing
'INR    09/06/2004  BMIDS744 Get ApplicationFactFind OptOutIndicator
'INR    17/06/2004  BMIDS744 Ignore FullBureauXML if FBBureauRefCategory is not 1
'                   in GetAndDeduplicateBureauRecords.
'INR    12/08/2004  BMIDS744 Added GetCurrentCreditCheckStatus
'HMA    28/07/2004  BMIDS748 Added RunUpgradeToFullAppCreditCheck
'JD     04/08/2004  BMIDS832 Added check for FBMATCHEDTO to match SP data.
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'HM     31/08/2005  MAR13 Added GetCurrentXMLCreditCheckResults
'SR     20/10/2005  MAR24 New method FindApplicationCreditCheckDetailList
'INR    23/11/2005  MAR645 Added CreateRuleOverride
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICreditCheckBO
Implements ObjectControl

Private Const cstrROOT_NODE_NAME = "CREDITCHECK"
Private Const cstrFULLBUREAUCAIS_NAME = "FULLBUREAUCAIS"
Private Const cstrFULLBUREAUVOTERSROLL_NAME = "FULLBUREAUVOTERSROLL"
Private Const cstrFULLBUREAUPUBLICINFO_NAME = "FULLBUREAUPUBLICINFO"
Private Const cstrFULLBUREAUCIFAS_NAME = "FULLBUREAUCIFAS"
Private Const cstrFULLBUREAUCAPS_NAME = "FULLBUREAUCAPS"
Private Const cstrFULLBUREAUASSOCIATIONS_NAME = "FULLBUREAUASSOCIATIONS"
Private Const cstrFULLBUREAUCORRECTIONLINES_NAME = "FULLBUREAUCORRECTIONLINES"
Private Const cstrFULLBUREAUDATAHEADER_NAME = "FULLBUREAUDATAHEADER"
Private Const cstrAPPLICATIONCREDITCHECK_NAME = "APPLICATIONCREDITCHECK"
Private Const cstrFULLLBUREAUSTANDARDHEADER_NAME = "FULLBUREAUSTANDARDHEADER"

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist



'BMIDS00336 MDC 14/08/2002
Public Function GetAndDeduplicateBureauRecords(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' procedure:    GetAndDeduplicateBureauRecords
' author:       Mark Coleman
' date:         14/08/2002
' description:
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------

On Error GoTo GetAndDeduplicateBureauRecordsVbErr

Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
    
Const cstrFunctionName As String = "GetAndDeduplicateBureauRecords"

    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False

    Set xmlRequestDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)

    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseNode = ICreditCheckBO_GetAndDeduplicateBureauRecords(xmlRequestDoc.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetAndDeduplicateBureauRecords = xmlResponseDoc.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetAndDeduplicateBureauRecordsExit:
        
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

GetAndDeduplicateBureauRecordsVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetAndDeduplicateBureauRecords = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetAndDeduplicateBureauRecordsExit

End Function
'BMIDS00336 MDC 14/08/2002 - End

Private Sub GetApplicationCreditCheckKeyData(ByVal vXMLRequest As IXMLDOMNode, _
                                             ByRef strCreditCheckRefNo As String, _
                                             ByRef strCreditCheckGuid As String)

On Error GoTo GetVbErr
                                    
    Const cstrFunctionName As String = "GetApplicationCreditCheckKeyData"
                                        
    Dim objICreditCheckDO As ICreditCheckDO
                                    
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTable As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strSequenceNumber As String
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTable = xmlDoc.createElement("APPLICATIONCREDITCHECK")
    xmlNewRequest.appendChild xmlTable
                                
    xmlTable.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTable.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    strSequenceNumber = m_objXmlAssist.GetTagValue(vXMLRequest, "SEQUENCENUMBER")
    If Len(strSequenceNumber) > 0 Then
        Set xmlNode = xmlDoc.createElement("SEQUENCENUMBER")
        xmlNode.Text = strSequenceNumber
        xmlTable.appendChild xmlNode
    End If
                                
    Set xmlResponse = objICreditCheckDO.FindApplicationCreditCheckList(xmlNewRequest)
    Set xmlNodeList = xmlResponse.selectNodes(".//APPLICATIONCREDITCHECK")
    Set xmlNode = xmlNodeList.Item(xmlNodeList.length - 1)
    
    strCreditCheckGuid = xmlNode.selectSingleNode(".//CREDITCHECKGUID").Text
    strCreditCheckRefNo = xmlNode.selectSingleNode(".//CREDITCHECKREFERENCENUMBER").Text
    
    Set objICreditCheckDO = Nothing
    Set xmlDoc = Nothing
    Set xmlTable = Nothing
    Set xmlNewRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    m_objContext.SetComplete
    
    Exit Sub
    
GetVbErr:
    If m_objErrAssist.IsWarning = True Then
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set objICreditCheckDO = Nothing
    Set xmlDoc = Nothing
    Set xmlTable = Nothing
    Set xmlNewRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    m_objContext.SetComplete
    
    m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 276
                                    
End Sub

Public Function GetCurrentCreditCheckResults(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCurrentCreditCheckResultsVbErr
    
    Const strFunctionName As String = "GetCurrentCreditCheckResults"

    
Dim objXmlIn As FreeThreadedDOMDocument40
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode
    
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentCreditCheckResults(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode

    GetCurrentCreditCheckResults = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentCreditCheckResultsExit:
        
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

GetCurrentCreditCheckResultsVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetCurrentCreditCheckResults = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetCurrentCreditCheckResultsExit
End Function

Public Function GetCurrentXMLCreditCheckResults(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCurrentXMLCreditCheckResultsVbErr
    
    Const strFunctionName As String = "GetCurrentXMLCreditCheckResults"

    
Dim objXmlIn As FreeThreadedDOMDocument40
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode
    
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentXMLCreditCheckResults(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode

    GetCurrentXMLCreditCheckResults = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentXMLCreditCheckResultsExit:
        
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

GetCurrentXMLCreditCheckResultsVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetCurrentXMLCreditCheckResults = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetCurrentXMLCreditCheckResultsExit
End Function

Public Function GetCurrentCAISSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCurrentCAISSummaryVbErr
    
    Const strFunctionName As String = "GetCurrentCAISSummary"

    
Dim objXmlIn As FreeThreadedDOMDocument40
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode
    
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentCAISSummary(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode

    GetCurrentCAISSummary = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentCAISSummaryExit:
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

GetCurrentCAISSummaryVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetCurrentCAISSummary = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetCurrentCAISSummaryExit

End Function

Public Function RunFullBureau(ByVal vstrXMLRequest As String) As String

On Error GoTo RunFullBureauVbErr
    
    Const strFunctionName As String = "RunFullBureau"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseNode = ICreditCheckBO_RunFullBureau(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    RunFullBureau = xmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RunFullBureauExit:
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

RunFullBureauVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    RunFullBureau = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume RunFullBureauExit

End Function

Private Function ValidateBeforeFBDownLoad(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

On Error GoTo ValidateVbErr
    
    Const cstrFunctionName As String = "ValidateBeforeFBDownLoad"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlClassDef As FreeThreadedDOMDocument40
    
    Dim objIClassDef As ICreditCheckClassDef
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssit As IADOAssist
    Dim objSQLAssist As SQLAssist
    
    Dim strSuccessIndicator As String
    Dim strCreditCheckGuid As String
    Dim strCondition As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssit = New ADOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssit = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    '----------------------------------------------------------------------------------------------------
    'Check whether the last inital creditcheck for this application was successful. If no, raise an error
    '----------------------------------------------------------------------------------------------------
    Set xmlRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATIONCREDITCHECK")
    xmlRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    'xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//SEQUENCENUMBER").cloneNode(True)
    
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDef = objIClassDef.LoadApplicationCreditCheckData()
    
    Set xmlTempResponseNode = objIDOAssist.FindList(xmlRequest, xmlClassDef, "SEQUENCENUMBER")
    Set xmlNodeList = xmlTempResponseNode.selectNodes("APPLICATIONCREDITCHECK")
    Set xmlNode = xmlNodeList.Item(xmlNodeList.length - 1)
    
    strSuccessIndicator = m_objXmlAssist.GetTagValue(xmlNode, "SUCCESSINDICATOR")
    strCreditCheckGuid = m_objXmlAssist.GetTagValue(xmlNode, "CREDITCHECKGUID")
    
    If strSuccessIndicator <> "1" Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4808
    End If
    
    '------------------------------------------------------------------------------
    'Check whether FBDownload was already done for this case. If so, raise an error
    '------------------------------------------------------------------------------
    Set objSQLAssist = New SQLAssist
    
    'BMIDS744 INR Changed to speed up searches on database.
    strCondition = "CREDITCHECKGUID=" & objSQLAssist.FormatGuid(strCreditCheckGuid)
'    strCondition = "CREDITCHECKGUID=" & objSQLAssist.FormatString(strCreditCheckGuid)
    
    If objIADOAssit.CheckRecordExists("FULLBUREAUSTANDARDHEADER", strCondition) Then
        ' Raise error
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4809
    End If
    
    'BMIDS00336 MDC 03/09/2002 - Return Sequence Number
    m_objXmlAssist.CopyNode "SEQUENCENUMBER", "SEQUENCENUMBER", xmlNode, xmlResponseElem, False
    'BMIDS00336 MDC 03/09/2002 - End
    
    Set ValidateBeforeFBDownLoad = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateVbExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlClassDef = Nothing

    Set objIClassDef = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssit = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function
    
ValidateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ValidateBeforeFBDownLoad = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume ValidateVbExit

End Function

Private Function ICreditCheckBO_CheckCreditCheckResponse(ByVal vxmlResponseToCheck As IXMLDOMElement, ByVal vxmlResponseToAddTo As IXMLDOMElement, Optional ByVal vblnRaiseError As Boolean = False) As Long
' header ----------------------------------------------------------------------------------
' description:
'   takes an xml response and returns the error number. Re raises the error if
'   vblnRaiseError is set to true
'   This function is Credit Check specific. It will filter out Credit Check
'   failure error messages and not re-raise them.
' pass:
'   vxmlResponseToCheck   xml Response to check for errors
'   vxmlResponseToAddTo   xml Element to add the warnings to
'   vblnRaiseError        Determines whether to raise the error
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "ICreditCheckBO_CheckCreditCheckResponse"

    
    Dim objXmlAssist As New XMLAssist
    Dim xmlErrNumElement As IXMLDOMElement
    Dim xmlErrSourceElement As IXMLDOMElement
    Dim xmlErrDescElement As IXMLDOMElement
    Dim xmlFirstChildElem As IXMLDOMNode
    Dim lngErrorNumber As Long
    
    Dim xmlMessageList As IXMLDOMNodeList
    Dim xmlMessageElem As IXMLDOMElement
    
    If vxmlResponseToCheck Is Nothing Or vxmlResponseToAddTo Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Response to Check or Response to add to missing"
    End If
    
    If vxmlResponseToCheck.nodeName <> "RESPONSE" Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "RESPONSE must be top level tag"
    End If
    
    ' If the error is to be re-raised get the appropriate data out of the response
    ' and raise the error
    If vxmlResponseToCheck.getAttribute("TYPE") = "WARNING" Then
        vxmlResponseToAddTo.setAttribute "TYPE", "WARNING"
        
        Set xmlFirstChildElem = vxmlResponseToCheck.firstChild
        Set xmlMessageList = vxmlResponseToCheck.getElementsByTagName("MESSAGE")

        ' insert messages at the top of the response
        For Each xmlMessageElem In xmlMessageList
            vxmlResponseToAddTo.insertBefore xmlMessageElem.cloneNode(True), _
                                             xmlFirstChildElem
        Next
    ElseIf vxmlResponseToCheck.getAttribute("TYPE") <> "SUCCESS" Then
        Set xmlErrNumElement = vxmlResponseToCheck.getElementsByTagName("NUMBER").Item(0)
        lngErrorNumber = CLng(xmlErrNumElement.Text)
        'Filter out specific credit check failures
        'Errors: 280    Insufficient data for credit check
        '        282    Credit Check interface currently unavailable
        '        548    Credit Check failed
        Select Case lngErrorNumber - vbObjectError - 512
        Case 280, 282, 548
            'Do not raise an error
            vblnRaiseError = False
        End Select
        If vblnRaiseError = True Then
            Set xmlErrSourceElement = vxmlResponseToCheck.getElementsByTagName("SOURCE").Item(0)
            Set xmlErrDescElement = vxmlResponseToCheck.getElementsByTagName("DESCRIPTION").Item(0)
            Err.Raise lngErrorNumber, xmlErrSourceElement.Text, xmlErrDescElement.Text
        End If
    End If

    ICreditCheckBO_CheckCreditCheckResponse = lngErrorNumber
    
    
ICreditCheckBO_CheckCreditCheckResponseExit:

    Set objXmlAssist = Nothing
    Set xmlErrNumElement = Nothing
    Set xmlErrSourceElement = Nothing
    Set xmlErrDescElement = Nothing
    Set xmlFirstChildElem = Nothing
    Exit Function
    
ICreditCheckBO_CheckCreditCheckResponseVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    ICreditCheckBO_CheckCreditCheckResponse = Err.Number
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_CheckCreditCheckResponseExit


End Function
'SR 18/10/2005 : MAR24
Private Function ICreditCheckBO_FindApplicationCreditCheckDetailsList( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    Const strFunctionName As String = "ICreditCheckBO_FindApplicationCreditCheckDetailsList"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    If vXMLRequest.nodeName = "APPLICATIONCREDITCHECKDETAILS" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("APPLICATIONCREDITCHECKDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONCREDITCHECKDETAILS" & " tag not found"
    End If

    Set xmlDataNode = objICreditCheckDO.FindApplicationCreditCheckDetailsList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindApplicationCreditCheckDetailsList = xmlResponseElem
FindListExit:
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objICreditCheckDO = Nothing
    
    Exit Function
FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindApplicationCreditCheckDetailsList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
End Function

Private Function ICreditCheckBO_FindApplicationCreditCheckList( _
                        ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindApplicationCreditCheckList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vXMLRequest.nodeName = cstrAPPLICATIONCREDITCHECK_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrAPPLICATIONCREDITCHECK_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrAPPLICATIONCREDITCHECK_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindApplicationCreditCheckList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindApplicationCreditCheckList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindApplicationCreditCheckList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
End Function

Private Function ICreditCheckBO_FindFullBureauAssociationsList( _
                    ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauAssociationsList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUASSOCIATIONS_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUASSOCIATIONS_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUASSOCIATIONS_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauAssociationsList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauAssociationsList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauAssociationsList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function ICreditCheckBO_FindFullBureauCAPSList( _
            ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauCAPSList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUCAPS_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUCAPS_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUCAPS_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauCAPSList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauCAPSList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauCAPSList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function ICreditCheckBO_FindFullBureauCIFASList( _
                ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauCIFASList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUCIFAS_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUCIFAS_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUCIFAS_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauCIFASList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauCIFASList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauCIFASList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Private Function ICreditCheckBO_FindFullBureauCorrectionLinesList( _
                ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauCorrectionLinesList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUCORRECTIONLINES_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUCORRECTIONLINES_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUCORRECTIONLINES_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauCorrectionLinesList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauCorrectionLinesList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function ICreditCheckBO_FindFullBureauPublicInfoList( _
                ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauPublicInfoList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUPUBLICINFO_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUPUBLICINFO_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUPUBLICINFO_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauPublicInfoList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauPublicInfoList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauPublicInfoList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function ICreditCheckBO_FindFullBureauVotersRollList( _
                    ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_FindFullBureauVotersRollList"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vobjRequestNode.nodeName = cstrFULLBUREAUVOTERSROLL_NAME Then
        Set xmlRequestNode = vobjRequestNode
    Else
        Set xmlRequestNode = vobjRequestNode.getElementsByTagName(cstrFULLBUREAUVOTERSROLL_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUVOTERSROLL_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.FindFullBureauVotersRollList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_FindFullBureauVotersRollList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_FindFullBureauVotersRollList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit


End Function

'BMIDS00336 MDC 14/08/2002
Private Function ICreditCheckBO_GetAndDeduplicateBureauRecords(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetAndDeduplicateBureauRecordsVbErr

Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlApplicationElement As IXMLDOMElement
Dim xmlCustomerListElement As IXMLDOMElement
Dim xmlGuidElement As IXMLDOMElement
Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
Dim xmlTempRequestElem As IXMLDOMElement
Dim xmlTempParentElem As IXMLDOMElement
Dim xmlTempElement As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim xmlAssociationListNode As IXMLDOMNode
Dim xmlCustomerResponseList As IXMLDOMElement
Dim xmlAliasResponseList As IXMLDOMElement
Dim xmlCAISResponseList As IXMLDOMElement
Dim xmlPublicInfoResponseList As IXMLDOMElement
Dim xmlFBHeaderRequestElem As IXMLDOMElement
Dim xmlFBHeaderElem As IXMLDOMElement
Dim xmlFBHeaderBlockIdElem As IXMLDOMElement
Dim xmlFBHeaderBlockSeqElem As IXMLDOMElement
Dim xmlCustomerDetailsElem As IXMLDOMElement
Dim xmlCAISElem As IXMLDOMElement
Dim xmlPublicInfoElem As IXMLDOMElement
Dim xmlBureauDataElem As IXMLDOMElement
Dim xmlVotersRollListNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode  'BM0251 MDC 15/01/2003

Dim strCreditCheckRefNo As String
Dim strCreditCheckGuid As String
Dim strCustomerForename As String
Dim strCustomerSurname As String
Dim strCustomerTitle As String   'BMIDS904
'GD BM0373 Added START

'Dim strCustomerVersionNumber As String
Dim strFirstCharForename As String
Dim strMonth As String
Dim strYear As String
Dim strFullDateOfBirth As String
Dim xmlTempDoc As FreeThreadedDOMDocument40

'GD BM0373 Added END
Dim strStartDate As String
Dim strBalance As String
Dim strPayment As String
Dim strAmountInPounds As String
Dim strAmountInPence As String
Dim strSatisfiedDate As String
'BMIDS630
Dim strRegistrationDate As String
Dim strPattern As String
'GD BM0486 START
'Node lists
Dim xmlFullBureauCAISList As IXMLDOMNodeList
Dim xmlFullBureauCAISListDeDupe As IXMLDOMNodeList
Dim xmlMatchList As IXMLDOMNodeList
'Elements
'Dim xmlCAISElem As IXMLDOMElement
Dim xmlAddressElem As IXMLDOMElement
Dim xmlCAISSettledDateElem As IXMLDOMElement
Dim xmlCAISResponseListDeDupe As IXMLDOMElement
'Dim xmlMatchList As IXMLDOMNodeList
Dim xmlTempElem As IXMLDOMElement
'Simple types
Dim intFullBureauCAISListLength As Integer
Dim intIndex As Integer
Dim strAddressType As String
Dim strCAISSettledDate As String
Dim intMatchListIndex As Integer
Dim strFBHeaderSequence As String
'GD BM0486 END
'DRC BMIDS693 START
Dim xmlCustomerNodeList As IXMLDOMNodeList
Dim xmlAliasNodeList As IXMLDOMNodeList
Dim xmlAliasNode As IXMLDOMNode
Dim xmlAliasElement As IXMLDOMElement
Dim strCustomerNo As String
Dim strCustVerNo As String
Dim strAliasSurname As String
Dim objCustomerBO As CustomerBO
Dim strTempResponse As String
Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
'DRC BMIDS693 END
'BMIDS744
Dim strFBBureauRefCategory As String
Dim strFBMatchedTo As String  ' JD BMIDS832
        
Const cstrFunctionName As String = "ICreditCheckBO_GetAndDeduplicateBureauRecords"

    '=================================================================================
    ' Initialise
    '=================================================================================
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    'GD BM0373 Added START
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    'GD BM0373 Added END
    'Initialise objects
    Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
    xmlTempRequestDoc.validateOnParse = False
    xmlTempRequestDoc.setProperty "NewParser", True
    Set xmlAliasResponseList = xmlResponseDoc.createElement("FULLBUREAUASSOCIATIONSLIST")
    Set xmlCAISResponseList = xmlResponseDoc.createElement("FULLBUREAUCAISLIST")
    Set xmlPublicInfoResponseList = xmlResponseDoc.createElement("FULLBUREAUPUBLICINFOLIST")
    Set xmlVotersRollListNode = xmlResponseDoc.createElement("FULLBUREAUVOTERSROLLLIST")
    Set xmlBureauDataElem = xmlResponseDoc.createElement("BUREAUDATA")
    xmlResponseElem.appendChild xmlBureauDataElem
    
    '=================================================================================
    ' Validate
    '=================================================================================
        
    'Validate request
    Set xmlApplicationElement = m_objXmlAssist.GetMandatoryNode(vXMLRequest, _
            "APPLICATION[(APPLICATIONNUMBER != '') and (APPLICATIONFACTFINDNUMBER != '')]")
    Set xmlCustomerListElement = m_objXmlAssist.GetMandatoryNode(vXMLRequest, _
            "CUSTOMERLIST[CUSTOMER[(CUSTOMERNUMBER != '') and (CUSTOMERVERSIONNUMBER != '')]]")
    

    '=================================================================================
    ' BMIDS693 DRC
    ' Retrieve Alias data already declared in Omiga
    '=================================================================================
    Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    ' Only want records with AliasType = "ALIAS"
    Set xmlTempElement = xmlTempRequestDoc.createElement("ALIASRECORDSONLY")
    xmlTempElement.Text = "1"
    xmlTempRequestElem.appendChild xmlTempElement
    ' Only want records declared in Omiga4
    Set xmlTempElement = xmlTempRequestDoc.createElement("DECLAREDRECORDSONLY")
    xmlTempElement.Text = "1"
    xmlTempRequestElem.appendChild xmlTempElement
    
    'Loop through each customer in the request
    Set xmlNodeList = xmlCustomerListElement.selectNodes(".//CUSTOMER")
    For Each xmlNode In xmlNodeList
        Set xmlTempNode = xmlNode.cloneNode(True)
        m_objXmlAssist.ChangeNodeName xmlTempNode, "CUSTOMER", "ALIAS"
        xmlTempRequestElem.appendChild xmlTempNode
        ' Convert the Surname and first forename to Uppercase only
        Set xmlTempElem = xmlNode.selectSingleNode(".//SURNAME")
        xmlTempElem.Text = ToUpperAlphaOnly(xmlTempElem.Text)
        Set xmlTempElem = xmlNode.selectSingleNode(".//FIRSTFORENAME")
        xmlTempElem.Text = ToUpperAlphaOnly(xmlTempElem.Text)
    Next
    strTempResponse = objCustomerBO.FindAliasList(xmlTempRequestElem.xml)
    Set xmlTempResponseDoc = m_objXmlAssist.load(strTempResponse, TypeName(Me), cstrFunctionName)
    Set xmlTempResponse = xmlTempResponseDoc.documentElement
    'Check response ignoring record not found error
    If CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName) Then
           Set xmlCustomerNodeList = xmlCustomerListElement.selectNodes(".//CUSTOMER")
        For Each xmlNode In xmlCustomerNodeList
          strCustomerNo = xmlNode.selectSingleNode(".//CUSTOMERNUMBER").Text
          strCustVerNo = xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
          'make a list of aliases found for this customer
           Set xmlAliasNodeList = xmlTempResponse.selectNodes( _
        ".//ALIAS[(translate(CUSTOMERNUMBER, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerNo) & _
           """) and (CUSTOMERVERSIONNUMBER = """ & strCustVerNo & """ )]")
           For Each xmlAliasNode In xmlAliasNodeList
              strAliasSurname = xmlAliasNode.selectSingleNode(".//SURNAME").Text
              Set xmlAliasElement = xmlTempRequestDoc.createElement("ALIASSURNAME")
              xmlAliasElement.Text = ToUpperAlphaOnly(strAliasSurname)
              xmlNode.appendChild xmlAliasElement
           Next
          
        Next
    
    End If
      
    '=================================================================================
    ' Retrieve Alias/Association Bureau data
    '=================================================================================
    'Get Credit Check Keys
         Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    xmlTempRequestElem.appendChild xmlApplicationElement
    Call GetApplicationCreditCheckKeyData(xmlTempRequestElem, strCreditCheckRefNo, strCreditCheckGuid)

    'Get Alias/Association data
    Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    Set xmlTempParentElem = xmlTempRequestDoc.createElement(cstrFULLBUREAUASSOCIATIONS_NAME)
    xmlTempRequestElem.appendChild xmlTempParentElem
    Set xmlGuidElement = xmlTempRequestDoc.createElement("CREDITCHECKGUID")
    xmlGuidElement.Text = strCreditCheckGuid
    xmlTempParentElem.appendChild xmlGuidElement
    
    Set xmlAssociationListNode = ICreditCheckBO_FindFullBureauAssociationsList(xmlTempRequestElem)
    'Check response ignoring record not found error
    Call CheckSuccessOrRecordNotFound(xmlAssociationListNode, cstrFunctionName)

    'Set up Request for call(s) to GetFullBureauDataHeaderData
    Set xmlFBHeaderRequestElem = vXMLRequest.cloneNode(False)
    Set xmlFBHeaderElem = xmlTempRequestDoc.createElement(cstrFULLBUREAUDATAHEADER_NAME)
    xmlFBHeaderRequestElem.appendChild xmlFBHeaderElem
    xmlFBHeaderElem.appendChild xmlGuidElement
    Set xmlFBHeaderBlockIdElem = xmlTempRequestDoc.createElement("FBBLOCKID")
    xmlFBHeaderElem.appendChild xmlFBHeaderBlockIdElem
    Set xmlFBHeaderBlockSeqElem = xmlTempRequestDoc.createElement("FBHEADERSEQUENCE")
    xmlFBHeaderElem.appendChild xmlFBHeaderBlockSeqElem
     'Loop through each record returned
    Set xmlNodeList = xmlAssociationListNode.selectNodes(".//FULLBUREAUASSOCIATIONS")
    For Each xmlNode In xmlNodeList
        'BMIDS744 Ignore if FBBureauRefCategory is not 1
        strFBBureauRefCategory = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBBUREAUREFCATEGORY", False, False)
        If (StrComp(strFBBureauRefCategory, "1", vbTextCompare) = 0) Then
            
            'Get related Bureau Header record
            xmlFBHeaderBlockIdElem.Text = "BL01"
            m_objXmlAssist.SetMandatoryChildText xmlFBHeaderElem, "FBHEADERSEQUENCE", m_objXmlAssist.GetMandatoryElementText(xmlNode, "FBHEADERSEQUENCE")
            'm_objXmlAssist.CopyNode "FBHEADERSEQUENCE", "FBHEADERSEQUENCE", xmlNode, xmlFBHeaderElem, False
            Set xmlTempResponse = ICreditCheckBO_GetFullBureauDataHeaderData(xmlFBHeaderRequestElem)
    '        Set xmlTempResponse = ICreditCheckBO_GetFullBureauDataHeaderData(xmlTempRequestElem)
            
            'Check response ignoring record not found error
            If CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName) Then
                'Try to find matching customer
                strCustomerForename = ToUpperAlphaOnly(m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBFORENAME", False, False))
                strCustomerSurname = ToUpperAlphaOnly(m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBSURNAME", False, False))
                
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 START 1 (ALIAS BL01)~~~~~~~~~~~~~~~~~~~~~
                'BMIDS693
                'strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """]"
                strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                                     """)) and translate(FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerForename) & """]"
                                     
                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                'blnMatchFound = False
                If (xmlCustomerDetailsElem Is Nothing) Then
                    'Check first initial...as below
                 'BMIDS630 INR should be checking the first Initial if we get to here,
                 'regardless of the length of strCustomerForename
    '                If Len(strCustomerForename) = 1 Then ' 1 letter returned from EXPERIAN for Forename
                    'Check if first character of forenames match
                        strFirstCharForename = Left$(strCustomerForename, 1)
                        Set xmlTempDoc.documentElement = vXMLRequest ' vxmlRequest is the origin on xmlCustomerListElement
                        xmlTempDoc.setProperty "SelectionLanguage", "XPath"
                        strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                        strFirstCharForename = ToUpperAlphaOnly(strFirstCharForename)
                        'BMIDS693
                        'strPattern = "CUSTOMER[SURNAME[translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & strCustomerSurname & """] and FIRSTFORENAME[translate(substring(.,1,1),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" & strFirstCharForename & "']]"
                        strPattern = "CUSTOMER[((SURNAME = """ & strCustomerSurname & """) or ( ALIASSURNAME = """ & strCustomerSurname & _
                        """))  and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']    ]"
     
                        Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                        xmlTempDoc.setProperty "SelectionLanguage", "XSLPattern" 'reset
    '                End If
                End If
                'If found, append Customer details to the Alias/Association Node
                If Not (xmlCustomerDetailsElem Is Nothing) Then
                    xmlNode.appendChild xmlCustomerDetailsElem.cloneNode(True)
                End If
                'GD END OF INSERTED CODE
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 END 1 (ALIAS BL01)'~~~~~~~~~~~~~~~~~~~~~
    
    
                '... and append the Alias/Association node to the list
                xmlAliasResponseList.appendChild xmlNode
            End If
        End If 'End FBBureauRefCategory check
    Next
    
    
    'Append the Alias/Association list to the main response
    If xmlAliasResponseList.hasChildNodes Then
        xmlBureauDataElem.appendChild xmlAliasResponseList
    End If
    
    
    '=================================================================================
    ' Retrieve and deduplicate CAIS Bureau data
    '=================================================================================

    'Get CAIS data
    Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    Set xmlTempParentElem = xmlTempRequestDoc.createElement(cstrFULLBUREAUCAIS_NAME)
    xmlTempRequestElem.appendChild xmlTempParentElem
    Set xmlGuidElement = xmlTempRequestDoc.createElement("CREDITCHECKGUID")
    xmlGuidElement.Text = strCreditCheckGuid
    xmlTempParentElem.appendChild xmlGuidElement

    Set xmlTempResponse = ICreditCheckBO_GetCurrentFullBureauCAISSummary(xmlTempRequestElem)
    
    'Check response ignoring record not found error
    Call CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName)

    xmlFBHeaderBlockIdElem.Text = "BC01"

    Set xmlNodeList = xmlTempResponse.selectNodes(".//FULLBUREAUCAIS")
    For Each xmlNode In xmlNodeList
        'BMIDS744 Ignore if FBBureauRefCategory is not 1
        strFBBureauRefCategory = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBBUREAUREFCATEGORY", False, False)
        If (StrComp(strFBBureauRefCategory, "1", vbTextCompare) = 0) Then
        
            'Reset customer list
            Set xmlCustomerResponseList = xmlResponseDoc.createElement("CUSTOMERLIST")
            xmlNode.appendChild xmlCustomerResponseList
            
            'Get related Bureau Header record
            'Set xmlTempElement = m_objXmlAssist.GetNode(xmlFBHeaderElem, "FBHEADERSEQUENCE", True)
            xmlFBHeaderBlockSeqElem.Text = m_objXmlAssist.GetNodeValue(xmlNode, "FBHEADERSEQUENCE", True, True)
            Set xmlTempResponse = ICreditCheckBO_GetFullBureauDataHeaderData(xmlFBHeaderRequestElem)
            Debug.Print "CAIS Ref: " & xmlFBHeaderBlockSeqElem.Text
            
            'Check response ignoring record not found error
            If CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName) Then
            
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 START 2.1 (CAIS BC01)~~~~~~~~~~~~~~~~~~~~~
                
                'Try to find matching customer
                strCustomerForename = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBFORENAME", False, False)
                strCustomerForename = ToUpperAlphaOnly(strCustomerForename)
                strCustomerSurname = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBSURNAME", False, False)
                strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                'GD START OF CODE REMOVAL
    '''''            'BMIDS01055 MDC 21/11/2002 - Handle apostrophes in names
    '''''            strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """]"
    '''''            'BMIDS01055 MDC 21/11/2002 - End
    '''''            Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                'GD END OF CODE REMOVAL
                'GD START OF CODE INSERTED
                'strCustomerVersionNumber = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//CUSTOMERVERSIONNUMBER", False, False)
                strFullDateOfBirth = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBDATEOFBIRTH", False, False)
                'strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and DATEOFBIRTH $ieq$ """ & strFullDateOfBirth & """ and CUSTOMERVERSIONNUMBER $ieq$ """ & strCustomerVersionNumber & """]"
                'DRC BMIDS693
                strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                                     """))and translate(DATEOFBIRTH, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strFullDateOfBirth) & """ ]"
                
                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                'Check FULL DATE : Format on  : 27/11/1968
                If (xmlCustomerDetailsElem Is Nothing) Then
                 ' Check PARTIAL DATE ... as below
                    strYear = Format$(strFullDateOfBirth, "yyyy")
                    strMonth = Format$(strFullDateOfBirth, "mm")
                    
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                
                    Set xmlTempDoc.documentElement = vXMLRequest ' vxmlRequest is the origin on xmlCustomerListElement
                    xmlTempDoc.setProperty "SelectionLanguage", "XPath"
                    strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                    'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and DATEOFBIRTH[substring(.,4,2) = '" & strMonth & "']  and DATEOFBIRTH[substring(.,string-length(.) - 3,4) = '" & strYear & "']   and CUSTOMERVERSIONNUMBER = """ & strCustomerVersionNumber & """]"
                    'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and DATEOFBIRTH[substring(.,4,2) = '" & strMonth & "']  and DATEOFBIRTH[substring(.,string-length(.) - 3,4) = '" & strYear & "']]"
       '             strPattern = "CUSTOMER[SURNAME[translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & strCustomerSurname & """] and DATEOFBIRTH[substring(.,4,2) = '" & strMonth & "']  and DATEOFBIRTH[substring(.,string-length(.) - 3,4) = '" & strYear & "']]"
                     'DRC BMIDS 693
                     strPattern = "CUSTOMER[((SURNAME = """ & strCustomerSurname & """) or ( ALIASSURNAME = """ & strCustomerSurname & _
                                """)) and DATEOFBIRTH[substring(.,4,2) = '" & strMonth & "'] and DATEOFBIRTH[substring(.,string-length(.) - 3,4) = '" & strYear & "']]"
                    Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                    xmlTempDoc.setProperty "SelectionLanguage", "XSLPattern" 'reset
                    If xmlCustomerDetailsElem Is Nothing Then
                        'Try and match on SURNAME AND FORENAME AND CUSTOMERVERSION
                        'strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """  and CUSTOMERVERSIONNUMBER $ieq$ """ & strCustomerVersionNumber & """]"
                        'DRC BMIDS693
                        strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                                      """))  and translate(FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerForename) & """]"
                        
                        Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                        If xmlCustomerDetailsElem Is Nothing Then
                            'JD BMIDS832 try to match using FBMatchedTo field
                            strFBMatchedTo = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBMATCHEDTO", False, False)
                            If strFBMatchedTo = "1" Or strFBMatchedTo = "2" Or _
                               strFBMatchedTo = "3" Or strFBMatchedTo = "9" Then
                                strPattern = "CUSTOMER[CUSTOMERORDER = '1']"
                                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                            Else
                                strPattern = "CUSTOMER[CUSTOMERORDER = '2']"
                                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                            End If
                        End If
                    End If
    '                If Not (xmlCustomerDetailsElem Is Nothing) Then
    '                    'Match Found
    '                    'ASSIGN EXPERIAN RECORD TO THIS APPLICANT
    '                Else
    '                    'No Match Found
    '                    'SET TO UNASSIGNED
    '                End If
                End If
                
                
                
                'GD END OF CODE INSERTED
                
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 END 2.1 (CAIS BC01)~~~~~~~~~~~~~~~~~~~~~
                
            End If
            
            'Check if Joint Record
            'BM0251 MDC 15/01/2003
            'If Len(Trim$(m_objXmlAssist.GetNodeValue(xmlNode, "FBCAISJOINTACCOUNTFLAG", False, False))) = 0 Then
            If Len(Trim$(m_objXmlAssist.GetNodeValue(xmlNode, "FBJOINTACCOUNTFLAG", False, False))) = 0 Then
            'BM0251 MDC 15/01/2003 - End
                'Not a joint account
                
                    'If found, append Customer details to the CAIS record
                    If Not xmlCustomerDetailsElem Is Nothing Then
                        xmlCustomerResponseList.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        
                        'BMIDS530 KRW 03/02/2004 Only append CAIS node if assigned to a customer
                        '... and append the CAIS node to the list
                        xmlCAISResponseList.appendChild xmlNode
                        'BMIDS530 End
                    End If
                    
                    'BMIDS530 KRW 03/02/2004
                    ''... and append the CAIS node to the list
                    'xmlCAISResponseList.appendChild xmlNode
                    'BMIDS530 End
            Else
                'Joint Account
        
                'Do we already have a record that is a duplicate of this?
                strStartDate = m_objXmlAssist.GetNodeValue(xmlNode, "FBCAISSTARTDATE", False, False)
                strBalance = m_objXmlAssist.GetNodeValue(xmlNode, "FBCAISBALANCE", False, False)
                strPayment = m_objXmlAssist.GetNodeValue(xmlNode, "FBCAISMONTHLYPAYMENT", False, False)
                strPattern = "FULLBUREAUCAIS[FBCAISSTARTDATE='" & strStartDate & "' and FBCAISBALANCE='" _
                                        & strBalance & "' and FBCAISMONTHLYPAYMENT='" & strPayment & "']"
                Set xmlCAISElem = m_objXmlAssist.GetNode(xmlCAISResponseList, strPattern, False)
                If Not xmlCAISElem Is Nothing And Not xmlCustomerDetailsElem Is Nothing Then
                    'Duplicate found, so append Customer Details to it
                    Set xmlTempElement = m_objXmlAssist.GetNode(xmlCAISElem, "CUSTOMERLIST", False)
                    If xmlTempElement Is Nothing Then
                        Set xmlTempElement = xmlTempRequestDoc.createElement("CUSTOMERLIST")
                        xmlTempElement.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        xmlCAISElem.appendChild xmlTempElement
                    Else
    
                        'BM0251 MDC 15/01/2003 - Check that customer not already added
                        'xmlTempElement.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        
                        strCustomerForename = m_objXmlAssist.GetNodeValue(xmlCustomerDetailsElem, "FIRSTFORENAME", False, False)
                        strCustomerSurname = m_objXmlAssist.GetNodeValue(xmlCustomerDetailsElem, "SURNAME", False, False)
                        
                        '~~~~~~~~~~~~~~~~~~~~~GD BM0373 START 2.2 (CAIS BC01)~~~~~~~~~~~~~~~~~~~~~
                        'strCustomerVersionNumber = m_objXmlAssist.GetNodeValue(xmlCustomerDetailsElem, "CUSTOMERVERSIONNUMBER", False, False)
                        
                        'GD CHANGEDstrPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """]"
                        'GD ADDED - Need to Look at CUSTOMERVERSIONNUMBER too.
                        'strPattern = "CUSTOMER[SURNAME $ieq$ """ & . & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """  and CUSTOMERVERSIONNUMBER $ieq$ """ & strCustomerVersionNumber & """]"
                        'DRC BMIDS693
                        strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                                      """)) and translate(FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerForename) & """ ]"
                        
                        
                        Set xmlTempNode = m_objXmlAssist.GetNode(xmlTempElement, strPattern, False)
                        If xmlTempNode Is Nothing Then
                            xmlTempElement.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        End If
                        'BM0251 MDC 15/01/2003 - End
                        '~~~~~~~~~~~~~~~~~~~~~GD BM0373 END 2.2 (CAIS BC01)~~~~~~~~~~~~~~~~~~~~~
                    End If
                Else
                    'No record already so add new one
                    If Not xmlCustomerDetailsElem Is Nothing Then
                        xmlCustomerResponseList.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        
                        'BMIDS530 KRW 03/02/2004 'BMIDS530 KRW 03/02/2004 Only append CAIS node if assigned to a customer
                        '... and append the CAIS node to the list
                        xmlCAISResponseList.appendChild xmlNode
                        'BMIDS530 KRW 03/02/2004 End
                    End If
                    
                    'BMIDS530 KRW 03/02/2004
                    ''... and append the CAIS node to the list
                    'xmlCAISResponseList.appendChild xmlNode
                    'BMIDS530 KRW 03/02/2004 End
                End If
            End If
        End If  'FBBureauRefCategory check
    Next
'GD BM0486 START
    'Now remove duplicates based on FBCAISSETTLEDDATE and FBCAISCMLADDRESSTYPE
    'Create a clone of xmlCAISResponseList called xmlCAISResponseListDeDupe.
    'Keep the original one intact, for looping through.
    Set xmlCAISResponseListDeDupe = xmlCAISResponseList.cloneNode(True)
    Set xmlFullBureauCAISList = xmlCAISResponseList.selectNodes(".//FULLBUREAUCAIS")
    'This will be the list where we remove duplicates from
    Set xmlFullBureauCAISListDeDupe = xmlCAISResponseListDeDupe.selectNodes(".//FULLBUREAUCAIS")
    
    'Set xmlFullBureauCAISDedupeList = xmlFullBureauCAISList.cloneNode(True)
    intFullBureauCAISListLength = xmlFullBureauCAISList.length
    For intIndex = 0 To (intFullBureauCAISListLength - 1)
        Set xmlCAISElem = xmlFullBureauCAISList.Item(intIndex)
        Set xmlAddressElem = xmlCAISElem.selectSingleNode("FBCAISCMLADDRESSTYPE")
        strAddressType = xmlAddressElem.Text
        If Trim$(strAddressType) = "R" Then
            'Reposession
            'Check for duplicate address type account with the same settled date
            Set xmlCAISSettledDateElem = xmlCAISElem.selectSingleNode("FBCAISSETTLEDDATE")
            strCAISSettledDate = xmlCAISSettledDateElem.Text
            If IsDate(strCAISSettledDate) Then
                Set xmlMatchList = xmlCAISResponseListDeDupe.selectNodes("FULLBUREAUCAIS[((FBCAISCMLADDRESSTYPE = 'F') or (FBCAISCMLADDRESSTYPE = 'A')) and FBCAISSETTLEDDATE = '" & strCAISSettledDate & "']")
                If xmlMatchList.length > 0 Then
                    'Remove the nodes in xmlMatchList from xmlFullBureauCAISListDeDupe
                    For intMatchListIndex = 0 To (xmlMatchList.length - 1)
                        strFBHeaderSequence = xmlMatchList.Item(intMatchListIndex).selectSingleNode("FBHEADERSEQUENCE").Text
                        Set xmlTempElem = xmlCAISResponseListDeDupe.selectSingleNode("FULLBUREAUCAIS[FBHEADERSEQUENCE = '" & strFBHeaderSequence & "']")
                        xmlCAISResponseListDeDupe.removeChild xmlTempElem
                    Next
                End If
            End If
        End If
    Next intIndex


    'Append the CAIS list to the main response.
    If xmlCAISResponseListDeDupe.hasChildNodes Then
        xmlBureauDataElem.appendChild xmlCAISResponseListDeDupe
    End If

'GD START OF CODE REMOVAL
'''''    'Append the CAIS list to the main response
'''''    If xmlCAISResponseList.hasChildNodes Then
'''''        xmlBureauDataElem.appendChild xmlCAISResponseList
'''''    End If
'GD END OF CODE REMOVAL
    
'GD BM0486 END


    '=================================================================================
    ' Retrieve and deduplicate Public Information Bureau data
    '=================================================================================

    'Get Public Info data (sorted by date descending)
    Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    Set xmlTempParentElem = xmlTempRequestDoc.createElement(cstrFULLBUREAUPUBLICINFO_NAME)
    xmlTempParentElem.setAttribute "SORT", "FBPUBLICINFOREGISTRATIONDATE DESC"
    xmlTempRequestElem.appendChild xmlTempParentElem
    Set xmlGuidElement = xmlTempRequestDoc.createElement("CREDITCHECKGUID")
    xmlGuidElement.Text = strCreditCheckGuid
    xmlTempParentElem.appendChild xmlGuidElement

    Set xmlTempResponse = ICreditCheckBO_GetCurrentPublicInfoSummary(xmlTempRequestElem)
    
    'Check response ignoring record not found error
    Call CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName)

    'BMIDS00561 MDC 07/10/2002 - Correct Block Id
    xmlFBHeaderBlockIdElem.Text = "BJ01"  '"BE01"
    'BMIDS00561 MDC 07/10/2002 - End

    Set xmlNodeList = xmlTempResponse.selectNodes(".//FULLBUREAUPUBLICINFO")
    For Each xmlNode In xmlNodeList
        'BMIDS744 Ignore if FBBureauRefCategory is not 1
        strFBBureauRefCategory = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBBUREAUREFCATEGORY", False, False)
        If (StrComp(strFBBureauRefCategory, "1", vbTextCompare) = 0) Then
        
            'Reset customer list
            Set xmlCustomerResponseList = xmlResponseDoc.createElement("CUSTOMERLIST")
            xmlNode.appendChild xmlCustomerResponseList
            
            'Get related Bureau Header record
            xmlFBHeaderBlockSeqElem.Text = m_objXmlAssist.GetNodeValue(xmlNode, "FBHEADERSEQUENCE", True, True)
            Set xmlTempResponse = ICreditCheckBO_GetFullBureauDataHeaderData(xmlFBHeaderRequestElem)
            
            'Check response ignoring record not found error
            If CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName) Then
            
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 START 3 (PUBLIC INFORMATION BJ01)~~~~~~~~~~~~~~~~~~~~~
    
                'Try to find matching customer
                strCustomerForename = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBFORENAME", False, False)
                strCustomerForename = ToUpperAlphaOnly(strCustomerForename)
                strCustomerSurname = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBSURNAME", False, False)
                strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                strCustomerTitle = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBTITLE", False, False)  'BMIDS904
                strCustomerTitle = ToUpperAlphaOnly(strCustomerTitle)                                        'BMIDS904
                
                'strCustomerVersionNumber = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//CUSTOMERVERSIONNUMBER", False, False)
                'GD START OF CODE REMOVAL
    '''''            strCustomerForename = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBFORENAME", False, False)
    '''''            strCustomerSurname = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBSURNAME", False, False)
    '''''            'BMIDS01055 MDC 21/11/2002 - Handle apostrophes in names
    '''''            strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ _
    '''''                                                                        & strCustomerForename & """]"
    '''''            'BMIDS01055 MDC 21/11/2002 - End
    '''''           Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
    
                'GD END OF CODE REMOVAL
                'GD START OF INSERTED CODE
                'strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """  and CUSTOMERVERSIONNUMBER $ieq$ """ & strCustomerVersionNumber & """]"
                'DRC BMIDS693
                strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                              """)) and translate(FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerForename) & """ ]"
                
                
                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                'blnMatchFound = False
                If (xmlCustomerDetailsElem Is Nothing) Then
                     ' Check first initial...as below
                 'BMIDS630 INR should be checking the first Initial if we get to here,
                 'regardless of the length of strCustomerForename
    '                If Len(strCustomerForename) = 1 Then ' 1 letter returned from EXPERIAN for Forename
                        'Check if first character of forenames match
                        strFirstCharForename = Left$(strCustomerForename, 1)
                        Set xmlTempDoc.documentElement = vXMLRequest ' vxmlRequest is the origin on xmlCustomerListElement
                        xmlTempDoc.setProperty "SelectionLanguage", "XPath"
                        
                        'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']    and CUSTOMERVERSIONNUMBER = """ & strCustomerVersionNumber & """]"
                        'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']    ]"
                        strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                        strFirstCharForename = ToUpperAlphaOnly(strFirstCharForename)
                        'DRC BMIDS693
                        'BMIDS904  Add Title to match criteria.
                        strPattern = "CUSTOMER[((SURNAME = """ & strCustomerSurname & """) or ( ALIASSURNAME = """ & strCustomerSurname & _
                        """))  and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "'] and TITLE = """ & _
                        strCustomerTitle & """]"
    
                        Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                        xmlTempDoc.setProperty "SelectionLanguage", "XSLPattern" 'reset
                        If (xmlCustomerDetailsElem Is Nothing) Then
                            'JD BMIDS832 try to match using FBMatchedTo field
                            strFBMatchedTo = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBMATCHEDTO", False, False)
                            If strFBMatchedTo = "1" Or strFBMatchedTo = "2" Or _
                               strFBMatchedTo = "3" Or strFBMatchedTo = "9" Then
                                strPattern = "CUSTOMER[CUSTOMERORDER = '1']"
                                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                            Else
                                strPattern = "CUSTOMER[CUSTOMERORDER = '2']"
                                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                            End If
                        End If
    '                End If
                End If
                
                
                'GD END OF INSERTED CODE
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 END 3 (PUBLIC INFORMATION BJ01)~~~~~~~~~~~~~~~~~~~~~
                
            End If
            
            'Check if Joint Record
            'GD BM0592 START
            'If Len(Trim$(m_objXmlAssist.GetNodeValue(xmlNode, "FBPUBLICINFOJOINTACCOUNTFLAG", False, False))) = 0 Then
            If Len(Trim$(m_objXmlAssist.GetNodeValue(xmlNode, "FBJOINTACCOUNT", False, False))) = 0 Then
            'GD BM0592 END
                'Not a joint account
                
                    'If found, append Customer details to the Public Info record
                    If Not xmlCustomerDetailsElem Is Nothing Then
                        xmlCustomerResponseList.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        
                        'BMIDS530 KRW 03/02/2004 Only append the public info record if assigned to a customer
                        '... and append the Public Info node to the list
                        xmlPublicInfoResponseList.appendChild xmlNode
                        'BMIDS530 KRW 03/02/2004
                    End If
                    
                    'BMIDS530 KRW 03/02/2004
                    '... and append the Public Info node to the list
                    'xmlPublicInfoResponseList.appendChild xmlNode
                    'BMIDS530 KRW 03/02/2004 End
            Else
                'Joint Record
        
                'Do we already have a record that is a duplicate of this?
                strAmountInPounds = m_objXmlAssist.GetNodeValue(xmlNode, "FBPUBLICINFOAMOUNTINPOUNDS", False, False)
                strAmountInPence = m_objXmlAssist.GetNodeValue(xmlNode, "FBPUBLICINFOAMOUNTINPENCE", False, False)
                strSatisfiedDate = m_objXmlAssist.GetNodeValue(xmlNode, "FBPUBLICINFOSATISFIEDDATE", False, False)
                'BMIDS630 Check RegistrationDate to help determine joint records
                strRegistrationDate = m_objXmlAssist.GetNodeValue(xmlNode, "FBPUBLICINFOREGISTRATIONDATE", False, False)
                strPattern = "FULLBUREAUPUBLICINFO[FBPUBLICINFOAMOUNTINPOUNDS='" & strAmountInPounds & "' and FBPUBLICINFOAMOUNTINPENCE='" _
                                        & strAmountInPence & "' and FBPUBLICINFOSATISFIEDDATE='" & strSatisfiedDate _
                                        & "' and FBPUBLICINFOREGISTRATIONDATE='" & strRegistrationDate & "']"
                
                Set xmlPublicInfoElem = m_objXmlAssist.GetNode(xmlPublicInfoResponseList, strPattern, False)
                If Not xmlPublicInfoElem Is Nothing And Not xmlCustomerDetailsElem Is Nothing Then
                    'Duplicate found, so append Customer Details to it
                    Set xmlTempElement = m_objXmlAssist.GetNode(xmlPublicInfoElem, "CUSTOMERLIST", False)
                    If xmlTempElement Is Nothing Then
                        Set xmlTempElement = xmlTempRequestDoc.createElement("CUSTOMERLIST")
                        xmlTempElement.appendChild xmlCustomerDetailsElem.cloneNode(True)
                        xmlPublicInfoElem.appendChild xmlTempElement
                    Else
                        xmlTempElement.appendChild xmlCustomerDetailsElem.cloneNode(True)
                    End If
                Else
                    'No record already so add new one
                    If Not xmlCustomerDetailsElem Is Nothing Then
                        xmlCustomerResponseList.appendChild xmlCustomerDetailsElem.cloneNode(True)
                    
                        'BMIDS530 KRW 03/02/2004 Only append the public info record if assigned to a customer
                        '... and append the CAIS node to the list
                        xmlPublicInfoResponseList.appendChild xmlNode
                        'BMIDS530 KRW 03/02/2004 End
                    End If
                    
                    'BMIDS530 KRW 03/02/2004
                    ''... and append the CAIS node to the list
                    'xmlPublicInfoResponseList.appendChild xmlNode
                    'BMIDS530 KRW 03/02/2004 End
                End If
            End If
        End If  'FBBureauRefCategory check
    Next

    'Append the Public Info list to the main response
    If xmlPublicInfoResponseList.hasChildNodes Then
        xmlBureauDataElem.appendChild xmlPublicInfoResponseList
    End If


    '=================================================================================
    ' Retrieve Voters Roll Bureau data
    '=================================================================================
    
    'Get Voters Roll data
    Set xmlTempRequestElem = vXMLRequest.cloneNode(False)
    Set xmlTempParentElem = xmlTempRequestDoc.createElement(cstrFULLBUREAUVOTERSROLL_NAME)
    xmlTempRequestElem.appendChild xmlTempParentElem
    Set xmlGuidElement = xmlTempRequestDoc.createElement("CREDITCHECKGUID")
    xmlGuidElement.Text = strCreditCheckGuid
    xmlTempParentElem.appendChild xmlGuidElement
    
    Set xmlTempResponse = ICreditCheckBO_FindFullBureauVotersRollList(xmlTempRequestElem)
    'Check response ignoring record not found error
    Call CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName)

    'BMIDS00561 MDC 07/10/2002
    xmlFBHeaderBlockIdElem.Text = "BE02"    '"BC01"
    'BMIDS00561 MDC 07/10/2002 - End

    Set xmlNodeList = xmlTempResponse.selectNodes(".//FULLBUREAUVOTERSROLL")
    For Each xmlNode In xmlNodeList
        'BMIDS744 Ignore if FBBureauRefCategory is not 1
        strFBBureauRefCategory = m_objXmlAssist.GetNodeValue(xmlNode, ".//FBBUREAUREFCATEGORY", False, False)
        If (StrComp(strFBBureauRefCategory, "1", vbTextCompare) = 0) Then
        
            'Get related Bureau Header record
            'BMIDS00561 MDC 07/10/2002
            'xmlFBHeaderBlockIdElem.Text = "BL01"
            'BMIDS00561 MDC 07/10/2002 - End
    
            'BMIDS00985 MDC 19/11/2002
            'Reset customer list
            Set xmlCustomerResponseList = xmlResponseDoc.createElement("CUSTOMERLIST")
            xmlNode.appendChild xmlCustomerResponseList
            'BMIDS00985 MDC 19/11/2002 - End
    
            'Get related Bureau Header record
            xmlFBHeaderBlockSeqElem.Text = m_objXmlAssist.GetNodeValue(xmlNode, "FBHEADERSEQUENCE", True, True)
            Set xmlTempResponse = ICreditCheckBO_GetFullBureauDataHeaderData(xmlFBHeaderRequestElem)
            
            'Check response ignoring record not found error
            If CheckSuccessOrRecordNotFound(xmlTempResponse, cstrFunctionName) Then
                
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 START 4 (VOTORS ROLE BE02)~~~~~~~~~~~~~~~~~~~~~
                'Try to find matching customer
                strCustomerForename = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBFORENAME", False, False)
                strCustomerForename = ToUpperAlphaOnly(strCustomerForename)
                strCustomerSurname = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//FBSURNAME", False, False)
                strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                'strCustomerVersionNumber = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//CUSTOMERVERSIONNUMBER", False, False)
                'GD START OF CODE REMOVAL
    '''''            'BMIDS01055 MDC 21/11/2002 - Handle apostrophes in names
    '''''            strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ _
    '''''                                                                        & strCustomerForename & """]"
    '''''            'BMIDS01055 MDC 21/11/2002 - End
    '''''            Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                'GD END OF CODE REMOVAL
                'GD START OF CODE INSERTED
                'strPattern = "CUSTOMER[SURNAME $ieq$ """ & strCustomerSurname & """ and FIRSTFORENAME $ieq$ """ & strCustomerForename & """  and CUSTOMERVERSIONNUMBER $ieq$ """ & strCustomerVersionNumber & """]"
                'DRC BMIDS693
                strPattern = "CUSTOMER[((translate(SURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & """) or (translate(ALIASSURNAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerSurname) & _
                               """)) and translate(FIRSTFORENAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = """ & UCase$(strCustomerForename) & """ ]"
                
                
                Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                
                If (xmlCustomerDetailsElem Is Nothing) Then
                 ' Check first initial...as below
                 'BMIDS630 INR should be checking the first Initial if we get to here,
                 'regardless of the length of strCustomerForename
    '                If Len(strCustomerForename) = 1 Then ' 1 letter returned from EXPERIAN for Forename
                        'Check if first character of forenames match
                        strFirstCharForename = Left$(strCustomerForename, 1)
                        Set xmlTempDoc.documentElement = vXMLRequest ' vxmlRequest is the origin on xmlCustomerListElement
                        xmlTempDoc.setProperty "SelectionLanguage", "XPath"
                        strCustomerSurname = ToUpperAlphaOnly(strCustomerSurname)
                        strFirstCharForename = ToUpperAlphaOnly(strFirstCharForename)
                        'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']    and CUSTOMERVERSIONNUMBER = """ & strCustomerVersionNumber & """]"
                        'strPattern = "CUSTOMER[SURNAME = """ & strCustomerSurname & """ and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']]"
                        'DRC BMIDS693
                        strPattern = "CUSTOMER[((SURNAME = """ & strCustomerSurname & """) or( ALIASSURNAME = """ & strCustomerSurname & _
                                      """)) and FIRSTFORENAME[substring(.,1,1) = '" & strFirstCharForename & "']    ]"
                        
                        Set xmlCustomerDetailsElem = m_objXmlAssist.GetNode(xmlCustomerListElement, strPattern, False)
                        xmlTempDoc.setProperty "SelectionLanguage", "XSLPattern" 'reset
    '                End If
                End If
                'GD END OF CODE INSERTED
                'If found, append Customer details to the Alias/Association Node
                If Not xmlCustomerDetailsElem Is Nothing Then
                    xmlCustomerResponseList.appendChild xmlCustomerDetailsElem.cloneNode(True)  'BMIDS00985 MDC 19/11/2002
                End If
                
                '~~~~~~~~~~~~~~~~~~~~~GD BM0373 END 4 (VOTORS ROLE BE02)~~~~~~~~~~~~~~~~~~~~~
                
                '... and append the Voters Roll node to the list
                
                
                xmlVotersRollListNode.appendChild xmlNode
                
                
            End If
        End If  'FBBureauRefCategory check
    Next
    
    'Append the Voters Roll list to the main response
    If xmlVotersRollListNode.hasChildNodes Then
        xmlBureauDataElem.appendChild xmlVotersRollListNode
    End If


    '=================================================================================
    ' Main Response
    '=================================================================================
    If Not xmlBureauDataElem.hasChildNodes Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
    End If
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set ICreditCheckBO_GetAndDeduplicateBureauRecords = xmlResponseElem
    
GetAndDeduplicateBureauRecordsExit:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplicationElement = Nothing
    Set xmlCustomerListElement = Nothing
    Set xmlGuidElement = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlTempRequestElem = Nothing
    Set xmlTempParentElem = Nothing
    Set xmlTempElement = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlAssociationListNode = Nothing
    Set xmlCustomerResponseList = Nothing
    Set xmlAliasResponseList = Nothing
    Set xmlCAISResponseList = Nothing
    Set xmlPublicInfoResponseList = Nothing
    Set xmlFBHeaderRequestElem = Nothing
    Set xmlFBHeaderElem = Nothing
    Set xmlFBHeaderBlockIdElem = Nothing
    Set xmlCustomerDetailsElem = Nothing
    Set xmlCAISElem = Nothing
    Set xmlPublicInfoElem = Nothing
    Set xmlBureauDataElem = Nothing
    Set xmlTempNode = Nothing   'BM0251 MDC 15/01/2003
    
    Exit Function
    
GetAndDeduplicateBureauRecordsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetAndDeduplicateBureauRecords = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetAndDeduplicateBureauRecordsExit

End Function
'BMIDS00336 MDC 14/08/2002 - End


'BMIDS00336 MDC 14/08/2002
Private Function CheckSuccessOrRecordNotFound(ByVal xmlResponseToCheck As IXMLDOMNode, _
                                            ByVal strFunction As String) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    CheckSuccessOrRecordNotFound
' author:       Mark Coleman
' date:         14/08/2002
' description:  Helper function to check if a response is Success or Record Not Found error
'               Used to ignore 'Record Not Found' errors
' pass:         xmlResponseToCheck  XML Input specifying the XML Response data to be checked
'               strFunction         Name of the calling function
' return:       True if success, False if record Not Found
' errors:       If an error other than Record Not Found, raise it
'------------------------------------------------------------------------------------------
On Error GoTo CheckSuccessOrRecordNotFoundVbErr

Const strFunctionName As String = "CheckSuccessOrRecordNotFound"
    
Dim lngErr As Long

    'Check response ignoring record not found error
    lngErr = m_objErrAssist.CheckResponse(xmlResponseToCheck, xmlResponseToCheck, False)
    If lngErr <> 0 Then
        lngErr = m_objErrAssist.GetOmigaErrorNumber(lngErr)
    End If
    If lngErr <> 0 And lngErr <> oeRecordNotFound Then
        m_objErrAssist.ThrowError TypeName(Me), strFunction, lngErr
    ElseIf lngErr = oeRecordNotFound Then
        CheckSuccessOrRecordNotFound = False
    Else
        CheckSuccessOrRecordNotFound = True
    End If

    Exit Function
    
CheckSuccessOrRecordNotFoundVbErr:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'BMIDS00336 MDC 14/08/2002 - End

Private Function ICreditCheckBO_GetCCAndFBStatus( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const cstrFunctionName As String = "ICreditCheckBO_GetCCAndFBStatus"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseElem As IXMLDOMElement
    Dim xmlTable As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlApplCreditCheckList As IXMLDOMNodeList

    Dim objICreditCheckDO As ICreditCheckDO

    ' Create default response
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set xmlTable = xmlOut.createElement("APPLICATIONCREDITCHECK")
    xmlTable.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTable.appendChild vXMLRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    Dim strCreditCheckGuid As String
    Dim bRecordFound As Boolean
    bRecordFound = True
    
    Set xmlTempResponseElem = objICreditCheckDO.FindApplicationCreditCheckList(xmlTable)
           
    If bRecordFound Then
        Set xmlApplCreditCheckList = xmlTempResponseElem.selectNodes(".//APPLICATIONCREDITCHECK")
        Dim iListLength As Integer, iLoopIndex As Integer
        iListLength = xmlApplCreditCheckList.length
        
        ' Append the details of last credit check run to Response
        Set xmlTable = xmlOut.createElement("LASTCREDITCHECK")
        xmlResponseElem.appendChild xmlTable
        Set xmlNode = xmlApplCreditCheckList.Item(iListLength - 1).cloneNode(True)
        xmlTable.appendChild xmlNode
                
        Dim strLastCreditCheckSuccess As String
        strLastCreditCheckSuccess = xmlNode.selectSingleNode(".//SUCCESSINDICATOR").Text
                        
        ' If the last credit check run was not successful, get the details of the last one which was successful
        
        If strLastCreditCheckSuccess = 1 Then
            strCreditCheckGuid = xmlNode.selectSingleNode(".//CREDITCHECKGUID").Text
        Else
            For iLoopIndex = iListLength - 2 To 0 Step -1
                Set xmlNode = xmlApplCreditCheckList.Item(iLoopIndex)
                If xmlNode.selectSingleNode(".//SUCCESSINDICATOR").Text = 1 Then
                    Set xmlTable = xmlOut.createElement("LASTSUCCESSFULCREDITCHECK")
                    xmlTable.appendChild xmlNode
                    xmlResponseElem.appendChild xmlTable
                    strCreditCheckGuid = xmlNode.selectSingleNode(".//CREDITCHECKGUID").Text
                    Exit For
                End If
            Next iLoopIndex
        End If
    End If

    ' Chek whether FullBureau Download was run for the last successful credit check run
    Set xmlTempResponseElem = Nothing
    If strCreditCheckGuid <> "" Then
        Set xmlTable = xmlOut.createElement("FULLBUREAUSTANDARDHEADER")
        Set xmlNode = xmlOut.createElement("CREDITCHECKGUID")
        xmlNode.Text = strCreditCheckGuid
        xmlTable.appendChild xmlNode
           
        bRecordFound = True
        Set xmlTempResponseElem = objICreditCheckDO.GetFullBureauStandardHeaderData(xmlTable)
         If bRecordFound Then
            xmlResponseElem.appendChild xmlTempResponseElem
        End If
    End If

    Set ICreditCheckBO_GetCCAndFBStatus = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlTable = Nothing
    Set xmlNode = Nothing
    Set xmlApplCreditCheckList = Nothing
    Set objICreditCheckDO = Nothing
    
    Exit Function
GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        bRecordFound = False
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCCAndFBStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit
End Function

Private Function ICreditCheckBO_GetCurrentAliasSummary( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentAliasSummary"


    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAlias As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlFBDataHeader As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode, _
        xmlTable As IXMLDOMNode, _
        xmlCorrectionRef As IXMLDOMNode, _
        xmlCorrectionLines As IXMLDOMNode

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'CIFAS' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlAlias = objICreditCheckDO.FindFullBureauAssociationsList(vXMLRequest)
    
    Dim strCorrectionRef As String
    
    For Each xmlNode In xmlAlias.selectNodes(".//FULLBUREAUASSOCIATIONS")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If strCorrectionRef <> "" Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
    Next xmlNode
    
    xmlResponseElem.appendChild xmlAlias
    
    Set ICreditCheckBO_GetCurrentAliasSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:
    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlAlias = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function
GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentAliasSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function ICreditCheckBO_GetCurrentCAPSSummary( _
                ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentCAPSSummary"


    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlCAPS As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlFBDataHeader As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode, _
        xmlTable As IXMLDOMNode, _
        xmlCorrectionRef As IXMLDOMNode, _
        xmlCorrectionLines As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'CIFAS' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlCAPS = objICreditCheckDO.FindFullBureauCAPSList(vXMLRequest)
    
    Dim strCorrectionRef As String
    
    For Each xmlNode In xmlCAPS.selectNodes(".//FULLBUREAUCAPS")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If strCorrectionRef <> "" Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
         
    Next xmlNode
    
    xmlResponseElem.appendChild xmlCAPS
    
    Set ICreditCheckBO_GetCurrentCAPSSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:

    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCAPS = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function

GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentCAPSSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit
End Function

Private Function ICreditCheckBO_GetCurrentCIFASSummary( _
                ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentCIFASSummary"


    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlCIFAS As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlFBDataHeader As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode, _
        xmlTable As IXMLDOMNode, _
        xmlCorrectionRef As IXMLDOMNode, _
        xmlCorrectionLines As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'CIFAS' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlCIFAS = objICreditCheckDO.FindFullBureauCIFASList(vXMLRequest)
    
    Dim strCorrectionRef As String
   
    For Each xmlNode In xmlCIFAS.selectNodes(".//FULLBUREAUCIFAS")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If strCorrectionRef <> "" Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
    Next xmlNode
    
    xmlResponseElem.appendChild xmlCIFAS
    
    Set ICreditCheckBO_GetCurrentCIFASSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:

    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCIFAS = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function

GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentCIFASSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function ICreditCheckBO_GetCurrentFullBureauResults( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentFullBureauResults"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement, xmlNode As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode, xmlTable As IXMLDOMNode

    Dim objICreditCheckDO As ICreditCheckDO
    
    'BMIDS744
    Dim objIDOAssist As IDOAssist
    Dim objIApplClassDef As IApplicationClassDef
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlAFFTPData As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement

    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
    Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strCCRNo As String, strCCGuid As String, strSequenceNo As String
    GetApplicationCreditCheckKeyData vXMLRequest, strCCRNo, strCCGuid
    
    ' Attach Credit Check key dat to response
    Set xmlNode = xmlOut.createElement("CREDITCHECKGUID")
    xmlNode.Text = strCCGuid
    xmlResponseElem.appendChild xmlNode
    
    Set xmlNode = xmlOut.createElement("CREDITCHECKREFERENCENUMBER")
    xmlNode.Text = strCCRNo
    xmlResponseElem.appendChild xmlNode
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    
    ' Fetch data from 'FullBureauStandardheader' and append it to xmlResponse
    Set xmlTable = xmlOut.createElement(cstrFULLLBUREAUSTANDARDHEADER_NAME)
    Set xmlNode = xmlOut.createElement("CREDITCHECKGUID")
    xmlNode.Text = strCCGuid
    xmlTable.appendChild xmlNode
    
    Set xmlTempResponse = objICreditCheckDO.GetFullBureauStandardHeaderData(xmlTable)
    xmlResponseElem.appendChild xmlTempResponse
    
    ' Fetch data from view 'FBRESULTS' and append to Response
    Set xmlTable = xmlOut.createElement("FBRESULTS")
    Set xmlNode = xmlOut.createElement("CREDITCHECKGUID")
    xmlNode.Text = strCCGuid
    xmlTable.appendChild xmlNode
    
    Set xmlTempResponse = objICreditCheckDO.FindCurrentFullBureauResultsList(xmlTable)
    xmlResponseElem.appendChild xmlTempResponse

    'BMIDS744 Need to get the OptOutIndicator
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vXMLRequest.ownerDocument)
    Set xmlTableElem = vXMLRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = vXMLRequest.selectSingleNode(".//" & "APPLICATIONNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = vXMLRequest.selectSingleNode(".//" & "APPLICATIONFACTFINDNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData(AFFThirdPartyData)
    Set xmlAFFTPData = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
     
    Set xmlElement = xmlApplClassDefDoc.createElement("OPTOUTINDICATOR")
    xmlElement.Text = xmlAFFTPData.selectSingleNode(".//" & "OPTOUTINDICATOR").Text
    xmlResponseElem.appendChild xmlElement
     
    Set ICreditCheckBO_GetCurrentFullBureauResults = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
     
GetExit:
    
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTable = Nothing
    Set xmlOut = Nothing
    Set objICreditCheckDO = Nothing

    'BMIDS744
    Set objIDOAssist = Nothing
    Set objIApplClassDef = Nothing
    Set xmlApplClassDefDoc = Nothing
    Set xmlAFFTPData = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    
    Exit Function
GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentFullBureauResults = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function ICreditCheckBO_GetCurrentPublicInfoSummary( _
                ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentPublicInfoSummary"

    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPublicInfo As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTable As IXMLDOMNode
    Dim xmlCorrectionRef As IXMLDOMNode
    Dim xmlCorrectionLines As IXMLDOMNode
    Dim strCorrectionRef As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'PublicInfo' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlPublicInfo = objICreditCheckDO.FindFullBureauPublicInfoList(vXMLRequest)
        
    For Each xmlNode In xmlPublicInfo.selectNodes(".//FULLBUREAUPUBLICINFO")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If Len(strCorrectionRef) > 0 Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
    Next xmlNode
    
    xmlResponseElem.appendChild xmlPublicInfo
    
    Set ICreditCheckBO_GetCurrentPublicInfoSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:

    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPublicInfo = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function

GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentPublicInfoSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function ICreditCheckBO_GetCurrentVotersRollSummary( _
            ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentVotersRollSummary"


    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlVotersRoll As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlFBDataHeader As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode, _
        xmlTable As IXMLDOMNode, _
        xmlCorrectionRef As IXMLDOMNode, _
        xmlCorrectionLines As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'VotersRoll' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlVotersRoll = objICreditCheckDO.FindFullBureauVotersRollList(vobjRequestNode)
    
    Dim strCorrectionRef As String
    
    For Each xmlNode In xmlVotersRoll.selectNodes(".//FULLBUREAUVOTERSROLL")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If strCorrectionRef <> "" Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
    Next xmlNode
    
    xmlResponseElem.appendChild xmlVotersRoll
    
    Set ICreditCheckBO_GetCurrentVotersRollSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:

    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlVotersRoll = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function

GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentVotersRollSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function GetFBCorrectionLines(ByVal vobjRequestNode As IXMLDOMElement, _
                                      ByVal strTableName As String) As IXMLDOMNode

    Dim objICreditCheckDO As ICreditCheckDO
    Dim xmlNewRequest As IXMLDOMNode, xmlResponse As IXMLDOMNode
    Dim lngErrNo As Long
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    
    Set xmlNewRequest = vobjRequestNode.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlNewRequest, strTableName, "FULLBUREAUCORRECTIONLINES"
    
    On Error Resume Next
    Set xmlResponse = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
    lngErrNo = Err.Number
     
    On Error GoTo 0
    If lngErrNo <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), "GetFBCorrectionLines", 294
    End If
     
    Set GetFBCorrectionLines = xmlResponse
    
    Set xmlResponse = Nothing
    Set xmlNewRequest = Nothing
    Set objICreditCheckDO = Nothing
    
End Function

Private Function GetCustomerAddress(ByVal vobjRequestNode As IXMLDOMElement, _
                                    ByVal strFBAddressIndicator As String) As IXMLDOMNode

'    Dim objCustomerBO As CustomerBO
'    Dim xmlNewRequest As IXMLDOMNode, xmlTable As IXMLDOMNode
'
'    Dim xmlDoc As New DomDocument, xmlNode As IXMLDOMNode
'    Dim xmlCustomerAddressList As IXMLDOMNodeList
'
'    Dim strResponse As String, lngErrNo As Long
'
'    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
'    Set xmlTable = xmlDoc.createElement("CUSTOMERADDRESS")
'    xmlNewRequest.appendChild xmlTable
'
'    xmlTable.appendChild vobjRequestNode.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
'    xmlTable.appendChild vobjRequestNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
'
'    ' append CustomerAddressType
'    Set xmlNode = xmlDoc.createElement("ADDRESSTYPE")
'    xmlNode.Text = IIf(strFBAddressIndicator = "C", "1", "3")
'    xmlTable.appendChild xmlNode
'
'    Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
'    strResponse = objCustomerBO.FindCustomerAddressList(xmlNewRequest.xml)
'
'    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), "GetCustomerAddress")
'    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement)
'
'    If lngErrNo <> 0 Then
'        m_objErrAssist.ThrowError TypeName(Me), "GetCustomerAddress", 295
'    End If
'
'    If strFBAddressIndicator = "C" Then
'        Set GetCustomerAddress = xmlDoc.selectSingleNode(".//CUSTOMERADDRESS")
'    Else
'        Set xmlCustomerAddressList = xmlDoc.selectNodes(".//CUSTOMERADDRESS")
'        If strFBAddressIndicator = "P" Then
'            Set GetCustomerAddress = xmlCustomerAddressList.Item(0)
'        Else
'            If strFBAddressIndicator = "3" Then
'                Set GetCustomerAddress = xmlCustomerAddressList.Item(1)
'            Else
'                Set GetCustomerAddress = xmlCustomerAddressList.Item(2)
'            End If
'        End If
'    End If
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'GetCustomerAddressExit:
'    Set xmlDoc = Nothing
'    Set xmlNewRequest = Nothing
'    Set xmlTable = Nothing
'    Set xmlNode = Nothing
'
'    Set objCustomerBO = Nothing
'
'    Exit Function
'
'
'GetCustomerAddressVbErr:
'    Set xmlDoc = Nothing
'    Set xmlNewRequest = Nothing
'    Set xmlTable = Nothing
'    Set xmlNode = Nothing
'
'    Set objCustomerBO = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.setcomplete
'    End If
'
'    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function GetFullBureauDataHeaderData(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetFullBureauDataHeaderData"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetFullBureauDataHeaderData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetFullBureauDataHeaderData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetFullBureauDataHeaderData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit


End Function

Public Function GetFullBureauRequest(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetFullBureauRequest"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ICreditCheckBO_GetFullBureauRequest(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    GetFullBureauRequest = xmlTempResponseNode.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetFullBureauRequest = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function ICreditCheckBO_GetCurrentFullBureauCAISSummary( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentFullBureauCAISSummary"

    Dim objICreditCheckDO As ICreditCheckDO

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlCAIS As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTable As IXMLDOMNode
    Dim xmlCorrectionRef As IXMLDOMNode
    Dim xmlCorrectionLines As IXMLDOMNode
    Dim strCorrectionRef As String
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    ' Build request for fetching records from 'FullBureauCorrectionLines'
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlTable = xmlOut.createElement("FULLBUREAUCORRECTIONLINES")
     xmlNewRequest.appendChild xmlTable
    
    Set xmlCorrectionRef = xmlOut.createElement("FBNOTICEOFCORRECTIONREF")
    xmlTable.appendChild xmlCorrectionRef
    
    'Fetch records from 'CAIS' table corresponding to this 'CreditCheckGuid' and 'BlockId'
    Set xmlCAIS = objICreditCheckDO.FindFullBureauCAISList(vXMLRequest)
        
    For Each xmlNode In xmlCAIS.selectNodes(".//FULLBUREAUCAIS")
        ' Fetch corresponding records from table 'CorrectionLines'
        strCorrectionRef = xmlNode.selectSingleNode(".//FBNOTICEOFCORRECTIONREF").Text
            
        If Len(strCorrectionRef) > 0 Then
            xmlCorrectionRef.Text = strCorrectionRef
                
            Set xmlCorrectionLines = objICreditCheckDO.FindFullBureauCorrectionLinesList(xmlNewRequest)
            xmlNode.appendChild xmlCorrectionLines
        End If
    Next xmlNode
    
    xmlResponseElem.appendChild xmlCAIS
    
    Set ICreditCheckBO_GetCurrentFullBureauCAISSummary = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetExit:

    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCAIS = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlCorrectionLines = Nothing
    Set xmlCorrectionRef = Nothing
        
    Exit Function

GetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentFullBureauCAISSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetExit

End Function

Private Function ICreditCheckBO_GetCurrentXMLCreditCheckResults(ByVal vobjRequestNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode

On Error GoTo ICreditCheckBO_GetCurrentXMLCreditCheckResultsVbErr

    Dim objICreditCheckDO As ICreditCheckDO

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set ICreditCheckBO_GetCurrentXMLCreditCheckResults = objICreditCheckDO.GetCurrentXMLCreditCheckResults(vobjRequestNode)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_GetCurrentXMLCreditCheckResultsExit:

    Set objICreditCheckDO = Nothing
    Exit Function
    
ICreditCheckBO_GetCurrentXMLCreditCheckResultsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_GetCurrentXMLCreditCheckResults = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_GetCurrentXMLCreditCheckResultsExit
    

End Function

Private Function ICreditCheckBO_GetFullBureauDataHeaderData( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_GetFullBureauDataHeaderData"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vXMLRequest.nodeName = cstrFULLBUREAUDATAHEADER_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrFULLBUREAUDATAHEADER_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLBUREAUDATAHEADER_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.GetFullBureauDataHeaderData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICreditCheckBO_GetFullBureauDataHeaderData = xmlResponseElem
    
    m_objContext.SetComplete
    
GetDataExit:
    
    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetFullBureauDataHeaderData = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function ICreditCheckBO_GetFullBureauRequest( _
                ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_GetFullBureauRequest"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode
    
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
        
    Set xmlTempResponseNode = ICreditCheckBO_ValidateFBMandatoryData(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set xmlTempResponseNode = objICreditCheckDO.GetFullBureauRequest(vXMLRequest)
    
    Set ICreditCheckBO_GetFullBureauRequest = xmlTempResponseNode
    
    m_objContext.SetComplete
    
GetDataExit:
    
    Set objICreditCheckDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetFullBureauRequest = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function ICreditCheckBO_GetFullBureauStandardHeaderData( _
                ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_GetFullBureauStandardHeaderData"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
                
    If vXMLRequest.nodeName = cstrFULLLBUREAUSTANDARDHEADER_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrFULLLBUREAUSTANDARDHEADER_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrFULLLBUREAUSTANDARDHEADER_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICreditCheckDO.GetFullBureauStandardHeaderData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICreditCheckBO_GetFullBureauStandardHeaderData = xmlResponseElem
    
    m_objContext.SetComplete
    
GetDataExit:
    
    Set xmlRequestNode = Nothing
    Set objICreditCheckDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetFullBureauStandardHeaderData = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function ICreditCheckBO_RunCreditScore(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' procedure:    ICreditCheckBO_RunCreditScore
' author:       Mark Coleman
' date:         22/08/2002
' description:  Runs the Credit Scoring process saving results from bureau
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo ICreditCheckBO_RunCreditScoreVbErr

    Dim objIBureauDO As IExperianDO
    
    If m_objContext Is Nothing Then
        Set objIBureauDO = New ExperianDO
    Else
        Set objIBureauDO = m_objContext.CreateInstance(gstrEXPERIAN_COMPONENT & ".ExperianDO")
    End If
    
    Set ICreditCheckBO_RunCreditScore = objIBureauDO.RunCreditCheck(vXMLRequest)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_RunCreditScoreExit:

    Set objIBureauDO = Nothing
    Exit Function

ICreditCheckBO_RunCreditScoreVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_RunCreditScore = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_RunCreditScoreExit

End Function

Private Function ICreditCheckBO_RunFullBureau( _
                ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

On Error GoTo ICreditCheckBO_RunFullBureauVbErr
    
    Const strFunctionName As String = "ICreditCheckBO_RunFullBureau"

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIBureauDO As IExperianDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIBureauDO = New ExperianDO
    Else
        Set objIBureauDO = m_objContext.CreateInstance(gstrEXPERIAN_COMPONENT & ".ExperianDO")
    End If
    
    ' Validate before running download
    Set xmlTempResponseNode = ValidateBeforeFBDownLoad(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    'BMIDS00336 MDC 03/09/2002
    m_objXmlAssist.CopyNode "SEQUENCENUMBER", "SEQUENCENUMBER", xmlTempResponseNode, vXMLRequest, False
    'BMIDS00336 MDC 03/09/2002 - End
    
    ' Build the Full Bureau Request tag and append it to the Request to be passed to Experian
    Set xmlTempResponseNode = ICreditCheckBO_GetFullBureauRequest(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    vXMLRequest.appendChild xmlTempResponseNode.selectSingleNode(".//FBREQUEST")
    vXMLRequest.appendChild xmlTempResponseNode.selectSingleNode(".//CREDITCHECKGUID")
    
    Set xmlTempResponseNode = objIBureauDO.RunFullBureau(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICreditCheckBO_RunFullBureau = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_RunFullBureauExit:

    Set objIBureauDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICreditCheckBO_RunFullBureauVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_RunFullBureau = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_RunFullBureauExit
End Function

Private Function ICreditCheckBO_ValidateFBMandatoryData( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ICreditCheckBO_ValidateFBMandatoryData"

        
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Mandatory data to be ascertained..
    
    Set ICreditCheckBO_ValidateFBMandatoryData = xmlResponseElem
    
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
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_ValidateFBMandatoryData = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume ValidateExit
    
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

Private Function ICreditCheckBO_GetCurrentCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo ICreditCheckBO_GetCurrentCreditCheckResultsVbErr

    Dim objICreditCheckDO As ICreditCheckDO

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set ICreditCheckBO_GetCurrentCreditCheckResults = objICreditCheckDO.GetCurrentCreditCheckResults(vobjRequestNode)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_GetCurrentCreditCheckResultsExit:

    Set objICreditCheckDO = Nothing
    Exit Function
    
ICreditCheckBO_GetCurrentCreditCheckResultsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_GetCurrentCreditCheckResults = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_GetCurrentCreditCheckResultsExit
    
End Function

Private Function ICreditCheckBO_GetCurrentCAISSummary(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo GetCurrentCAISSummaryVbErr

    Const strFunctionName As String = "ICreditCheckBO_GetCurrentCAISSummary"


    Dim objICreditCheckDO As ICreditCheckDO
    Dim objResponse As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set objResponse = objICreditCheckDO.GetCurrentCAISSummary(vobjRequestNode)
    Set ICreditCheckBO_GetCurrentCAISSummary = objResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentCAISSummaryExit:
    Set objICreditCheckDO = Nothing
    Exit Function

GetCurrentCAISSummaryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_GetCurrentCAISSummary = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetCurrentCAISSummaryExit

End Function

Private Function ICreditCheckBO_GetDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo ICreditCheckBO_GetDemoOptionsVbErr

    Dim objICreditCheckDO As ICreditCheckDO

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set ICreditCheckBO_GetDemoOptions = objICreditCheckDO.GetDemoOptions(vobjRequestNode)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_GetDemoOptionsExit:

    Set objICreditCheckDO = Nothing
    Exit Function

ICreditCheckBO_GetDemoOptionsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_GetDemoOptions = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_GetDemoOptionsExit

End Function

Private Function ICreditCheckBO_RunCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' procedure:    ICreditCheckBO_RunCreditCheck
' author:       Mark Coleman
' date:         22/08/2002
' description:  Run Credit Score followed by a Full Bureau download
' pass:         vobjRequestNode  XML Input request specifying the data and operation
' return:       XML Response to indicate Success or Error
'------------------------------------------------------------------------------------------

On Error GoTo ICreditCheckBO_RunCreditCheckVbErr

Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMElement
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
Dim xmlApplicationNode As IXMLDOMNode
Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlTempNode As IXMLDOMNode

Dim objParamBO As GlobalParameterBO

Dim strResponse As String
Dim strLog As String
'BMIDS682
Dim strAddressTargeting As String

Const cstrFunctionName As String = "ICreditCheckBO_RunCreditCheck"

    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    
    'Validate Request
    
    Set xmlTempNode = m_objXmlAssist.GetMandatoryNode(vobjRequestNode, _
            "APPLICATION[(APPLICATIONNUMBER != '') and (APPLICATIONFACTFINDNUMBER != '')]")
    Set xmlTempNode = m_objXmlAssist.GetMandatoryNode(vobjRequestNode, _
            "CUSTOMERLIST[CUSTOMER[(CUSTOMERNUMBER != '') and (CUSTOMERVERSIONNUMBER != '')]]")
    
    'Determine if logging Bureau requests and Responses
    Set objParamBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    strResponse = objParamBO.GetCurrentParameter("CreditCheckLog")
    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
        strLog = m_objXmlAssist.GetElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='CreditCheckLog']/BOOLEAN")
    End If
    
    Set xmlAttrib = vobjRequestNode.ownerDocument.createAttribute("LOG")
    xmlAttrib.Value = strLog
    vobjRequestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    
    'Run the Credit Score...
    Set xmlTempResponse = ICreditCheckBO_RunCreditScore(vobjRequestNode)
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    
    'BMIDS682 Only proceed to FullBureau if we are not doing addresstargeting
    strAddressTargeting = m_objXmlAssist.GetNodeValue(xmlTempResponse, ".//ADDRESSTARGETING", True, True)
    
    'If Address Targeting is Required Don't run FullBureau
    If (StrComp(strAddressTargeting, "NO") = 0) Then
    
        Set xmlApplicationNode = m_objXmlAssist.GetMandatoryNode(vobjRequestNode, "APPLICATION")
        Set xmlTempRequest = vobjRequestNode.cloneNode(False)
        xmlTempRequest.setAttribute "LOG", strLog
        m_objXmlAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", xmlApplicationNode, xmlTempRequest, False
        m_objXmlAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", xmlApplicationNode, xmlTempRequest, False
        
        '...and then the Full Bureau Download
        Set xmlTempResponse = ICreditCheckBO_RunFullBureau(xmlTempRequest)
        m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
        
        'Response
        xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        Set ICreditCheckBO_RunCreditCheck = xmlResponseElem
    Else
        'return our Address Targeting response instead
        Set ICreditCheckBO_RunCreditCheck = xmlTempResponse
    End If
    
ICreditCheckBO_RunCreditCheckExit:
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlResponseElem = Nothing
    Set xmlTempResponse = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objParamBO = Nothing
    Set xmlAttrib = Nothing
    Set xmlTempNode = Nothing
    Exit Function

ICreditCheckBO_RunCreditCheckVbErr:
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set ICreditCheckBO_RunCreditCheck = m_objErrAssist.CreateErrorResponseEx()
    
    '   go to clean-up section
    Resume ICreditCheckBO_RunCreditCheckExit


End Function

Private Function ICreditCheckBO_RunUpgradeToFullAppCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' procedure:    ICreditCheckBO_RunUpgradeToFullAppCreditCheck
' author:       Helen Aldred
' date:         28/07/2004
' description:  Upgrades an enquiry 'footprint' Credit Score to a Full Application 'footprint'
' pass:         vobjRequestNode  XML Input request specifying the data and operation
' return:       XML Response to indicate Success or Error
'------------------------------------------------------------------------------------------

    On Error GoTo ICreditCheckBO_RunUpgradeToFullAppCreditCheckVbErr

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlTempNode As IXMLDOMNode

    Dim objParamBO As GlobalParameterBO
    Dim objIBureauDO As IExperianDO
    
    Dim strResponse As String
    Dim strLog As String

    Const cstrFunctionName As String = "ICreditCheckBO_RunUpgradeToFullAppCreditCheck"
    
    If m_objContext Is Nothing Then
        Set objIBureauDO = New ExperianDO
    Else
        Set objIBureauDO = m_objContext.CreateInstance(gstrEXPERIAN_COMPONENT & ".ExperianDO")
    End If

    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    
    'Validate Request
    
    Set xmlTempNode = m_objXmlAssist.GetMandatoryNode(vobjRequestNode, _
            "APPLICATION[(@APPLICATIONNUMBER != '') and (@APPLICATIONFACTFINDNUMBER != '')]")
    
    'Determine if logging Bureau requests and Responses
    Set objParamBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    strResponse = objParamBO.GetCurrentParameter("CreditCheckLog")
    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), cstrFunctionName)
    If m_objErrAssist.CheckResponse(xmlTempResponseDoc.documentElement, xmlResponseElem, False) = 0 Then
        strLog = m_objXmlAssist.GetElementText(xmlTempResponseDoc.documentElement, "GLOBALPARAMETER[NAME='CreditCheckLog']/BOOLEAN")
    End If
    
    Set xmlAttrib = vobjRequestNode.ownerDocument.createAttribute("LOG")
    xmlAttrib.Value = strLog
    vobjRequestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    
    'Run the Upgrade
    Set ICreditCheckBO_RunUpgradeToFullAppCreditCheck = objIBureauDO.RunUpgradeToFullAppCreditCheck(vobjRequestNode)
    
ICreditCheckBO_RunUpgradeToFullAppCreditCheckExit:
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlResponseElem = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objParamBO = Nothing
    Set objIBureauDO = Nothing
    Set xmlAttrib = Nothing
    Set xmlTempNode = Nothing
    Exit Function

ICreditCheckBO_RunUpgradeToFullAppCreditCheckVbErr:
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set ICreditCheckBO_RunUpgradeToFullAppCreditCheck = m_objErrAssist.CreateErrorResponseEx()
    
    ' go to clean-up section
    Resume ICreditCheckBO_RunUpgradeToFullAppCreditCheckExit


End Function

Public Function RunCreditCheck(ByVal vstrXMLRequest As String) As String

On Error GoTo RunCreditCheckVbErr
    
    Const strFunctionName As String = "RunCreditCheck"

    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlTempResponseNode = ICreditCheckBO_RunCreditCheck(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode
   
    RunCreditCheck = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RunCreditCheckExit:
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

RunCreditCheckVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    RunCreditCheck = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume RunCreditCheckExit

End Function

'BMIDS748
Public Function RunUpgradeToFullAppCreditCheck(ByVal vstrXMLRequest As String) As String

On Error GoTo RunUpgradeToFullAppCreditCheckVbErr
    
    Const strFunctionName As String = "RunUpgradeToFullAppCreditCheck"

    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlTempResponseNode = ICreditCheckBO_RunUpgradeToFullAppCreditCheck(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode
   
    RunUpgradeToFullAppCreditCheck = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RunUpgradeToFullAppCreditCheckExit:
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

RunUpgradeToFullAppCreditCheckVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    RunUpgradeToFullAppCreditCheck = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume RunUpgradeToFullAppCreditCheckExit

End Function

Public Function GetDemoOptions(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDemoOptionsVbErr

Const strFunctionName As String = "GetDemoOptions"

Dim objXmlIn As New FreeThreadedDOMDocument40
objXmlIn.validateOnParse = False
objXmlIn.setProperty "NewParser", True
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode

    objXmlOut.async = False
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlTempResponseNode = ICreditCheckBO_GetDemoOptions(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode
    
    GetDemoOptions = objXmlOut.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDemoOptionsExit:
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    Exit Function
    
GetDemoOptionsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetDemoOptions = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetDemoOptionsExit

End Function

Public Function UpdateDemoOptions(ByVal vstrXMLRequest As String) As String

On Error GoTo UpdateDemoOptionsVbErr

Const strFunctionName As String = "UpdateDemoOptions"

Dim objXmlIn As FreeThreadedDOMDocument40
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode

    objXmlOut.async = False
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlTempResponseNode = ICreditCheckBO_UpdateDemoOptions(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode
    
    UpdateDemoOptions = objXmlOut.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateDemoOptionsExit:
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    Exit Function
    
UpdateDemoOptionsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    UpdateDemoOptions = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume UpdateDemoOptionsExit


End Function

Private Function ICreditCheckBO_UpdateDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo ICreditCheckBO_UpdateDemoOptionsVbErr
    
    Dim objICreditCheckDO As ICreditCheckDO

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set ICreditCheckBO_UpdateDemoOptions = objICreditCheckDO.UpdateDemoOptions(vobjRequestNode)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_UpdateDemoOptionsExit:

    Set objICreditCheckDO = Nothing
    Exit Function

ICreditCheckBO_UpdateDemoOptionsVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_UpdateDemoOptions = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_UpdateDemoOptionsExit

End Function

Public Function FindFullBureauAssociationsList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauAssociationsList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauAssociationsList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauAssociationsList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauAssociationsList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function FindFullBureauVotersRollList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauVotersRollList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauVotersRollList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauVotersRollList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauVotersRollList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

'SR 18/10/2005 : MAR24
Public Function FindApplicationCreditCheckDetailsList(ByVal vstrXMLRequest As String) As String
On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindApplicationCreditCheckDetailsList"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindApplicationCreditCheckDetailsList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindApplicationCreditCheckDetailsList = xmlResponseElem.xml
    
    m_objContext.SetComplete
FindListExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
      
    Exit Function
FindListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindApplicationCreditCheckDetailsList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
End Function

Public Function FindApplicationCreditCheckList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindApplicationCreditCheckList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindApplicationCreditCheckList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindApplicationCreditCheckList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindApplicationCreditCheckList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function FindFullBureauCorrectionLinesList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauCorrectionLinesList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauCorrectionLinesList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauCorrectionLinesList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauCorrectionLinesList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function FindFullBureauPublicInfoList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauPublicInfoList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauPublicInfoList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauPublicInfoList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauPublicInfoList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function FindFullBureauCAPSList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauCAPSList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauCAPSList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauCAPSList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauCAPSList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function FindFullBureauCIFASList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindFullBureauCIFASList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_FindFullBureauCIFASList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindFullBureauCIFASList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindFullBureauCIFASList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Public Function GetCurrentVotersRollSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentVotersRollSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentVotersRollSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentVotersRollSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentVotersRollSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCCAndFBStatus(ByVal vstrXMLRequest As String) As String
'------------------------------------------------------------------------------------
'Get the status data of last credit check (from table ApplicationCreditCheck'.
'Also, If it was not successful, get the latest status data of latestCreditCheck that
'was successful. Then find out whether Full Bureau was run for the last successful
'credit check run
'------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const cstrFunctionName As String = "GetCCAndFBStatus"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCCAndFBStatus(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCCAndFBStatus = xmlResponseElem.xml
    
    m_objContext.SetComplete
    
GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetCCAndFBStatus = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentCIFASSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentCIFASSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentCIFASSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentCIFASSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentCIFASSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentCAPSSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentCAPSSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentCAPSSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentCAPSSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentCAPSSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentAliasSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentAliasSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentAliasSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentAliasSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentAliasSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentFullBureauResults(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentFullBureauResults"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentFullBureauResults(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentFullBureauResults = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentFullBureauResults = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentFullBureauCAISSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentFullBureauCAISSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentFullBureauCAISSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentFullBureauCAISSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentFullBureauCAISSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetCurrentPublicInfoSummary(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCurrentPublicInfoSummary"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentPublicInfoSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCurrentPublicInfoSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCurrentPublicInfoSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function GetFullBureauStandardHeaderData(ByVal vstrXMLRequest As String) As String

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetFullBureauStandardHeaderData"
        
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_GetFullBureauStandardHeaderData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetFullBureauStandardHeaderData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetFullBureauStandardHeaderData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDataExit

End Function



'DRC 24/02/2004 BMIDS693
'
Private Function ToUpperAlphaOnly(ByVal vstrIn As String) As String
' header ----------------------------------------------------------------------------------
' procedure:    ToUpperAlphaOnly
' author:       Dave Crossley
' date:         12/02/2004
' description:  Helper function to remove any non alpha characters & return
'               result in uppercase
' pass:         Raw string
' return:       String containing only upper case letters
' errors:       If an error found, raise it
'------------------------------------------------------------------------------------------
On Error GoTo ToUpperAlphaOnlyErr

Const strFunctionName As String = "ToUpperAlphaOnly"

Dim strOut As String
strOut = ""
Dim intIndex As Integer
intIndex = 1
Dim strChar As String
Dim intTextLen, intAscValue As Integer
intTextLen = Len(vstrIn)

    
    While intIndex <= intTextLen
       strChar = Mid$(vstrIn, intIndex, 1)
       intAscValue = Asc(strChar)
       If (intAscValue >= 65 And intAscValue <= 90) Then
         ' uppercase
          strOut = strOut & strChar
       Else
          If (intAscValue >= 97 And intAscValue <= 122) Then
          ' lowercase
           strOut = strOut & UCase$(strChar)
          End If
       End If
       intIndex = intIndex + 1
    
    Wend
       

  ToUpperAlphaOnly = strOut
  Exit Function
    
ToUpperAlphaOnlyErr:
    
    Err.Raise Err.Number, Err.Source, Err.Description

  
End Function

'BMIDS744
Public Function GetCurrentCreditCheckStatus(ByVal vstrXMLRequest As String) As String

On Error GoTo GetCurrentCreditCheckStatusVbErr
    
    Const strFunctionName As String = "GetCurrentCreditCheckStatus"

    
Dim objXmlIn As FreeThreadedDOMDocument40
Dim objXmlResponseElem As IXMLDOMElement
Dim objXmlDataNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim xmlTempResponseNode As IXMLDOMNode
    
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseNode = ICreditCheckBO_GetCurrentCreditCheckStatus(objXmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    m_objXmlAssist.AttachResponseData objXmlResponseElem, xmlTempResponseNode

    GetCurrentCreditCheckStatus = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentCreditCheckStatusExit:
        
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    
    Exit Function

GetCurrentCreditCheckStatusVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetCurrentCreditCheckStatus = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetCurrentCreditCheckStatusExit
End Function


Private Function ICreditCheckBO_GetCurrentCreditCheckStatus(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo ICreditCheckBO_GetCurrentCreditCheckStatusVbErr

    Dim objICreditCheckDO As ICreditCheckDO

    If m_objContext Is Nothing Then
        Set objICreditCheckDO = New CreditCheckDO
    Else
        Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    End If
    
    Set ICreditCheckBO_GetCurrentCreditCheckStatus = objICreditCheckDO.GetCreditCheckStatus(vobjRequestNode)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICreditCheckBO_GetCurrentCreditCheckStatusExit:

    Set objICreditCheckDO = Nothing
    Exit Function
    
ICreditCheckBO_GetCurrentCreditCheckStatusVbErr:
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set ICreditCheckBO_GetCurrentCreditCheckStatus = m_objErrAssist.CreateErrorResponseEx()
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ICreditCheckBO_GetCurrentCreditCheckStatusExit
    
End Function

'INR 23/11/2005 : MAR645
Public Function CreateRuleOverride(ByVal vstrXMLRequest As String) As String
On Error GoTo CreateRuleOverrideVbErr

    Const strFunctionName As String = "CreateRuleOverride"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
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
    Set xmlTempResponseNode = ICreditCheckBO_CreateRuleOverride(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateRuleOverride = xmlResponseElem.xml
    
    m_objContext.SetComplete
CreateRuleOverrideExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
      
    Exit Function
CreateRuleOverrideVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateRuleOverride = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CreateRuleOverrideExit
End Function

'INR 23/11/2005 : MAR645
Private Function ICreditCheckBO_CreateRuleOverride( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateRuleOverrideVbErr
    Const strFunctionName As String = "ICreditCheckBO_CreateRuleOverride"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICreditCheckDO As ICreditCheckDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICreditCheckDO = m_objContext.CreateInstance(App.Title & ".CreditCheckDO")
    If vXMLRequest.nodeName = "SMRULEOVERRIDE" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("SMRULEOVERRIDE").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "SMRULEOVERRIDE" & " tag not found"
    End If

    Set xmlDataNode = objICreditCheckDO.CreateRuleOverride(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set ICreditCheckBO_CreateRuleOverride = xmlResponseElem
CreateRuleOverrideExit:
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objICreditCheckDO = Nothing
    
    Exit Function
CreateRuleOverrideVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICreditCheckBO_CreateRuleOverride = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CreateRuleOverrideExit
End Function


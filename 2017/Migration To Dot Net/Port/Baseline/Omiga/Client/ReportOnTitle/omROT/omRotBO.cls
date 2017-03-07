VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omRotBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Code template for Report on Title Business Object

'Dependencies:  omRotTxBO, omRotDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' JR    05/03/01    Created.
' JR    21/03/01    AQR SYS2125: GetReportOnTitleData returns object not set error if XML response is empty
'                                ValidateReportOnTitle comment out ref to AppProc to avoid build errors
' JR    21/03/01    SYS2048:     Include more functionlaity on ValidateReportonTitle and complete
'                                FindRotFurtherQuestionsList method
' JR    20/04/01                 Completed Validation method now that IK has implemented ROT Rules.
'                                Added UpdateROTAddress
'GD     20/11/01    SYS2989      FindROTFurtherQuestionsList - Raise an error if RECORD NOT FOUND in call to
'                                   omAppProc.AppProcBO.FindAdditionalQuestionList, and pass to GUI
' DB    20/12/01    SYS3520      Changes to ValidateReportOnTitle
' CL    13/05/02    SYS4565      Changes to 'DoRequest' and addition of function 'DoClientRequest'
'-------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext
Private gobjTxBO As IomRotTxBO
Private gobjAppBO As ApplicationBO
Private gobjAppProcBO As omAppProcBO


'-----------------------------------------------------------------------------------------------
'BMIDS History.
'
'Prog   Date        Description
'DB     13/11/2002  BMIDS00862 Removed any solicitors validation as part of ROT.
'DRC    08/12/2003  BMIDS668 - Changed Offer expiry message
'
'-------------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function OmRotRequest(ByVal vstrXMLRequest As String) As String

    On Error GoTo OmRotRequestVbErr
    
    Const strFunctionName As String = "OmRotRequest"
    
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
    
    If gobjContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If
    gobjContext.SetComplete

    ' ensure global objects released before running request
    Set gobjTxBO = Nothing
    'Set gobjDO = Nothing
    
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
    
    OmRotRequest = xmlOutDoc.xml
    Debug.Print OmRotRequest
    
OmRotRequestExit:
    
    ' clean up ================================================================
    
    ' release global objects
    Set gobjTxBO = Nothing
    'Set gobjDO = Nothing

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

OmRotRequestVbErr:
    
    ' create ERROR response block from Err object
    OmRotRequest = errCreateErrorResponse()
    Debug.Print OmRotRequest
    
    ' perform clean up
    GoTo OmRotRequestExit

End Function

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

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
    
    strOperation = UCase$(strOperation)
    
    Select Case strOperation
    
        Case "CREATEREPORTONTITLE"
            CreateReportOnTitle vxmlRequestNode, vxmlResponseNode
        
        Case "GETREPORTONTITLEDATA"
            GetReportOnTitleData vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEREPORTONTITLE"
            UpdateReportOnTitle vxmlRequestNode, vxmlResponseNode
            
        Case "FINDROTFURTHERQUESTIONSLIST"
            FindROTFurtherQuestionsList vxmlRequestNode, vxmlResponseNode
        
        ' JR SYS2048
        Case "VALIDATEREPORTONTITLE"
            ValidateReportOnTitle vxmlRequestNode, vxmlResponseNode
                    
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateReportOnTitleExit
    
    Const strFunctionName As String = "CreateReportOnTitle"
    
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim xmlNode As IXMLDOMNode
    
    ' raise error if no REPORTONTITLE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "REPORTONTITLE")
    
    'get applicationnumber, applicationfactfindnumber
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".omRotTxBO")
    End If
    
    gobjTxBO.CreateReportOnTitle vxmlRequestNode, vxmlResponseNode
    
CreateReportOnTitleExit:
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub GetReportOnTitleData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetReportOnTitleDataExit
    
Const strFunctionName As String = "GetReportOnTitleData"
    
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlTempAddressElement As IXMLDOMElement
Dim xmlTempElement As IXMLDOMElement

Dim strAppNumber As String
Dim strAppFFNumber As String
Dim strRotGuid As String

Dim objGetRotData As IomRotDO
    
    
    'Validate REQUEST & select REPORTONTITLE node
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "REPORTONTITLE")
    
    'Get Applicationnumber, Applicationfactfindnumber
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    Set objGetRotData = gobjContext.CreateInstance(App.Title & ".omRotDO")
    
    objGetRotData.GetRotAndSolBankAccount vxmlRequestNode, vxmlResponseNode
    
    Set xmlNode = vxmlResponseNode.selectSingleNode("REPORTONTITLE")
    'JR SYS2125 - raise error if returned response is empty
    If xmlNode Is Nothing Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    Set xmlNodeList = vxmlResponseNode.selectNodes("REPORTONTITLE")
    
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlTempAddressElement = xmlDoc.createElement("GETROTADDRESSES")
    
    For Each xmlNode In xmlNodeList
        'Create element SOLBANKACCT and append data to responsenode
        Set xmlTempElement = xmlDoc.createElement("SOLBANKACCT")
        xmlTempElement.setAttribute "BANKNAME", xmlGetAttributeText(xmlNode, "BANKNAME")
        xmlTempElement.setAttribute "BANKACCOUNTNAME", xmlGetAttributeText(xmlNode, "BANKACCOUNTNAME")
        xmlTempElement.setAttribute "BANKACCOUNTNUMBER", xmlGetAttributeText(xmlNode, "BANKACCOUNTNUMBER")
        xmlTempElement.setAttribute "BANKSORTCODE", xmlGetAttributeText(xmlNode, "BANKSORTCODE")
        xmlTempElement.setAttribute "ACCOUNTCURRENCY", xmlGetAttributeText(xmlNode, "ACCOUNTCURRENCY")
        Set xmlTempResponse = xmlDoc.appendChild(xmlTempElement)
        xmlNode.appendChild xmlTempResponse
    
        'Remove duplicated attributes
        xmlNode.Attributes.removeNamedItem "BANKNAME"
        xmlNode.Attributes.removeNamedItem "BANKACCOUNTNAME"
        xmlNode.Attributes.removeNamedItem "BANKACCOUNTNUMBER"
        xmlNode.Attributes.removeNamedItem "BANKSORTCODE"
        xmlNode.Attributes.removeNamedItem "ACCOUNTCURRENCY"
        
        'Complete XML response with address details
        strRotGuid = xmlGetAttributeText(xmlNode, "ROTGUID")
        xmlTempAddressElement.setAttribute "ROTGUID", strRotGuid
        xmlTempRequest.appendChild xmlTempAddressElement
        
        objGetRotData.GetRotAddresses xmlTempRequest, xmlNode
    Next
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetReportOnTitleDataExit:

    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlNodeList = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempAddressElement = Nothing
    Set xmlTempElement = Nothing
    Set objGetRotData = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateReportOnTitleExit
    
Const strFunctionName As String = "UpdateReportOnTitle"
    
Dim xmlNode As IXMLDOMNode

    ' validate REQUEST & select REPORTONTITLE node
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "REPORTONTITLE")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".omRotTxBO")
    End If
    
    gobjTxBO.UpdateReportOnTitle vxmlRequestNode, vxmlResponseNode
    
UpdateReportOnTitleExit:

    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindROTFurtherQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindROTFurtherQuestionsListExit
    
Const strFunctionName As String = "FindROTFurtherQuestionsList"
    
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlApplnQuestionsDoc As FreeThreadedDOMDocument40

Dim xmlPatternMatch As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlAdditionalQuestionsNode As IXMLDOMNode
Dim xmlApplnAddQuestionsNode As IXMLDOMNode

Dim xmlAdditionalQuestionList As IXMLDOMNodeList

Dim xmlRequestElement As IXMLDOMElement

Dim strResponse As String
Dim lngRet As Long
Dim strValue As String
Dim strQuestionRef As String
Dim strAppNumber As String
Dim strAppFFNumber As String
Dim iType As Integer
Dim strPattern As String 'xml pattern matching string
        
    ' Validate REQUEST & select APPLICATION node
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' Get mandatory fields
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    iType = xmlGetMandatoryAttributeText(xmlNode, "TYPE")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlApplnQuestionsDoc = New FreeThreadedDOMDocument40
    xmlApplnQuestionsDoc.validateOnParse = False
    xmlApplnQuestionsDoc.setProperty "NewParser", True
        
    If gobjAppProcBO Is Nothing Then
        Set gobjAppProcBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    End If
    
    '******************************************************************
    ' Get AdditionalQuestions from omAppProcBO
    '******************************************************************
    Set xmlRequestElement = xmlDoc.createElement("REQUEST")
    xmlRequestElement.setAttribute "OPERATION", "FindAdditionalQuestionList"
    Set xmlTempRequest = xmlRequestElement
    
    Set xmlRequestElement = xmlDoc.createElement("QUESTION")
    xmlRequestElement.setAttribute "TYPE", iType
    xmlRequestElement.setAttribute "DELETEFLAG", 0
    
    xmlTempRequest.appendChild xmlRequestElement
    
    strResponse = gobjAppProcBO.OmAppProcRequest(xmlTempRequest.xml)
    
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    
    'GD     20/11/01    SYS2989
    'If Not ((lngRet = 0) Or (lngRet = oeRecordNotFound)) Then
     '   errThrowError strFunctionName, lngRet
    'End If
    If Not ((lngRet = 0)) Then
        errThrowError strFunctionName, lngRet
    End If
    
    xmlDoc.loadXML strResponse
    Set xmlAdditionalQuestionsNode = xmlDoc.selectSingleNode(".//ADDITIONALQUESTIONS")
    If Not xmlAdditionalQuestionsNode Is Nothing Then
        Set xmlAdditionalQuestionList = xmlDoc.selectNodes(".//ADDITIONALQUESTIONS")
    End If
        
    '******************************************************************
    ' Get ApplnAdditionalQuestionsList
    '******************************************************************
    Set xmlRequestElement = xmlDoc.createElement("REQUEST")
    xmlRequestElement.setAttribute "OPERATION", "FindAppAdditionalQuestionList"
    
    Set xmlTempRequest = xmlRequestElement
       
    Set xmlRequestElement = xmlDoc.createElement("ADDITIONALQUESTIONS")
    xmlRequestElement.setAttribute "APPLICATIONNUMBER", strAppNumber
    xmlRequestElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNumber
    xmlRequestElement.setAttribute "TYPE", iType
            
    xmlTempRequest.appendChild xmlRequestElement
        
    strResponse = gobjAppProcBO.OmAppProcRequest(xmlTempRequest.xml)
            
    'Check response for errors
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    If Not ((lngRet = 0) Or (lngRet = oeRecordNotFound)) Then
        errThrowError strFunctionName, lngRet
    End If
    xmlApplnQuestionsDoc.loadXML strResponse
    
    '***************************************************************************
    ' Check for a corresponding record on ApplnAdditionalQuestionNode
    ' for each record on AdditionalQuestionNode.Extract required attributes and
    ' append to response if match or create a ApplnAdditionalQuestions Record
    ' if no match
    '***************************************************************************
    For Each xmlAdditionalQuestionsNode In xmlAdditionalQuestionList
        strQuestionRef = xmlGetMandatoryAttributeText(xmlAdditionalQuestionsNode, "QUESTIONREFERENCE")
        strPattern = ".//QUESTIONS[@QUESTIONREFERENCE='" & strQuestionRef & "']"
        Set xmlPatternMatch = xmlGetNode(xmlApplnQuestionsDoc, strPattern)
        If Not xmlPatternMatch Is Nothing Then
            ' Append required attributes the existing Additional Question node
            strValue = xmlGetAttributeText(xmlPatternMatch, "RESPONSE")
            xmlSetAttributeValue xmlAdditionalQuestionsNode, "RESPONSE", strValue
        
            strValue = xmlGetAttributeText(xmlPatternMatch, "FURTHERDETAILS")
            xmlSetAttributeValue xmlAdditionalQuestionsNode, "FURTHERDETAILS", strValue
        Else
            ' Call omAppProc.CreateAppAdditionalQuestions to create a record
            Set xmlRequestElement = xmlDoc.createElement("REQUEST")
            xmlRequestElement.setAttribute "OPERATION", "CreateAppAdditionalQuestions"
    
            Set xmlTempRequest = xmlRequestElement
       
            Set xmlRequestElement = xmlDoc.createElement("APPLNADDITIONALQUESTIONS")
            xmlRequestElement.setAttribute "APPLICATIONNUMBER", strAppNumber
            xmlRequestElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNumber
            xmlRequestElement.setAttribute "QUESTIONREFERENCE", strQuestionRef
            xmlRequestElement.setAttribute "TYPE", 20
            
            xmlTempRequest.appendChild xmlRequestElement
            
            strResponse = gobjAppProcBO.OmAppProcRequest(xmlTempRequest.xml)
            
            lngRet = errCheckXMLResponse(strResponse, False)
            If lngRet <> 0 Then
                lngRet = errGetOmigaErrorNumber(lngRet)
            End If
            If Not (lngRet = 0) Then
                errThrowError strFunctionName, lngRet
            End If
        End If
    
        vxmlResponseNode.appendChild xmlAdditionalQuestionsNode
    Next
    
FindROTFurtherQuestionsListExit:

    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlApplnQuestionsDoc = Nothing
    Set xmlPatternMatch = Nothing
    Set xmlTempRequest = Nothing
    Set xmlAdditionalQuestionsNode = Nothing
    Set xmlApplnAddQuestionsNode = Nothing
    Set xmlAdditionalQuestionList = Nothing
    Set xmlRequestElement = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ValidateReportOnTitle( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
'
' Code in NOT INHERITED
'
    On Error GoTo ValidateReportOnTitleExit
        
    Const strFunctionName As String = "ValidateReportOnTitle"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlIn As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTPResponseNode As IXMLDOMNode
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'Dim xmlROTResponseNode As IXMLDOMNode
    Dim xmlPanelResponseNode As IXMLDOMNode
    Dim xmlAppOfferResponseNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlTempNodeElement As IXMLDOMElement
    Dim xmlRotResponseElement As IXMLDOMElement
    
    Dim objATP As omApp.ApplicationThirdPartyBO
    Dim objTP As ThirdPartyBO
    
    Dim strRuleResult As String
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'Dim strRotBankAccountNumber As String
    'Dim strRotBankSortCode As String
    
    Dim strPnlBankAccountNumber As String
    Dim strPnlBankSortCode As String
    Dim strTPDirGuid As String
    Dim strPnlDirGuid As String
    
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'Dim strSolBankAccMessage As String
    
    Dim strReissueOfferMsg As String
    Dim strOfferIssueDate As String
    Dim strResponse As String
    
    Dim lngRet As Long

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objATP = gobjContext.CreateInstance("omApp.ApplicationThirdPartyBO")
    If gobjAppBO Is Nothing Then
        Set gobjAppBO = gobjContext.CreateInstance("omAPP.ApplicationBO")
    End If
    Set objTP = gobjContext.CreateInstance("omTP.ThirdPartyBO")
    
    'Validate REQUEST and retrieve the required attributes
    Set xmlIn = xmlGetRequestNode(vxmlRequestNode)
     
    'Select APPLICATION node
    Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
        
    'Get mandatory Applicationnumber and Applicationfactfindnumber
    strAppNumber = xmlGetMandatoryAttributeText(xmlIn, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlIn, "APPLICATIONFACTFINDNUMBER")
       
    'Set-up the XML Response Tag
    Set xmlRotResponseElement = xmlDoc.createElement("ROT")
    
    '*******************************************************************
    'Get Application Offer Data
    '*******************************************************************
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlTempElement = xmlDoc.createElement("APPLICATIONOFFER")
    
    Set xmlTempNodeElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlTempNodeElement.Text = strAppNumber
    xmlTempElement.appendChild xmlTempNodeElement
    
    Set xmlTempNodeElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempNodeElement.Text = strAppFFNumber
    xmlTempElement.appendChild xmlTempNodeElement
    
    xmlTempRequest.appendChild xmlTempElement
    
    strResponse = gobjAppBO.GetApplicationOfferData(xmlTempRequest.xml)
    
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
        errThrowError strFunctionName, lngRet
    End If
    
    xmlDoc.loadXML strResponse
    Set xmlAppOfferResponseNode = xmlDoc.selectSingleNode(".//APPLICATIONOFFER")
        
    If Not xmlAppOfferResponseNode Is Nothing Then
        strOfferIssueDate = xmlGetNodeAsDate(xmlAppOfferResponseNode, "OFFERISSUEDATE")
        If strOfferIssueDate <> "" Then
        If Date > xmlGetNodeAsDate(xmlAppOfferResponseNode, "OFFEREXPIRYDATE") Then
                xmlRotResponseElement.setAttribute "REISSUEOFFERIND", 1
                'DRC BMIDS668 - changed message
                strReissueOfferMsg = "The Offer has expired.  Re-underwrite the application."

               ' strReissueOfferMsg = "The Offer has expired. A re-issue offer task has " & _
               '                      "been generated"
                xmlRotResponseElement.setAttribute "REISSUEOFFERMESSAGE", strReissueOfferMsg
            End If
        End If
    End If
    
    RunRotRules vxmlRequestNode, vxmlResponseNode
    
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    '*******************************************************************
    'Get ROT Solicitors Details
    '*******************************************************************
'    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'    Set xmlTempElement = xmlDoc.createElement("REPORTONTITLE")
'    xmlTempElement.setAttribute "APPLICATIONNUMBER", strAppNumber
'    xmlTempElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNumber
'    xmlTempRequest.appendChild xmlTempElement
'
'    Set xmlTempResponseNode = xmlDoc.createElement("RESPONSE")
'
'    GetReportOnTitleData xmlTempRequest, xmlTempResponseNode
'
'    Set xmlROTResponseNode = xmlTempResponseNode.selectSingleNode(".//SOLBANKACCT")
'    If Not xmlROTResponseNode Is Nothing Then
'        strRotBankSortCode = xmlGetAttributeText(xmlROTResponseNode, "BANKSORTCODE")
'        strRotBankAccountNumber = xmlGetAttributeText(xmlROTResponseNode, "BANKACCOUNTNUMBER")
'    End If
     'DB END

    '*******************************************************************
    'Get Application Third Party solicitors details
    '*******************************************************************
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlTempElement = xmlDoc.createElement("APPLICATIONLEGALREP")

    Set xmlTempNodeElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlTempNodeElement.Text = strAppNumber
    xmlTempElement.appendChild xmlTempNodeElement

    Set xmlTempNodeElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempNodeElement.Text = strAppFFNumber
    xmlTempElement.appendChild xmlTempNodeElement

    xmlTempRequest.appendChild xmlTempElement

    strResponse = objATP.GetLegalRep(xmlTempRequest.xml)

    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
        errThrowError strFunctionName, lngRet
    End If

    xmlDoc.loadXML strResponse
    Set xmlTPResponseNode = xmlDoc.selectSingleNode(".//APPLICATIONLEGALREP")
    If Not xmlTPResponseNode Is Nothing Then
        strTPDirGuid = xmlGetNodeText(xmlTPResponseNode, "DIRECTORYGUID")
    End If

    '*********************************************************************
    'Get Panel bank account data using ThirdPartyBO.GetThirdParty method
    '*********************************************************************
    If strTPDirGuid <> "" Then
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        Set xmlTempElement = xmlDoc.createElement("NAMEANDADDRESSDIRECTORY")
        
        Set xmlTempNodeElement = xmlDoc.createElement("DIRECTORYGUID")
        xmlTempNodeElement.Text = strTPDirGuid
        xmlTempElement.appendChild xmlTempNodeElement
        
        xmlTempRequest.appendChild xmlTempElement
              
        strResponse = objTP.GetThirdParty(xmlTempRequest.xml)
                
        lngRet = errCheckXMLResponse(strResponse, False)
        If lngRet <> 0 Then
            lngRet = errGetOmigaErrorNumber(lngRet)
        End If
        If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
            errThrowError strFunctionName, lngRet
        End If
                
        xmlDoc.loadXML strResponse
        Set xmlTempResponseNode = xmlDoc.selectSingleNode(".//PANELBANKACCOUNTLIST")
        If Not xmlTempResponseNode Is Nothing Then
            Set xmlPanelResponseNode = xmlTempResponseNode.selectSingleNode(".//PANELBANKACCOUNT")
            Set xmlNodeList = xmlTempResponseNode.selectNodes("PANELBANKACCOUNT")
                  
            For Each xmlPanelResponseNode In xmlNodeList
                If strTPDirGuid = xmlGetNodeText(xmlPanelResponseNode, "DIRECTORYGUID") Then
                    strPnlDirGuid = xmlGetNodeText(xmlPanelResponseNode, "DIRECTORYGUID")
                    strPnlBankAccountNumber = xmlGetNodeText(xmlPanelResponseNode, "ACCOUNTNUMBER")
                    strPnlBankSortCode = xmlGetNodeText(xmlPanelResponseNode, "BANKSORTCODE")
                    
                    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids,
                    '                therefore they do not need to have any validation against them.
                    'DB SYS3520 20/12/01 - Inserted validate ROT into main IF statement.
'                    If strRotBankSortCode <> strPnlBankSortCode Or _
'                        strRotBankAccountNumber <> strPnlBankAccountNumber Then
'                        xmlRotResponseElement.setAttribute "SOLICITORBANKACCMATCH", 0
'                        strSolBankAccMessage = "Solicitor's bank details do not match those held on the " & _
'                        "database.  Please confirm bank details with Solicitor"
'                        xmlRotResponseElement.setAttribute "SOLICITORBANKACCMESSAGE", strSolBankAccMessage
'
'                    Else
'                        xmlRotResponseElement.setAttribute "SOLICITORBANKACCMATCH", 1
'                    End If
                     ' DB END
                        vxmlResponseNode.appendChild xmlRotResponseElement
                    ' DB End
                    Exit For
                End If
            Next
        End If
     End If
        
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'DB SYS3520 20/12/01 - If no guid value, use the solicitor bank details in ROT
    'If strTPDirGuid = "" Then
    '    xmlRotResponseElement.setAttribute "SOLICITORBANKACCMATCH", 1
    'End If
    'DB END
    
    vxmlResponseNode.appendChild xmlRotResponseElement
    ' DB End
           
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids,
    '                therefore do not need to validate them.
    '*******************************************************************
    'Validate ROT and panel bank account details
    '*******************************************************************
'    If strRotBankSortCode <> strPnlBankSortCode Or _
'    strRotBankAccountNumber <> strPnlBankAccountNumber Then
'        xmlRotResponseElement.setAttribute "SOLICITORBANKACCMATCH", 0
'        strSolBankAccMessage = "Solicitor's bank details do not match those held on the " & _
'                               "database.  Please confirm bank details with Solicitor"
'        xmlRotResponseElement.setAttribute "SOLICITORBANKACCMESSAGE", strSolBankAccMessage
'    End If
    
    'Comlpete Response by appending
'    vxmlResponseNode.appendChild xmlRotResponseElement
    'DB END
    
ValidateReportOnTitleExit:

    Set xmlIn = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTPResponseNode = Nothing
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'Set xmlROTResponseNode = Nothing
    Set xmlPanelResponseNode = Nothing
    Set xmlAppOfferResponseNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub RunRotRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo RunRotRulesExit
    Const strFunctionName As String = "RunRotRules"

    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40, _
        xmlThisResponseDoc As FreeThreadedDOMDocument40, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationNode As IXMLDOMNode, _
        xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement

    Dim objAPRules As Object, _
        gobjAppBOProc As Object
    
    Dim strComments As String
        
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "RunROTRules"
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    
    GetApplicationData xmlApplicationNode, xmlThisRequestNode
        
    GetGlobalParams xmlThisRequestNode
    
    GetAdditionalQuestions xmlThisRequestNode
    
    Set objAPRules = gobjContext.CreateInstance("omAPRules.APRulesBO")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    xmlThisResponseDoc.loadXML objAPRules.RunRules(xmlThisRequestDoc.xml)
    Set objAPRules = Nothing
    
    Set xmlResponseList = _
        xmlThisResponseDoc.selectNodes("RESPONSE/ROTRULE[@RESULT='0']")
    
    If xmlResponseList.length > 0 Then
    
        For Each xmlThisNode In xmlResponseList
        
            If Len(strComments) > 0 Then
                strComments = strComments & ";"
            End If
            
            strComments = strComments & xmlGetAttributeText(xmlThisNode, "MESSAGE")
        
        Next
        
        ' build CreateApplicationReviewHistory request =================================
        
        ' strip child nodes from request
        For Each xmlThisNode In xmlThisRequestNode.childNodes
            xmlThisRequestNode.removeChild xmlThisNode
        Next
        
        ' set operation value
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = _
            "CreateApplicationReviewHistory"
    
        Set xmlThisElem = xmlThisRequestDoc.createElement("APPLICATIONREVIEWHISTORY")
        xmlCopyAttribute xmlApplicationNode, xmlThisElem, "APPLICATIONNUMBER"
        xmlCopyAttribute xmlApplicationNode, xmlThisElem, "APPLICATIONFACTFINDNUMBER"
        xmlThisElem.setAttribute "REVIEWREASON", 90
        xmlSetSysDateToNodeAttrib xmlThisElem, "REVIEWDATETIME"
        xmlThisElem.setAttribute _
            "REVIEWUSERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
        xmlThisElem.setAttribute _
            "REVIEWUNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
        xmlThisElem.setAttribute "REVIEWCOMMENTS", strComments

        ' Complete the XML request by appending
        xmlThisRequestNode.appendChild xmlThisElem
        
        Debug.Print "omAppProc.omAppProcBO Request"
        Debug.Print xmlThisRequestDoc.xml

        Set gobjAppBOProc = gobjContext.CreateInstance("omAppProc.omAppProcBO")
        
        xmlThisResponseDoc.loadXML _
            gobjAppBOProc.OmAppProcRequest(xmlThisRequestDoc.xml)
        
        Debug.Print "omAppProc.omAppProcBO Response"
        Debug.Print xmlThisResponseDoc.xml
        
        Set gobjAppBOProc = Nothing
        
        errCheckXMLResponseNode _
            xmlThisResponseDoc.selectSingleNode("RESPONSE"), _
            vxmlResponseNode, True
    
    End If
    
RunRotRulesExit:

    Set objAPRules = Nothing
    Set gobjAppBOProc = Nothing
    
    Set xmlResponseList = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlThisNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetApplicationData( _
    ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
    ByVal vxmlParentNode As IXMLDOMNode)
    
    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
    
    vxmlParentNode.appendChild _
        xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION").cloneNode(True)
    
GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub GetGlobalParams(ByVal vxmlParentNode As IXMLDOMNode)
    
    On Error GoTo GetGlobalParamsExit
    Const cstrFunctionName As String = "GetGlobalParams"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("GLOBALPARAMETER")
    
    xmlElem.setAttribute "_SCHEMA_", "GLOBALPARAMETER"
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetGlobalParams request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetGlobalParams response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/GLOBALPARAMETER") Is Nothing Then
        errThrowError _
            cstrFunctionName, _
            oeXMLMissingElement, _
            "No GLOBALPARAMETER details"
    End If
    
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next
    
GetGlobalParamsExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub GetAdditionalQuestions(ByVal vxmlParentNode As IXMLDOMNode)
    
    On Error GoTo GetAdditionalQuestionsExit
    Const cstrFunctionName As String = "GetAdditionalQuestions"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("ADDITIONALQUESTIONS")
    
    xmlElem.setAttribute "_SCHEMA_", "ADDITIONALQUESTIONS"
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetAdditionalQuestions request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetAdditionalQuestions response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/ADDITIONALQUESTIONS") Is Nothing Then
        errThrowError _
            cstrFunctionName, _
            oeXMLMissingElement, _
            "No GLOBALPARAMETER details"
    End If
    
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next
    
GetAdditionalQuestionsExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub Terminate( _
    ByVal vxmlDoc As FreeThreadedDOMDocument40, _
    ByVal vstrFunctionName As String)
    
    vxmlDoc.save ("c:\temp\OmRotBO.xml")
    
    Shell _
        "c:\program files\plus!\microsoft internet\iexplore.exe c:\temp\OmRotBO.xml", _
        vbMaximizedFocus
    
    gobjContext.SetAbort
    
    Err.Raise _
        oeUnspecifiedError, _
        TypeName(Me) & "." & vstrFunctionName, _
        "OmTmBO Terminate Request"
        
End Sub


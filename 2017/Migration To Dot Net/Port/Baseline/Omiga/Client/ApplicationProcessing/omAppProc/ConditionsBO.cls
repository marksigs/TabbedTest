VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ConditionsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl
Implements IConditionsBO

Private gobjContext As ObjectContext

Private gobjTxBO As IConditionsTxBO
Private gobjDO As IConditionsDO

'BMIDS History:
'
'Prog   Date        Description
'GD     09/09/02    BMIDS00313 - APWP2 - BM004 Added FindNonProductConditionsList,FindAppProductConditionsList,GetApplicationData
'                                               Changed IConditionsBO_FindConditionsList
'GD     21/10/2002  BMIDS00650 Added IConditionsBO_GetApplicationData
'GD     05/11/2002  BMIDS00824 Changes to IConditionsBO_GetApplicationData to put COMBOs and GLOBALPARAMETERSs inside APPLICATION element
'GD     07/11/2002  BMIDS00876 Removed debugging code.
'MDC    19/12/2002  BM0206 Ensure that an error is thrown in GetAcceptedOrActiveQuoteData
'------------------------------------------------------------------------------------------

Private Sub IConditionsBO_CreateApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationConditionsExit

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strValue As String

    Const cstrFunctionName As String = "IConditionsBO_CreateApplicationConditions"

    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONCONDITIONS element has an APPLICATIONNUMBER attribute value
        For Each xmlNode In xmlNodeList
            xmlGetMandatoryAttributeText xmlNode, "APPLICATIONNUMBER"
            strValue = xmlGetAttributeText(xmlNode, "CONDITIONTEXT")
            xmlSetAttributeValue xmlNode, "CONDITIONABBREVIATIONTEXT", Left(strValue, 50)
        Next
    Else
        errThrowError cstrFunctionName, oeXMLMissingElement, "Missing APPLICATIONCONDITIONS element"
    End If
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ConditionsTxBO")
    End If
    
    gobjTxBO.CreateApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    
CreateApplicationConditionsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsBO_CreateDefaultApplicationConditions( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateDefaultApplicationConditionsExit

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strValue As String

    Const cstrFunctionName As String = "IConditionsBO_CreateDefaultApplicationConditions"

    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONCONDITIONS element has an APPLICATIONNUMBER attribute value
        For Each xmlNode In xmlNodeList
            xmlGetMandatoryAttributeText xmlNode, "APPLICATIONNUMBER"
        Next
    Else
        errThrowError cstrFunctionName, oeXMLMissingElement, "Missing APPLICATIONCONDITIONS element"
    End If
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ConditionsTxBO")
    End If
    
    gobjTxBO.CreateDefaultApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    
CreateDefaultApplicationConditionsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsBO_DeleteAllApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteAllApplicationConditionsExit

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strValue As String

    Const cstrFunctionName As String = "IConditionsBO_DeleteAllApplicationConditions"

    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONCONDITIONS element has an APPLICATIONNUMBER, APPLNCONDITIONSSEQ attribute value
        For Each xmlNode In xmlNodeList
             xmlGetMandatoryAttributeText xmlNode, "APPLICATIONNUMBER"
        Next
    Else
        errThrowError cstrFunctionName, oeXMLMissingElement, "Missing APPLICATIONCONDITIONS element"
    End If
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ConditionsTxBO")
    End If
    
    gobjTxBO.DeleteAllApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    
DeleteAllApplicationConditionsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsBO_DeleteApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteApplicationConditionsExit

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strValue As String

    Const cstrFunctionName As String = "IConditionsBO_DeleteApplicationConditions"

    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONCONDITIONS element has an APPLICATIONNUMBER, APPLNCONDITIONSSEQ attribute value
        For Each xmlNode In xmlNodeList
             xmlGetMandatoryAttributeText xmlNode, "APPLICATIONNUMBER"
            xmlGetMandatoryAttributeText xmlNode, "APPLNCONDITIONSSEQ"
        Next
    Else
        errThrowError cstrFunctionName, oeXMLMissingElement, "Missing APPLICATIONCONDITIONS element"
    End If
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ConditionsTxBO")
    End If
    
    gobjTxBO.DeleteApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    
DeleteApplicationConditionsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsBO_FindApplicationConditionsList( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindApplicationConditionsListExit

    Const strFunctionName As String = "IConditionsBO_FindApplicationConditionsList"
    
    Dim xmlApplicationConditions As IXMLDOMNode

     ' validate REQUEST & select CASETASK node
    Set xmlApplicationConditions = xmlGetMandatoryNode(vxmlRequestParentNode, "APPLICATIONCONDITIONS")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
    End If
    
    gobjDO.FindApplicationConditionsList xmlApplicationConditions, vxmlResponseNode
    
FindApplicationConditionsListExit:
    
    Set xmlApplicationConditions = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub IConditionsBO_FindConditionsList( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)


On Error GoTo FindConditionsListExit

    Const strFunctionName As String = "IConditionsBO_FindConditionsList"
    
    Dim iIndex As Integer
    
    Dim xmlConditions As IXMLDOMNode
    Dim objConditionsTxBO As ConditionsTxBO
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim xmlAppElem As IXMLDOMElement
    Dim xmlAppCondElem As IXMLDOMElement
    Dim xmlGetApplicationDataResp As IXMLDOMElement
    Dim xmlRunResolveDataRulesResp As IXMLDOMElement
    Dim xmlRunResolveDataRulesReq As IXMLDOMElement
    Dim xmlCondList As IXMLDOMNodeList
    Dim xmlCondElem As IXMLDOMElement
    Dim xmlCondListElem As IXMLDOMElement
    Dim xmlApplicationSchemaElem As IXMLDOMElement
    
    Dim xmlConditionsLogFileElem As IXMLDOMElement
    Dim strConditionsErrorFilePath As String
    
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
     ' validate REQUEST & select CONDITIONS node
    Set xmlConditions = xmlGetMandatoryNode(vxmlRequestParentNode, "CONDITIONS")
    Set objConditionsTxBO = New ConditionsTxBO
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
    End If
    Set xmlAppElem = vxmlRequestParentNode.selectSingleNode(".//APPLICATION")
    
    gobjDO.FindConditionsList xmlConditions, vxmlResponseNode
    If Not (xmlAppElem Is Nothing) Then
    

        '---------------------------Get Application Data-------------------------------------------
        strApplicationNumber = xmlGetMandatoryAttributeText(xmlAppElem, "APPLICATIONNUMBER")
        strApplicationFFNumber = xmlGetMandatoryAttributeText(xmlAppElem, "APPLICATIONFACTFINDNUMBER")
        Set xmlAppCondElem = xmlDoc.createElement("APPLICATIONCONDITIONS")
        Set xmlGetApplicationDataResp = xmlDoc.createElement("RESPONSE")
        xmlSetAttributeValue xmlAppCondElem, "APPLICATIONNUMBER", strApplicationNumber
        xmlSetAttributeValue xmlAppCondElem, "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
        'BMIDS00650
        'objConditionsTxBO.GetApplicationData xmlAppCondElem, xmlGetApplicationDataResp
        IConditionsBO_GetApplicationData xmlAppCondElem, xmlGetApplicationDataResp
        '----------------------------------------------------------------------------------------
        
        '----------------------------Run Resolve Data Rules----------------------------------------
        Set xmlRunResolveDataRulesResp = xmlDoc.createElement("RESPONSE")
        Set xmlCondList = vxmlResponseNode.selectNodes(".//CONDITIONS")
        Set xmlApplicationSchemaElem = xmlGetApplicationDataResp.selectSingleNode(".//APPLICATION")
        xmlChangeNodeName xmlApplicationSchemaElem, "APPLICATION", "SCHEMADATA"
        
        
        'Create CONDITIONSLIST and attach all the conditions to it.
        Set xmlCondListElem = xmlDoc.createElement("CONDITIONSLIST")
        For iIndex = 0 To (xmlCondList.length - 1)
            Set xmlCondElem = xmlCondList.Item(iIndex)
            xmlCondListElem.appendChild xmlCondElem
            
        Next
        Set xmlRunResolveDataRulesReq = xmlDoc.createElement("REQUEST")
        xmlRunResolveDataRulesReq.appendChild xmlCondListElem
        xmlRunResolveDataRulesReq.appendChild xmlApplicationSchemaElem
        'Get error file name XXXX
        'Set xmlConditionsLogFileElem = xmlDoc.createElement("CONDITIONRULESLOGFILE")
        'strConditionsErrorFilePath = GetGlobalParamString("OfferConditionResolveFailure")
        'xmlConditionsLogFileElem.setAttribute "FILENAME", strCondition
        'xmlRunResolveDataRulesReq.appendChild xmlConditionsLogFileElem
        
        
        
        objConditionsTxBO.RunResolveDataRules xmlRunResolveDataRulesReq, xmlRunResolveDataRulesResp
        'Dim i As Integer
        
        'i = errCheckError("FindConditionsList", "omAppProc.ConditionsBO", xmlRunResolveDataRulesResp)
        'errcheckerror
        '----------------------------------------------------------------------------------------
        'Append to response
        Set xmlCondList = xmlRunResolveDataRulesResp.selectNodes(".//CONDITIONS")
        For iIndex = 0 To (xmlCondList.length - 1)
            Set xmlCondElem = xmlCondList.Item(iIndex)
            vxmlResponseNode.appendChild xmlCondElem
        Next iIndex
        

    End If
    
FindConditionsListExit:
    
    Set xmlConditions = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub IConditionsBO_UpdateApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateApplicationConditionsExit

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strValue As String

    Const cstrFunctionName As String = "IConditionsBO_UpdateApplicationConditions"

    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONCONDITIONS element has an APPLICATIONNUMBER, APPLNCONDITIONSSEQ attribute value
        For Each xmlNode In xmlNodeList
             xmlGetMandatoryAttributeText xmlNode, "APPLICATIONNUMBER"
            xmlGetMandatoryAttributeText xmlNode, "APPLNCONDITIONSSEQ"
        Next
    Else
        errThrowError cstrFunctionName, oeXMLMissingElement, "Missing APPLICATIONCONDITIONS element"
    End If
        
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ConditionsTxBO")
    End If
    
    gobjTxBO.UpdateApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    
UpdateApplicationConditionsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

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

Private Sub IConditionsBO_FindNonProductConditionsList( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)


On Error GoTo FindNonProductConditionsListExit

    Const strFunctionName As String = "IConditionsBO_FindNonProductConditionsList"
    
    Dim xmlApplicationConditions As IXMLDOMNode

    Set xmlApplicationConditions = xmlGetMandatoryNode(vxmlRequestParentNode, "CONDITIONS")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
    End If
    
    gobjDO.FindNonProductConditionsList xmlApplicationConditions, vxmlResponseNode
    
FindNonProductConditionsListExit:
    
    Set xmlApplicationConditions = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsBO_FindAppProductConditionsList( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)


On Error GoTo FindAppProductConditionsListExit

    Const strFunctionName As String = "IConditionsBO_FindAppProductConditionsList"
    
'    Dim xmlApplicationConditions As IXMLDOMNode
'
'    Set xmlApplicationConditions = xmlGetMandatoryNode(vxmlRequestParentNode, "CONDITIONS")
    'Method in DO not required?
    '-----------------------------------------------------------------------------------------------------------
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    'STEP 0 Validate the REQUEST
    Dim xmlConditionsElem As IXMLDOMElement
    Dim xmlRequestClone As IXMLDOMElement
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim lngErrorNum As Long
    Dim strErrorDesc As String
    Dim blnRecordNotFoundNon As Boolean
    
    
    'Check CONDITIONS tag
    Set xmlConditionsElem = xmlGetMandatoryNode(vxmlRequestParentNode, ".//CONDITIONS")
    'Check APPLICATIONNUMBER present
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlConditionsElem, "APPLICATIONNUMBER")
    'Check APPLICATIONFACTFINDNUMBER present
    strApplicationFactFindNumber = xmlGetMandatoryAttributeText(xmlConditionsElem, "APPLICATIONFACTFINDNUMBER")
    'Will not reach here if CONDITIONS tag, ApplicationNumber or ApplicationFactFindNumber absent
    
    Set xmlRequestClone = vxmlRequestParentNode.cloneNode(True)
    
    'STEP 1 Call ConditionsBO.FindNonProductConditionsList
    '-----------------------------------------------------
    'no need to go through the public interface as method is in same class
    Dim xmlNonProductConditionsListResp As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlNonProductConditionsListResp = xmlOutDoc.appendChild(xmlResponseElem)
    Err.Clear
    On Error Resume Next
    IConditionsBO_FindNonProductConditionsList vxmlRequestParentNode, xmlNonProductConditionsListResp
    'Check for Record Not Found
    blnRecordNotFoundNon = False
    ' Need to ignore record not found error
    If Err.Number <> 0 Then
        lngErrorNum = errGetOmigaErrorNumber(Err.Number)
        strErrorDesc = Err.Description
    Else
        lngErrorNum = 0
    End If
    On Error GoTo FindAppProductConditionsListExit
    
    If lngErrorNum <> 0 Then
        If lngErrorNum <> oeRecordNotFound Then
            errThrowError strFunctionName, lngErrorNum, strErrorDesc
        Else
            blnRecordNotFoundNon = True
        End If
    End If

    
    '-----------------------------------------------------------------------------------------------------------
    
    'STEP 2 Set xmlNonProductConditionsListResp = Total RESPONSE from the call in STEP 1
    '------------------------------------------------------------------------------------
    
    'xmlNonProductConditionsListResp now contains the first piece of XML we're after
    
    '-----------------------------------------------------------------------------------------------------------
    
    
    'STEP 3 Call ApplicationQuoteBO.GetAcceptedOrActiveQuoteData(AppNum, AppFFNum)
    '-----------------------------------------------------------------------------
    '<REQUEST>
    '   <APPLICATION>
    '      <APPLICATIONNUMBER>S00004162</APPLICATIONNUMBER>
    '      <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
    '   </APPLICATION>
    '</REQUEST>
    Dim strAcceptedOrActiveQuoteRequest As String
    Dim strAcceptedOrActiveQuoteResponse As String
    Dim strMSQNumber As String
    Dim xmlApplicationNumberElem As IXMLDOMElement
    Dim xmlApplicationFFNumberElem As IXMLDOMElement
    Dim xmlMSQElem As IXMLDOMElement
    
    Dim objApplicationQuoteBO As ApplicationQuoteBO
    Dim xmlAcceptedOrActiveQuoteResponse As IXMLDOMElement
    'Convert from attributes to elements
    Set xmlApplicationNumberElem = xmlOutDoc.createElement("APPLICATIONNUMBER")
    xmlApplicationNumberElem.Text = strApplicationNumber
    Set xmlApplicationFFNumberElem = xmlOutDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlApplicationFFNumberElem.Text = strApplicationFactFindNumber
    'Build the Element-based REQUEST
    'BM0206 MDC 19/12/2002 - Ensure that an error is thrown in GetAcceptedOrActiveQuoteData
    'strAcceptedOrActiveQuoteRequest = "<REQUEST>"
    strAcceptedOrActiveQuoteRequest = "<REQUEST THROWERROR='1'>"
    'BM0206 MDC 19/12/2002 - End
    strAcceptedOrActiveQuoteRequest = strAcceptedOrActiveQuoteRequest & "<APPLICATION>"
    strAcceptedOrActiveQuoteRequest = strAcceptedOrActiveQuoteRequest & xmlApplicationNumberElem.xml
    strAcceptedOrActiveQuoteRequest = strAcceptedOrActiveQuoteRequest & xmlApplicationFFNumberElem.xml
    strAcceptedOrActiveQuoteRequest = strAcceptedOrActiveQuoteRequest & "</APPLICATION>"
    strAcceptedOrActiveQuoteRequest = strAcceptedOrActiveQuoteRequest & "</REQUEST>"
    
    If gobjContext Is Nothing Then
        Set objApplicationQuoteBO = New ApplicationQuoteBO
    Else
        Set objApplicationQuoteBO = _
            gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    End If
    strAcceptedOrActiveQuoteResponse = objApplicationQuoteBO.GetAcceptedOrActiveQuoteData(strAcceptedOrActiveQuoteRequest)
    
    'Raise an error if there is one.
    errCheckXMLResponse strAcceptedOrActiveQuoteResponse, True
    'Extract the MortgageSubQuoteNumber
    Set xmlOutDoc = xmlLoad(strAcceptedOrActiveQuoteResponse, strFunctionName)
    Set xmlMSQElem = xmlOutDoc.documentElement.selectSingleNode(".//MORTGAGESUBQUOTENUMBER")
    If Not (xmlMSQElem Is Nothing) Then
        strMSQNumber = xmlMSQElem.Text
        If Len(strMSQNumber) = 0 Then
            'Raise error - MORTGAGESUBQUOTENUMBER is present but has no value
            errThrowError strFunctionName, oeXMLMissingElementText, "MORTGAGESUBQUOTENUMBER is present but has no value"
        
        End If
        
    Else
        'Raise error - Missing MORTGAGESUBQUOTENUMBER
        errThrowError strFunctionName, oeXMLMissingElement, "MORTGAGESUBQUOTENUMBER element is absent"
    End If
    '-----------------------------------------------------------------------------------------------------------
    'STEP 4 Get MSQNumber from RESPONSE from STEP 3
    '-----------------------------------------------------------------------------------------------------------
    'strMSQNumber has now been obtained from above RESPONSE
    
    
    'STEP 5 Call MortgageSubQuoteBO.GetAllLoanComponentData(AppNum, AppFFNum, MSQNumber)
    '-----------------------------------------------------------------------------------------------------------
    '<REQUEST>
    '   <MORTGAGESUBQUOTE>
    '      <APPLICATIONNUMBER>S00004162</APPLICATIONNUMBER>
    '      <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
    '      <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
    '   </MORTGAGESUBQUOTE>
    '</REQUEST>
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    Dim strLoanComponentRequest As String
    Dim strLoanComponentResponse As String
    Dim xmlLoanComponentListElem As IXMLDOMElement
    Dim xmlMortgageProductList As IXMLDOMNodeList
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    Dim strTempResponse As String
    Dim xmlProductConditionsListResp As IXMLDOMElement
    Dim xmlProductConditionsList As IXMLDOMNodeList
    Dim xmlNonProductConditionsList As IXMLDOMNodeList
    Dim xmlErrorNumber As IXMLDOMElement
    Dim blnRecordNotFoundProd As Boolean
    Dim objMortgageProductBO As MortgageProductBO
    Dim iIndex As Integer
    Dim xmlConditionsListElem As IXMLDOMElement
    
    If gobjContext Is Nothing Then
        Set objMortgageSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objMortgageSubQuoteBO = _
            gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If

    'objMortgageSubQuoteBO.GetLoanComponentData OR (New Method) objMortgageSubQuoteBO.GetALLLoanComponentData ????
    strLoanComponentRequest = "<REQUEST>"
    strLoanComponentRequest = strLoanComponentRequest & "<MORTGAGESUBQUOTE>"
    strLoanComponentRequest = strLoanComponentRequest & xmlApplicationNumberElem.xml
    strLoanComponentRequest = strLoanComponentRequest & xmlApplicationFFNumberElem.xml
    strLoanComponentRequest = strLoanComponentRequest & xmlMSQElem.xml
    strLoanComponentRequest = strLoanComponentRequest & "</MORTGAGESUBQUOTE>"
    strLoanComponentRequest = strLoanComponentRequest & "</REQUEST>"
    
    'strLoanComponentResponse = objMortgageSubQuoteBO.GetAllLoanComponentData(strLoanComponentRequest)
    strLoanComponentResponse = objMortgageSubQuoteBO.FindLoanComponentList(strLoanComponentRequest)
    'Raise an error if there is one.
    errCheckXMLResponse strLoanComponentResponse, True
    Set xmlOutDoc = xmlLoad(strLoanComponentResponse, strFunctionName)
    Set xmlLoanComponentListElem = xmlOutDoc.documentElement.selectSingleNode(".//LOANCOMPONENTLIST")
    xmlChangeNodeName xmlLoanComponentListElem, "LOANCOMPONENTLIST", "MORTGAGEPRODUCTLIST"
    Set xmlMortgageProductList = xmlLoanComponentListElem.selectNodes(".//LOANCOMPONENT")
    'Loop through each node in the list and change the name from LOANCOMPONENT to MORTGAGEPRODUCT
    For iIndex = 0 To (xmlMortgageProductList.length - 1)
        Set xmlTempElem = xmlMortgageProductList.Item(iIndex)
        xmlChangeNodeName xmlTempElem, "LOANCOMPONENT", "MORTGAGEPRODUCT"
    Next iIndex
    
    'STEP 6 Call MortgageproductBO.FindProductConditionsList with MORTGAGEPRODUCTCODEs and STARTDATEs from above
    '-----------------------------------------------------------------------------------------------------------
    '<REQUEST OPERATION="FindProductConditionsList">
    '   <MORTGAGEPRODUCTLIST>
    '      <MORTGAGEPRODUCT>
    '         <MORTGAGEPRODUCTCODE>swp2</MORTGAGEPRODUCTCODE>
    '         <STARTDATE>01/01/2001</STARTDATE>
    '      </MORTGAGEPRODUCT>
    '      <MORTGAGEPRODUCT>
    '         <MORTGAGEPRODUCTCODE>12345</MORTGAGEPRODUCTCODE>
    '         <STARTDATE>13/06/2002 09:00:00</STARTDATE>
    '      </MORTGAGEPRODUCT>
    '      <MORTGAGEPRODUCT>
    '         <MORTGAGEPRODUCTCODE>somecode</MORTGAGEPRODUCTCODE>
    '         <STARTDATE>09/09/2001</STARTDATE>
    '      </MORTGAGEPRODUCT>
    '   </MORTGAGEPRODUCTLIST>
    '</REQUEST>
    If gobjContext Is Nothing Then
        Set objMortgageProductBO = New MortgageProductBO
    Else
        Set objMortgageProductBO = _
            gobjContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
    Set xmlTempElem = xmlGetRequestNode(vxmlRequestParentNode)
    xmlTempElem.appendChild xmlLoanComponentListElem
    
    strTempResponse = objMortgageProductBO.FindProductConditionsList(xmlTempElem.xml)
    'Raise an error if there is one.
    'Record Not Found is OK.
    'errCheckXMLResponse strTempResponse, False
    
    
    'STEP 7 set xmlProductConditionsListResp = RESPONSE from above call and convert to attribute based xml
    '-----------------------------------------------------------------------------------------------------------
    Set xmlOutDoc = xmlLoad(strTempResponse, strFunctionName)
    'Check for record not found
    Set xmlErrorNumber = xmlOutDoc.documentElement.selectSingleNode("ERROR/NUMBER")
    blnRecordNotFoundProd = False
    If Not (xmlErrorNumber Is Nothing) Then
        
        lngErrorNum = errGetOmigaErrorNumber(CSafeLng(xmlErrorNumber.Text))
        If lngErrorNum <> oeRecordNotFound Then
            errThrowError strFunctionName, lngErrorNum, strErrorDesc
        Else
            blnRecordNotFoundProd = True
        End If
    End If
    '4 cases can potentially exist now
    'CASE 1 Both calls find records (Non and Prod)
    'CASE 2 Non finds records, Prod doesn't
    'CASE 3 Neither calls find records (Non and Prod)
    'CASE 4 Non doesn't find any records, Prod does
    '
    'CASE 3
     If blnRecordNotFoundNon And blnRecordNotFoundProd Then
        'raise record not found error and exit method
        errThrowError strFunctionName, oeRecordNotFound
    End If
    Set xmlNonProductConditionsList = xmlNonProductConditionsListResp.selectNodes(".//CONDITIONS")
    Set xmlProductConditionsListResp = xmlOutDoc.documentElement.selectSingleNode(".//CONDITIONSLIST")
    Set xmlProductConditionsListResp = xmlCreateAttributeBasedResponse(xmlProductConditionsListResp, True)
    'STEP 8 Concatenate xmlNonProductConditionsList(STEP 2) and xmlProductConditionsListResp(STEP7) and Return
    '-----------------------------------------------------------------------------------------------------------
    Set xmlConditionsListElem = xmlOutDoc.createElement("CONDITIONSLIST")
    
    If (blnRecordNotFoundNon = True And blnRecordNotFoundProd = False) Or (blnRecordNotFoundNon = False And blnRecordNotFoundProd = False) Then
        For iIndex = 0 To (xmlNonProductConditionsList.length - 1)
            Set xmlTempElem = xmlNonProductConditionsList.Item(iIndex)
            xmlProductConditionsListResp.appendChild xmlTempElem
        Next iIndex
    
        'Append to RESPONSE Node
        vxmlResponseNode.appendChild xmlProductConditionsListResp
    Else
        For iIndex = 0 To (xmlNonProductConditionsList.length - 1)
            Set xmlTempElem = xmlNonProductConditionsList.Item(iIndex)
            xmlConditionsListElem.appendChild xmlTempElem
        Next iIndex
    
        'Append to RESPONSE Node
        vxmlResponseNode.appendChild xmlConditionsListElem
    
    End If
FindAppProductConditionsListExit:
    
    'Set xmlOutDoc = FreeThreadedDOMDocument40
    Set xmlConditionsElem = Nothing
    Set xmlRequestClone = Nothing
    Set xmlNonProductConditionsListResp = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplicationNumberElem = Nothing
    Set xmlApplicationFFNumberElem = Nothing
    Set xmlMSQElem = Nothing
    Set objApplicationQuoteBO = Nothing
    Set xmlAcceptedOrActiveQuoteResponse = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set xmlLoanComponentListElem = Nothing
    Set xmlMortgageProductList = Nothing
    Set xmlTempElem = Nothing
    Set xmlTempRequest = Nothing
    Set xmlProductConditionsListResp = Nothing
    Set xmlProductConditionsList = Nothing
    Set xmlNonProductConditionsList = Nothing
    Set xmlErrorNumber = Nothing
    Set objMortgageProductBO = Nothing
    Set xmlConditionsListElem = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub



'BMIDS00650
Private Sub IConditionsBO_GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                               ByVal vxmlParentNode As IXMLDOMNode)

    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "IConditionsBO_GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    'GD BMIDS00650
    'For Combos
    Dim xmlComboElement As IXMLDOMElement
    Dim xmlComboAttribute As IXMLDOMAttribute
    'For Global Parameters
    Dim xmlGlobalParameterElement As IXMLDOMElement
    Dim xmlGlobalParameterAttribute As IXMLDOMAttribute
    Dim strMaxAge As String
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    'GD BMIDS00650 - change combolookup to YES
'''    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
'''    xmlAttrib.Text = "YES"
'''    xmlNode.Attributes.setNamedItem xmlAttrib
    
    'GD BMIDS00650
'''    'Get Any Appropriate COMBOs
'''    '1 .. SpecialSchemes
'''    Set xmlComboElement = xmlRequestDoc.createElement("COMBO")
'''    Set xmlComboAttribute = xmlRequestDoc.createAttribute("_SCHEMA_")
'''    xmlComboAttribute.Text = "COMBO"
'''    xmlComboElement.Attributes.setNamedItem xmlComboAttribute
'''    Set xmlComboAttribute = xmlRequestDoc.createAttribute("GROUPNAME")
'''    xmlComboAttribute.Text = "SpecialSchemes"
'''    xmlComboElement.Attributes.setNamedItem xmlComboAttribute
'''    'Now xmlComboElement = <COMBO _SCHEMA_="COMBO" GROUPNAME="SpecialSchemes"/>
    
'''    'Get Any Appropriate GLOBALPARAMETERs
'''    '1 .. MaximumAge
'''    Set xmlGlobalParameterElement = xmlRequestDoc.createElement("GLOBALPARAMETER")
'''    Set xmlGlobalParameterAttribute = xmlRequestDoc.createAttribute("_SCHEMA_")
'''    xmlGlobalParameterAttribute.Text = "GLOBALPARAMETER"
'''    xmlGlobalParameterElement.Attributes.setNamedItem xmlGlobalParameterAttribute
'''    Set xmlGlobalParameterAttribute = xmlRequestDoc.createAttribute("NAME")
'''    xmlGlobalParameterAttribute.Text = "MaximumAge"
'''    xmlGlobalParameterElement.Attributes.setNamedItem xmlGlobalParameterAttribute
'''    'Now xmlGlobalParameterElement = <COMBO _SCHEMA_="GLOBALPARAMETER" NAME="MaximumAge"/>
'''    'GetGlobalParam
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Text = "APCondRBTemplate"
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
    'GD BMIDS00650
'''    xmlNode.appendChild xmlComboElement
'''    xmlNode.appendChild xmlGlobalParameterElement
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    Dim strTemp As String
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    strTemp = objRB.OmDataRequest(xmlRequestDoc.xml)
    xmlRequestDoc.loadXML strTemp
    Set objRB = Nothing
''    'GD DEBUG START BMIDS00876 removed debugging code
''            Dim hFile As Integer
''            Dim sFileName As String
''            Dim vstrXML As String
''
''            'IN vstrXML
''            vstrXML = strTemp
''            'IN sFilename
''            sFileName = "C:\Projects\BMIDS\MiddleTierCode\VBSs\omAppProc\GetApplicationData_output1.xml"
''            hFile = FreeFile
''
''            Kill sFileName
''
''            Open sFileName For Output Access Write As hFile
''
''            Print #hFile, vstrXML
''            Close hFile
''
''    'GD DEBUG END
    
    
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
    ' START GD BMIDS00650
    'Add all COMBO nodes underneath the APPLICATION NODE
    Dim xmlComboNodeList As IXMLDOMNodeList
    Dim iComboCount As Integer
    Dim iIndex As Integer
    
    Set xmlComboNodeList = xmlRequestDoc.selectNodes(".//RESPONSE/COMBO")
    iComboCount = xmlComboNodeList.length
    
    Set xmlElem = xmlRequestDoc.selectSingleNode(".//APPLICATION")
    For iIndex = 0 To (iComboCount - 1)
        Set xmlComboElement = xmlComboNodeList.Item(iIndex)
        xmlElem.appendChild xmlComboElement
    Next
    'Add all GLOBALPARAMETERS nodes underneath the APPLICATION NODE
    Dim xmlGlobalParameterNodeList As IXMLDOMNodeList
    Dim iGlobalParameterCount As Integer

    
    Set xmlGlobalParameterNodeList = xmlRequestDoc.selectNodes(".//RESPONSE/GLOBALPARAMETER")
    iGlobalParameterCount = xmlGlobalParameterNodeList.length
    
    Set xmlElem = xmlRequestDoc.selectSingleNode(".//APPLICATION")
    For iIndex = 0 To (iGlobalParameterCount - 1)
        Set xmlGlobalParameterElement = xmlGlobalParameterNodeList.Item(iIndex)
        xmlElem.appendChild xmlGlobalParameterElement
    Next
    
    ' END
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next

GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

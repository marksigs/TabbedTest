VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omAppProcBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omAppProcBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Application Processing Business Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    05/03/01    SYS1879 Added Application Summary
'DJP/GD 12/03/01    Valuer Instructions (GD) and Valuation Report (DP)
'JR     26/03/01    SYS2048 Added AdditionalQuestions
'PSC    30/03/01    SYS2188 Added processing for batch payments
'SA     18/05/01    SYS2256 Crested new method CreateReviewHistory & plugged in code for validation rules.
'CL     13/05/02    SYS4565 Modifications to method 'DoRequest' and addition of method 'DoCustomerRequest'
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'MV     28/05/2002  BMIDS00013 - Amended DoRequest() and Added a New function FindAdditionalQuestionsList()
'MV     14/06/2002  BMIDS00013 - Removed FindAdditionalQuestionsList()
'                               and Amended GetAdditionalQuestionsList() in DoRequest
'DPF    22/08/2002  BMIDS00344 - APWP3: have added call to new method -
'                                CreateValuationReportUsingTemplate
'MV     03/09/2002  BMIDS00353  Added New Method  GETLATESTVALUATIONAMOUNTS to DoRequest()
'GD     09/09/02    BMIDS00313 - APWP2 - BM004 Added FindNonProductConditionsList, FindAppProductConditionsList to DoRequest()
'
'------------------------------------------------------------------------------------------------
'INGDUK
'JD     28/10/2005  MAR342       Added new method CreateValuationReportNoInst
'PSC    25/03/2006  MARS1485     Apply BBG1939 Performance fixes
'------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function OmAppProcRequest(ByVal vstrXMLRequest As String) As String
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
    
    Const strFunctionName As String = "OmAppProcRequest"
    
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
    
    ' create REQUEST & RESPONSE FreeThreadedDOMDocuments
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    Set xmlInDoc = xmlLoad(vstrXMLRequest, strFunctionName)
    
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

    'if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
           
    'if we have created a Valuer Instruction, pass back the sequence no of the record created
    If xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateValuerInstructions" Then
        Dim intSeqNo As Integer
        intSeqNo = xmlRequestNode.firstChild.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO").Text
        xmlResponseElem.setAttribute "INSTRUCTIONSEQUENCENO", intSeqNo
    End If
        
    OmAppProcRequest = xmlOutDoc.xml
    
OmRequestExit:
    
    ' clean up ================================================================
    
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
    OmAppProcRequest = errCreateErrorResponse()
    Debug.Print OmAppProcRequest
    
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
    Dim xmlThisElem As IXMLDOMElement   'SA SYS2256
        
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    'PSC 25/03/2006 MARS1485 - Start
    Select Case strOperation
    
        Case "CREATEEMPLOYERSREFERENCE"
            CreateEmployersReference vxmlRequestNode, vxmlResponseNode
            
        Case "CREATECURREMPLOYERSREF"
            CreateCurrentEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "CREATEPREVEMPLOYERSREF"
            CreatePreviousEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "CREATEACCOUNTANTSREF"
            CreateAccountantsReference vxmlRequestNode, vxmlResponseNode

        Case "CREATELENDERSREFERENCE"
            CreateLendersReference vxmlRequestNode, vxmlResponseNode

        Case "CREATECURRLENDERSREF"
            CreateCurrentLendersReference vxmlRequestNode, vxmlResponseNode

        Case "CREATEPREVLENDERSREF"
            CreatePreviousLendersReference vxmlRequestNode, vxmlResponseNode

        Case "CREATELANDLORDSREF"
            CreateLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "CREATECURRLANDLORDSREF"
            CreateCurrentLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "CREATEPREVLANDLORDSREF"
            CreatePreviousLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "GETEMPLOYERSREFERENCE"
            GetEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "GETCURREMPLOYERSREF"
            GetCurrentEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "GETPREVEMPLOYERSREF"
            GetPreviousEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "GETACCOUNTANTSREF"
            GetAccountantsReference vxmlRequestNode, vxmlResponseNode

        Case "GETLANDLORDSREF"
            GetLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "GETCURRLANDLORDSREF"
            GetCurrentLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "GETPREVLANDLORDSREF"
            GetPreviousLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "GETLENDERSREFERENCE"
            GetLendersReference vxmlRequestNode, vxmlResponseNode

        Case "GETCURRLENDERSREF"
            GetCurrentLendersReference vxmlRequestNode, vxmlResponseNode

        Case "GETPREVLENDERSREF"
            GetPreviousLendersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATEEMPLOYERSREFERENCE"
            UpdateEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATECURREMPLOYERSREF"
            UpdateCurrentEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATEPREVEMPLOYERSREF"
            UpdatePreviousEmployersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATEACCOUNTANTSREF"
            UpdateAccountantsReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATELANDLORDSREF"
            UpdateLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATECURRLANDLORDSREF"
            UpdateCurrentLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATEPREVLANDLORDSREF"
            UpdatePreviousLandlordsReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATELENDERSREFERENCE"
            UpdateLendersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATECURRLENDERSREF"
            UpdateCurrentLendersReference vxmlRequestNode, vxmlResponseNode

        Case "UPDATEPREVLENDERSREF"
            UpdatePreviousLendersReference vxmlRequestNode, vxmlResponseNode

        Case "GETEMPDETAILSFORREF"
            GetEmpDetailsForReference vxmlRequestNode, vxmlResponseNode

        Case "GETLOANDETAILSFORREF"
            GetLoanDetailsForReference vxmlRequestNode, vxmlResponseNode

        Case "GETTENANCYDETAILSFORREF"
            GetTenancyDetailsForReference vxmlRequestNode, vxmlResponseNode
            
        Case "CREATEAPPLICATIONREVIEWHISTORY"
            CreateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
        
        Case "FINDAPPLICATIONREVIEWHISTORYLIST"
            FindApplicationReviewHistoryList vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEAPPLICATIONREVIEWHISTORY"
            UpdateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
        
        Case "VALIDATEUSERLOGON"
            ValidateUserLogon vxmlRequestNode, vxmlResponseNode
        
        Case "GETLOANANDRISKSDATA"
            GetLoanAndRisksData vxmlRequestNode, vxmlResponseNode
  
        Case "DELETEAPPLICATIONCONDITIONS"
            DeleteApplicationConditions vxmlRequestNode, vxmlResponseNode
            
        Case "FINDAPPLICATIONCONDITIONSLIST"
            FindApplicationConditionsList vxmlRequestNode, vxmlResponseNode
        
        Case "FINDCONDITIONSLIST"
            FindConditionsList vxmlRequestNode, vxmlResponseNode
                 
        Case "UPDATEAPPLICATIONCONDITIONS"
            UpdateApplicationConditions vxmlRequestNode, vxmlResponseNode
            
        Case "DELETEALLAPPLICATIONCONDITIONS"
            DeleteAllApplicationConditions vxmlRequestNode, vxmlResponseNode
            
         Case "CREATEAPPLICATIONCONDITIONS"
            CreateApplicationConditions vxmlRequestNode, vxmlResponseNode
         
         Case "CREATEDEFAULTAPPLICATIONCONDITIONS"
            CreateDefaultApplicationConditions vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "CREATEVALUATIONREPORT"
            CreateValuationReport vxmlRequestNode, vxmlResponseNode
        
        'APWP3 - DPF on the 22 Aug 2002
        Case "CREATEVALUATIONREPORTUSINGTEMPLATE"
            CreateValuationReportUsingTemplate vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "GETVALUATIONREPORT"
            GetValuationReport vxmlRequestNode, vxmlResponseNode
        
        'APWP3 - DPF on the 27 Aug 2002
        Case "GETVALUATIONREPORTFORTEMPLATE"
            GetValuationReportForTemplate vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "UPDATEVALUATIONREPORT"
            UpdateValuationReport vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "VALIDATEVALUATIONREPORT"
            ValidateValuationReport vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "VALIDATEINVOICEAMOUNT"
            ValidateInvoiceAmount vxmlRequestNode, vxmlResponseNode
        
        'SYS1839 (DP)
        Case "GETVALUATIONPROPERTYDETAILS"
            GetValuationPropertyDetails vxmlRequestNode, vxmlResponseNode
        
        'SYS2039 (GD)
        Case "GETVALUERINSTRUCTIONS"
            GetValuerInstructions vxmlRequestNode, vxmlResponseNode
        
        'SYS2039 (GD)
        Case "UPDATEVALUERINSTRUCTIONS"
            UpdateValuerInstructions vxmlRequestNode, vxmlResponseNode
        
        'SYS2039 (GD)
        Case "CREATEVALUERINSTRUCTIONS"
            CreateValuerInstructions vxmlRequestNode, vxmlResponseNode
            
        'SYS2039 (GD)
        Case "FINDVALUERINSTRUCTIONLIST"
            FindValuerInstructionList vxmlRequestNode, vxmlResponseNode
            
        'SYS2039 (GD)
        Case "GETVALUATIONFEE"
            GetValuationFee vxmlRequestNode, vxmlResponseNode
         
        'SYS2039 (GD)
        Case "GETVALUATIONPROPERTYDETAILS"
            GetValuationPropertyDetails vxmlRequestNode, vxmlResponseNode
        
        'SYS2048 (JR)
        Case "FINDADDITIONALQUESTIONLIST"
            FindAdditionalQuestionsList vxmlRequestNode, vxmlResponseNode
        
        'SYS2048 (JR)
        Case "FINDAPPADDITIONALQUESTIONLIST"
            FindAppAdditionalQuestionsList vxmlRequestNode, vxmlResponseNode
        
        'SYS2048 (JR)
        Case "CREATEAPPADDITIONALQUESTIONS"
            CreateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
        
        'SYS2048 (JR)
        Case "UPDATEAPPADDITIONALQUESTIONS"
            UpdateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
        
        ' PSC 30/03/01 SYS2188 - Start
        Case "FINDINSTRUCTIONSFORBATCH"
            FindInstructionsForBatch vxmlRequestNode, vxmlResponseNode
            
        Case "PROCESSVALUERBATCH"
            ProcessValuerBatch vxmlRequestNode, vxmlResponseNode
        
        Case "CREATEVALUATIONAUDIT"
            CreateValuationAudit vxmlRequestNode, vxmlResponseNode
        
        Case "FINDFAILEDINSTRUCTIONS"
            FindFailedInstructions vxmlRequestNode, vxmlResponseNode
        
        Case "PROCESSVALUATIONINSTRUCTION"
            ProcessValuationInstruction vxmlRequestNode, vxmlResponseNode
        ' PSC 30/03/01 SYS2188 - End
        'SYS2256 (SA)
        Case "VALIDATECURREMPLOYERSREF"
            ValidateCurrentEmployersReference vxmlRequestNode, vxmlResponseNode
        'SYS2256 (SA)
        Case "VALIDATEPREVEMPLOYERSREF"
            ValidatePreviousEmployersReference vxmlRequestNode, vxmlResponseNode
        'SYS2256 (SA)
        Case "VALIDATECURRLENDERSREF"
            ValidateCurrentLendersReference vxmlRequestNode, vxmlResponseNode
        'SYS2256 (SA)
        Case "VALIDATEPREVLENDERSREF"
            ValidatePreviousLendersReference vxmlRequestNode, vxmlResponseNode
        'SYS2256 (SA)
        Case "VALIDATECURRLANDLORDSREF"
            ValidateCurrentLandlordsReference vxmlRequestNode, vxmlResponseNode
        'SYS2256 (SA)
        Case "VALIDATEPREVLANDLORDSREF"
            ValidatePreviousLandlordsReference vxmlRequestNode, vxmlResponseNode
        
        'SYS2256 (SA)
        Case "VALIDATEACCOUNTANTSREF"
            ValidateAccountantsReference vxmlRequestNode, vxmlResponseNode
        
        'BMIDS - DCWP2- BM071 - BMIDS00013 - MV - 14/06/2002
        Case "GETADDITIONALQUESTIONSLIST"
            GetAdditionalQuestionsList vxmlRequestNode, vxmlResponseNode
    
        'BMIDS - CMWP4 - BM087 - BMIDS00353 - MV - 02/09/2002
        Case "GETLATESTVALUATIONAMOUNTS"
            GetLatestValuationAmounts vxmlRequestNode, vxmlResponseNode
            
        'JD MAR40
        Case "CREATEVALUATIONREPORTNOINST"
            CreateValuationReportNoInstruction vxmlRequestNode, vxmlResponseNode
            
        'BMIDS00313 - APWP2 - BM004 Added FindNonProductConditionsList,FindAppProductConditionsList

         Case "FINDNONPRODUCTCONDITIONSLIST"
            FindNonProductConditionsList vxmlRequestNode, vxmlResponseNode
            
         Case "FINDAPPPRODUCTCONDITIONSLIST"
            FindAppProductConditionsList vxmlRequestNode, vxmlResponseNode
            
        Case Else
            errThrowError _
                "AppProcRequest", _
                oeNotImplemented, _
                vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    End Select
    'PSC 25/03/2006 MARS1485 - End
    
DoRequestExit:
    
    Set xmlThisElem = Nothing   'SA SYS2256
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateReviewHistory(ByVal vstrNodeToSelect As String, _
                                ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode, _
                                ByVal strReviewReason As String)
' header ----------------------------------------------------------------------------------
' description:
'
' pass:
'   vstrNodeToSelect
'       Node to Select
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
'   strReviewReason
'       Application Review Reason
' return:
'       none
' exceptions:
'
'------------------------------------------------------------------------------------------
    On Error GoTo CreateReviewHistoryExit
    
    Const strFunctionName As String = "CreateReviewHistory"
    Dim strSearch As String, _
        strComments As String
        
    Dim xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement, _
        xmlRequestDoc As FreeThreadedDOMDocument40, _
        xmlResponseDoc As FreeThreadedDOMDocument40, _
        xmlAttrib As IXMLDOMAttribute
        
        
    Dim iUnderReviewInd As Integer
    
    strSearch = vstrNodeToSelect
    strSearch = strSearch & "[@RESULT='0']"

    Set xmlResponseList = vxmlResponseNode.selectNodes(strSearch)
    If xmlResponseList.length > 0 Then

        iUnderReviewInd = 1 ' 1 or more rules have failed

        For Each xmlThisNode In xmlResponseList

            If Len(strComments) > 0 Then
                strComments = strComments & ";"
            End If

            strComments = strComments & xmlGetAttributeText(xmlThisNode, "MESSAGE")

        Next

        ' Build CreateApplicationReviewHistory request
        ' Strip child nodes from request
        For Each xmlThisNode In vxmlRequestNode.childNodes
            vxmlRequestNode.removeChild xmlThisNode
        Next

        ' Set operation value
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
        xmlRequestDoc.validateOnParse = False
        xmlRequestDoc.setProperty "NewParser", True
        xmlRequestDoc.async = False
        Set xmlThisRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateApplicationReviewHistory"
        
        Set xmlThisElem = xmlRequestDoc.createElement("APPLICATIONREVIEWHISTORY")
        
        Set xmlApplicationNode = vxmlResponseNode.selectSingleNode("CASETASK")
        xmlThisElem.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlApplicationNode, "CASEID")
        xmlThisElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
        
        xmlThisElem.setAttribute "REVIEWREASON", strReviewReason
        xmlSetSysDateToNodeAttrib xmlThisElem, "REVIEWDATETIME"
        xmlThisElem.setAttribute _
            "REVIEWUSERID", xmlGetAttributeText(xmlApplicationNode, "USERID")
        xmlThisElem.setAttribute _
            "REVIEWUNITID", xmlGetAttributeText(xmlApplicationNode, "UNITID")
        xmlThisElem.setAttribute "REVIEWCOMMENTS", strComments
       
       ' Complete the XML request by appending
        xmlThisRequestNode.appendChild xmlThisElem

        Debug.Print "omAppProc.omAppProcBO Request"
        Debug.Print xmlRequestDoc.xml

        ' Now call to create application history
        Set xmlResponseDoc = New FreeThreadedDOMDocument40
        xmlResponseDoc.validateOnParse = False
        xmlResponseDoc.setProperty "NewParser", True
        xmlResponseDoc.loadXML OmAppProcRequest(xmlRequestDoc.xml)
        
        Debug.Print "omAppProc.omAppProcBO Response"
        Debug.Print xmlResponseDoc.xml

        errCheckXMLResponseNode _
            xmlResponseDoc.selectSingleNode("RESPONSE"), _
            vxmlResponseNode, True

        'Set response node to go back to the GUI
        'Strip child nodes from request
    
        For Each xmlThisNode In vxmlResponseNode.childNodes
            vxmlResponseNode.removeChild xmlThisNode
        Next

        Set xmlThisElem = xmlResponseDoc.createElement("APPSTATUS")

        Set xmlAttrib = xmlResponseDoc.createAttribute("UNDERREVIEWIND")
        xmlAttrib.Text = Trim$(Str(iUnderReviewInd))
        xmlThisElem.Attributes.setNamedItem xmlAttrib
        vxmlResponseNode.appendChild xmlThisElem
    Else
        ' No rules have failed - set up Appstatus
        'Set response node to go back to the GUI
        'Strip child nodes from response
        For Each xmlThisNode In vxmlResponseNode.childNodes
            vxmlResponseNode.removeChild xmlThisNode
        Next
        Set xmlResponseDoc = New FreeThreadedDOMDocument40
        xmlResponseDoc.validateOnParse = False
        xmlResponseDoc.setProperty "NewParser", True
        Set xmlThisElem = xmlResponseDoc.createElement("APPSTATUS")
        Set xmlAttrib = xmlResponseDoc.createAttribute("UNDERREVIEWIND")
        xmlAttrib.Text = "0"
        xmlThisElem.Attributes.setNamedItem xmlAttrib
        vxmlResponseNode.appendChild xmlThisElem
    End If
    
CreateReviewHistoryExit:
    
    Set xmlResponseList = Nothing
    Set xmlThisNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

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

'PSC 25/03/2006 MARS1485 - Start
Private Sub CreateEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                     ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateEmployersReferenceExit
    Const strFunctionName As String = "CreateEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateEmployersReference vxmlRequestNode, vxmlResponseNode
    
CreateEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateCurrentEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateCurrentEmployersReferenceExit
    Const strFunctionName As String = "CreateCurrentEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrentEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreatePreviousEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreatePreviousEmployersReferenceExit
    Const strFunctionName As String = "CreatePreviousEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreatePrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
CreatePreviousEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateAccountantsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateAccountantsReferenceExit
    Const strFunctionName As String = "CreateAccountantsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateAccountantsRef vxmlRequestNode, vxmlResponseNode
    
CreateAccountantsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
                                           ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateApplicationReviewHistoryExit
    Const strFunctionName As String = "CreateApplicationReviewHistory"
    
    Dim objAppSummaryBO As IAppSummaryBO
    
    If Not gobjContext Is Nothing Then
        Set objAppSummaryBO = gobjContext.CreateInstance("omAppProc.AppSummaryBO")
    Else
        Set objAppSummaryBO = New AppSummaryBO
    End If
    
    objAppSummaryBO.CreateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
    
CreateApplicationReviewHistoryExit:
    
    Set objAppSummaryBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindApplicationReviewHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                             ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo FindApplicationReviewHistoryListExit
    Const strFunctionName As String = "FindApplicationReviewHistoryList"
    
    Dim objAppSummaryBO As IAppSummaryBO
    
    If Not gobjContext Is Nothing Then
        Set objAppSummaryBO = gobjContext.CreateInstance("omAppProc.AppSummaryBO")
    Else
        Set objAppSummaryBO = New AppSummaryBO
    End If
    
    objAppSummaryBO.FindApplicationReviewHistoryList vxmlRequestNode, vxmlResponseNode
    
FindApplicationReviewHistoryListExit:
    
    Set objAppSummaryBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub DeleteApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo DeleteApplicationConditionsExit
    Const strFunctionName As String = "DeleteApplicationConditions"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.DeleteApplicationConditions vxmlRequestNode, vxmlResponseNode
    
DeleteApplicationConditionsExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindApplicationConditionsList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindApplicationConditionsListExit
    Const strFunctionName As String = "FindApplicationConditionsList"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.FindApplicationConditionsList vxmlRequestNode, vxmlResponseNode
    
FindApplicationConditionsListExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindAppProductConditionsList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindAppProductConditionsListExit
    Const strFunctionName As String = "FindAppProductConditionsList"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.FindAppProductConditionsList vxmlRequestNode, vxmlResponseNode
    
FindAppProductConditionsListExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindNonProductConditionsList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindNonProductConditionsListExit
    Const strFunctionName As String = "FindNonProductConditionsList"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.FindNonProductConditionsList vxmlRequestNode, vxmlResponseNode
    
FindNonProductConditionsListExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub FindConditionsList(ByVal vxmlRequestNode As IXMLDOMNode, _
                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindConditionsListExit
    Const strFunctionName As String = "FindConditionsList"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.FindConditionsList vxmlRequestNode, vxmlResponseNode
    
FindConditionsListExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo UpdateApplicationConditionsExit
    Const strFunctionName As String = "UpdateApplicationConditions"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.UpdateApplicationConditions vxmlRequestNode, vxmlResponseNode
    
UpdateApplicationConditionsExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo CreateApplicationConditionsExit
    Const strFunctionName As String = "CreateApplicationConditions"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.CreateApplicationConditions vxmlRequestNode, vxmlResponseNode
    
CreateApplicationConditionsExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateDefaultApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo CreateDefaultApplicationConditionsExit
    Const strFunctionName As String = "CreateDefaultApplicationConditions"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.CreateDefaultApplicationConditions vxmlRequestNode, vxmlResponseNode
    
CreateDefaultApplicationConditionsExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo CreateValuationReportExit
    Const strFunctionName As String = "CreateValuationReport"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.CreateValuationReport vxmlRequestNode, vxmlResponseNode
    
CreateValuationReportExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateValuationReportUsingTemplate(ByVal vxmlRequestNode As IXMLDOMNode, _
                                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo CreateValuationReportUsingTemplateExit
    Const strFunctionName As String = "CreateValuationReportUsingTemplate"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.CreateValuationReportUsingTemplate vxmlRequestNode, vxmlResponseNode
    
CreateValuationReportUsingTemplateExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo UpdateValuationReportExit
    Const strFunctionName As String = "UpdateValuationReport"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.UpdateValuationReport vxmlRequestNode, vxmlResponseNode
    
UpdateValuationReportExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo GetValuationReportExit
    Const strFunctionName As String = "GetValuationReport"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetValuationReport vxmlRequestNode, vxmlResponseNode
    
GetValuationReportExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetValuationReportForTemplate(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo GetValuationReportForTemplateExit
    Const strFunctionName As String = "GetValuationReportForTemplate"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetValuationReportForTemplate vxmlRequestNode, vxmlResponseNode
    
GetValuationReportForTemplateExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub ValidateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo ValidateValuationReportExit
    Const strFunctionName As String = "ValidateValuationReport"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.ValidateValuationReport vxmlRequestNode, vxmlResponseNode
    
ValidateValuationReportExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub FindInstructionsForBatch(ByVal vxmlRequestNode As IXMLDOMNode, _
                                     ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindInstructionsForBatchExit
    Const strFunctionName As String = "FindInstructionsForBatch"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.FindInstructionsForBatch vxmlRequestNode, vxmlResponseNode
    
FindInstructionsForBatchExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ProcessValuerBatch(ByVal vxmlRequestNode As IXMLDOMNode, _
                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo ProcessValuerBatchExit
    Const strFunctionName As String = "ProcessValuerBatch"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.ProcessValuerBatch vxmlRequestNode, vxmlResponseNode
    
ProcessValuerBatchExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateValuationAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
                                 ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo CreateValuationAuditExit
    Const strFunctionName As String = "CreateValuationAudit"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.CreateValuationAudit vxmlRequestNode, vxmlResponseNode
    
CreateValuationAuditExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub FindFailedInstructions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                   ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo FindFailedInstructionsExit
    Const strFunctionName As String = "FindFailedInstructions"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.FindFailedInstructions vxmlRequestNode, vxmlResponseNode
    
FindFailedInstructionsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ProcessValuationInstruction(ByVal vxmlRequestNode As IXMLDOMNode, _
                                         ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo ProcessValuationInstructionExit
    Const strFunctionName As String = "ProcessValuationInstruction"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.ProcessValuationInstruction vxmlRequestNode, vxmlResponseNode
    
ProcessValuationInstructionExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetLatestValuationAmounts(ByVal vxmlRequestNode As IXMLDOMNode, _
                                      ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo GetLatestValuationAmountsExit
    Const strFunctionName As String = "GetLatestValuationAmounts"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetLatestValuationAmounts vxmlRequestNode, vxmlResponseNode
    
GetLatestValuationAmountsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub ValidateInvoiceAmount(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo ValidateInvoiceAmountExit
    Const strFunctionName As String = "ValidateInvoiceAmount"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.ValidateInvoiceAmount vxmlRequestNode, vxmlResponseNode
    
ValidateInvoiceAmountExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetValuationPropertyDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo GetValuationPropertyDetailsExit
    Const strFunctionName As String = "GetValuationPropertyDetails"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetValuationPropertyDetails vxmlRequestNode, vxmlResponseNode
    
GetValuationPropertyDetailsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub GetValuerInstructions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
                                     
    On Error GoTo GetValuerInstructionsExit
    Const strFunctionName As String = "GetValuerInstructions"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetValuerInstructions vxmlRequestNode, vxmlResponseNode
    
GetValuerInstructionsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateValuerInstructions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateValuerInstructionsExit
    Const strFunctionName As String = "UpdateValuerInstructions"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.UpdateValuerInstructions vxmlRequestNode, vxmlResponseNode
    
UpdateValuerInstructionsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub FindValuerInstructionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo FindValuerInstructionListExit
    Const strFunctionName As String = "FindValuerInstructionList"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.FindValuerInstructionList vxmlRequestNode, vxmlResponseNode
    
FindValuerInstructionListExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetValuationFee(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetValuationFeeExit
    Const strFunctionName As String = "GetValuationFee"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.GetValuationFee vxmlRequestNode, vxmlResponseNode
    
GetValuationFeeExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub FindAdditionalQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo FindAdditionalQuestionsListExit
    Const strFunctionName As String = "FindAdditionalQuestionsList"
    
    Dim objAdditionalQuestionsBO As IAdditionalQuestionsBO
    
    If Not gobjContext Is Nothing Then
        Set objAdditionalQuestionsBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsBO")
    Else
        Set objAdditionalQuestionsBO = New AdditionalQuestionsBO
    End If
    
    objAdditionalQuestionsBO.FindAdditionalQuestionList vxmlRequestNode, vxmlResponseNode
    
FindAdditionalQuestionsListExit:
    
    Set objAdditionalQuestionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub FindAppAdditionalQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo FindAppAdditionalQuestionsListExit
    Const strFunctionName As String = "FindAppAdditionalQuestionsList"
    
    Dim objAdditionalQuestionsBO As IAdditionalQuestionsBO
    
    If Not gobjContext Is Nothing Then
        Set objAdditionalQuestionsBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsBO")
    Else
        Set objAdditionalQuestionsBO = New AdditionalQuestionsBO
    End If
    
    objAdditionalQuestionsBO.FindAppAdditionalQuestionList vxmlRequestNode, vxmlResponseNode
    
FindAppAdditionalQuestionsListExit:
    
    Set objAdditionalQuestionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateAppAdditionalQuestionsExit
    Const strFunctionName As String = "CreateAppAdditionalQuestions"
    
    Dim objAdditionalQuestionsBO As IAdditionalQuestionsBO
    
    If Not gobjContext Is Nothing Then
        Set objAdditionalQuestionsBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsBO")
    Else
        Set objAdditionalQuestionsBO = New AdditionalQuestionsBO
    End If
    
    objAdditionalQuestionsBO.CreateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
    
CreateAppAdditionalQuestionsExit:
    
    Set objAdditionalQuestionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateAppAdditionalQuestionsExit
    Const strFunctionName As String = "UpdateAppAdditionalQuestions"
    
    Dim objAdditionalQuestionsBO As IAdditionalQuestionsBO
    
    If Not gobjContext Is Nothing Then
        Set objAdditionalQuestionsBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsBO")
    Else
        Set objAdditionalQuestionsBO = New AdditionalQuestionsBO
    End If
    
    objAdditionalQuestionsBO.UpdateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
    
UpdateAppAdditionalQuestionsExit:
    
    Set objAdditionalQuestionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetAdditionalQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetAdditionalQuestionsListExit
    Const strFunctionName As String = "GetAdditionalQuestionsList"
    
    Dim objAdditionalQuestionsBO As IAdditionalQuestionsBO
    
    If Not gobjContext Is Nothing Then
        Set objAdditionalQuestionsBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsBO")
    Else
        Set objAdditionalQuestionsBO = New AdditionalQuestionsBO
    End If
    
    objAdditionalQuestionsBO.GetAdditionalQuestionsList vxmlRequestNode, vxmlResponseNode
    
GetAdditionalQuestionsListExit:
    
    Set objAdditionalQuestionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub CreateValuationReportNoInstruction(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateValuationReportNoInstructionExit
    Const strFunctionName As String = "CreateValuationReportNoInstruction"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.CreateValuationReportNoInst vxmlRequestNode, vxmlResponseNode
    
CreateValuationReportNoInstructionExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateValuerInstructions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateValuerInstructionsExit
    Const strFunctionName As String = "CreateValuerInstructions"
    
    Dim objValuationBO As IValuationBO
    
    If Not gobjContext Is Nothing Then
        Set objValuationBO = gobjContext.CreateInstance("omAppProc.ValuationBO")
    Else
        Set objValuationBO = New ValuationBO
    End If
    
    objValuationBO.CreateValuerInstructions vxmlRequestNode, vxmlResponseNode
    
CreateValuerInstructionsExit:
    
    Set objValuationBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ValidateUserLogon(ByVal vxmlRequestNode As IXMLDOMNode, _
                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidateUserLogonExit
    Const strFunctionName As String = "ValidateUserLogon"
    
    Dim objAppSummaryBO As IAppSummaryBO
    
    If Not gobjContext Is Nothing Then
        Set objAppSummaryBO = gobjContext.CreateInstance("omAppProc.AppSummaryBO")
    Else
        Set objAppSummaryBO = New AppSummaryBO
    End If
    
    objAppSummaryBO.ValidateUserLogon vxmlRequestNode, vxmlResponseNode
    
ValidateUserLogonExit:
    
    Set objAppSummaryBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetLoanAndRisksData(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetLoanAndRisksDataExit
    Const strFunctionName As String = "GetLoanAndRisksData"
    
    Dim objAppSummaryBO As IAppSummaryBO
    
    If Not gobjContext Is Nothing Then
        Set objAppSummaryBO = gobjContext.CreateInstance("omAppProc.AppSummaryBO")
    Else
        Set objAppSummaryBO = New AppSummaryBO
    End If
    
    objAppSummaryBO.GetLoanAndRisksData vxmlRequestNode, vxmlResponseNode
    
GetLoanAndRisksDataExit:
    
    Set objAppSummaryBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
                                           ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateApplicationReviewHistoryExit
    Const strFunctionName As String = "UpdateApplicationReviewHistory"
    
    Dim objAppSummaryBO As IAppSummaryBO
    
    If Not gobjContext Is Nothing Then
        Set objAppSummaryBO = gobjContext.CreateInstance("omAppProc.AppSummaryBO")
    Else
        Set objAppSummaryBO = New AppSummaryBO
    End If
    
    objAppSummaryBO.UpdateApplicationReviewHistory vxmlRequestNode, vxmlResponseNode
    
UpdateApplicationReviewHistoryExit:
    
    Set objAppSummaryBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateLendersReferenceExit
    Const strFunctionName As String = "CreateLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateLendersReference vxmlRequestNode, vxmlResponseNode
    
CreateLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CreateCurrentLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateCurrentLendersReferenceExit
    Const strFunctionName As String = "CreateCurrentLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrentLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreatePreviousLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreatePreviousLendersReferenceExit
    Const strFunctionName As String = "CreatePreviousLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreatePrevLendersRef vxmlRequestNode, vxmlResponseNode
    
CreatePreviousLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateLandlordsReferenceExit
    Const strFunctionName As String = "CreateLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateLandlordsRef vxmlRequestNode, vxmlResponseNode
    
CreateLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateCurrentLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreateCurrentLandlordsReferenceExit
    Const strFunctionName As String = "CreateCurrentLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreateCurrLandlordsRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrentLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdateEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateEmployersReferenceExit
    Const strFunctionName As String = "UpdateEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateEmployersReference vxmlRequestNode, vxmlResponseNode
    
UpdateEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateLendersReferenceExit
    Const strFunctionName As String = "UpdateLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateLendersReference vxmlRequestNode, vxmlResponseNode
    
UpdateLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdateCurrentEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateCurrentEmployersReferenceExit
    Const strFunctionName As String = "UpdateCurrentEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrentEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdatePreviousEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                             ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdatePreviousEmployersReferenceExit
    Const strFunctionName As String = "UpdatePreviousEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdatePrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
UpdatePreviousEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateAccountantsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateAccountantsReferenceExit
    Const strFunctionName As String = "UpdateAccountantsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateAccountantsRef vxmlRequestNode, vxmlResponseNode
    
UpdateAccountantsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdateLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateLandlordsReferenceExit
    Const strFunctionName As String = "UpdateLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateLandlordsRef vxmlRequestNode, vxmlResponseNode
    
UpdateLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateCurrentLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateCurrentLandlordsReferenceExit
    Const strFunctionName As String = "UpdateCurrentLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateCurrLandlordsRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrentLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub UpdateCurrentLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdateCurrentLendersReferenceExit
    Const strFunctionName As String = "UpdateCurrentLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdateCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrentLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdatePreviousLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdatePreviousLandlordsReferenceExit
    Const strFunctionName As String = "UpdatePreviousLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdatePrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
UpdatePreviousLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub UpdatePreviousLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo UpdatePreviousLendersReferenceExit
    Const strFunctionName As String = "UpdatePreviousLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.UpdatePrevLendersRef vxmlRequestNode, vxmlResponseNode
    
UpdatePreviousLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreatePreviousLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo CreatePreviousLandlordsReferenceExit
    Const strFunctionName As String = "CreatePreviousLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.CreatePrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
CreatePreviousLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetEmployersReferenceExit
    Const strFunctionName As String = "GetEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetEmployersReference vxmlRequestNode, vxmlResponseNode
    
GetEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetLendersReferenceExit
    Const strFunctionName As String = "GetLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetLendersReference vxmlRequestNode, vxmlResponseNode
    
GetLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetEmpDetailsForReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetEmpDetailsForReferenceExit
    Const strFunctionName As String = "GetEmpDetailsForReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetEmpDetailsForRef vxmlRequestNode, vxmlResponseNode
    
GetEmpDetailsForReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetLoanDetailsForReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetLoanDetailsForReferenceExit
    Const strFunctionName As String = "GetLoanDetailsForReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetLoanDetailsForRef vxmlRequestNode, vxmlResponseNode
    
GetLoanDetailsForReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetTenancyDetailsForReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetTenancyDetailsForReferenceExit
    Const strFunctionName As String = "GetTenancyDetailsForReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetTenancyDetailsForRef vxmlRequestNode, vxmlResponseNode
    
GetTenancyDetailsForReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetCurrentEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetCurrentEmployersReferenceExit
    Const strFunctionName As String = "GetCurrentEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
GetCurrentEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetPreviousEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetPreviousEmployersReferenceExit
    Const strFunctionName As String = "GetPreviousEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetPrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
GetPreviousEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetPreviousLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetPreviousLandlordsReferenceExit
    Const strFunctionName As String = "GetPreviousLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetPrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
GetPreviousLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetPreviousLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetPreviousLendersReferenceExit
    Const strFunctionName As String = "GetPreviousLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetPrevLendersRef vxmlRequestNode, vxmlResponseNode
    
GetPreviousLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetAccountantsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetAccountantsReferenceExit
    Const strFunctionName As String = "GetAccountantsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetAccountantsRef vxmlRequestNode, vxmlResponseNode
    
GetAccountantsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetLandlordsReferenceExit
    Const strFunctionName As String = "GetLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetLandlordsRef vxmlRequestNode, vxmlResponseNode
    
GetLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetCurrentLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetCurrentLandlordsReferenceExit
    Const strFunctionName As String = "GetCurrentLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetLandlordsRef vxmlRequestNode, vxmlResponseNode
    
GetCurrentLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetCurrentLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo GetCurrentLendersReferenceExit
    Const strFunctionName As String = "GetCurrentLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.GetCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
GetCurrentLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub ValidateCurrentEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidateCurrentEmployersReferenceExit
    Const strFunctionName As String = "ValidateCurrentEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidateCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "CURRENTEMPLOYERREFRULE", vxmlRequestNode, vxmlResponseNode, "30"
    
ValidateCurrentEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub ValidateCurrentLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidateCurrentLendersReferenceExit
    Const strFunctionName As String = "ValidateCurrentLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidateCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "CURRENTLENDERREFRULE", vxmlRequestNode, vxmlResponseNode, "50"
    
ValidateCurrentLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub ValidatePreviousEmployersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidatePreviousEmployersReferenceExit
    Const strFunctionName As String = "ValidatePreviousEmployersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidatePrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "PREVIOUSEMPLOYERREFRULE", vxmlRequestNode, vxmlResponseNode, "40"
    
ValidatePreviousEmployersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub ValidatePreviousLendersReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidatePreviousLendersReferenceExit
    Const strFunctionName As String = "ValidatePreviousLendersReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidatePrevLendersRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "PREVIOUSLENDERREFRULE", vxmlRequestNode, vxmlResponseNode, "60"
        
ValidatePreviousLendersReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ValidateCurrentLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidateCurrentLandlordsReferenceExit
    Const strFunctionName As String = "ValidateCurrentLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidateCurrLandlordsRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "CURRENTLANDLORDEFRULE", vxmlRequestNode, vxmlResponseNode, "70"
    
ValidateCurrentLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ValidatePreviousLandlordsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidatePreviousLandlordsReferenceExit
    Const strFunctionName As String = "ValidatePreviousLandlordsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If
    
    objRefsBO.ValidatePrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "PREVIOUSLANDLORDEFRULE", vxmlRequestNode, vxmlResponseNode, "80"
    
ValidatePreviousLandlordsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub ValidateAccountantsReference(ByVal vxmlRequestNode As IXMLDOMNode, _
                                               ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo ValidateAccountantsReferenceExit
    Const strFunctionName As String = "ValidateAccountantsReference"
    
    Dim objRefsBO As IReferencesBO
    
    If Not gobjContext Is Nothing Then
        Set objRefsBO = gobjContext.CreateInstance("omAppProc.referencesBO")
    Else
        Set objRefsBO = New referencesBO
    End If

    objRefsBO.ValidateAccountantsRef vxmlRequestNode, vxmlResponseNode
    
    CreateReviewHistory "ACCOUNTANTSREFRULE", vxmlRequestNode, vxmlResponseNode, "20"
    
ValidateAccountantsReferenceExit:
    
    Set objRefsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub DeleteAllApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                                     
    On Error GoTo DeleteApplicationConditionsExit
    Const strFunctionName As String = "DeleteAllApplicationConditions"
    
    Dim objConditionsBO As IConditionsBO
    
    If Not gobjContext Is Nothing Then
        Set objConditionsBO = gobjContext.CreateInstance("omAppProc.ConditionsBO")
    Else
        Set objConditionsBO = New ConditionsBO
    End If
    
    objConditionsBO.DeleteAllApplicationConditions vxmlRequestNode, vxmlResponseNode
    
DeleteApplicationConditionsExit:
    
    Set objConditionsBO = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub
'PSC 25/03/2006 MARS1485 - End


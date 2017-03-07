VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Postprocessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <PostProcessing>
' File:         <Postprocessing.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionRulesBO.doc
' Purpose:      To perform any client based post processing tasks for the ingestion engine
' History:
'   AQR                 AUTHOR          DESCRIPTION                         DATE
'   ---                 ------          -----------                         ------
'   IDWP1               Martin O'Hanlon Created                             17/07/2002
'   IDWP1 - BMIDS00462  Martin O'Hanlon Made changes for the marking        16/10/2002
'                                       payments made, and modified the
'                                       mortgage application stage
'   BMIDS00999          Martin O'Hanlon Changed to take stings from the     20/11/2002
'                                       core component
'   BMIDS01073          Martin O'Hanlon Removed CreateActivity              25/11/2002
'   BMIDS01097          Martin O'Hanlon Fixed bugs noted from BMIDS UAT     27/11/2002
'   BM0209              SR              Modifed method                      02/01/2003
'                                       'UpdateQuotationRecord'
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

'CONSTANTS

'TASK MANAGEMENT
' Global for Omiga
Const cstrSourceApplication As String = "Omiga" ' Omiga Application
Const cstrActivityID As String = "10" ' Mortgage Application
' Specific for Create Application
Const cstrCreateApplicationStageID As String = "30" ' Mortgage Application stage

Private gobjContext As ObjectContext

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'------------------------------------------------------------------------
' Procedure OmRequest
' Author: MO
' Date: 20/11/2002
' Purpose: post-processing public method, this will be called at the
'    post-processing phase by the DataIngestion component
' Input parameters: XML Request string
' Output parameters: XML Response string
'------------------------------------------------------------------------
Public Function OmRequest(ByVal vstrXmlIn As String) As String

On Error GoTo OmRequestErrorOccured
        
    Dim xmlRequest As IXMLDOMNode
    Dim xmlOmigaFormattedRequest As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    
    'extract the main nodes from the rules request
    ExtractPostprocessingRulesRequest vstrXmlIn, xmlRequest, xmlOmigaFormattedRequest, xmlResponse
    
    'do the request
    DoRequest xmlRequest, xmlOmigaFormattedRequest, xmlResponse
    
    'return the response xml string
    OmRequest = BuildPostprocessingRulesResponse(xmlRequest, xmlOmigaFormattedRequest, xmlResponse)
    
OmRequestExit:

    'destroy objects
    Set xmlRequest = Nothing
    Set xmlOmigaFormattedRequest = Nothing
    Set xmlResponse = Nothing
    
Exit Function

OmRequestErrorOccured:

    Err.Raise Err.Number, "DoRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume OmRequestExit

End Function

'------------------------------------------------------------------------
' Procedure DoRequest
' Author: MO
' Date: 19/07/2002
' Purpose: This is the method which performs the correct activity based
'           on the operation
' Input parameters: vxmlRequestNode, OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DoRequestErrorOccured

    Dim objDOM As FreeThreadedDOMDocument40
    
    'run specific post-processing function dependant on operation
    Select Case xmlGetMandatoryNode(vxmlRequestNode, "@OPERATION").Text
        
        Case "CREATEAPPLICATION"
            
            ' MO - 25/11/2002 - BMIDS01073
            'create the current activity record
            'CreateActivity BuildCreateActivityRequest(vxmlOmigaFormatRequestNode, cstrSourceApplication, cstrActivityID, cstrCreateApplicationStageID)
            
            'calculate the mortgage costs
            CalculateMortgageCosts BuildMortgageCalcRequest(vxmlOmigaFormatRequestNode)
            
            'process any one off costs that are going to be added to the loan
    ' SR 19/12/2002 : SR Debug
    Dim mtmStart As Single, mtmStop As Single
    mtmStart = Timer
    ' SR 19/12/2002 : SR Debug - End
            ProcessOneOffCosts vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode, cstrCreateApplicationStageID
    ' SR 19/12/2002 : SR Debug
    mtmStop = Timer
    Debug.Print "PostProcessing.ProcessOneOffCosts : " & Format(mtmStop - mtmStart, "0.00000000")
    ' SR 19/12/2002 : SR Debug - End
    
            'get the total net monthly cost and update the quotation record
    ' SR 19/12/2002 : SR Debug
    mtmStart = Timer
    ' SR 19/12/2002 : SR Debug - End
            UpdateQuotationRecord vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
    ' SR 19/12/2002 : SR Debug
    mtmStop = Timer
    Debug.Print "PostProcessing.UpdateQuotationRecord : " & Format(mtmStop - mtmStart, "0.00000000")
    ' SR 19/12/2002 : SR Debug - End
                    
            'make any payments to this application, which were taken prior to ingestion
    ' SR 19/12/2002 : SR Debug
    mtmStart = Timer
    ' SR 19/12/2002 : SR Debug - End

            MakeOneOffCostPayments vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
   ' SR 19/12/2002 : SR Debug
    mtmStop = Timer
    Debug.Print "PostProcessing.MakeOneOffCostPayments : " & Format(mtmStop - mtmStart, "0.00000000")
    ' SR 19/12/2002 : SR Debug - End

    End Select
    
    'copy the Omiga formatted request to the response
    Set objDOM = New FreeThreadedDOMDocument40
    objDOM.validateOnParse = False
    objDOM.setProperty "NewParser", True
    vxmlResponseNode.appendChild objDOM.createElement("OMIGADATA")
    vxmlResponseNode.selectSingleNode("OMIGADATA").appendChild vxmlOmigaFormatRequestNode.cloneNode(True)

    
DoRequestExit:
    
    'destroy objects
    Set objDOM = Nothing

Exit Sub

DoRequestErrorOccured:

    Err.Raise Err.Number, Err.Source & " : DoRequest", Err.Description, Err.HelpFile, Err.HelpContext
    Resume DoRequestExit
    
End Sub

'------------------------------------------------------------------------
' Procedure CreateActivity
' Author: MO
' Date: 30/07/2002
' Purpose: Creates the activity within Task Management for the current stage
' Input parameters: vstrRequest
' Output parameters: None
'------------------------------------------------------------------------
Private Sub CreateActivity(ByVal vstrRequest As String)

    On Error GoTo CreateActivityErrorOccured
    
    Dim strResponse As String
    Dim objOmTm As Object
    
    Set objOmTm = gobjContext.CreateInstance("omTm.omTmBO")
    strResponse = objOmTm.OmTmRequest(vstrRequest)
    Set objOmTm = Nothing
    
    errCheckXMLResponse strResponse
    
CreateActivityExit:
    
    Set objOmTm = Nothing
    
    Exit Sub
    
CreateActivityErrorOccured:
    
    Err.Raise Err.Number, "CreateActivity." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume CreateActivityExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildCreateActivityRequest
' Author: MO
' Date: 30/07/2002
' Purpose: Creates the activity request to create the current stage for
'           an application
' Input parameters: vxmlRequestNode (omiga formatted)
' Output parameters: String (omTM Request for createactivity)
'------------------------------------------------------------------------
Private Function BuildCreateActivityRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal strSourceApplication As String, ByVal strActivityId As String, ByVal strStageId As String) As String
    
On Error GoTo BuildCreateActivityRequestErrorOccured
    
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlActivityRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    Set xmlActivityRequestDoc = New FreeThreadedDOMDocument40
    xmlActivityRequestDoc.validateOnParse = False
    xmlActivityRequestDoc.setProperty "NewParser", True
    xmlActivityRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlActivityRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    xmlNode.Attributes.getNamedItem("OPERATION").Text = "CreateActivity"
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    
    strApplicationNumber = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    strApplicationFactFindNumber = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    
    Set xmlElem = xmlActivityRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
   
    Set xmlElem = xmlActivityRequestDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "SOURCEAPPLICATION", strSourceApplication
    xmlElem.setAttribute "ACTIVITYID", strActivityId
    xmlElem.setAttribute "STAGEID", strStageId
    xmlElem.setAttribute "CASEID", strApplicationNumber
    xmlNode.appendChild xmlElem
    
    BuildCreateActivityRequest = xmlNode.xml
    
BuildCreateActivityRequestExit:

    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlActivityRequestDoc = Nothing

Exit Function

BuildCreateActivityRequestErrorOccured:

    Err.Raise Err.Number, "BuildCreateActivityRequest." & Err.Source, Err.Description
    Resume BuildCreateActivityRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure CalculateMortgageCosts
' Author: MO
' Date: 14/08/2002
' Purpose: Calculates the mortgage costs
' Input parameters: vstrRequest
' Output parameters: None
'------------------------------------------------------------------------
Private Sub CalculateMortgageCosts(ByVal vstrRequest As String)

    On Error GoTo CalculateMortgageCostsErrorOccured
        
    Dim objOmAQ As Object
    Dim strResponse As String
    
    Set objOmAQ = gobjContext.CreateInstance("omAQ.ApplicationQuoteBO")
    
    strResponse = objOmAQ.CalculateMortgageCosts(vstrRequest)
    
    Set objOmAQ = Nothing
        
    errCheckXMLResponse strResponse
    
CalculateMortgageCostsExit:
    
    Set objOmAQ = Nothing
    
Exit Sub
    
CalculateMortgageCostsErrorOccured:
    
    Err.Raise Err.Number, "CalculateMortgageCosts." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume CalculateMortgageCostsExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildMortgageCalcRequest
' Author: MO
' Date: 14/08/2002
' Purpose: Creates the request to calculate the mortgage cost
' Input parameters: vxmlRequestNode (omiga formatted)
' Output parameters: String (omTM Request for calcmortgagecosts)
'------------------------------------------------------------------------
Private Function BuildMortgageCalcRequest(vxmlRequestNode As IXMLDOMNode) As String
    
On Error GoTo BuildMortgageCalcRequestErrorOccured
    
    Dim xmlCalcRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerVersionNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlOneOffCostListNode As IXMLDOMNode
    Dim xmlOneOffCostNode As IXMLDOMNode
    
    Set xmlCalcRequestDoc = New FreeThreadedDOMDocument40
    xmlCalcRequestDoc.validateOnParse = False
    xmlCalcRequestDoc.setProperty "NewParser", True
    xmlCalcRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlCalcRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    Set xmlElem = xmlCalcRequestDoc.createElement("MORTGAGECOSTS")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlCalcRequestDoc.createElement("CONTEXT")
    xmlElem.Text = "Cost Modelling"
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlCalcRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlCalcRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlCalcRequestDoc.createElement("QUOTATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/@QUOTATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlCalcRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    ' should not be necessary for BMIDS, the node is required but the value isnt!
    Set xmlElem = xmlCalcRequestDoc.createElement("LIFESUBQUOTENUMBER")
    xmlNode.appendChild xmlElem
        
    Set xmlElem = xmlCalcRequestDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerListNode = xmlNode.appendChild(xmlElem)
    
    For Each xmlCustomerVersionNode In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMER/CUSTOMERVERSION")
    
        Set xmlElem = xmlCalcRequestDoc.createElement("CUSTOMER")
        Set xmlCustomerNode = xmlCustomerListNode.appendChild(xmlElem)
    
        Set xmlElem = xmlCalcRequestDoc.createElement("CUSTOMERNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
        Set xmlElem = xmlCalcRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
    Next
    
    'MO - 25/11/2002 - Removed old debugging code, this is not run!
    'Set xmlOneOffCostListNode = xmlCalcRequestDoc.createElement("ONEOFFCOSTLIST")
    'xmlNode.appendChild xmlOneOffCostListNode
    '
    'Set xmlOneOffCostNode = xmlCalcRequestDoc.createElement("ONEOFFCOST")
    'xmlOneOffCostListNode.appendChild xmlOneOffCostNode
   '
   ' Set xmlElem = xmlCalcRequestDoc.createElement("IDENTIFIER")
    'xmlElem.Text = "ARR"
    'xmlOneOffCostNode.appendChild xmlElem
    '
    'Set xmlElem = xmlCalcRequestDoc.createElement("AMOUNT")
    'xmlElem.Text = "299"
    'xmlOneOffCostNode.appendChild xmlElem
   '
   ' Set xmlElem = xmlCalcRequestDoc.createElement("ADDEDTOLOAN")
   ' xmlElem.Text = "1"
   ' xmlOneOffCostNode.appendChild xmlElem
   '
   ' Set xmlElem = xmlCalcRequestDoc.createElement("AUTOADDIND")
   ' xmlElem.Text = "0"
   ' xmlOneOffCostNode.appendChild xmlElem
   '
   ' Set xmlElem = xmlCalcRequestDoc.createElement("INCLUDEINCOSTS")
   ' xmlElem.Text = "1"
   ' xmlOneOffCostNode.appendChild xmlElem
   
    BuildMortgageCalcRequest = xmlCalcRequestDoc.xml
    
BuildMortgageCalcRequestExit:
    
    Set xmlCalcRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCustomerNode = Nothing

Exit Function

BuildMortgageCalcRequestErrorOccured:

    Err.Raise Err.Number, "BuildMortgageCalcRequest." & Err.Source, Err.Description
    Resume BuildMortgageCalcRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure UpdateQuotationRecord
' Author: MO
' Date: 14/08/2002
' Purpose: Updates the quotation record with the values that were generated
'           by the calculation procedure
' Input parameters: vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
' Output parameters:
'------------------------------------------------------------------------
Private Sub UpdateQuotationRecord(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateQuotationRecordErrorOccured
    
    Dim objOmAQApplicationBO As Object
    Dim objOmCMQuotationBO As Object
    Dim objOmAQApplicationQuoteBo As Object
    Dim objOmCMMSQBO As Object ' SR 20/12/2002 : BM0209
    Dim objResponseDOM As FreeThreadedDOMDocument40
    Dim strResponse As String
    Dim strTotalNetMonthlyCost As String
    
    'get the total net monthly cost, which was generated by the calculation
    Set objOmAQApplicationBO = gobjContext.CreateInstance("omAQ.ApplicationQuoteBO")
    Set objOmCMMSQBO = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
    ' SR 20/12/2002 : BM0209 - just retrieve MortgageSubQuote data
    ' strResponse = objOmAQApplicationBO.GetLoanCompositionDetails(BuildGetLoanCompositionDetailsRequest(vxmlOmigaFormatRequestNode))
    strResponse = objOmCMMSQBO.GetData(BuildGetMSQDataRequest(vxmlOmigaFormatRequestNode))
    ' SR 20/12/2002 : BM0209 - End
    
    errCheckXMLResponse strResponse
    
    Set objResponseDOM = New FreeThreadedDOMDocument40
    objResponseDOM.validateOnParse = False
    objResponseDOM.setProperty "NewParser", True
    
    'pull out the total net monthly cost from the response
    objResponseDOM.loadXML strResponse
    
    strTotalNetMonthlyCost = xmlGetMandatoryNode(objResponseDOM.documentElement, "MORTGAGESUBQUOTE/TOTALNETMONTHLYCOST").Text
    
    Set objResponseDOM = Nothing
    
    Set objOmAQApplicationBO = Nothing
    
    'save this to the quotation record
    Set objOmCMQuotationBO = gobjContext.CreateInstance("omCM.QuotationBO")
    
    strResponse = objOmCMQuotationBO.StoreQuotation(BuildUpdateQuotationRequest(vxmlOmigaFormatRequestNode, strTotalNetMonthlyCost))
    
    errCheckXMLResponse strResponse
    
    Set objOmCMQuotationBO = Nothing
    
    'set the active quotation record
    Set objOmAQApplicationQuoteBo = gobjContext.CreateInstance("omAQ.ApplicationQuoteBO")
    
    strResponse = objOmAQApplicationQuoteBo.AcceptQuotation(BuildAcceptQuotationRequest(vxmlOmigaFormatRequestNode))
    
    errCheckXMLResponse strResponse
    
    Set objOmAQApplicationQuoteBo = Nothing
    
UpdateQuotationRecordExit:
    
   'destroy objects
   Set objOmAQApplicationBO = Nothing
   Set objOmCMQuotationBO = Nothing
   Set objOmAQApplicationQuoteBo = Nothing
   Set objResponseDOM = Nothing
   
Exit Sub

UpdateQuotationRecordErrorOccured:

    Err.Raise Err.Number, "UpdateQuotationRecord." & Err.Source, Err.Description
    Resume UpdateQuotationRecordExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildGetLoanCompositionDetailsRequest
' Author: MO
' Date: 22/08/2002
' Purpose: Creates the request to retreive the loan composition details
' Input parameters: vxmlOmigaFormatRequestNode
' Output parameters:
'------------------------------------------------------------------------
Private Function BuildGetLoanCompositionDetailsRequest(vxmlOmigaFormatRequestNode As IXMLDOMNode) As String
    
On Error GoTo BuildGetLoanCompositionDetailsRequestErrorOccured
    
    Dim xmlGetLoanCompRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlGetLoanCompRequestDoc = New FreeThreadedDOMDocument40
    xmlGetLoanCompRequestDoc.validateOnParse = False
    xmlGetLoanCompRequestDoc.setProperty "NewParser", True
    xmlGetLoanCompRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlGetLoanCompRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlGetLoanCompRequestDoc.createElement("LOANCOMPOSITION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlGetLoanCompRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetLoanCompRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
   
    Set xmlElem = xmlGetLoanCompRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetLoanCompRequestDoc.createElement("QUOTATIONTYPE")
    xmlElem.Text = "2" ' Cost modelling
    xmlNode.appendChild xmlElem
    
    BuildGetLoanCompositionDetailsRequest = xmlGetLoanCompRequestDoc.xml
    
BuildGetLoanCompositionDetailsRequestExit:
    
    Set xmlGetLoanCompRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

BuildGetLoanCompositionDetailsRequestErrorOccured:

    Err.Raise Err.Number, "BuildGetLoanCompositionDetailsRequest." & Err.Source, Err.Description
    Resume BuildGetLoanCompositionDetailsRequestExit
    
End Function

Private Function BuildGetMSQDataRequest(vxmlOmigaFormatRequestNode As IXMLDOMNode) As String

On Error GoTo BuildGetMSQDataRequestErrorOccured
    
    Dim xmlGetMSQRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlGetMSQRequestDoc = New FreeThreadedDOMDocument40
    xmlGetMSQRequestDoc.validateOnParse = False
    xmlGetMSQRequestDoc.setProperty "NewParser", True
    xmlGetMSQRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlGetMSQRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlGetMSQRequestDoc.createElement("MORTGAGESUBQUOTE")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlGetMSQRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetMSQRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
   
    Set xmlElem = xmlGetMSQRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    BuildGetMSQDataRequest = xmlGetMSQRequestDoc.xml
    
BuildGetMSQDataRequestExit:
    
    Set xmlGetMSQRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

BuildGetMSQDataRequestErrorOccured:

    Err.Raise Err.Number, "BuildGetMSQDataRequest." & Err.Source, Err.Description
    Resume BuildGetMSQDataRequestExit

End Function


'------------------------------------------------------------------------
' Procedure BuildUpdateQuotationRequest
' Author: MO
' Date: 22/08/2002
' Purpose: Creates the request to update the quotation record
' Input parameters: vxmlOmigaFormatRequestNode
' Output parameters:
'------------------------------------------------------------------------
Private Function BuildUpdateQuotationRequest(vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal strTotalQuotationCost As String) As String
    
On Error GoTo BuildUpdateQuotationRequestErrorOccured
    
    Dim xmlUpdateQuoteRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlUpdateQuoteRequestDoc = New FreeThreadedDOMDocument40
    xmlUpdateQuoteRequestDoc.validateOnParse = False
    xmlUpdateQuoteRequestDoc.setProperty "NewParser", True
    xmlUpdateQuoteRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlUpdateQuoteRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("UPDATE")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("QUOTATION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
   
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("QUOTATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/@QUOTATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("TOTALQUOTATIONCOST")
    xmlElem.Text = strTotalQuotationCost
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlUpdateQuoteRequestDoc.createElement("QUOTATIONCOMPLETE")
    xmlElem.Text = "1"
    xmlNode.appendChild xmlElem
    
    BuildUpdateQuotationRequest = xmlUpdateQuoteRequestDoc.xml
    
BuildUpdateQuotationRequestExit:
    
    Set xmlUpdateQuoteRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

BuildUpdateQuotationRequestErrorOccured:

    Err.Raise Err.Number, "BuildUpdateQuotationRequest." & Err.Source, Err.Description
    Resume BuildUpdateQuotationRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildAcceptQuotationRequest
' Author: MO
' Date: 22/08/2002
' Purpose: Creates the request to accept the quoation
' Input parameters: vxmlOmigaFormatRequestNode
' Output parameters:
'------------------------------------------------------------------------
Private Function BuildAcceptQuotationRequest(vxmlOmigaFormatRequestNode As IXMLDOMNode) As String
    
On Error GoTo BuildAcceptQuotationRequestErrorOccured
    
    Dim xmlAcceptQuoteRequestDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlAcceptQuoteRequestDoc = New FreeThreadedDOMDocument40
    xmlAcceptQuoteRequestDoc.validateOnParse = False
    xmlAcceptQuoteRequestDoc.setProperty "NewParser", True
    xmlAcceptQuoteRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlAcceptQuoteRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlAcceptQuoteRequestDoc.createElement("QUOTATION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlAcceptQuoteRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlAcceptQuoteRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
   
    BuildAcceptQuotationRequest = xmlAcceptQuoteRequestDoc.xml
    
BuildAcceptQuotationRequestExit:
    
    Set xmlAcceptQuoteRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

BuildAcceptQuotationRequestErrorOccured:

    Err.Raise Err.Number, "BuildAcceptQuotationRequest." & Err.Source, Err.Description
    Resume BuildAcceptQuotationRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure GetLoanCompositionDetails
' Author: MO
' Date: 22/08/2002
' Purpose: Gets the loan composition details
' Input parameters: vxmlOmigaFormatRequestNode
' Output parameters:
'------------------------------------------------------------------------
Private Function GetLoanCompositionDetails(vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal strStageId As String) As IXMLDOMNode
    
On Error GoTo GetLoanCompositionDetailsErrorOccured
    
    Dim xmlGetLoanCompositionRequestDoc As FreeThreadedDOMDocument40
    Dim objOmAQApplicationBO As Object
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strResponse As String
    
    Set xmlGetLoanCompositionRequestDoc = New FreeThreadedDOMDocument40
    xmlGetLoanCompositionRequestDoc.validateOnParse = False
    xmlGetLoanCompositionRequestDoc.setProperty "NewParser", True
    xmlGetLoanCompositionRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlGetLoanCompositionRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlGetLoanCompositionRequestDoc.createElement("LOANCOMPOSITION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlGetLoanCompositionRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetLoanCompositionRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlGetLoanCompositionRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetLoanCompositionRequestDoc.createElement("CURRENTSTAGEID")
    xmlElem.Text = strStageId
    xmlNode.appendChild xmlElem

    Set objOmAQApplicationBO = CreateObject("omAQ.ApplicationQuoteBO")
    
    strResponse = objOmAQApplicationBO.GetLoanCompositionDetails(xmlGetLoanCompositionRequestDoc.xml)
    
    errCheckXMLResponse strResponse
    
    xmlGetLoanCompositionRequestDoc.loadXML strResponse
    
    Set GetLoanCompositionDetails = xmlGetLoanCompositionRequestDoc.documentElement
    
GetLoanCompositionDetailsExit:
    
    Set xmlGetLoanCompositionRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

GetLoanCompositionDetailsErrorOccured:

    Err.Raise Err.Number, "GetLoanCompositionDetails." & Err.Source, Err.Description
    Resume GetLoanCompositionDetailsExit
    
End Function

'------------------------------------------------------------------------
' Procedure GetOneOffCostsDetail
' Author: MO
' Date: 22/08/2002
' Purpose: Gets the one off costs detail
' Input parameters: vxmlOmigaFormatRequestNode
' Output parameters:
'------------------------------------------------------------------------
Private Function GetOneOffCostsDetails(vxmlOmigaFormatRequestNode As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo GetOneOffCostsDetailsErrorOccured
    
    Dim xmlGetOneOffCostsDetailsRequestDoc As FreeThreadedDOMDocument40
    Dim objOmCmMortgageSubQuoteBO As Object
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strResponse As String
    
    Set xmlGetOneOffCostsDetailsRequestDoc = New FreeThreadedDOMDocument40
    xmlGetOneOffCostsDetailsRequestDoc.validateOnParse = False
    xmlGetOneOffCostsDetailsRequestDoc.setProperty "NewParser", True
    xmlGetOneOffCostsDetailsRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlGetOneOffCostsDetailsRequestDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))
    
    Set xmlElem = xmlGetOneOffCostsDetailsRequestDoc.createElement("ONEOFFCOST")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlGetOneOffCostsDetailsRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlGetOneOffCostsDetailsRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlGetOneOffCostsDetailsRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set objOmCmMortgageSubQuoteBO = CreateObject("omCM.MortgageSubQuoteBO")
    
    strResponse = objOmCmMortgageSubQuoteBO.GetOneOffCostsDetails(xmlGetOneOffCostsDetailsRequestDoc.xml)
    
    errCheckXMLResponse strResponse
    
    xmlGetOneOffCostsDetailsRequestDoc.loadXML strResponse
    
    Set GetOneOffCostsDetails = xmlGetOneOffCostsDetailsRequestDoc.documentElement

GetOneOffCostsDetailsExit:
    
    Set xmlGetOneOffCostsDetailsRequestDoc = Nothing
    Set objOmCmMortgageSubQuoteBO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

Exit Function

GetOneOffCostsDetailsErrorOccured:

    Err.Raise Err.Number, "GetOneOffCostsDetails." & Err.Source, Err.Description
    Resume GetOneOffCostsDetailsExit
    
End Function

'------------------------------------------------------------------------
' Procedure ProcessOneOffCosts
' Author: MO
' Date: 27/08/2002
' Purpose: Process any one off costs that have been added to the loan and
'           recosts the mortgage application
' Input parameters: vstrRequest, omigaformattedrequestxml, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub ProcessOneOffCosts(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode, ByVal strStageId As String)

    On Error GoTo ProcessOneOffCostsErrorOccured
    
    Dim strFeeControl As String
    Dim strFeeControlArray() As String
    Dim lngCount As Long
    Dim objProcessOneOffCostsDOM As FreeThreadedDOMDocument40
    Dim objOneOffCostList As IXMLDOMNode
    Dim objOneOffCostNode As IXMLDOMNode
    Dim objFeeNode As IXMLDOMNode
    Dim objOneOffCostDetail As IXMLDOMNode
    Dim objMortgageOneOffCost As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    Dim strEloansFeeType As String
    Dim omAQApplicationQuoteBO As Object
    Dim blnAddedToLoanInd As Boolean
    Dim dblAmountAddedToLoan As Double
    Dim strResponse As String
    Dim objRecostResponseDOM As FreeThreadedDOMDocument40
    Dim omCMQuotationBO As Object
    
    'get the fee control list from the request xml
    strFeeControl = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/FeeControl").Text
    
    'make sure that a feecontrol is present
     If Len(strFeeControl) > 0 Then
        
        blnAddedToLoanInd = False
    
        'split the fee control into the array
        'the fee control looks a bit like this 0,0,0,120CI,0,0,0,0,120CI,0,0,0
        'there is an item for each fee type, going in order
        'a zero means it hasnt been added to the loan
        'any other value indicates the product number and whether its been added to capital
        '   and interest (CI) or interest only (IO)
        strFeeControlArray = Split(strFeeControl, ",")
        
        Set objOneOffCostDetail = GetOneOffCostsDetails(vxmlOmigaFormatRequestNode)
        
        Set objProcessOneOffCostsDOM = New FreeThreadedDOMDocument40
        objProcessOneOffCostsDOM.validateOnParse = False
        objProcessOneOffCostsDOM.setProperty "NewParser", True
        
        'create a new oneoffcostdetaillist
        Set objOneOffCostList = objProcessOneOffCostsDOM.createElement("ONEOFFCOSTLIST")
        
        'loop through the one off costs that have been calculated by omiga
        For Each objMortgageOneOffCost In objOneOffCostDetail.selectNodes("ONEOFFCOST/MORTGAGEONEOFFCOSTLIST/MORTGAGEONEOFFCOST")
            
            If Val(objMortgageOneOffCost.selectSingleNode("AMOUNT").Text) > 0 Then
                
                'create a new oneoffcost node
                Set objOneOffCostNode = objProcessOneOffCostsDOM.createElement("ONEOFFCOST")
                objOneOffCostList.appendChild objOneOffCostNode
                
                'put on the combovalidationtype node
                Set objNode = objProcessOneOffCostsDOM.createElement("COMBOVALIDATIONTYPE")
                objNode.Text = GetValidationTypeForFee(xmlGetMandatoryNode(objMortgageOneOffCost, "MORTGAGEONEOFFCOSTTYPE").Text)
                objOneOffCostNode.appendChild objNode
        
                'put on the name node
                Set objNode = objProcessOneOffCostsDOM.createElement("NAME")
                objOneOffCostNode.appendChild objNode
                
                'put on the amount node
                Set objNode = objProcessOneOffCostsDOM.createElement("AMOUNT")
                objNode.Text = objMortgageOneOffCost.selectSingleNode("AMOUNT").Text
                objOneOffCostNode.appendChild objNode
                
                'has this one off cost been added to the loan?
                
                'put on the added node
                Set objNode = objProcessOneOffCostsDOM.createElement("ADDEDTOLOAN")
                
                'does this fee map to eloans fees?
                strEloansFeeType = GetEloansFeeType(xmlGetMandatoryNode(objMortgageOneOffCost, "MORTGAGEONEOFFCOSTTYPE").Text)
                
                If Len(strEloansFeeType) > 0 Then
                    'get the fee control from the fee control array
                    If strFeeControlArray(CLng(strEloansFeeType)) <> "0" Then
                        objNode.Text = "1"
                        blnAddedToLoanInd = True
                        dblAmountAddedToLoan = dblAmountAddedToLoan + CDbl(objMortgageOneOffCost.selectSingleNode("AMOUNT").Text)
                    Else
                        objNode.Text = "0"
                    End If
                Else
                    objNode.Text = "0"
                End If
            
                objOneOffCostNode.appendChild objNode
                        
            End If
        Next
        
        'only recost and process the one off costs if any fees have been added to the loan
        If blnAddedToLoanInd = True Then
            
            'PROCESS ADDED ONE OFF COSTS
            Set omAQApplicationQuoteBO = gobjContext.CreateInstance("omAQ.ApplicationQuoteBO")
            
            strResponse = omAQApplicationQuoteBO.ProcessAddedOneOffCosts(BuildProcessAddedCostsRequest(vxmlOmigaFormatRequestNode, objOneOffCostList))
            
            errCheckXMLResponse strResponse
            
            'RECOST MORTGAGE AND LIFE COVER
            strResponse = omAQApplicationQuoteBO.RecostMortgageComponentAndLifeCover(BuildRecostMortgageAndLifeCoverRequest(vxmlOmigaFormatRequestNode, objOneOffCostList, dblAmountAddedToLoan, strStageId))
            
            errCheckXMLResponse strResponse
            
            Set omAQApplicationQuoteBO = Nothing
            
            Set objRecostResponseDOM = New FreeThreadedDOMDocument40
            objRecostResponseDOM.validateOnParse = False
            objRecostResponseDOM.setProperty "NewParser", True
            
            objRecostResponseDOM.loadXML strResponse
            
            'SAVE ONE OFF COSTS DETAILS
            Set omCMQuotationBO = gobjContext.CreateInstance("omCM.QuotationBO")
            
            strResponse = omCMQuotationBO.SaveOneOffCostDetails(BuildSaveOneOffCostDetailsRequest(vxmlOmigaFormatRequestNode, objOneOffCostList, objRecostResponseDOM.documentElement, dblAmountAddedToLoan))

            Set omCMQuotationBO = Nothing
            
            errCheckXMLResponse strResponse
            
        End If
        
    Else
        Err.Raise direFeeControlStringMissing, "ProcessOneOffCosts." & Err.Source, "The fee control string which dictates which fees may have been added to the loan, was missing."
    End If
    
ProcessOneOffCostsExit:
    
    'destroy objects
    Set objOneOffCostDetail = Nothing
    Set objOneOffCostList = Nothing
    Set objProcessOneOffCostsDOM = Nothing
    Set objOneOffCostDetail = Nothing
    Set objFeeNode = Nothing
    Set objMortgageOneOffCost = Nothing
    Set omAQApplicationQuoteBO = Nothing
    Set objRecostResponseDOM = Nothing
    
Exit Sub
    
ProcessOneOffCostsErrorOccured:
    
    Err.Raise Err.Number, "ProcessOneOffCosts." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume ProcessOneOffCostsExit
    
End Sub

'------------------------------------------------------------------------
' Procedure GetValidationTypeForFee
' Author: MO
' Date: 27/08/2002
' Purpose: Returns the validation fee type for a fee type id
' Input parameters: FeeTypeId
' Note:  These fee types are hardcoded as the application needs the
'           specific validation type that is used in cost modelling,
'           and as a value can have more than one validation type, it
'           would not be possible to know which one was needed.
' Output parameters: FeeValidationType
'------------------------------------------------------------------------
Private Function GetValidationTypeForFee(ByVal strFeeType As String) As String

    On Error GoTo GetValidationTypeForFeeErrorOccured
    
    Select Case strFeeType
        Case "1" ' Arrangement
            GetValidationTypeForFee = "ARR"
        Case "2" ' Porting
            GetValidationTypeForFee = "POR"
        Case "3" ' Valuation
            GetValidationTypeForFee = "VAL"
        Case "4" ' Reinspection
            GetValidationTypeForFee = "REI"
        Case "5" ' Administration
            GetValidationTypeForFee = "ADM"
        Case "6" ' Stamp Duty
            GetValidationTypeForFee = "STA"
        Case "7" ' Sealing Fee
            GetValidationTypeForFee = "SEA"
        Case "8" ' Deeds release fee
            GetValidationTypeForFee = "DEE"
        Case "9" ' TT
            GetValidationTypeForFee = "TTF"
        Case "10" ' Legal
            GetValidationTypeForFee = "LEG"
        Case "11" ' Other
            GetValidationTypeForFee = "OTH"
        Case "12" ' HPAF
            GetValidationTypeForFee = "MIG"
        Case "13" ' Total indemnity
            GetValidationTypeForFee = "TID"
        Case "14" ' Revaluation
            GetValidationTypeForFee = "REV"
        Case Else ' Anything else
            'raise an error
            Err.Raise direFeeValuationTypeNotKnown, "GetValidationTypeForFee", "The omiga validation type for FeeType " & strFeeType & " is not known."
    End Select
    
GetValidationTypeForFeeExit:
    
    'destroy objects
    
Exit Function
    
GetValidationTypeForFeeErrorOccured:
    
    Err.Raise Err.Number, "GetValidationTypeForFee." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume GetValidationTypeForFeeExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildRecostMortgageAndLifeCoverRequest
' Author: MO
' Date: 30/07/2002
' Purpose: Creates the request to recost the mortgage costs
' Input parameters: vxmlOmigaFormattedRequestNode (omiga formatted)
' Output parameters: String (omTM Request for createactivity)
'------------------------------------------------------------------------
Private Function BuildRecostMortgageAndLifeCoverRequest(ByVal vxmlOmigaFormattedRequestNode As IXMLDOMNode, ByVal vxmlOneOffCostListXML As IXMLDOMNode, ByVal dblAmountAddedToLoan As Double, ByVal strStageId As String) As String
    
On Error GoTo BuildRecostMortgageAndLifeCoverRequestErrorOccured
    
    Dim xmlRecostMortgageRequestDoc As FreeThreadedDOMDocument40
    Dim objLoanCompositionDetails As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustomerVersionNode As IXMLDOMNode
    
    Set objLoanCompositionDetails = GetLoanCompositionDetails(vxmlOmigaFormattedRequestNode, strStageId)
    
    Set xmlRecostMortgageRequestDoc = New FreeThreadedDOMDocument40
    xmlRecostMortgageRequestDoc.validateOnParse = False
    xmlRecostMortgageRequestDoc.setProperty "NewParser", True
    xmlRecostMortgageRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlRecostMortgageRequestDoc.appendChild(vxmlOmigaFormattedRequestNode.cloneNode(False))
    
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("APPLICATIONQUOTE")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CONTEXT")
    xmlElem.Text = "Cost Modelling"
    xmlNode.appendChild xmlElem
    
    'increase the total loan amount of the first loan component, by the increase in value
    xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/TOTALLOANCOMPONENTAMOUNT").Text = CDbl(xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/TOTALLOANCOMPONENTAMOUNT").Text) + dblAmountAddedToLoan
    
    'if the first loan component is a part and part populate the interestonly and capitalandinterest amounts
    If xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/REPAYMENTMETHOD").Text = "3" Then 'part and part
        
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("INTERESTONLYAMOUNT")
        xmlElem.Text = xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/@INTERESTONLYELEMENT").Text
        xmlNode.appendChild xmlElem
        
        'add the amount added to the loan to capital and interest portion
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CAPITALANDINTERESTAMOUNT")
        xmlElem.Text = CDbl(xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/@CAPITALANDINTERESTELEMENT").Text) + dblAmountAddedToLoan
        xmlNode.appendChild xmlElem
        
    Else
        
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("INTERESTONLYAMOUNT")
        xmlElem.Text = "0"
        xmlNode.appendChild xmlElem
    
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CAPITALANDINTERESTAMOUNT")
        xmlElem.Text = "0"
        xmlNode.appendChild xmlElem
        
    End If
    
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("APPLICATIONDATE")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/@APPLICATIONDATE").Text
    xmlNode.appendChild xmlElem
    
    ' append the loan component list node to the application quote
    xmlNode.appendChild xmlGetMandatoryNode(objLoanCompositionDetails, "MORTGAGESUBQUOTE/LOANCOMPONENTLIST").cloneNode(True)
    
    ' append the one off cost list to the application quote
    xmlNode.appendChild vxmlOneOffCostListXML.cloneNode(True)
    
    'create the customer list and append it to the application quote
    Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerListNode = xmlNode.appendChild(xmlElem)
    
    For Each xmlCustomerVersionNode In vxmlOmigaFormattedRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMER/CUSTOMERVERSION")
    
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CUSTOMER")
        Set xmlCustomerNode = xmlCustomerListNode.appendChild(xmlElem)
    
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CUSTOMERNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
        Set xmlElem = xmlRecostMortgageRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
    Next
    
    BuildRecostMortgageAndLifeCoverRequest = xmlRecostMortgageRequestDoc.xml
    
BuildRecostMortgageAndLifeCoverRequestExit:

    'destroy objects
    Set xmlRecostMortgageRequestDoc = Nothing
    Set objLoanCompositionDetails = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerVersionNode = Nothing
    
Exit Function

BuildRecostMortgageAndLifeCoverRequestErrorOccured:

    Err.Raise Err.Number, "BuildRecostMortgageAndLifeCoverRequest." & Err.Source, Err.Description
    Resume BuildRecostMortgageAndLifeCoverRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure GetEloansFeeType
' Author: MO
' Date: 27/08/2002
' Purpose: Returns the validation fee type for a fee type id
' Input parameters: OmigaMortgageOneOffCostType
' Output parameters: EloansFeeTypeId
'------------------------------------------------------------------------
Private Function GetEloansFeeType(ByVal strOmigaMortgageOneOffCostType As String) As String

    On Error GoTo GetEloansFeeTypeErrorOccured
    
    Select Case strOmigaMortgageOneOffCostType
        
        Case "1" ' Arrangement fee
            GetEloansFeeType = "1"
        Case "3" ' Valuation fee
            GetEloansFeeType = "0"
        Case "5" ' Administration fee
            GetEloansFeeType = "4"
        Case "9" ' TT fee
            GetEloansFeeType = "9"
        Case "10" ' Legal fee
            GetEloansFeeType = "2"
        Case "12" ' MIG/HPAF fee
            GetEloansFeeType = "7"
    
    End Select
    
GetEloansFeeTypeExit:
    
Exit Function
    
GetEloansFeeTypeErrorOccured:
    
    Err.Raise Err.Number, "GetEloansFeeType." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume GetEloansFeeTypeExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildProcessAddedCostsRequest
' Author: MO
' Date: 30/07/2002
' Purpose: Creates the request to process any added one off costs
' Input parameters: vxmlOmigaFormattedRequestNode (omiga formatted)
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildProcessAddedCostsRequest(ByVal vxmlOmigaFormattedRequestNode As IXMLDOMNode, ByVal vxmlOneOffCostListXML As IXMLDOMNode) As String
    
On Error GoTo BuildProcessAddedCostsRequestErrorOccured
    
    Dim xmlProcessAddedRequestDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustomerVersionNode As IXMLDOMNode
    
    Set xmlProcessAddedRequestDoc = New FreeThreadedDOMDocument40
    xmlProcessAddedRequestDoc.validateOnParse = False
    xmlProcessAddedRequestDoc.setProperty "NewParser", True
    xmlProcessAddedRequestDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlProcessAddedRequestDoc.appendChild(vxmlOmigaFormattedRequestNode.cloneNode(False))
    
    Set xmlElem = xmlProcessAddedRequestDoc.createElement("ADDEDONEOFFCOSTS")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlProcessAddedRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlProcessAddedRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlProcessAddedRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    ' append the one off cost list to the application quote
    xmlNode.appendChild vxmlOneOffCostListXML.cloneNode(True)
    
    'create the customer list and append it to the application quote
    Set xmlElem = xmlProcessAddedRequestDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerListNode = xmlNode.appendChild(xmlElem)
    
    For Each xmlCustomerVersionNode In vxmlOmigaFormattedRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMER/CUSTOMERVERSION")
    
        Set xmlElem = xmlProcessAddedRequestDoc.createElement("CUSTOMER")
        Set xmlCustomerNode = xmlCustomerListNode.appendChild(xmlElem)
    
        Set xmlElem = xmlProcessAddedRequestDoc.createElement("CUSTOMERNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
        Set xmlElem = xmlProcessAddedRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlElem.Text = xmlGetMandatoryAttribute(xmlCustomerVersionNode, "CUSTOMERVERSIONNUMBER").Text
        xmlCustomerNode.appendChild xmlElem
    
    Next
    
    BuildProcessAddedCostsRequest = xmlProcessAddedRequestDoc.xml
    
BuildProcessAddedCostsRequestExit:

    'destroy objects
    Set xmlProcessAddedRequestDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerVersionNode = Nothing
    
Exit Function

BuildProcessAddedCostsRequestErrorOccured:

    Err.Raise Err.Number, "BuildProcessAddedCostsRequest." & Err.Source, Err.Description
    Resume BuildProcessAddedCostsRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildSaveOneOffCostDetailsRequest
' Author: MO
' Date: 28/08/2002
' Purpose: Creates the request to save any one off costs details
' Input parameters: vxmlOmigaFormattedRequestNode (omiga formatted)
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildSaveOneOffCostDetailsRequest(ByVal vxmlOmigaFormattedRequestNode As IXMLDOMNode, ByVal vxmlOneOffCostListXML As IXMLDOMNode, ByVal vxmlRecostMortgageResponseXML As IXMLDOMNode, ByVal dblAmountAddedToLoan As Double) As String
    
On Error GoTo BuildSaveOneOffCostDetailsRequestErrorOccured
    
    Dim xmlSaveDetailsDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    
    Set xmlSaveDetailsDoc = New FreeThreadedDOMDocument40
    xmlSaveDetailsDoc.validateOnParse = False
    xmlSaveDetailsDoc.setProperty "NewParser", True
    xmlSaveDetailsDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlSaveDetailsDoc.appendChild(vxmlOmigaFormattedRequestNode.cloneNode(False))

    xmlNode.Attributes.setNamedItem(xmlSaveDetailsDoc.createAttribute("OPERATION")).Text = "UPDATE"
    xmlNode.Attributes.setNamedItem(xmlSaveDetailsDoc.createAttribute("ACTION")).Text = "UPDATE"
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("QUOTATION")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/@MORTGAGESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("LIFESUBQUOTENUMBER")
    xmlElem.Text = xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/LIFESUBQUOTE/@LIFESUBQUOTENUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("LOANCOMPONENTSEQUENCENUMBER")
    xmlElem.Text = "1"
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("MIGIPT")
    xmlElem.Text = "0"
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("TOTALLOANAMOUNT")
    xmlElem.Text = CDbl(xmlGetMandatoryNode(vxmlOmigaFormattedRequestNode, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@TOTALLOANAMOUNT").Text) + dblAmountAddedToLoan
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("INTERESTONLYAMOUNT")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "LOANCOMPONENTLIST/LOANCOMPONENT/INTERESTONLYELEMENT").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("CAPITALANDINTERESTAMOUNT")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "LOANCOMPONENTLIST/LOANCOMPONENT/CAPITALANDINTERESTELEMENT").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("TOTALGROSSMONTHLYCOST")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "TOTALGROSSMONTHLYCOST").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("TOTALNETMONTHLYCOST")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "TOTALNETMONTHLYCOST").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("TOTALACCRUEDINTEREST")
    xmlElem.Text = xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "TOTALACCRUEDINTEREST").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlSaveDetailsDoc.createElement("CONTEXT")
    xmlElem.Text = "UPDATE"
    xmlNode.appendChild xmlElem
    
    ' append the one off cost list to the quotation
    xmlNode.appendChild vxmlOneOffCostListXML.cloneNode(True)
    
    ' append the loan component list to the quotation
    xmlNode.appendChild xmlGetMandatoryNode(vxmlRecostMortgageResponseXML, "LOANCOMPONENTLIST").cloneNode(True)
        
    BuildSaveOneOffCostDetailsRequest = xmlSaveDetailsDoc.xml
    
BuildSaveOneOffCostDetailsRequestExit:

    'destroy objects
    Set xmlSaveDetailsDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    
Exit Function

BuildSaveOneOffCostDetailsRequestErrorOccured:

    Err.Raise Err.Number, "BuildSaveOneOffCostDetailsRequest." & Err.Source, Err.Description
    Resume BuildSaveOneOffCostDetailsRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildCreateAdhocTaskRequest
' Author: MO
' Date: 06/09/2002
' Purpose: Creates the request to create an adhoc task
' Input parameters: vxmlOmigaFormattedRequestNode (omiga formatted)
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildCreateAdhocTaskRequest(ByVal vxmlOmigaFormattedRequestNode As IXMLDOMNode, ByVal vxmlOneOffCostListXML As IXMLDOMNode, ByVal vxmlRecostMortgageResponseXML As IXMLDOMNode, ByVal dblAmountAddedToLoan As Double) As String
    
On Error GoTo BuildCreateAdhocTaskRequestErrorOccured
    
    Dim xmlCreateAdhocTaskDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    
    Set xmlCreateAdhocTaskDoc = New FreeThreadedDOMDocument40
    xmlCreateAdhocTaskDoc.validateOnParse = False
    xmlCreateAdhocTaskDoc.setProperty "NewParser", True
    xmlCreateAdhocTaskDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlCreateAdhocTaskDoc.appendChild(vxmlOmigaFormattedRequestNode.cloneNode(False))

        
    BuildCreateAdhocTaskRequest = xmlCreateAdhocTaskDoc.xml
    
BuildCreateAdhocTaskRequestExit:

    'destroy objects
    Set xmlCreateAdhocTaskDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    
Exit Function

BuildCreateAdhocTaskRequestErrorOccured:

    Err.Raise Err.Number, "BuildCreateAdhocTaskRequest." & Err.Source, Err.Description
    Resume BuildCreateAdhocTaskRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure MakeOneOffCostPayments
' Author: MO
' Date: 16/10/2002
' Purpose: Process any one off costs payments that have been made prior
'           to the application being ingested
' Input parameters: vstrRequest, omigaformattedrequestxml, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub MakeOneOffCostPayments(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo MakeOneOffCostPaymentsErrorOccured
    
    Dim strArrangementFeeAmountPaid As String
    Dim strValuationFeeAmountPaid As String
    Dim objCardPaymentNode As IXMLDOMNode
    
    'get the card payment node
    Set objCardPaymentNode = xmlGetMandatoryNode(vxmlRequestNode, "ProposalRoot/Proposal/CardPayment")
    
    'have any card payment been taken in GIWS
    If xmlGetMandatoryNode(objCardPaymentNode, "RequiredYN").Text = "1" Then
        
        'was the payment successful?
        If xmlGetMandatoryNode(objCardPaymentNode, "PaymentSuccessfulYN").Text = "1" Then
        
            strArrangementFeeAmountPaid = xmlGetMandatoryNode(objCardPaymentNode, "PaymentArrangementFeeAmount").Text
            strValuationFeeAmountPaid = xmlGetMandatoryNode(objCardPaymentNode, "PaymentValuationFeeAmount").Text
            
            'has an arrangement fee been paid ?
            If CDbl("0" & strArrangementFeeAmountPaid) > 0 Then
                'pay the arrangement fee, feetype = 1
                PayFeeByCreditCard vxmlOmigaFormatRequestNode, objCardPaymentNode, 1, strArrangementFeeAmountPaid
            End If
            
            'has a valuation fee been paid ?
            If CDbl("0" & strValuationFeeAmountPaid) > 0 Then
                'pay the arrangement fee, feetype = 3
                PayFeeByCreditCard vxmlOmigaFormatRequestNode, objCardPaymentNode, 3, strValuationFeeAmountPaid
            End If
        
        End If
        
    End If
    
MakeOneOffCostPaymentsExit:
    
    'destroy objects
    Set objCardPaymentNode = Nothing
    
Exit Sub
    
MakeOneOffCostPaymentsErrorOccured:
    
    Err.Raise Err.Number, "MakeOneOffCostPayments." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume MakeOneOffCostPaymentsExit
    
End Sub

'------------------------------------------------------------------------
' Procedure PayFeeByCreditCard
' Author: MO
' Date: 17/10/2002
' Purpose: Makes a credit card payment amount against a single fee type
' Input parameters: vxmlOmigaFormatRequestNode, CardPaymentNode, FeeType, Amount, UserId
' Output parameters: None
'------------------------------------------------------------------------
Private Sub PayFeeByCreditCard(ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlCardPaymentNode As IXMLDOMNode, ByVal strFeeType As String, ByVal strAmount As String)

    On Error GoTo PayFeeByCreditCardErrorOccured
    
    Const cstrPaymentEvent As String = "10" ' Combo : PaymentEvent : "Payment"
    ' Text placed into the payment notes, #T is replaced with the transcation code, #A with the autorisation code and #P with the ingested application number (ProposalID)
    Const cstrNotes As String = "This payment was taken via the Internet using the Secured Card Payment application:" & vbCrLf & _
                                "P.O.S document reference number = #P (used for searching against Secured Card Payments)" & vbCrLf & _
                                "Transaction code = #T" & vbCrLf & _
                                "Authorisation code = #A" & vbCrLf

    Const cstrPaymentMethod As String = "2" ' Combo : FeePaymentMehod : "CreditCard"
    
    Dim strResponse As String
    Dim strFeeTypeSequenceNumber As String
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
    Dim objFindFeeTypeDOM As FreeThreadedDOMDocument40
    Dim strPaymentNotes As String
    
    'check to see if this fee type has actually been calculated and placed in APPLICATIONFEETYPE
    strResponse = CallPaymentProcessing(BuildFindFeeTypeListRequest(vxmlOmigaFormatRequestNode, strFeeType))
    
    'check to see if a no record found error was generated
    On Error Resume Next
    
    errCheckXMLResponse strResponse
            
    'get the error details
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
            
    On Error GoTo PayFeeByCreditCardErrorOccured
    
    'has an error has occured
    If lngErrorNumber <> 0 Then
        
        'is this error that the record cuold not be found?
        If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
            Err.Raise direApplicationFeeTypeMissing, "PayFeeByCreditCard." & strErrorSource, "An error occured when processing the payment for Fee Type " & strFeeType & " (please see Omiga 4 Combo Group OneOffCost for details).  The Fee Type could not be found in the Omiga 4 APPLICATIONFEETYPE table, this is likely because the fee was not calculated or was added to the loan."
        Else
            Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
        End If
        
    End If
    
    'clear the error
    Err.Clear
    
    'find the FeeTypeSequenceNumber
    Set objFindFeeTypeDOM = xmlLoad(strResponse)
    
    strFeeTypeSequenceNumber = xmlGetMandatoryNode(objFindFeeTypeDOM.documentElement, "APPLICATIONFEETYPE/@FEETYPESEQUENCENUMBER").Text
    
    'build up the payment notes, with the trans, auth codes and ingested application number
    strPaymentNotes = Replace(cstrNotes, "#T", xmlGetMandatoryNode(vxmlCardPaymentNode, "PaymentTransactionCode").Text)
    strPaymentNotes = Replace(strPaymentNotes, "#A", xmlGetMandatoryNode(vxmlCardPaymentNode, "PaymentAuthCode").Text)
    strPaymentNotes = Replace(strPaymentNotes, "#P", xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@INGESTIONAPPLICATIONNUMBER").Text)
    
    'make the call to create the fee type payment
    strResponse = CallPaymentProcessing(BuildCreateFeeTypePaymentRequest(vxmlOmigaFormatRequestNode, strFeeType, cstrPaymentEvent, strAmount, strPaymentNotes, strFeeTypeSequenceNumber, cstrPaymentMethod))
    
    errCheckXMLResponse strResponse
    
PayFeeByCreditCardExit:
    
    'destroy objects
    Set objFindFeeTypeDOM = Nothing
    
Exit Sub
    
PayFeeByCreditCardErrorOccured:
    
    Err.Raise Err.Number, "PayFeeByCreditCard." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume PayFeeByCreditCardExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildCreateFeeTypePaymentRequest
' Author: MO
' Date: 17/10/2002
' Purpose: Builds the request for a fee type payment request
' Input parameters: OmigaFormatRequest, FeeType, PaymentEvent, Amount, Notes, FeeTypeSequenceNumber, PaymentMethod
' Output parameters: None
'------------------------------------------------------------------------
Private Function BuildCreateFeeTypePaymentRequest(ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal strFeeType As String, ByVal strPaymentEvent As String, ByVal strAmount As String, strNotes As String, strFeeTypeSequenceNumber As String, ByVal strPaymentMethod As String) As String

    On Error GoTo BuildCreateFeeTypePaymentRequestErrorOccured
    
    Dim xmlFeeTypePaymentDoc As FreeThreadedDOMDocument40
    Dim xmlFeePaymentNode As IXMLDOMElement
    Dim xmlPaymentRecordNode As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    
    Set xmlFeeTypePaymentDoc = New FreeThreadedDOMDocument40
    xmlFeeTypePaymentDoc.validateOnParse = False
    xmlFeeTypePaymentDoc.setProperty "NewParser", True
    xmlFeeTypePaymentDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlFeeTypePaymentDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))

    xmlNode.Attributes.setNamedItem(xmlFeeTypePaymentDoc.createAttribute("OPERATION")).Text = "CREATEFEETYPEPAYMENT"
    
    Set xmlFeePaymentNode = xmlFeeTypePaymentDoc.createElement("FEEPAYMENT")
    xmlNode.appendChild xmlFeePaymentNode
    
    Set xmlPaymentRecordNode = xmlFeeTypePaymentDoc.createElement("PAYMENTRECORD")
    xmlNode.appendChild xmlPaymentRecordNode
    
    'build the fee payment node
    xmlFeePaymentNode.setAttribute "APPLICATIONNUMBER", xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlFeePaymentNode.setAttribute "FEETYPE", strFeeType
    xmlFeePaymentNode.setAttribute "PAYMENTEVENT", strPaymentEvent
    xmlFeePaymentNode.setAttribute "AMOUNTPAID", strAmount
    xmlFeePaymentNode.setAttribute "DATEOFPAYMENT", Now
    xmlFeePaymentNode.setAttribute "REFUNDDATE", ""
    xmlFeePaymentNode.setAttribute "REFUNDAMOUNT", ""
    xmlFeePaymentNode.setAttribute "NOTES", strNotes
    xmlFeePaymentNode.setAttribute "FEETYPESEQUENCENUMBER", strFeeTypeSequenceNumber
    
    'build the payment record node
    xmlPaymentRecordNode.setAttribute "APPLICATIONNUMBER", xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlPaymentRecordNode.setAttribute "AMOUNT", strAmount
    xmlPaymentRecordNode.setAttribute "USERID", xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "@USERID").Text
    xmlPaymentRecordNode.setAttribute "PAYMENTMETHOD", strPaymentMethod
    
    BuildCreateFeeTypePaymentRequest = xmlFeeTypePaymentDoc.xml
        
BuildCreateFeeTypePaymentRequestExit:
    
    'destroy objects
    Set xmlFeePaymentNode = Nothing
    Set xmlPaymentRecordNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlFeeTypePaymentDoc = Nothing
    
Exit Function
    
BuildCreateFeeTypePaymentRequestErrorOccured:
    
    Err.Raise Err.Number, "BuildCreateFeeTypePaymentRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume BuildCreateFeeTypePaymentRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure CallPaymentProcessing
' Author: MO
' Date: 17/10/2002
' Purpose: Calls payment processing BO ans passed a relevant request
' Input parameters: RequestXML
' Output parameters: ResponseXML
'------------------------------------------------------------------------
Private Function CallPaymentProcessing(ByVal strPaymentProcessingRequest As String) As String

    On Error GoTo CallPaymentProcessingErrorOccured
    
    Dim objOmPayProcBO As Object
    
    Set objOmPayProcBO = CreateObject("omPayProc.PaymentProcessingBO")
    
    CallPaymentProcessing = objOmPayProcBO.omPayProcRequest(strPaymentProcessingRequest)
    
CallPaymentProcessingExit:
    
    'destroy objects
    Set objOmPayProcBO = Nothing
    
Exit Function
    
CallPaymentProcessingErrorOccured:
    
    Err.Raise Err.Number, "CallPaymentProcessing." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume CallPaymentProcessingExit
    
End Function

'------------------------------------------------------------------------
' Procedure BuildFindFeeTypeListRequest
' Author: MO
' Date: 17/10/2002
' Purpose: Builds the request to find a specific application fee type
' Input parameters: OmigaFormatRequest, FeeType
' Output parameters: request
'------------------------------------------------------------------------
Private Function BuildFindFeeTypeListRequest(ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal strFeeType As String) As String

    On Error GoTo BuildFindFeeTypeListRequestErrorOccured
    
    Dim xmlFeeTypeListDoc As FreeThreadedDOMDocument40
    Dim xmlApplicationFeeTypeNode As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMNode
    
    Set xmlFeeTypeListDoc = New FreeThreadedDOMDocument40
    xmlFeeTypeListDoc.validateOnParse = False
    xmlFeeTypeListDoc.setProperty "NewParser", True
    xmlFeeTypeListDoc.async = False
    
    ' clone REQUEST node
    Set xmlNode = xmlFeeTypeListDoc.appendChild(vxmlOmigaFormatRequestNode.cloneNode(False))

    xmlNode.Attributes.setNamedItem(xmlFeeTypeListDoc.createAttribute("OPERATION")).Text = "FINDFEETYPELIST"
    
    Set xmlApplicationFeeTypeNode = xmlFeeTypeListDoc.createElement("APPLICATIONFEETYPE")
    xmlNode.appendChild xmlApplicationFeeTypeNode
    
    'build the application fee type node
    xmlApplicationFeeTypeNode.setAttribute "APPLICATIONNUMBER", xmlGetMandatoryNode(vxmlOmigaFormatRequestNode, "APPLICATION/@APPLICATIONNUMBER").Text
    xmlApplicationFeeTypeNode.setAttribute "FEETYPE", strFeeType
    
    BuildFindFeeTypeListRequest = xmlFeeTypeListDoc.xml
        
BuildFindFeeTypeListRequestExit:
    
    'destroy objects
    Set xmlApplicationFeeTypeNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlFeeTypeListDoc = Nothing
    
Exit Function
    
BuildFindFeeTypeListRequestErrorOccured:
    
    Err.Raise Err.Number, "BuildFindFeeTypeListRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume BuildFindFeeTypeListRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure ExtractPostprocessingRulesRequest
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the request from preprocessing and extracts the 3 main nodes
'           of the operation
' Input parameters: vstrRulesResponse holding the response from the rules component
'                   vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Sub ExtractPostprocessingRulesRequest(ByVal strRulesResponse As String, ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ExtractPostprocessingRulesRequestErrorOccured
    
    Dim xmlRulesRequestDOM As FreeThreadedDOMDocument40
    
    Set xmlRulesRequestDOM = xmlLoad(strRulesResponse)
    
    'get the original request node
    Set vxmlRequestNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "ORIGINALREQUEST/REQUEST").cloneNode(True)
    
    'get the omiga formatted request node
    Set vxmlOmigaFormatRequestNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "OMIGAFORMATTEDREQUEST/REQUEST").cloneNode(True)
    
    'get the current response node
    Set vxmlResponseNode = xmlGetMandatoryNode(xmlRulesRequestDOM.documentElement, "CURRENTRESPONSE/RESPONSE").cloneNode(True)
    
ExtractPostprocessingRulesRequestExit:
    
    'destroy objects
    Set xmlRulesRequestDOM = Nothing
    
    Exit Sub
    
ExtractPostprocessingRulesRequestErrorOccured:

    Err.Raise Err.Number, "ExtractPostprocessingRulesRequest." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume ExtractPostprocessingRulesRequestExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildPostprocessingRulesResponse
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the 3 main nodes of the operation and converts them into
'           the string response needed for the Rules component
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildPostprocessingRulesResponse(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As String

    On Error GoTo BuildPostprocessingRulesResponseErrorOccured
    
    Dim xmlRulesResponseDOM As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOriginalRequestNode As IXMLDOMNode
    Dim xmlOmigaFormattedRequestNode As IXMLDOMNode
    Dim xmlCurrentResponseNode As IXMLDOMNode
    
    Set xmlRulesResponseDOM = New FreeThreadedDOMDocument40
    xmlRulesResponseDOM.validateOnParse = False
    xmlRulesResponseDOM.setProperty "NewParser", True
    'create the top level request node
    Set xmlRequestNode = xmlRulesResponseDOM.createElement("REQUEST")
    xmlRulesResponseDOM.appendChild xmlRequestNode
    
    'create the original request node
    Set xmlOriginalRequestNode = xmlRulesResponseDOM.createElement("ORIGINALREQUEST")
    xmlRequestNode.appendChild xmlOriginalRequestNode
    'append the original request node to it!
    xmlOriginalRequestNode.appendChild vxmlRequestNode.cloneNode(True)
    
    'create the omiga formatted request node
    Set xmlOmigaFormattedRequestNode = xmlRulesResponseDOM.createElement("OMIGAFORMATTEDREQUEST")
    xmlRequestNode.appendChild xmlOmigaFormattedRequestNode
    'append the omiga formatted request to it
    xmlOmigaFormattedRequestNode.appendChild vxmlOmigaFormatRequestNode.cloneNode(True)
    
    'create the current response node
    Set xmlCurrentResponseNode = xmlRulesResponseDOM.createElement("CURRENTRESPONSE")
    xmlRequestNode.appendChild xmlCurrentResponseNode
    'append the response node to it
    xmlCurrentResponseNode.appendChild vxmlResponseNode.cloneNode(True)
        
    'return the xml as a string
    BuildPostprocessingRulesResponse = xmlRulesResponseDOM.xml

BuildPostprocessingRulesResponseExit:
    
    'destroy objects
    Set xmlRulesResponseDOM = Nothing
    Set xmlRequestNode = Nothing
    Set xmlOriginalRequestNode = Nothing
    Set xmlOmigaFormattedRequestNode = Nothing
    Set xmlCurrentResponseNode = Nothing

    Exit Function
    
BuildPostprocessingRulesResponseErrorOccured:

    Err.Raise Err.Number, "BuildPostprocessingRulesResponse." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume BuildPostprocessingRulesResponseExit
    
End Function


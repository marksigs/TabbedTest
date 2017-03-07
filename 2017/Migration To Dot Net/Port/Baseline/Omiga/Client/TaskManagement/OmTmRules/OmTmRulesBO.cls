VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OmTmRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------Module header sample----------------------------------
'
' Class Module OmTmRulesBO
' File: OmTmRulesBO.cls
' Author: DCROSSLEY
' Date: 30/05/2002
' Purpose: Stage and task rules
'
'------------------------------------------------------------------------
'----------------------------------------------------------------------
'EPSOM Phase 2 Specific History:
'Prog   Date        AQR         Description
'LH     10/01/2007  EP2_647     New stage & task rules, improved error handling
'LH     11/01/2007  EP2_647     New stage & task rules
'LH     12/01/2007  EP2_647     Changed rule ref names of emp task rules in
'                               accordance with spec change.  Added further comments.
'LH     13/01/2007  EP2_647     Added task rule: ID_Check_Task_Rule & slight amendments after testing
'LH     08/02/2007  EP2_1010    Changed rules according to spec change.
'LH     15/02/2007  EP2_1432    Change to IsUnpackaged()
'LH     21/02/2007  EP2_1010    Changed rules according to spec change.
'OS     05/03/2007  EP2_1010    Changed IsUnPackagedDirect()
'OS     06/03/2007  EP2_1701    Created IsUnPackagedWithBroker()
'OS     07/03/2007  EP2_1736    Created IsKYCDocumentationRequired() and modified ID_Check_Task_Rule()
'DRC    12/03/2007  EP2_1490    Need to check for mismatches between task rule references in the
'                               in the code and the refeference on the task in the database
'AShaw  14/03/2007  EP2_1140    Non_Returning_Cust_OrTOE_Add_Cust_Task_Rule EP2_1140 - Add extra condition - IsAppType_TransferOfEquity.
'LDM    15/03/07    EP2_1749    fix IsUnPackaged()
'DRC    23/03/2007  EP2_1596    Completion Exception Rule - invert logic
'LDM    31/03/2007  EP2_1851    change the rule reference for rule 41(CancelReason_Expiry_Unpackaged_With_Broker_Task to Unpackaged_With_Broker_CancelReason_Not_Expiry_Task)
'PE     09/04/2007  EP2_1851    Changed IsUnPackagedDirect to check PACKAGERINDICATOR attribute.
'AShaw  09/04/2007  EP2_1262    Add code to support OrganisationTasks.
'DS     11/04/2007  EP2_1140    In IsTOE_WithAdditionalCustomer() function Commented condition that checks for if customer role type not 'Applicant'
'PE     20/04/2007  EP2_2471    2nd Applicant ID Confirmed Task Missing
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'The stage & task rules spec can be found in VSS Mars Docs:
'$/Epsom/2.Product Specific Details - Epsom/B.Parameters/VMS Phase 2 Stages & Tasks v5 291206.xls

Option Explicit

Implements ObjectControl

Private gstrRuleName As String
Private gobjContext As ObjectContext

Private Enum GlobalParameterType
    Amount = 1
    Percentage = 2
    BooleanType = 3
    StringType = 4
End Enum

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function RunRules(ByVal vxmlRequest As String) As String

    gobjContext.SetComplete

    On Error GoTo RunRulesVbErr

    Dim xmlRequestDoc As MSXML2.DOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlStageRequestNode As IXMLDOMNode
    
    Dim xmlResponseDoc As MSXML2.DOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTaskNode As IXMLDOMNode
    Dim xmlCaseStageResponseNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim sExceptionReason As String             ' BMIDS691
    Dim sExceptionReason_Text As String        ' BMIDS691
    
    Dim strStageId As String, _
        strActivityId As String, _
        strRuleRef As String
    
    Dim blnNextstageOk As Boolean
    Dim blnExceptionStage As Boolean
    
    Set xmlRequestDoc = New MSXML2.DOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    xmlRequestDoc.loadXML vxmlRequest
    
    ' ik_debug
    'xmlRequestDoc.save "c:\omiga4Trace\omTmRulesRequest.xml"
    
    Set xmlResponseDoc = New MSXML2.DOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    ' can be overriden
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseNode = xmlResponseDoc.appendChild(xmlResponseElem)
    
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("REQUEST")
    
    Set xmlApplicationRequestNode = xmlRequestDoc.selectSingleNode("REQUEST/APPLICATION")
    If xmlApplicationRequestNode Is Nothing Then
        RaiseError ("Application data not found")
    End If
    
    Set xmlStageRequestNode = xmlRequestDoc.selectSingleNode("REQUEST/STAGE")
    If xmlStageRequestNode Is Nothing Then
        RaiseError ("Stage data not found")
    End If
    
    'test stage rule against stage
    blnNextstageOk = TestStage(xmlStageRequestNode, xmlApplicationRequestNode, xmlResponseNode)
    
    'if stage run passed, run task rules for stage
    If blnNextstageOk Then
        
        Set xmlElem = xmlResponseNode.ownerDocument.createElement("CASESTAGE")
        xmlElem.setAttribute "STAGEID", xmlStageRequestNode.Attributes.getNamedItem("STAGEID").Text
        Set xmlCaseStageResponseNode = xmlResponseNode.appendChild(xmlElem)
        
        'add tasks to the case stage if the task rule criteria is met or if
        'the task is a default task
        TestTasks xmlStageRequestNode, xmlApplicationRequestNode, xmlCaseStageResponseNode
        
    End If

RunRulesExit:
    
    ' ik_debug
    'xmlResponseDoc.save "c:\omiga4Trace\omTmRulesResponse.xml"
    
    RunRules = xmlResponseDoc.xml
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlStageRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlCaseStageResponseNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    Exit Function
    
RunRulesVbErr:

    If Err.Number <> 0 Then
        RunRules = CreateErrorResponse
    End If

    xmlResponseElem.setAttribute "TYPE", "APPERR"       'MAR1021 GHun
    Set xmlElem = xmlResponseDoc.createElement("ERROR")
    Set xmlNode = xmlResponseElem.appendChild(xmlElem)

    Set xmlElem = xmlResponseDoc.createElement("NUMBER")        'MAR1021 GHun
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlResponseDoc.createElement("SOURCE")        'MAR1021 GHun
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("DESCRIPTION")       'MAR1021 GHun
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    
    GoTo RunRulesExit

End Function

Private Function TestStage( _
    ByVal vxmlStageNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) _
    As Boolean
    
    Dim blnNextstageOk As Boolean
    Dim strRuleRef As String

    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
            
    If Not vxmlStageNode.Attributes.getNamedItem("STAGERULEREFERENCE") Is Nothing Then
        strRuleRef = vxmlStageNode.getAttribute("STAGERULEREFERENCE")
    End If
    
    Select Case UCase(strRuleRef)
        Case "KFI_STAGE"
            blnNextstageOk = KFI_Stage_Rule(vxmlApplicationNode)
            
        Case "PSWOFFER_STAGE"
            blnNextstageOk = PSWOffer_Stage_Rule(vxmlApplicationNode)
            
        Case "DIP_STAGE"
            blnNextstageOk = DIP_Stage_Rule(vxmlApplicationNode)
            
        Case "DATACAPTURE_STAGE"
            blnNextstageOk = DataCapture_Stage_Rule(vxmlApplicationNode)
            
        Case "DATAVERIFICATION_STAGE"
            blnNextstageOk = DataVerification_Stage_Rule(vxmlApplicationNode)
            
        Case "OFFER_STAGE"
            blnNextstageOk = Offer_Stage_Rule(vxmlApplicationNode)
            
        Case "OFFERISSUED_STAGE"
            blnNextstageOk = OfferIssued_Stage_Rule(vxmlApplicationNode)
            
        Case "PRECOMPLETION_STAGE"
            blnNextstageOk = PreCompletion_Stage_Rule(vxmlApplicationNode)
            
        Case "COMPLETION_STAGE"
            blnNextstageOk = Completion_Stage_Rule(vxmlApplicationNode)
            
        Case "COMPLETION_EXCEPTION_STAGE"
            blnNextstageOk = Completion_Exception_Stage_Rule(vxmlApplicationNode)
            
        Case Else
            blnNextstageOk = True
    End Select
     
    TestStage = blnNextstageOk
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing

End Function

Private Sub TestTasks(ByVal vxmlStageNode As IXMLDOMNode, _
                      ByVal vxmlApplicationNode As IXMLDOMNode, _
                      ByVal vxmlResponseNode As IXMLDOMNode)
        
    Dim xmlTaskNode As IXMLDOMElement
    Dim strRuleRef As String
    
    For Each xmlTaskNode In vxmlStageNode.childNodes
    
        strRuleRef = "NO RULE"
    
        If Not xmlTaskNode.Attributes.getNamedItem("TASKRULEREFERENCE") Is Nothing Then
            strRuleRef = UCase(xmlTaskNode.getAttribute("TASKRULEREFERENCE"))
            gstrRuleName = strRuleRef 'used for error handling
        End If
        
        Select Case strRuleRef
            
            Case "SELFEMP_OR_CONTRACT_TASK"
                SelfEmp_Or_Contract_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                        
            Case "EMP_FULLSTATUS_TASK"
                Emp_FullStatus_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "SELFEMP_OR_CON_FULLSTATUS_TASK"
                SelfEmp_Or_Con_FullStatus_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "PORTED_LOAN_TASK"
                Ported_Loan_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "TOE_TASK"
                ToE_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "PRODUCT_SWITCH_TASK"
                Product_Switch_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADDBORR_TOE_ORADDBORRTOE_TASK"
                AddBorr_TOE_OrAddBorrTOE_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADDBORR_ORADDBORRTOE_TASK"
                AddBorr_OrAddBorrTOE_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "CREDIT_LIMIT_INC_TASK"
                Credit_Limit_Inc_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "BROKER_UNPACKAGED_TASK"
                Broker_Unpackaged_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "PACKAGED_TASK"
                Packaged_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "UNPACKAGED_TASK"
                Unpackaged_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "PAPER_CHANNEL_TASK"
                Paper_Channel_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "WEB_CHANNEL_TASK"
                Web_Channel_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "BROKER_TASK"
                Broker_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "CONTINUOUS_BORR_TASK"
                Continuous_Borr_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADDBORR_TOE_TASK"
                AddBorr_TOE_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NON_CONTINUOUS_BORR_TASK"
                Non_Continuous_Borr_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "EXCL_NON_CONTINUOUS_BORR_TASK"
                Excl_Non_Continuous_Borr_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "RETURNING_BORR_TASK"
                Returning_Borr_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "PACKAGED_CANCELREASON_NOT_EXPIRY_TASK"
                Packaged_CancelReason_Not_Expiry_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "UNPACKAGED_NOBROKER_CANCELREASON_NOT_EXPIRY_TASK"
                UnPackaged_NoBroker_CancelReason_Not_Expiry_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NO_SOLS_ORTHIRDPARTYSOLS_TASK"
                No_Sols_ORThirdpartySols_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "UNPACKAGED_NOBROKER_TASK"
                Unpackaged_NoBroker_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            'LDM 31/3/07 EP2_1851
'            Case "CANCELREASON_EXPIRY_UNPACKAGED_WITH_BROKER_TASK"
 '               CancelReason_Expiry_Unpackaged_With_Broker_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            Case "UNPACKAGED_WITH_BROKER_CANCELREASON_NOT_EXPIRY_TASK"
                Unpackaged_With_Broker_CancelReason_Not_Expiry_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "EXCL_NON_RETURNING_CUST_TASK"
                Excl_Non_Returning_Cust_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "CONTINUOUS_BORR_NON_GUARANTOR_ENG_WALES_TASK"
                Continuous_Borr_Non_Guarantor_Eng_Wales_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "CONTINUOUS_BORR_NON_GUARANTOR_NIRELAND_TASK"
                Continuous_Borr_Non_Guarantor_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "CONTINUOUS_BORR_NON_GUARANTOR_SCOTLAND_TASK"
                Continuous_Borr_Non_Guarantor_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "CONTINUOUS_BORR_GUARANTOR_ENG_WALES_TASK"
                Continuous_Borr_Guarantor_Eng_Wales_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "CONTINUOUS_BORR_GUARANTOR_NIRELAND_TASK"
                Continuous_Borr_Guarantor_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "CONTINUOUS_BORR_GUARANTOR_SCOTLAND_TASK"
                Continuous_Borr_Guarantor_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                  
            Case "NON_CONTINUOUS_BORR_NON_GUARANTOR_ENG_WALES_TASK"
                Non_Continuous_Borr_Non_Guarantor_Eng_Wales_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_CONTINUOUS_BORR_NON_GUARANTOR_NIRELAND_TASK"
                Non_Continuous_Borr_Non_Guarantor_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_CONTINUOUS_BORR_NON_GUARANTOR_SCOTLAND_TASK"
                Non_Continuous_Borr_Non_Guarantor_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_CONTINUOUS_BORR_GUARANTOR_ENG_WALES_TASK"
                Non_Continuous_Borr_Guarantor_Eng_Wales_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_CONTINUOUS_BORR_GUARANTOR_NIRELAND_TASK"
                Non_Continuous_Borr_Guarantor_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_CONTINUOUS_BORR_GUARANTOR_SCOTLAND_TASK"
                Non_Continuous_Borr_Guarantor_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode

            Case "NON_RETURNING_CUST_ORTOE_ADD_CUST_TASK"
                Non_Returning_Cust_OrTOE_Add_Cust_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NON_RETURNING_CUST_TASK"
                Non_Returning_Cust_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "RETURNING_CUST_TASK"
                Returning_Cust_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "RETURNING_CUST_INCL_FULL_UNDERWRITING_TASK"
                Returning_Cust_Incl_Full_Underwriting_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ID_CHECK_TASK"
                ID_Check_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "JOINT_APPS_TASK"
                Joint_Apps_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "UNPACKAGED_NON_CONTINUOUS_BORR_TASK"
                Unpackaged_Non_Continuous_Borr_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
                
            Case "UNPACKAGED_EMP_FULLSTATUS_TASK"
                Unpackaged_Emp_FullStatus_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "UNPACKAGED_SELFEMP_ORCON_FULLSTATUS_TASK"
                Unpackaged_SelfEmp_ORCon_FullStatus_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "THIRDPARTY_SOLS_TASK"
                Thirdparty_Sols_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NON_CONTINUOUS_BORR_ENG_WALES_TASK"
                Non_Continuous_Borr_Eng_Wales_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NON_CONTINUOUS_BORR_NIRELAND_TASK"
                Non_Continuous_Borr_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "NON_CONTINUOUS_BORR_SCOTLAND_TASK"
                Non_Continuous_Borr_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADD_BORR_THIRDPARTY_SOLS_ENGLAND_TASK"
                Add_Borr_Thirdparty_Sols_England_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADD_BORR_THIRDPARTY_SOLS_NIRELAND_TASK"
                Add_Borr_Thirdparty_Sols_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "ADD_BORR_THIRDPARTY_SOLS_SCOTLAND_TASK"
                Add_Borr_Thirdparty_Sols_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "TOE_ENGLAND_TASK"
                TOE_England_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "TOE_NIRELAND_TASK"
                TOE_NIreland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
            
            Case "TOE_SCOTLAND_TASK"
                TOE_Scotland_Task_Rule xmlTaskNode, vxmlApplicationNode, vxmlResponseNode
    
            Case "NO RULE"
                 'EP2_1490 no rule - is default task?
                If Not xmlTaskNode.Attributes.getNamedItem("DEFAULTFLAG") Is Nothing Then
                    If xmlTaskNode.getAttribute("DEFAULTFLAG") = "1" Then
                        AddDefaultCaseTask xmlTaskNode, vxmlResponseNode
                    End If
                End If
            Case Else
                'EP2_1490 Rule could not be identified
                CreateErrorResponse (vxmlResponseNode)
                       
        End Select
    
    Next

    Set xmlTaskNode = Nothing

End Sub

'*************************************************************************************
'Stage rules
'*************************************************************************************

Private Function KFI_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsWebChannel(vxmlApplicationNode) Then
        KFI_Stage_Rule = True
    End If
    
End Function

Private Function PSWOffer_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If IsAppType_ProductSwitch(vxmlApplicationNode) Then
        PSWOffer_Stage_Rule = True
    End If
    
End Function

Private Function DIP_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) Then
        DIP_Stage_Rule = True
    End If
    
End Function

Private Function DataCapture_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) Then
        DataCapture_Stage_Rule = True
    End If
    
End Function

Private Function DataVerification_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) Then
        DataVerification_Stage_Rule = True
    End If
    
End Function

Private Function Offer_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) Then
        Offer_Stage_Rule = True
    End If
    
End Function

Private Function OfferIssued_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) Then
        OfferIssued_Stage_Rule = True
    End If
    
End Function

Private Function PreCompletion_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) And _
    Not IsAppType_TransferOfEquity(vxmlApplicationNode) And _
    Not IsAppType_CreditLimitIncrease(vxmlApplicationNode) Then
        PreCompletion_Stage_Rule = True
    End If
    
End Function

Private Function Completion_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
            
    If Not IsAppType_ProductSwitch(vxmlApplicationNode) And _
    Not IsAppType_TransferOfEquity(vxmlApplicationNode) And _
    Not IsAppType_CreditLimitIncrease(vxmlApplicationNode) Then
        Completion_Stage_Rule = True
    End If
    
End Function

Private Function Completion_Exception_Stage_Rule(ByVal vxmlApplicationNode As IXMLDOMNode) As Boolean
    'EP2_1596 - Change this rule to opposite logic
    If IsAppType_ProductSwitch(vxmlApplicationNode) Or _
       IsAppType_TransferOfEquity(vxmlApplicationNode) Or _
       IsAppType_CreditLimitIncrease(vxmlApplicationNode) Then
        Completion_Exception_Stage_Rule = True
    End If
    
End Function

'*************************************************************************************
'Task rules
'*************************************************************************************

'Purpose: creates task if employment type = Self Employed or Contractor
'EP2_1262 - Use new AddOrganisationTask method.
Private Sub SelfEmp_Or_Contract_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlCustomerRole As IXMLDOMElement
    
    For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If IsSelfEmployed(xmlCustomerRole) Or _
        IsContractor(xmlCustomerRole) Then
            AddOrganisationTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
            Exit For
        End If
    Next
    
End Sub

'Purpose: creates task If employment type = Employed and Full Status
'EP2_1262 - Use new AddOrganisationTask method.
Private Sub Emp_FullStatus_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlCustomerRole As IXMLDOMElement
    
    For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If IsEmployed(xmlCustomerRole) And _
        IsIncomeFullStatus(vxmlApplicationNode) Then
            AddOrganisationTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
            Exit For
        End If
    Next
    
End Sub

'Purpose: If employment type = Self Emp/Contract and Full Status
'EP2_1262 - Use new AddOrganisationTask method.
Private Sub SelfEmp_Or_Con_FullStatus_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)

    Dim xmlCustomerRole As IXMLDOMElement
    
    For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If (IsSelfEmployed(xmlCustomerRole) Or _
        IsContractor(xmlCustomerRole)) And _
        IsIncomeFullStatus(vxmlApplicationNode) Then
            AddOrganisationTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
            Exit For
        End If
    Next
    
End Sub

'Purpose: creates task if ported loan
Private Sub Ported_Loan_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_PortedLoan(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is TOE
Private Sub ToE_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_TransferOfEquity(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is PSW
Private Sub Product_Switch_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ProductSwitch(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is additional borrowing or
'additional borrowing with TOE or TOE
Private Sub AddBorr_TOE_OrAddBorrTOE_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_AdditionalBorrowing(vxmlApplicationNode) Or _
    IsAppType_ABTransferOfEquity(vxmlApplicationNode) Or _
    IsAppType_TransferOfEquity(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is additional borrowing or
'additional borrowing with TOE
Private Sub AddBorr_OrAddBorrTOE_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_AdditionalBorrowing(vxmlApplicationNode) Or _
    IsAppType_ABTransferOfEquity(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is credit limit increase
Private Sub Credit_Limit_Inc_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_CreditLimitIncrease(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if submitted by Broker and Unpackaged
Private Sub Broker_Unpackaged_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If Not IsPackaged(vxmlApplicationNode) And _
    (IsBroker(vxmlApplicationNode) Or _
    IsUnPackagedWithBroker(vxmlApplicationNode)) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if packaged
Private Sub Packaged_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsPackaged(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if unpackaged
Private Sub Unpackaged_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)

    If IsUnPackaged(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If

End Sub

'Purpose: creates task if channel = paper
Private Sub Paper_Channel_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsPaperChannel(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if channel = web
Private Sub Web_Channel_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsWebChannel(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If

End Sub

'Purpose: creates task if broker involved (i.e. recorded in DC010)
Private Sub Broker_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlApplicationIntroducer As IXMLDOMElement
    
    If IsUnPackaged(vxmlApplicationNode) = False Or IsBroker(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type validation type = 'CB'
Private Sub Continuous_Borr_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is Additional Borrowing with ToE
Private Sub AddBorr_TOE_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ABTransferOfEquity(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrowing
Private Sub Non_Continuous_Borr_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is NOT non continuous borrowing
Private Sub Excl_Non_Continuous_Borr_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If Not IsAppType_NonContinuousBorrowing(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is returning borrower
Private Sub Returning_Borr_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ReturningBorrower(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if cancellation reason is not expiry and packaged
Private Sub Packaged_CancelReason_Not_Expiry_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsCancellationReasonNotExpiry(vxmlApplicationNode) And _
    IsPackaged(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if cancellation reason is not expiry and
'unpackaged and no broker involved
Private Sub UnPackaged_NoBroker_CancelReason_Not_Expiry_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsCancellationReasonNotExpiry(vxmlApplicationNode) And _
    IsUnPackagedDirect(vxmlApplicationNode) And _
    Not IsBroker(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if no panel solicitor details or third party sols details (i.e. empty!)
'OR If Third Party Solicitor present
Private Sub No_Sols_ORThirdpartySols_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim blLegalRep_ThirdPartyGUID As Boolean
    
    blLegalRep_ThirdPartyGUID = IsLegalRep_ThirdParty(vxmlApplicationNode)
    
    If blLegalRep_ThirdPartyGUID Or _
    (Not blLegalRep_ThirdPartyGUID And Not IsLegalRep_Directory(vxmlApplicationNode)) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if unpackaged and no broker involved
Private Sub Unpackaged_NoBroker_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsUnPackagedDirect(vxmlApplicationNode) And _
    Not IsBroker(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if cancellation reason is not expiry
'and unpackaged and broker involved
Private Sub Unpackaged_With_Broker_CancelReason_Not_Expiry_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsCancellationReasonNotExpiry(vxmlApplicationNode) And _
    IsUnPackagedDirect(vxmlApplicationNode) And _
    IsBroker(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is NOT non returning customer
Private Sub Excl_Non_Returning_Cust_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If Not IsAppType_NonReturningCustomer(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'England and Wales and no guarantor
Private Sub Continuous_Borr_Non_Guarantor_Eng_Wales_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'Northern Ireland and no guarantor
Private Sub Continuous_Borr_Non_Guarantor_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'Scotland and no guarantor
Private Sub Continuous_Borr_Non_Guarantor_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'England and Wales and guarantor
Private Sub Continuous_Borr_Guarantor_Eng_Wales_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'Northern Ireland and guarantor
Private Sub Continuous_Borr_Guarantor_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is continuous borrower and
'Scotland and guarantor
Private Sub Continuous_Borr_Guarantor_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'England and Wales and no guarantor
Private Sub Non_Continuous_Borr_Non_Guarantor_Eng_Wales_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Northern Ireland and no guarantor
Private Sub Non_Continuous_Borr_Non_Guarantor_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Scotland and no guarantor
Private Sub Non_Continuous_Borr_Non_Guarantor_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    Not IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'England and Wales and guarantor
Private Sub Non_Continuous_Borr_Guarantor_Eng_Wales_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Northern Ireland and guarantor
Private Sub Non_Continuous_Borr_Guarantor_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Scotland and guarantor
Private Sub Non_Continuous_Borr_Guarantor_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsGuarantor(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower or
'TOE with additional customer.
' EP2_1140 - Add extra condition - IsAppType_TransferOfEquity.
Private Sub Non_Returning_Cust_OrTOE_Add_Cust_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonReturningCustomer(vxmlApplicationNode) Or _
    (IsTOE_WithAdditionalCustomer(vxmlApplicationNode) And IsAppType_TransferOfEquity(vxmlApplicationNode)) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non returning customer
Private Sub Non_Returning_Cust_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonReturningCustomer(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is returning customer
Private Sub Returning_Cust_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_ReturningCustomer(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is returning customer or
'or if app type is returning borrower
Private Sub Returning_Cust_Incl_Full_Underwriting_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)

    If IsAppType_ReturningCustomer(vxmlApplicationNode) Or _
    IsAppType_ReturningBorrower(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: Creates task if app type is non returning customer
'OR TOE with additional customer
'AND (if KYC record exists and customer not validated)
'(i.e. a KYC check has been performed (for that customer)
'and the status returned is id not verified/satisfactory;
'or when the KYC check interface or task has failed
'or "Documentation required" response from AuthenticatePlus).
Private Sub ID_Check_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)

    Dim xmlCustomerRole As IXMLDOMNode
    
    If IsAppType_NonReturningCustomer(vxmlApplicationNode) Or _
    IsTOE_WithAdditionalCustomer(vxmlApplicationNode) Then
        For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
            If IsKYCAndCustomerNotValidated(vxmlApplicationNode, xmlCustomerRole) Or _
               IsKYCDocumentationRequired(vxmlApplicationNode, xmlCustomerRole) Then
                    AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
            End If
        Next
    End If
    
End Sub

'Purpose: creates task if a joint app (2 or more apps exist)
Private Sub Joint_Apps_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)

    If IsJointApp(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if unpackaged and non continuous borrower
Private Sub Unpackaged_Non_Continuous_Borr_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If Not IsPackaged(vxmlApplicationNode) And _
    IsAppType_NonContinuousBorrowing(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If

End Sub

'Purpose: creates task if unpackaged and employment type = Employed and Full Status
'EP2_1262 - Use new AddOrganisationTask method.
Private Sub Unpackaged_Emp_FullStatus_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlCustomerRole As IXMLDOMElement
        
    If IsUnPackagedDirect(vxmlApplicationNode) And _
    IsIncomeFullStatus(vxmlApplicationNode) Then
        For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
            If IsEmployed(xmlCustomerRole) Then
                AddOrganisationTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
                Exit For
            End If
        Next
    End If
    
End Sub

'Purpose: creates task if unpackaged and employment type = Self Emp/Contract and Full Status
'EP2_1262 - Use new AddOrganisationTask method.
Private Sub Unpackaged_SelfEmp_ORCon_FullStatus_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlCustomerRole As IXMLDOMElement
        
    If IsUnPackagedDirect(vxmlApplicationNode) And _
    IsIncomeFullStatus(vxmlApplicationNode) Then
        For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
            If IsSelfEmployed(xmlCustomerRole) Or _
            IsContractor(xmlCustomerRole) Then
                AddOrganisationTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerRole
                Exit For
            End If
        Next
    End If
    
End Sub

'Purpose: creates task if Third party solicitor record present
Private Sub Thirdparty_Sols_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsLegalRep_ThirdParty(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If

End Sub

'Purpose: creates task if app type is non continuous borrower and
'England and Wales
Private Sub Non_Continuous_Borr_Eng_Wales_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Northern Ireland
Private Sub Non_Continuous_Borr_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is non continuous borrower and
'Scotland
Private Sub Non_Continuous_Borr_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_NonContinuousBorrowing(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task (if app type is additional borrowing or
'app type is additional borrowing TOE) and England & Wales and
'Third party solicitor record present
Private Sub Add_Borr_Thirdparty_Sols_England_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If (IsAppType_AdditionalBorrowing(vxmlApplicationNode) Or _
    IsAppType_ABTransferOfEquity(vxmlApplicationNode)) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) And _
    IsLegalRep_ThirdParty(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task (if app type is additional borrowing or
'app type is additional borrowing TOE) and Northern Ireland and
'Third party solicitor record present
Private Sub Add_Borr_Thirdparty_Sols_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If (IsAppType_AdditionalBorrowing(vxmlApplicationNode) Or _
    IsAppType_ABTransferOfEquity(vxmlApplicationNode)) And _
    IsPropertyInNIreland(vxmlApplicationNode) And _
    IsLegalRep_ThirdParty(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task (if app type is additional borrowing or
'app type is additional borrowing TOE) and Scotland and
'Third party solicitor record present
Private Sub Add_Borr_Thirdparty_Sols_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If (IsAppType_AdditionalBorrowing(vxmlApplicationNode) Or _
    IsAppType_ABTransferOfEquity(vxmlApplicationNode)) And _
    IsPropertyInScotland(vxmlApplicationNode) And _
    IsLegalRep_ThirdParty(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is transfer of equity and England & Wales
Private Sub TOE_England_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_TransferOfEquity(vxmlApplicationNode) And _
    IsPropertyInEnglandAndWales(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is transfer of equity and Northern Ireland
Private Sub TOE_NIreland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_TransferOfEquity(vxmlApplicationNode) And _
    IsPropertyInNIreland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'Purpose: creates task if app type is transfer of equity and Scotland
Private Sub TOE_Scotland_Task_Rule( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    If IsAppType_TransferOfEquity(vxmlApplicationNode) And _
    IsPropertyInScotland(vxmlApplicationNode) Then
        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
    End If
    
End Sub

'*************************************************************************************
'Helper routines
'*************************************************************************************
Private Function IsTOE_WithAdditionalCustomer(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    Dim xmlCustomerRole As IXMLDOMElement
    Dim xmlCustomerVersion As IXMLDOMElement
    
    For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'if customer role type not 'Applicant'
        'DS EP2_1140 - Commented condition that checks for if customer role type not 'Applicant'
        'If xmlCustomerRole.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_A") Is Nothing Then
            Set xmlCustomerVersion = xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION")
            If Not xmlCustomerVersion Is Nothing Then
                If Not xmlCustomerVersion.Attributes.getNamedItem("NEWTOECUSTOMERIND") Is Nothing Then
                    If xmlCustomerVersion.Attributes.getNamedItem("NEWTOECUSTOMERIND").nodeValue = "1" Then
                        IsTOE_WithAdditionalCustomer = True
                        Exit For
                    End If
                End If
            End If
        'DS EP2_1140
        'End If
    Next

End Function

'If KYC record exists and customer not validated then return TRUE, otherwise FALSE
Private Function IsKYCAndCustomerNotValidated(ByVal vxmlApplicationNode As IXMLDOMElement, _
                                              ByRef xmlCustomerRole As IXMLDOMElement) As Boolean
    Dim xmlKYC As IXMLDOMElement
    
    If Not xmlCustomerRole Is Nothing Then
        Set xmlKYC = xmlCustomerRole.selectSingleNode("./KNOWYOURCUSTOMERCHECK")
        If Not xmlKYC Is Nothing Then
            If Not xmlKYC.Attributes.getNamedItem("SUCCESSINDICATOR") Is Nothing Then
                If xmlKYC.Attributes.getNamedItem("SUCCESSINDICATOR").nodeValue = "0" Then
                    IsKYCAndCustomerNotValidated = True
                End If
            End If
        Else
            'No KYC record exists
            IsKYCAndCustomerNotValidated = True
        End If
    End If

End Function

'Check for "Documentation required" response from AuthenticatePlus
Private Function IsKYCDocumentationRequired(ByVal vxmlApplicationNode As IXMLDOMElement, _
                                            ByRef xmlCustomerRole As IXMLDOMElement) As Boolean
    Dim xmlKYA As IXMLDOMElement
    
    If Not xmlCustomerRole Is Nothing Then
        Set xmlKYA = xmlCustomerRole.selectSingleNode("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS")
        If Not xmlKYA Is Nothing Then
            If Not xmlKYA.Attributes.getNamedItem("AUTHDECISION") Is Nothing Then
                If xmlKYA.Attributes.getNamedItem("AUTHDECISION").nodeValue = "NA00" Then
                    IsKYCDocumentationRequired = True
                End If
            End If
        End If
    End If

End Function

Private Function IsGuarantor(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    Dim xmlCustomerRole As IXMLDOMElement
    
    For Each xmlCustomerRole In vxmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_G='true']")
        IsGuarantor = True
        Exit For
    Next

End Function

Private Function IsJointApp(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    Dim xmlCustomerRole As IXMLDOMNodeList
    
    Set xmlCustomerRole = vxmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_G='true']")
    If xmlCustomerRole.length > 1 Then
        IsJointApp = True
    End If
    
End Function


Private Function IsPaperChannel(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlApplicationDataNode As IXMLDOMElement
    
    Set xmlApplicationDataNode = vxmlApplicationNode.selectSingleNode("./APPLICATIONDATA")
    
    If Not xmlApplicationDataNode Is Nothing Then
        If Not xmlApplicationDataNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
            If xmlApplicationDataNode.Attributes.getNamedItem("CHANNELID").nodeValue = GetGlobalParameter(vxmlApplicationNode, "PaperChannelId", Amount) Then
                IsPaperChannel = True
            End If
        End If
    End If

End Function

Private Function IsWebChannel(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlApplicationDataNode As IXMLDOMElement
    
    Set xmlApplicationDataNode = vxmlApplicationNode.selectSingleNode("./APPLICATIONDATA")
    
    If Not xmlApplicationDataNode Is Nothing Then
        If Not xmlApplicationDataNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
            If xmlApplicationDataNode.Attributes.getNamedItem("CHANNELID").nodeValue = GetGlobalParameter(vxmlApplicationNode, "WebChannelId", Amount) Then
                IsWebChannel = True
            End If
        End If
    End If
    
End Function
Private Function IsCancellationReasonNotExpiry(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlCaseStage As IXMLDOMElement
    
    Set xmlCaseStage = vxmlApplicationNode.selectSingleNode("./CASEACTIVITY/CASESTAGE")
    If Not xmlCaseStage Is Nothing Then
        'CE = Case Expired
        If xmlCaseStage.Attributes.getNamedItem("EXCEPTIONREASON_TYPE_CE") Is Nothing Then
            IsCancellationReasonNotExpiry = True
        End If
    End If
    
End Function

Private Function IsLegalRep_ThirdParty(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlLegalRep As IXMLDOMElement

    Set xmlLegalRep = vxmlApplicationNode.selectSingleNode("./APPLICATIONLEGALREP")
    If Not xmlLegalRep Is Nothing Then
        If Not xmlLegalRep.Attributes.getNamedItem("THIRDPARTYGUID") Is Nothing Then
            IsLegalRep_ThirdParty = True
        End If
    End If

End Function

Private Function IsLegalRep_Directory(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlLegalRep As IXMLDOMElement

    Set xmlLegalRep = vxmlApplicationNode.selectSingleNode("./APPLICATIONLEGALREP")
    If Not xmlLegalRep Is Nothing Then
        If Not xmlLegalRep.Attributes.getNamedItem("DIRECTORYGUID") Is Nothing Then
            IsLegalRep_Directory = True
        End If
    End If

End Function

Private Function IsPropertyInEnglandAndWales(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlProperty As IXMLDOMElement

    Set xmlProperty = vxmlApplicationNode.selectSingleNode("./NEWPROPERTY")
    If Not xmlProperty Is Nothing Then
        If Not xmlProperty.Attributes.getNamedItem("PROPERTYLOCATION_TYPE_E") Is Nothing Then
            IsPropertyInEnglandAndWales = True
        End If
    End If

End Function

Private Function IsPropertyInNIreland(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlProperty As IXMLDOMElement

    Set xmlProperty = vxmlApplicationNode.selectSingleNode("./NEWPROPERTY")
    If Not xmlProperty Is Nothing Then
        If Not xmlProperty.Attributes.getNamedItem("PROPERTYLOCATION_TYPE_I") Is Nothing Then
            IsPropertyInNIreland = True
        End If
    End If

End Function

Private Function IsPropertyInScotland(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlProperty As IXMLDOMElement

    Set xmlProperty = vxmlApplicationNode.selectSingleNode("./NEWPROPERTY")
    If Not xmlProperty Is Nothing Then
        If Not xmlProperty.Attributes.getNamedItem("PROPERTYLOCATION_TYPE_S") Is Nothing Then
            IsPropertyInScotland = True
        End If
    End If

End Function

Private Function IsBroker(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    Dim xmlApplicationIntroducer As IXMLDOMElement

    For Each xmlApplicationIntroducer In vxmlApplicationNode.selectNodes("APPLICATIONINTRODUCER")
        If Not xmlApplicationIntroducer Is Nothing Then
            If Not xmlApplicationIntroducer.Attributes.getNamedItem("ARFIRMID") Is Nothing Then
                IsBroker = True
                Exit For
            End If
        End If
    Next

End Function

Private Function IsPackaged(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
'IF ApplicationIntroducer.PrincipalFirmID record exists THEN
'  For each ApplicationIntroducer.PrincipalFirmID record
'      IF ANY record found has PrincipalFirm.PackagerIndicator = TRUE // RULE is TRUE the case is packaged (may also have a Broker)
      
    Dim xmlApplicationIntroducer As IXMLDOMElement
    Dim xmlPrincipalFirms As IXMLDOMNodeList
    Dim xmlPrincipalFirm As IXMLDOMNode
    
    For Each xmlApplicationIntroducer In vxmlApplicationNode.selectNodes("APPLICATIONINTRODUCER")
        Set xmlPrincipalFirms = xmlApplicationIntroducer.selectNodes("./PRINCIPALFIRM")
        For Each xmlPrincipalFirm In xmlPrincipalFirms
            If Not xmlPrincipalFirm.Attributes.getNamedItem("PACKAGERINDICATOR") Is Nothing Then
                If xmlPrincipalFirm.Attributes.getNamedItem("PACKAGERINDICATOR").nodeValue = "1" Then
                    IsPackaged = True
                    Exit For
                End If
            End If
        Next
    Next

End Function

Private Function IsUnPackaged(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
'if at least 1 ApplicationIntroducer.PrincipalFirmID record exists WHERE PrincipalFirm.PackagerIndicator = TRUE
'then the Case is Packaged else the case is unpackaged
'LDM 19/3/07 EP2_1749

    Dim xmlPrincipalFirms As IXMLDOMNodeList
    
    Set xmlPrincipalFirms = vxmlApplicationNode.selectNodes( _
    "APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='1']")

    IsUnPackaged = True
    If Not xmlPrincipalFirms Is Nothing Then
        If xmlPrincipalFirms.length > 0 Then
            IsUnPackaged = False  'Case has a packaged introducer so case is packaged.
        End If
    End If
    

End Function


Private Function IsUnPackagedDirect(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
'IF no ApplicationIntroducer.PrincipalFirmID record

    Dim xmlApplicationIntroducer As IXMLDOMElement
    Dim xmlPrincipalFirms As IXMLDOMNodeList
    Dim xmlPrincipalFirm As IXMLDOMNode
    
    IsUnPackagedDirect = True
    
    For Each xmlApplicationIntroducer In vxmlApplicationNode.selectNodes("APPLICATIONINTRODUCER")
        Set xmlPrincipalFirms = xmlApplicationIntroducer.selectNodes("./PRINCIPALFIRM[@PACKAGERINDICATOR='1']")
        For Each xmlPrincipalFirm In xmlPrincipalFirms
            IsUnPackagedDirect = False
            Exit For
        Next
    Next
    
End Function

Private Function IsUnPackagedWithBroker(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
'IF ApplicationIntroducer.PrincipalFirmID record with PACKAGERINDICATOR = False

    Dim xmlApplicationIntroducer As IXMLDOMElement
    Dim xmlPrincipalFirms As IXMLDOMNodeList
    Dim xmlPrincipalFirm As IXMLDOMNode
    
    IsUnPackagedWithBroker = False
    
    For Each xmlApplicationIntroducer In vxmlApplicationNode.selectNodes("APPLICATIONINTRODUCER")
        Set xmlPrincipalFirms = xmlApplicationIntroducer.selectNodes("./PRINCIPALFIRM[@PACKAGERINDICATOR='0']")
        For Each xmlPrincipalFirm In xmlPrincipalFirms
            IsUnPackagedWithBroker = True
            Exit Function
        Next
    Next
    
End Function


Private Function IsAppType_PortedLoan(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_NP") Is Nothing Then
        IsAppType_PortedLoan = True
    End If

End Function

Private Function IsAppType_NonReturningCustomer(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_NRC") Is Nothing Then
        IsAppType_NonReturningCustomer = True
    End If

End Function
Private Function IsAppType_TransferOfEquity(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_TOE") Is Nothing Then
        IsAppType_TransferOfEquity = True
    End If

End Function

Private Function IsAppType_ProductSwitch(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PSW") Is Nothing Then
        IsAppType_ProductSwitch = True
    End If

End Function

Private Function IsAppType_ContinuousBorrowing(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_CB") Is Nothing Then
        IsAppType_ContinuousBorrowing = True
    End If

End Function

Private Function IsAppType_ReturningBorrower(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_RET") Is Nothing Then
        IsAppType_ReturningBorrower = True
    End If

End Function

Private Function IsAppType_ReturningCustomer(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_RC") Is Nothing Then
        IsAppType_ReturningCustomer = True
    End If

End Function

Private Function IsAppType_NonContinuousBorrowing(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_NCB") Is Nothing Then
        IsAppType_NonContinuousBorrowing = True
    End If

End Function

Private Function IsAppType_AdditionalBorrowing(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_ABO") Is Nothing Then
        IsAppType_AdditionalBorrowing = True
    End If

End Function

Private Function IsAppType_ABTransferOfEquity(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_ABTOE") Is Nothing Then
        IsAppType_ABTransferOfEquity = True
    End If

End Function

Private Function IsAppType_CreditLimitIncrease(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_CLI") Is Nothing Then
        IsAppType_CreditLimitIncrease = True
    End If

End Function

Private Function IsPackager(ByVal vxmlApplicationNode As IXMLDOMElement)

    Dim xmlApplicationIntroducer As IXMLDOMElement

    For Each xmlApplicationIntroducer In vxmlApplicationNode.selectNodes("APPLICATIONINTRODUCER")
        If Not xmlApplicationIntroducer.selectSingleNode("./INTRODUCER[@INTRODUCERTYPE_TYPE_P='true']") Is Nothing Then
            IsPackager = True
            Exit For
        End If
    Next

End Function

Private Function IsIncomeFullStatus(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    
    If Not vxmlApplicationNode.Attributes.getNamedItem("APPLICATIONINCOMESTATUS_TYPE_FS") Is Nothing Then
        IsIncomeFullStatus = True
    End If
    
End Function


Private Sub AddTask( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMElement)
    
    Dim xmlCaseTaskElem As IXMLDOMElement

    Set xmlCaseTaskElem = vxmlCaseStageResponseNode.ownerDocument.createElement("CASETASK")
    xmlCaseTaskElem.setAttribute "TASKID", vxmlTaskNode.getAttribute("TASKID")
    xmlCaseTaskElem.setAttribute "CASETASKNAME", vxmlTaskNode.getAttribute("TASKNAME")
    
    vxmlCaseStageResponseNode.appendChild xmlCaseTaskElem
    
    Set xmlCaseTaskElem = Nothing

End Sub

Private Sub AddCustomerTask( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMElement, _
    ByVal vxmlCustomerNode As IXMLDOMElement)
    
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlCustVerElem As IXMLDOMElement
    Dim strName As String

    Set xmlCaseTaskElem = vxmlCaseStageResponseNode.ownerDocument.createElement("CASETASK")
    xmlCaseTaskElem.setAttribute "TASKID", vxmlTaskNode.getAttribute("TASKID")
    xmlCaseTaskElem.setAttribute "CUSTOMERNUMBER", vxmlCustomerNode.getAttribute("CUSTOMERNUMBER")
    
    If vxmlCustomerNode.selectSingleNode("CUSTOMERVERSION[@FIRSTFORENAME or @SURNAME]") Is Nothing Then
    
        xmlCaseTaskElem.setAttribute "CASETASKNAME", vxmlTaskNode.getAttribute("TASKNAME")
        
    Else
    
        Set xmlCustVerElem = vxmlCustomerNode.selectSingleNode("CUSTOMERVERSION")
        
        If Not xmlCustVerElem.Attributes.getNamedItem("FIRSTFORENAME") Is Nothing Then
            strName = xmlCustVerElem.getAttribute("FIRSTFORENAME") & " "
        End If
        
        If Not xmlCustVerElem.Attributes.getNamedItem("SURNAME") Is Nothing Then
            strName = strName & xmlCustVerElem.getAttribute("SURNAME")
        End If

        xmlCaseTaskElem.setAttribute _
            "CASETASKNAME", _
            vxmlTaskNode.getAttribute("TASKNAME") & _
            " for " & strName
            
    End If
    
    vxmlCaseStageResponseNode.appendChild xmlCaseTaskElem
    
    Set xmlCaseTaskElem = Nothing

End Sub

Private Sub AddOrganisationTask( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMElement, _
    ByVal vxmlCustomerNode As IXMLDOMElement)
    
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlCustVerElem As IXMLDOMElement
    Dim xmlEmployerElem As IXMLDOMElement
    Dim xmlTPElem As IXMLDOMElement
    Dim strName As String
    Dim strCompanyName As String

    ' Create new task
    Set xmlCaseTaskElem = vxmlCaseStageResponseNode.ownerDocument.createElement("CASETASK")
    
    ' Add TaskID and CustomerID
    xmlCaseTaskElem.setAttribute "TASKID", vxmlTaskNode.getAttribute("TASKID")
    xmlCaseTaskElem.setAttribute "CUSTOMERIDENTIFIER", vxmlCustomerNode.getAttribute("CUSTOMERNUMBER")
     
    ' Add EmployerSeqNumber.
    Set xmlEmployerElem = vxmlCustomerNode.selectSingleNode("CUSTOMERVERSION/EMPLOYMENT")
    xmlCaseTaskElem.setAttribute "CONTEXT", xmlEmployerElem.getAttribute("EMPLOYMENTSEQUENCENUMBER")
    
    ' Get the strCompanyName from either DirectoryGUId or TPGUID (ONLY 1 will be present)
    If xmlEmployerElem.getAttribute("THIRDPARTYGUID") <> "" Then
        Set xmlTPElem = xmlEmployerElem.selectSingleNode("THIRDPARTY")
        strCompanyName = xmlTPElem.getAttribute("COMPANYNAME")
    ElseIf xmlEmployerElem.getAttribute("DIRECTORYGUID") <> "" Then
        Set xmlTPElem = xmlEmployerElem.selectSingleNode("NAMEANDADDRESSDIRECTORY")
        strCompanyName = xmlTPElem.getAttribute("COMPANYNAME")
    End If
    
    ' Now create the full CaseTaskName.
    ' Get node to obtain Customer Name.
    Set xmlCustVerElem = vxmlCustomerNode.selectSingleNode("CUSTOMERVERSION")
    ' Add Forename if present
    If Not xmlCustVerElem.Attributes.getNamedItem("FIRSTFORENAME") Is Nothing Then
        strName = xmlCustVerElem.getAttribute("FIRSTFORENAME") & " "
    End If
    ' Add Surname if present
    If Not xmlCustVerElem.Attributes.getNamedItem("SURNAME") Is Nothing Then
         strName = strName & xmlCustVerElem.getAttribute("SURNAME")
    End If
    
    ' CaseTask name = "Company Task for customer"
    xmlCaseTaskElem.setAttribute "CASETASKNAME", strCompanyName & " " & vxmlTaskNode.getAttribute("TASKNAME") & " for " & strName
       
       
    ' Flag to Confirm Applicant details have been added.
    xmlCaseTaskElem.setAttribute "APPLICANTDETAILSADDED", "1"
    
    ' Add Task to full list.
    vxmlCaseStageResponseNode.appendChild xmlCaseTaskElem
    
    ' Clear the working storage node.
    Set xmlCaseTaskElem = Nothing
    
End Sub


Private Function IsSelfCert(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean
    If Not vxmlApplicationNode.Attributes.getNamedItem("SELFCERTIND") Is Nothing Then
        If vxmlApplicationNode.getAttribute("SELFCERTIND") = "1" Then
            IsSelfCert = True
        End If
    End If
End Function

Private Function IsNonConforming(ByVal vxmlApplicationNode As IXMLDOMElement) As Boolean

    If Not vxmlApplicationNode.selectSingleNode("QUOTATION") Is Nothing Then
        If vxmlApplicationNode.selectNodes("QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/SPECIALGROUP[@GROUPTYPE='Prime' or @GROUPTYPE='Near Prime' or @GROUPTYPE='Super -Light']").length <> _
           vxmlApplicationNode.selectNodes("QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT").length _
        Then
            IsNonConforming = True
        End If
    End If

End Function

Private Function IsEmployed(xmlCustomerRole As IXMLDOMElement) As Boolean

    If Not xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@EMPLOYMENTSTATUS_TYPE_EMP='true']") Is Nothing Then
        IsEmployed = True
    End If

End Function

Private Function IsSelfEmployed(xmlCustomerRole As IXMLDOMElement) As Boolean

    If Not xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@EMPLOYMENTSTATUS_TYPE_SELF='true']") Is Nothing Then
        IsSelfEmployed = True
    End If
    
End Function

Private Function IsContractor(xmlCustomerRole As IXMLDOMElement) As Boolean

    If Not xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then
        IsContractor = True
    End If

End Function


Private Sub AddDefaultCaseTask( _
    ByVal vxmlTaskNode As IXMLDOMElement, _
    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
    
    Dim xmlCaseTaskElem As IXMLDOMElement
                        
    Set xmlCaseTaskElem = vxmlCaseStageResponseNode.ownerDocument.createElement("CASETASK")
    xmlCaseTaskElem.setAttribute "TASKID", vxmlTaskNode.getAttribute("TASKID")
    vxmlCaseStageResponseNode.appendChild xmlCaseTaskElem
    
    Set xmlCaseTaskElem = Nothing
    
End Sub


Private Function CreateErrorResponse(Optional xmlResponse As DOMDocument40) As String

    Dim xmlResponseDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = True
    xmlResponseDoc.async = True
    xmlResponseDoc.setProperty "NewParser", True

    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
    
    If Err.Number > vbObjectError Then
        xmlElem.setAttribute "TYPE", "APPERR"
    Else
        xmlElem.setAttribute "TYPE", "SYSERR"
    End If
    Set xmlNode = xmlResponseDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("ERROR")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("NUMBER")
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("SOURCE")
    xmlElem.Text = "omTMRulesBO"
    
    If Len(gstrRuleName) > 0 Then
        xmlElem.Text = xmlElem.Text & "." & gstrRuleName & "_Rule"
    End If
    
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("DESCRIPTION")
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    
    If Not xmlResponse Is Nothing Then
        Set xmlResponse = xmlResponseDoc
    End If
    
    CreateErrorResponse = xmlResponseDoc.xml
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing
    
End Function
    
Private Function GetGlobalParameter(ByVal xmlApplicationNode As IXMLDOMNode, _
                                    ByVal sName As String, _
                                    ByVal eGlobalParameterType As GlobalParameterType) As String
   
    Dim xmlNode As IXMLDOMNode
    Dim sParamValue As String
    Dim sGlobalParameterType As String
    
'    Private Enum GlobalParameterType
'        Amount = 1
'        Percentage = 2
'        BooleanType = 3
'        StringType = 4
'    End Enum

    Select Case eGlobalParameterType
        Case Amount
            sGlobalParameterType = "AMOUNT"
            
        Case Percentage
            sGlobalParameterType = "PERCENTAGE"
            
        Case BooleanType
            sGlobalParameterType = "BOOLEAN"
            
        Case StringType
            sGlobalParameterType = "STRING"
    End Select
    
    Set xmlNode = xmlApplicationNode.selectSingleNode("./GLOBALPARAMETER[@NAME='" & sName & "']")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem(sGlobalParameterType) Is Nothing Then
            sParamValue = xmlNode.Attributes.getNamedItem(sGlobalParameterType).nodeValue
        End If
    Else
        RaiseError "Globalparameter '" & sName & "' not found"
    End If
        
    GetGlobalParameter = sParamValue
    
End Function

Private Sub RaiseError(ByVal vstrDescription)

    Err.Raise vbObjectError + 512 + 999, "omTMRulesBO", vstrDescription

End Sub

'Private Sub IDCheckRule( _
'    ByVal vxmlTaskNode As IXMLDOMElement, _
'    ByVal vxmlApplicationNode As IXMLDOMNode, _
'    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
'
'    Dim xmlCustomerNode As IXMLDOMElement
'
'    For Each xmlCustomerNode In vxmlApplicationNode.selectNodes("CUSTOMERROLE")
'
'        If Not xmlCustomerNode.selectSingleNode("KNOWYOURCUSTOMERCHECK[1]/KNOWYOURCUSTOMERAUTHENTICATERESULTS[@AUTHDECISION='NA00' or @AUTHDECISION='RA00']") Is Nothing Then
'
'            AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerNode
'
'        End If
'
'    Next
'
'    Set xmlCustomerNode = Nothing
'
'End Sub
'
'Private Sub AccountantCheckRule( _
'    ByVal vxmlTaskNode As IXMLDOMElement, _
'    ByVal vxmlApplicationNode As IXMLDOMElement, _
'    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
'
'    Dim xmlCustomerNode As IXMLDOMElement
'
'    For Each xmlCustomerNode In vxmlApplicationNode.selectNodes("CUSTOMERROLE")
'
'        If xmlCustomerNode.selectNodes("CUSTOMERVERSION/EMPLOYMENT[not(@DATELEFTORCEASEDTRADING) and @EMPLOYMENTSTATUS_TYPE_SELF]").length <> 0 _
'        And IsNonConforming(vxmlApplicationNode) _
'        Then
'
'            AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerNode
'
'        End If
'
'    Next
'
'    Set xmlCustomerNode = Nothing
'
'End Sub
'
'Private Sub OccupationCheckRule( _
'    ByVal vxmlTaskNode As IXMLDOMElement, _
'    ByVal vxmlApplicationNode As IXMLDOMElement, _
'    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
'
'    Dim xmlCustomerNode As IXMLDOMElement
'
'    For Each xmlCustomerNode In vxmlApplicationNode.selectNodes("CUSTOMERROLE")
'
'        If IsSelfCert(vxmlApplicationNode) And IsNonConforming(vxmlApplicationNode) Then
'
'            AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerNode
'
'        Else
'
'            If IsSelfCert(vxmlApplicationNode) _
'            And xmlCustomerNode.selectNodes("CUSTOMERVERSION/EMPLOYMENT[not(@DATELEFTORCEASEDTRADING) and @EMPLOYMENTSTATUS_TYPE_EMP]").length <> 0 _
'            Then
'
'                AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerNode
'
'            Else
'
'                If xmlCustomerNode.selectNodes("CUSTOMERVERSION/EMPLOYMENT[not(@DATELEFTORCEASEDTRADING) and @EMPLOYMENTSTATUS_TYPE_SELF]").length <> 0 _
'                And IsNonConforming(vxmlApplicationNode) _
'                Then
'
'                    AddCustomerTask vxmlTaskNode, vxmlCaseStageResponseNode, xmlCustomerNode
'
'                End If
'
'            End If
'
'        End If
'
'    Next
'
'    Set xmlCustomerNode = Nothing
'
'End Sub
'
'Private Sub DeedOfConsentRule( _
'    ByVal vxmlTaskNode As IXMLDOMElement, _
'    ByVal vxmlApplicationNode As IXMLDOMNode, _
'    ByVal vxmlCaseStageResponseNode As IXMLDOMNode)
'
'    If vxmlApplicationNode.selectNodes("OTHERRESIDENT").length <> 0 Then
'        AddTask vxmlTaskNode, vxmlCaseStageResponseNode
'    End If
'
'End Sub



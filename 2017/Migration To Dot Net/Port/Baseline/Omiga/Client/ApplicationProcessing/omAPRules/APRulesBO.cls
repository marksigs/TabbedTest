VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "APRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      APRulesBO.cls.cls
'Copyright:     Copyright © 2000 Marlborough Stirling

'Description:   Application Processing Rules
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NoTransactions
'------------------------------------------------------------------------------------------
'History:
'
' Prog      Date     Description
' PSC       04/12/01 SYS3314 Amend current and previous landlord's reference rule to
'                            take into account the fact that 'Yes' has been entered
' PSC       04/12/01 SYS3314 Amend current and previous lender's reference rule to
'                            take into account the fact that 'Yes' has been entered
'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog      Date        Description
' GD        03/10/2002  BMIDS00313 - APWP2 - Added WriteConditionsErrorLog,ValidateAndResolveConditionText,ResolveDataRules
'                                               Changed ConditionsRules
' GD        21/10/2002  BMIDS00650 New private method ResolveConditionRule(taken processing out of ValidateAndResolveConditionText())
'                       Added New Rule for Condition 703 (with embedded tag <<BTLGT65>>)
' GD        06/11/2002  BMIDS00759  Changed ResolveConditionRule and ConditionsRules so null attribute is handled in BTLGT65
' GD        11/11/2002  BMIDS00913 - missing code to check for BTL(related to change done for BMIDS00759)
'------------------------------------------------------------------------------------------
'Mars History:
'
' Prog      Date        Description
' GHun      22/02/2006  MAR1316 Commented out most of the rules as they are obsolete
'------------------------------------------------------------------------------------------
'Epsom History:
'
' Prog      Date        Description
' PE        29/08/2006  EP1099 - Removed conditions 701 & 202. Added conditions 3, 4, 22, 46, 47, 57, 62, 72
' PE        15/09/2006  EP1144 - Removed conditions 22, 46, 57, 62 and 72. Added condition 10.
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext

Const glngErrBase As Long = vbObjectError + 512

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function RunRules(ByVal strXmlIn As String)
    
    On Error GoTo RunRulesVbErr
    
    Const strFunctionName As String = "RunRules"

    Dim xmlIn As FreeThreadedDOMDocument40, _
        xmlOut As FreeThreadedDOMDocument40, _
        xmlElem As IXMLDOMElement, _
        xmlNode As IXMLDOMNode, _
        xmlRequestNode As IXMLDOMNode, _
        xmlResponseElem As IXMLDOMElement, _
        xmlResponseNode As IXMLDOMNode, _
        xmlApplicationNode  As IXMLDOMNode

    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strOperation As String
        
    Dim bParseOK As Boolean
    
    gobjContext.SetComplete
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    xmlIn.async = False
    bParseOK = xmlIn.loadXML(strXmlIn)
    
    If Not bParseOK Then
        Err.Raise glngErrBase + 1, strFunctionName, xmlParseErr(xmlIn)
    End If
        
    Set xmlRequestNode = xmlIn.selectSingleNode("REQUEST")
    Set xmlApplicationNode = xmlIn.selectSingleNode("REQUEST/APPLICATION")
    
    If xmlApplicationNode Is Nothing Then
        Err.Raise glngErrBase + 2, strFunctionName, "REQUEST missing APPLICATION node"
    End If
    
    If Not xmlAttributeExists(xmlRequestNode, "OPERATION") Then
        Err.Raise glngErrBase + 3, strFunctionName, "REQUEST missing OPERATION attribute"
    End If
    
    strOperation = xmlGetAttributeText(xmlRequestNode, "OPERATION")
    
    Select Case strOperation
    
        Case "RunValuationReportRules"
            ValuationReportRules xmlApplicationNode, xmlResponseNode
            
        Case "RunCurrEmployersRefRules"
            CurrEmployersRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunPrevEmployersRefRules"
            PrevEmployersRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunCurrLendersRefRules"
            CurrLendersRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunPrevLendersRefRules"
            PrevLendersRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunCurrLandlordRefRules"
            CurrLandlordRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunPrevLandlordRefRules"
            PrevLandlordRefRules xmlRequestNode, xmlResponseNode
            
        Case "RunROTRules"
            ROTRules xmlRequestNode, xmlResponseNode
    
        ' JR SYS2025 25/04/01
        Case "RunConditionRules"
            ConditionsRules xmlRequestNode, xmlResponseNode
            
        ' GD BMIDS0313
        Case "RunResolveDataRules"
            ResolveDataRules xmlRequestNode, xmlResponseNode
            
    End Select
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
RunRulesExit:
    
    RunRules = xmlOut.xml
    Debug.Print RunRules
        
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing

    Exit Function
    
RunRulesVbErr:
    xmlResponseElem.setAttribute "TYPE", "ERROR"
    
    Set xmlElem = xmlOut.createElement("ERROR")
    Set xmlNode = xmlResponseElem.appendChild(xmlElem)
    
    Set xmlElem = xmlOut.createElement("NUMBER")
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlOut.createElement("SOURCE")
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlOut.createElement("DESCRIPTION")
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    
    GoTo RunRulesExit

End Function


Private Sub ValuationReportRules( _
    ByVal vxmlApplicationNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrRuleResponseTag As String = "VALUATIONREPORTRULE"
    
    Dim xmlReportNode As IXMLDOMNode
    
    Dim lngTotalRetention As Long, _
        lngAmountRequested As Long
        
    Dim strSearch As String
    
    Set xmlReportNode = _
        vxmlApplicationNode.selectSingleNode("VALUERINSTRUCTION/VALUATIONREPORT")
    
    If Not xmlReportNode Is Nothing Then
        
        lngTotalRetention = xmlGetAttributeAsLong(xmlReportNode, "RETENTIONWORKS") + _
                            xmlGetAttributeAsLong(xmlReportNode, "RETENTIONSROADS")
                            
        If lngTotalRetention = 0 Then
            
            If xmlGetAttributeAsLong(xmlReportNode, "PRESENTVALUATION") > 0 And _
                xmlGetAttributeAsLong(xmlReportNode, "POSTWORKSVALUATION") > 0 _
            Then
                
                AddResponse _
                    vxmlResponseNode, _
                    cstrRuleResponseTag, _
                    "There are no Retentions for the Post Works Valuation", _
                    xmlGetAttributeText(xmlReportNode, "INSTRUCTIONSEQUENCENO")
            
            End If
        
        Else
                
            strSearch = _
                "QUOTATION[@QUOTATIONNUMBER='" & _
                xmlGetAttributeText(vxmlApplicationNode, "ACCEPTEDQUOTENUMBER") & _
                "']/MORTGAGESUBQUOTE/@AMOUNTREQUESTED"
                
            lngAmountRequested = xmlGetNodeAsLong(vxmlApplicationNode, strSearch)
                
            If lngTotalRetention > lngAmountRequested Then
                
                AddResponse _
                    vxmlResponseNode, _
                    cstrRuleResponseTag, _
                    "The total amount of Retention exceeds the Loan Requested", _
                    xmlGetAttributeText(xmlReportNode, "INSTRUCTIONSEQUENCENO")
            
            End If
        
        End If
        
    End If
                                    
End Sub

Private Sub CurrEmployersRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

'MAR1316 GHun
'    Const cstrFunctionName As String = "CurrEmployersRefRules"
'    Const cstrRuleResponseTag As String = "CURRENTEMPLOYERREFRULE"
'
'    Dim xmlEmploymentNode As IXMLDOMNode, _
'        xmlEmployerRefNode As IXMLDOMNode, _
'        xmlEarnedIncomeNode As IXMLDOMNode
'
'    Dim strSearch As String, _
'        strCustNo As String, _
'        strCustVersNo As String, _
'        strEmpSeqNo As String
'
'    Dim dtDateStarted As Date, _
'        dtRefDateStarted As Date
'
'    Dim intToleranceParam As Integer, _
'        intPayFreq As Integer
'
'    Dim lngIncome As Long, _
'        lngAnnualIncome As Long, _
'        lngToleranceAmount As Long
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERNUMBER") Then
'        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERNUMBER attribute"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERVERSIONNUMBER") Then
'        strCustVersNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
'    Else
'        strCustVersNo = "1"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER") Then
'        strEmpSeqNo = xmlGetAttributeText(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing EMPLOYMENTSEQUENCENUMBER attribute"
'    End If
'
'    strSearch = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch = strSearch & "[@CUSTOMERNUMBER='" & strCustNo & "' and @CUSTOMERVERSIONNUMBER='" & strCustVersNo & "']"
'    strSearch = strSearch & "/CUSTOMER/INCOME/EMPLOYMENT"
'    strSearch = strSearch & "[@EMPLOYMENTSEQUENCENUMBER='" & strEmpSeqNo & "']"
'
'    Set xmlEmploymentNode = vxmlRequestNode.selectSingleNode(strSearch)
'
'    If xmlEmploymentNode Is Nothing Then
'        Exit Sub
'    End If
'
'    Set xmlEmployerRefNode = xmlEmploymentNode.selectSingleNode("CURRENTEMPLOYERREF")
'
'    If xmlEmployerRefNode Is Nothing Then
'        Exit Sub
'    End If
'
'    If xmlAttributeExists(xmlEmployerRefNode, "EMPLOYMENTTYPE") Then
'        If xmlGetAttributeText(xmlEmployerRefNode, "EMPLOYMENTTYPE") <> _
'            xmlGetAttributeText(xmlEmploymentNode, "EMPLOYMENTTYPE") _
'        Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Employment Type inconsistent", _
'                strEmpSeqNo
'        End If
'    End If
'
'    If xmlAttributeExists(xmlEmployerRefNode, "DATESTARTED") Then
'
'        intToleranceParam = _
'            xmlGetNodeAsInteger(vxmlRequestNode, "GLOBALPARAMETER[@NAME='TimeInEmploymentTolerance']/@AMOUNT")
'
'        dtDateStarted = xmlGetAttributeAsDate(xmlEmploymentNode, "DATESTARTEDORESTABLISHED")
'        dtRefDateStarted = xmlGetAttributeAsDate(xmlEmployerRefNode, "DATESTARTED")
'
'        If Abs(DateDiff("d", dtDateStarted, dtRefDateStarted)) > intToleranceParam Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Declared time in Employment falls outside of accepted tolerance levels", _
'                strEmpSeqNo
'        End If
'
'    End If
'
'    If xmlAttributeExists(xmlEmployerRefNode, "GROSSBASICINCOME") Then
'
'        intToleranceParam = _
'            xmlGetNodeAsInteger(vxmlRequestNode, "GLOBALPARAMETER[@NAME='BasicIncomeTolerance']/@PERCENTAGE")
'
'        Set xmlEarnedIncomeNode = _
'            xmlEmploymentNode.selectSingleNode("EMPLOYEDDETAIL/EARNEDINCOME[@EARNEDINCOMETYPE='1']")
'
'        If Not xmlEarnedIncomeNode Is Nothing Then
'            intPayFreq = xmlGetAttributeAsInteger(xmlEarnedIncomeNode, "PAYMENTFREQUENCYTYPE")
'            lngIncome = xmlGetAttributeAsLong(xmlEarnedIncomeNode, "EARNEDINCOMEAMOUNT")
'            lngAnnualIncome = lngIncome * intPayFreq
'        End If
'
'        Set xmlEarnedIncomeNode = _
'            xmlEmploymentNode.selectSingleNode("EMPLOYEDDETAIL/EARNEDINCOME[@EARNEDINCOMETYPE='6']")
'
'        If Not xmlEarnedIncomeNode Is Nothing Then
'            intPayFreq = xmlGetAttributeAsInteger(xmlEarnedIncomeNode, "PAYMENTFREQUENCYTYPE")
'            lngIncome = xmlGetAttributeAsLong(xmlEarnedIncomeNode, "EARNEDINCOMEAMOUNT")
'            lngAnnualIncome = lngAnnualIncome + (lngIncome * intPayFreq)
'        End If
'
'        lngToleranceAmount = (lngAnnualIncome * intToleranceParam) / 100
'
'        lngIncome = xmlGetAttributeAsLong(xmlEmployerRefNode, "GROSSBASICINCOME")
'
'        If Abs(lngAnnualIncome - lngIncome) > lngToleranceAmount Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Declared Basic Income falls outside of accepted tolerance levels", _
'                strEmpSeqNo
'        End If
'
'    End If
'
'    If xmlGetAttributeAsBoolean(xmlEmployerRefNode, "SIGNEDSTAMPIND") = False Then
'        AddResponse _
'            vxmlResponseNode, _
'            cstrRuleResponseTag, _
'            "Employer's Reference not Signed and Stamped", _
'            strEmpSeqNo
'    End If
'MAR1316 End

End Sub

Private Sub PrevEmployersRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "PrevEmployersRefRules"
'    Const cstrRuleResponseTag As String = "PREVIOUSEMPLOYERREFRULE"
'
'    Dim xmlEmploymentNode As IXMLDOMNode, _
'        xmlEmployerRefNode As IXMLDOMNode
'
'    Dim strSearch As String, _
'        strCustNo As String, _
'        strCustVersNo As String, _
'        strEmpSeqNo As String
'
'    Dim dtDateStarted As Date, _
'        dtRefDateStarted As Date
'
'    Dim intToleranceParam As Integer
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERNUMBER") Then
'        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERNUMBER attribute"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERVERSIONNUMBER") Then
'        strCustVersNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
'    Else
'        strCustVersNo = "1"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER") Then
'        strEmpSeqNo = xmlGetAttributeText(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing EMPLOYMENTSEQUENCENUMBER attribute"
'    End If
'
'    strSearch = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch = strSearch & "[@CUSTOMERNUMBER='" & strCustNo & "' and @CUSTOMERVERSIONNUMBER='" & strCustVersNo & "']"
'    strSearch = strSearch & "/CUSTOMER/INCOME/EMPLOYMENT"
'    strSearch = strSearch & "[@EMPLOYMENTSEQUENCENUMBER='" & strEmpSeqNo & "']"
'
'    Set xmlEmploymentNode = vxmlRequestNode.selectSingleNode(strSearch)
'
'    If xmlEmploymentNode Is Nothing Then
'        Exit Sub
'    End If
'
'    Set xmlEmployerRefNode = xmlEmploymentNode.selectSingleNode("PREVIOUSEMPLOYERREF")
'
'    If xmlEmployerRefNode Is Nothing Then
'        Exit Sub
'    End If
'
'    intToleranceParam = _
'        xmlGetNodeAsInteger(vxmlRequestNode, "GLOBALPARAMETER[@NAME='TimeInEmploymentTolerance']/@AMOUNT")
'
'    If xmlAttributeExists(xmlEmployerRefNode, "DATESTARTED") Then
'
'        dtDateStarted = xmlGetAttributeAsDate(xmlEmploymentNode, "DATESTARTEDORESTABLISHED")
'        dtRefDateStarted = xmlGetAttributeAsDate(xmlEmployerRefNode, "DATESTARTED")
'
'        If Abs(DateDiff("d", dtDateStarted, dtRefDateStarted)) > intToleranceParam Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Declared time in Employment falls outside of accepted tolerance levels", _
'                strEmpSeqNo
'        End If
'
'    End If
'
'    If xmlAttributeExists(xmlEmployerRefNode, "DATELEFT") Then
'
'        dtDateStarted = xmlGetAttributeAsDate(xmlEmploymentNode, "DATELEFTORCEASEDTRADING")
'        dtRefDateStarted = xmlGetAttributeAsDate(xmlEmployerRefNode, "DATELEFT")
'
'        If Abs(DateDiff("d", dtDateStarted, dtRefDateStarted)) > intToleranceParam Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Declared time in Previous Employment falls outside of accepted tolerance levels", _
'                strEmpSeqNo
'        End If
'
'    End If
'MAR1316 End

End Sub

Private Sub CurrLendersRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "CurrLendersRefRules"
'    Const cstrRuleResponseTag As String = "CURRENTLENDERREFRULE"
'
'    Dim strSearch1 As String, _
'        strSearch2 As String, _
'        strAccountGuid As String
'
'    Dim lngRefOsBalance As Long, _
'        lngCustOsBalance As Long
'
'    Dim xmlNode As IXMLDOMNode
'    Dim xmlNodeList As IXMLDOMNodeList
'
'    ' JR - SYS2256, check for mandatory AccountGuid
'    If xmlAttributeExists(vxmlRequestNode, "ACCOUNTGUID") Then
'        strAccountGuid = xmlGetAttributeText(vxmlRequestNode, "ACCOUNTGUID")
'    Else
'        Err.Raise _
'            glngErrBase + 5, _
'            cstrFunctionName, _
'            "REQUEST missing ACCOUNTGUID attribute"
'    End If
'
'    strSearch1 = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch1 = strSearch1 & "/CUSTOMER/MORTGAGEACCOUNT"
'    strSearch1 = strSearch1 & "[@ACCOUNTGUID='" & strAccountGuid & "']"
'
'    strSearch2 = strSearch1 & "/LENDERSREFERENCE[@GOODACCOUNTCONDUCT]"
'
'    Set xmlNode = vxmlRequestNode.selectSingleNode(strSearch2)
'
'    If xmlNode Is Nothing Then Exit Sub     'SA SYS2256 10/5/01
'
'    ' SA SYS2256 Add in check to see if attribute exists
'    ' PSC 04/12/01 SYS3314
'    If xmlAttributeExists(xmlNode, "GOODACCOUNTCONDUCT") Then
'        If xmlGetAttributeAsBoolean(xmlNode, "GOODACCOUNTCONDUCT") = False Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Account Conduct", _
'                strAccountGuid
'        End If
'    'SA SYS2256 If they haven't answered a crucial question like this we need to highlight the fact! {
'    Else
'        AddResponse _
'            vxmlResponseNode, _
'            cstrRuleResponseTag, _
'            "Unanswered Question - Good Account Conduct ", _
'            strAccountGuid
'    'SA SYS2256 10/5/01 }
'    End If
'
'    strSearch2 = strSearch1 & "/CURRENTLENDERSREF/@BALANCEOUTSTANDING"
'
'    lngRefOsBalance = xmlGetNodeAsLong(vxmlRequestNode, strSearch2)
'
'    If lngRefOsBalance <> 0 Then
'
'        strSearch2 = strSearch1 & "/MORTGAGELOAN"
'        Set xmlNodeList = vxmlRequestNode.selectNodes(strSearch2)
'
'        For Each xmlNode In xmlNodeList
'
'            lngCustOsBalance = lngCustOsBalance + _
'                xmlGetAttributeAsLong(xmlNode, "OUTSTANDINGBALANCE")
'
'        Next
'
'        If lngRefOsBalance > lngCustOsBalance Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Outstanding Balance is greater than that declared", _
'                strAccountGuid
'        End If
'
'    End If
'MAR1316 End

End Sub

Private Sub PrevLendersRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "PrevLendersRefRules"
'    Const cstrRuleResponseTag As String = "PREVIOUSLENDERREFRULE"
'
'    Dim strSearch As String, _
'        strAccountGuid As String
'
'    Dim xmlNode As IXMLDOMNode
'
'    ' JR - SYS2256, check for mandatory AccountGuid
'    If xmlAttributeExists(vxmlRequestNode, "ACCOUNTGUID") Then
'        strAccountGuid = xmlGetAttributeText(vxmlRequestNode, "ACCOUNTGUID")
'    Else
'        Err.Raise _
'            glngErrBase + 5, _
'            cstrFunctionName, _
'            "REQUEST missing ACCOUNTGUID attribute"
'    End If
'
'    strSearch = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch = strSearch & "/CUSTOMER/MORTGAGEACCOUNT"
'    strSearch = strSearch & "[@ACCOUNTGUID='" & strAccountGuid & "']"
'    strSearch = strSearch & "/LENDERSREFERENCE[@GOODACCOUNTCONDUCT]"
'
'    Set xmlNode = vxmlRequestNode.selectSingleNode(strSearch)
'
'    If xmlNode Is Nothing Then Exit Sub     'SA SYS2256 10/5/01
'
'    'SA SYS2256 Also need to check they have answered the question
'    ' PSC 04/12/01 SYS3314
'    If xmlAttributeExists(xmlNode, "GOODACCOUNTCONDUCT") Then
'        If xmlGetAttributeAsBoolean(xmlNode, "GOODACCOUNTCONDUCT") = False Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Account Conduct", _
'                strAccountGuid
'        End If
'    'SA SYS2256 If they haven't answered a crucial question like this we need to highlight the fact! {
'    Else
'        AddResponse _
'            vxmlResponseNode, _
'            cstrRuleResponseTag, _
'            "Unanswered Question - Account Conduct ", _
'            strAccountGuid
'    'SA SYS2256 10/5/01 }
'    End If
'MAR1316 End

End Sub

Private Sub CurrLandlordRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "CurrLandlordRefRules"
'    Const cstrRuleResponseTag As String = "CURRENTLANDLORDEFRULE"
'
'    Dim strSearch As String, _
'        strCustNo As String, _
'        strCustVersNo As String, _
'        strAddressSeqNo As String
'
'    Dim xmlNode As IXMLDOMNode
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERNUMBER") Then
'        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERNUMBER attribute"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERVERSIONNUMBER") Then
'        strCustVersNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
'    Else
'        strCustVersNo = "1"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERADDRESSSEQUENCENUMBER") Then
'        strAddressSeqNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERADDRESSSEQUENCENUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 6, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERADDRESSSEQUENCENUMBER attribute"
'    End If
'
'    strSearch = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch = strSearch & "[@CUSTOMERNUMBER='" & strCustNo & "' and @CUSTOMERVERSIONNUMBER='" & strCustVersNo & "']"
'    strSearch = strSearch & "/CUSTOMER/CUSTOMERADDRESS"
'
'    strSearch = strSearch & "[@CUSTOMERADDRESSSEQUENCENUMBER='" & strAddressSeqNo & "']"
'    strSearch = strSearch & "/LANDLORDSREFERENCE[@SATISFACTORYCONDUCT]"
'
'    Set xmlNode = vxmlRequestNode.selectSingleNode(strSearch)
'
'    If xmlNode Is Nothing Then Exit Sub     'SA SYS2256 10/5/01
'
'    ' SA SYS2256 Check if they've answered the question!
'    ' PSC 04/12/01 SYS3314
'    If xmlAttributeExists(xmlNode, "SATISFACTORYCONDUCT") Then
'       If xmlGetAttributeAsBoolean(xmlNode, "SATISFACTORYCONDUCT") = False Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Unsatisfactory Conduct", _
'                strAddressSeqNo
'        End If
'    'SA SYS2256 10/5/01 Raise response if they haven't answered question{
'    Else
'        AddResponse _
'            vxmlResponseNode, _
'            cstrRuleResponseTag, _
'            "Unanswered Question: Satisfactory Conduct", _
'            strAddressSeqNo
'    'SA SYS2256 10/5/01 }
'    End If
'MAR1316 End

End Sub

Private Sub PrevLandlordRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "PrevLandlordRefRules"
'    Const cstrRuleResponseTag As String = "PREVIOUSLANDLORDEFRULE"
'
'    Dim strSearch As String, _
'        strCustNo As String, _
'        strCustVersNo As String, _
'        strAddressSeqNo As String
'
'    Dim xmlNode As IXMLDOMNode
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERNUMBER") Then
'        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 4, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERNUMBER attribute"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERVERSIONNUMBER") Then
'        strCustVersNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
'    Else
'        strCustVersNo = "1"
'    End If
'
'    If xmlAttributeExists(vxmlRequestNode, "CUSTOMERADDRESSSEQUENCENUMBER") Then
'        strAddressSeqNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERADDRESSSEQUENCENUMBER")
'    Else
'        Err.Raise _
'            glngErrBase + 6, _
'            cstrFunctionName, _
'            "REQUEST missing CUSTOMERADDRESSSEQUENCENUMBER attribute"
'    End If
'
'    strSearch = "APPLICATION/APPLICATIONCUSTOMERROLE"
'    strSearch = strSearch & "[@CUSTOMERNUMBER='" & strCustNo & "' and @CUSTOMERVERSIONNUMBER='" & strCustVersNo & "']"
'    strSearch = strSearch & "/CUSTOMER/CUSTOMERADDRESS"
'    strSearch = strSearch & "[@CUSTOMERADDRESSSEQUENCENUMBER='" & strAddressSeqNo & "']"
'    strSearch = strSearch & "/LANDLORDSREFERENCE[@SATISFACTORYCONDUCT]"
'
'    Set xmlNode = vxmlRequestNode.selectSingleNode(strSearch)
'
'    If xmlNode Is Nothing Then Exit Sub     'SA SYS2256 10/5/01
'
'    ' PSC 04/12/01 SYS3314
'    If xmlAttributeExists(xmlNode, "SATISFACTORYCONDUCT") Then
'       If xmlGetAttributeAsBoolean(xmlNode, "SATISFACTORYCONDUCT") = False Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Unsatisfactory Conduct", _
'                strAddressSeqNo
'        End If
'    'SA SYS2256 10/5/01 Raise response if they haven't answered question{
'    Else
'        AddResponse _
'            vxmlResponseNode, _
'            cstrRuleResponseTag, _
'            "Unanswered Question: Satisfactory Conduct", _
'            strAddressSeqNo
'    'SA SYS2256 10/5/01 }
'    End If
'MAR1316 End

End Sub

Private Sub ROTRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
'MAR1316 GHun
'    Const cstrFunctionName As String = "ROTRules"
'    Const cstrRuleResponseTag As String = "ROTRULE"
'
'    Dim xmlROTNode As IXMLDOMNode, _
'        xmlNewPropertyNode As IXMLDOMNode, _
'        xmlAdditionalQsList As IXMLDOMNodeList, _
'        xmlAdditionalQNode As IXMLDOMNode
'
'    Dim strSearch As String
'
'    Set xmlROTNode = vxmlRequestNode.selectSingleNode("APPLICATION/REPORTONTITLE")
'
'    If Not xmlROTNode Is Nothing Then
'
'        If xmlGetAttributeAsBoolean(xmlROTNode, "HMLREGISTERED") = True And _
'            xmlAttributeExists(xmlROTNode, "TITLENUMBER") = False _
'        Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "HMLR Registered but no Title Number"
'        End If
'
'        Set xmlNewPropertyNode = vxmlRequestNode.selectSingleNode("APPLICATION/NEWPROPERTY")
'
'        If xmlGetAttributeAsInteger(xmlNewPropertyNode, "TENURETYPE") <> _
'            xmlGetAttributeAsInteger(xmlROTNode, "PROPERTYTENURE") _
'        Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                "Property Tenure inconsistent"
'        End If
'
'    End If
'
'    Set xmlAdditionalQsList = _
'        vxmlRequestNode.selectNodes("ADDITIONALQUESTIONS[@TYPE='20' and @UNDERREVIEWIND='1']")
'
'    For Each xmlAdditionalQNode In xmlAdditionalQsList
'
'        strSearch = "APPLICATION/APPLNADDITIONALQUESTIONS"
'        strSearch = _
'            strSearch & "[@QUESTIONREFERENCE='" & _
'            xmlGetAttributeText(xmlAdditionalQNode, "QUESTIONREFERENCE") & "'"
'        strSearch = strSearch & "and @RESPONSE='1']"
'
'        If Not vxmlRequestNode.selectSingleNode(strSearch) Is Nothing Then
'            AddResponse _
'                vxmlResponseNode, _
'                cstrRuleResponseTag, _
'                xmlGetAttributeText(xmlAdditionalQNode, "QUESTIONSHORTTEXT") & _
'                    " answered True"
'        End If
'
'    Next
'MAR1316 End

End Sub

Private Sub ConditionsRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)


    Const cstrRuleResponseTag As String = "CONDITIONSRULES"
    
    Dim xmlConditionsNode As IXMLDOMNode
    
    Dim lngMiningReportsReq As Long, _
        lngTimberDampReportReq As Long, _
        strTypeOfApp As String
        
    Dim strSearch As String
    
    
    'JLD SYS3051 ensure there is always 1 condition returned
    Set xmlConditionsNode = _
        vxmlRequestNode.selectSingleNode("APPLICATION")
    'START These are Default Conditions - Always added
    ' Peter Edney - EP1099 - 29/08/06
    ' AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "701"
    ' AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "702"
    'END
    
    ' Peter Edney - EP1099 - 29/08/06
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "3"
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "4"
    
    ' Peter Edney - EP1144 - 15/09/06
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "10"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "22"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "46"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "47"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "57"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "62"
    'AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "72"
    
    'START These are Product Specific Conditions -
    'These conditions only get added if the Products attached to this Application have been linked (Thru' MORTGAGEPRODUCTCONDITION table) thru Supervisor
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "510"
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "511"
    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "512"
    'END

    If Not xmlConditionsNode Is Nothing Then
        
        strTypeOfApp = xmlGetAttributeText(xmlConditionsNode, "TYPEOFAPPLICATION")
        If Not strTypeOfApp = "" Then
            AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "503"
        End If
    End If
        
    Set xmlConditionsNode = _
        vxmlRequestNode.selectSingleNode("APPLICATION/VALUERINSTRUCTION/VALUATIONREPORT")
    
    If Not xmlConditionsNode Is Nothing Then
        
        lngMiningReportsReq = xmlGetAttributeAsLong(xmlConditionsNode, "MININGREPORTSREQ")
        If lngMiningReportsReq = 1 Then
            AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "303"
        End If
        
        lngTimberDampReportReq = xmlGetAttributeAsLong(xmlConditionsNode, "TIMBERDAMPREPORTREQ")
        If lngTimberDampReportReq = 1 Then
            AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "302"
        End If
        
    End If
    'GD BMIDS00650 Added BTL Condition (703) which contains <<BTLGT65>>  - START
    Dim blnBTL As Boolean, blnNotSpecialistScheme As Boolean, blnWhenTermReachedApplicantOver65 As Boolean
    Dim xmlSchemaData As IXMLDOMElement, xmlBTLElem As IXMLDOMElement, xmlNoneElem As IXMLDOMElement, xmlTempElem As IXMLDOMElement, xmlLoanComponent As IXMLDOMElement
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim strBTLValueid As String, strNoneValueID As String, strSpecialScheme As String, strDOB As String, strAcceptedQuoteNumber As String
    Dim iMaxterm As Integer, iCount As Integer, iIndex As Integer, iTotalMonths As Integer
    Dim dteWhenApplicant65 As Date, dteWhenTermReached As Date
    
        '-------------------- Is it a BTL application --------------------
        ' ------------------- Is it a Non SpecialistScheme Application --------------------
        blnBTL = False
        blnNotSpecialistScheme = False
        Set xmlSchemaData = vxmlRequestNode.selectSingleNode(".//APPLICATION")
        Set xmlBTLElem = xmlSchemaData.selectSingleNode(".//COMBO/COMBOVALIDATION[@GROUPNAME = 'SpecialSchemes' and @VALIDATIONTYPE = 'BTL']")
        strBTLValueid = xmlBTLElem.Attributes.getNamedItem("VALUEID").Text
        Set xmlNoneElem = xmlSchemaData.selectSingleNode(".//COMBO/COMBOVALIDATION[@GROUPNAME = 'SpecialSchemes' and @VALIDATIONTYPE = 'N']")
        strNoneValueID = xmlNoneElem.Attributes.getNamedItem("VALUEID").Text
        'GD BMIDS00759 START
'''        strSpecialScheme = xmlSchemaData.getAttribute("SPECIALSCHEME")
'''        If strSpecialScheme = Null Or strSpecialScheme = strNoneValueID Then
'''            blnNotSpecialistScheme = True
'''        End If
        Dim xmlTempAttrib As IXMLDOMAttribute
        Set xmlTempAttrib = xmlSchemaData.Attributes.getNamedItem("SPECIALSCHEME")
        If Not (xmlTempAttrib Is Nothing) Then
            strSpecialScheme = xmlTempAttrib.Text
        Else
            strSpecialScheme = ""
        End If
        
        If strSpecialScheme = "" Or strSpecialScheme = strNoneValueID Then
            blnNotSpecialistScheme = True
        End If

        'GD BMIDS00759 END
        If strSpecialScheme = strBTLValueid Then
            blnBTL = True
        End If
        

        '-------------------- Get Applicant's Date Of Birth --------------------
        
        Set xmlTempElem = xmlSchemaData.selectSingleNode(".//APPLICATIONCUSTOMERROLE/CUSTOMER")

        If Not xmlTempElem Is Nothing Then
            strDOB = xmlTempElem.Attributes.getNamedItem("DATEOFBIRTH").Text
        End If
        'Debug.Print "Applicants DOB is " & strDOB
                
        '-------------------- Get the Accepted Quote Number or the Active Quote Number--------------------
        Dim xmlTempAttribute As IXMLDOMAttribute
        Set xmlTempAttribute = xmlSchemaData.Attributes.getNamedItem("ACCEPTEDQUOTENUMBER")
        If xmlTempAttribute Is Nothing Then
            Set xmlTempAttribute = xmlSchemaData.Attributes.getNamedItem("ACTIVEQUOTENUMBER")
        End If
        
        strAcceptedQuoteNumber = xmlTempAttribute.Text
        'Debug.Print "Accepted quote number is " & strAcceptedQuoteNumber
        
        '-------------------- Get the Loan Component With The Biggest Term --------------------
        Set xmlLoanComponentList = xmlSchemaData.selectNodes(".//QUOTATION[@QUOTATIONNUMBER = '" & strAcceptedQuoteNumber & "']/MORTGAGESUBQUOTE/LOANCOMPONENT")
        
        'For Each LOANCOMPONENT in xmlLoanComponentList find the one with the biggest term (in months and years) = intBiggestTerm in months
        iMaxterm = -1
        iCount = xmlLoanComponentList.length
        For iIndex = 0 To (iCount - 1)
            Set xmlLoanComponent = xmlLoanComponentList.Item(iIndex)
            'Get the total term in months
            iTotalMonths = (CInt(xmlLoanComponent.Attributes.getNamedItem("TERMINYEARS").Text) * 12) + CInt(xmlLoanComponent.Attributes.getNamedItem("TERMINMONTHS").Text)
            If iTotalMonths > iMaxterm Then
                iMaxterm = iTotalMonths
            End If
        Next
       
        '-------------------- Calculate Date when applicant is 65 years old --------------------
        dteWhenApplicant65 = DateAdd("m", CInt(xmlSchemaData.selectSingleNode(".//GLOBALPARAMETER[@NAME = 'MaximumAge']").Attributes.getNamedItem("AMOUNT").Text) * 12, CDate(strDOB))
        
        '-------------------- Calculate Date when term would be reached as Now() + iMaxTerm --------------------
         dteWhenTermReached = DateAdd("m", iMaxterm, Now())
         
        '-------------------- Will Applicant be over 65 when (max) term reached? --------------------
        blnWhenTermReachedApplicantOver65 = (dteWhenApplicant65 < dteWhenTermReached)
        
        '-------------------- Decide if Condition is to be added --------------------
        
        If blnWhenTermReachedApplicantOver65 Then
            If blnBTL Then
        'ADD CONDITION
                AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "703"
            Else
                If blnNotSpecialistScheme Then
                    'ADD CONDITION
                    AddResponse vxmlResponseNode, cstrRuleResponseTag, "", "703"
                End If
            End If
        End If
        
    'GD BMIDS00650 Added BTL Condition (703) which contains <<BTLGT65>> END
        
End Sub

Private Sub AddResponse( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vstrRuleResponseTag As String, _
    ByVal vstrRuleMessage As String, _
    Optional ByVal vstrContext As String = "")
    
    Dim thisNode As IXMLDOMElement
    
    Set thisNode = vxmlResponseNode.ownerDocument.createElement(vstrRuleResponseTag)
    thisNode.setAttribute "RESULT", "0"
    thisNode.setAttribute "MESSAGE", vstrRuleMessage
    
    If Len(vstrContext) > 0 Then
        thisNode.setAttribute "CONTEXT", vstrContext
    End If
    
    vxmlResponseNode.appendChild thisNode
    
    Set thisNode = Nothing

End Sub

Private Function xmlParseErr(xmlDoc)
    
    Dim sError As String
    sError = "xml Parser error" & vbCrLf
    sError = sError & "reason : " & xmlDoc.parseError.reason & vbCrLf
    sError = sError & "line No. : " & xmlDoc.parseError.Line & vbCrLf
    sError = sError & "character : " & xmlDoc.parseError.linepos & vbCrLf
    sError = sError & "source text : " & xmlDoc.parseError.srcText & vbCrLf
    xmlParseErr = sError

End Function

Private Function xmlChildNodeExists( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Boolean
    
    If Not vxmlParentNode Is Nothing Then
        If Not vxmlParentNode.selectSingleNode(vstrSearchPattern) Is Nothing Then
            xmlChildNodeExists = True
        End If
    End If

End Function

Private Function xmlGetNodeText( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As String
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.selectSingleNode(vstrSearchPattern) Is Nothing Then
            xmlGetNodeText = vxmlNode.selectSingleNode(vstrSearchPattern).Text
        End If
    End If

End Function

Private Function xmlGetNodeAsBoolean( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Boolean
    
    Dim strValue As String
    
    strValue = xmlGetNodeText(vxmlNode, vstrSearchPattern)
    
    strValue = UCase(xmlGetAttributeText(vxmlNode, vstrSearchPattern))
    
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetNodeAsBoolean = True
    End If

End Function

Private Function xmlGetNodeAsDate( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Date
    
    Dim strValue As String
    
    strValue = xmlGetNodeText(vxmlNode, vstrSearchPattern)
    
    If IsDate(strValue) Then
        xmlGetNodeAsDate = CDate(strValue)
    End If

End Function

Private Function xmlGetNodeAsDouble( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Double
    
    Dim strValue As String
    
    strValue = xmlGetNodeText(vxmlNode, vstrSearchPattern)
    
    If IsNumeric(strValue) Then
        xmlGetNodeAsDouble = CDbl(strValue)
    End If

End Function

Private Function xmlGetNodeAsInteger( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Integer
    
    Dim strValue As String
    
    strValue = xmlGetNodeText(vxmlNode, vstrSearchPattern)
    
    If IsNumeric(strValue) Then
        xmlGetNodeAsInteger = CInt(strValue)
    End If

End Function

Private Function xmlGetNodeAsLong( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrSearchPattern As String) _
    As Long
    
    Dim strValue As String
    
    strValue = xmlGetNodeText(vxmlNode, vstrSearchPattern)
    
    If IsNumeric(strValue) Then
        xmlGetNodeAsLong = CLng(strValue)
    End If

End Function

Private Function xmlAttributeExists( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            If Len(vxmlNode.Attributes.getNamedItem(vstrAttribName).Text) > 0 Then
                xmlAttributeExists = True
            End If
        End If
    End If

End Function

Private Function xmlGetAttributeText( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As String
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            xmlGetAttributeText = vxmlNode.Attributes.getNamedItem(vstrAttribName).Text
        End If
    End If

End Function

Private Function xmlGetAttributeAsBoolean( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    Dim strValue As String
    
    strValue = UCase(xmlGetAttributeText(vxmlNode, vstrAttribName))
    
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetAttributeAsBoolean = True
    End If
    
End Function

Private Function xmlGetAttributeAsDate( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Date
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsDate(strValue) Then
        xmlGetAttributeAsDate = CDate(strValue)
    End If

End Function

Private Function xmlGetAttributeAsDouble( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Double
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsDouble = CDbl(strValue)
    End If

End Function

Private Function xmlGetAttributeAsInteger( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Integer
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsInteger = CInt(strValue)
    End If

End Function

Private Function xmlGetAttributeAsLong( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Long
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsLong = CLng(strValue)
    End If

End Function

Private Sub ResolveDataRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    'On Error GoTo ResolveDataRulesError
    
    Const cstrRuleResponseTag As String = "ResolveDataRules"
    Dim xmlConditionsList As IXMLDOMNodeList
    Dim xmlConditionsElem As IXMLDOMElement
    
    Dim xmlResolvedConditionDescElem As IXMLDOMElement
    Dim xmlResolvedConditionDescList As IXMLDOMElement
    Dim xmlResolvedConditionDescNodeList As IXMLDOMNodeList
    Dim xmlSchemaData As IXMLDOMElement
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim iNumConds As Integer
    Dim iIndex As Integer
    Dim strDescription As String
    Dim strErrorMsg As String
    Dim strFileErrorMessage As String
    Dim blnOneOrMoreFailed As Boolean
    Dim colTotalRules As Collection
    Dim colRules As Collection
    Dim strConditionReference As String
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strDateTime As String
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strFilename As String
    
    strFilename = xmlGetAttributeText(vxmlRequestNode.selectSingleNode(".//CONDITIONRULESLOGFILE"), "FILENAME")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set colTotalRules = New Collection
    Set xmlResolvedConditionDescList = xmlDoc.createElement("RESOLVEDCONDITIONDESCLIST")
    
    Set xmlConditionsList = vxmlRequestNode.selectNodes(".//CONDITIONS")
    Set xmlSchemaData = vxmlRequestNode.selectSingleNode(".//SCHEMADATA")
    'Get AppNum, AppFFNum, DateTime to use in error log (if necessary)
    strApplicationNumber = xmlGetAttributeText(xmlSchemaData, "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(xmlSchemaData, "APPLICATIONFACTFINDNUMBER")
    strDateTime = Now()
    iNumConds = xmlConditionsList.length
    blnOneOrMoreFailed = False
    strFileErrorMessage = ""
    strFileErrorMessage = strFileErrorMessage & "Application Number :" & strApplicationNumber
    strFileErrorMessage = strFileErrorMessage & vbCrLf & "Application Fact Find Number :" & strApplicationFFNumber
    strFileErrorMessage = strFileErrorMessage & vbCrLf & "Date Time :" & strDateTime
    strFileErrorMessage = strFileErrorMessage & vbCrLf & "----------------------------------" & vbCrLf
    For iIndex = 0 To (iNumConds - 1)
        
        Set xmlConditionsElem = xmlConditionsList.Item(iIndex)
        'Get Description
        Set xmlResolvedConditionDescElem = xmlDoc.createElement("RESOLVEDCONDITION")

        strDescription = xmlGetAttributeText(xmlConditionsElem, "CONDITIONDESCRIPTION")
        'If there is nothing in the CONDITIONDESCRIPTION, no need to resolve anything
        If Len(strDescription) <> 0 Then
            Set colRules = New Collection
            strErrorMsg = ""
            'Check valid <<..>>. Report error in strErrorMsg (if any). colRules contains all RULES (if any)...
            ' ...found in strDescription (if no error)

            If Not (ValidateAndResolveConditionText(xmlSchemaData, strDescription, strErrorMsg, colRules)) Then
                'Leave strDescription as is.
                blnOneOrMoreFailed = True
                
                'Add to strFileErrorMessage
                'APPLICATIONNUMBER
                'APPLICATIONFACTFINFNUMBER
                'DATE/TIME
                'CONDITIONREFERENCE
                '
                strConditionReference = xmlGetAttributeText(xmlConditionsElem, "CONDITIONREFERENCE")
                strFileErrorMessage = strFileErrorMessage & "Condition Reference :" & strConditionReference & vbCrLf & "------------------------" & vbCrLf
                strFileErrorMessage = strFileErrorMessage & strErrorMsg & vbCrLf
                
            End If
            
        End If
        
        'Add all the conditions back to the response, whether resolved or not..

        
        xmlResolvedConditionDescElem.setAttribute "TEXT", strDescription

        xmlResolvedConditionDescList.appendChild xmlResolvedConditionDescElem
    Next iIndex
    If blnOneOrMoreFailed Then
        'Write Error File
        WriteConditionsErrorLog strFilename, strFileErrorMessage, False
        Err.Raise _
            vbObjectError + 512 + 7, _
            cstrRuleResponseTag, _
            "Resolution of conditions failed. One or more of the embedded rules defined in your conditions text does not have a corresponding rule definition. Call Helpdesk for details of which rules failed."
        
        
        
    Else
        'Update CONDITIONSLIST.CONDITIONSDESCRIPTION with corresponding value from xmlResolvedConditionDescList
        Set xmlResolvedConditionDescNodeList = xmlResolvedConditionDescList.selectNodes("RESOLVEDCONDITION")
        For iIndex = 0 To (iNumConds - 1)
            Set xmlConditionsElem = xmlConditionsList.Item(iIndex)
            Set xmlResolvedConditionDescElem = xmlResolvedConditionDescNodeList.Item(iIndex)
            xmlSetAttributeValue xmlConditionsElem, "CONDITIONDESCRIPTION", xmlGetAttributeText(xmlResolvedConditionDescElem, "TEXT")
        Next iIndex
    
    End If
    
    vxmlResponseNode.appendChild vxmlRequestNode.selectSingleNode(".//CONDITIONSLIST")
    
'ResolveDataRulesExit:
    
    Set xmlConditionsList = Nothing
    Set xmlConditionsElem = Nothing
    Set xmlResolvedConditionDescElem = Nothing
    Set xmlResolvedConditionDescList = Nothing
    Set xmlResolvedConditionDescNodeList = Nothing
    Set xmlSchemaData = Nothing
    Set xmlDoc = Nothing
    Set colTotalRules = Nothing
    Set colRules = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing

End Sub

Private Sub xmlSetAttributeValue(ByVal vxmlDestNode As IXMLDOMNode, _
                                ByVal vstrAttribName As String, _
                                ByVal vstrAttribValue As String)

' header ----------------------------------------------------------------------------------
' description:
'   Creates a xml attribute on vxmlDestNode,
'   will overwrite existing attribute value
' pass:
'   vxmlDestNode - Destination Node for new attribute to be placed on
'   vstrAttribName - Attrbiute name of the new attribute
'   vstrAttribValue - Attribute value for new attribute on vxmlDestNode
'------------------------------------------------------------------------------------------
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName)
    xmlAttrib.Value = vstrAttribValue
        
    vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    
    Set xmlAttrib = Nothing

End Sub

Private Function ValidateAndResolveConditionText(ByVal xmlSchemaData As IXMLDOMElement, ByRef strParseString As String, ByRef strErrorResult As String, ByRef colRules As Collection) As Boolean

On Error GoTo ValidateAndResolveConditionTextError

'xmlSchemaData
'               xml element containing all the schema data


'strParseString
'               is the actual conditions text, potentially containing "sometext<<MYRULE>>somemoretext<<MYOTHERRULE>>"
'               Will be returned containing the resolved text.

'strErrorResult
'               will be built up in this method. Passed in as empty string. On completion will contain
'               a description of any error found, or nothing if no error found.

'colRules
'               will be built up in this method. Passed in as empty object. On completion will contain a collection of all
'               the rules found, eg.MYRULE, MYOTHERRULE. Will be returned empty if any errors found.



    Const strFunctionName As String = "ValidateAndResolveConditionText"
    Dim blnResult As Boolean 'result of function
    Dim intLength As Integer 'length of string in txtDescription
    Dim intIndex As Integer 'loop counter
    Dim blnOpenExpected As Boolean
    Dim strErrorMsg As String
    Dim colDBRefs As Collection
    Dim intLeftCount As Integer
    Dim intRightCount As Integer
    Dim strChar As String
    Dim intBegin As Integer
    Dim intEnd As Integer
    Const strLeftBrace = "<<"
    Const strRightBrace = ">>"

    intLeftCount = 0
    intRightCount = 0
    blnResult = True

    intLength = Len(strParseString)
    Set colDBRefs = New Collection

    intIndex = 1
    intLeftCount = 0
    intRightCount = 0
    blnOpenExpected = True
    intBegin = 1
    intEnd = 1

    strErrorMsg = ""
    While (blnResult And (intIndex <= intLength))
        strChar = Mid(strParseString, intIndex, 2)
        'Debug.Print "#" & strChar & "#"
        If blnOpenExpected Then
            If strChar = strLeftBrace Then 'found a '<<' now expecting a '>>'
                blnOpenExpected = False
                intLeftCount = intLeftCount + 1
                intBegin = intIndex

            End If

            If strChar = strRightBrace Then
                strErrorMsg = "Expecting a '" & strLeftBrace & "', but found a '" & strRightBrace & "'."
                blnResult = False
            End If

        Else 'Close Expected
                If strChar = strRightBrace Then 'found a '>>' now expecting a '<<'
                    blnOpenExpected = True
                    intRightCount = intRightCount + 1
                    intEnd = intIndex
                    'Debug.Print "BRACE CONTENT : " & Mid(strParseString, intBegin + 2, intEnd - intBegin - 2)
                    colDBRefs.Add Mid(strParseString, intBegin + 2, intEnd - intBegin - 2)
                End If

                If strChar = strLeftBrace Then
                    strErrorMsg = "Expecting a '" & strRightBrace & "', but found a '" & strLeftBrace & "'."
                    blnResult = False
                End If

            'End If
        End If

        intIndex = intIndex + 1

    Wend

    If intLeftCount <> intRightCount Then
        blnResult = False
        If Len(strErrorMsg) <> 0 Then
            strErrorMsg = "Unmatching number of '" & strLeftBrace & "' and '" & strRightBrace & "'." & strErrorMsg
        Else
            strErrorMsg = "Unmatching number of '" & strLeftBrace & "' and '" & strRightBrace & "'."
        End If
    End If
    If blnResult Then

        
        Set colRules = colDBRefs
        
    Else
        strErrorResult = strErrorMsg
        ValidateAndResolveConditionText = blnResult 'Exit and return false
    End If
    'Will only reach here if strParseString is valid
    '----------------------------------------------------------------------------------
    '------------------------------BEGIN RULE RESOLUTION--------------------------------

    Dim strRule As String
    Dim strResolvedText As String
    Dim iDescStringIndex As Integer
    Dim iRuleStringLength As Integer
    Dim iLenDescription As Integer
    Dim iLenRule As Integer
    Dim strTemp As String
    Dim strLeft As String
    Dim strRight As String
    
    Dim xmlTempElem As IXMLDOMElement
    
    intLength = colRules.Count
    For intIndex = 1 To intLength
        strRule = colRules.Item(intIndex)
        'The same rule may occur more than once in strDescription
        iRuleStringLength = Len(strRule)
        ResolveConditionRule xmlSchemaData, strRule, strResolvedText, blnResult, strErrorResult

        If blnResult Then
            'Look for strRule
            'Replace with strResolvedText
            iDescStringIndex = 1
            iLenDescription = Len(strParseString)
            strRule = "<<" & strRule & ">>"
            strRule = UCase(strRule)
            iLenRule = Len(strRule)
            
            If iLenDescription > 0 Then
                While ((iDescStringIndex < iLenDescription) And (iLenRule <= iLenDescription))
                    strTemp = UCase(Mid(strParseString, iDescStringIndex, iLenRule))
                    'Debug.Print strTemp
                    If strTemp = strRule Then ' Found a match
                        'Grab the lhs part of the Description
                        strLeft = Left(strParseString, iDescStringIndex - 1)
                        'Grab the rhs part of the Description
                        strRight = Right(strParseString, iLenDescription - (iLenRule + iDescStringIndex) + 1)
                        'Re-Build the Description with the resolved text in
                        strParseString = strLeft & strResolvedText & strRight
                        iLenDescription = Len(strParseString)
                        'Adjust the pointer posn depending on the length of the inserted rule
                        iDescStringIndex = iDescStringIndex + (Len(strResolvedText))
                    Else
                        iDescStringIndex = iDescStringIndex + 1
                    End If
                Wend
            End If
        
        End If
    Next intIndex
    ValidateAndResolveConditionText = blnResult
    
ValidateAndResolveConditionTextExit:
'Clean up

Exit Function

ValidateAndResolveConditionTextError:

Err.Raise glngErrBase + 8, "ValidateAndResolveConditionText", "An error, '" & Err.Description & "', occured trying to validate and resolve :'" & strParseString & "'"

GoTo ValidateAndResolveConditionTextExit

End Function

Private Function WriteConditionsErrorLog(sFilename As String, sErrorText As String, Optional bAppend As Boolean = True) As Boolean
'BMIDS00313
    On Error GoTo Failed
    Dim fd As Integer

Const strFunctionName As String = "WriteConditionsErrorLog"
    fd = FreeFile
    If bAppend Then
        Open sFilename For Append Access Write As fd
    Else
        Open sFilename For Output Access Write As fd
    End If

    Print #fd, sErrorText
    Close #fd
    WriteConditionsErrorLog = True
    
    Exit Function
Failed:
    WriteConditionsErrorLog = False
End Function

Private Sub ResolveConditionRule(ByVal xmlSchemaData As IXMLDOMElement, ByVal strRule As String, ByRef strResolvedText As String, ByRef blnResult, ByRef strErrorResult)
'GD BMIDS00650
    
On Error GoTo ResolveConditionRuleError
    'Generic local variables
    Dim xmlTempElem As IXMLDOMElement
    'Dim xmlTempElem1 As IXMLDOMElement
    Dim xmlTempAttrib As IXMLDOMAttribute
    
    Dim strTemp As String
    Dim strTemp1 As String
    'Dim strTemp2 As String
    Dim iIndex As Integer
    Dim iCount As Integer
    
    
    'Define local variables for use with BTLGT65 rule
    Dim blnBTL As Boolean
    Dim blnNotSpecialistScheme As Boolean
    Dim blnWhenTermReachedApplicantOver65 As Boolean
    Dim strDOB As String
    

    Dim dteWhenApplicant65 As Date
    Dim dteWhenTermReached As Date
    Dim strValueID As String
    Dim strAcceptedQuoteNumber As String
    Dim iMaxterm As Integer
    Dim iTotalMonths As Integer
    Dim strBTLValueid As String
    Dim strNoneValueID As String
    Dim strSpecialScheme As String
    Dim strCompName As String
    
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlNoneElem As IXMLDOMElement
    Dim xmlBTLElem As IXMLDOMElement
    Dim xmlLoanComponentList As IXMLDOMNodeList
    
    strRule = UCase(strRule)
    Select Case strRule

        Case "PURCHASEPRICE"
            strResolvedText = xmlGetAttributeText(xmlSchemaData, "PURCHASEPRICEORESTIMATEDVALUE")

        Case "LOANAMOUNT"
            Set xmlTempElem = xmlSchemaData.selectSingleNode(".//MORTGAGESUBQUOTE")

            strResolvedText = xmlGetAttributeText(xmlTempElem, "TOTALLOANAMOUNT")
            
        Case "BTLGT65"
            '-------------------- Is it a BTL application --------------------
            ' ------------------- Is it a Non SpecialistScheme Application --------------------
            blnBTL = False
            blnNotSpecialistScheme = False
            strCompName = "Birmingham Midshires"
            Set xmlBTLElem = xmlSchemaData.selectSingleNode(".//COMBO/COMBOVALIDATION[@GROUPNAME = 'SpecialSchemes' and @VALIDATIONTYPE = 'BTL']")
            strBTLValueid = xmlBTLElem.Attributes.getNamedItem("VALUEID").Text
            Set xmlNoneElem = xmlSchemaData.selectSingleNode(".//COMBO/COMBOVALIDATION[@GROUPNAME = 'SpecialSchemes' and @VALIDATIONTYPE = 'N']")
            strNoneValueID = xmlNoneElem.Attributes.getNamedItem("VALUEID").Text
        'GD BMIDS00759 START
'''        strSpecialScheme = xmlSchemaData.getAttribute("SPECIALSCHEME")
'''        If strSpecialScheme = Null Or strSpecialScheme = strNoneValueID Then
'''            blnNotSpecialistScheme = True
'''        End If
            'Dim xmlTempAttrib As IXMLDOMAttribute
            Set xmlTempAttrib = xmlSchemaData.Attributes.getNamedItem("SPECIALSCHEME")
            If Not (xmlTempAttrib Is Nothing) Then
                strSpecialScheme = xmlTempAttrib.Text
            Else
                strSpecialScheme = ""
            End If
            
            If strSpecialScheme = "" Or strSpecialScheme = strNoneValueID Then
                blnNotSpecialistScheme = True
            End If

        'GD BMIDS00759 END
        'GD BMIDS00913 START
        If strSpecialScheme = strBTLValueid Then
            blnBTL = True
        End If
        'GD BMIDS00913 END
            '-------------------- Get Applicant's Date Of Birth --------------------
            
            Set xmlTempElem = xmlSchemaData.selectSingleNode(".//APPLICATIONCUSTOMERROLE/CUSTOMER")

            If Not xmlTempElem Is Nothing Then
                strDOB = xmlTempElem.Attributes.getNamedItem("DATEOFBIRTH").Text
            End If
            'Debug.Print "Applicants DOB is " & strDOB
            
            '-------------------- Get the Accepted Quote Number or the Active Quote Number--------------------
            Dim xmlTempAttribute As IXMLDOMAttribute
            Set xmlTempAttribute = xmlSchemaData.Attributes.getNamedItem("ACCEPTEDQUOTENUMBER")
            If xmlTempAttribute Is Nothing Then
                Set xmlTempAttribute = xmlSchemaData.Attributes.getNamedItem("ACTIVEQUOTENUMBER")
            End If
            
            strAcceptedQuoteNumber = xmlTempAttribute.Text
            'Debug.Print "Accepted quote number is " & strAcceptedQuoteNumber
            
            '-------------------- Get the Loan Component With The Biggest Term --------------------
            Set xmlLoanComponentList = xmlSchemaData.selectNodes(".//QUOTATION[@QUOTATIONNUMBER = '" & strAcceptedQuoteNumber & "']/MORTGAGESUBQUOTE/LOANCOMPONENT")
            
            'For Each LOANCOMPONENT in xmlLoanComponentList find the one with the biggest term (in months and years) = intBiggestTerm in months
            iMaxterm = -1
            iCount = xmlLoanComponentList.length
            For iIndex = 0 To (iCount - 1)
                Set xmlLoanComponent = xmlLoanComponentList.Item(iIndex)
                'Get the total term in months
                iTotalMonths = (CInt(xmlLoanComponent.Attributes.getNamedItem("TERMINYEARS").Text) * 12) + CInt(xmlLoanComponent.Attributes.getNamedItem("TERMINMONTHS").Text)
                If iTotalMonths > iMaxterm Then
                    iMaxterm = iTotalMonths
                End If
            Next
           
            '-------------------- Calculate Date when applicant is 65 years old --------------------
            'Get GlobalParameter MaximumAge
            
            dteWhenApplicant65 = DateAdd("m", CInt(xmlSchemaData.selectSingleNode(".//GLOBALPARAMETER[@NAME = 'MaximumAge']").Attributes.getNamedItem("AMOUNT").Text) * 12, CDate(strDOB))
            
            '-------------------- Calculate Date when term would be reached as Now() + iMaxTerm --------------------
             dteWhenTermReached = DateAdd("m", iMaxterm, Now())
             
            '-------------------- Will Applicant be over 65 when (max) term reached? --------------------
            blnWhenTermReachedApplicantOver65 = (dteWhenApplicant65 < dteWhenTermReached)
            
            '-------------------- Decide the contents of the resolved text --------------------
            
            'Some default value, which should never appear
            strResolvedText = "Applicant will NOT be over 65 when the term is reached."
            If blnWhenTermReachedApplicantOver65 Then
                If blnBTL Then
                    strResolvedText = "Based on the mortgage term requested, you will have a mortgage debt outstanding " & _
                              "when you reach 65. Whilst the mortgage may be supported by rental income, you should carefully consider " & _
                                "the implications of taking your mortgage into retirement. If you have an adviser, you should consult them in this regard, " & _
                                "alternatively you should consult and independent financial adviser."
                End If
                If blnNotSpecialistScheme Then
                     strResolvedText = "Based on the mortgage term requested, you will have a mortgage debt outstanding " & _
                              "when you reach 65. You have confirmed to us the level of pension you believe you will receive in " & _
                                "retirement. By agreeing to proceed with the term requested " & strCompName & ", is not " & _
                                "in any way implying agreement that the pension income specified is accurate or achievable " & _
                                "or that you will be able to fund the mortgage into retirement. You need to carefully " & _
                                "consider the implications of taking your mortgage into retirement. If you have an " & _
                                "adviser you should consult them in this regard, alternatively you should consult " & _
                                "an independent financial advisor."
                   
                End If
                
            End If
            
        
        
        
        Case "RULE1"
            strResolvedText = "RESOLVEDRULE1"

        Case "RULE2"
            strResolvedText = "RESOLVEDRULE2"

        Case "RULE3"
            strResolvedText = "RESOLVEDRULE3"

        Case "RULE4"
            strResolvedText = "RESOLVEDRULE4"

        Case "RULE5"
            strResolvedText = "RESOLVEDRULE5"

        Case "RULE6"
            strResolvedText = "RESOLVEDRULE6"

    Case Else
        blnResult = False
        strErrorResult = strErrorResult & "Rule : '" & strRule & "' Found in text, but no rule has been defined for it yet." & vbCrLf

    End Select
    
ResolveConditionRuleExit:
'Clean up
    Set xmlLoanComponent = Nothing
    Set xmlNoneElem = Nothing
    Set xmlBTLElem = Nothing
    Set xmlLoanComponentList = Nothing

Exit Sub

ResolveConditionRuleError:

    Err.Raise glngErrBase + 7, "ResolveConditionRule", "An error :'" & Err.Description & "', occured trying to resolve rule " & strRule

    GoTo ResolveConditionRuleExit

End Sub

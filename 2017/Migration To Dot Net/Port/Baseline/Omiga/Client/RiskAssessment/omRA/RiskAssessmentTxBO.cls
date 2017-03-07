VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "RiskAssessmentTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4TxBO' with actual class name
' *****************************************************************************
'
'Workfile:      Omiga4TxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Provides update/insert for Risk Assessment Rule processing
'
'Dependencies:  RiskAssessmentDO
'
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     03/04/00    Created
'MH     31/05/00    Running a RA now updates the underwriter decision automagically
'CL     07/04/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------

Option Explicit

Implements IRiskAssessmentTxBO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrROOT_NODE_NAME = "RISKASSESSMENTRULEOVERRIDE"

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required

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

Private Function IRiskAssessmentTxBO_CreateRuleOverride(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates one or more rule overrides based on the supplied data.
'               Uses RiskAssessmentDO to persist/retrieve the necessary data.
'               Input XML will look like this:
'                    <REQUEST>
'                            <RISKASSESSMENTRULEOVERRIDE>
'                               <APPLICATIONNUMBER>00005010</APPLICATIONNUMBER>
'                               <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'                               <STAGENUMBER>1</STAGENUMBER>
'                               <RISKASSESSMENTSEQUENCENUMBER>1</RISKASSESSMENTSEQUENCENUMBER>
'                               <RARULENUMBER>7</RARULENUMBER>
'                               <USERID>Blobby</USERID>
'                               <PASSWORD>....</PASSWORD>
'                               <RAOVERRIDEDATETIME>30/03/2000 10:01:02</RAOVERRIDEDATETIME>
'                               <RAOVERRIDEREASON>He seemed like a nice chap</RAOVERRIDEREASON>
'                               <RAOVERRIDEAPPROVALREFERENCE>TTFN</RAOVERRIDEAPPROVALREFERENCE>
'                            </RISKASSESSMENTRULEOVERRIDE>
'                    </REQUEST>
'
'           Relevant rule will be retrieved, the authority checked and then an update allowed.
'
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateRuleOverrideVbErr

    Const strFunctionName As String = "IRiskAssessmentTxBO_CreateRuleOverride"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlRule As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAuthority As IXMLDOMElement
    Dim objOrg As Object
    
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    'Instantiate
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
        Set objOrg = CreateObject(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
        Set objOrg = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    End If
    
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = "RISKASSESSMENTRULEOVERRIDE" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    'Validate the userid and password
    Dim strUserId As String
    Dim strPassword As String
    Dim strRequest As String
    Dim strUnitId As String
    
    strUserId = m_objXmlAssist.GetElementText(xmlElement, "USERID")
    strUnitId = m_objXmlAssist.GetElementText(xmlElement, "UNITID")
    strPassword = m_objXmlAssist.GetElementText(xmlElement, "USERPASSWORD")
    
    strRequest = _
        "<REQUEST USERID=""" & strUserId & """ MACHINEID=""RiskAssessmentTxBO"">" & _
            "<OMIGAUSER>" & _
            "<USERID>" & strUserId & "</USERID>" & _
            "<PASSWORDVALUE>" & strPassword & "</PASSWORDVALUE>" & _
            "<UNITID>" & strUnitId & "</UNITID>" & _
            "<AUDITRECORDTYPE>1</AUDITRECORDTYPE>" & _
            "</OMIGAUSER>" & _
        "</REQUEST>"

    If m_objErrAssist.CheckXMLResponse(objOrg.ValidateUserLogon(strRequest)) <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111, "Not authorised"
    End If
    
    'Do the work
    
    '1. Get the Rule using the supplied rule override keys
    Dim xmlRuleRequest As IXMLDOMElement
    Set xmlRuleRequest = xmlElement.ownerDocument.createElement("REQUEST")
    Call m_objXmlAssist.CopyNode(xmlElement.nodeName, "RISKASSESSMENTRULE", xmlElement, xmlRuleRequest, True)
    
    Set xmlRule = objIRiskAssessmentDO.GetRiskAssessmentRule(xmlRuleRequest)
    
    '2. Get the users RA Authority
    Set xmlAuthority = objIRiskAssessmentDO.GetUserRiskAssessmentAuthority(xmlElement.getElementsByTagName("USERID").Item(0))
    
    '3. Validate that they have enough authority
    If CLng(xmlAuthority.Text) <= CLng(m_objXmlAssist.GetMandatoryElementText(xmlRule, "RARULERESULT")) Then
       m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111, "Not authorised - insufficient authority"
    End If
        
    '4. Perform the create
    objIRiskAssessmentDO.CreateRuleOverride xmlElement
    
    'Get out
    Set IRiskAssessmentTxBO_CreateRuleOverride = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateRuleOverrideExit:

    Set objIRiskAssessmentDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IRiskAssessmentTxBO_CreateRuleOverride = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateRuleOverrideExit

End Function

Private Function IRiskAssessmentTxBO_UpdateRAUnderwriterDecision(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates the underwriter decision - typical sets it.
'               Uses RiskAssessmentDO to persist/retrieve the necessary data.
'               Input XML will look like this:
'                    <REQUEST>
'                            <RISKASSESSMENT>
'                               <APPLICATIONNUMBER>00005010</APPLICATIONNUMBER>
'                               <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'                               <STAGENUMBER>1</STAGENUMBER>
'                               <RISKASSESSMENTSEQUENCENUMBER>1</RISKASSESSMENTSEQUENCENUMBER>
'                               <UNDERWRITERDECISION></UNDERWRITERDECISION>
'                    </REQUEST>
'
'           Relevant rule will be retrieved, the authority checked and then an update allowed.
'
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRAUnderwriterDecisionVbErr

    Const strFunctionName As String = "IRiskAssessmentTxBO_UpdateRAUnderwriterDecision"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlRule As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    'Instantiate
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    End If
    
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = "RISKASSESSMENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName("RISKASSESSMENT").Item(0)
    End If
        
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "RISKASSESSMENT tag not found"
    End If
    
    objIRiskAssessmentDO.UpdateRAUnderwriterDecision xmlElement
    
    'Get out
    Set IRiskAssessmentTxBO_UpdateRAUnderwriterDecision = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateRAUnderwriterDecisionExit:

    Set objIRiskAssessmentDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateRAUnderwriterDecisionVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IRiskAssessmentTxBO_UpdateRAUnderwriterDecision = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateRAUnderwriterDecisionExit

End Function

Private Sub IRiskAssessmentTxBO_StoreRiskAssResults( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByVal vobjXmlResponse As IXMLDOMNode)
    
On Error GoTo StoreRiskAssResultsVbErr
    
    Dim objRiskAssessmentDO As IRiskAssessmentDO
    Dim xmlResponse As IXMLDOMNode
    
    Set objRiskAssessmentDO = _
        m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")

    Set xmlResponse = objRiskAssessmentDO.StoreRiskAssResults( _
        vobjXmlRequest, vobjXmlResponse)
   
    Set objRiskAssessmentDO = Nothing
    
    'XML response from the DO will either be a success or an error was raised.
    'The response returned includes all the key fields that are needed to do the next bit.
    Set xmlResponse = IRiskAssessmentTxBO_UpdateRAUnderwriterDecision(xmlResponse)
    
    If Not m_objErrAssist.CheckResponse(xmlResponse, vobjXmlResponse, True) Then
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    Else
        If Not m_objContext Is Nothing Then
            m_objContext.SetAbort
        End If
    End If
    Exit Sub
    
StoreRiskAssResultsVbErr:
    
    Set objRiskAssessmentDO = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Err.Raise _
        Err.Number, _
        "[RiskAssessmentTxBO].StoreRiskAssResults." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext

End Sub



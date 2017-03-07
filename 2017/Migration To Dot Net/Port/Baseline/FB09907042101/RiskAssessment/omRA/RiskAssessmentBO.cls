VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RiskAssessmentBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      RiskAssessmentBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:
'   This module concentrates on the management of RiskAssessment requirements of the GUI
'   It will retrieve the summary of the risk assessment for a particular application
'   together with any rules which have a non-zero risk score.
'
'   It will also allow authorised users to override a bad risk and lend the money anyway.
'
'Dependencies: RiskAssessmentBO, RiskAssessmentTxBO, OrganisationDO
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     30/03/00    Created
'PSC    06/03/01    SYS1879 Add GetLatestRiskAssesmentScore
'PSC    18/10/01    SYS2821 Amend so that RunRiskAssessment expects an enclosing
'                   tag
'MEVA   25/04/02    SYS1030 Risk Assessment uses RARULESCORE rather than RARULERESULT
'CL     07/04/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------

Option Explicit

Implements IRiskAssessmentBO
Implements ObjectControl

'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrROOT_NODE_NAME = "RISKASSESSMENT"

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
'

'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    18/12/2002  BM0179 Only include non scorecard rules in RA Score & UW Decision
'SR     09/08/2004  BMIDS815 - new method CheckRiskAssessmentExistsForApplication
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'DRC    16/11/2005  MAR33 - need to pass the repsonse from the rules back to the caller to indicate what happened
'INR    29/11/2005  MAR725 Store the Underwriters decision
'IK     30/11/2005  MAR700 - error returning DECISION
'INR    01/12/2005  MAR779 Should be using Underwritersdecision for validation
'HMA    23/02/2006  MAR1315 Add GetLatestRANoStage
'------------------------------------------------------------------------------------------
'EPSOM History:
'
'Prog   Date        Description
'MAH    28/12/2006  E2_607 Resolved Type MisMatch Errors when attribures no returned to numeric variables
'IK     23/01/2007  EP2_929 fix to AppendAffordability
'LH     31/01/2007  EP2_1130 fix to AppendAffordability
'IK     19/02/2007  EP2_1429 delete any APPLICATIONUNDERWRITING record when risk ass. run
'------------------------------------------------------------------------------------------
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

Public Function GetRiskAssessment(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a Risk Assessment. Delegate to the Implementation interface version
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetRiskAssessmentVbErr

    Const strFunctionName As String = "GetRiskAssessment"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_GetRiskAssessment(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetRiskAssessment = xmlResponseElem.xml


GetRiskAssessmentExit:

    '   clean-up section, release object references

    Exit Function

GetRiskAssessmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetRiskAssessment = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetRiskAssessmentExit

End Function

Public Function GetPreviousOverrideReasons(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Returns the override reasons that have been used previously on a rule.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPreviousOverrideReasonsVbErr

    Const strFunctionName As String = "GetPreviousOverrideReasons"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If Not m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    Else
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    End If
    
    Set xmlTempResponseNode = objIRiskAssessmentDO.GetPreviousOverrideReasons(xmlIn.getElementsByTagName("REQUEST").Item(0))
    
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetPreviousOverrideReasons = xmlResponseElem.xml


GetPreviousOverrideReasonsExit:

    '   clean-up section, release object references

    Exit Function

GetPreviousOverrideReasonsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPreviousOverrideReasons = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetPreviousOverrideReasonsExit

End Function

Public Function GetUserRiskAssessmentAuthority(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  What is the user authorised for? Delegate to the Implementation interface version
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetUserRiskAssessmentAuthorityVbErr

    Const strFunctionName As String = "GetUserRiskAssessmentAuthority"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_GetUserRiskAssessmentAuthority(xmlIn.documentElement)
    xmlResponseElem.appendChild xmlTempResponseNode

    GetUserRiskAssessmentAuthority = xmlResponseElem.xml


GetUserRiskAssessmentAuthorityExit:

    '   clean-up section, release object references

    Exit Function

GetUserRiskAssessmentAuthorityVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetUserRiskAssessmentAuthority = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetUserRiskAssessmentAuthorityExit

End Function

Public Function GetLatestRiskAssessment(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the latest Risk Assessment. Delegate to the Implementation interface version
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLatestRiskAssessmentVbErr

    Const strFunctionName As String = "GetLatestRiskAssessment"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_GetLatestRiskAssessment(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetLatestRiskAssessment = xmlResponseElem.xml


GetLatestRiskAssessmentExit:

    '   clean-up section, release object references

    Exit Function

GetLatestRiskAssessmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLatestRiskAssessment = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetLatestRiskAssessmentExit

End Function

'MAR1315
Public Function GetLatestRANoStage(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the latest Risk Assessment regardless of stage.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLatestRANoStageVbErr

    Const strFunctionName As String = "GetLatestRANoStage"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_GetLatestRANoStage(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetLatestRANoStage = xmlResponseElem.xml


GetLatestRANoStageExit:

    Set xmlIn = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOut = Nothing

    Exit Function

GetLatestRANoStageVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLatestRANoStage = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetLatestRANoStageExit

End Function

Public Function GetRiskAssessmentApplicationStages(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Find all stages within an application that have data
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetRiskAssessmentApplicationStagesVbErr

    Const strFunctionName As String = "GetRiskAssessmentApplicationStages"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objDO As IRiskAssessmentDO
    
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If Not m_objContext Is Nothing Then
        Set objDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    Else
        Set objDO = New RiskAssessmentDO
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = objDO.GetRAApplicationStages(xmlIn.documentElement)
    xmlResponseElem.appendChild xmlTempResponseNode

    GetRiskAssessmentApplicationStages = xmlResponseElem.xml


GetRiskAssessmentApplicationStagesExit:

    '   clean-up section, release object references

    Exit Function

GetRiskAssessmentApplicationStagesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetRiskAssessmentApplicationStages = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetRiskAssessmentApplicationStagesExit

End Function

Public Function ValidateRuleOverride(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Validate an override. Delegate to the Implementation interface version
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateRuleOverrideVbErr

    Const strFunctionName As String = "ValidateRuleOverride"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_ValidateRuleOverride(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    ValidateRuleOverride = xmlResponseElem.xml


ValidateRuleOverrideExit:

    '   clean-up section, release object references

    Exit Function

ValidateRuleOverrideVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateRuleOverride = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume ValidateRuleOverrideExit

End Function

Public Function CreateRuleOverride(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create an override. Delegate to the Implementation interface version
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateRuleOverrideVbErr

    Const strFunctionName As String = "CreateRuleOverride"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IRiskAssessmentBO_CreateRuleOverride(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    CreateRuleOverride = xmlResponseElem.xml


CreateRuleOverrideExit:

    '   clean-up section, release object references

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

    ' go to clean-up section
    Resume CreateRuleOverrideExit

End Function

Public Function UpdateRAUnderwriterDecision(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: Allows a user to change the underwriter's decision from Pass/Fail
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
'    <REQUEST>
'            <RISKASSESSMENT>
'               <APPLICATIONNUMBER>00005010</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'               <STAGENUMBER>1</STAGENUMBER>
'               <RISKASSESSMENTSEQUENCENUMBER>1</RISKASSESSMENTSEQUENCENUMBER>
'               <UNDERWRITERDECISION></UNDERWRITERDECISION>
'            </RISKASSESSMENTRULEOVERRIDE>
'    </REQUEST>

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRAUnderwriterDecisionVbErr

    Const strFunctionName As String = "UpdateRAUnderwriterDecision"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objIRiskAssessmentTxBO As IRiskAssessmentTxBO

    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentTxBO = New RiskAssessmentTxBO
    Else
        Set objIRiskAssessmentTxBO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentTxBO")
    End If
    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = objIRiskAssessmentTxBO.UpdateRAUnderwriterDecision(xmlIn.documentElement)
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    UpdateRAUnderwriterDecision = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateRAUnderwriterDecisionExit:

    '   clean-up section, release object references

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

    UpdateRAUnderwriterDecision = m_objErrAssist.CreateErrorResponse

     If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
   ' go to clean-up section
    Resume UpdateRAUnderwriterDecisionExit

End Function

Private Function IRiskAssessmentBO_CreateRuleOverride(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo CreateRuleOverrideVbErr

    Const strFunctionName As String = "IRiskAssessmentBO_CreateRuleOverride"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIRiskAssessmentTxBO As IRiskAssessmentTxBO
    
    'Init
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentTxBO = New RiskAssessmentTxBO
    Else
        Set objIRiskAssessmentTxBO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentTxBO")
    End If
    
    'Delegate the work to something that has transaction control
    Set IRiskAssessmentBO_CreateRuleOverride = _
                                objIRiskAssessmentTxBO.CreateRuleOverride(vxmlRequest)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateRuleOverrideExit:

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

    Set IRiskAssessmentBO_CreateRuleOverride = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateRuleOverrideExit


End Function

Private Function IRiskAssessmentBO_GetLatestRiskAssessment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'  This is the same as GetRiskAssessment but returns the latest one.
'  The DO will detect the missing field and GetLatest instead.
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IRiskAssessmentBO_GetLatestRiskAssessment"

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set IRiskAssessmentBO_GetLatestRiskAssessment = IRiskAssessmentBO_GetRiskAssessment(vxmlRequest)

GetDataExit:

   Exit Function

GetDataVbErr:

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IRiskAssessmentBO_GetLatestRiskAssessment = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit


End Function

'MAR1315 Add function
Private Function IRiskAssessmentBO_GetLatestRANoStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'
' Detect the latest stage for which a Risk Assessment exists and then call GetRiskAssessment
' to return the latest set.
'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IRiskAssessmentBO_GetLatestRANoStage"

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlStageResponse As IXMLDOMNode
    Dim xmlStage As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlApplicationFactFind As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objDO As IRiskAssessmentDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    
    Set xmlApplication = m_objXmlAssist.GetMandatoryNode(vxmlRequest, "APPLICATIONNUMBER")
    Set xmlApplicationFactFind = m_objXmlAssist.GetNode(vxmlRequest, "APPLICATIONFACTFINDNUMBER")
    
    If Not m_objContext Is Nothing Then
        Set objDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    Else
        Set objDO = New RiskAssessmentDO
    End If
        
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Get all the stages for which case assessment results are found
    Set xmlStageResponse = objDO.GetRAApplicationStages(vxmlRequest)

    'Check that at least one stage has been found
    If (xmlStageResponse.selectNodes("APPLICATIONSTAGE").length > 0) Then

        'Get latest stage - result from GetRAApplicationStages is returned with latest stage first.
        Set xmlStage = xmlStageResponse.firstChild.selectSingleNode("STAGENUMBER")
    
        'Get the latest risk assessment for the latest stage
        'When no sequence number is passed into GetRiskAssessment, the latest is returned.
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlElem = xmlOut.createElement("RISKASSESSMENT")
    
        xmlElem.appendChild xmlApplication
        xmlElem.appendChild xmlApplicationFactFind
        xmlElem.appendChild xmlStage
    
        xmlNewRequest.appendChild xmlElem

        Set IRiskAssessmentBO_GetLatestRANoStage = IRiskAssessmentBO_GetRiskAssessment(xmlNewRequest)
    Else
        Set IRiskAssessmentBO_GetLatestRANoStage = xmlResponseElem
    End If

GetDataExit:

    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlStageResponse = Nothing
    Set xmlStage = Nothing
    Set xmlApplication = Nothing
    Set xmlApplicationFactFind = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElem = Nothing
    Set objDO = Nothing

   Exit Function

GetDataVbErr:

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IRiskAssessmentBO_GetLatestRANoStage = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit


End Function

Private Function IRiskAssessmentBO_GetRiskAssessment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'      business object
'      Returns the RiskAssessment summary record and all of the associated risk
'       assessment rules together with the override data.
'   Input
'       <REQUEST>
'       <RISKASSESSMENT>
'           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'           <STAGENUMBER></STAGENUMBER>
'           <RISKASSESSMENTSEQUENCENUMBER></RISKASSESSMENTSEQUENCENUMBER> - if this is missing, get latest
'       </RISKASSESSMENT>
'       </REQUEST>
'
'   Output:
'           <RISKASSESSMENT>  - The basic deatils
'           </RISKASSESSMENT>
'           <RISKASSESSMENTRULESLIST> - all of the rules
'           </RISKASSESSMENTRULESLIST>
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IRiskAssessmentBO_GetRiskAssessment"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIRiskAssessmentDO As IRiskAssessmentDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    End If

    'Get RISKASSESSMENT node
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If

    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If

    Set xmlDataNode = objIRiskAssessmentDO.GetRiskAssessment(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    Set IRiskAssessmentBO_GetRiskAssessment = xmlResponseElem

GetDataExit:

    Set xmlRequestNode = Nothing
    Set objIRiskAssessmentDO = Nothing
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

    Set IRiskAssessmentBO_GetRiskAssessment = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function IRiskAssessmentBO_ValidateRuleOverride(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Const strFunctionName As String = "IRiskAssessmentBO_ValidateRuleOverride"
    On Error GoTo GetDataVbErr
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlCompetency As IXMLDOMElement
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    'Build a default response
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Instantiate DO's
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    End If
    
    'Get the request
    If vxmlRequest.nodeName = "RISKASSESSMENTRULEOVERRIDE" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("RISKASSESSMENTRULEOVERRIDE").Item(0)
    End If
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    '---------------------
    'Process the request
    '---------------------
    'Assess how much risk the user can handle
    Set xmlCompetency = objIRiskAssessmentDO.GetUserRiskAssessmentAuthority(vxmlRequest.getElementsByTagName("USERID").Item(0))
   
    If UserHasAuthority(xmlCompetency, xmlRequestNode) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInternalError, "User not authorised"
    End If
    
    Set IRiskAssessmentBO_ValidateRuleOverride = xmlResponseElem
    
GetDataExit:


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

    Set IRiskAssessmentBO_ValidateRuleOverride = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function IRiskAssessmentBO_GetUserRiskAssessmentAuthority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Const strFunctionName As String = "IRiskAssessmentBO_GetUserRiskAssessmentAuthority"
    On Error GoTo GetDataVbErr
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
     
    'Build a default response
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Instantiate DO's
    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    End If
    
    Set IRiskAssessmentBO_GetUserRiskAssessmentAuthority = objIRiskAssessmentDO.GetUserRiskAssessmentAuthority(vxmlRequest)
    
GetDataExit:


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

    Set IRiskAssessmentBO_GetUserRiskAssessmentAuthority = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function RunRiskAssessment(ByVal strXmlIn As String) As String
On Error GoTo RunRiskAssessmentVbErr

    Const strFunctionName As String = "RunRiskAssessment"

    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    objXmlOut.async = False

    ' *************************************************************************
    ' validate request block
    ' *************************************************************************
    Set objXmlIn = m_objXmlAssist.load(strXmlIn, TypeName(Me), strFunctionName)
    Set objXmlRequestNode = m_objXmlAssist.GetMandatoryNode(objXmlIn, "REQUEST")
    
    'ik_debug
    'objXmlIn.Save "c:\omiga4Trace\omRA.RiskAssessmentBO.RunRiskAssessment.Request.xml"

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    'ik_MAR700_20051130
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
    
    IRiskAssessmentBO_RunRiskAssessment objXmlRequestNode, objXmlResponseNode
    
    'ik_MAR700_20051130
    RunRiskAssessment = objXmlOut.xml

RunRiskAssessmentExit:

    Set objXmlRequestNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlResponseNode = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing

    Exit Function

RunRiskAssessmentVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    RunRiskAssessment = m_objErrAssist.CreateErrorResponse

    '   go to clean-up section
    Resume RunRiskAssessmentExit

End Function

Private Sub IRiskAssessmentBO_RunRiskAssessment( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByRef vobjXmlResponse As IXMLDOMNode)
    
    Dim objXmlRulesResponse As FreeThreadedDOMDocument40
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRulesRequestNode As IXMLDOMNode
    Dim xmlRiskAssessmentNode As IXMLDOMNode
    
    ' late binding allows Script Component or COM Component
    Dim objRulesExec As Object
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strRequest As String, _
        strRulesResponse As String
    
    strApplicationNumber = _
        m_objXmlAssist.GetMandatoryElementText(vobjXmlRequest, ".//APPLICATIONNUMBER")
        
    strApplicationFactFindNumber = _
        m_objXmlAssist.GetMandatoryElementText(vobjXmlRequest, ".//APPLICATIONFACTFINDNUMBER")
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlRulesRequestNode = _
        xmlRequestDoc.appendChild(vobjXmlRequest.cloneNode(False))
        
    ' PSC 18/10/01 SYS2821 - Start
    Set xmlRiskAssessmentNode = m_objXmlAssist.GetMandatoryNode(vobjXmlRequest, "RISKASSESSMENT")

    'IK 19/02/2007 EP2_1429
    Call DeleteApplicationUnderwriting(xmlRiskAssessmentNode)
        
    GetApplicationData xmlRiskAssessmentNode, xmlRulesRequestNode, vobjXmlResponse
    ' PSC 18/10/01 SYS2821 - End

    'Appends Affordability data if required
    'EP2_929
    AppendAffordability xmlRequestDoc.documentElement, vobjXmlResponse 'EP2_128 MCh 29/11/2006
    
    Set objRulesExec = m_objContext.CreateInstance(gstrRISK_ASSESSMENT_RULES_COMP & ".RARulesBO")

    strRulesResponse = objRulesExec.RiskAssessmentRules(xmlRequestDoc.xml)

    Set objRulesExec = Nothing

    Set objXmlRulesResponse = _
        m_objXmlAssist.load(strRulesResponse, TypeName(Me), "RunRiskAssessment")
        
    ParseRulesResults vobjXmlResponse, objXmlRulesResponse
    'MAR33 DRC - Put decision into request
    'ik_MAR700_20051130
    If Not (vobjXmlResponse.selectSingleNode("DECISION") Is Nothing) Then
        xmlRiskAssessmentNode.appendChild vobjXmlResponse.selectSingleNode("DECISION").cloneNode(True)
    End If

    ' *************************************************************************
    ' store rule results
    ' *************************************************************************
    
    StoreResults vobjXmlRequest, objXmlRulesResponse
    'MAR725 Need to store the Underwriters Decision
    'EP2_1429
    'StoreUWResults vobjXmlRequest
    
    'ik_MAR700_20051130
    'Set vobjXmlResponse = objXmlRulesResponse
'    Set vobjXmlResponse = _
'        m_objXmlAssist.load(strRulesResponse, TypeName(Me), "RunRiskAssessment")
    
    Set xmlRulesRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    Set objXmlRulesResponse = Nothing
    Set xmlRiskAssessmentNode = Nothing

End Sub

Private Sub GetApplicationData( _
    ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
    ByVal vxmlApplicationDestNode As IXMLDOMNode, _
    ByVal vobjXmlResponse As IXMLDOMNode)
    
    Const cstrFunctionName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    
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
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Text = "RARulesTemplate"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("APPLICATIONNUMBER").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = m_objContext.CreateInstance("omRB.OmRequestDO")
    xmlResponseDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlResponseDoc.xml
    
    m_objErrAssist.CheckResponse _
        xmlResponseDoc.selectSingleNode("RESPONSE"), _
        vobjXmlResponse, _
        True

    For Each xmlNode In xmlResponseDoc.selectSingleNode("RESPONSE").childNodes
        vxmlApplicationDestNode.appendChild xmlNode.cloneNode(True)
    Next
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
End Sub

' MCh EP2_128 29/11/2006
' AppendAffordability: If AFDBAffordability = true (Changed to RAIncludeAffordability EP2_444)
'                      Calculate the Total Allowable Income for the application
'                      Calculate the ApproximateMonthlyCost for the application
' EP2_929 add response node parameter
' EP2_1313 - Exclude Guarantors if not required.

Private Sub AppendAffordability( _
    ByVal vxmlApplicationDestNode As MSXML2.IXMLDOMNode, _
    ByRef vobjXmlResponse As IXMLDOMNode)

    'Dim blnAFDBAffordability As Boolean --&& EP2_444
    Dim blnRAIncludeAffordability As Boolean '++&& EP2_444
    Dim lstCustomerVersion As IXMLDOMNodeList
    Dim ndCustomerVersion As IXMLDOMNode
    Dim ndIncomeSummary As IXMLDOMNode
    ' EP2_1313
    Dim lstCustomerRole As IXMLDOMNodeList
    Dim ndCustomerRole As IXMLDOMNode

    
    Dim dblTotalAllowableIncome As Double
    Dim dblTotalApproximateMonthlyCost As Double
    Dim attNetConfirmedAllowableIncome As IXMLDOMAttribute
    Dim dblNetConfirmedAllowableIncome As Double
    Dim attNetAllowableAnnualIncome As IXMLDOMAttribute
    Dim dblNetAllowableAnnualIncome As Double
    
    Dim objApplicationQuoteBO As Object
    Dim xmlDoc As DOMDocument
    Dim xmlRequest As IXMLDOMElement
    Dim xmlApplicationNumber As IXMLDOMElement
    Dim xmlApplicationFactFindNumber As IXMLDOMElement
    Dim xmlApplication As IXMLDOMElement
    Dim xmlSrcApplication As IXMLDOMElement
    
    Dim strAttributeValue As String ' EP_607 MAH 28/12/2006
    
    Dim bIncludeGuarantors As Boolean    ' EP2_1313 Global Param?
    Dim sCustRoleType As String          ' EP2_1313 Customer Role Type?

'temp?
Dim strString As String

    'Attribute may not exist so continue if error raised
'    blnAFDBAffordability = m_objXmlAssist.GetAttributeValue(vxmlApplicationDestNode, _
'                                                            "GLOBALPARAMETER[@NAME='AFDBAffordability']", _
'                                                            "BOOLEAN")' --&& EP2_444

    If Not vxmlApplicationDestNode.selectSingleNode("APPLICATION/GLOBALPARAMETER[@NAME='RAIncludeAffordability']") Is Nothing Then
        blnRAIncludeAffordability = _
            m_objXmlAssist.GetAttributeValue( _
            vxmlApplicationDestNode, _
            "GLOBALPARAMETER[@NAME='RAIncludeAffordability']", _
            "BOOLEAN") '++&& EP2_444
    End If
                
'    If blnAFDBAffordability Then --&& EP2_444
    If blnRAIncludeAffordability Then '++&& EP2_444
        
        ' EP2_1313 - Get Global Param
        If Not vxmlApplicationDestNode.selectSingleNode("APPLICATION/GLOBALPARAMETER[@NAME='IncludeGuarantors']") Is Nothing Then
            bIncludeGuarantors = _
                m_objXmlAssist.GetAttributeValue( _
                vxmlApplicationDestNode, _
                "GLOBALPARAMETER[@NAME='IncludeGuarantors']", _
                "BOOLEAN") ' EP2_1313
        End If
                   
        
        ' EP2_1313 - New loop to obtain Customer Role Type.
        Set lstCustomerRole = vxmlApplicationDestNode.selectNodes(".//CUSTOMERROLE")
        
        For Each ndCustomerRole In lstCustomerRole
        
            ' Get the Customer Role type for this customer.
            sCustRoleType = m_objXmlAssist.GetAttributeValue(ndCustomerRole, "CUSTOMERROLE", "CUSTOMERROLETYPE")
            
            ' Now loop round the CustomerVersions.
            Set lstCustomerVersion = ndCustomerRole.selectNodes(".//CUSTOMERVERSION")
        
            For Each ndCustomerVersion In lstCustomerVersion
            
                ' EP2_1313 Exclude guarantors if specified
                If bIncludeGuarantors = True Or sCustRoleType <> "2" Then
                    Set ndIncomeSummary = ndCustomerVersion.selectSingleNode("./INCOMESUMMARY")
                    
                    If Not (ndIncomeSummary Is Nothing) Then
                        ' EP_607 MAH 28/12/2006 Start
                        '                dblNetConfirmedAllowableIncome = m_objXmlAssist.GetAttributeValue(ndIncomeSummary, _
                        '                                                                                  "INCOMESUMMARY", _
                        '                                                                                  "NETCONFIRMEDALLOWABLEINCOME")
                        '
                        '                dblNetAllowableAnnualIncome = m_objXmlAssist.GetAttributeValue(ndIncomeSummary, _
                        '                                                                               "INCOMESUMMARY", _
                        '                                                                               "NETALLOWABLEANNUALINCOME")
                        strAttributeValue = m_objXmlAssist.GetAttributeValue(ndIncomeSummary, _
                                                "INCOMESUMMARY", _
                                                "NETCONFIRMEDALLOWABLEINCOME")
                          
                        dblNetConfirmedAllowableIncome = CDbl(IIf(strAttributeValue = vbNullString, 0, strAttributeValue))
                        strAttributeValue = m_objXmlAssist.GetAttributeValue(ndIncomeSummary, _
                                                "INCOMESUMMARY", _
                                                "NETALLOWABLEANNUALINCOME")
                        dblNetAllowableAnnualIncome = CDbl(IIf(strAttributeValue = vbNullString, 0, strAttributeValue))
                        ' EP_607 MAH 28/12/2006 End
                        dblTotalAllowableIncome = dblTotalAllowableIncome + IIf(dblNetConfirmedAllowableIncome > 0, dblNetConfirmedAllowableIncome / 12, dblNetAllowableAnnualIncome / 12)
                    End If
                
                End If ' EP2_1313
            
            Next  'End ndCustomerVersion In lstCustomerVersion
        
        Next ' End ndCustomerRole in lstCustomerRole

        '---------------------------------------------------------
        'Build XML for omAQ.ApplicationQuoteBO.CalcAllowableInc() call
        
        Set xmlDoc = New MSXML2.DOMDocument
        Set xmlRequest = xmlDoc.createElement("REQUEST")
        
        Set xmlApplication = xmlDoc.createElement("APPLICATION")
        
        Set xmlApplicationNumber = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlApplicationNumber.Text = m_objXmlAssist.GetMandatoryAttribute(m_objXmlAssist.GetMandatoryNode(vxmlApplicationDestNode, "//APPLICATION"), "APPLICATIONNUMBER")

        Set xmlApplicationFactFindNumber = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlApplicationFactFindNumber.Text = m_objXmlAssist.GetMandatoryAttribute(m_objXmlAssist.GetMandatoryNode(vxmlApplicationDestNode, "//APPLICATION"), "APPLICATIONFACTFINDNUMBER")

        xmlRequest.appendChild xmlApplication
        xmlApplication.appendChild xmlApplicationNumber
        xmlApplication.appendChild xmlApplicationFactFindNumber
        xmlDoc.appendChild xmlRequest

        '----------------------------------------------------------

        Set objApplicationQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
' EP_607 MAH 28/12/2006 Start
        'dblTotalApproximateMonthlyCost = objApplicationQuoteBO.CalculateApproximateMonthlyCost(xmlRequest.xml)
        strAttributeValue = objApplicationQuoteBO.CalculateApproximateMonthlyCost(xmlRequest.xml)
        
        'EP2_929 - test for error response
        If Left(strAttributeValue, 10) = "<RESPONSE " Then
            xmlDoc.loadXML strAttributeValue
    
            m_objErrAssist.CheckResponse _
                xmlDoc.selectSingleNode("RESPONSE"), _
                vobjXmlResponse, _
                True
                
        Else
            dblTotalApproximateMonthlyCost = CDbl(IIf(IsNumeric(strAttributeValue), strAttributeValue, 0))
        End If
        
' EP_607 MAH 28/12/2006 End
        Set xmlApplication = m_objXmlAssist.GetMandatoryNode(vxmlApplicationDestNode, "APPLICATION")
        xmlApplication.setAttribute "TOTALALLOWABLEINCOME", dblTotalAllowableIncome
        xmlApplication.setAttribute "TOTALAPPROXIMATEMONTHLYCOST", dblTotalApproximateMonthlyCost
    
    End If

End Sub

Private Sub ParseRulesResults( _
    ByVal vobjResponseNode As IXMLDOMNode, _
    ByVal vobjRulesResponse As FreeThreadedDOMDocument40)
    
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    'ik_MAR700_20051130
    Dim objXmlAttrib As IXMLDOMAttribute

    Dim objRiskAssessmentDO As IRiskAssessmentDO

    Dim strMessage As String, _
        strFailCount As String, _
        strDecision As String
    Dim intFailCount As Integer 'BM0179 MDC 18/12/2002

    If m_objXmlAssist.GetAttributeFromNode(vobjRulesResponse.firstChild, "TYPE") <> "SUCCESS" Then
        
        If Not vobjRulesResponse.selectSingleNode("RESPONSE/MESSAGE") Is Nothing Then
            strMessage = vobjRulesResponse.selectSingleNode("RESPONSE/MESSAGE").Text
        Else
            strMessage = "unspecified error"
        End If
        
        m_objErrAssist.ThrowError _
            TypeName(Me), _
            "ParseRulesResults", _
            oeInternalError, _
            strMessage
     'MARS269 DRC 31/10/2005
    Else
      
        'Copy  Rules Response header to Response
        'ik_MAR700_20051130
        For Each objXmlAttrib In vobjRulesResponse.firstChild.Attributes
            vobjResponseNode.Attributes.setNamedItem objXmlAttrib.cloneNode(True)
        Next
        For Each objXmlNode In vobjRulesResponse.firstChild.childNodes
            vobjResponseNode.appendChild objXmlNode.cloneNode(True)
        Next
    
    End If
     'MARS269 End
    
    'BM0179 MDC 18/12/2002
    'strFailCount = _
        m_objXmlAssist.GetAttributeFromNode(vobjRulesResponse.firstChild, "FAILCOUNT")
    'intFailCount = vobjRulesResponse.selectNodes("RULE[@SCORE>0 and @SCORECARDIND='0']").length
    'MARS269 DRC 31/10/2005
    intFailCount = CInt(m_objXmlAssist.GetAttributeFromNode(vobjRulesResponse.firstChild, "FAILLEVEL"))
        
'    If Not strFailCount = Empty Then
    
'        If IsNumeric(strFailCount) Then
            
            Set objRiskAssessmentDO = _
                m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
                
            'strDecision = _
                objRiskAssessmentDO.GetRiskAssessmentDecisionParameter( _
                    CInt(strFailCount))
            strDecision = objRiskAssessmentDO.GetRiskAssessmentDecisionParameter(intFailCount)

            Set objRiskAssessmentDO = Nothing
                
            Set objXmlNode = _
                vobjRulesResponse.createAttribute("DECISION")
                
            objXmlNode.Text = strDecision
                    
            vobjRulesResponse.firstChild.Attributes.setNamedItem objXmlNode
            
            Set objXmlElem = _
                vobjResponseNode.ownerDocument.createElement("DECISION")
            objXmlElem.Text = strDecision
            Set objXmlNode = vobjResponseNode.appendChild(objXmlElem)
                
            Set objXmlElem = Nothing
            Set objXmlNode = Nothing
        
'        End If
    
'    End If
    'BM0179 MDC 18/12/2002 - End

End Sub

Private Sub StoreResults( _
    ByVal vobjXmlRequestNode As IXMLDOMNode, _
    ByVal vobjRulesResponse As FreeThreadedDOMDocument40)
    
    Dim objXmlRulesResponseNode As IXMLDOMNode
    Dim objRiskAssessmentTxBO As IRiskAssessmentTxBO
    
    Set objXmlRulesResponseNode = _
        vobjRulesResponse.firstChild
        
    Set objRiskAssessmentTxBO = _
        m_objContext.CreateInstance(App.Title & ".RiskAssessmentTxBO")
        
    objRiskAssessmentTxBO.StoreRiskAssResults _
        vobjXmlRequestNode, _
        objXmlRulesResponseNode
        
    Set objRiskAssessmentTxBO = Nothing
    Set objXmlRulesResponseNode = Nothing

End Sub

'MAR725 Store the Underwriters decision
Private Sub StoreUWResults(ByVal vobjXmlRequestNode As IXMLDOMNode)

    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempParentElem As IXMLDOMNode
    Dim xmlStoreRequest As IXMLDOMElement
    Dim xmlRiskAssess As IXMLDOMNode
    Dim objApplicationBO As ApplicationBO
    Dim userID As String
    Dim decisionText As String
    Dim decisionCompare As String
    Dim objComboDO As ComboDO
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True

    Set xmlStoreRequest = vobjXmlRequestNode.cloneNode(False)
    Set xmlRiskAssess = vobjXmlRequestNode.selectSingleNode(".//RISKASSESSMENT")
    userID = xmlStoreRequest.getAttribute("USERID")
    decisionText = xmlRiskAssess.selectSingleNode(".//DECISION").Text
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    decisionCompare = objComboDO.GetFirstComboValueId("RiskAssessmentDecision", "R")
   
    If (StrComp(decisionText, decisionCompare) = 0) Then
        
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        'MAR779 Should be using Underwritersdecision for validation
        decisionText = objComboDO.GetFirstComboValueId("Underwritersdecision", "R")

        'We have a refer decision and need to save it
        Set xmlTempParentElem = xmlRequestDoc.createElement("APPLICATIONUNDERWRITING")
        
        Set xmlTempElem = xmlRequestDoc.createElement("APPLICATIONNUMBER")
        xmlTempElem.Text = xmlRiskAssess.selectSingleNode(".//APPLICATIONNUMBER").Text
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlTempElem.Text = xmlRiskAssess.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("UNDERWRITERSDECISION")
        xmlTempElem.Text = decisionText
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("UNDERWRITERDECISIONDATETIME")
        xmlTempElem.Text = CStr(Now)
        xmlTempParentElem.appendChild xmlTempElem
        
        Set xmlTempElem = xmlRequestDoc.createElement("USERID")
        xmlTempElem.Text = userID
        xmlTempParentElem.appendChild xmlTempElem
    
        xmlStoreRequest.appendChild xmlTempParentElem
        
        objApplicationBO.SaveApplicationUnderwriting xmlStoreRequest.xml
    
    End If
    
    Set xmlRequestDoc = Nothing
    Set xmlTempElem = Nothing
    Set xmlTempParentElem = Nothing
    Set xmlStoreRequest = Nothing
    Set xmlRiskAssess = Nothing
        
    Set objApplicationBO = Nothing
    Set objComboDO = Nothing
    
End Sub

Private Function UserHasAuthority(ByVal vxmlCompetency As IXMLDOMNode, ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
On Error GoTo errhandler
    UserHasAuthority = (CLng(vxmlCompetency.Text) >= CLng(m_objXmlAssist.GetMandatoryNode(vxmlRequestNode, "RARULESCORE").Text))
    Exit Function

errhandler:
    Debug.Assert False
    UserHasAuthority = False
End Function

Public Function GetLatestRiskAssessmentScore(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLatestRiskAssessmentScoreVbErr

    Const strFunctionName As String = "GetLatestRiskAssessmentScore"
        
        
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
    Set xmlTempResponseNode = IRiskAssessmentBO_GetLatestRiskAssessmentScore(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLatestRiskAssessmentScore = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetLatestRiskAssessmentScoreExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetLatestRiskAssessmentScoreVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLatestRiskAssessmentScore = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetLatestRiskAssessmentScoreExit

End Function

Private Function IRiskAssessmentBO_GetLatestRiskAssessmentScore(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLatestRiskAssessmentScoreVbErr
    
    Const strFunctionName As String = "IRiskAssessmentBO_GetLatestRiskAssessmentScore"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIRiskAssessmentDO.GetLatestRiskAssessmentScore(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IRiskAssessmentBO_GetLatestRiskAssessmentScore = xmlResponseElem
    
    m_objContext.SetComplete
    
GetLatestRiskAssessmentScoreExit:
    
    Set xmlRequestNode = Nothing
    Set objIRiskAssessmentDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetLatestRiskAssessmentScoreVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IRiskAssessmentBO_GetLatestRiskAssessmentScore = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetLatestRiskAssessmentScoreExit

End Function


Public Function CheckRiskAssessmentExistsForApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CheckRiskAssessmentExistsForApplicationVbErr

    Const strFunctionName As String = "CheckRiskAssessmentExistsForApplication"
        
        
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
    Set xmlTempResponseNode = IRiskAssessmentBO_CheckRiskAssessmentExistsForApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CheckRiskAssessmentExistsForApplication = xmlResponseElem.xml
    
    m_objContext.SetComplete

CheckRiskAssessmentExistsForApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

CheckRiskAssessmentExistsForApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CheckRiskAssessmentExistsForApplication = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CheckRiskAssessmentExistsForApplicationExit

End Function

' SR 09/08/2004 : BMIDS815 - new method
Private Function IRiskAssessmentBO_CheckRiskAssessmentExistsForApplication( _
                                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IRiskAssessmentBO_CheckRiskAssessmentExistsForApplication"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIRiskAssessmentDO As IRiskAssessmentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIRiskAssessmentDO = New RiskAssessmentDO
    Else
        Set objIRiskAssessmentDO = m_objContext.CreateInstance(App.Title & ".RiskAssessmentDO")
    End If

    'Get RISKASSESSMENT node
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If

    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objIRiskAssessmentDO.CheckRiskAssessmentExistsForApplication(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    Set IRiskAssessmentBO_CheckRiskAssessmentExistsForApplication = xmlResponseElem

GetDataExit:

    Set xmlRequestNode = Nothing
    Set objIRiskAssessmentDO = Nothing
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

    Set IRiskAssessmentBO_CheckRiskAssessmentExistsForApplication = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetDataExit

End Function

'IK 19/02/2007 EP2_1429
Private Sub DeleteApplicationUnderwriting(ByVal vxmlRiskAssessmentNode As IXMLDOMNode)
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object

    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "DELETE"
    xmlElem.setAttribute "SCHEMA_NAME", "epsomCRUD"
    xmlElem.setAttribute "ENTITY_REF", "APPLICATIONUNDERWRITING"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONUNDERWRITING")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRiskAssessmentNode.selectSingleNode("APPLICATIONNUMBER").Text
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vxmlRiskAssessmentNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    crudObj.OmRequest xmlDoc.xml
    Set crudObj = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

End Sub


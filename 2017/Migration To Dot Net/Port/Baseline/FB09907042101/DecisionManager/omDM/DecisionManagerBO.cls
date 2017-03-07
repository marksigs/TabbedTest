VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DecisionManagerBO"
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
'MC     11/07/00    Created
'MC     01/09/00    SYS1174 Use default UserID if one not specified in Request
'MC     15/09/00    Only run Credit Check/Risk Assessment for AIP.
'------------------------------------------------------------------------------------------

Option Explicit

Implements IDecisionManagerBO
Implements ObjectControl

'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrROOT_NODE_NAME = "DECISIONMANAGER"

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Private Enum SERVICE_TYPE
    stCREDIT_CHECK
    stRISK_ASSESSMENT
    stOM4TO3DOWNLOAD
End Enum

Public Function RunScoreAndDownload(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' XML structure: <REQUEST>
'                    <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                    <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                    <DOWNLOAD>
'                       <TYPE>AIP</TYPE>
'                       <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                    </DOWNLOAD>
'                </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo RunScoreAndDownloadVbErr

    Const strFunctionName As String = "RunScoreAndDownload"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IDecisionManagerBO_RunScoreAndDownload(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RunScoreAndDownload = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RunScoreAndDownloadExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RunScoreAndDownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    RunScoreAndDownload = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RunScoreAndDownloadExit

End Function

Private Function IsServiceAvailable(vstServiceType As SERVICE_TYPE) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check against GlobalParameter table whether a particular service is available.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo IsServiceAvailableVbErr
    
    Const strFunctionName As String = "IsServiceAvailable"

    
    Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim strParamName As String
    
    Select Case vstServiceType
    Case stCREDIT_CHECK
        strParamName = "CreditCheckAvailable"
    Case stRISK_ASSESSMENT
        strParamName = "RiskAssessmentAvailable"
    Case stOM4TO3DOWNLOAD
        strParamName = "Omiga4To3DownloadAvailable"
    Case Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeNotImplemented, "Specified service type not handled"
    End Select
    
    IsServiceAvailable = objGlobalParamDO.GetCurrentParameterByType( _
        strParamName, "BOOLEAN")
    
    Set objGlobalParamDO = Nothing
    
    Exit Function
    
IsServiceAvailableVbErr:
    
    Set objGlobalParamDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function LogSubmissionError( _
    ByVal vstrDescription As String, _
    ByVal vstServiceType As SERVICE_TYPE)
' header ----------------------------------------------------------------------------------
' description:
'   Log a Submission process error.
' pass:
' return:   n/a
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "LogSubmissionError"

    
    Dim strMsg As String
    Dim strSource As String
    
    Select Case vstServiceType
    Case stCREDIT_CHECK
        strSource = "Credit Check"
    Case stRISK_ASSESSMENT
        strSource = "Risk Assessment"
    Case stOM4TO3DOWNLOAD
        strSource = "Om4To3Download"
    Case Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeNotImplemented, "Specified service type not handled"
    End Select
    
    strSource = App.Title & ".SubmitBO." & strFunctionName
    
    strMsg = _
        strSource & " Submission Failure - " & vbCr & _
        "Details: " & vstrDescription & vbCr
        
    App.LogEvent strMsg, vbLogEventTypeError
    
    ' FIXME Remove this eventually
    Debug.Print strMsg

End Function

Private Function GetSubmissionStatusComboId(ByVal vstrValidation As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the combo id for supplied combo validation value from combo "SubmissionStatus".
' pass:
'   vstrValidation
'       combo validation value.
' return:
'   Combo Id value.
'------------------------------------------------------------------------------------------
On Error GoTo GetSubmissionStatusComboIdVbErr
    
    Const strFunctionName As String = "GetSubmissionStatusComboId"

    
    Dim objComboDO As ComboDO
    
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    GetSubmissionStatusComboId = objComboDO.GetFirstComboValueId( _
        "SubmissionStatus", vstrValidation)
    
    Set objComboDO = Nothing
    
    Exit Function
    
GetSubmissionStatusComboIdVbErr:
    
    Set objComboDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IDecisionManagerBO_RunScoreAndDownload(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo RunScoreAndDownloadVbErr

'XML In Format
' NB: OUTPUTFILE and ASYNC attributes are optional
'
'<REQUEST  USERID="Username" ACTION="SUBMIT" OUTPUTFILE="C:\DOWNLOAD.XML" ASYNC="Yes">
'   <APPLICATIONNUMBER>00063460</APPLICATIONNUMBER>
'   <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   <UNITID>UnitName</UNITID>
'   <USERID>UserName</USERID>
'   <TYPE>AIP</TYPE>
'   <APPLICATIONSTAGE>
'       <APPLICATIONNUMBER>00063460</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <NEWSTAGENUMBER>35</NEWSTAGENUMBER>
'       <STAGENAME>AIP Complete</STAGENAME>
'   </APPLICATIONSTAGE>
'</REQUEST>

    Const strFunctionName As String = "IDecisionManagerBO_RunScoreAndDownload"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Dim xmlDummyElem As IXMLDOMElement
    Dim xmlRequestClone As IXMLDOMElement
    
    Dim objISubmission As ISubmitBO
    Dim objICreditCheck As ICreditCheckBO
    Dim objIRiskAssessment As IRiskAssessmentBO
    Dim objIDownload As IOmiga4ToOmiga3BO
    Dim objIApplication As IApplicationBO
    Dim lngError As Long
    Dim lngCCError As Long
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If m_objContext Is Nothing Then
        Set objISubmission = New SubmitBO
        Set objICreditCheck = New CreditCheckBO
        Set objIRiskAssessment = New RiskAssessmentBO
        Set objIDownload = New Omiga4toOmiga3BO
        Set objIApplication = New ApplicationBO
    Else
        Set objISubmission = m_objContext.CreateInstance(gstrSUBMISSION_COMPONENT & ".SubmitBO")
        Set objICreditCheck = m_objContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
        Set objIRiskAssessment = m_objContext.CreateInstance(gstrRISKASSESSMENT_COMPONENT & ".RiskAssessmentBO")
        Set objIDownload = m_objContext.CreateInstance(gstrDOWNLOAD_COMPONENT & ".Omiga4toOmiga3BO")
        Set objIApplication = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If

    'If no UserID specified, get default one
    If IsNull(vxmlRequest.getAttribute("USERID")) Then
        'Get default UserId
        vxmlRequest.setAttribute "USERID", GetDefaultUserId()
    End If

    'Perform Credit Check (if AIP & available)
    Dim sType As String
    sType = m_objXmlAssist.GetMandatoryElementText(vxmlRequest, ".//TYPE")
    If IsServiceAvailable(stCREDIT_CHECK) And sType = "AIP" Then
        Set xmlTempResponseNode = objICreditCheck.RunCreditCheck(vxmlRequest)
        lngCCError = objICreditCheck.CheckCreditCheckResponse(xmlTempResponseNode, xmlResponseElem, True)
        If lngCCError <> 0 Then
            'Log error...
            LogSubmissionError xmlTempResponseNode.xml, stCREDIT_CHECK
            '...and update Submission Status
            Set xmlSubElement = vxmlRequest.ownerDocument.createElement("APPLICATIONSUBMISSIONSTATUS")
            vxmlRequest.appendChild xmlSubElement
            Set xmlElement = vxmlRequest.ownerDocument.createElement("SUBMISSIONSTATUS")
            xmlElement.Text = GetSubmissionStatusComboId("CCFailed")
            xmlSubElement.appendChild xmlElement
            m_objXmlAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", vxmlRequest, xmlSubElement, False
            m_objXmlAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", vxmlRequest, xmlSubElement, False
            'Attempt to retrieve existing submission record. If it exists update else create
            Set xmlTempResponseNode = objISubmission.GetSubmissionStatusForApplication(vxmlRequest)
            lngError = m_objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem)
            If lngError = 0 Then
                Set xmlTempResponseNode = objISubmission.UpdateSubmissionStatus(vxmlRequest)
            Else
                Set xmlTempResponseNode = objISubmission.CreateSubmissionStatus(vxmlRequest)
            End If
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        End If
    Else
        lngCCError = 1
    End If
    
    'Perform Risk Assessment (if available and Credit Check successful)
    If lngCCError = 0 And IsServiceAvailable(stRISK_ASSESSMENT) Then
        objIRiskAssessment.RunRiskAssessment vxmlRequest, xmlTempResponseNode
        lngError = m_objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem, True)
    End If
    
    'Perform Download to Omiga3 (if available)
    If IsServiceAvailable(stOM4TO3DOWNLOAD) Then
        'The Request needs to be cloned if running asynchronously
        Set xmlRequestClone = vxmlRequest.cloneNode(True)
        
        xmlRequestClone.removeAttribute "ACTION"
        xmlRequestClone.setAttribute "ACTION", "DOWNLOAD"
        Set xmlDummyElem = xmlOut.createElement("PERFORMANCE")
        Set xmlTempResponseNode = objIDownload.Download(xmlRequestClone, xmlDummyElem)
        lngError = m_objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem, True)
    End If
    
    'Update Application Stage
    Set xmlElement = vxmlRequest.selectSingleNode(".//APPLICATIONSTAGE")
    If Not xmlElement Is Nothing Then
        vxmlRequest.setAttribute "ACTION", "CREATE"
        'Rename NewStageNumber node to StageNumber as using StageNumber
        'initially causes an error in Risk Assessment
        m_objXmlAssist.ChangeNodeName vxmlRequest, "NEWSTAGENUMBER", "STAGENUMBER"
        Set xmlTempResponseNode = objIApplication.CreateApplicationStage(vxmlRequest)
        lngError = m_objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem, True)
    End If
    
    Set IDecisionManagerBO_RunScoreAndDownload = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RunScoreAndDownloadExit:
    
    Set objISubmission = Nothing
    Set objICreditCheck = Nothing
    Set objIRiskAssessment = Nothing
    Set objIDownload = Nothing
    Set objIApplication = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRequestClone = Nothing
    
    Exit Function

RunScoreAndDownloadVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IDecisionManagerBO_RunScoreAndDownload = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RunScoreAndDownloadExit


End Function

Private Function GetDefaultUserId() As String

On Error GoTo GetDefaultUserIdVbErr
    
    Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    GetDefaultUserId = objGlobalParamDO.GetCurrentParameterByType("DefaultUserId", "STRING")
    
GetDefaultUserIdVbExit:
    Set objGlobalParamDO = Nothing
    Exit Function
    
GetDefaultUserIdVbErr:
    
    'If error occurs return empty string
    GetDefaultUserId = ""
    GoTo GetDefaultUserIdVbExit
    
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


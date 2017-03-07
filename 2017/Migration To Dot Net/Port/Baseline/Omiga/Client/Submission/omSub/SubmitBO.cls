VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "SubmitBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SubmitBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'RF     27/03/00    Created
'MC     30/06/00    Added GetSubmissionStatusForApplication method
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
Option Explicit

Implements ISubmitBO
Implements ObjectControl

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

Private Const cstrROOT_NODE_NAME = "APPLICATIONSUBMISSIONSTATUS"

Public Function CreateSubmissionStatus(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Create a single instance of the persistant data associated with this
'   business object
' pass:
'   vstrXmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubmissionStatusVbErr

    Const strFunctionName As String = "CreateSubmissionStatus"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
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
    
    ' Run validation checks
    Set xmlTempResponseNode = ISubmitBO_Validate( _
        xmlIn.documentElement, bomCreateSubmissionStatus)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), _
        strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ISubmitBO_CreateSubmissionStatus(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateSubmissionStatus = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateSubmissionStatusExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateSubmissionStatusVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateSubmissionStatus = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateSubmissionStatusExit

End Function

Private Function ISubmitBO_CreateSubmissionStatus(ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create a single instance of the persistant data associated with this
'   business object. Invokes appropriate transactioning (Tx) business object
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubmissionStatusVbErr
    
    Const strFunctionName As String = "ISubmitBO_CreateSubmissionStatus"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objISubmitTxBO As ISubmitTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objISubmitTxBO = New SubmitTxBO
    Else
        Set objISubmitTxBO = m_objContext.CreateInstance(App.Title & ".SubmitTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objISubmitTxBO.CreateSubmissionStatus(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ISubmitBO_CreateSubmissionStatus = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateSubmissionStatusExit:

    Set objISubmitTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateSubmissionStatusVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_CreateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateSubmissionStatusExit

End Function

Public Function DoCreditCheck(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest
'       Format:
'           <REQUEST></REQUEST>
' return:
'   xml Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo DoCreditCheckVbErr

    Const strFunctionName As String = "DoCreditCheck"

        
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
    Set xmlTempResponseNode = ISubmitBO_DoCreditCheck(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DoCreditCheck = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DoCreditCheckExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DoCreditCheckVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Dim strResponse As String
    
    strResponse = m_objErrAssist.CreateErrorResponse
    
    LogSubmissionError strResponse, stCREDIT_CHECK
    
    DoCreditCheck = strResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume DoCreditCheckExit

End Function

Private Function ISubmitBO_DoCreditCheck(ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitBO_DoCreditCheckVbErr
    
    Const strFunctionName As String = "ISubmitBO_DoCreditCheck"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCCResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------------------------------------------
    ' check if Credit Check service is available
    '------------------------------------------------------------------------------------------
    
    If IsServiceAvailable(stCREDIT_CHECK) = True Then
    
        '------------------------------------------------------------------------------------------
        ' look for an Application to process
        '------------------------------------------------------------------------------------------
        
        Dim xmlFindAppResponseNode As IXMLDOMNode
        
        Err.Clear
        On Error Resume Next
        Set xmlFindAppResponseNode = FindApplication(vxmlRequest, stCREDIT_CHECK)
        m_objErrAssist.SaveErr
        On Error GoTo ISubmitBO_DoCreditCheckVbErr
        
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> _
                oeRecordNotFound Then
                m_objErrAssist.ReRaise
            ' Else nothing to process
            End If
        Else
        
            Dim objISubmitTxBO As ISubmitTxBO
            
            If m_objContext Is Nothing Then
                Set objISubmitTxBO = New SubmitTxBO
            Else
                Set objISubmitTxBO = _
                    m_objContext.CreateInstance(App.Title & ".SubmitTxBO")
            End If
            
            '------------------------------------------------------------------------------------------
            ' mark Application as having processing "in progress"
            '------------------------------------------------------------------------------------------
            
            Set xmlTempNode = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
            Set xmlElem = xmlDoc.createElement("APPLICATIONSUBMISSIONSTATUS")
            Set xmlNode = xmlTempNode.appendChild(xmlElem)
            
            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONNUMBER").cloneNode(True)

            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
                
            Set xmlElem = xmlDoc.createElement("USERID")
            xmlElem.Text = m_objXmlAssist.GetMandatoryAttribute( _
                xmlTempNode, "USERID")
            xmlNode.appendChild xmlElem

            Set xmlElem = xmlDoc.createElement("UNITID")
            xmlElem.Text = m_objXmlAssist.GetMandatoryAttribute( _
                xmlTempNode, "UNITID")
            xmlNode.appendChild xmlElem

            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            xmlElem.Text = GetSubmissionStatusComboId("CCInProgress")
            xmlNode.appendChild xmlElem
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            xmlNode.removeChild xmlElem
            
            '------------------------------------------------------------------------------------------
            ' do the credit check
            '------------------------------------------------------------------------------------------
        
            m_objXmlAssist.ChangeNodeName xmlNode, "APPLICATIONSUBMISSIONSTATUS", "CREDITCHECK"
            
            Set xmlCCResponseNode = _
                objISubmitTxBO.DoCreditCheck(xmlTempNode)
            
            Dim lngErrNo As Long
            lngErrNo = m_objErrAssist.CheckXMLResponse( _
                xmlCCResponseNode.xml, False)
        
            '------------------------------------------------------------------------------------------
            ' Update the ApplicationSubmissionStatus.
            '------------------------------------------------------------------------------------------
            
            m_objXmlAssist.ChangeNodeName xmlNode, "CREDITCHECK", "APPLICATIONSUBMISSIONSTATUS"
            
            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            
            If lngErrNo = 0 Then
                xmlElem.Text = GetSubmissionStatusComboId("CCComplete")
            Else
                xmlElem.Text = GetSubmissionStatusComboId("CCFailed")
                LogSubmissionError xmlCCResponseNode.xml, stCREDIT_CHECK
            End If
            
            xmlNode.appendChild xmlElem
            
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        End If
        
    End If
        
    Set ISubmitBO_DoCreditCheck = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ISubmitBO_DoCreditCheckExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCCResponseNode = Nothing
    Set objISubmitTxBO = Nothing
    
    Exit Function

ISubmitBO_DoCreditCheckVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_DoCreditCheck = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ISubmitBO_DoCreditCheckExit

End Function

Public Function DoOmiga4To3Download(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest
'       Format:
'           <REQUEST>
'               <TYPE></TYPE>    (Valid values are "AIP" or "FormFill")
'           </REQUEST>
' return:
'   xml Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo DoOmiga4To3DownloadVbErr

    Const strFunctionName As String = "DoOmiga4To3Download"

        
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
    Set xmlTempResponseNode = ISubmitBO_DoOmiga4To3Download(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DoOmiga4To3Download = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DoOmiga4To3DownloadExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DoOmiga4To3DownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Dim strResponse As String
    
    strResponse = m_objErrAssist.CreateErrorResponse
    
    LogSubmissionError strResponse, stOM4TO3DOWNLOAD
    
    DoOmiga4To3Download = strResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DoOmiga4To3DownloadExit

End Function

Private Function ISubmitBO_DoOmiga4To3Download(ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request data stream.
'       Format:
'           <REQUEST>
'               <TYPE></TYPE>    (Valid values are "AIP" or "FormFill")
'           </REQUEST>
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitBO_DoOmiga4To3DownloadVbErr
    
    Const strFunctionName As String = "ISubmitBO_DoOmiga4To3Download"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlDownloadResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------------------------------------------
    ' check if OM4TO3DOWNLOAD service is available
    '------------------------------------------------------------------------------------------
    
    If IsServiceAvailable(stOM4TO3DOWNLOAD) = True Then
    
        '------------------------------------------------------------------------------------------
        ' look for an Application to process
        '------------------------------------------------------------------------------------------
        
        Dim xmlFindAppResponseNode As IXMLDOMNode
        Dim strDownloadType As String
        
        strDownloadType = m_objXmlAssist.GetMandatoryElementText(vxmlRequest, ".//TYPE")
        
        Err.Clear
        On Error Resume Next
        Set xmlFindAppResponseNode = _
            FindApplication(vxmlRequest, stOM4TO3DOWNLOAD, strDownloadType)
        m_objErrAssist.SaveErr
        On Error GoTo ISubmitBO_DoOmiga4To3DownloadVbErr
        
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> _
                oeRecordNotFound Then
                m_objErrAssist.ReRaise
            ' Else nothing to process
            End If
        Else
        
            Dim objISubmitTxBO As ISubmitTxBO
            
            If m_objContext Is Nothing Then
                Set objISubmitTxBO = New SubmitTxBO
            Else
                Set objISubmitTxBO = _
                    m_objContext.CreateInstance(App.Title & ".SubmitTxBO")
            End If
            
            '------------------------------------------------------------------------------------------
            ' mark Application as having processing "in progress"
            '------------------------------------------------------------------------------------------
            
            Set xmlTempNode = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
            Set xmlElem = xmlDoc.createElement("APPLICATIONSUBMISSIONSTATUS")
            Set xmlNode = xmlTempNode.appendChild(xmlElem)
            
            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONNUMBER").cloneNode(True)

            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONFACTFINDNUMBER").cloneNode(True)

            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            
            Select Case strDownloadType
            Case "AIP"
                xmlElem.Text = GetSubmissionStatusComboId("AIPDLInProgress")
            Case "FormFill"
                xmlElem.Text = GetSubmissionStatusComboId("MADLInProgress")
            Case Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, "Download type is not valid: " & strDownloadType
            End Select

            xmlNode.appendChild xmlElem
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            xmlNode.removeChild xmlElem
            
            '------------------------------------------------------------------------------------------
            ' do the download
            '------------------------------------------------------------------------------------------
            
            m_objXmlAssist.ChangeNodeName xmlNode, "APPLICATIONSUBMISSIONSTATUS", "DOWNLOAD"
            
            Set xmlElem = xmlDoc.createElement("TYPE")
            xmlElem.Text = strDownloadType
            xmlNode.appendChild xmlElem
        
            Set xmlDownloadResponseNode = _
                objISubmitTxBO.DoOmiga4To3Download(xmlTempNode)
            
            Dim lngErrNo As Long
            lngErrNo = m_objErrAssist.CheckXMLResponse( _
                xmlDownloadResponseNode.xml, False)
            
            '------------------------------------------------------------------------------------------
            ' Update the ApplicationSubmissionStatus.
            '------------------------------------------------------------------------------------------
            
            m_objXmlAssist.ChangeNodeName xmlNode, "DOWNLOAD", "APPLICATIONSUBMISSIONSTATUS"
            
            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            
            Select Case strDownloadType
            Case "AIP"
                If lngErrNo = 0 Then
                    xmlElem.Text = GetSubmissionStatusComboId("AIPDLComplete")
                Else
                    xmlElem.Text = GetSubmissionStatusComboId("AIPDLFailed")
                    LogSubmissionError xmlDownloadResponseNode.xml, stOM4TO3DOWNLOAD
                End If
            Case "FormFill"
                If lngErrNo = 0 Then
                    xmlElem.Text = GetSubmissionStatusComboId("MADLComplete")
                Else
                    xmlElem.Text = GetSubmissionStatusComboId("MADLFailed")
                    LogSubmissionError xmlDownloadResponseNode.xml, stOM4TO3DOWNLOAD
                End If
            Case Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, "Download type is not valid: " & strDownloadType
            End Select
            
            xmlNode.appendChild xmlElem
            
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        End If
        
    End If
        
    Set ISubmitBO_DoOmiga4To3Download = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ISubmitBO_DoOmiga4To3DownloadExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDownloadResponseNode = Nothing
    Set objISubmitTxBO = Nothing
    
    Exit Function

ISubmitBO_DoOmiga4To3DownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_DoOmiga4To3Download = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ISubmitBO_DoOmiga4To3DownloadExit

End Function

Public Function DoRiskAssessment(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest
'       Format:
'           <REQUEST></REQUEST>
' return:
'   xml Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo DoRiskAssessmentVbErr

    Const strFunctionName As String = "DoRiskAssessment"

        
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
    Set xmlTempResponseNode = ISubmitBO_DoRiskAssessment(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DoRiskAssessment = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DoRiskAssessmentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DoRiskAssessmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Dim strResponse As String
    
    strResponse = m_objErrAssist.CreateErrorResponse
    
    LogSubmissionError strResponse, stRISK_ASSESSMENT
    
    DoRiskAssessment = strResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume DoRiskAssessmentExit

End Function

Private Function ISubmitBO_DoRiskAssessment( _
    ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request data stream.
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitBO_DoRiskAssessmentVbErr
    
    Const strFunctionName As String = "ISubmitBO_DoRiskAssessment"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlRAResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------------------------------------------
    ' check if RiskAssessment service is available
    '------------------------------------------------------------------------------------------
    
    If IsServiceAvailable(stRISK_ASSESSMENT) = True Then
    
        '------------------------------------------------------------------------------------------
        ' look for an Application to process
        '------------------------------------------------------------------------------------------
        
        Dim xmlFindAppResponseNode As IXMLDOMNode
        
        Err.Clear
        On Error Resume Next
        Set xmlFindAppResponseNode = FindApplication(vxmlRequest, stRISK_ASSESSMENT)
        m_objErrAssist.SaveErr
        On Error GoTo ISubmitBO_DoRiskAssessmentVbErr
        
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> _
                oeRecordNotFound Then
                m_objErrAssist.ReRaise
            ' Else nothing to process
            End If
        Else
        
            Dim objISubmitTxBO As ISubmitTxBO
            
            If m_objContext Is Nothing Then
                Set objISubmitTxBO = New SubmitTxBO
            Else
                Set objISubmitTxBO = _
                    m_objContext.CreateInstance(App.Title & ".SubmitTxBO")
            End If
            
            '------------------------------------------------------------------------------------------
            ' mark Application as having processing "in progress"
            '------------------------------------------------------------------------------------------
            
            Set xmlTempNode = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
            Set xmlElem = xmlDoc.createElement("APPLICATIONSUBMISSIONSTATUS")
            Set xmlNode = xmlTempNode.appendChild(xmlElem)
            
            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONNUMBER").cloneNode(True)

            xmlNode.appendChild m_objXmlAssist.GetMandatoryNode( _
                xmlFindAppResponseNode, ".//APPLICATIONFACTFINDNUMBER").cloneNode(True)

            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            
            xmlElem.Text = GetSubmissionStatusComboId("RAInProgress")
            
            xmlNode.appendChild xmlElem
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            xmlNode.removeChild xmlElem
            
            '------------------------------------------------------------------------------------------
            ' do the risk assessment
            '------------------------------------------------------------------------------------------
        
            m_objXmlAssist.ChangeNodeName xmlNode, "APPLICATIONSUBMISSIONSTATUS", "RISKASSESSMENT"
            
            Set xmlRAResponseNode = _
                objISubmitTxBO.DoRiskAssessment(xmlTempNode)
        
            Dim lngErrNo As Long
            lngErrNo = m_objErrAssist.CheckXMLResponse(xmlRAResponseNode.xml, False)
            
            '------------------------------------------------------------------------------------------
            ' Update the ApplicationSubmissionStatus.
            '------------------------------------------------------------------------------------------
            
            m_objXmlAssist.ChangeNodeName xmlNode, "RISKASSESSMENT", "APPLICATIONSUBMISSIONSTATUS"
            
            Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
            
            If lngErrNo = 0 Then
                xmlElem.Text = GetSubmissionStatusComboId("RAComplete")
            Else
                xmlElem.Text = GetSubmissionStatusComboId("RAFailed")
                LogSubmissionError xmlRAResponseNode.xml, stRISK_ASSESSMENT
            End If
            
            xmlNode.appendChild xmlElem
            
            Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(xmlNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        End If
        
    End If
        
    Set ISubmitBO_DoRiskAssessment = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ISubmitBO_DoRiskAssessmentExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRAResponseNode = Nothing
    Set objISubmitTxBO = Nothing
    
    Exit Function

ISubmitBO_DoRiskAssessmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_DoRiskAssessment = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ISubmitBO_DoRiskAssessmentExit

End Function

Private Function ISubmitBO_GetSubmissionStatusForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetSubmissionStatusForApplicationVbErr
    
    Const strFunctionName As String = "ISubmitBO_GetSubmissionStatusForApplication"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objResponse As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objISubmitDO As ISubmitDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objISubmitDO = New SubmitDO
    Else
        Set objISubmitDO = m_objContext.CreateInstance(App.Title & ".SubmitDO")
    End If
    
    Set objResponse = objISubmitDO.GetSubmissionStatusForApplication(vxmlTableElement)
    'm_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    xmlResponseElem.appendChild objResponse
    
    Set ISubmitBO_GetSubmissionStatusForApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetSubmissionStatusForApplicationExit:

    Set objISubmitDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

GetSubmissionStatusForApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_GetSubmissionStatusForApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetSubmissionStatusForApplicationExit

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

Private Function ISubmitBO_Validate( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal veboMethod As ISubmitBOMethod) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   performs hard coded validation for this object
' pass:
'   vxmlRequest
'       xml Request data stream containing details of action to be performed and
'       data required. This is the full request block as received from the client.
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ISubmitBO_Validate"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objISubmitDO As ISubmitDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objISubmitDO = New SubmitDO
    Else
        Set objISubmitDO = m_objContext.CreateInstance(App.Title & ".SubmitDO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
     
    'objISubmitDO.AddDerivedData xmlRequestNode
            
    ' TODO ************************************************************************
    '   add hard coded validation here, & raise error if it fails
    ' *****************************************************************************
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set ISubmitBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objISubmitDO = Nothing
    
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

    Set ISubmitBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function


Public Function UpdateSubmissionStatus(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this business object.
' pass:
'   vstrXmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubmissionStatusVbErr

    Const strFunctionName As String = "UpdateSubmissionStatus"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
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
    
    ' Run validation checks
    Set xmlTempResponseNode = ISubmitBO_Validate( _
        xmlIn.documentElement, bomUpdateSubmissionStatus)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), _
        strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ISubmitBO_UpdateSubmissionStatus(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateSubmissionStatus = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateSubmissionStatusExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateSubmissionStatusVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateSubmissionStatus = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateSubmissionStatusExit
    
End Function


Private Function ISubmitBO_UpdateSubmissionStatus(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this business object.
'   Invokes appropriate transactioning (Tx) business object.
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubmissionStatusVbErr
    
    Const strFunctionName As String = "ISubmitBO_UpdateSubmissionStatus"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objISubmitTxBO As ISubmitTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objISubmitTxBO = New SubmitTxBO
    Else
        Set objISubmitTxBO = m_objContext.CreateInstance(App.Title & ".SubmitTxBO")
    End If
    
    Set xmlTempResponseNode = objISubmitTxBO.UpdateSubmissionStatus(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ISubmitBO_UpdateSubmissionStatus = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateSubmissionStatusExit:

    Set objISubmitTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateSubmissionStatusVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitBO_UpdateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateSubmissionStatusExit

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
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function FindApplication( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vstServiceType As SERVICE_TYPE, _
    Optional ByVal vstrDownloadType As String = "") _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Find an Application requiring processing via the specified service.
' pass:
' return:
'   True if an applicable application has been found, otherwise False.
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationVbErr
    
    Const strFunctionName As String = "FindApplication"

    
    Dim objISubmitDO As ISubmitDO
    
    If m_objContext Is Nothing Then
        Set objISubmitDO = New SubmitDO
    Else
        Set objISubmitDO = m_objContext.CreateInstance(App.Title & ".SubmitDO")
    End If
    
    Dim strSubmissionStatus As String
    
    Select Case vstServiceType
    Case stCREDIT_CHECK
        strSubmissionStatus = GetSubmissionStatusComboId("AIPImpComplete")
    Case stRISK_ASSESSMENT
        strSubmissionStatus = GetSubmissionStatusComboId("CCComplete")
     Case stOM4TO3DOWNLOAD
        Select Case vstrDownloadType
        Case "AIP"
            strSubmissionStatus = GetSubmissionStatusComboId("RAComplete")
        Case "FormFill"
            strSubmissionStatus = GetSubmissionStatusComboId("AIPDLComplete")
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInvalidParameter, "Download type is not valid: " & vstrDownloadType
        End Select
    Case Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeNotImplemented, "Specified service type not handled"
    End Select
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlDoc.documentElement = vxmlRequest.cloneNode(True)
    
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = m_objXmlAssist.GetRequestNodeEx(xmlDoc.documentElement)
    
    Dim xmlElem As IXMLDOMElement
    Set xmlElem = xmlDoc.createElement("SUBMISSIONSTATUS")
    xmlElem.Text = strSubmissionStatus
    xmlNode.appendChild xmlElem

    Set FindApplication = objISubmitDO.GetSubmissionStatusData(xmlNode)
    
    Set objISubmitDO = Nothing
    Set xmlDoc = Nothing
    
    Exit Function
    
FindApplicationVbErr:
    
    Set objISubmitDO = Nothing
    Set xmlDoc = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetComplete
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
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


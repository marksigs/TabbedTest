VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "SubmitTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SubmitTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:
'
'Dependencies:  SubmitDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     04/04/00    Created.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ISubmitTxBO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================

Private Const cstrROOT_NODE_NAME = "APPLICATIONSUBMISSIONSTATUS"

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
'Private m_objXmlAssist As XMLAssist             ' not required

Private Function ISubmitTxBO_CreateSubmissionStatus( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create multiple instances of the persistant data associated with this
'   business object. Invokes appropriate Data Object(s) to perform persistance.
' pass:
'   vxmlRequest
'       xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubmissionStatusVbErr

    Const strFunctionName As String = "ISubmitTxBO_CreateSubmissionStatus"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objISubmitDO As ISubmitDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objISubmitDO = New SubmitDO
    Else
        Set objISubmitDO = m_objContext.CreateInstance(App.Title & ".SubmitDO")
    End If
    
    objISubmitDO.CreateSubmissionStatus xmlElement
    
    Set ISubmitTxBO_CreateSubmissionStatus = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateSubmissionStatusExit:

    Set objISubmitDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set ISubmitTxBO_CreateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateSubmissionStatusExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    'Set m_objXmlAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    'Set m_objXmlAssist = Nothing

End Sub

Private Function ISubmitTxBO_UpdateSubmissionStatus( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Update multiple instances of the persistant data associated with this business object.
'   Invokes appropriate Data Object(s) to perform persistance.
' pass:
'   vxmlRequest
'       xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubmissionStatusVbErr

    Const strFunctionName As String = "ISubmitTxBO_UpdateSubmissionStatus"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objISubmitDO As ISubmitDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objISubmitDO = New SubmitDO
    Else
        Set objISubmitDO = m_objContext.CreateInstance(App.Title & ".SubmitDO")
    End If
    
    objISubmitDO.UpdateSubmissionStatus xmlElement
    
    Set ISubmitTxBO_UpdateSubmissionStatus = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateSubmissionStatusExit:
    
    Set objISubmitDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set ISubmitTxBO_UpdateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateSubmissionStatusExit

End Function

Private Function ISubmitTxBO_DoRiskAssessment( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request element
'       Format:
'           <REQUEST>
'               <RISKASSESSMENT>
'                   <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'               </RISKASSESSMENT>
'           </REQUEST>
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitTxBO_DoRiskAssessmentVbErr

    Const strFunctionName As String = "ISubmitTxBO_DoRiskAssessment"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    Dim objRiskAssBO As RiskAssessmentBO
    
    If m_objContext Is Nothing Then
        Set objRiskAssBO = New RiskAssessmentBO
    Else
        Set objRiskAssBO = _
            m_objContext.CreateInstance(gstrRISK_ASSESSMENT & ".RiskAssessmentBO")
    End If
    
    Dim strResponse As String
    strResponse = objRiskAssBO.RunRiskAssessment(vxmlRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    
    Set ISubmitTxBO_DoRiskAssessment = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
    
ISubmitTxBO_DoRiskAssessmentExit:

    Set objRiskAssBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ISubmitTxBO_DoRiskAssessmentVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitTxBO_DoRiskAssessment = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume ISubmitTxBO_DoRiskAssessmentExit

End Function


Private Function ISubmitTxBO_DoCreditCheck( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request element
'       Format:
'           <REQUEST>
'               <CREDITCHECK>
'                   <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                   <USERID></USERID>
'                   <UNITID<>/UNITID>
'               </CREDITCHECK>
'           </REQUEST>
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitTxBO_DoCreditCheckVbErr

    Const strFunctionName As String = "ISubmitTxBO_DoCreditCheck"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objCreditCheckBO As CreditCheckBO
    
    If m_objContext Is Nothing Then
        Set objCreditCheckBO = New CreditCheckBO
    Else
        Set objCreditCheckBO = _
            m_objContext.CreateInstance(gstrCREDIT_CHECK & ".CreditCheckBO")
    End If
    
    Dim strResponse As String
    strResponse = objCreditCheckBO.RunCreditCheck(vxmlRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True

    Set ISubmitTxBO_DoCreditCheck = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
    
ISubmitTxBO_DoCreditCheckExit:

    Set objCreditCheckBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ISubmitTxBO_DoCreditCheckVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitTxBO_DoCreditCheck = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume ISubmitTxBO_DoCreditCheckExit

End Function


Private Function ISubmitTxBO_DoOmiga4To3Download( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request element
'       Format:
'           <REQUEST>
'               <DOWNLOAD>
'                   <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                   <TYPE></TYPE>    (Valid values are "AIP" or "FormFill")
'               </DOWNLOAD>
'           </REQUEST>
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ISubmitTxBO_DoOmiga4To3DownloadVbErr

    Const strFunctionName As String = "ISubmitTxBO_DoOmiga4To3Download"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    Dim objOmiga4toOmiga3BO As Omiga4toOmiga3BO
    
    If m_objContext Is Nothing Then
        Set objOmiga4toOmiga3BO = New Omiga4toOmiga3BO
    Else
        Set objOmiga4toOmiga3BO = _
            m_objContext.CreateInstance(gstrOMIGA4ToOMIGA3DOWNLOAD & ".Omiga4toOmiga3BO")
    End If
    
    vxmlRequest.setAttribute "ACTION", "DOWNLOAD"
    
    Dim strResponse As String
    strResponse = objOmiga4toOmiga3BO.Download(vxmlRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    
    Set ISubmitTxBO_DoOmiga4To3Download = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ISubmitTxBO_DoOmiga4To3DownloadExit:

    Set objOmiga4toOmiga3BO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ISubmitTxBO_DoOmiga4To3DownloadVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ISubmitTxBO_DoOmiga4To3Download = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume ISubmitTxBO_DoOmiga4To3DownloadExit

End Function




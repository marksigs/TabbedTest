VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AuditTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AuditTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Transaction Business objects class for omAU
'
'Dependencies:  List any other dependent components
'               e.g. AuditDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date         Description
'RF     30/09/99    Created
'RF     30/11/99    Change MTSTransactionMode from RequiresNewTransaction to RequiresTransaction
'                   because of problem in locking applications
'CL     18/10/00    core00004 Modifications made to conform to coding templates
'MV     06/03/01    SYS2001: Commenting in CreateAccessAudit
'------------------------------------------------------------------------------------------
Option Explicit
Implements IAuditTxBO
Implements ObjectControl
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private Const cstrROOT_NODE_NAME = "ACCESSAUDIT"
Private Function IAuditTxBO_CreateAccessAudit(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccessAuditVbErr
    Const strFunctionName As String = "IAuditTxBO_CreateAccessAudit"
    'Check to see that MTS is up and running
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    'Dim xmlTempResponseElem As IXMLDOMElement
    Dim objIAuditDO As IAuditDO
    Set xmlOut = New FreeThreadedDOMDocument40
    'Create a new DOM document with RESPONSE node with attribute TYPE as SUCCESS
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
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
    'Open an instance of AuditDO and view through the IAuditDO interface
    Set objIAuditDO = m_objContext.CreateInstance(App.Title & ".AuditDO")
    objIAuditDO.CreateAccessAudit xmlElement
    'Set xmlTempResponseElem = objIAuditDO.CreateAccessAudit(xmlElement)
    'xmlResponseElem.appendChild xmlTempResponseElem
    Set IAuditTxBO_CreateAccessAudit = xmlResponseElem
    m_objContext.SetComplete
CreateAccessAuditExit:
    Set objIAuditDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
CreateAccessAuditVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IAuditTxBO_CreateAccessAudit = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    ' go to clean-up section
    Resume CreateAccessAuditExit
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

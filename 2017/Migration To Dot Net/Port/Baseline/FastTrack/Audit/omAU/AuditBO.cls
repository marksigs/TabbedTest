VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AuditBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AuditBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Business objects class for omAU
'
'Dependencies:  AuditTxBO, AuditDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     30/09/99    Created based on template version 30/09/99.
'RF     27/01/00    Pick up ScriptInterface object from omBase.
'MC     16/05/00    SYS0210 - Synchronise the password change date/time with corresponding
'                   access audit record
'CL     18/10/00    Core00004 Modifications made to conform to coding templates
'MV     06/03/01    SYS2001: changed  the Return Data from xmlTempResponseNode.xml to xmlResponseElem.xml
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        Description
'TK     30/11/2004  E2EM00002504 - Performance related fixes.
'------------------------------------------------------------------------------------------
Option Explicit
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Implements IAuditBO
Implements ObjectControl
Private Const cstrROOT_NODE_NAME = "ACCESSAUDIT"
 
Private Function IAuditBO_Validate(ByVal vxmlRequest As IXMLDOMElement, _
                                   ByVal veboMethod As IAuditBOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr
    Const strFunctionName As String = "IAuditBO_Validate"
    'If MTS is not running then throw an error
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
'    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim strUserId As String
    Dim strAUDITRECORDTYPE As String
    Dim strSUCCESSINDICATOR As String
    'Create a new DOM document and add a response node "SUCCESS"
    'this will be used if we have a success from the DO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    'Find the node REQUEST within the imported xml
    If vxmlRequest.nodeName = "REQUEST" Then
        Set xmlRequestElem = vxmlRequest
    Else
        Set xmlRequestElem = m_objXmlAssist.GetNode(vxmlRequest, ".//REQUEST")
    End If
            
    'Get the value of the attribute USERID withing the REQUEST node
    strUserId = m_objXmlAssist.GetAttributeValue(xmlRequestElem, "REQUEST", "USERID")
    'If no value is found in the above node then throw an error
    If Len(strUserId) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected USERID attribute in REQUEST tag"
    End If
        
    'Get the text of the AUDITRECORDTYPE node
    strAUDITRECORDTYPE = m_objXmlAssist.GetNodeValue(xmlRequestElem, ".//AUDITRECORDTYPE")
    'Get the text of the SUCCESSINDICATOR node
    strAUDITRECORDTYPE = m_objXmlAssist.GetNodeValue(xmlRequestElem, ".//SUCCESSINDICATOR")
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Set IAuditBO_Validate = xmlResponseElem
        
          
ValidateExit:
'    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
       
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
    Set IAuditBO_Validate = m_objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' go to clean-up section
    Resume ValidateExit
End Function
Public Function CreateAccessAudit(ByVal vstrRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Create an accessaudit record.
'   For a change to password, a change password record is also created.
'   For an application access or release, an applicationaccess record is also created.
' pass:
'   vstrRequest
'       xml Request data stream containing data to be persisted.
'       USERID and MACHINEID from REQUEST element are passed down to DO.
'       Format:
'        <REQUEST USERID= MACHINEID=>
'            <ACCESSAUDIT>
'                <AUDITRECORDTYPE></AUDITRECORDTYPE>
'                <SUCCESSINDICATOR></SUCCESSINDICATOR>
'                <ONBEHALFOFUSERID>Optional</ONBEHALFOFUSERID>
'                <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                <PASSWORDCREATIONDATE>Optional</PASSWORDCREATIONDATE>
'            </ACCESSAUDIT>
'        </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccessAuditVbErr
    
    Dim strFunctionName As String
    strFunctionName = "CreateAccessAudit"
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
         
    Set xmlIn = m_objXmlAssist.load(vstrRequest, TypeName(Me), strFunctionName)
    Set xmlTempResponseNode = IAuditBO_Validate(xmlIn.documentElement, bomCreateAccessAudit)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    '------------------------------------------------------------------------------------------
    ' create the record(s) using the updated xml request string
    '------------------------------------------------------------------------------------------
    'Using the IAuditBO interface ?????????
    Set xmlTempResponseNode = IAuditBO_CreateAccessAudit(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    'CreateAccessAudit = xmlTempResponseNode.xml
    CreateAccessAudit = xmlResponseElem.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
           
CreateAccessAuditExit:
    
    'clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
           
                
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
    CreateAccessAudit = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    'go to clean-up section
    Resume CreateAccessAuditExit
End Function
Public Function GetNumberOfFailedAttempts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfFailedAttemptsVbErr
    Const strFunctionName As String = "GetNumberOfFailedAttempts"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IAuditBO_GetNumberOfFailedAttempts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetNumberOfFailedAttempts = xmlResponseElem.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
GetNumberOfFailedAttemptsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
GetNumberOfFailedAttemptsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetNumberOfFailedAttempts = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ' go to clean-up section
    Resume GetNumberOfFailedAttemptsExit
End Function
Private Function IAuditBO_CreateAccessAudit(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create an accessaudit record.
'   For a change to password, a change password record is also created.
'   For an application access or release, an applicationaccess record is also created.
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted.
'       USERID and MACHINEID from REQUEST element are passed down to DO.
'       Format:
'        <REQUEST USERID= MACHINEID=>
'            <ACCESSAUDIT>
'                <AUDITRECORDTYPE></AUDITRECORDTYPE>
'                <SUCCESSINDICATOR></SUCCESSINDICATOR>
'                <ONBEHALFOFUSERID>Optional</ONBEHALFOFUSERID>
'                <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                <PASSWORDCREATIONDATE>Optional</PASSWORDCREATIONDATE>
'            </ACCESSAUDIT>
'        </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccessAuditVbErr
    
    Const strFunctionName As String = "CreateAccessAudit"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAuditTxBO As IAuditTxBO
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    '------------------------------------------------------------------------------------------
    ' Add USERID and MACHINEID from REQUEST element to the body of the request.
    ' MACHINEID is optional.
    '------------------------------------------------------------------------------------------
    Dim strMachineId As String
    Dim strUserId As String
    'Get the value of attribute MACHINEID in node REQUEST
    strMachineId = m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "MACHINEID")
    'Get the value of the attribute USERID withing the REQUEST node
    strUserId = m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "USERID")
     
    Set xmlTableNode = m_objXmlAssist.GetNode(vxmlRequest, ".//ACCESSAUDIT")
    Set xmlElem = vxmlRequest.ownerDocument.createElement("USERID")
    'Make the text of the new element equal to the text of item
    xmlElem.Text = strUserId
    xmlTableNode.appendChild xmlElem
    'If there is a Machine ID then enter this into the MACHINEID node
    If Len(strMachineId) > 0 Then
        Set xmlElem = vxmlRequest.ownerDocument.createElement("MACHINEID")
        xmlElem.Text = strMachineId
        xmlTableNode.appendChild xmlElem
    End If
    If m_objContext Is Nothing Then
        Set objIAuditTxBO = New AuditTxBO
    Else
        Set objIAuditTxBO = m_objContext.CreateInstance(App.Title & ".AuditTxBO")
    End If
    ' call Business Transaction Object Create function
    Set xmlTempResponseNode = objIAuditTxBO.CreateAccessAudit(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    'Set IAuditBO_CreateAccessAudit = xmlTempResponseNode
    Set IAuditBO_CreateAccessAudit = xmlResponseElem
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
CreateAccessAuditExit:
 
    Set objIAuditTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlElem = Nothing
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
    Set IAuditBO_CreateAccessAudit = m_objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' go to clean-up section
    Resume CreateAccessAuditExit
End Function
Private Function IAuditBO_GetNumberOfFailedAttempts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfFailedAttemptsVbErr
    
    Const strFunctionName As String = "IAuditBO_GetNumberOfFailedAttempts"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIAuditDO As IAuditDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIAuditDO = New AuditDO
    Else
        Set objIAuditDO = m_objContext.CreateInstance(App.Title & ".AuditDO")
    End If
    Dim lngNoAttempts As Long
            
    lngNoAttempts = objIAuditDO.GetNumberOfFailedAttempts(vxmlRequest)
    Set xmlDataNode = xmlOut.createElement("NUMBEROFATTEMPTS")
    xmlDataNode.Text = CStr(lngNoAttempts)
    xmlResponseElem.appendChild xmlDataNode
         
    Set IAuditBO_GetNumberOfFailedAttempts = xmlResponseElem
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
GetNumberOfFailedAttemptsExit:
    
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIAuditDO = Nothing
    Exit Function
GetNumberOfFailedAttemptsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IAuditBO_GetNumberOfFailedAttempts = m_objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ' go to clean-up section
    Resume GetNumberOfFailedAttemptsExit
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
    Set m_objXmlAssist = Nothing
    Set m_objErrAssist = Nothing
End Sub

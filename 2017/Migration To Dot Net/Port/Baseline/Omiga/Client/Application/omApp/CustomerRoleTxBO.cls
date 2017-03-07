VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CustomerRoleTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerRoleTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Dependencies:  CustomerRoleDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     07/02/00    Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerRoleTxBO
Implements ObjectControl
 
Private Const cstrROOT_NODE_NAME = "CUSTOMERROLE"

Private Const cstrCOMPONENT_NAME = "omApp"
Private Const cstrCUSTOMER_COMPONENET = "omCust"

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist

Private Function ICustomerRoleTxBO_Create(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ICustomerRoleTxBO_Create"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
   
    Dim objICustomerRoleDO As ICustomerRoleDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerRoleDO = New CustomerRoleDO
    Else
        Set objICustomerRoleDO = m_objContext.CreateInstance(App.Title & ".CustomerRoleDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerRoleDO.Create xmlElement
        CreateLock xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objICustomerRoleDO.Create xmlElement
            CreateLock xmlElement
        Next xmlElement
    End If
    
    Set ICustomerRoleTxBO_Create = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objICustomerRoleDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerRoleTxBO_Create = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function ICustomerRoleTxBO_Delete(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "ICustomerRoleTxBO_Delete"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
   
    Dim objICustomerRoleDO As ICustomerRoleDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerRoleDO = New CustomerRoleDO
    Else
        Set objICustomerRoleDO = m_objContext.CreateInstance(App.Title & ".CustomerRoleDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerRoleDO.Delete xmlElement
        DeleteCustomerLocks xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objICustomerRoleDO.Delete xmlElement
            DeleteCustomerLocks xmlElement
        Next xmlElement
    End If
    
    Set ICustomerRoleTxBO_Delete = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteExit:

    Set objICustomerRoleDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerRoleTxBO_Delete = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function ICustomerRoleTxBO_Update(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "ICustomerRoleTxBO_Update"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objICustomerRoleDO As ICustomerRoleDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerRoleDO = New CustomerRoleDO
    Else
        Set objICustomerRoleDO = m_objContext.CreateInstance(App.Title & ".CustomerRoleDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerRoleDO.Update xmlElement
    Else
        For Each xmlNode In xmlNodeList
            objICustomerRoleDO.Update xmlNode
        Next xmlNode
    End If
    
    Set ICustomerRoleTxBO_Update = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objICustomerRoleDO = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerRoleTxBO_Update = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Sub CreateLock(ByVal objCustomerRoleElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Creates a CustomerLockApplicationLock for the customer
' pass:         objCustomerRoleElem  Customer role element
'
' return:       n/a
'
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "CreateLock"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objRequestDoc As New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Dim objTypeElem As IXMLDOMElement
    Dim objTableElem As IXMLDOMElement
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objAppManagerBO As ApplicationManagerBO

    If objCustomerRoleElem Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4InvalidParameter, _
                                "Customer Role element is set to nothing"
    End If

    Dim strApplicationNo As String
    Dim strCustomerNo As String

    strApplicationNo = objXmlAssist.GetTagValue(objCustomerRoleElem, "APPLICATIONNUMBER")
    strCustomerNo = objXmlAssist.GetTagValue(objCustomerRoleElem, "CUSTOMERNUMBER")

    If objContext Is Nothing Then
        Set objAppManagerBO = New ApplicationManagerBO
    Else
        Set objAppManagerBO = objContext.CreateInstance(cstrCOMPONENT_NAME & _
                                                             ".ApplicationManagerBO")
    End If

    ' Set up create request for CustomerLockApplicationLock
    Dim strResponse As String

    Set objRequestNode = objXmlAssist.GetRequestNode(objCustomerRoleElem.ownerDocument)
    objRequestDoc.appendChild objRequestNode
    Set objTypeElem = objRequestDoc.createElement("CREATE")
    objRequestNode.appendChild objTypeElem
    Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCKAPPLICATIONLOCK")
    objTypeElem.appendChild objTableElem
    Set objElement = objRequestDoc.createElement("APPLICATIONNUMBER")
    objElement.Text = strApplicationNo
    objTableElem.appendChild objElement
    Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
    objElement.Text = strCustomerNo
    objTableElem.appendChild objElement

    strResponse = objAppManagerBO.CreateCustomerLockApplicationLock(objRequestDoc.xml)
    objErrAssist.CheckXMLResponse strResponse, True

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objRequestDoc = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objAppManagerBO = Nothing

    Exit Sub

CreateLockVbErr:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objRequestDoc = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objAppManagerBO = Nothing

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub DeleteCustomerLocks(ByVal objCustomerRoleElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Deletes the locks held on the customer
' pass:         objCustomerRoleElem  Customer role element
'
' return:       n/a
'
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerLocksVbErr

    Const strFunctionName As String = "DeleteCustomerLocks"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objRequestDoc As New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Dim objTypeElem As IXMLDOMElement
    Dim objTableElem As IXMLDOMElement
    Dim objRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objAppManagerBO As ApplicationManagerBO
    Dim objCustomerBO As CustomerBO

    If objCustomerRoleElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeinvalidparameter, _
                                "Customer Role element is set to nothing"
    End If

    Dim strApplicationNo As String
    Dim strCustomerNo As String

    strApplicationNo = objXmlAssist.GetTagValue(objCustomerRoleElem, "APPLICATIONNUMBER")
    strCustomerNo = objXmlAssist.GetTagValue(objCustomerRoleElem, "CUSTOMERNUMBER")

    If objContext Is Nothing Then
        Set objAppManagerBO = New ApplicationManagerBO
        Set objCustomerBO = New CustomerBO
    Else
        Set objAppManagerBO = objContext.CreateInstance(cstrCOMPONENT_NAME & _
                                                             ".ApplicationManagerBO")
        Set objCustomerBO = objContext.CreateInstance(cstrCUSTOMER_COMPONENET & _
                                                            ".CustomerBO")
    End If

    ' Set up delete request for CustomerLockApplicationLock
    Dim strResponse As String

    Set objRequestNode = objRequestDoc.createElement("REQUEST")
    objRequestDoc.appendChild objRequestNode
    Set objTypeElem = objRequestDoc.createElement("DELETE")
    objRequestNode.appendChild objTypeElem
    Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCKAPPLICATIONLOCK")
    objTypeElem.appendChild objTableElem
    Set objElement = objRequestDoc.createElement("APPLICATIONNUMBER")
    objElement.Text = strApplicationNo
    objTableElem.appendChild objElement
    Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
    objElement.Text = strCustomerNo
    objTableElem.appendChild objElement

    strResponse = objAppManagerBO.DeleteCustomerLockApplicationLock(objRequestNode.xml)
    objErrAssist.CheckXMLResponse strResponse, True

    ' Delete Customer Lock
    Set objRequestNode = objRequestDoc.createElement("REQUEST")
    Set objTableElem = objRequestDoc.createElement("CUSTOMERLOCK")
    objRequestNode.appendChild objTableElem

    Set objElement = objRequestDoc.createElement("CUSTOMERNUMBER")
    objElement.Text = strCustomerNo
    objTableElem.appendChild objElement
    
    strResponse = objCustomerBO.DeleteCustomerLock(objRequestNode.xml)
    objErrAssist.CheckXMLResponse strResponse, True
        
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objRequestDoc = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objAppManagerBO = Nothing
    Set objCustomerBO = Nothing

    Exit Sub

DeleteCustomerLocksVbErr:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objRequestDoc = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objAppManagerBO = Nothing
    Set objCustomerBO = Nothing
    
    '   Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ContactBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'ContactBO' with actual class name
' *****************************************************************************

'Workfile:      ContactBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Contact Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (ContactTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. ContactTxBO, ContactDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'RF     30/09/99    Applied changes raised by code review of 30/09/99, including:
'                   removed AnonInterfaceFunction
'                   removed calls to Validate
'                   improved error handling
'RF     04/10/99    Added profiling
'PSC    21/10/99    Amend warning processing
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'JLD    23/10/01    SYS2839 Error in GetContact
'SG     06/12/01    SYS3361 Instantiated object in IContactBO_GetContact
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit

Implements IContactBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================


'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    01/10/2003  BMIDS612 - Correct SaveContact to return errors correctly
'------------------------------------------------------------------------------------------


Private Function IContactBO_DetermineContactOperation(ByVal vxmlElement As IXMLDOMElement) As Integer

' header ----------------------------------------------------------------------------------
' description:  Saves the group connection records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

    On Error GoTo DetermineOperationVbErr
    
Const strFunctionName As String = "IContactBO_DetermineContactOperation"


Dim objErrAssist As New ErrAssist
Dim objXmlAssist As New XMLAssist
Dim objIContactClassDef As IContactClassDef

Dim botOperationType As BOOPERATIONTYPE

'SG 28/05/02 SYS4767 START
Dim xmlTelNode As IXMLDOMNode
Dim blnFoundNonDelete As Boolean
Dim xmlTelNodeList As IXMLDOMNodeList
Dim botTelephoneOPType As BOOPERATIONTYPE
'SG 28/05/02 SYS4767 END

    Set objIContactClassDef = New ContactClassDef

    botOperationType = booNone
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlElement, objIContactClassDef.LoadContactData)
    
    'SG 28/05/02 SYS4767 START
    'MSMS0044 - Check the operation to perform on all telephone records. If any
    'of them are not DELETE then we must alter the contact's delete operation
    'to an UPDATE.
    If botOperationType = booDelete Then
        'Get a list of all the telephone numbers.
        Set xmlTelNodeList = vxmlElement.selectNodes(".//CONTACTTELEPHONEDETAILS")
        
        For Each xmlTelNode In xmlTelNodeList
            botTelephoneOPType = objXmlAssist.DetermineOperationEx(xmlTelNode, objIContactClassDef.LoadContactTelephoneData())
        
            'If we find a telephone number which isn't going to be deleted then
            'we must stop the contact from being deleted.
            If botTelephoneOPType <> booDelete Then
                blnFoundNonDelete = True
                Exit For
            End If
        Next xmlTelNode
        
        'Alter the contact's operation to an update as we want to preserve the
        'telephone record(s).
        If blnFoundNonDelete Then
            botOperationType = booUpdate
        End If
    End If
    'MSMS0044 - End.
    'SG 28/05/02 SYS4767 END


DetermineOperationExit:
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIContactClassDef = Nothing

    IContactBO_DetermineContactOperation = botOperationType
    Exit Function

DetermineOperationVbErr:
    Set objXmlAssist = Nothing
    Set objIContactClassDef = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    IContactBO_DetermineContactOperation = booNone

    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function IContactBO_SaveContact(ByVal vxmlRequest As IXMLDOMElement, _
                                        Optional ByRef rbotOperationType As Integer = 0) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the group connection records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

    On Error GoTo SaveContactVbErr
    
Const strFunctionName As String = "IContactBO_SaveContact"


Dim objErrAssist As New ErrAssist

Dim botOperationType As BOOPERATIONTYPE

    botOperationType = rbotOperationType
    If rbotOperationType = booNone Then
        botOperationType = IContactBO_DetermineContactOperation(vxmlRequest)
        rbotOperationType = botOperationType
    End If

    '
    ' Execute the appropriate operation
    '
    Select Case botOperationType
        Case booCreate: Set IContactBO_SaveContact = IContactBO_CreateContact(vxmlRequest)
        Case booUpdate: Set IContactBO_SaveContact = IContactBO_UpdateContact(vxmlRequest)
        Case booDelete: Set IContactBO_SaveContact = IContactBO_DeleteContact(vxmlRequest)
        Case Else
            ' N/A
    End Select

SaveContactExit:
    Set objErrAssist = Nothing

    Exit Function

SaveContactVbErr:
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ' PSC 01/10/2003 BMIDS612
    Set IContactBO_SaveContact = objErrAssist.CreateErrorResponseEx
    
    '   go to clean-up section
    Resume SaveContactExit
End Function


Private Function IContactBO_CreateContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateContact a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactVbErr
    
    Const strFunctionName As String = "IContactBO_CreateContact"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIContactTxBO As IContactTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    If objContext Is Nothing Then
        Set objIContactTxBO = New ContactTxBO
    Else
        Set objIContactTxBO = objContext.CreateInstance(App.Title & ".ContactTxBO")
    End If
        
    '   call Business Transaction Object CreateContact function
    Set xmlTempResponseNode = objIContactTxBO.CreateContact(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If

    Set IContactBO_CreateContact = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateContactVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IContactBO_CreateContact = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IContactBO_DeleteContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteContact a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteContactVbErr
    
    Const strFunctionName As String = "IContactBO_DeleteContact"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIContactTxBO As IContactTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "DELETE"
    
    If objContext Is Nothing Then
        Set objIContactTxBO = New ContactTxBO
    Else
        Set objIContactTxBO = objContext.CreateInstance(App.Title & ".ContactTxBO")
    End If
           
    '   call Business Transaction Object DeleteContact function
    Set xmlTempResponseNode = objIContactTxBO.DeleteContact(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IContactBO_DeleteContact = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objErrAssist = Nothing
    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteContactVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IContactBO_DeleteContact = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IContactBO_GetContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetContactVbErr
    
    Const strFunctionName As String = "IContactBO_GetContact"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40   'JLD
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIContactDO As IContactDO
    
    Set xmlOut = New FreeThreadedDOMDocument40     'SG 06/12/01 SYS3361
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
        
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIContactDO = New ContactDO
    Else
        Set objIContactDO = objContext.CreateInstance(App.Title & ".ContactDO")
    End If
                
    If vxmlRequest.nodeName = "CONTACTDETAILS" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CONTACTDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CONTACTDETAILS tag not found"
    End If
        
    Set xmlDataNode = objIContactDO.GetContact(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IContactBO_GetContact = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIContactDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetContactVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIContactDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IContactBO_GetContact = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IContactBO_UpdateContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateContact a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactVbErr
    
    Const strFunctionName As String = "IContactBO_UpdateContact"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIContactTxBO As IContactTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "UPDATE"

    If objContext Is Nothing Then
        Set objIContactTxBO = New ContactTxBO
    Else
        Set objIContactTxBO = objContext.CreateInstance(App.Title & ".ContactTxBO")
    End If
    
    Set xmlTempResponseNode = objIContactTxBO.UpdateContact(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IContactBO_UpdateContact = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objErrAssist = Nothing
    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateContactVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIContactTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
        
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IContactBO_UpdateContact = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


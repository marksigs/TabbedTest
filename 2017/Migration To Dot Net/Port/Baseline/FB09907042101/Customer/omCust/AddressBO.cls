VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AddressBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AddressBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. AddressTxBO, AddressDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'SR     02/11/99    Created
'MH     12/05/00    SYS0618 Added postcode validation to address
'JLD    30/08/00    SYS1496 Added public interface for postcode validation method.
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
Implements IAddressBO

Private m_objErrAssist As New ErrAssist

'------------------------------------------------------------------------------------------
'BMids History:
'Prog   Date        Description
'MDC    30/09/2002  BMIDS00510 - Fix error in GetAddress
'GHun   29/01/2003  BM0242  Amend ValidatePostcodeTag to support partial postcodes
'INR    22/05/2003  BM0506 No longer validate the postcode from create or update methods.
'------------------------------------------------------------------------------------------

Private Function IAddressBO_DetermineAddressOperation(ByVal vxmlElement As IXMLDOMElement) As Integer

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
    
Const strFunctionName As String = "IAddressBO_DetermineAddressOperation"


Dim objXmlAssist As New XMLAssist
Dim objIAddressClassDef As IAddressClassDef
Dim objIAddressDO As IAddressDO

Dim botOperationType As BOOPERATIONTYPE

    Set objIAddressDO = New AddressDO
    Set objIAddressClassDef = New AddressClassDef

    botOperationType = booNone
    botOperationType = objXmlAssist.DetermineOperationEx(vxmlElement, objIAddressClassDef.LoadAddressData)

    If (botOperationType = booUpdate) And objIAddressDO.IsEmptyAddress(vxmlElement) Then
        ' With no relevant information left for the address, it needs deleting
        botOperationType = booDelete
    End If

DetermineOperationExit:
    Set objXmlAssist = Nothing
    Set objIAddressDO = Nothing
    
    IAddressBO_DetermineAddressOperation = botOperationType
    Exit Function

DetermineOperationVbErr:
    Set objXmlAssist = Nothing
    Set objIAddressDO = Nothing

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    IAddressBO_DetermineAddressOperation = booNone

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IAddressBO_IsEmptyAddress(ByVal vstrXMLRequest As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Determines whether the address is regarded as 'empty'
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
'
'               Note: This method should NOT BE CALLED INTERNALLY (because it creates a DOMDocument)
'               Instead, call the AddressDO.IsEmptyAddress.

' return:       Boolean
'------------------------------------------------------------------------------------------
On Error GoTo IsEmptyAddressVbErr

    Const strFunctionName As String = "IAddressBO_IsEmptyAddress"


    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlAddressNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objIAddressDO As IAddressDO

    Set objIAddressDO = New AddressDO

    '
    ' Note: This method should NOT BE CALLED INTERNALLY (because it creates a DOMDocument)
    ' Instead, call the private version.
    '
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Find table tag in the request XML
    Set xmlAddressNode = xmlIn.documentElement.selectSingleNode("ADDRESS")
    If xmlAddressNode Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ADDRESS tag is missing"
    End If

    IAddressBO_IsEmptyAddress = objIAddressDO.IsEmptyAddress(xmlAddressNode)

    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlAddressNode = Nothing
    Set objIAddressDO = Nothing

Exit Function

IsEmptyAddressVbErr:
    '   clean-up code
    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlAddressNode = Nothing
    Set objIAddressDO = Nothing

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IAddressBO_SaveAddress(ByVal vxmlRequest As IXMLDOMElement, _
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

    On Error GoTo SaveAddressVbErr
    
Const strFunctionName As String = "IAddressBO_SaveAddress"


Dim botOperationType As BOOPERATIONTYPE

    botOperationType = rbotOperationType
    If rbotOperationType = booNone Then
        botOperationType = IAddressBO_DetermineAddressOperation(vxmlRequest)
        rbotOperationType = botOperationType
    End If

    '
    ' Execute the appropriate operation
    '
    Select Case botOperationType
        Case booCreate: Set IAddressBO_SaveAddress = IAddressBO_CreateAddress(vxmlRequest)
        Case booUpdate: Set IAddressBO_SaveAddress = IAddressBO_UpdateAddress(vxmlRequest)
        Case booDelete: Set IAddressBO_SaveAddress = IAddressBO_DeleteAddress(vxmlRequest)
        Case Else
            ' N/A
    End Select

SaveAddressExit:
    Exit Function

SaveAddressVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IAddressBO_SaveAddress = m_objErrAssist.CreateErrorResponseEx
    
    '   go to clean-up section
    Resume SaveAddressExit

End Function

Private Function IAddressBO_CreateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateAddress a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateAddressVbErr
    
    Const strFunctionName As String = "IAddressBO_CreateAddress"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAddressTxBO As IAddressTxBO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    If objContext Is Nothing Then
        Set objIAddressTxBO = New AddressTxBO
    Else
        Set objIAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IAddressBO_ValidateAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    '   call Business Transaction Object CreateAddress function
    Set xmlTempResponseNode = objIAddressTxBO.CreateAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    Set IAddressBO_CreateAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IAddressBO_CreateAddress = m_objErrAssist.CreateErrorResponseEx()
    

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IAddressBO_DeleteAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteAddress a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAddressVbErr
    
    Const strFunctionName As String = "IAddressBO_DeleteAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAddressTxBO As IAddressTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "DELETE"
    
    If objContext Is Nothing Then
        Set objIAddressTxBO = New AddressTxBO
    Else
        Set objIAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If
           
    '   call Business Transaction Object DeleteAddress function
    Set xmlTempResponseNode = objIAddressTxBO.DeleteAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IAddressBO_DeleteAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IAddressBO_DeleteAddress = m_objErrAssist.CreateErrorResponseEx()
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IAddressBO_GetAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAddressVbErr
    
    Const strFunctionName As String = "IAddressBO_GetAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIAddressDO As IAddressDO
    
    'BMIDS00510 MDC 30/09/2002
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    'BMIDS00510 MDC 30/09/2002 - End
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIAddressDO = New AddressDO
    Else
        Set objIAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
                
    If vxmlRequest.nodeName = "ADDRESS" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("ADDRESS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ADDRESS tag not found"
    End If
        
    Set xmlDataNode = objIAddressDO.GetAddress(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IAddressBO_GetAddress = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
   Set objIAddressDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIAddressDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IAddressBO_GetAddress = m_objErrAssist.CreateErrorResponseEx()
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IAddressBO_UpdateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateAddress a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAddressVbErr
    
    Const strFunctionName As String = "IAddressBO_UpdateAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAddressTxBO As IAddressTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "UPDATE"

    If objContext Is Nothing Then
        Set objIAddressTxBO = New AddressTxBO
    Else
        Set objIAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If
    
    ' Run validation checks
    Set xmlTempResponseNode = IAddressBO_ValidateAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlTempResponseNode = objIAddressTxBO.UpdateAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IAddressBO_UpdateAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIAddressTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
        
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IAddressBO_UpdateAddress = m_objErrAssist.CreateErrorResponseEx()
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IAddressBO_ValidateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateAddressVbErr

    Const strFunctionName As String = "IAddressBO_ValidateAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'BM0506 No longer validate the postcode.
'    ValidatePostcodeTag vxmlRequest

    Set IAddressBO_ValidateAddress = xmlResponseElem.cloneNode(True)

    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IAddressBO_ValidateAddress = m_objErrAssist.CreateErrorResponseEx()
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function ValidatePostCode(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Validates the postcode past in
' pass:         vstrXmlRequest  xml Request data stream containing postcode to check
'               <VALIDATEPOSTCODE>
'                   <POSTCODE>GL50 7ty</POSTCODE>
'               </VALIDATEPOSTCODE>
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidatePostCode"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    ValidatePostcodeTag xmlIn.documentElement
    'objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ValidatePostCode = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ValidateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidatePostCode = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Private Sub ValidatePostcodeTag(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Postcode validation identical to that in scScreenFunctions
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'------------------------------------------------------------------------------------------
On Error GoTo ValidatePostcodeTagVbErr

    Const strFunctionName As String = "ValidatePostcodeTag"


    Dim xmlPostCode As IXMLDOMElement
    Dim strPostcode As String
    Dim blnSuccess As Boolean
    
    'Ignore optional or absent postcodes
    Set xmlPostCode = vxmlRequest.getElementsByTagName("POSTCODE").Item(0)
    
    If xmlPostCode Is Nothing Then Exit Sub

    strPostcode = Trim$(xmlPostCode.Text)
    'BM0242 Remove duplicate spaces
    Do While InStr(1, strPostcode, "  ")
        strPostcode = Replace(strPostcode, "  ", " ")
    Loop
    'BM0242 End
    xmlPostCode.Text = strPostcode
    
    If strPostcode = vbNullString Then Exit Sub
    
    blnSuccess = False
    
    Select Case Len(strPostcode)
    'BM0242
    Case 1
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@")
    Case 2
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@#")
    'BM0242 End
    Case 3
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@#") Or _
            ValidatePostcodeWithMask(strPostcode, "@#@") Or _
            ValidatePostcodeWithMask(strPostcode, "@##") Or _
            ValidatePostcodeWithMask(strPostcode, "IOM")
    Case 4
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@#@") Or _
            ValidatePostcodeWithMask(strPostcode, "@# #") Or _
            ValidatePostcodeWithMask(strPostcode, "BT74") Or _
            ValidatePostcodeWithMask(strPostcode, "@@##") 'BM0242
    Case 5
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@# #") Or _
            ValidatePostcodeWithMask(strPostcode, "@#@ #") Or _
            ValidatePostcodeWithMask(strPostcode, "@## #") Or _
            ValidatePostcodeWithMask(strPostcode, "@# #@") 'BM0242
    Case 6
        'BM0242
        'blnSuccess = blnSuccess Or _
        '    ValidatePostcodeWithMask(strPostcode, "@@#@ #") Or _
        '    ValidatePostcodeWithMask(strPostcode, "@@## #") Or _
        '    ValidatePostcodeWithMask(strPostcode, "@# #@@")
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@#@ #") Or _
            ValidatePostcodeWithMask(strPostcode, "@@## #") Or _
            ValidatePostcodeWithMask(strPostcode, "@# #@@") Or _
            ValidatePostcodeWithMask(strPostcode, "@## #@") Or _
            ValidatePostcodeWithMask(strPostcode, "@@# #@") Or _
            ValidatePostcodeWithMask(strPostcode, "@#@ #@")
        'BM0242 End
    Case 7
        'BM0242
        'blnSuccess = blnSuccess Or _
        '    ValidatePostcodeWithMask(strPostcode, "@@# #@@") Or _
        '    ValidatePostcodeWithMask(strPostcode, "@#@ #@@") Or _
        '    ValidatePostcodeWithMask(strPostcode, "@## #@@")
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@# #@@") Or _
            ValidatePostcodeWithMask(strPostcode, "@#@ #@@") Or _
            ValidatePostcodeWithMask(strPostcode, "@## #@@") Or _
            ValidatePostcodeWithMask(strPostcode, "@@## #@") Or _
            ValidatePostcodeWithMask(strPostcode, "@@#@ #@")
        'BM0242 End
    Case 8
        blnSuccess = blnSuccess Or _
            ValidatePostcodeWithMask(strPostcode, "@@#@ #@@") Or _
            ValidatePostcodeWithMask(strPostcode, "@@## #@@")
    Case Else
        blnSuccess = False
    End Select

    If Not blnSuccess Then _
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 543
    Exit Sub

ValidatePostcodeTagVbErr:
    
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ValidatePostcodeWithMask(ByVal vstrPostcode As String, ByVal vstrMask As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Compares a postcode with a mask. If it matches then all is well
'   This is the same code as in scScreenfunctions.
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'------------------------------------------------------------------------------------------
    Dim blnSuccess As Boolean
    Dim intLen As Integer
    Dim intPos As Integer
    Dim strChar As String
    
    intLen = Len(vstrMask)
    blnSuccess = True
'    Debug.Assert vstrPostcode <> "IOM"
    
    For intPos = 1 To intLen
        strChar = UCase$(Mid$(vstrPostcode, intPos, 1))
        Select Case Mid$(vstrMask, intPos, 1)
        Case "@":
            blnSuccess = blnSuccess And Not (strChar < "A" Or strChar > "Z")
        Case "#"
            blnSuccess = blnSuccess And Not (strChar < "0" Or strChar > "9")
        Case Else
            blnSuccess = blnSuccess And (Mid$(vstrMask, intPos, 1) = strChar)
        End Select
    Next intPos
    
    ValidatePostcodeWithMask = blnSuccess
End Function



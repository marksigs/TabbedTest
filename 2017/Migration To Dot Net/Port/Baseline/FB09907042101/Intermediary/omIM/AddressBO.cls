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
'------------------------------------------------------------------------------------------

Option Explicit
 
Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object
'               invokes AddressTxBO business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "Create"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXMLAssist As New XMLAssist
        
    Dim objAddressTxBO As AddressTxBO
    Dim strXML As String

    If objContext Is Nothing Then
        Set objAddressTxBO = New AddressTxBO
    Else
        Set objAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   call Business Transaction Object Create function
    Validate vstrXMLRequest
    strXML = objAddressTxBO.Create(vstrXMLRequest)
    xmlOut.firstChild.appendChild objXMLAssist.load(strXML, TypeName(Me), strFunctionName).documentElement

    Create = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    Set objAddressTxBO = Nothing

    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Create = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function Update(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes AddresstxBO business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXMLAssist As New XMLAssist

    Dim objAddressTxBO As AddressTxBO
    
    If objContext Is Nothing Then
        Set objAddressTxBO = New AddressTxBO
    Else
        Set objAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Validate vstrXMLRequest
    objAddressTxBO.Update vstrXMLRequest
        
    Update = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    Set objAddressTxBO = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Update = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume UpdateExit
    
End Function


Public Function Delete(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes AddressTxBO  object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "Delete"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objAddressTxBO As AddressTxBO

    If objContext Is Nothing Then
        Set objAddressTxBO = New AddressTxBO
    Else
        Set objAddressTxBO = objContext.CreateInstance(App.Title & ".AddressTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
           
    '   call Business Transaction Object Delete function
   objAddressTxBO.Delete (vstrXMLRequest)
    
    Delete = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objAddressTxBO = Nothing
    Set objXMLAssist = Nothing

    Exit Function

DeleteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Delete = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume DeleteExit

End Function

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "GetData"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
' TODO ************************************************************************
'   Remove oeNotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented

' TODO ************************************************************************
'   replace Omiga4DO reference with actual DO required
' *****************************************************************************
    Dim objAddressDO As AddressDO
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objAddressDO = New AddressDO
    Else
        Set objAddressDO = objContext.CreateInstance(App.Title & ".Omiga4DO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName("ADDRESS").Item(0)
    
    '   call Data Object GetData function
    Dim strXMLData As String
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ADDRESS" & " tag not found"
    End If
        
    strXMLData = objAddressDO.GetData(xmlRequestNode.xml)
    Set xmlData = objXMLAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
    
    GetData = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDataExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXMLAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objAddressDO = Nothing

    Exit Function

GetDataVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume GetDataExit

End Function

Private Sub Validate(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "Validate"

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' call scripting interface
    '-----------------------------------------------------------------------------

    Dim objScriptInterface As New ScriptInterface
    objScriptInterface.RunScript App.Title, _
                                 TypeName(Me), _
                                 strFunctionName, _
                                 vstrXMLRequest
    
    '   clean-up code
    Set objScriptInterface = Nothing
    Set xmlIn = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing

Exit Sub

ValidateVbErr:

    '   clean-up code
    Set objScriptInterface = Nothing
    Set xmlIn = Nothing
    Set objXMLAssist = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function IsEmptyAddress(ByVal vstrXMLRequest As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Determines whether the address is regarded as 'empty'
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
' return:       Boolean
'------------------------------------------------------------------------------------------
On Error GoTo IsEmptyAddressVbErr

    Const strFunctionName As String = "IsEmptyAddress"

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlRootElement As IXMLDOMElement

    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Find table tag in the request XML
    Set xmlNodeList = xmlIn.getElementsByTagName("ADDRESS")
    If xmlNodeList.length > 0 Then
        Set xmlRootElement = xmlNodeList.Item(0)
    End If


    If Not (xmlRootElement Is Nothing) Then
        IsEmptyAddress = (objXMLAssist.GetTagValue(xmlRootElement, "POSTCODE") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "FLATNUMBER") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "BUILDINGORHOUSENAME") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "BUILDINGORHOUSENUMBER") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "STREET") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "DISTRICT") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "TOWN") = "") And _
                         (objXMLAssist.GetTagValue(xmlRootElement, "COUNTY") = "")
    Else
        IsEmptyAddress = True
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' call scripting interface
    '-----------------------------------------------------------------------------

    Set xmlIn = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlRootElement = Nothing

Exit Function

IsEmptyAddressVbErr:

    '   clean-up code
    Set xmlIn = Nothing
    Set objXMLAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlRootElement = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SaveAddress(ByVal vstrXMLRequest As String, _
                            Optional ByRef rbotOperationType As Variant) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the address records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

    On Error GoTo SaveAddressVbErr
    
Const strFunctionName As String = "SaveAddress"

Dim objErrAssist As New ErrAssist

Dim botOperationType As BOOPERATIONTYPE

    botOperationType = booNone
    If Not IsMissing(rbotOperationType) Then botOperationType = rbotOperationType

    If botOperationType = booNone Then
        botOperationType = DetermineOperation(vstrXMLRequest)
    End If


    Select Case botOperationType
        Case booCreate: SaveAddress = Create(vstrXMLRequest)
        Case booUpdate: SaveAddress = Update(vstrXMLRequest)
        Case booDelete: SaveAddress = Delete(vstrXMLRequest)
        Case Else
            ' N/A
    End Select


    If Not IsMissing(rbotOperationType) Then
        rbotOperationType = botOperationType
    End If

SaveAddressExit:
    Set objErrAssist = Nothing

    Exit Function

SaveAddressVbErr:
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveAddress = objErrAssist.CreateErrorResponse
    
    '   go to clean-up section
    Resume SaveAddressExit

End Function

Public Function DetermineOperation(ByVal vstrXMLRequest As String) As Integer
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
    
Const strFunctionName As String = "DetermineOperation"

Dim objErrAssist As New ErrAssist
Dim objXMLAssist As New XMLAssist
Dim objAddressDO As New AddressDO

Dim botOperationType As BOOPERATIONTYPE

    botOperationType = booNone

    botOperationType = objXMLAssist.DetermineOperation(vstrXMLRequest, objAddressDO.LoadAddressData)

    If IsEmptyAddress(vstrXMLRequest) Then
        If botOperationType = booCreate Then
            ' No point in creating an address for which no relevant information is specified
            botOperationType = booNone
        ElseIf botOperationType = booUpdate Then
            ' With no relevant information left for the address, it needs deleting
            botOperationType = booDelete
        End If
    End If

DetermineOperationExit:
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set objAddressDO = Nothing

    DetermineOperation = botOperationType
    Exit Function

DetermineOperationVbErr:
    Set objXMLAssist = Nothing
    Set objAddressDO = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DetermineOperation = objErrAssist.CreateErrorResponse
    
    '   go to clean-up section
    Resume DetermineOperationExit

End Function



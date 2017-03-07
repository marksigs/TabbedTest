VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AttitudeToBorrowingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AttitudeToBorrowingBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (AttitudeToBorrowingTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  AttitudeToBorrowingTxBO,
'               AttitudeToBorrowingDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'RF    01/08/99 Created
'RF    27/01/00 Pick up script interface from base component.
'DRC   05/10/01 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "ATTITUDETOBORROWING"


Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   create a single instance of the persistant data associated with this
'   business object
'   invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       Create          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "Create"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlNode As IXMLDOMNode

    Dim objAttitudeToBorrowingTxBO As AttitudeToBorrowingTxBO

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingTxBO = New AttitudeToBorrowingTxBO
    Else
        Set objAttitudeToBorrowingTxBO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    Validate vstrXMLRequest
    
    ' call Business Transaction Object Create function
    objAttitudeToBorrowingTxBO.Create (vstrXMLRequest)
    
    ' if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Create = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objAttitudeToBorrowingTxBO = Nothing

    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Create = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume CreateExit

End Function

Public Function Update(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       Update          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "Update"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
' TODO ************************************************************************
'   replace AttitudeToBorrowingTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objAttitudeToBorrowingTxBO As AttitudeToBorrowingTxBO
    
    If objContext Is Nothing Then
        Set objAttitudeToBorrowingTxBO = New AttitudeToBorrowingTxBO
    Else
        Set objAttitudeToBorrowingTxBO = objContext.CreateInstance(App.Title & ".AttitudeToBorrowingTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    
    If Not objXmlResponseElem Is Nothing Then
        Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    End If
    
    ' Call scripting interface
    Validate (vstrXMLRequest)
    
    objAttitudeToBorrowingTxBO.Update vstrXMLRequest
    
    '   if we get here, everything has completed OK
    If Not objXmlResponseElem Is Nothing Then
        objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    End If
    
    Update = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objAttitudeToBorrowingTxBO = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As FreeThreadedDOMDocument40
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

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
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       Delete          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "Delete"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
    
' TODO ************************************************************************
'   replace AttitudeToBorrowingTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objAttitudeToBorrowingTxBO As AttitudeToBorrowingTxBO

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingTxBO = New AttitudeToBorrowingTxBO
    Else
        Set objAttitudeToBorrowingTxBO = objContext.CreateInstance(App.Title & ".AttitudeToBorrowingTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
           
    '   call Business Transaction Object Delete function
   objAttitudeToBorrowingTxBO.Delete (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Delete = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objAttitudeToBorrowingTxBO = Nothing

    Exit Function

DeleteVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

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
'
' return:       GetData         xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "GetData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    

' TODO ************************************************************************
'   replace AttitudeToBorrowingDO reference with actual DO required
' *****************************************************************************
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
                
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    '   call Data Object GetData function
    Dim strXMLData As String
    
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objAttitudeToBorrowingDO.GetData(objXmlRequestNode.xml)
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If

    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    GetData = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDataExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objAttitudeToBorrowingDO = Nothing

    Exit Function

GetDataVbErr:
    
    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
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
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "Validate"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlDerived As New FreeThreadedDOMDocument40
    objXmlDerived.validateOnParse = False
    objXmlDerived.setProperty "NewParser", True
    Dim objXmlBeforeData As IXMLDOMElement
    Dim objXmlAfterData As IXMLDOMElement
    Dim objXmlAfterNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist

' TODO ************************************************************************
'   replace AttitudeToBorrowingDO reference with actual DO required
' *****************************************************************************
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
    
' TODO ************************************************************************
'   add processing here
'   e.g.
'   examine input xml request & extract before image node (if present)
'   & after image node
' *****************************************************************************
    
    If objXmlIn.getElementsByTagName("UPDATE").length > 0 Then
        Dim int0 As Integer
        For int0 = 0 To objXmlIn.getElementsByTagName("UPDATE").length - 1
            Select Case objXmlIn.getElementsByTagName("UPDATE").Item(int0).Attributes.getNamedItem("TYPE").Text
                Case "BEFORE"
                    Set objXmlBeforeData = objXmlIn.getElementsByTagName("UPDATE").Item(int0).firstChild
                Case "AFTER"
                    Set objXmlAfterNode = objXmlIn.getElementsByTagName("UPDATE").Item(int0)
                    Set objXmlAfterData = objXmlAfterNode.firstChild
            End Select
        Next
    Else
        If objXmlIn.getElementsByTagName("CREATE").length > 0 Then
            Set objXmlAfterNode = objXmlIn.getElementsByTagName("CREATE").Item(int0)
            Set objXmlAfterData = objXmlAfterNode.firstChild
        End If
    End If
    
'   replace 'after' image with xml plus any derived values

    If Not objXmlAfterData Is Nothing Then
        Dim strXMLAfterData As String
        strXMLAfterData = objAttitudeToBorrowingDO.AddDerivedData(objXmlAfterData.xml)
        Set objXmlDerived = objXmlAssist.load(strXMLAfterData, TypeName(Me), strFunctionName)
        objXmlAfterNode.replaceChild objXmlDerived.documentElement, objXmlAfterData
    End If
    
' TODO ************************************************************************
'   add processing here
'   e.g.
'   raise error if Surname changed
' *****************************************************************************

    If (Not objXmlBeforeData Is Nothing) And (Not objXmlAfterData Is Nothing) Then
        Dim strWasSurname As String, strIsSurname As String
        If objXmlBeforeData.getElementsByTagName("SURNAME").length > 0 Then
            strWasSurname = objXmlBeforeData.getElementsByTagName("SURNAME").Item(0).Text
        End If
        If objXmlAfterData.getElementsByTagName("SURNAME").length > 0 Then
            strIsSurname = objXmlAfterData.getElementsByTagName("SURNAME").Item(0).Text
        End If
        If strWasSurname <> strIsSurname Then
            Dim objErrAssist As New ErrAssist
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4SurnameChange
        End If
    End If
    
'   -----------------------------------------------------------------------------
'   hard-coded validation OK
'   call scripting interface
'   -----------------------------------------------------------------------------

'    Dim objScriptInterface As New omBase.ScriptInterface
    Dim objScriptInterface As ScriptInterface
    
    Set objScriptInterface = New ScriptInterface
    
    objScriptInterface.RunScript App.Title, _
                                 TypeName(Me), _
                                 strFunctionName, _
                                 vstrXMLRequest
    
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objAttitudeToBorrowingDO = Nothing

Exit Sub

ValidateVbErr:

    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objAttitudeToBorrowingDO = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
    
' TODO ************************************************************************
'   replace AnonInterfaceFunction with actual function name
' *****************************************************************************

Public Function AnonInterfaceFunction(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  add description here
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       AnonInterfaceFunction  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AnonInterfaceFunctionVbErr
    
    Const strFunctionName As String = "AnonInterfaceFunction"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    
' TODO ************************************************************************
'   add processing here
'   replace 'not implemented' error
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    AnonInterfaceFunction = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

AnonInterfaceFunctionExit:
    
    Set objErrAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing

    Exit Function

AnonInterfaceFunctionVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    AnonInterfaceFunction = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume AnonInterfaceFunctionExit

End Function

Public Function FindList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       FindList         xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "FindList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    Dim strXMLData As String
    
    '   call Data Object FindList function
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objAttitudeToBorrowingDO.FindList(objXmlRequestNode.xml)
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    FindList = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objAttitudeToBorrowingDO = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindListExit

End Function

Public Function DeleteAll(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       DeleteAll       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "DeleteAll"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
' TODO ************************************************************************
'   replace AttitudeToBorrowingTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objAttitudeToBorrowingTxBO As AttitudeToBorrowingTxBO

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingTxBO = New AttitudeToBorrowingTxBO
    Else
        Set objAttitudeToBorrowingTxBO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   call Business Transaction Object DeleteAll function
    objAttitudeToBorrowingTxBO.DeleteAll (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    DeleteAll = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteAllExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objAttitudeToBorrowingTxBO = Nothing

    Exit Function

DeleteAllVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAll = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume DeleteAllExit

End Function




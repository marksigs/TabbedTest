VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuickQuoteOutgoingsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuoteOutgoingsBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (QuickQuoteOutgoingsTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:  QuickQuoteOutgoingsTxBO,
'               QuickQuoteOutgoingsDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'RF    01/08/99 Created
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "QUICKQUOTEOUTGOINGS"

Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
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

   Dim objQuickQuoteOutgoingsTxBO As QuickQuoteOutgoingsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsTxBO = New QuickQuoteOutgoingsTxBO
    Else
        Set objQuickQuoteOutgoingsTxBO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    Validate (vstrXMLRequest)
    
    '   call Business Transaction Object Create function
    objQuickQuoteOutgoingsTxBO.Create (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
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
    Set objQuickQuoteOutgoingsTxBO = Nothing

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
'   replace QuickQuoteOutgoingsTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objQuickQuoteOutgoingsTxBO As QuickQuoteOutgoingsTxBO
    
    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsTxBO = New QuickQuoteOutgoingsTxBO
    Else
        Set objQuickQuoteOutgoingsTxBO = objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    
    If Not objXmlResponseElem Is Nothing Then
        Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    End If
    
    ' Call scripting interface
    Validate (vstrXMLRequest)
    
    objQuickQuoteOutgoingsTxBO.Update vstrXMLRequest
    
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
    Set objQuickQuoteOutgoingsTxBO = Nothing
    
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
'   replace QuickQuoteOutgoingsTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objQuickQuoteOutgoingsTxBO As QuickQuoteOutgoingsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsTxBO = New QuickQuoteOutgoingsTxBO
    Else
        Set objQuickQuoteOutgoingsTxBO = objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
           
    '   call Business Transaction Object Delete function
   objQuickQuoteOutgoingsTxBO.Delete (vstrXMLRequest)
    
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
    Set objQuickQuoteOutgoingsTxBO = Nothing

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
'   replace QuickQuoteOutgoingsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteOutgoingsDO As QuickQuoteOutgoingsDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    

    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsDO = New QuickQuoteOutgoingsDO
    Else
        Set objQuickQuoteOutgoingsDO = objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsDO")
    End If
                
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    '   call Data Object GetData function
    Dim strXMLData As String
    
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objQuickQuoteOutgoingsDO.GetData(objXmlRequestNode.xml)
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
    Set objQuickQuoteOutgoingsDO = Nothing

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
    Dim objXmlAssist As New XMLAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objErrAssist As New ErrAssist

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' if AMOUNT is not null, FREQUENCYSEQUENCENUMBER may not be null;
    ' if FREQUENCYSEQUENCENUMBER is not null, AMOUNT may not be null;
    
    Set xmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    Dim intNumNodes As Integer
    Dim intNode As Integer
    Dim strFREQUENCYSEQUENCENUMBER, strAMOUNT As String
    
    intNumNodes = xmlNodeList.length
    
    ' fime - check this does both applicants
    For intNode = 0 To intNumNodes - 1
        strFREQUENCYSEQUENCENUMBER = objXmlAssist.GetTagValue(xmlNodeList.Item(intNode), "FREQUENCYSEQUENCENUMBER")
        strAMOUNT = objXmlAssist.GetTagValue(xmlNodeList.Item(intNode), "AMOUNT")
        
        If (Len(strFREQUENCYSEQUENCENUMBER) > 0 And Len(strAMOUNT) = 0) Or _
            (Len(strFREQUENCYSEQUENCENUMBER) = 0 And Len(strAMOUNT) > 0) Then
            ' "Quick Quote Outgoings - mandatory field missing"
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err124
        End If
    Next
    
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' call scripting interface
    '-----------------------------------------------------------------------------

    'Dim objScriptInterface As New omBase.ScriptInterface
    Dim objScriptInterface As ScriptInterface
    
    Set objScriptInterface = New ScriptInterface
    
    objScriptInterface.RunScript App.Title, TypeName(Me), _
                                 strFunctionName, vstrXMLRequest
    
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

Exit Sub

ValidateVbErr:

    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
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

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
' TODO ************************************************************************
'   replace QuickQuoteOutgoingsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteOutgoingsDO As QuickQuoteOutgoingsDO

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsDO = New QuickQuoteOutgoingsDO
    Else
        Set objQuickQuoteOutgoingsDO = objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    Dim strXMLData As String
    
    '   call Data Object FindList function
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objQuickQuoteOutgoingsDO.FindList(objXmlRequestNode.xml)
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
    Set objQuickQuoteOutgoingsDO = Nothing

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
    
    Dim objQuickQuoteOutgoingsTxBO As QuickQuoteOutgoingsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsTxBO = New QuickQuoteOutgoingsTxBO
    Else
        Set objQuickQuoteOutgoingsTxBO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   call Business Transaction Object DeleteAll function
    objQuickQuoteOutgoingsTxBO.DeleteAll vstrXMLRequest
    
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
    Set objQuickQuoteOutgoingsTxBO = Nothing

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
        ' fixme - is this where we catch DeleteAll no records affected error?
        'objContext.SetAbort
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume DeleteAllExit

End Function




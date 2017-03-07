VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "GlobalParameterBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      GlobalParameterBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Global Parameter Business Object which 'supports transactions' only
'               Methods which do not require transaction support reside in this
'               class. Any methods that require transactions will be delegated to
'               GlobalParameterTxBO
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:  GlobalParameterTxBO, GlobalParameterDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'PSC    21/07/99 Created
'MV     29/11/00 Created New Function FindCurrentParameterList
'ASm    12/01/01 SYS1817: oeNoDataForParameter enum removed and replaced with specific error number
'PSC    12/03/01 SYS2024: Added IsTaskManager()
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    10/04/2003  BM0493 Added GetCurrentParameterListEx
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
'Core History
'
'Prog   Date        Description
'AS     22/03/2006  CORE175: BMIDS609 GlobalParameterBO Functions not getting the object context
'------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl
 
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
'Private m_objIDOAssist As IDOAssist ' AS 22/03/2006 CORE175 Never used
 
 '=============================================
 'Constant Declaration Section
 '=============================================
' application name
Private Const strAppName As String = "Omiga4"
' registry section for database connection info
Private Const strREGISTRY_SECTION As String = "System Configuration"
 '=============================================
 'Variable Declaration Section
 '=============================================

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
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
    Dim objGlobalParameterTxBO As GlobalParameterTxBO
    If m_objContext Is Nothing Then
        Set objGlobalParameterTxBO = New GlobalParameterTxBO
    Else
        Set objGlobalParameterTxBO = m_objContext.CreateInstance(App.Title & ".GlobalParameterTxBO")
    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
' TODO ************************************************************************
'   add processing here
' *****************************************************************************
    
'   ---------------------------------------------------------------------------
'   perform hard-coded validation & call scripting interface
'   ---------------------------------------------------------------------------
    Validate (vstrXMLRequest)
    '   call Business Transaction Object Create function
    objGlobalParameterTxBO.Create (vstrXMLRequest)
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Create = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
CreateExit:
    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objGlobalParameterTxBO = Nothing
    Exit Function
CreateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    Create = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
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
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
    Dim objGlobalParameterTxBO As GlobalParameterTxBO
    If m_objContext Is Nothing Then
        Set objGlobalParameterTxBO = New GlobalParameterTxBO
    Else
        Set objGlobalParameterTxBO = m_objContext.CreateInstance(App.Title & ".GlobalParameterTxBO")
    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    If Not objXmlResponseElem Is Nothing Then
        Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    End If
    ' Call scripting interface
    Validate (vstrXMLRequest)
    objGlobalParameterTxBO.Update vstrXMLRequest
    '   if we get here, everything has completed OK
    If Not objXmlResponseElem Is Nothing Then
        objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    End If
    Update = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
UpdateExit:
    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objGlobalParameterTxBO = Nothing
    Exit Function
UpdateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    Update = m_objErrAssist.CreateErrorResponse
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
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    Dim objGlobalParameterTxBO As GlobalParameterTxBO
    If m_objContext Is Nothing Then
        Set objGlobalParameterTxBO = New GlobalParameterTxBO
    Else
        Set objGlobalParameterTxBO = m_objContext.CreateInstance(App.Title & ".GlobalParameterTxBO")
    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
           
    '   call Business Transaction Object Delete function
   objGlobalParameterTxBO.Delete (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Delete = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
DeleteExit:
    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objGlobalParameterTxBO = Nothing
    Exit Function
DeleteVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    Delete = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
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
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objGlobalParameterDO As GlobalParameterDO
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName("GLOBALPARAMETER").Item(0)
    '   call Data Object GetData function
    Dim strXMLData As String
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objGlobalParameterDO.GetData(objXmlRequestNode.xml)
        Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "GLOBALPARAMETER tag not found"
    End If
        
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    GetData = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
GetDataExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlRequestNode = Nothing
    Set objGlobalParameterDO = Nothing
    Exit Function

GetDataVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    GetData = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
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
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlDerived As New FreeThreadedDOMDocument40
    Dim objXmlBeforeData As IXMLDOMElement
    Dim objXmlAfterData As IXMLDOMElement
    Dim objXmlAfterNode As IXMLDOMNode
    Dim objGlobalParameterDO As GlobalParameterDO
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
' TODO ************************************************************************
'   add processing here
'   e.g.
'   examine input xml request & extract before image node (if present)
'   & after image node
' *****************************************************************************
    
    If objXmlIn.getElementsByTagName("UPDATE").Length > 0 Then
        Dim int0 As Integer
        For int0 = 0 To objXmlIn.getElementsByTagName("UPDATE").Length - 1
            Select Case objXmlIn.getElementsByTagName("UPDATE").Item(int0).Attributes.getNamedItem("TYPE").Text
                Case "BEFORE"
                    Set objXmlBeforeData = objXmlIn.getElementsByTagName("UPDATE").Item(int0).firstChild
                Case "AFTER"
                    Set objXmlAfterNode = objXmlIn.getElementsByTagName("UPDATE").Item(int0)
                    Set objXmlAfterData = objXmlAfterNode.firstChild
            End Select
        Next
    Else
        If objXmlIn.getElementsByTagName("CREATE").Length > 0 Then
            Set objXmlAfterNode = objXmlIn.getElementsByTagName("CREATE").Item(int0)
            Set objXmlAfterData = objXmlAfterNode.firstChild
        End If
    End If
'   replace 'after' image with xml plus any derived values
    If Not objXmlAfterData Is Nothing Then
        Dim strXMLAfterData As String
        strXMLAfterData = objGlobalParameterDO.AddDerivedData(objXmlAfterData.xml)
        Set objXmlDerived = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        Set objXmlAfterNode = objXmlAfterNode.replaceChild(objXmlDerived.documentElement, objXmlAfterData)
    End If
' TODO ************************************************************************
'   add processing here
'   e.g.
'   raise error if Surname changed
' *****************************************************************************
    If (Not objXmlBeforeData Is Nothing) And (Not objXmlAfterData Is Nothing) Then
        Dim strWasSurname As String, strIsSurname As String
        If objXmlBeforeData.getElementsByTagName("SURNAME").Length > 0 Then
            strWasSurname = objXmlBeforeData.getElementsByTagName("SURNAME").Item(0).Text
        End If
        If objXmlAfterData.getElementsByTagName("SURNAME").Length > 0 Then
            strIsSurname = objXmlAfterData.getElementsByTagName("SURNAME").Item(0).Text
        End If
        If strWasSurname <> strIsSurname Then
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4SurnameChange
        End If
    End If
'   -----------------------------------------------------------------------------
'   hard-coded validation OK
'   call scripting interface
'   -----------------------------------------------------------------------------
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
    Set objGlobalParameterDO = Nothing
Exit Sub
ValidateVbErr:
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
    Set objGlobalParameterDO = Nothing
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
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
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    Dim objGlobalParameterDO As GlobalParameterDO
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName("GLOBALPARAMETER").Item(0)
    Dim strXMLData As String
    '   call Data Object FindList function
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objGlobalParameterDO.FindList(objXmlRequestNode.xml)
        Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "GLOBALPARAMETER tag not found"
    End If
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    FindList = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
FindListExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlRequestNode = Nothing
    Set objGlobalParameterDO = Nothing
    Exit Function

FindListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    FindList = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   go to clean-up section
    Resume FindListExit
End Function
Public Function FindCurrentParameterList(ByVal vstrXMLRequest As String) As String
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
On Error GoTo FindCurrentParameterListVbErr
    
    Const strFunctionName As String = "FindCurrentParameterList"
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim strXMLData As String
    Dim xmlParamList As IXMLDOMElement
    Dim xmlNameList As IXMLDOMNodeList
    Dim xmlDataDoc As New FreeThreadedDOMDocument40
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlElement As IXMLDOMElement
    Dim lngErrNumber As Long
    
    'Create Response Object
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'append child
    Set objXmlDataNode = xmlOut.appendChild(xmlResponseElem)
    'set type attribute and value as suscess
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    'load Request into an XML Document
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    'Check for GLobalParameter Node
    Set xmlRequestNode = objXmlIn.getElementsByTagName("GLOBALPARAMETER").Item(0)
    If xmlRequestNode Is Nothing Then
        'If null then raise error
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingPrimaryTag, "GLOBALPARAMETER tag not found"
    End If
    'store all name nodes
    Set xmlNameList = objXmlIn.selectNodes(".//GLOBALPARAMETER/NAME")
    'Create an elemetn for response object
    Set xmlParamList = xmlOut.createElement("GLOBALPARAMETERLIST")
    xmlResponseElem.appendChild xmlParamList
    'loop through the name list node
    For Each xmlElement In xmlNameList
        'If a parameter is not found then raise an cutomised error message
        On Error GoTo FindCurrentParameterListRecordNotFoundVbErr
        strXMLData = objGlobalParameterDO.GetCurrentParameter(xmlElement.Text)
        On Error GoTo FindCurrentParameterListVbErr
        If lngErrNumber > 0 Then
            If lngErrNumber = oeRecordNotFound Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 291, , xmlElement.Text
            Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, m_objErrAssist.GetOmigaErrorNumber(lngErrNumber)
            End If
        End If
        'If the parameter exists then
        If Len(strXMLData) > 0 Then
            'Load into an XMlDoc
            Set xmlDataDoc = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
            'append to the response object
            xmlParamList.appendChild xmlDataDoc.documentElement
        End If
        strXMLData = vbNullString
    Next
    FindCurrentParameterList = xmlResponseElem.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindCurrentParameterListExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set xmlOut = Nothing
    Set xmlRequestNode = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlParamList = Nothing
    Set xmlNameList = Nothing
    Set xmlDataDoc = Nothing
    Set xmlElement = Nothing
    Exit Function

FindCurrentParameterListRecordNotFoundVbErr:
    lngErrNumber = m_objErrAssist.GetOmigaErrorNumber(Err.Number)
    Resume Next

FindCurrentParameterListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    FindCurrentParameterList = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   go to clean-up section
    Resume FindCurrentParameterListExit
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
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    Dim objGlobalParameterTxBO As GlobalParameterTxBO
    If m_objContext Is Nothing Then
        Set objGlobalParameterTxBO = New GlobalParameterTxBO
    Else
        Set objGlobalParameterTxBO = m_objContext.CreateInstance(App.Title & ".GlobalParameterTxBO")
    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   call Business Transaction Object DeleteAll function
    objGlobalParameterTxBO.DeleteAll (vstrXMLRequest)
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    DeleteAll = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
DeleteAllExit:
    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objGlobalParameterTxBO = Nothing
    Exit Function

DeleteAllVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    DeleteAll = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    '   go to clean-up section
    Resume DeleteAllExit
End Function
Public Function GetCurrentParameter(ByVal vstrParameterName As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the current values for the parameter passed in
' pass:         vstrParameterName   Parameter for which current data is required
' return:       GetCurrentParameter xml Response data stream containing results of
'               operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterVbErr
    
    Const strFunctionName As String = "GetCurrentParameter"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objGlobalParameterDO As GlobalParameterDO
    
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    Dim strXMLData As String
          
    strXMLData = objGlobalParameterDO.GetCurrentParameter(vstrParameterName)
    Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
        
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    GetCurrentParameter = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetCurrentParameterExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objGlobalParameterDO = Nothing
    Exit Function

GetCurrentParameterVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    GetCurrentParameter = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   go to clean-up section
    Resume GetCurrentParameterExit
End Function
Public Function IsMultipleLender() As String
' header ----------------------------------------------------------------------------------
' description:  Indicates whether or not the application is in mutiple lender mode
' pass:         n/a
' return:       IsMultipleLender xml Response data stream containing results of
'               operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsMultipleLenderVbErr
    
    Const strFunctionName As String = "IsMultipleLender"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Dim strItem As String
    Dim strRegDir As String
    strRegDir = "SOFTWARE\" + strAppName + "\" + strREGISTRY_SECTION
    strItem = QueryValue(HKEY_LOCAL_MACHINE, strRegDir, "Multiple Lender")
    If (Len(strItem) = 0) Or _
        (strItem <> "1") Then
        strItem = "0"
    End If
    Set objXmlNode = objXmlOut.createElement("MULTIPLELENDER")
    objXmlNode.Text = strItem
    objXmlDataNode.appendChild objXmlNode
    IsMultipleLender = objXmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
IsMultipleLenderExit:
    
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlOut = Nothing
    Set objXmlNode = Nothing
    Exit Function
    
IsMultipleLenderVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    IsMultipleLender = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   go to clean-up section
    Resume IsMultipleLenderExit
End Function
Public Function IsTaskManager() As Boolean
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo IsTaskManagerVbErr
    Const cstrFunctionName As String = "IsTaskManager"
                                           
    Dim strItem As String
    Dim strRegDir As String
    Dim blnTaskManager As Boolean
    
    blnTaskManager = True
    strRegDir = "SOFTWARE\" + strAppName + "\" + strREGISTRY_SECTION
    strItem = QueryValue(HKEY_LOCAL_MACHINE, strRegDir, "Task Manager")
    If (Len(strItem) = 0) Or (strItem <> "1") Then
        blnTaskManager = False
    End If
                                          
    IsTaskManager = blnTaskManager
    m_objContext.SetComplete
IsTaskManagerExit:
    
    Exit Function
IsTaskManagerVbErr:
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    IsTaskManager = False
    m_objContext.SetComplete
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
'BM0493 MDC 10/04/2003
Public Function GetCurrentParameterListEx(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the current values for the parameter passed in
' pass:         vstrParameterName   Parameter for which current data is required
' return:       GetCurrentParameter xml Response data stream containing results of
'               operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterListExVbErr
    
    Const strFunctionName As String = "GetCurrentParameterListEx"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objGlobalParameterDO As GlobalParameterDO
    
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    Dim strXMLData As String
          
    strXMLData = objGlobalParameterDO.GetCurrentParameterListEx(vstrXMLRequest)
    Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    objXmlResponseElem.appendChild objXmlData.documentElement
        
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    GetCurrentParameterListEx = objXmlResponseElem.xml

GetCurrentParameterListExExit:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Set objXmlResponseElem = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objGlobalParameterDO = Nothing
    Exit Function

GetCurrentParameterListExVbErr:
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetCurrentParameterListEx = m_objErrAssist.CreateErrorResponse
        
    '   go to clean-up section
    Resume GetCurrentParameterListExExit
End Function
'BM0493 MDC 10/04/2003 - End

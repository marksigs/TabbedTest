VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "GlobalBandedParameterBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      GlobalBandedParameterBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Global Banded Parameter Business Object which 'supports transactions' only.
'               Methods which do not require transaction support reside in this
'               class. Any methods that require transactions will be delegated to
'               GlobalBandedParameterTxBO
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'IK    30/06/99 Created
'DRC    3/10/01 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
 '=============================================
 'Constant Declaration Section
 '=============================================
 '=============================================
 'Variable Declaration Section
 '=============================================
' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
Private Const cstrTABLE_NAME = "GLOBALBANDEDPARAMETER"
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
    
    Dim strFunctionName As String
    strFunctionName = "Create"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
'    Dim objErrAssist As New ErrAssist
    Dim objXmlNode As IXMLDOMNode
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
' TODO ************************************************************************
'   replace Omiga4TxBO reference with actual TxBO required
' *****************************************************************************
'    Dim objOmiga4TxBO As Omiga4TxBO
'
'    If objContext Is Nothing Then
'        Set objOmiga4TxBO = New Omiga4TxBO
'    Else
'        Set objOmiga4TxBO = objContext.CreateInstance(App.Title & ".Omiga4TxBO")
'    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
' TODO ************************************************************************
'   add processing here
' *****************************************************************************
    
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    Validate (vstrXMLRequest)
    '   call Business Transaction Object Create function
'    objOmiga4TxBO.Create (vstrXMLRequest)
    
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
'    Set objErrAssist = Nothing
'    Set objOmiga4TxBO = Nothing
    Exit Function
CreateVbErr:
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
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
    
    Dim strFunctionName As String
    strFunctionName = "Update"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
' TODO ************************************************************************
'   replace Omiga4TxBO reference with actual TxBO required
' *****************************************************************************
'    Dim objOmiga4TxBO As Omiga4TxBO
'
'    If objContext Is Nothing Then
'        Set objOmiga4TxBO = New Omiga4TxBO
'    Else
'        Set objOmiga4TxBO = objContext.CreateInstance(App.Title & ".Omiga4TxBO")
'    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    If Not objXmlResponseElem Is Nothing Then
        Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    End If
    ' Call scripting interface
    Validate (vstrXMLRequest)
'    objOmiga4TxBO.Update vstrXMLRequest
    
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
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
'    Set objOmiga4TxBO = Nothing
    
    Exit Function
UpdateVbErr:
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Update = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
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
    
    Dim strFunctionName As String
    strFunctionName = "Delete"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace Omiga4TxBO reference with actual TxBO required
' *****************************************************************************
'    Dim objOmiga4TxBO As Omiga4TxBO
'
'    If objContext Is Nothing Then
'        Set objOmiga4TxBO = New Omiga4TxBO
'    Else
'        Set objOmiga4TxBO = objContext.CreateInstance(App.Title & ".Omiga4TxBO")
'    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
           
    '   call Business Transaction Object Delete function
'   objOmiga4TxBO.Delete (vstrXMLRequest)
    
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
'    Set objErrAssist = Nothing
'    Set objOmiga4TxBO = Nothing
    Exit Function
DeleteVbErr:
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
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
    
    Dim strFunctionName As String
    strFunctionName = "GetData"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace Omiga4DO reference with actual DO required
' *****************************************************************************
'    Dim objOmiga4DO As Omiga4DO
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
'    If objContext Is Nothing Then
'        Set objOmiga4DO = New Omiga4DO
'    Else
'        Set objOmiga4DO = objContext.CreateInstance(App.Title & ".Omiga4DO")
'    End If
                
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    '   call Data Object GetData function
    Dim strXMLData As String
    If Not objXmlRequestNode Is Nothing Then
'        strXMLData = objOmiga4DO.GetData(objXmlRequestNode.xml)
        Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
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
'    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
'    Set objErrAssist = Nothing
'    Set objOmiga4DO = Nothing
    Exit Function
GetDataVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
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
    Dim strFunctionName As String
    strFunctionName = "Validate"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlDerived As New FreeThreadedDOMDocument40
    Dim objXmlBeforeData As IXMLDOMElement
    Dim objXmlAfterData As IXMLDOMElement
    Dim objXmlAfterNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
' TODO ************************************************************************
'   replace Omiga4DO reference with actual DO required
' *****************************************************************************
'    Dim objOmiga4DO As Omiga4DO
    
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
'    If objContext Is Nothing Then
'        Set objOmiga4DO = New Omiga4DO
'    Else
'        Set objOmiga4DO = objContext.CreateInstance(App.Title & ".Omiga4DO")
'    End If
    
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
'        strXMLAfterData = objOmiga4DO.AddDerivedData(objXmlAfterData.xml)
        Set objXmlDerived = m_objXmlAssist.load(strXMLAfterData, TypeName(Me), strFunctionName)
        objXmlAfterNode.replaceChild objXmlDerived.documentElement, objXmlAfterData
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
            Dim objErrAssist As New ErrAssist
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4SurnameChange
        End If
    End If
'   -----------------------------------------------------------------------------
'   hard-coded validation OK
'   call scripting interface
'   -----------------------------------------------------------------------------
    Dim objScriptInterface As New ScriptInterface
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
'    Set objXmlAssist = Nothing
'    Set objOmiga4DO = Nothing
Exit Sub
ValidateVbErr:
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
'    Set objXmlAssist = Nothing
'    Set objOmiga4DO = Nothing
    
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
    
    Dim strFunctionName As String
    strFunctionName = "FindList"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace Omiga4DO reference with actual DO required
' *****************************************************************************
'    Dim objOmiga4DO As Omiga4DO
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
'    If objContext Is Nothing Then
'        Set objOmiga4DO = New Omiga4DO
'    Else
'        Set objOmiga4DO = objContext.CreateInstance(App.Title & ".Omiga4DO")
'    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    Dim strXMLData As String
    '   call Data Object FindList function
    If Not objXmlRequestNode Is Nothing Then
'        strXMLData = objOmiga4DO.FindList(objXmlRequestNode.xml)
        Set objXmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
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
'    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
'    Set objErrAssist = Nothing
'    Set objOmiga4DO = Nothing
    Exit Function
FindListVbErr:
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindList = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
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
    
    Dim strFunctionName As String
    strFunctionName = "DeleteAll"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace Omiga4TxBO reference with actual TxBO required
' *****************************************************************************
'    Dim objOmiga4TxBO As Omiga4TxBO
'
'    If objContext Is Nothing Then
'        Set objOmiga4TxBO = New Omiga4TxBO
'    Else
'        Set objOmiga4TxBO = objContext.CreateInstance(App.Title & ".Omiga4TxBO")
'    End If
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   call Business Transaction Object DeleteAll function
'    objOmiga4TxBO.DeleteAll (vstrXMLRequest)
    
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
'    Set objErrAssist = Nothing
'    Set objOmiga4TxBO = Nothing
    Exit Function
DeleteAllVbErr:
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    DeleteAll = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    '   go to clean-up section
    Resume DeleteAllExit
End Function
Public Function GetCurrentParameter(ByVal vstrParameterName As String, _
                                    ByVal vstrQualifier As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the current values for the parameter and qualifier passed in
' pass:         vstrParameterName   Parameter for which current data is required
'               vstrQualifier       Qualifier to select correct band
' return:       GetCurrentParameter xml Response data stream containing results of
'               operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetCurrentParameter"
'    Dim objContext As ObjectContext
    strFunctionName = "GetData"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    'Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objGlobalBandedParameterDO As GlobalBandedParameterDO
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    If m_objContext Is Nothing Then
        Set objGlobalBandedParameterDO = New GlobalBandedParameterDO
    Else
        Set objGlobalBandedParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalBandedParameterDO")
    End If
    Dim strXMLData As String
          
    strXMLData = objGlobalBandedParameterDO.GetCurrentParameter(vstrParameterName, vstrQualifier)
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
'    Set objXmlAssist = Nothing
'    Set objXmlRequestNode = Nothing
'    Set objErrAssist = Nothing
    Set objGlobalBandedParameterDO = Nothing
    Exit Function
GetCurrentParameterVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
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

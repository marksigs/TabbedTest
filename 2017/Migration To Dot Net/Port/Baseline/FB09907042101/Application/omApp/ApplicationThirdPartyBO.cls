VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationThirdPartyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'ApplicationThirdPartyBO' with actual class name
' *****************************************************************************

'Workfile:      ApplicationThirdPartyBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for ApplicationThirdParty Business Object which 'supports transactions'
'               only
'               i.e. any code that does not require transaction support will reside in
'               this first level module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level)
'               module (ApplicationThirdPartyTxBO in this example)
'
'Dependencies:  List any other dependent components
'               e.g. ApplicationThirdPartyTxBO, ApplicationThirdPartyDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     11/01/00    Created
'JR     20/04/01    SYS2048 - Commented out objContext.setabort/complete in GetLegalRep method
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit

Implements IApplicationThirdPartyBO

'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     31/07/2002  BMIDS00075 - Created SaveLegalRep() ; IApplicationThirdPartyBO_SaveLegalRep()
'------------------------------------------------------------------------------------------
Public Function CreateArchitect(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateArchitect a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateArchitectVbErr

    Const strFunctionName As String = "CreateArchitect"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONARCHITECT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_CreateArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateArchitect = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateArchitectExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateArchitect = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArchitectExit

End Function

Public Function CreateBankBuildingSociety(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateBankBuildingSociety a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateBankBuildingSocietyVbErr

    Const strFunctionName As String = "CreateBankBuildingSociety"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONBANKBUILDINGSOC/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_CreateBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateBankBuildingSociety = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateBankBuildingSocietyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateBankBuildingSociety = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankBuildingSocietyExit

End Function

Public Function CreateBuilder(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateBuilder a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateBuilderVbErr

    Const strFunctionName As String = "CreateBuilder"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONBUILDER/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_CreateBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateBuilder = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateBuilderExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateBuilder = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBuilderExit

End Function

Public Function CreateEstateAgent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateEstateAgent a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateEstateAgentVbErr

    Const strFunctionName As String = "CreateEstateAgent"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONESTATEAGENT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_CreateEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateEstateAgent = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateEstateAgentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateEstateAgent = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEstateAgentExit

End Function

Public Function CreateLegalRep(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateLegalRep a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateLegalRepVbErr

    Const strFunctionName As String = "CreateLegalRep"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONLEGALREP/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_CreateLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateLegalRep = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateLegalRepExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateLegalRep = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLegalRepExit

End Function

Public Function DeleteArchitect(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteArchitect a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArchitectVbErr

    Const strFunctionName As String = "DeleteArchitect"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_DeleteArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteArchitect = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteArchitectExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteArchitect = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArchitectExit

End Function

Public Function DeleteBankBuildingSociety(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteBankBuildingSociety a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankBuildingSocietyVbErr

    Const strFunctionName As String = "DeleteBankBuildingSociety"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_DeleteBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteBankBuildingSociety = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankBuildingSocietyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteBankBuildingSociety = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankBuildingSocietyExit

End Function

Public Function DeleteBuilder(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteBuilder a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBuilderVbErr

    Const strFunctionName As String = "DeleteBuilder"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_DeleteBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteBuilder = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBuilderExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteBuilder = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBuilderExit

End Function

Public Function DeleteEstateAgent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteEstateAgent a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEstateAgentVbErr

    Const strFunctionName As String = "DeleteEstateAgent"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_DeleteEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteEstateAgent = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteEstateAgentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteEstateAgent = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEstateAgentExit

End Function

Public Function DeleteLegalRep(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteLegalRep a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLegalRepVbErr

    Const strFunctionName As String = "DeleteLegalRep"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_DeleteLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteLegalRep = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLegalRepExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteLegalRep = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLegalRepExit

End Function

Public Function FindApplicationThirdPartyList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindApplicationThirdPartyListVbErr

    Const strFunctionName As String = "FindApplicationThirdPartyList"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_FindApplicationThirdPartyList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindApplicationThirdPartyList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindApplicationThirdPartyListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindApplicationThirdPartyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindApplicationThirdPartyList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationThirdPartyListExit

End Function

Public Function FindBankBuildingSocietyList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindBankBuildingSocietyListVbErr

    Const strFunctionName As String = "FindBankBuildingSocietyList"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_FindBankBuildingSocietyList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindBankBuildingSocietyList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindBankBuildingSocietyListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindBankBuildingSocietyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBankBuildingSocietyList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankBuildingSocietyListExit

End Function

Public Function GetArchitect(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetArchitectVbErr

    Const strFunctionName As String = "GetArchitect"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_GetArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetArchitect = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetArchitectExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetArchitect = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetArchitectExit

End Function

Public Function GetBankBuildingSociety(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetBankBuildingSocietyVbErr

    Const strFunctionName As String = "GetBankBuildingSociety"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_GetBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetBankBuildingSociety = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetBankBuildingSocietyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetBankBuildingSociety = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBankBuildingSocietyExit

End Function

Public Function GetBuilder(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetBuilderVbErr

    Const strFunctionName As String = "GetBuilder"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_GetBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetBuilder = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetBuilderExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetBuilder = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBuilderExit

End Function

Public Function GetEstateAgent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetEstateAgentVbErr

    Const strFunctionName As String = "GetEstateAgent"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_GetEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetEstateAgent = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetEstateAgentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetEstateAgent = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEstateAgentExit

End Function

Public Function GetLegalRep(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLegalRepVbErr

    Const strFunctionName As String = "GetLegalRep"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationThirdPartyBO_GetLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetLegalRep = xmlResponseElem.xml
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If

GetLegalRepExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLegalRep = objErrAssist.CreateErrorResponse

' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetAbort
'    End If
    
    ' go to clean-up section
    Resume GetLegalRepExit

End Function

Public Function UpdateArchitect(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateArchitect a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArchitectVbErr

    Const strFunctionName As String = "UpdateArchitect"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONARCHITECT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_UpdateArchitect(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateArchitect = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateArchitectExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateArchitect = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArchitectExit
    
End Function

Public Function UpdateBankBuildingSociety(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateBankBuildingSociety a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankBuildingSocietyVbErr

    Const strFunctionName As String = "UpdateBankBuildingSociety"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONBANKBUILDINGSOC/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_UpdateBankBuildingSociety(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateBankBuildingSociety = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateBankBuildingSocietyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateBankBuildingSociety = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankBuildingSocietyExit
    
End Function

Public Function UpdateBuilder(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateBuilder a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBuilderVbErr

    Const strFunctionName As String = "UpdateBuilder"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONBUILDER/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_UpdateBuilder(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateBuilder = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateBuilderExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateBuilder = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBuilderExit
    
End Function

Public Function UpdateEstateAgent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateEstateAgent a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEstateAgentVbErr

    Const strFunctionName As String = "UpdateEstateAgent"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONESTATEAGENT/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_UpdateEstateAgent(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateEstateAgent = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateEstateAgentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateEstateAgent = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEstateAgentExit
    
End Function

Public Function UpdateLegalRep(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateLegalRep a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLegalRepVbErr

    Const strFunctionName As String = "UpdateLegalRep"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim strDirectoryAddressIndicator As String
    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' CreateEstateAgent default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strDirectoryAddressIndicator = objXmlAssist.GetTagValue(xmlIn.documentElement, "DIRECTORYADDRESSINDICATOR")
    If strDirectoryAddressIndicator = "" Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "DIRECTORYADDRESSINDICATOR must be specified."
    End If

    ' Ensure that the directory GUID is valid in the context of the entity type
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//APPLICATIONLEGALREP/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = IApplicationThirdPartyBO_ValidateLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_UpdateLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLegalRep = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateLegalRepExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateLegalRep = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLegalRepExit
    
End Function

Public Function SaveLegalRep(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo SaveLegalRepVbErr

    Const strFunctionName As String = "SaveLegalRep"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationThirdPartyBO_SaveLegalRep(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveLegalRep = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveLegalRepExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveLegalRep = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLegalRepExit

End Function

Private Function IApplicationThirdPartyBO_CreateArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateArchitect a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateArchitectVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_CreateArchitect"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateArchitect function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.CreateArchitect(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationThirdPartyBO_CreateArchitect = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateArchitectExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_CreateArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArchitectExit
End Function

Private Function IApplicationThirdPartyBO_CreateBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankBuildingSociety a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankBuildingSocietyVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_CreateBankBuildingSociety"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateBankBuildingSociety function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.CreateBankBuildingSociety(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationThirdPartyBO_CreateBankBuildingSociety = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankBuildingSocietyExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_CreateBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyBO_CreateBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBuilder a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBuilderVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_CreateBuilder"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateBuilder function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.CreateBuilder(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationThirdPartyBO_CreateBuilder = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBuilderExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_CreateBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBuilderExit
End Function

Private Function IApplicationThirdPartyBO_CreateEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateEstateAgent a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateEstateAgentVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_CreateEstateAgent"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateEstateAgent function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.CreateEstateAgent(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationThirdPartyBO_CreateEstateAgent = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateEstateAgentExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_CreateEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEstateAgentExit
End Function

Private Function IApplicationThirdPartyBO_CreateLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLegalRep a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_CreateLegalRep"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateLegalRep function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.CreateLegalRep(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationThirdPartyBO_CreateLegalRep = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLegalRepExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_CreateLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLegalRepExit
End Function

Private Function IApplicationThirdPartyBO_DeleteArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteArchitect a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArchitectVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_DeleteArchitect"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteArchitect function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.DeleteArchitect(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_DeleteArchitect = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteArchitectExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_DeleteArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArchitectExit
End Function

Private Function IApplicationThirdPartyBO_DeleteBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankBuildingSociety a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankBuildingSocietyVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_DeleteBankBuildingSociety"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteBankBuildingSociety function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.DeleteBankBuildingSociety(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_DeleteBankBuildingSociety = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBankBuildingSocietyExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_DeleteBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyBO_DeleteBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBuilder a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBuilderVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_DeleteBuilder"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteBuilder function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.DeleteBuilder(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_DeleteBuilder = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteBuilderExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_DeleteBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBuilderExit
End Function

Private Function IApplicationThirdPartyBO_DeleteEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteEstateAgent a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEstateAgentVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_DeleteEstateAgent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteEstateAgent function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.DeleteEstateAgent(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_DeleteEstateAgent = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteEstateAgentExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_DeleteEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEstateAgentExit
End Function

Private Function IApplicationThirdPartyBO_DeleteLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLegalRep a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_DeleteLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteLegalRep function
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.DeleteLegalRep(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_DeleteLegalRep = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLegalRepExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_DeleteLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLegalRepExit
End Function

Private Function IApplicationThirdPartyBO_FindApplicationThirdPartyList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationThirdPartyListVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_FindApplicationThirdPartyList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATION" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATION").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATION" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.FindApplicationThirdPartyList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IApplicationThirdPartyBO_FindApplicationThirdPartyList = xmlResponseElem.cloneNode(True)

FindApplicationThirdPartyListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindApplicationThirdPartyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_FindApplicationThirdPartyList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationThirdPartyListExit
End Function

Private Function IApplicationThirdPartyBO_FindBankBuildingSocietyList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindBankBuildingSocietyListVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_FindBankBuildingSocietyList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONBANKBUILDINGSOC" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONBANKBUILDINGSOC").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONBANKBUILDINGSOC" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.FindBankBuildingSocietyList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IApplicationThirdPartyBO_FindBankBuildingSocietyList = xmlResponseElem.cloneNode(True)

FindBankBuildingSocietyListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindBankBuildingSocietyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_FindBankBuildingSocietyList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindBankBuildingSocietyListExit
End Function

Private Function IApplicationThirdPartyBO_GetArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetArchitectVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_GetArchitect"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONARCHITECT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONARCHITECT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONARCHITECT" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.GetArchitect(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationThirdPartyBO_GetArchitect = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetArchitectExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_GetArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetArchitectExit
End Function

Private Function IApplicationThirdPartyBO_GetBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetBankBuildingSocietyVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_GetBankBuildingSociety"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONBANKBUILDINGSOC" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONBANKBUILDINGSOC").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONBANKBUILDINGSOC" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.GetBankBuildingSociety(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationThirdPartyBO_GetBankBuildingSociety = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetBankBuildingSocietyExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_GetBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyBO_GetBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetBuilderVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_GetBuilder"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONBUILDER" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONBUILDER").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONBUILDER" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.GetBuilder(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationThirdPartyBO_GetBuilder = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetBuilderExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_GetBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBuilderExit
End Function

Private Function IApplicationThirdPartyBO_GetEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetEstateAgentVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_GetEstateAgent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONESTATEAGENT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONESTATEAGENT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONESTATEAGENT" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.GetEstateAgent(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationThirdPartyBO_GetEstateAgent = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetEstateAgentExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_GetEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetEstateAgentExit
End Function

Private Function IApplicationThirdPartyBO_GetLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_GetLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "APPLICATIONLEGALREP" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONLEGALREP").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONLEGALREP" & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationThirdPartyDO.GetLegalRep(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationThirdPartyBO_GetLegalRep = xmlResponseElem.cloneNode(True)
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
    
GetLegalRepExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_GetLegalRep = objErrAssist.CreateErrorResponseEx()

' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
    
    ' go to clean-up section
    Resume GetLegalRepExit
End Function

Private Function IApplicationThirdPartyBO_SaveLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo SaveLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_SaveLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.SaveLegalRep(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_SaveLegalRep = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveLegalRepExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_SaveLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLegalRepExit
End Function

Private Function IApplicationThirdPartyBO_UpdateArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateArchitect a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArchitectVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_UpdateArchitect"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.UpdateArchitect(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_UpdateArchitect = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateArchitectExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateArchitectVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_UpdateArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArchitectExit
End Function

Private Function IApplicationThirdPartyBO_UpdateBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankBuildingSociety a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankBuildingSocietyVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_UpdateBankBuildingSociety"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.UpdateBankBuildingSociety(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_UpdateBankBuildingSociety = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankBuildingSocietyExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateBankBuildingSocietyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_UpdateBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyBO_UpdateBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBuilder a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBuilderVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_UpdateBuilder"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.UpdateBuilder(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_UpdateBuilder = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBuilderExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateBuilderVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_UpdateBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBuilderExit
End Function

Private Function IApplicationThirdPartyBO_UpdateEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateEstateAgent a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEstateAgentVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_UpdateEstateAgent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.UpdateEstateAgent(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_UpdateEstateAgent = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEstateAgentExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateEstateAgentVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_UpdateEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEstateAgentExit
End Function

Private Function IApplicationThirdPartyBO_UpdateLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLegalRep a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyBO_UpdateLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationThirdPartyTxBO As IApplicationThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationThirdPartyTxBO = New ApplicationThirdPartyTxBO
    Else
        Set objIApplicationThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationThirdPartyTxBO.UpdateLegalRep(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationThirdPartyBO_UpdateLegalRep = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateLegalRepExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateLegalRepVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_UpdateLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLegalRepExit
End Function

Private Function IApplicationThirdPartyBO_ValidateArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyBO_ValidateArchitect"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationThirdPartyBO_ValidateArchitect = xmlResponseElem.cloneNode(True)
    
ValidateArchitectExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateArchitectVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_ValidateArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateArchitectExit

End Function

Private Function IApplicationThirdPartyBO_ValidateBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyBO_ValidateBankBuildingSociety"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationThirdPartyBO_ValidateBankBuildingSociety = xmlResponseElem.cloneNode(True)
    
ValidateBankBuildingSocietyExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateBankBuildingSocietyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_ValidateBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyBO_ValidateBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyBO_ValidateBuilder"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationThirdPartyBO_ValidateBuilder = xmlResponseElem.cloneNode(True)
    
ValidateBuilderExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateBuilderVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_ValidateBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateBuilderExit
End Function

Private Function IApplicationThirdPartyBO_ValidateEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyBO_ValidateEstateAgent"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationThirdPartyBO_ValidateEstateAgent = xmlResponseElem.cloneNode(True)
    
ValidateEstateAgentExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateEstateAgentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_ValidateEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateEstateAgentExit
End Function

Private Function IApplicationThirdPartyBO_ValidateLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyBO_ValidateLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationThirdPartyBO_ValidateLegalRep = xmlResponseElem.cloneNode(True)
    
ValidateLegalRepExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateLegalRepVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationThirdPartyBO_ValidateLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateLegalRepExit
End Function


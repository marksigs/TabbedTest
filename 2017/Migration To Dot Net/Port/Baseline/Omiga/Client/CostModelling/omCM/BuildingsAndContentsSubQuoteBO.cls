VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BuildingsAndContentsSubQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BuildingsAndContentsSubQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for BuildingsAndContentsSubQuote Business Object which 'supports transactions'
'               only
'               i.e. any code that does not require transaction support will reside in
'               this first level module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level)
'               module (BuildingsAndContentsSubQuoteTxBO in this example)
'
'Dependencies:  List any other dependent components
'               e.g. BuildingsAndContentsSubQuoteTxBO, BuildingsAndContentsSubQuoteDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'JLD    21/12/1999  Created
'APS    05/04/2000  Added Valuable processing to SaveSubQuote
'APS    05/05/2000  Handle "Record Not Found" on FindValuablesList as its ok not to have any valuables
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit

Implements IBuildingsAndContentsSubQuoteBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================


'BMIDS History:**************************************************************************
'
'Prog   Date        Description
'GD     16/07/2002  BMIDS00045 - IBuildingsAndContentsSubQuoteBO_SaveSubQuote reworked
'MC     16/06/2004  BMIDS763    GetInsuranceAdminFee Member Functions added
'
'****************************************************************************************

Public Function CreateSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateSubQuoteDetails"
    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_CreateSubQuoteDetails(xmlIn.documentElement)
    xmlResponseElem.appendChild xmlTempResponseNode
    objErrAssist.CheckResponse xmlResponseElem, xmlResponseElem, True
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    'xmlResponseElem.appendChild xmlTempResponseNode
    CreateSubQuoteDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateSubQuoteDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function BuildBCCalcInterface(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo BuildInterfaceVbErr

    Const strFunctionName As String = "BuildBCCalcInterface"

    
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
        
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_BuildBCCalcInterface(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    BuildBCCalcInterface = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

BuildInterfaceExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

BuildInterfaceVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    BuildBCCalcInterface = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume BuildInterfaceExit

End Function

'*****************************************************************************
'   DATE    :   16/06/2004
'   CC      :   BMIDS763
'*****************************************************************************
Public Function GetInsuranceAdminFee(ByVal vstrXMLRequest As String) As String
        
    On Error GoTo ErrorGetInsuranceAdminFee

    Const strFunctionName As String = "GetInsuranceAdminFee"
    
    Dim objContext              As ObjectContext
    Dim xmlIn                   As FreeThreadedDOMDocument40
    Dim objErrAssist            As ErrAssist
    Dim objXmlAssist            As XMLAssist
    Dim objIScriptInterface     As IScriptInterface
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlResponseElem         As IXMLDOMElement
    Dim xmlTempResponseNode     As IXMLDOMNode
    Dim sReturn                 As String
    
    '*=[MC]Instantiate Object
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    '*=[MC]Get COM+ Context
    Set objContext = GetObjectContext()
        
    '*=[MC]Create a default response block and assume result as success
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    '*=[MC]Temp ResponseNode to pass the check response.
    Set xmlTempResponseNode = xmlOut.documentElement
    
    '*=[MC]Call Interface Method
    Call IBuildingsAndContentsSubQuoteBO_GetInsuranceAdminFee(xmlIn.documentElement, xmlResponseElem)
    '*=[MC]check the response. If valid attach response
    Call objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem, True)
    Call objXmlAssist.AttachResponseData(xmlResponseElem, xmlTempResponseNode)
    '*=Write to File
    Call objXmlAssist.WriteXMLToFile(xmlResponseElem.xml)
    '*=Return XML
    sReturn = xmlResponseElem.xml

ExitGetInsuranceAdminFee:
    '*=[MC]Cleanup here
    GetInsuranceAdminFee = sReturn
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ErrorGetInsuranceAdminFee:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        '*=Rebind XML
        sReturn = xmlResponseElem.xml
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    sReturn = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ExitGetInsuranceAdminFee
    
    Exit Function
    Resume
End Function


Public Function CreateValuablesOverLimit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateValuablesOverLimit"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_CreateValuablesOverLimit(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CreateValuablesOverLimit = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateValuablesOverLimit = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function CreateNewSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr

    Const strFunctionName As String = "CreateNewSubQuote"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_CreateNewSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CreateNewSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateNewSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateNewSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewSubQuoteExit

End Function

Public Function CreateFirstSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr

    Const strFunctionName As String = "CreateFirstSubQuote"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_CreateFirstSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    
    CreateFirstSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateFirstSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateFirstSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateFirstSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstSubQuoteExit

End Function

Public Function SaveBCNotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveBCNotesVbErr

    Const strFunctionName As String = "SaveBCNotes"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn  As FreeThreadedDOMDocument40, _
        xmlOut As New FreeThreadedDOMDocument40
        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_SaveBCNotes(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveBCNotes = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveBCNotesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing

    Exit Function

SaveBCNotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    SaveBCNotes = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveBCNotesExit

End Function

Public Function SaveSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveSubQuoteVbErr

    Const strFunctionName As String = "SaveSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn  As FreeThreadedDOMDocument40, _
        xmlOut As New FreeThreadedDOMDocument40
        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_SaveSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing

    Exit Function

SaveSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    SaveSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveSubQuoteExit

End Function

Public Function UpdateSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "UpdateSubQuoteDetails"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_UpdateSubQuoteDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateSubQuoteDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    UpdateSubQuoteDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
    
End Function

Public Function ValidateSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXMLRequest  xml Request data stream of the format
'
' return:      xml Response data stream containing results of operation
'              either: TYPE="SUCCESS" and 1 (True) if Quote is Valid Else 0 (False)
'              or: TYPE="SYSERR" and <ERROR> element
'NOTE ** CURRENTLY THIS METHOD JUST RETURNS TRUE FOR ALL QUOTES (NOT FULLY IMPLEMENTED)
'------------------------------------------------------------------------------------------
On Error GoTo ValidateSubQuoteVbErr
    
    Const strFunctionName As String = "ValidateSubQuote"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_ValidateSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ValidateSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlIn = Nothing
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
    Exit Function

ValidateSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ValidateSubQuote = objErrAssist.CreateErrorResponse
    
    ' go to clean-up section
    Resume ValidateSubQuoteExit

End Function

Public Function UpdateValuablesOverLimit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "UpdateValuablesOverLimit"

    
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
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_UpdateValuablesOverLimit(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateValuablesOverLimit = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    UpdateValuablesOverLimit = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
    
End Function

Public Function DeleteValuablesOverLimit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "DeleteValuablesOverLimit"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_DeleteValuablesOverLimit(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteValuablesOverLimit = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    DeleteValuablesOverLimit = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Public Function GetBCNotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetBCNotesVbErr

    Const strFunctionName As String = "GetBCNotes"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_GetBCNotes(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetBCNotes = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetBCNotesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetBCNotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetBCNotes = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBCNotesExit

End Function

Public Function GetPaymentFrequency(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPaymentFrequencyVbErr

    Const strFunctionName As String = "GetPaymentFrequency"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_GetPaymentFrequency(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPaymentFrequency = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetPaymentFrequencyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPaymentFrequencyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetPaymentFrequency = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPaymentFrequencyExit

End Function

Public Function GetSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetSubQuoteDetailsVbErr

    Const strFunctionName As String = "GetSubQuoteDetails"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_GetSubQuoteDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetSubQuoteDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetSubQuoteDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetSubQuoteDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetSubQuoteDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetSubQuoteDetailsExit

End Function

Public Function FindValuablesOverLimitList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindValuablesOverLimitList"

    
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
    Set xmlTempResponseNode = _
            IBuildingsAndContentsSubQuoteBO_FindValuablesOverLimitList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindValuablesOverLimitList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindValuablesOverLimitList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function FindList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindList"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_FindList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function DeleteAll(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr

    Const strFunctionName As String = "DeleteAll"

    
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
    Set xmlTempResponseNode = IBuildingsAndContentsSubQuoteBO_DeleteAll(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DeleteAll = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteAllExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteAllVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    DeleteAll = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteAllExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_BuildBCCalcInterface(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create an XML containing the interface required for the insurance calculations
' pass:         vxmlRequest  xml Request data stream containing data to be interfaced
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo BuildInterfaceVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_BuildBCCalcInterface"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim strBuildingsCoverAmt As String
    Dim strContentsCoverAmt As String
    Dim strPolarisSchemeVersion As String
    Dim strPolarisSchemeIdentifier As String
    Dim strPolarisApplicationVersion As String
    Dim strPolarisApplicationIdentifier As String
    Dim strPolarisTransactionType As String
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTopNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlPostCodeNode As IXMLDOMNode
    Dim iNoOfClaims As Integer
    Dim strQuotationType As String
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strResponse As String
    Dim objResponseDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim objComboDO As ComboDO
    
    If objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objComboDO = New ComboDO
    Else
        Set objGlobalParameterDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' APS 05/04/00 - Changed to use new GetCurrentParameterByType
    Dim vFieldValue As Variant
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("BuildingsCoverAmount", "AMOUNT")
    strBuildingsCoverAmt = CStr(vFieldValue)
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("ContentsCoverAmount", "AMOUNT")
    strContentsCoverAmt = CStr(vFieldValue)
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("PolarisSchemeVersion", "STRING")
    strPolarisSchemeVersion = CStr(vFieldValue)
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("PolarisSchemeIdentifier", "STRING")
    strPolarisSchemeIdentifier = CStr(vFieldValue)
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("PolarisApplicationVersion", "STRING")
    strPolarisApplicationVersion = CStr(vFieldValue)
    
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("PolarisApplicationIdentifier", "STRING")
    strPolarisApplicationIdentifier = CStr(vFieldValue)
        
    vFieldValue = objGlobalParameterDO.GetCurrentParameterByType("PolarisTransactionType", "STRING")
    strPolarisTransactionType = CStr(vFieldValue)
        
    Set xmlPostCodeNode = vXMLRequest.getElementsByTagName("PROPERTYPOSTCODE").Item(0)
    If Not xmlPostCodeNode Is Nothing Then
        If xmlPostCodeNode.Text = "" Then
            Set xmlPostCodeNode = Nothing
        End If
    End If
    
    Set xmlTempNode = vXMLRequest.getElementsByTagName("COVERTYPE").Item(0)
    strQuotationType = xmlTempNode.Text
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Start Interface
    Set xmlTempElem = xmlOut.createElement("BCCALCINPUT")
    Set xmlTopNode = xmlResponseElem.appendChild(xmlTempElem)

    'Company Code
    Set xmlTempElem = xmlOut.createElement("COMPANYCODE")
    xmlTempElem.Text = "5"
    xmlTopNode.appendChild xmlTempElem
    
    'Product Number
    Set xmlTempElem = xmlOut.createElement("SCHEMEIDENTIFIER")
    xmlTempElem.Text = strPolarisSchemeIdentifier
    xmlTopNode.appendChild xmlTempElem
    
    'Product/Scheme Version
    Set xmlTempElem = xmlOut.createElement("SCHEMEVERSION")
    xmlTempElem.Text = strPolarisSchemeVersion
    xmlTopNode.appendChild xmlTempElem

    'Application Identifier
    Set xmlTempElem = xmlOut.createElement("APPLICATIONIDENTIFIER")
    xmlTempElem.Text = strPolarisApplicationIdentifier
    xmlTopNode.appendChild xmlTempElem

    'Application Version
    Set xmlTempElem = xmlOut.createElement("APPLICATIONVERSION")
    xmlTempElem.Text = strPolarisApplicationVersion
    xmlTopNode.appendChild xmlTempElem

    'Transaction
    Set xmlTempElem = xmlOut.createElement("TRANSACTIONMESSAGETYPE")
    xmlTempElem.Text = strPolarisTransactionType
    xmlTopNode.appendChild xmlTempElem

    'Number Of Claims in Last 5 Years
    Set xmlTempElem = xmlOut.createElement("FIVEYEARCLAIMCOUNT")
    If xmlPostCodeNode Is Nothing Then
        xmlTempElem.Text = "0"
    Else
        Set xmlTempNode = vXMLRequest.getElementsByTagName("NUMBEROFCLAIMSLAST5YEARS").Item(0)
        iNoOfClaims = CInt(xmlTempNode.Text)
        xmlTempElem.Text = CStr(iNoOfClaims)
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Quotation Date
    Set xmlTempElem = xmlOut.createElement("QUOTATIONDATE")
    xmlTempElem.Text = Format(Now, "yyyymmdd")
    xmlTopNode.appendChild xmlTempElem
    
    'Buildings Value
    Set xmlTempElem = xmlOut.createElement("BUILDINGSVALUE")
    If strQuotationType = "3" Then
        'Contents only
        xmlTempElem.Text = "0"
    Else
        xmlTempElem.Text = strBuildingsCoverAmt
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Alarm Indicator
    Set xmlTempElem = xmlOut.createElement("ALARMINDICATOR")
    If xmlPostCodeNode Is Nothing Then
        xmlTempElem.Text = "N"
    Else
        Set xmlTempNode = vXMLRequest.getElementsByTagName("ALARMFITTED").Item(0)
        'APS 07/03/00 - object variable or with block not set error as we were using
        ' xmlTempNode set to nothing and we were attempting to assign to the text property
        If Not xmlTempNode Is Nothing Then
            If xmlTempNode.Text = "1" Then
                xmlTempElem.Text = "Y"
            Else
                xmlTempElem.Text = "N"
            End If
        Else
            xmlTempElem.Text = "N"
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Wall Construction
    Set xmlTempElem = xmlOut.createElement("WALLCONSTRUCTION")
    Set xmlTempNode = vXMLRequest.getElementsByTagName("WALLCONSTRUCTION").Item(0)
    If Not xmlTempNode Is Nothing Then
        If xmlTempNode.Text = "" Or xmlTempNode.Text = "0" Then
            xmlTempElem.Text = "15"
        Else
            If xmlTempElem.getAttribute("TYPE") = "S" Then
                xmlTempElem.Text = "15"
            Else
                xmlTempElem.Text = "99"
            End If
        End If
    Else
        xmlTempElem.Text = "15"
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Roof Construction
    Set xmlTempElem = xmlOut.createElement("ROOFCONSTRUCTION")
    Set xmlTempNode = vXMLRequest.getElementsByTagName("ROOFCONSTRUCTION").Item(0)
    If Not xmlTempNode Is Nothing Then
        If xmlTempNode.Text = "" Or xmlTempNode.Text = "0" Then
            xmlTempElem.Text = "11"
        Else
            If xmlTempElem.getAttribute("TYPE") = "S" Then
                xmlTempElem.Text = "11"
            Else
                xmlTempElem.Text = "99"
            End If
        End If
    Else
        xmlTempElem.Text = "11"
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Number of Bedrooms
    Set xmlTempElem = xmlOut.createElement("BEDROOMCOUNT")
    Set xmlTempNode = vXMLRequest.getElementsByTagName("NUMBEROFBEDROOMS").Item(0)
    'APS 07/03/00 - object variable or with block not set error as we were using
    ' xmlTempNode when set to nothing and attempting to assign to the text property
    If Not xmlTempNode Is Nothing Then
        xmlTempElem.Text = xmlTempNode.Text
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Buildings Accidental Damage Required
    Set xmlTempElem = xmlOut.createElement("BUILDACCIDENTALDAMAGEREQ")
    If strQuotationType = "3" Then
        'Contents only
        xmlTempElem.Text = "N"
    Else
        Set xmlTempNode = vXMLRequest.getElementsByTagName("ACCIDENTALBUILDCOVERREQUIRED").Item(0)
        'APS 07/03/00 - object variable or with block not set error as we were using
        ' xmlTempNode when set to nothing and attempting to assign to the text property
        If Not xmlTempNode Is Nothing Then
            If xmlTempNode.Text = "" Or xmlTempNode.Text = "0" Then
                xmlTempElem.Text = "N"
            Else
                xmlTempElem.Text = "Y"
            End If
        Else
            xmlTempElem.Text = "N"
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Buildings Voluntary Excess
    Set xmlTempElem = xmlOut.createElement("BUILDVOLUNTARYEXCESS")
    If strQuotationType = "3" Then
        'Contents only
        xmlTempElem.Text = "0"
    Else
        If xmlPostCodeNode Is Nothing Then
            xmlTempElem.Text = "50"
        Else
            Set xmlTempNode = vXMLRequest.getElementsByTagName("BUILDINGSEXCESSOPTION").Item(0)
            'APS 07/03/00 - object variable or with block not set error as we were using
            ' xmlTempNode when set to nothing and attempting to assign to the text property
            If Not xmlTempNode Is Nothing Then
                If xmlTempNode.Text = "" Then
                    xmlTempElem.Text = "0"
                Else
                    xmlTempElem.Text = objComboDO.GetFirstComboValidation("BuildingExcessAmount", xmlTempNode.Text)
                End If
            Else
                xmlTempElem.Text = "0"
            End If
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'No Subsidence Risk Indicator
    Set xmlTempElem = xmlOut.createElement("NOSUBSIDENCEINDICATOR")
    If xmlPostCodeNode Is Nothing Or strQuotationType = "3" Then
        xmlTempElem.Text = "Y"
    Else
        Set xmlTempNode = vXMLRequest.getElementsByTagName("PROPERTYAFFECTED").Item(0)
        'APS 07/03/00 - object variable or with block not set error as we were using
        ' xmlTempNode when set to nothing and attempting to assign to the text property
        If Not xmlTempNode Is Nothing Then
            If xmlTempNode.Text = "1" Then
                xmlTempElem.Text = "N"
            Else
                xmlTempElem.Text = "Y"
            End If
        Else
            xmlTempElem.Text = "Y"
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Property post code
    Set xmlTempElem = xmlOut.createElement("PROPERTYPOSTCODE")
    If xmlPostCodeNode Is Nothing Then
        xmlTempElem.Text = ""
    Else
        'Dim strLeft As String
        'Dim strRight As String
        'Dim strText As String
        'strText = xmlPostCodeNode.Text
        'Dim iSpacePos As Integer
        'iSpacePos = InStr(1, strText, " ", vbTextCompare)
        'strLeft = Left(strText, iSpacePos - 1)
        'strRight = Right(strText, (Len(strText) - iSpacePos))
        'xmlTempElem.Text = strLeft + LTrim(strRight)
        xmlTempElem.Text = xmlPostCodeNode.Text
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Cover away from home
    Set xmlTempElem = xmlOut.createElement("COVERAWAYFROMHOME")
    Set xmlTempNode = vXMLRequest.getElementsByTagName("BELONGINGSFROMHOMEREQUIRED").Item(0)
    If Not xmlTempNode Is Nothing Then
        If strQuotationType = "2" Or xmlPostCodeNode Is Nothing Or _
           xmlTempNode.Text = "" Or xmlTempNode.Text = "0" Then
            xmlTempElem.Text = ""
        Else
            xmlTempElem.Text = "7500"
        End If
    Else
        xmlTempElem.Text = ""
    End If
    xmlTopNode.appendChild xmlTempElem
        
    'Specified Item Occurrence count
    Set xmlTempElem = xmlOut.createElement("SPECIALOCCURRENCECOUNT")
    Set xmlNodeList = vXMLRequest.getElementsByTagName("VALUABLESOVERLIMIT")
    'Dim xmlElement As IXMLDOMElement
    'Dim maxSeqNo As Integer
    'maxSeqNo = 0
    'For Each xmlElement In xmlNodeList
    '    Set xmlTempNode = xmlElement.getElementsByTagName("VALUABLESLIMITSEQUENCENUMBER")
    '    If CInt(xmlTempNode.Text) > maxSeqNo Then
    '        maxSeqNo = CInt(xmlTempNode.Text)
    '    End If
    'Next
    'xmlTempElem.Text = CStr(maxSeqNo)
    If xmlNodeList.length > 10 Then
        xmlTempElem.Text = "10"
    Else
        xmlTempElem.Text = CStr(xmlNodeList.length)
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Specified item value
    Dim xmlElement As IXMLDOMElement
    Dim count As Integer
    count = 0
    Set xmlTempElem = xmlOut.createElement("SPECIALITEMLIST")
    Dim xmlElemItemValue As IXMLDOMElement
    For Each xmlElement In xmlNodeList
        If count < 10 Then
            Set xmlTempNode = xmlElement.getElementsByTagName("INSURABLEVALUE").Item(0)
            Set xmlElemItemValue = xmlOut.createElement("SPECIALITEMVALUE")
            xmlElemItemValue.Text = xmlTempNode.Text
            xmlTempElem.appendChild xmlElemItemValue
            count = count + 1
        End If
    Next
    Set xmlElemItemValue = Nothing
    xmlTopNode.appendChild xmlTempElem
    
    'Contents amount required
    Set xmlTempElem = xmlOut.createElement("CONTENTSCOVERAMTRQD")
    If strQuotationType = "2" Then
        'buildings only
        xmlTempElem.Text = ""
    Else
        xmlTempElem.Text = strContentsCoverAmt
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Contents Voluntary Excess
    Set xmlTempElem = xmlOut.createElement("CONTENTSVOLUNTARYXSAMOUNT")
    If strQuotationType = "2" Or xmlPostCodeNode Is Nothing Then
        xmlTempElem.Text = "0"
    Else
        Set xmlTempNode = vXMLRequest.getElementsByTagName("CONTENTSEXCESSOPTION").Item(0)
        'APS 07/03/00 - object variable or with block not set error as we were using
        ' xmlTempNode when set to nothing and attempting to assign to the text property
        If Not xmlTempNode Is Nothing Then
            If xmlTempNode.Text = "" Then
                xmlTempElem.Text = "0"
            Else
                xmlTempElem.Text = objComboDO.GetFirstComboValidation("ContentsExcessAmount", xmlTempNode.Text)
            End If
        Else
            xmlTempElem.Text = "0"
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'contents accidental damage
    Set xmlTempElem = xmlOut.createElement("CONTENTSACCIDENTALDAMAGEINDICATOR")
    If strQuotationType = "2" Then
        xmlTempElem.Text = "N"
    Else
        If xmlPostCodeNode Is Nothing Then
            xmlTempElem.Text = "Y"
        Else
            Set xmlTempNode = vXMLRequest.getElementsByTagName("ACCIDENTALCONTENTCOVERREQUIRED").Item(0)
            'APS 07/03/00 - object variable or with block not set error as we were using
            ' xmlTempNode when set to nothing and attempting to assign to the text property
            If Not xmlTempNode Is Nothing Then
                If xmlTempNode.Text = "" Then
                    xmlTempElem.Text = "N"
                Else
                    xmlTempElem.Text = "Y"
                End If
            Else
                xmlTempElem.Text = "N"
            End If
        End If
    End If
    xmlTopNode.appendChild xmlTempElem
    
    'Legal Expenses
    Set xmlTempElem = xmlOut.createElement("LEGALEXPENSEINDICATOR")
    xmlTempElem.Text = "Y"
    xmlTopNode.appendChild xmlTempElem
    
    'Proposer repeat count
    Set xmlTempNode = vXMLRequest.getElementsByTagName("CUSTOMERNUMBER2").Item(0)
    Set xmlTempElem = xmlOut.createElement("CLIENTDATACOUNT")
    If xmlTempNode Is Nothing Then
        xmlTempElem.Text = "1"
    Else
        If xmlTempNode.Text = "" Then
            xmlTempElem.Text = "1"
        Else
            xmlTempElem.Text = "2"
        End If
    End If
    Dim iProposerRepeat As Integer
    iProposerRepeat = CInt(xmlTempElem.Text)
    xmlTopNode.appendChild xmlTempElem
        
    Dim xmlTempGroupListElem As IXMLDOMElement
    Dim xmlTempClientDataElem As IXMLDOMElement
    Dim strDate As String
    Dim nYear As Long
    Dim nMonth As Long
    Dim nDay As Long
    Dim i As Integer
    For i = 1 To iProposerRepeat
        If i = 1 Then
            Set xmlTempGroupListElem = xmlOut.createElement("CLIENTDATAGROUPLIST")
        End If
        Set xmlTempClientDataElem = xmlOut.createElement("CLIENTDATA")
        'Proposer date of birth
        Set xmlTempNode = vXMLRequest.getElementsByTagName("DATEOFBIRTH").Item(0)
        Set xmlTempElem = xmlOut.createElement("CLIENTDATEOFBIRTH")
        strDate = xmlTempNode.Text
        nYear = Year(CDate(strDate))
        nMonth = Month(CDate(strDate))
        nDay = Day(CDate(strDate))
        strDate = CStr(nYear)
        If nMonth < 10 Then
            strDate = strDate + "0"
        End If
        strDate = strDate + CStr(nMonth)
        If nDay < 10 Then
            strDate = strDate + "0"
        End If
        strDate = strDate + CStr(nDay)
        xmlTempElem.Text = strDate
        xmlTempClientDataElem.appendChild xmlTempElem
    
        'Refused household insurance
        Set xmlTempNode = vXMLRequest.getElementsByTagName("DECLARATIONVALUE").Item(0)
        Set xmlTempElem = xmlOut.createElement("CLIENTDECLINEINDICATOR")
        If xmlPostCodeNode Is Nothing Or xmlTempNode Is Nothing Then
            xmlTempElem.Text = "N"
        Else
            If xmlTempNode.Text = "1" Then
                xmlTempElem.Text = "Y"
            Else
                xmlTempElem.Text = "N"
            End If
        End If
        xmlTempClientDataElem.appendChild xmlTempElem
            
        'Dishonesty convictions or bankruptcy
        Set xmlTempElem = xmlOut.createElement("CLIENTCONVICTEDINDICATOR")
        If xmlPostCodeNode Is Nothing Then
            xmlTempElem.Text = "N"
        Else
            Dim strBankruptcyApp As String
            strBankruptcyApp = "BANKRUPTCYAPPLICANT" & CStr(i)
            Set xmlTempNode = vXMLRequest.getElementsByTagName(strBankruptcyApp).Item(0)
            If xmlTempNode.Text = "1" Then
                xmlTempElem.Text = "Y"
            Else
                xmlTempElem.Text = "N"
            End If
        End If
        xmlTempClientDataElem.appendChild xmlTempElem
        
        'we need a CLIENTOCCUPATIONCODE even though we don't use it
        Set xmlTempElem = xmlOut.createElement("CLIENTOCCUPATIONCODE")
        xmlTempClientDataElem.appendChild xmlTempElem
    
        xmlTempGroupListElem.appendChild xmlTempClientDataElem
    Next
    xmlTopNode.appendChild xmlTempGroupListElem
    
    Set IBuildingsAndContentsSubQuoteBO_BuildBCCalcInterface = xmlOut.documentElement
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
BuildInterfaceExit:

    Set objErrAssist = Nothing
    Set xmlTempElem = Nothing
    Set xmlTopNode = Nothing
    Set xmlOut = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlTempClientDataElem = Nothing
    Set xmlTempGroupListElem = Nothing
    Set objComboDO = Nothing
    
    Exit Function

BuildInterfaceVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTopNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_BuildBCCalcInterface = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume BuildInterfaceExit
    

End Function

Private Function IBuildingsAndContentsSubQuoteBO_CreateFirstSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_CreateFirstSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
        
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.CreateFirstSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    Set IBuildingsAndContentsSubQuoteBO_CreateFirstSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateFirstVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_CreateFirstSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_CreateNewSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_CreateNewSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.CreateNewSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    Set IBuildingsAndContentsSubQuoteBO_CreateNewSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_CreateNewSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewExit


End Function

Private Function IBuildingsAndContentsSubQuoteBO_CreateSubQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_CreateSubQuoteDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.CreateSubQuoteDetails(vXMLRequest)
   
    Set IBuildingsAndContentsSubQuoteBO_CreateSubQuoteDetails = xmlTempResponseNode.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_CreateSubQuoteDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_CreateValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_CreateValuablesOverLimit"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.CreateValuablesOverLimit(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IBuildingsAndContentsSubQuoteBO_CreateValuablesOverLimit = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_CreateValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_DeleteAllValuables(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete all instances of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllValuablesVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_DeleteAllValuables"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
       
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    
    '------------------------------------------------------------------------------------------
    ' validate parameters
    '------------------------------------------------------------------------------------------
    Dim strTagValue, strTagName As String
           
    strTagName = "APPLICATIONNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If
        
    '------------------------------------------------------------------------------------------
    ' call Business Transaction Object DeleteAllValuables function
    '------------------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.DeleteAllValuables(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IBuildingsAndContentsSubQuoteBO_DeleteAllValuables = xmlResponseElem.cloneNode(True)
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteAllValuablesExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

DeleteAllValuablesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_DeleteAllValuables = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteAllValuablesExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_DeleteValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_DeleteValuablesOverLimit"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.DeleteValuablesOverLimit(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IBuildingsAndContentsSubQuoteBO_DeleteValuablesOverLimit = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_DeleteValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_DeleteAll(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete all instances of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_DeleteAll"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
' TODO ************************************************************************
'   Remove oeNotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
    
' TODO ************************************************************************
'   replace BuildingsAndContentsSubQuoteTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    
    '------------------------------------------------------------------------------------------
    ' validate parameters
    '------------------------------------------------------------------------------------------
    
    Dim strTagValue, strTagName As String
        
' TODO ************************************************************************
'   replace example tag names with actual tag names required
' *****************************************************************************
    
    strTagName = "APPLICATIONNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If
        
    '------------------------------------------------------------------------------------------
    ' call Business Transaction Object DeleteAll function
    '------------------------------------------------------------------------------------------
    
    objIBuildingsAndContentsSubQuoteTxBO.DeleteAll vXMLRequest
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IBuildingsAndContentsSubQuoteBO_DeleteAll = xmlResponseElem.cloneNode(True)
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteAllExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

DeleteAllVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_DeleteAll = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteAllExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_FindList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_FindList"


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
        
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILS" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BCSUBQUOTEDETAILS tag not found"
    End If
        
    Set xmlDataNode = objIBuildingsAndContentsSubQuoteDO.FindList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IBuildingsAndContentsSubQuoteBO_FindList = xmlResponseElem.cloneNode(True)

FindListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_FindList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_FindValuablesOverLimitList( _
                    ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_FindValuablesOverLimitList"


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
        
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "VALUABLESOVERLIMIT" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("VALUABLESOVERLIMIT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "VALUABLESOVERLIMIT tag not found"
    End If
        
    Set xmlDataNode = _
        objIBuildingsAndContentsSubQuoteDO.FindValuablesOverLimitList(xmlRequestNode)
    If Not xmlDataNode Is Nothing Then  '** THIS IS TEMPORARY STUFF - TILL CORRECTING ABOVE DO METHOD
        xmlResponseElem.appendChild xmlDataNode
    End If
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IBuildingsAndContentsSubQuoteBO_FindValuablesOverLimitList = xmlResponseElem

FindListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_FindValuablesOverLimitList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit
    
End Function

Private Function IBuildingsAndContentsSubQuoteBO_GetBCNotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetBCNotesVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_GetBCNotes"


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
    
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "BUILDINGSANDCONTENTSNOTES" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("BUILDINGSANDCONTENTSNOTES").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BUILDINGSANDCONTENTSNOTES tag not found"
    End If
        
    Set xmlDataNode = objIBCSubQuoteDO.GetBCNotes(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IBuildingsAndContentsSubQuoteBO_GetBCNotes = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetBCNotesExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetBCNotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_GetBCNotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBCNotesExit


End Function

'*****************************************************************************
'   DATE    :   16/06/2004
'   CC      :   BMIDS763
'*****************************************************************************
Private Sub IBuildingsAndContentsSubQuoteBO_GetInsuranceAdminFee(ByVal vXMLRequest As IXMLDOMElement, ByVal vXMLResponse As IXMLDOMElement)
        
    Dim xmlBCTypeNode       As IXMLDOMNode
    Dim objBCSubQuoteDO     As IBuildingsAndContentsSubQuoteDO
    Dim objContext          As ObjectContext
    
    On Error GoTo ErrorInsuranceAdminFee
    
    Set objContext = GetObjectContext()
    
    If Not objContext Is Nothing Then
        Set objBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    Else
        Set objBCSubQuoteDO = CreateObject(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    Set xmlBCTypeNode = vXMLRequest.selectSingleNode("GETINSURANCEADMINFEE")
    
    If Not xmlBCTypeNode Is Nothing Then
        objBCSubQuoteDO.GetInsuranceAdminFee xmlBCTypeNode, vXMLResponse
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ExitInsuranceAdminFee:
    
    Set objBCSubQuoteDO = Nothing
    Set xmlBCTypeNode = Nothing
    Exit Sub
    
ErrorInsuranceAdminFee:
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    If Err.number <> 0 Then
        Err.Raise Err.number, Err.Source, Err.Description
    Else
        Resume ExitInsuranceAdminFee
    End If
End Sub

Private Function IBuildingsAndContentsSubQuoteBO_GetPaymentFrequency(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPaymentFreqVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_GetPaymentFrequency"


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
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "BUILDINGSANDCONTENTSDETAILS" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("BUILDINGSANDCONTENTSDETAILS").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BUILDINGSANDCONTENTSDETAILS tag not found"
    End If
        
    Set xmlDataNode = objIBCSubQuoteDO.GetPaymentFrequency(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IBuildingsAndContentsSubQuoteBO_GetPaymentFrequency = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetPaymentFreqVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set IBuildingsAndContentsSubQuoteBO_GetPaymentFrequency = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IBuildingsAndContentsSubQuoteBO_GetSubQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetSubQuoteDetailsVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_GetSubQuoteDetails"


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
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILSNOTES" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILSNOTES").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BCSUBQUOTEDETAILSNOTES tag not found"
    End If
        
    Set xmlDataNode = objIBCSubQuoteDO.GetSubQuoteDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
    
    objXmlAssist.ChangeNodeName xmlRequestNode, "BCSUBQUOTEDETAILSNOTES", "VALUABLESOVERLIMIT"
    
    'APS 05/05/2000 - It is ok not to have any valuables
    Dim iSaveErr As Long
    On Error Resume Next
    Set xmlDataNode = objIBCSubQuoteDO.FindValuablesOverLimitList(xmlRequestNode)
    iSaveErr = Err.number
    On Error GoTo GetSubQuoteDetailsVbErr
    If objErrAssist.GetOmigaErrorNumber(iSaveErr) <> oeRecordNotfound Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, iSaveErr
    End If
    
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    
    ' change the request back
    objXmlAssist.ChangeNodeName xmlRequestNode, "VALUABLESOVERLIMIT", "BCSUBQUOTEDETAILSNOTES"
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IBuildingsAndContentsSubQuoteBO_GetSubQuoteDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetSubQuoteDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set IBuildingsAndContentsSubQuoteBO_GetSubQuoteDetails = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IBuildingsAndContentsSubQuoteBO_SaveBCNotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveBCNotesVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_SaveBCNotes"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIClassDef As IomCMClassDef
    Dim botOperation As BOOPERATIONTYPE
    Dim xmlResponseElem As IXMLDOMElement
    Dim strAttrVal As String
    Dim strDescription As String
    

    Set objIClassDef = New omCMClassDef

    'botOperation = objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadBCNotesData)
    
    Dim objIBCSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    If objContext Is Nothing Then
        Set objIBCSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBCSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
        
    ' try to 'get' it first to determine operation
    Dim blnCreate As Boolean
    blnCreate = False
    Set xmlResponseElem = IBuildingsAndContentsSubQuoteBO_GetBCNotes(vXMLRequest)
    strAttrVal = objXmlAssist.GetAttributeValue(xmlResponseElem, "RESPONSE", "TYPE")
    If strAttrVal = "APPERR" Then
        strDescription = objXmlAssist.GetTagValue(xmlResponseElem, "DESCRIPTION")
        If strDescription = "Record not found" Then
            blnCreate = True
        End If
    End If
    
    If blnCreate = True Then
        Set IBuildingsAndContentsSubQuoteBO_SaveBCNotes = objIBCSubQuoteTxBO.CreateBCNotes(vXMLRequest)
    Else
        Set IBuildingsAndContentsSubQuoteBO_SaveBCNotes = objIBCSubQuoteTxBO.UpdateBCNotes(vXMLRequest)
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIClassDef = Nothing
    
    Exit Function

SaveBCNotesVbErr:
    Set objIClassDef = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set IBuildingsAndContentsSubQuoteBO_SaveBCNotes = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IBuildingsAndContentsSubQuoteBO_SaveSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveSubQuoteVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_SaveSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIBCSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO

    Dim xmlTempNode As IXMLDOMNode
    Dim xmlBCCalcNode As IXMLDOMNode
    Dim xmlTempBCCalcNode As IXMLDOMNode
    Dim xmlBCDetailsNode As IXMLDOMNode
    Dim xmlBCSubQuoteNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    
    If objContext Is Nothing Then
        Set objIBCSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO

    Else
        Set objIBCSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    
    Set xmlBCCalcNode = vXMLRequest.selectSingleNode(".//BCCALCOUTPUT")
    Set xmlBCDetailsNode = vXMLRequest.selectSingleNode(".//BCSUBQUOTEDETAILSNOTES/BUILDINGSANDCONTENTSDETAILS")
    Set xmlBCSubQuoteNode = vXMLRequest.selectSingleNode(".//BCSUBQUOTEDETAILSNOTES/BUILDINGSANDCONTENTSSUBQUOTE")
    
    ' Update the return text from the BCCALC output
    Set xmlTempBCCalcNode = xmlBCCalcNode.selectSingleNode("BCRETURNTEXT")
    Set xmlTempNode = xmlBCDetailsNode.selectSingleNode("BCRETURNTEXT")
    xmlTempNode.Text = xmlTempBCCalcNode.Text
    
    ' Update the status from the BCCALC output
    Set xmlTempBCCalcNode = xmlBCCalcNode.selectSingleNode("STATUS")
    Set xmlTempNode = xmlBCDetailsNode.selectSingleNode("STATUS")
    xmlTempNode.Text = xmlTempBCCalcNode.Text
    
    'APS 29/02/00 - Change CDbl function to ConvertAssist.CSafeBool function
    Dim objConvertAssist As New ConvertAssist
    

    Set objConvertAssist = Nothing
    
    Dim xmlNode As IXMLDOMNode
    
    ' APS 05/04/2000 - First delete all existing valuables for this application
    Set xmlNode = objXmlAssist.GetNode(vXMLRequest, "//VALUABLESOVERLIMITLIST/VALUABLESOVERLIMIT", False)
    If Not (xmlNode Is Nothing) Then
        Set xmlNode = objIBCSubQuoteTxBO.DeleteAllValuables(vXMLRequest)
        objErrAssist.CheckResponse xmlNode, IBuildingsAndContentsSubQuoteBO_SaveSubQuote
    
        ' APS 05/04/2000 - Second create all valuables as new valuables why? - much quicker than
        ' looping through each valuable deciding if we have to create or update the record
        Set xmlNode = objIBCSubQuoteTxBO.CreateValuablesOverLimit(vXMLRequest)
        objErrAssist.CheckResponse xmlNode, IBuildingsAndContentsSubQuoteBO_SaveSubQuote
    End If
    ' APS 05/04/2000 - Finally save the underlying BC Quote information
    Set IBuildingsAndContentsSubQuoteBO_SaveSubQuote = objIBCSubQuoteTxBO.UpdateSubQuoteDetails(vXMLRequest)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objIBCSubQuoteTxBO = Nothing
    Set xmlTempNode = Nothing
    Set xmlBCCalcNode = Nothing
    Set xmlTempBCCalcNode = Nothing
    Set xmlBCSubQuoteNode = Nothing
    Set xmlBCDetailsNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Exit Function

SaveSubQuoteVbErr:
    Set objXmlAssist = Nothing

    Set objIBCSubQuoteTxBO = Nothing
    Set xmlTempNode = Nothing
    Set xmlBCCalcNode = Nothing
    Set xmlTempBCCalcNode = Nothing
    Set xmlBCSubQuoteNode = Nothing
    Set xmlBCDetailsNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set IBuildingsAndContentsSubQuoteBO_SaveSubQuote = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.number, Err.Source, Err.Description
End Function


Private Function IBuildingsAndContentsSubQuoteBO_UpdateSubQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_UpdateSubQuoteDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.UpdateSubQuoteDetails(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IBuildingsAndContentsSubQuoteBO_UpdateSubQuoteDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_UpdateSubQuoteDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_UpdateValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_UpdateValuablesOverLimit"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIBuildingsAndContentsSubQuoteTxBO As IBuildingsAndContentsSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteTxBO = New BuildingsAndContentsSubQuoteTxBO
    Else
        Set objIBuildingsAndContentsSubQuoteTxBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteTxBO.UpdateValuablesOverLimit(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IBuildingsAndContentsSubQuoteBO_UpdateValuablesOverLimit = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:

    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_UpdateValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_Validate(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_Validate"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

' TODO ************************************************************************
'   replace BuildingsAndContentsSubQuoteDO reference with actual DO required
' *****************************************************************************
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
        
    ' Get node to add derived data to and apply validation to
    If vXMLRequest.nodeName = "UPDATE" Or vXMLRequest.nodeName = "CREATE" Then
        Set xmlRequestTypeNode = vXMLRequest
    ElseIf vXMLRequest.getElementsByTagName("UPDATE").length > 0 Then
        Set xmlRequestTypeNode = vXMLRequest.getElementsByTagName("UPDATE").Item(0)
    ElseIf vXMLRequest.getElementsByTagName("CREATE").length > 0 Then
        Set xmlRequestTypeNode = vXMLRequest.getElementsByTagName("CREATE").Item(0)
    End If
        
    If Not xmlRequestTypeNode Is Nothing Then
        Set xmlDataNode = xmlRequestTypeNode.firstChild

        ' Add in derived data
        If Not xmlDataNode Is Nothing Then
            objIBuildingsAndContentsSubQuoteDO.AddDerivedData xmlDataNode
        End If
    End If
    
' TODO ************************************************************************
'   add hard coded validation here, & raise error if it fails
' *****************************************************************************

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IBuildingsAndContentsSubQuoteBO_Validate = xmlResponseElem.cloneNode(True)
    
ValidateExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_Validate = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Private Function IBuildingsAndContentsSubQuoteBO_ValidateSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
    
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteBO_ValidateSubQuote"


    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlElement = xmlOut.createElement("BCSUBQUOTE")
    Set xmlResponseNode = xmlResponseElem.appendChild(xmlElement)
    Set xmlElement = xmlOut.createElement("VALIDBCSUBQUOTE")
    xmlElement.Text = 1
    xmlResponseNode.appendChild xmlElement
    
    Set IBuildingsAndContentsSubQuoteBO_ValidateSubQuote = xmlResponseElem.cloneNode(True)

 
ValidateExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set objErrAssist = Nothing
    
    Exit Function

ValidateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteBO_ValidateSubQuote = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ValidateExit
    

End Function


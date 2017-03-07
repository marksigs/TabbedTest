VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ApplicationThirdPartyTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'ApplicationThirdPartyTxBO' with actual class name
' *****************************************************************************
'
'Workfile:      ApplicationThirdPartyTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for ApplicationThirdParty Business Object which requires transactions
'
'Dependencies:  List any other dependent components
'               e.g. ApplicationThirdPartyDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     11/01/00    Created
'------------------------------------------------------------------------------------------

Option Explicit

' TODO ************************************************************************
'   Add interface definition class and amend this to implement it
' *****************************************************************************
Implements IApplicationThirdPartyTxBO
 
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     31/07/2002  BMIDS00075 - IApplicationThirdPartyTxBO_SaveLegalRep()
'------------------------------------------------------------------------------------------

Private Function IApplicationThirdPartyTxBO_CreateArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateArchitect multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_CreateArchitect"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONARCHITECT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONARCHITECT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONARCHITECT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.CreateArchitect xmlElement
    Else
        ' For each relevant node call the data object CreateArchitect
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.CreateArchitect xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_CreateArchitect = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateArchitectExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_CreateArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArchitectExit
End Function

Private Function IApplicationThirdPartyTxBO_CreateBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankBuildingSociety multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_CreateBankBuildingSociety"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBANKBUILDINGSOC" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBANKBUILDINGSOC")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBANKBUILDINGSOC" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.CreateBankBuildingSociety xmlElement
    Else
        ' For each relevant node call the data object CreateBankBuildingSociety
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.CreateBankBuildingSociety xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_CreateBankBuildingSociety = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankBuildingSocietyExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_CreateBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyTxBO_CreateBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBuilder multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_CreateBuilder"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBUILDER" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBUILDER")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBUILDER" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.CreateBuilder xmlElement
    Else
        ' For each relevant node call the data object CreateBuilder
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.CreateBuilder xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_CreateBuilder = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBuilderExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_CreateBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBuilderExit
End Function

Private Function IApplicationThirdPartyTxBO_CreateEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateEstateAgent multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_CreateEstateAgent"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONESTATEAGENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONESTATEAGENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONESTATEAGENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.CreateEstateAgent xmlElement
    Else
        ' For each relevant node call the data object CreateEstateAgent
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.CreateEstateAgent xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_CreateEstateAgent = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateEstateAgentExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_CreateEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEstateAgentExit
End Function

Private Function IApplicationThirdPartyTxBO_CreateLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLegalRep multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_CreateLegalRep"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONLEGALREP" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONLEGALREP")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONLEGALREP" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.CreateLegalRep xmlElement
    Else
        ' For each relevant node call the data object CreateLegalRep
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.CreateLegalRep xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_CreateLegalRep = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLegalRepExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_CreateLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLegalRepExit
End Function

Private Function IApplicationThirdPartyTxBO_DeleteArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteArchitect multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_DeleteArchitect"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONARCHITECT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONARCHITECT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONARCHITECT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.DeleteArchitect xmlElement
    Else
        ' For each relevant node call the data object DeleteArchitect
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.DeleteArchitect xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_DeleteArchitect = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteArchitectExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationThirdPartyTxBO_DeleteArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArchitectExit
End Function

Private Function IApplicationThirdPartyTxBO_DeleteBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankBuildingSociety multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_DeleteBankBuildingSociety"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBANKBUILDINGSOC" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBANKBUILDINGSOC")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBANKBUILDINGSOC" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.DeleteBankBuildingSociety xmlElement
    Else
        ' For each relevant node call the data object DeleteBankBuildingSociety
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.DeleteBankBuildingSociety xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_DeleteBankBuildingSociety = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteBankBuildingSocietyExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationThirdPartyTxBO_DeleteBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyTxBO_DeleteBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBuilder multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_DeleteBuilder"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBUILDER" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBUILDER")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBUILDER" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.DeleteBuilder xmlElement
    Else
        ' For each relevant node call the data object DeleteBuilder
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.DeleteBuilder xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_DeleteBuilder = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteBuilderExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationThirdPartyTxBO_DeleteBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBuilderExit
End Function

Private Function IApplicationThirdPartyTxBO_DeleteEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteEstateAgent multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_DeleteEstateAgent"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONESTATEAGENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONESTATEAGENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONESTATEAGENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.DeleteEstateAgent xmlElement
    Else
        ' For each relevant node call the data object DeleteEstateAgent
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.DeleteEstateAgent xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_DeleteEstateAgent = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteEstateAgentExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationThirdPartyTxBO_DeleteEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEstateAgentExit
End Function

Private Function IApplicationThirdPartyTxBO_DeleteLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLegalRep multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_DeleteLegalRep"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONLEGALREP" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONLEGALREP")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONLEGALREP" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.DeleteLegalRep xmlElement
    Else
        ' For each relevant node call the data object DeleteLegalRep
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.DeleteLegalRep xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_DeleteLegalRep = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteLegalRepExit:

    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set IApplicationThirdPartyTxBO_DeleteLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLegalRepExit
End Function

Private Function IApplicationThirdPartyTxBO_SaveLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo SaveLegalRepVbErr
    
    Const strFunctionName As String = "IApplicationThirdPartyTxBO_SaveLegalRep"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlTempRequest  As IXMLDOMElement
    Dim xmlPayeeHistoryNode   As IXMLDOMNode
    Dim xmlPayeeHistory As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist
    Dim objIApplicationThirdPartyBO As IApplicationThirdPartyBO
    Dim objPaymentProcessingBO  As Object
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As New SQLAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim strCondition As String
    Dim strApplNumber As String
    Dim strAFFNumber As String
    Dim strResponse  As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyBO = New ApplicationThirdPartyBO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIApplicationThirdPartyBO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyBO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
'Stage1:
    
    ' Check whether a corresponding record already exist in table 'ApplicationlegalRep', if so update it
    ' else create one
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    strCondition = " APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplNumber) & " AND " _
                   & " APPLICATIONFACTFINDNUMBER =" & strAFFNumber
                   
    If objIADOAssist.CheckRecordExists("APPLICATIONLEGALREP", strCondition) Then
        Set xmlTempResponse = objIApplicationThirdPartyBO.UpdateLegalRep(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    Else
        Set xmlTempResponse = objIApplicationThirdPartyBO.CreateLegalRep(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    End If
    
'Stage2:

    'Create PayeeHistoryDetails Record
    
    Set xmlPayeeHistoryNode = vxmlRequest.selectSingleNode(".//PAYEEHISTORY")
    
    If Not xmlPayeeHistoryNode Is Nothing Then
        
        'Create Request Node
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequest)
        xmlTempRequest.setAttribute "OPERATION", "CREATEPAYEEHISTORYDETAILS"
    
        xmlTempRequest.appendChild xmlPayeeHistoryNode
        
        'COnvert into Attribute Base for PhaseII component
        Set xmlPayeeHistory = xmlCreateAttributeBasedResponse(xmlTempRequest, True)
    
        'Create an instance of PaymentProcessing and call Request method
        Set objPaymentProcessingBO = objContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingBO")
        
        strResponse = objPaymentProcessingBO.omPayProcRequest(xmlPayeeHistory.xml)
        objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
    End If
    
'Exit:

    Set IApplicationThirdPartyTxBO_SaveLegalRep = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveLegalRepExit:

    Set objErrAssist = Nothing
    Set objIApplicationThirdPartyBO = Nothing
    Set objPaymentProcessingBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPayeeHistoryNode = Nothing
    Set xmlPayeeHistory = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    
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

    Set IApplicationThirdPartyTxBO_SaveLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLegalRepExit

End Function

Private Function IApplicationThirdPartyTxBO_UpdateArchitect(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateArchitect multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_UpdateArchitect"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONARCHITECT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONARCHITECT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONARCHITECT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.UpdateArchitect xmlElement
    Else
        ' For each relevant node call the data object UpdateArchitect
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.UpdateArchitect xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_UpdateArchitect = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateArchitectExit:
    
    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_UpdateArchitect = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArchitectExit
End Function

Private Function IApplicationThirdPartyTxBO_UpdateBankBuildingSociety(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankBuildingSociety multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_UpdateBankBuildingSociety"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBANKBUILDINGSOC" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBANKBUILDINGSOC")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBANKBUILDINGSOC" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.UpdateBankBuildingSociety xmlElement
    Else
        ' For each relevant node call the data object UpdateBankBuildingSociety
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.UpdateBankBuildingSociety xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_UpdateBankBuildingSociety = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankBuildingSocietyExit:
    
    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_UpdateBankBuildingSociety = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankBuildingSocietyExit
End Function

Private Function IApplicationThirdPartyTxBO_UpdateBuilder(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBuilder multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_UpdateBuilder"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONBUILDER" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONBUILDER")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONBUILDER" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.UpdateBuilder xmlElement
    Else
        ' For each relevant node call the data object UpdateBuilder
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.UpdateBuilder xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_UpdateBuilder = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBuilderExit:
    
    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_UpdateBuilder = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBuilderExit
End Function

Private Function IApplicationThirdPartyTxBO_UpdateEstateAgent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateEstateAgent multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_UpdateEstateAgent"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONESTATEAGENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONESTATEAGENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONESTATEAGENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.UpdateEstateAgent xmlElement
    Else
        ' For each relevant node call the data object UpdateEstateAgent
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.UpdateEstateAgent xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_UpdateEstateAgent = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEstateAgentExit:
    
    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_UpdateEstateAgent = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEstateAgentExit
End Function

Private Function IApplicationThirdPartyTxBO_UpdateLegalRep(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLegalRep multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyTxBO_UpdateLegalRep"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationThirdPartyDO As IApplicationThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATIONLEGALREP" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATIONLEGALREP")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATIONLEGALREP" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationThirdPartyDO = New ApplicationThirdPartyDO
    Else
        Set objIApplicationThirdPartyDO = objContext.CreateInstance(App.Title & ".ApplicationThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationThirdPartyDO.UpdateLegalRep xmlElement
    Else
        ' For each relevant node call the data object UpdateLegalRep
        For Each xmlElement In xmlNodeList
            objIApplicationThirdPartyDO.UpdateLegalRep xmlElement
        Next
    End If
    
    Set IApplicationThirdPartyTxBO_UpdateLegalRep = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateLegalRepExit:
    
    Set objIApplicationThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IApplicationThirdPartyTxBO_UpdateLegalRep = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLegalRepExit
End Function


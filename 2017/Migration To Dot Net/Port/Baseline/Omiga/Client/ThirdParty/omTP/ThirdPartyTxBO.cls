VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ThirdPartyTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'ThirdPartyTxBO' with actual class name
' *****************************************************************************

'Workfile:      ThirdPartyTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for ThirdParty Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  List any other dependent components
'               e.g. ThirdPartyDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'PSC    01/10/99    Added new error processing
'------------------------------------------------------------------------------------------

Option Explicit

Implements IThirdPartyTxBO

 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================


Private Function IThirdPartyTxBO_CreateDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IThirdPartyTxBO_CreateDirectory"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim objIThirdPartyDO As IThirdPartyDO
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim xmlKeysNode As IXMLDOMNode

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
               
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "NAMEANDADDRESSDIRECTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("NAMEANDADDRESSDIRECTORY")
    End If
    
    If xmlElement Is Nothing And xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "NAMEANDADDRESSDIRECTORY tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If

    ' Creat a 'GENERATEDKEYS' child node to the main response XML
    Set xmlKeysNode = xmlResponseElem.appendChild(xmlResponseElem.ownerDocument.createElement("GENERATEDKEYS"))

    If Not xmlElement Is Nothing Then
        objIThirdPartyDO.CreateDirectory xmlElement, xmlKeysNode
    Else
        ' For each relevant node call the data object CreateDirectory
        For Each xmlElement In xmlNodeList
            objIThirdPartyDO.CreateDirectory xmlElement, xmlKeysNode
        Next xmlElement
    End If

    Set IThirdPartyTxBO_CreateDirectory = xmlResponseElem.cloneNode(True)
     
    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlKeysNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
CreateExit:
    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlKeysNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
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

    Set IThirdPartyTxBO_CreateDirectory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IThirdPartyTxBO_CreateMortgageLenderDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLenderDirectory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLenderDirectoryVbErr

    Const strFunctionName As String = "IThirdPartyTxBO_CreateMortgageLenderDirectory"

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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGELENDERDIRECTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGELENDERDIRECTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGELENDERDIRECTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIThirdPartyDO.CreateMortgageLenderDirectory xmlElement
    Else
        ' For each relevant node call the data object CreateMortgageLenderDirectory
        For Each xmlElement In xmlNodeList
            objIThirdPartyDO.CreateMortgageLenderDirectory xmlElement
        Next
    End If
    
    Set IThirdPartyTxBO_CreateMortgageLenderDirectory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateMortgageLenderDirectoryExit:

    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateMortgageLenderDirectoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyTxBO_CreateMortgageLenderDirectory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageLenderDirectoryExit
End Function

Private Function IThirdPartyTxBO_CreateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateThirdParty multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyTxBO_CreateThirdParty"

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
    Dim objIThirdPartyDO As IThirdPartyDO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "THIRDPARTY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("THIRDPARTY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "THIRDPARTY tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If

    ' Creat a 'GENERATEDKEYS' child node to the main response XML
    Set xmlKeysNode = xmlResponseElem.appendChild(xmlResponseElem.ownerDocument.createElement("GENERATEDKEYS"))
    
    If Not xmlElement Is Nothing Then
        objIThirdPartyDO.CreateThirdParty xmlElement, xmlKeysNode
    Else
        ' For each relevant node call the data object CreateThirdParty
        For Each xmlElement In xmlNodeList
            objIThirdPartyDO.CreateThirdParty xmlElement, xmlKeysNode
        Next
    End If
    
    Set IThirdPartyTxBO_CreateThirdParty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateThirdPartyExit:

    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateThirdPartyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyTxBO_CreateThirdParty = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateThirdPartyExit

End Function

Private Function IThirdPartyTxBO_DeleteThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteThirdParty multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyTxBO_DeleteThirdParty"

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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "THIRDPARTY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("THIRDPARTY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "THIRDPARTY tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIThirdPartyDO.DeleteThirdParty xmlElement
    Else
        ' For each relevant node call the data object DeleteThirdParty
        For Each xmlElement In xmlNodeList
            objIThirdPartyDO.DeleteThirdParty xmlElement
        Next
    End If
    
    Set IThirdPartyTxBO_DeleteThirdParty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteThirdPartyExit:

    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteThirdPartyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyTxBO_DeleteThirdParty = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteThirdPartyExit

End Function

Private Function IThirdPartyTxBO_UpdateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateThirdParty multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyTxBO_UpdateThirdParty"

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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "THIRDPARTY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("THIRDPARTY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "THIRDPARTY tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIThirdPartyDO.UpdateThirdParty xmlElement
    Else
        ' For each relevant node call the data object UpdateThirdParty
        For Each xmlElement In xmlNodeList
            objIThirdPartyDO.UpdateThirdParty xmlElement
        Next
    End If
    
    Set IThirdPartyTxBO_UpdateThirdParty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateThirdPartyExit:
    
    Set objIThirdPartyDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateThirdPartyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyTxBO_UpdateThirdParty = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateThirdPartyExit

End Function



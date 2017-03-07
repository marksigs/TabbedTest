VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AttitudeToBorrowingTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AttitudeToBorrowingTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  AttitudeToBorrowingDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     01/08/99    Created
'APS    22/05/00    SYS0779 Removed Customer Registration component
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "ATTITUDETOBORROWING"
 
 
Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create multiple instances of the persistant data associated with this
'   business object
'   invokes appropriate Data Object(s) to perform persistance
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlUpdateNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objComboDO As ComboDO
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
    Dim objApplicationBO As ApplicationBO
    Dim strRequest, strResponse As String

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    Dim intNumNodes As Integer
    
    intNumNodes = xmlNodeList.length
    
    If intNumNodes = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
   
    ' APS 22/05/00 SYS0779 Removed Customer Registration component
    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
        Set objComboDO = New ComboDO
        Set objApplicationBO = New ApplicationBO
    Else
        Set objAttitudeToBorrowingDO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
        Set objComboDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objApplicationBO = _
            objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    Dim intNodeIndex As Integer
    
    '------------------------------------------------------------------------------------------
    ' update ApplicationFactFind with the total score
    '------------------------------------------------------------------------------------------
    
    Dim lngTotScore As Long
    Dim strQnResponse, strQnNumber, strValidationType  As String
    Dim strGroupName As String
    
    For intNodeIndex = 0 To intNumNodes - 1
    
        strQnResponse = objXmlAssist.GetTagValue( _
            xmlNodeList.Item(intNodeIndex), "RESPONSETOQUESTION")
        strQnNumber = objXmlAssist.GetTagValue( _
            xmlNodeList.Item(intNodeIndex), "QUESTIONNUMBER")
        strGroupName = "AttitudeToBorrowing_" & strQnNumber
            
        strRequest = "<LIST>"
        strRequest = strRequest & "<GROUPNAME>" & strGroupName & "</GROUPNAME>"
        strRequest = strRequest & "<VALUEID>" & strQnResponse & "</VALUEID>"
        strRequest = strRequest & "</LIST>"
            
        strResponse = objComboDO.GetComboValue(strRequest)
        
        Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
        strValidationType = objXmlAssist.GetTagValue( _
            xmlDoc.documentElement, "VALIDATIONTYPE")
            
        lngTotScore = lngTotScore + CLng(strValidationType)
        
    Next

    ' empty xmlDoc
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    ' add the <REQUEST> node
    Set xmlRequestNode = objXmlAssist.GetRequestNode(xmlIn)
    Set xmlTempNode = xmlDoc.appendChild(xmlRequestNode)
    
    Set xmlElem = xmlDoc.createElement("UPDATE")
    xmlElem.setAttribute "TYPE", "AFTER"
    Set xmlUpdateNode = xmlTempNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFIND")
    Set xmlTableNode = xmlUpdateNode.appendChild(xmlElem)
    
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlElem = xmlDoc.createElement(strTagName)
    xmlElem.Text = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    xmlTableNode.appendChild xmlElem
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlElem = xmlDoc.createElement(strTagName)
    xmlElem.Text = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    xmlTableNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("ATTITUDETOBORROWINGSCORE")
    xmlElem.Text = CStr(lngTotScore)
    xmlTableNode.appendChild xmlElem
    
    ' APS 22/05/00 Removed Customer Registration component
    strResponse = objApplicationBO.UpdateApplicationFactFind(xmlDoc.xml)
    
    objErrAssist.CheckXMLResponse strResponse, True
    
    '------------------------------------------------------------------------------------------
    ' call BO.DeleteAll
    '------------------------------------------------------------------------------------------
    
    ' empty xmlDoc
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    ' add the <REQUEST> node
    Set xmlRequestNode = objXmlAssist.GetRequestNode(xmlIn)
    Set xmlTempNode = xmlDoc.appendChild(xmlRequestNode)
    
    Set xmlElem = xmlDoc.createElement(cstrTABLE_NAME)
    Set xmlTableNode = xmlTempNode.appendChild(xmlElem)
    
    strTagName = "APPLICATIONNUMBER"
    Set xmlElem = xmlDoc.createElement(strTagName)
    xmlElem.Text = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    xmlTableNode.appendChild xmlElem
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlElem = xmlDoc.createElement(strTagName)
    xmlElem.Text = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    xmlTableNode.appendChild xmlElem
    
    On Error Resume Next
    
    DeleteAll xmlDoc.xml
    
    If Err.Number <> 0 Then
        If Err.Number = omiga4NoRowsAffected Then
            Err.Clear
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If
    
    On Error GoTo CreateVbErr
    
    '------------------------------------------------------------------------------------------
    ' call DO.Create
    '------------------------------------------------------------------------------------------
   
    ' For each relevant node call the data object create
    For intNodeIndex = 0 To intNumNodes - 1
        Set xmlElem = xmlNodeList.Item(intNodeIndex)
        
        objAttitudeToBorrowingDO.Create xmlElem.xml
    Next
 
    Set objAttitudeToBorrowingDO = Nothing
    Set objApplicationBO = Nothing
    Set objComboDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

CreateVbErr:
   
    Set objAttitudeToBorrowingDO = Nothing
    Set objApplicationBO = Nothing
    Set objComboDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing update data
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim objXmlUpdateElement As IXMLDOMElement

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace AttitudeToBorrowingDO reference with actual DO required
' *****************************************************************************
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get the after image to pass to the data object
    Dim blnFound As Boolean
    Dim intIndex As Integer
    Dim intNoOfUpdates As Integer
    
    blnFound = False
    intIndex = 0
    intNoOfUpdates = objXmlIn.getElementsByTagName("UPDATE").length
    
    ' Loop through each Update element to find the one with the after image
    Do While blnFound = False And intIndex < intNoOfUpdates
        Set objXmlUpdateElement = objXmlIn.getElementsByTagName("UPDATE").Item(intIndex)
        
        If objXmlUpdateElement.getAttribute("TYPE") = "AFTER" Then
            blnFound = True
        End If
        
        intIndex = intIndex + 1
    Loop
       
    If blnFound = False Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoAfterImagePresent
    End If

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = objXmlUpdateElement.getElementsByTagName(cstrTABLE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer
    
    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
    
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set xmlElem = xmlNodeList.Item(intNodeIndex)
        objAttitudeToBorrowingDO.Update xmlElem.xml
    Next
 
    Set objXmlUpdateElement = Nothing
    Set objAttitudeToBorrowingDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
    Exit Sub

UpdateVbErr:
   
    Set objXmlUpdateElement = Nothing
    Set objAttitudeToBorrowingDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Public Sub Delete(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform delete
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "Delete"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace AttitudeToBorrowingDO reference with actual DO required
' *****************************************************************************
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
   
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
    
    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set xmlElem = xmlNodeList.Item(intNodeIndex)
        objAttitudeToBorrowingDO.Delete xmlElem.xml
    Next
 
    Set objAttitudeToBorrowingDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:
   
    Set objAttitudeToBorrowingDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub DeleteAll(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete the data for all instances of the persistant data associated with
'   this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "DeleteAll"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objAttitudeToBorrowingDO As AttitudeToBorrowingDO
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim objXmlAssist As New XMLAssist

    If objContext Is Nothing Then
        Set objAttitudeToBorrowingDO = New AttitudeToBorrowingDO
    Else
        Set objAttitudeToBorrowingDO = _
            objContext.CreateInstance(App.Title & ".AttitudeToBorrowingDO")
    End If
    
    Dim strTagValue, strTagName As String
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    strTagName = "APPLICATIONNUMBER"
    strTagValue = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Expected " & strTagName & " tag"
    End If
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strTagValue = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Expected " & strTagName & " tag"
    End If
        
    ' Pass through to the data object
    objAttitudeToBorrowingDO.DeleteAll vstrXMLRequest
    
    Set objAttitudeToBorrowingDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub
    
DeleteAllVbErr:

    Set objAttitudeToBorrowingDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
        objContext.SetComplete
    End If
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



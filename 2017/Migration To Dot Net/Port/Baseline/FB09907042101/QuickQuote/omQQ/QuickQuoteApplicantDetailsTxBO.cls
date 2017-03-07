VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "QuickQuoteApplicantDetailsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuoteApplicantDetailsTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'Dependencies:
'               QuickQuoteApplicantDetailsDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     01/08/99    Created
'APS    02/03/00    SYS0324 - DeleteAll error not handled in SaveQuickQuoteData
'APS    22/05/00    SYS0779 Removed Customer Registration component
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "QUICKQUOTEAPPLICANTDETAILS"
 
 
Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement

    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
    
    Dim intNodeIndex As Integer
    
    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        
        objQuickQuoteApplicantDetailsDO.Create objElement.xml
    Next
 
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

CreateVbErr:
   
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
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
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objXmlUpdateElement As IXMLDOMElement

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    
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
    Set objXmlNodeList = objXmlUpdateElement.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer
    
    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
    
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objQuickQuoteApplicantDetailsDO.Update objElement.xml
    Next
 
    Set objXmlUpdateElement = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
    Exit Sub

UpdateVbErr:
   
    Set objXmlUpdateElement = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
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
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
     
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
   
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
    
    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objQuickQuoteApplicantDetailsDO.Delete objElement.xml
    Next
 
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:
   
    ' APS 01/03/00 SYS0324 - DeleteAll error not handled in SaveQuickQuoteData
    ' As DOAssist.Delete raises the omiga4NoRowsAffected we need
    ' functionality here to resume next because it is at this stage that we set
    ' abort
    If Err.Number = omiga4NoRowsAffected Then
        Resume Next
    End If
    
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
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
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "DeleteAll"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
        
    ' Pass through to the data object
    objQuickQuoteApplicantDetailsDO.DeleteAll (vstrXMLRequest)
    
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub
    
DeleteAllVbErr:

    Set objQuickQuoteApplicantDetailsDO = Nothing
    
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


Public Sub SaveQuickQuoteData(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Save a set of Quick Quote data, including some ApplicationFactFind data
'   (Purchase Price, O/S Loan Amount and Total Monthly Repayments)
' pass:
'   vstrXMLRequest
'       xml Request.
'       Format: As per QuickQuoteApplicantDetailsBO.SaveQuickQuoteData, i.e:
'       <REQUEST>
'            <QUICKQUOTEAPPLICANTDETAILSLIST>
'                <QUICKQUOTEAPPLICANTDETAILS>
'                    <QUICKQUOTEOUTGOINGSLIST>
'                        <QUICKQUOTEOUTGOINGS>
'                        </QUICKQUOTEOUTGOINGS>
'                    </QUICKQUOTEOUTGOINGSLIST>
'                    <QUICKQUOTEPERSONALDEBTSLIST>
'                        <QUICKQUOTEPERSONALDEBTS>
'                        </QUICKQUOTEPERSONALDEBTS>
'                    </QUICKQUOTEPERSONALDEBTSLIST>
'                </QUICKQUOTEAPPLICANTDETAILS>
'            </QUICKQUOTEAPPLICANTDETAILSLIST>
'            <UPDATE TYPE="AFTER">
'               <APPLICATIONFACTFIND>
'               </APPLICATIONFACTFIND>
'            </UPDATE>
'       </REQUEST>
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SaveQuickQuoteDataVbErr
    
    Const strFunctionName As String = "SaveQuickQuoteData"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlTempNode As IXMLDOMNode
    Dim objXmlApplicantNode As IXMLDOMNode
    Dim objQuickQuoteOutgoingsBO As QuickQuoteOutgoingsBO
    Dim objQuickQuotePersonalDebtsBO As QuickQuotePersonalDebtsBO
    Dim objApplicationBO As ApplicationBO
    Dim objXmlAssist As New XMLAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strTagName, strTableName, strListName As String
    Dim strResponse As String
    Dim intApplicant As Integer
    Dim intNumApplicants As Integer
    Dim lngErrNum As Long
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' APS 22/05/00 SYS0779 Removed Customer Registration component
    If objContext Is Nothing Then
        Set objQuickQuoteOutgoingsBO = New QuickQuoteOutgoingsBO
        Set objQuickQuotePersonalDebtsBO = New QuickQuotePersonalDebtsBO
        Set objApplicationBO = New ApplicationBO
    Else
        Set objQuickQuoteOutgoingsBO = objContext.CreateInstance( _
            App.Title & ".QuickQuoteOutgoingsBO")
        Set objQuickQuotePersonalDebtsBO = objContext.CreateInstance( _
            App.Title & ".QuickQuotePersonalDebtsBO")
        Set objApplicationBO = objContext.CreateInstance( _
            gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    ' check how many applicant details have been supplied
    intNumApplicants = _
        objXmlIn.getElementsByTagName("QUICKQUOTEAPPLICANTDETAILS").length
    
    '------------------------------------------------------------------------------------------
    ' delete existing Quick Quote outgoings details
    '------------------------------------------------------------------------------------------
    
    strTableName = "QUICKQUOTEOUTGOINGS"
    strListName = strTableName & "LIST"
    
    For intApplicant = 0 To intNumApplicants - 1
        
        Set objXmlApplicantNode = objXmlIn.getElementsByTagName( _
            "QUICKQUOTEAPPLICANTDETAILS").Item(intApplicant)
            
        ' empty objXmlDoc
        Set objXmlDoc = New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
    
        ' add the <REQUEST> node
        Set objXmlRequestNode = objXmlAssist.GetRequestNode(objXmlIn)
        Set objXmlTempNode = objXmlDoc.appendChild(objXmlRequestNode)
    
        ' add the table name
        Set xmlElem = objXmlDoc.createElement(strTableName)
        Set objXmlTempNode = objXmlTempNode.appendChild(xmlElem)
    
        ' add the criteria fields
    
        strTagName = "CUSTOMERNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem
        
        strTagName = "CUSTOMERVERSIONNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem
        
        strTagName = "APPLICATIONNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem
        
        strTagName = "APPLICATIONFACTFINDNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem
        
        ' APS 01/03/00 - Moved error processing into the DeleteAll method
        strResponse = objQuickQuoteOutgoingsBO.DeleteAll(objXmlDoc.xml)
        lngErrNum = objErrAssist.CheckXMLResponse(strResponse, True)
        'If lngErrNum <> 0 And lngErrNum <> omiga4NoRowsAffected Then
        '    objErrAssist.CheckXMLResponse strResponse, True
        'End If
        
    Next
    
    '------------------------------------------------------------------------------------------
    ' delete existing Quick Quote personal debts details
    '------------------------------------------------------------------------------------------
    
    strTableName = "QUICKQUOTEPERSONALDEBTS"
    strListName = strTableName & "LIST"
    
    For intApplicant = 0 To intNumApplicants - 1
    
        Set objXmlApplicantNode = objXmlIn.getElementsByTagName( _
            "QUICKQUOTEAPPLICANTDETAILS").Item(intApplicant)

        ' empty objXmlDoc
        Set objXmlDoc = New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True

        ' add the <REQUEST> node
        Set objXmlRequestNode = objXmlAssist.GetRequestNode(objXmlIn)
        Set objXmlTempNode = objXmlDoc.appendChild(objXmlRequestNode)

        ' add the table name
        Set xmlElem = objXmlDoc.createElement(strTableName)
        Set objXmlTempNode = objXmlTempNode.appendChild(xmlElem)

        ' add the criteria fields

        strTagName = "CUSTOMERNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem

        strTagName = "CUSTOMERVERSIONNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem

        strTagName = "APPLICATIONNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem

        strTagName = "APPLICATIONFACTFINDNUMBER"
        Set xmlElem = objXmlDoc.createElement(strTagName)
        xmlElem.Text = objXmlAssist.GetTagValue(objXmlApplicantNode, strTagName)
        objXmlTempNode.appendChild xmlElem
        
        ' APS 01/03/00 - Moved error processing into the DeleteAll method
        strResponse = objQuickQuotePersonalDebtsBO.DeleteAll(objXmlDoc.xml)
        lngErrNum = objErrAssist.CheckXMLResponse(strResponse, True)
        'If lngErrNum <> 0 And lngErrNum <> omiga4NoRowsAffected Then
        '    objErrAssist.CheckXMLResponse strResponse, True
        'End If
    Next
    
    '------------------------------------------------------------------------------------------
    ' save Quick Quote applicant details after deleting existing ones
    '------------------------------------------------------------------------------------------
    
    strTableName = "QUICKQUOTEAPPLICANTDETAILS"
    
    ' for each applicant, do the delete and ignore record not found error
    For intApplicant = 0 To intNumApplicants - 1
    
        Set objXmlApplicantNode = objXmlIn.getElementsByTagName( _
            strTableName).Item(intApplicant)
        
        ' APS 01/03/00 - Moved error processing to error handler of the
        ' Delete method so we avoid setting abort on a successful transaction
        'On Error Resume Next
        Me.Delete objXmlApplicantNode.xml
        'If Err.Number <> 0 And Err.Number <> omiga4RecordNotFound Then
        '    Err.Raise Err.Number, Err.Source, Err.Description
        'End If
        'On Error GoTo SaveQuickQuoteDataVbErr
    Next
    
    ' do the create
    Me.Create vstrXMLRequest
    
    '------------------------------------------------------------------------------------------
    ' create Quick Quote outgoings details and personal debts details
    '------------------------------------------------------------------------------------------
    
    strListName = "QUICKQUOTEOUTGOINGSLIST"
    
    For intApplicant = 0 To intNumApplicants - 1
        Set xmlElem = objXmlIn.getElementsByTagName( _
            strListName).Item(intApplicant)
        If Not xmlElem Is Nothing Then
            strResponse = objQuickQuoteOutgoingsBO.Create(xmlElem.xml)
            objErrAssist.CheckXMLResponse strResponse, True
        End If
    Next
    
    strListName = "QUICKQUOTEPERSONALDEBTSLIST"
    
    For intApplicant = 0 To intNumApplicants - 1
        Set xmlElem = objXmlIn.getElementsByTagName( _
            strListName).Item(intApplicant)
        If Not xmlElem Is Nothing Then
            strResponse = objQuickQuotePersonalDebtsBO.Create(xmlElem.xml)
            objErrAssist.CheckXMLResponse strResponse, True
        End If
    Next
    
    '------------------------------------------------------------------------------------------
    ' update ApplicationFactFind
    '------------------------------------------------------------------------------------------
    ' APS 22/05/00 SYS0779 Removed Customer Registration component
    objApplicationBO.UpdateApplicationFactFind vstrXMLRequest
    
    'SG 28/05/02 SYS4767 START
    'STB: MSMS0011 - Update the Application with any IntermediaryGUID specified.
    objApplicationBO.UpdateApplication vstrXMLRequest
    'SG 28/05/02 SYS4767 END
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objQuickQuoteOutgoingsBO = Nothing
    Set objQuickQuotePersonalDebtsBO = Nothing
    Set objApplicationBO = Nothing

    Exit Sub

SaveQuickQuoteDataVbErr:

    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objQuickQuoteOutgoingsBO = Nothing
    Set objQuickQuotePersonalDebtsBO = Nothing
    Set objApplicationBO = Nothing
    
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



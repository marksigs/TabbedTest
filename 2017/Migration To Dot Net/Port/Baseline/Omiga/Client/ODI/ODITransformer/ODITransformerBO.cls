VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODITransformerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ODITransformerBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     23/08/01    Expanded class created by LD.
'DS     30/04/02    Use FreeThreadedDOMDocument40.
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'STB    22/05/02    SYS4609 Enabled GetUserMortgageAdminDetails and SaveUserMortgageAdminDetails.
'STB    10/07/02    SYS5101 Enabled Save/CreateBaseRateSet().
'SG/JR  18/11/02    SYS5765/SYSMCP1256 Added GetThirdPartyDetails to DoRequest.
'PSC    16/01/07    EP2_855 Add GetAccountCustomers
'------------------------------------------------------------------------------------------
Option Explicit

'Implements ObjectControl
'
'Private gobjContext As ObjectContext
'
'Private gobjTxBO As IOm4Ph2TxBO
'Private gobjDO As IOm4Ph2DO
'
'Private Sub ObjectControl_Activate()
'    Set gobjContext = GetObjectContext()
'End Sub
'
'Private Function ObjectControl_CanBePooled() As Boolean
'    ObjectControl_CanBePooled = False
'End Function
'
'Private Sub ObjectControl_Deactivate()
'    Set gobjContext = Nothing
'End Sub

Public Function Request(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------
On Error GoTo RequestVbErr
    
    Const strFunctionName As String = "Request"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objODITransformerState As ODITransformerState
    
'    ' will be superceded by any subsequent SetAbort
'    gobjContext.SetComplete
'
'    ' ensure global objects released before running request
'    Set gobjTxBO = Nothing
'    Set gobjDO = Nothing
    
    ' create REQUEST & RESPONSE FreeThreadedDomDocuments
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    Set xmlInDoc = xmlLoad(vstrXMLRequest, strFunctionName)
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    '------------------------------------------------------------------------------------------
    ' initialise ODITransformerState
    '------------------------------------------------------------------------------------------

    ' initialise state from the request; if state is not present in the request
    ' (i.e. in the case of ValidateUserLogon), it will set the state to null.
    Set objODITransformerState = New ODITransformerState
    objODITransformerState.Initialise xmlInDoc
    
    '------------------------------------------------------------------------------------------
    ' process the operation(s)
    '------------------------------------------------------------------------------------------
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        DoRequest objODITransformerState, xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                ' copy attributes from REQUEST node to OPERATION node
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest objODITransformerState, xmlOperationNode, xmlResponseNode
        Next
    End If

    objODITransformerState.AddToResponseNode xmlResponseNode
    
    ' if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Request = xmlOutDoc.xml
    Debug.Print Request
    
RequestExit:
    
    ' clean up ================================================================
    
'    ' release global objects
'    Set gobjTxBO = Nothing
'    Set gobjDO = Nothing

    ' release XML DOM objects
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Set objODITransformerState = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Exit Function

RequestVbErr:
    
    ' create ERROR response block from Err object
    Request = errCreateErrorResponse()
    Debug.Print Request
    
    ' perform clean up
    GoTo RequestExit

End Function

Private Sub DoRequest( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       raises oeNotImplemented if OPERATION value not in list of known functions.
'------------------------------------------------------------------------------------------
    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "CREATEBASERATESET"
            CreateBaseRateSet objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "FINDACCOUNTDETAILS"
            FindAccountDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "FINDBUSINESSFORCUSTOMER"
            FindBusinessForCustomer objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "FINDCUSTOMERLIST"
            FindCustomerList objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "FINDMORTGAGEBALANCELIST"
            FindMortgageBalanceList objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "GETACCOUNTSUMMARY"
            GetAccountSummary objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "GETACCOUNTBALANCEANDSTATUS"
            GetAccountBalanceAndStatus objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "GETCUSTOMERDETAILS"
            GetCustomerDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "GETNEWACCOUNTNUMBER"
            GetNewAccountNumber objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "GETNEWCUSTOMERNUMBER"
            GetNewCustomerNumber objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "GETPAYMENTAMOUNT"
            GetPaymentAmount objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "LOGOFFUSER"
            LogOffUser objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        ' Completions does not use ODI
'        Case "PROCESSCOMPLETIONSINTERFACE"
'            ProcessCompletionsInterface objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "SAVEBASERATESET"
            SaveBaseRateSet objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "SAVEMORTGAGEPRODUCTDETAILS"
            SaveMortgageProductDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "SAVETHIRDPARTYDETAILS"
            SaveThirdPartyDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "VALIDATEUSERLOGON"
            ValidateUserLogon objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "SAVEUSERMORTGAGEADMINDETAILS"
            SaveUserMortgageAdminDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
        
        Case "GETUSERMORTGAGEADMINDETAILS"
            GetUserMortgageAdminDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "CREATEBASERATESET"
            CreateBaseRateSet objODITransformerState, vxmlRequestNode, vxmlResponseNode
            
        Case "UPDATEBASERATESET"
            SaveBaseRateSet objODITransformerState, vxmlRequestNode, vxmlResponseNode

        'SG 18/11/02 SYS5765
        Case "GETTHIRDPARTYDETAILS"
            GetThirdPartyDetails objODITransformerState, vxmlRequestNode, vxmlResponseNode
        ' PSC 17/01/2007 EP2_855
        Case "GETACCOUNTCUSTOMERS"
            GetAccountCustomers objODITransformerState, vxmlRequestNode, vxmlResponseNode
        ' PSC 17/01/2007 EP2_855 - End
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    '<VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName, TypeName(Me)

End Function


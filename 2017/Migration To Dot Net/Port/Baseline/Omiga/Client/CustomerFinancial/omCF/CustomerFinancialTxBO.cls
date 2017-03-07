VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CustomerFinancialTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4TxBO' with actual class name
' *****************************************************************************

'Workfile:      Omiga4TxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  List any other dependent components
'               e.g. CustomerFinancialDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'PSC    01/10/99    Added new error processing
'BG     27/06/00    Added SaveMortgageAccountAndLoan manager method which calls Create/Delete/UpdateMortgageAccount
'                   and Create/Delete/UpdateMortgageLoan respectively.
'SR     23/06/00    New method SaveLiability
'MC     05/07/00    SYS0726 Amended various methods to allow changing client for record.
'SR     17/07/00    Modified method SaveLiability - Do not validate before deletion operation
'SA     06/06/01    Modified method UpdateLiability to cope with changing clients.
'-------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerFinancialTxBO
Implements ObjectControl

Private Const cstrROOT_NODE_NAME = "MORTGAGEACCOUNT"

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist


'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006 Added CreateAccountRelationships
'                                    DeleteAccountRelationships
'                                    UpdateAccountRelationships
'                                    CreateSpecialFeature
'                                    UpdateSpecialFeature
'                                    DeleteSpecialFeature
'                                    FindSpecialFeatureList
'MV     14/08/2002  BMIDS00311  Core Ref AQR: SYS1705 ;Amended ICustomerFinancialTxBO_SaveMortgageAccountAndLoan
'MDC    08/10/2002  BMIDS00561  Return AccountGuid generated in CreateLiability
'SR     12/06/2004  BMIDS772    Modified methods CreateBankCard, CreateLiability, CreateRegularOutgoing
'                               DeleteBankCard, DeleteLiability, DeleteRegularOutgoing - Update FinancialSummary
'                               New private method 'UpdateFinancialSummary'
'SR     15/06/2004  BMIDS772    Modified methods CreateArrearsHistory, CreateBankruptcyHistory, CreateCCJHistory,
'                               CreateDeclinedMortgage, CreateMortgageAccount, DeleteArrearsHistory, DeleteBankruptcyHistory,
'                               DeleteCCJHistory, DeleteDeclinedMortgage, DeleteMortgageAccount
'-------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'SR     09/11/2005  MAR496 - modified function 'ICustomerFinancialTxBO_CreateLiability'
'                   Save BankCard details, if required
'JD     21/04/2006  MAR1371 CreateLiability: Check for >1 BankCreditCard records to create
'----------------------------------------------------------------------------------------

Private Const cstrACCOUNT_ROOT_NODE_NAME = "ACCOUNT"
Private Const cstrFEATURE_ROOT_NODE_NAME = "MORTGAGEACCOUNTSPECIALFEATURE"

Private Function ICustomerFinancialTxBO_CreateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateArrearsHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateArrearsHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "ARREARSHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("ARREARSHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "ARREARSHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateArrearsHistory xmlElement
    Else
        ' For each relevant node call the data object CreateArrearsHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateArrearsHistory xmlElement
        Next
    End If
    
     '-------------------------------------------------------------------------------------------------
    'SR 16/06/2004 : BMIDS772 - Update FinancialSummary record for with new ARREARSHISTORYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "ARREARSHISTORYINDICATOR"
    End If

    Set ICustomerFinancialTxBO_CreateArrearsHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateArrearsHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateArrearsHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateArrearsHistoryExit

End Function

Private Function ICustomerFinancialTxBO_CreateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankCard multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateBankCard"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKCREDITCARD" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKCREDITCARD")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKCREDITCARD" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateBankCard xmlElement
    Else
        ' For each relevant node call the data object CreateBankCard
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateBankCard xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 08/06/2004 : BMIDS772 - Update FinancialSummary record for with new BANKCARDINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "BANKCARDINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_CreateBankCard = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankCardExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateBankCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankCardExit

End Function

Private Function ICustomerFinancialTxBO_CreateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateBankruptcyHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateBankruptcyHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKRUPTCYHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKRUPTCYHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKRUPTCYHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateBankruptcyHistory xmlElement
    Else
        ' For each relevant node call the data object CreateBankruptcyHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateBankruptcyHistory xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 15/06/2004 : BMIDS772 - Update FinancialSummary record for with new BANKCARDINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "BANKRUPTCYHISTORYINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_CreateBankruptcyHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBankruptcyHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateBankruptcyHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialTxBO_CreateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateCCJHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateCCJHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "CCJHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CCJHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CCJHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateCCJHistory xmlElement
    Else
        ' For each relevant node call the data object CreateCCJHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateCCJHistory xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 16/06/2004 : BMIDS772 - Update FinancialSummary record for with new CCJHISTORYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "CCJHISTORYINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_CreateCCJHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCCJHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateCCJHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateCCJHistoryExit

End Function

Private Function ICustomerFinancialTxBO_CreateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateDeclinedMortgage multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateDeclinedMortgage"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "DECLINEDMORTGAGE" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("DECLINEDMORTGAGE")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DECLINEDMORTGAGE" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateDeclinedMortgage xmlElement
    Else
        ' For each relevant node call the data object CreateDeclinedMortgage
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateDeclinedMortgage xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 14/06/2004 : BMIDS772 - Update FinancialSummary record for with new DECLINEDMORTGAGEINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "DECLINEDMORTGAGEINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_CreateDeclinedMortgage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateDeclinedMortgageExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateDeclinedMortgageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateDeclinedMortgageExit

End Function

Private Function ICustomerFinancialTxBO_CreateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateGroupConnection multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateGroupConnection"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "GROUPCONNECTION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("GROUPCONNECTION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "GROUPCONNECTION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateGroupConnection xmlElement
    Else
        ' For each relevant node call the data object CreateGroupConnection
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateGroupConnection xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_CreateGroupConnection = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateGroupConnectionExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateGroupConnectionVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateGroupConnectionExit

End Function

Private Function ICustomerFinancialTxBO_CreateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLiability multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateLiability"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Dim xmlKeysNode As IXMLDOMNode 'BMIDS00561 MDC 08/10/2002
    Dim xmlCreditCardNodeList As IXMLDOMNodeList 'MAR1371
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    'BMIDS00561 MDC 08/10/2002
    Set xmlKeysNode = xmlOut.createElement("GENERATEDKEYS")
    'BMIDS00561 MDC 08/10/2002 - End
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LOANSLIABILITIES" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LOANSLIABILITIES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LOANSLIABILITIES" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateLiability xmlElement
        'BMIDS00561 MDC 08/10/2002
        m_objXmlAssist.SetMandatoryChildText xmlKeysNode, "ACCOUNTGUID", _
                                m_objXmlAssist.GetElementText(xmlElement, ".//ACCOUNTGUID")
        'BMIDS00561 MDC 08/10/2002 - End
    Else
        ' For each relevant node call the data object CreateLiability
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateLiability xmlElement
            'BMIDS00561 MDC 08/10/2002
            m_objXmlAssist.SetMandatoryChildText xmlKeysNode, "ACCOUNTGUID", _
                                    m_objXmlAssist.GetElementText(xmlElement, ".//ACCOUNTGUID")
            'BMIDS00561 MDC 08/10/2002 - End
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 11/06/2004 : BMIDS772 - Update FinancialSummary record for with new LOANLIABILITYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "LOANLIABILITYINDICATOR"
    End If
    
    'SR - 09/11/2005 :MAR496
    'JD MAR1371 Check for >1 BankCreditCard records to create
    Set xmlCreditCardNodeList = vxmlRequest.selectNodes(".//BANKCREDITCARD")
    'If Not vxmlRequest.selectSingleNode(".//BANKCREDITCARD") Is Nothing Then
    If xmlCreditCardNodeList.length > 0 Then
        For Each xmlElement In xmlCreditCardNodeList
            objICustomerFinancialDO.CreateBankCard xmlElement
        Next
    End If
    'SR 09/11/2005 : MAR496 - End
    
    'BMIDS00561 MDC 08/10/2002
    xmlResponseElem.appendChild xmlKeysNode.cloneNode(True)
    'BMIDS00561 MDC 08/10/2002 - End
    
    Set ICustomerFinancialTxBO_CreateLiability = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLiabilityExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing   'BMIDS00561 MDC 08/10/2002
    
    Exit Function

CreateLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLiabilityExit

End Function

Private Function ICustomerFinancialTxBO_CreateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateLifeProduct multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateLifeProduct"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LIFEPRODUCT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LIFEPRODUCT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LIFEPRODUCT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateLifeProduct xmlElement
    Else
        ' For each relevant node call the data object CreateLifeProduct
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateLifeProduct xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_CreateLifeProduct = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLifeProductExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateLifeProductVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLifeProductExit

End Function

Private Function ICustomerFinancialTxBO_CreateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageAccount multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateMortgageAccount"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGEACCOUNT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGEACCOUNT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateMortgageAccount xmlElement, vxmlKeysNode
    Else
        ' For each relevant node call the data object CreateMortgageAccount
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateMortgageAccount xmlElement, vxmlKeysNode
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 14/06/2004 : BMIDS772 - Update FinancialSummary record for with new EXISTINGMORTGAGEINDICATOR
    '                only If EXISTINGMORTGAGEINDICATOR node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "EXISTINGMORTGAGEINDICATOR"
    End If
    
    
    Set ICustomerFinancialTxBO_CreateMortgageAccount = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateMortgageAccountExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateMortgageAccountVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageAccountExit

End Function

Private Function ICustomerFinancialTxBO_CreateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLoan multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateMortgageLoan"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGELOAN" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGELOAN")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGELOAN" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateMortgageLoan xmlElement, vxmlKeysNode
    Else
        ' For each relevant node call the data object CreateMortgageLoan
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateMortgageLoan xmlElement, vxmlKeysNode
        Next
    End If
    
    Set ICustomerFinancialTxBO_CreateMortgageLoan = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
        
CreateMortgageLoanExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateMortgageLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageLoanExit

End Function

Private Function ICustomerFinancialTxBO_CreateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateRegularOutgoings multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateRegularOutgoings"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGELOAN" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("REGULAROUTGOINGS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "REGULAROUTGOINGS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.CreateRegularOutgoings xmlElement
    Else
        ' For each relevant node call the data object CreateRegularOutgoings
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.CreateRegularOutgoings xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 12/06/2004 : BMIDS772 - Update FinancialSummary record for with new REGULAROUTGOINGSINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "REGULAROUTGOINGSINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_CreateRegularOutgoings = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateRegularOutgoingsExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateRegularOutgoingsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateRegularOutgoingsExit

End Function

Private Function ICustomerFinancialTxBO_DeleteArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteArrearsHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteArrearsHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "ARREARSHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("ARREARSHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "ARREARSHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteArrearsHistory xmlElement
    Else
        ' For each relevant node call the data object DeleteArrearsHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteArrearsHistory xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 16/06/2004 : BMIDS772 - Update FinancialSummary record for with new ARREARSHISTORYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "ARREARSHISTORYINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteArrearsHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteArrearsHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteArrearsHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteArrearsHistoryExit

End Function

Private Function ICustomerFinancialTxBO_DeleteBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankCard multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteBankCard"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKCREDITCARD" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKCREDITCARD")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKCREDITCARD" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteBankCard xmlElement
    Else
        ' For each relevant node call the data object DeleteBankCard
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteBankCard xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 08/06/2004 : BMIDS772 - Update FinancialSummary record for with new BANKCARDINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "BANKCARDINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteBankCard = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteBankCardExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteBankCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankCardExit

End Function

Private Sub UpdateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement, ByVal strIndicator As String)
    
    Dim xmlFinancialSummary As IXMLDOMNode, xmlIndicatorNode As IXMLDOMNode
    Dim Cmd As ADODB.Command, Cn As ADODB.Connection
    Dim objIADOAssist As IADOAssist
    
    Dim strAppNo As String, strAFFNo As String
    Dim strIndicatorVal As String
    
    strAppNo = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNo = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        
    Set xmlFinancialSummary = vxmlRequest.selectSingleNode("FINANCIALSUMMARY")
    If Not xmlFinancialSummary Is Nothing Then
        strIndicatorVal = m_objXmlAssist.GetTagValue(xmlFinancialSummary, strIndicator)
    End If
    
    Set Cn = New ADODB.Connection
    Set Cmd = New ADODB.Command
    Set objIADOAssist = New ADOAssist
    
    Cn.ConnectionString = objIADOAssist.GetConnStr()
    Cn.open
    Set Cmd.ActiveConnection = Cn
    Cmd.CommandText = "usp_SaveFinancialSummary"
    Cmd.CommandType = adCmdStoredProc
    
    With Cmd
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, Len(strAppNo), strAppNo)
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, Len(strAFFNo), strAFFNo)
        .Parameters.Append .CreateParameter(, adBSTR, adParamInput, Len(strIndicator), strIndicator)
        If Len(strIndicatorVal) > 0 Then
               .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, strIndicatorVal)
        End If
    End With
    
    Cmd.Execute
    
    Cmd.ActiveConnection = Nothing
    Cn.Close
    
    m_objContext.SetComplete
    
    Set Cmd = Nothing
    Set Cn = Nothing
End Sub


Private Function ICustomerFinancialTxBO_DeleteBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteBankruptcyHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteBankruptcyHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKRUPTCYHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKRUPTCYHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKRUPTCYHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteBankruptcyHistory xmlElement
    Else
        ' For each relevant node call the data object DeleteBankruptcyHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteBankruptcyHistory xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 15/06/2004 : BMIDS772 - Update FinancialSummary record for with new BANKCARDINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "BANKRUPTCYHISTORYINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteBankruptcyHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteBankruptcyHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteBankruptcyHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialTxBO_DeleteCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteCCJHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteCCJHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "CCJHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CCJHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CCJHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteCCJHistory xmlElement
    Else
        ' For each relevant node call the data object DeleteCCJHistory
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteCCJHistory xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 16/06/2004 : BMIDS772 - Update FinancialSummary record for with new CCJHISTORYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "CCJHISTORYINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteCCJHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteCCJHistoryExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteCCJHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCCJHistoryExit

End Function

Private Function ICustomerFinancialTxBO_DeleteDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteDeclinedMortgage multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteDeclinedMortgage"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "DECLINEDMORTGAGE" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("DECLINEDMORTGAGE")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DECLINEDMORTGAGE" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteDeclinedMortgage xmlElement
    Else
        ' For each relevant node call the data object DeleteDeclinedMortgage
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteDeclinedMortgage xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 14/06/2004 : BMIDS772 - Update FinancialSummary record for with new DECLINEDMORTGAGEINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "DECLINEDMORTGAGEINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteDeclinedMortgage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteDeclinedMortgageExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteDeclinedMortgageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteDeclinedMortgageExit

End Function

Private Function ICustomerFinancialTxBO_DeleteGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteGroupConnection multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteGroupConnection"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "GROUPCONNECTION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("GROUPCONNECTION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "GROUPCONNECTION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteGroupConnection xmlElement
    Else
        ' For each relevant node call the data object DeleteGroupConnection
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteGroupConnection xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_DeleteGroupConnection = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteGroupConnectionExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteGroupConnectionVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteGroupConnectionExit

End Function

Private Function ICustomerFinancialTxBO_DeleteLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLiability multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteLiability"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LOANSLIABILITIES" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LOANSLIABILITIES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LOANSLIABILITIES" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteLiability xmlElement
    Else
        ' For each relevant node call the data object DeleteLiability
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteLiability xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 12/06/2004 : BMIDS772 - Update FinancialSummary record for with new LOANLIABILITYINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "LOANLIABILITYINDICATOR"
    End If

    Set ICustomerFinancialTxBO_DeleteLiability = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteLiabilityExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLiabilityExit

End Function

Private Function ICustomerFinancialTxBO_DeleteLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteLifeProduct multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteLifeProduct"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LIFEPRODUCT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LIFEPRODUCT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LIFEPRODUCT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteLifeProduct xmlElement
    Else
        ' For each relevant node call the data object DeleteLifeProduct
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteLifeProduct xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_DeleteLifeProduct = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteLifeProductExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteLifeProductVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLifeProductExit

End Function

Private Function ICustomerFinancialTxBO_DeleteMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageAccount multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteMortgageAccount"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGEACCOUNT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGEACCOUNT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteMortgageAccount xmlElement
    Else
        ' For each relevant node call the data object DeleteMortgageAccount
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteMortgageAccount xmlElement
        Next
    End If
    
    '-------------------------------------------------------------------------------------------------
    'SR 14/06/2004 : BMIDS772 - Update FinancialSummary record for with new EXISTINGMORTGAGEINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '-------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "EXISTINGMORTGAGEINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteMortgageAccount = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteMortgageAccountExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteMortgageAccountVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageAccountExit

End Function

Private Function ICustomerFinancialTxBO_DeleteMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteMortgageLoan multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteMortgageLoan"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGELOAN" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGELOAN")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGELOAN" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteMortgageLoan xmlElement
    Else
        ' For each relevant node call the data object DeleteMortgageLoan
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteMortgageLoan xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_DeleteMortgageLoan = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteMortgageLoanExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteMortgageLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageLoanExit

End Function

Private Function ICustomerFinancialTxBO_DeleteRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteRegularOutgoings multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo NoRowsAffectedErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteRegularOutgoings"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "REGULAROUTGOINGS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("REGULAROUTGOINGS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "REGULAROUTGOINGS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.DeleteRegularOutgoings xmlElement
    Else
        ' For each relevant node call the data object DeleteRegularOutgoings
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.DeleteRegularOutgoings xmlElement
        Next
    End If
    
    '--------------------------------------------------------------------------------------------
    'SR 11/06/2004 : BMIDS772 - Update FinancialSummary record for with new REGULAROUTGOINGSINDICATOR
    '                only If FINANCIALSUMMARY node is passed as part of the request
    '--------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//FINANCIALSUMMARY") Is Nothing Then
        UpdateFinancialSummary vxmlRequest, "REGULAROUTGOINGSINDICATOR"
    End If
    
    Set ICustomerFinancialTxBO_DeleteRegularOutgoings = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteRegularOutgoingsExit:

    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

NoRowsAffectedErr:
    
    If omiga4NoRowsAffected = Err.Number Then
        Resume Next
    End If

DeleteRegularOutgoingsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteRegularOutgoingsExit

End Function

Private Function ICustomerFinancialTxBO_SaveLiability( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo SaveLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_SaveLiability"

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
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim objICFBO As ICustomerFinancialBO

    Dim objIClassDef As ICFClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim botOperation As BOOPERATIONTYPE
  
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LOANSLIABILITIES" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LOANSLIABILITIES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "LOANSLIABILITIES" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
        Set objICFBO = New CustomerFinancialBO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
        Set objICFBO = objContext.CreateInstance(App.Title & ".CustomerFinancialBO")
    End If
    
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadLoansLiabilitiesData
    
    If Not xmlElement Is Nothing Then
        
        botOperation = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
        Select Case botOperation
            Case booCreate:
                Set xmlTempResponseNode = objICFBO.ValidateLiability(xmlElement)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                objICustomerFinancialDO.CreateLiability xmlElement
            Case booUpdate:
                Set xmlTempResponseNode = objICFBO.ValidateLiability(xmlElement)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                objICustomerFinancialDO.UpdateLiability xmlElement
            Case booDelete:
                objICustomerFinancialDO.DeleteLiability xmlElement
        End Select
    Else
        For Each xmlElement In xmlNodeList
            botOperation = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
            Select Case botOperation
                Case booCreate:
                    Set xmlTempResponseNode = objICFBO.ValidateLiability(xmlElement)
                    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                    objICustomerFinancialDO.CreateLiability xmlElement
                Case booUpdate:
                    Set xmlTempResponseNode = objICFBO.ValidateLiability(xmlElement)
                    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                    objICustomerFinancialDO.UpdateLiability xmlElement
                Case booDelete:
                    objICustomerFinancialDO.DeleteLiability xmlElement
            End Select
        Next xmlElement
    End If
        
    Set ICustomerFinancialTxBO_SaveLiability = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveLiabilityExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objICFBO = Nothing
    Set objIClassDef = Nothing
    
    Exit Function

SaveLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_SaveLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveLiabilityExit

End Function

Private Function ICustomerFinancialTxBO_SaveMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_SaveMortgageAccount"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim objIClassDef As ICFClassDef
    Dim botOperation As BOOPERATIONTYPE
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    Set objIClassDef = New CFClassDef

    botOperation = m_objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadMortgageAccountData)

    Select Case botOperation
        Case booCreate: Set ICustomerFinancialTxBO_SaveMortgageAccount = ICustomerFinancialTxBO_CreateMortgageAccount(vxmlRequest, vxmlGeneratedKeys)
        Case booUpdate: Set ICustomerFinancialTxBO_SaveMortgageAccount = ICustomerFinancialTxBO_UpdateMortgageAccount(vxmlRequest)
        Case booDelete: Set ICustomerFinancialTxBO_SaveMortgageAccount = ICustomerFinancialTxBO_DeleteMortgageAccount(vxmlRequest)
        Case Else
            ' N/A
    End Select
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveMortgageAccountExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveMortgageAccountVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_SaveMortgageAccount = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountExit

End Function

Private Function ICustomerFinancialTxBO_UpdateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateArrearsHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArrearsHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateArrearsHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "ARREARSHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("ARREARSHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "ARREARSHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteArrearsHistory xmlPreviousKeyNode
            objICustomerFinancialDO.CreateArrearsHistory xmlElement
        Else
            objICustomerFinancialDO.UpdateArrearsHistory xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateArrearsHistory
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteArrearsHistory xmlPreviousKeyNode
                objICustomerFinancialDO.CreateArrearsHistory xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateArrearsHistory xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateArrearsHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateArrearsHistoryExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateArrearsHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateArrearsHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateArrearsHistoryExit

End Function

Private Function ICustomerFinancialTxBO_UpdateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankCard multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankCardVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateBankCard"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKCREDITCARD" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKCREDITCARD")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKCREDITCARD" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteBankCard xmlPreviousKeyNode
            objICustomerFinancialDO.CreateBankCard xmlElement
        Else
            objICustomerFinancialDO.UpdateBankCard xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateBankCard
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteBankCard xmlPreviousKeyNode
                objICustomerFinancialDO.CreateBankCard xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateBankCard xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateBankCard = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankCardExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateBankCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateBankCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankCardExit

End Function

Private Function ICustomerFinancialTxBO_UpdateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateBankruptcyHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankruptcyHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateBankruptcyHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "BANKRUPTCYHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("BANKRUPTCYHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BANKRUPTCYHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteBankruptcyHistory xmlPreviousKeyNode
            objICustomerFinancialDO.CreateBankruptcyHistory xmlElement
        Else
            objICustomerFinancialDO.UpdateBankruptcyHistory xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateBankruptcyHistory
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteBankruptcyHistory xmlPreviousKeyNode
                objICustomerFinancialDO.CreateBankruptcyHistory xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateBankruptcyHistory xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateBankruptcyHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateBankruptcyHistoryExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateBankruptcyHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateBankruptcyHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateBankruptcyHistoryExit

End Function

Private Function ICustomerFinancialTxBO_UpdateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateCCJHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCCJHistoryVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateCCJHistory"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "CCJHISTORY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CCJHISTORY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CCJHISTORY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteCCJHistory xmlPreviousKeyNode
            objICustomerFinancialDO.CreateCCJHistory xmlElement
        Else
            objICustomerFinancialDO.UpdateCCJHistory xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateCCJHistory
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteCCJHistory xmlPreviousKeyNode
                objICustomerFinancialDO.CreateCCJHistory xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateCCJHistory xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateCCJHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateCCJHistoryExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateCCJHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateCCJHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCCJHistoryExit

End Function

Private Function ICustomerFinancialTxBO_UpdateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateDeclinedMortgage multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDeclinedMortgageVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateDeclinedMortgage"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "DECLINEDMORTGAGE" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("DECLINEDMORTGAGE")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DECLINEDMORTGAGE" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteDeclinedMortgage xmlPreviousKeyNode
            objICustomerFinancialDO.CreateDeclinedMortgage xmlElement
        Else
            objICustomerFinancialDO.UpdateDeclinedMortgage xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateDeclinedMortgage
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteDeclinedMortgage xmlPreviousKeyNode
                objICustomerFinancialDO.CreateDeclinedMortgage xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateDeclinedMortgage xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateDeclinedMortgage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateDeclinedMortgageExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateDeclinedMortgageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateDeclinedMortgage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateDeclinedMortgageExit

End Function

Private Function ICustomerFinancialTxBO_UpdateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateGroupConnection multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateGroupConnectionVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateGroupConnection"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "GROUPCONNECTION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("GROUPCONNECTION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "GROUPCONNECTION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.UpdateGroupConnection xmlElement
    Else
        ' For each relevant node call the data object UpdateGroupConnection
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.UpdateGroupConnection xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateGroupConnection = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateGroupConnectionExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateGroupConnectionVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateGroupConnection = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateGroupConnectionExit

End Function

Private Function ICustomerFinancialTxBO_UpdateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLiability multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLiabilityVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateLiability"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LOANSLIABILITIES" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LOANSLIABILITIES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LOANSLIABILITIES" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteLiability xmlPreviousKeyNode
            objICustomerFinancialDO.CreateLiability xmlElement
        Else
            objICustomerFinancialDO.UpdateLiability xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateLiability
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteLiability xmlPreviousKeyNode, vblnSimpleDelete:=True
                objICustomerFinancialDO.CreateLiability xmlElement
            'SYS2179 Check nodename before update - don't want to update it if customers changing.
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerFinancialDO.UpdateLiability xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateLiability = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateLiabilityExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLiabilityExit

End Function

Private Function ICustomerFinancialTxBO_UpdateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateLifeProduct multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeProductVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateLifeProduct"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Dim xmlPreviousKeyNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "LIFEPRODUCT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("LIFEPRODUCT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "LIFEPRODUCT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerFinancialDO.DeleteLifeProduct xmlPreviousKeyNode
            objICustomerFinancialDO.CreateLifeProduct xmlElement, vblnSimpleCreate:=True
            objICustomerFinancialDO.UpdateLifeProduct xmlElement, vblnUpdatePolicyRelationship:=False
        Else
            objICustomerFinancialDO.UpdateLifeProduct xmlElement
        End If
    Else
        ' For each relevant node call the data object UpdateLifeProduct
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerFinancialDO.DeleteLifeProduct xmlPreviousKeyNode, vblnSimpleDelete:=True
                objICustomerFinancialDO.CreateLifeProduct xmlElement, vblnSimpleCreate:=True
                objICustomerFinancialDO.UpdateLifeProduct xmlElement, vblnUpdatePolicyRelationship:=False
            Else
                objICustomerFinancialDO.UpdateLifeProduct xmlElement
            End If
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateLifeProduct = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdatelifeproductExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing
    
    Exit Function

UpdateLifeProductVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateLifeProduct = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatelifeproductExit

End Function

Private Function ICustomerFinancialTxBO_UpdateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageAccount multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageAccountVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateMortgageAccount"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGEACCOUNT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGEACCOUNT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGEACCOUNT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.UpdateMortgageAccount xmlElement
    Else
        ' For each relevant node call the data object UpdateMortgageAccount
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.UpdateMortgageAccount xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateMortgageAccount = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageAccountExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateMortgageAccountVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateMortgageAccount = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageAccountExit

End Function

Private Function ICustomerFinancialTxBO_SaveMortgageAccountAndLoan(ByVal vxmlRequest As IXMLDOMElement, _
                                                        ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo SaveMortgageAccountAndLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_SaveMortgageAccountAndLoan"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Dim objIADOAssist As IADOAssist
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    Dim xmlMortLoan As IXMLDOMNode
    Dim xmlMortLoanGuid As IXMLDOMNode
    Dim xmlTempMortLoanAcc As IXMLDOMNode
    Dim xmlchkMortLoan As IXMLDOMNode
    Dim xmlGUIDNode As IXMLDOMNode
    Dim strAccountGUID As String
    Dim strKey As String
    Dim blnRecordExists As Boolean
    Dim objIClassDef As ICFClassDef
    Dim botOperation As BOOPERATIONTYPE
    Dim xmlTempAccountNode As IXMLDOMNode
    Dim xmlTempLoanNode As IXMLDOMNode
    Dim xmlTempGenKeysNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim xmlMortLoanKey As IXMLDOMElement
    Dim xmlMortAccountKey As IXMLDOMElement
    
    Set objSQLAssist = New SQLAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
    
    
    Set objIClassDef = New CFClassDef

    botOperation = m_objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadMortgageAccountData)

    Select Case botOperation
    
        Case booCreate:
            Set xmlMortAccountKey = vxmlGeneratedKeys.ownerDocument.createElement("MORTGAGEACCOUNT")
            vxmlGeneratedKeys.appendChild xmlMortAccountKey
            Set xmlTempAccountNode = ICustomerFinancialTxBO_CreateMortgageAccount(vxmlRequest, xmlMortAccountKey)
            m_objErrAssist.CheckResponse xmlTempAccountNode, xmlResponseElem, True
                        
        Case booUpdate:
            Set xmlTempAccountNode = ICustomerFinancialTxBO_UpdateMortgageAccount(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempAccountNode, xmlResponseElem, True
            
        Case booDelete:
            Set xmlTempLoanNode = ICustomerFinancialTxBO_DeleteMortgageLoan(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempLoanNode, xmlResponseElem, True
            Set xmlTempAccountNode = ICustomerFinancialTxBO_DeleteMortgageAccount(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempAccountNode, xmlResponseElem, True
            
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeMissingPrimaryTag, _
                                      "Invalid XML format"
    End Select
    ' BG 27/06/00 If have performed delete above do not need to do anything with MortgageLoan, as have just deleted it.
    If botOperation <> booDelete Then
        botOperation = m_objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadMortgageLoanData)
        
        
        Select Case botOperation
        
            Case booCreate:
                Set xmlMortLoan = xmlElement.selectSingleNode("./MORTGAGELOAN")
                Set xmlMortLoanGuid = xmlMortLoan.selectSingleNode(".//ACCOUNTGUID")
                
                If xmlMortLoanGuid Is Nothing Then
                    Set xmlMortLoanGuid = xmlMortLoan.ownerDocument.createElement("ACCOUNTGUID")
                    xmlMortLoan.appendChild xmlMortLoanGuid
                End If
                
                
                strAccountGUID = xmlMortLoanGuid.Text
                
                If strAccountGUID = "" Then
                
                    Set xmlKeysNode = xmlMortAccountKey.selectSingleNode("./ACCOUNTGUID")
                    
                    If Not (xmlKeysNode Is Nothing) Then
                        
                        xmlMortLoanGuid.Text = xmlKeysNode.Text
                    Else
                        m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeMissingPrimaryTag, _
                                      "ACCOUNTGUID tag missing"
                    End If
                End If
                'BG 27/06/00 Add MorgageLoan tag to contain MortgageLoanGuid when generated.
                Set xmlMortLoanKey = vxmlGeneratedKeys.ownerDocument.createElement("MORTGAGELOAN")
                vxmlGeneratedKeys.appendChild xmlMortLoanKey
                Set xmlTempLoanNode = ICustomerFinancialTxBO_CreateMortgageLoan(vxmlRequest, xmlMortLoanKey)
                m_objErrAssist.CheckResponse xmlTempLoanNode, xmlResponseElem, True
                If Not xmlKeysNode Is Nothing Then
                    xmlMortLoanKey.appendChild xmlKeysNode.cloneNode(True)
                End If
                
            Case booUpdate:
                Set xmlTempLoanNode = ICustomerFinancialTxBO_UpdateMortgageLoan(vxmlRequest)
                m_objErrAssist.CheckResponse xmlTempLoanNode, xmlResponseElem, True
                
            Case booDelete:
                Set xmlTempLoanNode = ICustomerFinancialTxBO_DeleteMortgageLoan(vxmlRequest)
                m_objErrAssist.CheckResponse xmlTempLoanNode, xmlResponseElem, True
            End Select
    End If
    
    Set ICustomerFinancialTxBO_SaveMortgageAccountAndLoan = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveMortgageAccountAndLoanExit:
    
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    Set xmlMortLoan = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlMortLoanGuid = Nothing
    Set xmlTempMortLoanAcc = Nothing
    Set xmlchkMortLoan = Nothing
    Set objIClassDef = Nothing
    Set xmlTempAccountNode = Nothing
    Set xmlTempLoanNode = Nothing
    Set xmlTempGenKeysNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set objIClassDef = Nothing
    Set xmlMortLoanKey = Nothing
    Set xmlMortAccountKey = Nothing
    
    Exit Function

SaveMortgageAccountAndLoanVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_SaveMortgageAccountAndLoan = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMortgageAccountAndLoanExit

End Function

Private Function ICustomerFinancialTxBO_UpdateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateRegularOutgoings multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateRegularOutgoingsVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateRegularOutgoings"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "REGULAROUTGOINGS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("REGULAROUTGOINGS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "REGULAROUTGOINGS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.UpdateRegularOutgoings xmlElement
    Else
        ' For each relevant node call the data object UpdateRegularOutgoings
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.UpdateRegularOutgoings xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateRegularOutgoings = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateRegularOutgoingsExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateRegularOutgoingsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateRegularOutgoings = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateRegularOutgoingsExit

End Function

Private Function ICustomerFinancialTxBO_UpdateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateMortgageLoan multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageLoanVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateMortgageLoan"

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
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "MORTGAGELOAN" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MORTGAGELOAN")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGELOAN" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerFinancialDO = New CustomerFinancialDO
    Else
        Set objICustomerFinancialDO = objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerFinancialDO.UpdateMortgageLoan xmlElement
    Else
        ' For each relevant node call the data object UpdateMortgageLoan
        For Each xmlElement In xmlNodeList
            objICustomerFinancialDO.UpdateMortgageLoan xmlElement
        Next
    End If
    
    Set ICustomerFinancialTxBO_UpdateMortgageLoan = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageLoanExit:
    
    Set objICustomerFinancialDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateMortgageLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateMortgageLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageLoanExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing

End Sub

Private Function ICustomerFinancialTxBO_CreateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create instances of Account Relationship
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateAccountRelationships"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlAccountRel As IXMLDOMNode
    Dim xmlAccountRelList As IXMLDOMNode
    Dim xmlCustNo As IXMLDOMNode
    Dim xmlCustVersionNo As IXMLDOMNode
    Dim xmlAccountGuid As IXMLDOMNode
    Dim xmlCustRoleType As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim objComboDO As ComboDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrACCOUNT_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
   
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
        
    ' Set up DO request
    ' Check we have a customerlist with som customers
    Set xmlCustomerList = xmlElement.selectNodes("CUSTOMERLIST/CUSTOMER")
    
    If xmlCustomerList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "CUSTOMERLIST/CUSTOMER"
    End If
    
    ' Get applicant value id
    Dim strAppValueId As String
    Dim strCustRole As String
    Dim strGuid As String
    
    strGuid = m_objXmlAssist.GetMandatoryElementText(xmlElement, "ACCOUNTGUID")
    Set xmlAccountRelList = xmlOut.createElement("ACCOUNTRELATIONSHIPLIST")
    
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strAppValueId = objComboDO.GetFirstComboValueId("CustomerRoleType", "A")
    
    ' For each customer passed in update the DO request and call the method
    For Each xmlCustomer In xmlCustomerList
        Set xmlAccountRel = xmlOut.createElement("ACCOUNTRELATIONSHIP")
        
        Set xmlCustNo = xmlOut.createElement("CUSTOMERNUMBER")
        xmlAccountRel.appendChild xmlCustNo
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomer, "CUSTOMERNUMBER")
        xmlCustNo.Text = xmlNode.Text
        
        Set xmlCustVersionNo = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlAccountRel.appendChild xmlCustVersionNo
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        xmlCustVersionNo.Text = xmlNode.Text
        
        Set xmlAccountGuid = xmlOut.createElement("ACCOUNTGUID")
        xmlAccountRel.appendChild xmlAccountGuid
        xmlAccountGuid.Text = strGuid
        
        Set xmlCustRoleType = xmlOut.createElement("CUSTOMERROLETYPE")
        xmlAccountRel.appendChild xmlCustRoleType
        strCustRole = m_objXmlAssist.GetNodeValue(xmlCustomer, "CUSTOMERROLETYPE", False)
        
        If Len(strCustRole) > 0 Then
            xmlCustRoleType.Text = strCustRole
        Else
            xmlCustRoleType.Text = strAppValueId
        End If
        
        xmlAccountRelList.appendChild xmlAccountRel
    Next
    
    objICustomerFinancialDO.CreateAccountRelationships xmlAccountRelList

    Set ICustomerFinancialTxBO_CreateAccountRelationships = xmlResponseElem
    
    m_objContext.SetComplete
    
CreateAccountRelationshipsExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlAccountRel = Nothing
    Set xmlCustNo = Nothing
    Set xmlCustVersionNo = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlCustRoleType = Nothing
    Set xmlNode = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerList = Nothing
    Set objComboDO = Nothing

    
    
    Exit Function

CreateAccountRelationshipsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateAccountRelationshipsExit

End Function

Private Function ICustomerFinancialTxBO_DeleteAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Account Relationship
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteAccountRelationships"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlCopyNode As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrACCOUNT_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
   
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    
    
    Set xmlCopyNode = xmlElement.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlCopyNode, "ACCOUNT", "ACCOUNTRELATIONSHIP"
    
    objICustomerFinancialDO.DeleteAccountRelationships xmlCopyNode
    
    Set ICustomerFinancialTxBO_DeleteAccountRelationships = xmlResponseElem
    
    m_objContext.SetComplete
    
DeleteAccountRelationshipsExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCopyNode = Nothing
        
    
    Exit Function

DeleteAccountRelationshipsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteAccountRelationshipsExit

End Function

Private Function ICustomerFinancialTxBO_UpdateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instances of Account Relationship
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountRelationshipsVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateAccountRelationships"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlAccountRel As IXMLDOMNode
    Dim xmlCustNo As IXMLDOMNode
    Dim xmlCustVersionNo As IXMLDOMNode
    Dim xmlAccountGuid As IXMLDOMNode
    Dim xmlCustRoleType As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrACCOUNT_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrACCOUNT_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrACCOUNT_ROOT_NODE_NAME & " tag not found"
    End If
   
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
        
    ' Set up DO request
    Set xmlAccountRel = xmlOut.createElement("ACCOUNTRELATIONSHIP")
    Set xmlCustNo = xmlOut.createElement("CUSTOMERNUMBER")
    xmlAccountRel.appendChild xmlCustNo
    Set xmlCustVersionNo = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlAccountRel.appendChild xmlCustVersionNo
    Set xmlAccountGuid = xmlOut.createElement("ACCOUNTGUID")
    xmlAccountRel.appendChild xmlAccountGuid
    Set xmlCustRoleType = xmlOut.createElement("CUSTOMERROLETYPE")
    xmlAccountRel.appendChild xmlCustRoleType
    
    Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlElement, "ACCOUNTGUID")
    xmlAccountGuid.Text = xmlNode.Text
    
    ' Check we have a customerlist with som customers
    Set xmlCustomerList = xmlElement.selectNodes("CUSTOMERLIST/CUSTOMER")
    
    If xmlCustomerList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "CUSTOMERLIST/CUSTOMER"
    End If
    
    ' For each customer passed in update the DO request and call the method
    For Each xmlCustomer In xmlCustomerList
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomer, "CUSTOMERNUMBER")
        xmlCustNo.Text = xmlNode.Text
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        xmlCustVersionNo.Text = xmlNode.Text
        Set xmlNode = m_objXmlAssist.GetMandatoryNode(xmlCustomer, "CUSTOMERROLETYPE")
        xmlCustRoleType.Text = xmlNode.Text
        objICustomerFinancialDO.UpdateAccountRelationships xmlAccountRel
    Next
    
    Set ICustomerFinancialTxBO_UpdateAccountRelationships = xmlResponseElem
    
    m_objContext.SetComplete
    
UpdateAccountRelationshipsExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlAccountRel = Nothing
    Set xmlCustNo = Nothing
    Set xmlCustVersionNo = Nothing
    Set xmlAccountGuid = Nothing
    Set xmlCustRoleType = Nothing
    Set xmlNode = Nothing
    Set xmlCustomer = Nothing
    Set xmlCustomerList = Nothing

    
    
    Exit Function

UpdateAccountRelationshipsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateAccountRelationships = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateAccountRelationshipsExit

End Function

Private Function ICustomerFinancialTxBO_CreateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create instances of Mortgage Account Special Feature
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_CreateSpecialFeature"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    Dim xmlKeys As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrFEATURE_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrFEATURE_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrFEATURE_ROOT_NODE_NAME & " tag not found"
    End If
       
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    
    Set xmlKeys = xmlOut.createElement("GENERATEDKEYS")
    
    objICustomerFinancialDO.CreateSpecialFeature xmlElement, xmlKeys
    
    If xmlKeys.hasChildNodes Then
        xmlResponseElem.appendChild xmlKeys
    End If
    
    Set ICustomerFinancialTxBO_CreateSpecialFeature = xmlResponseElem
    
    m_objContext.SetComplete
    
CreateSpecialFeatureExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeys = Nothing
    
    
    Exit Function

CreateSpecialFeatureVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_CreateSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateSpecialFeatureExit

End Function

Private Function ICustomerFinancialTxBO_DeleteSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete instances of Mortgage Account Special Feature
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_DeleteSpecialFeature"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrFEATURE_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrFEATURE_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrFEATURE_ROOT_NODE_NAME & " tag not found"
    End If
   
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    
    objICustomerFinancialDO.DeleteSpecialFeature xmlElement
    
    Set ICustomerFinancialTxBO_DeleteSpecialFeature = xmlResponseElem
    
    m_objContext.SetComplete
    
DeleteSpecialFeatureExit:

    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    
    Exit Function

DeleteSpecialFeatureVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_DeleteSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeleteSpecialFeatureExit

End Function

Private Function ICustomerFinancialTxBO_UpdateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instances of Mortgage Account Special Feature
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSpecialFeatureVbErr

    Const strFunctionName As String = "ICustomerFinancialTxBO_UpdateSpecialFeature"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerFinancialDO As ICustomerFinancialDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrFEATURE_ROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrFEATURE_ROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrFEATURE_ROOT_NODE_NAME & " tag not found"
    End If
   
    Set objICustomerFinancialDO = m_objContext.CreateInstance(App.Title & ".CustomerFinancialDO")
    
    objICustomerFinancialDO.UpdateSpecialFeature xmlElement
    
    Set ICustomerFinancialTxBO_UpdateSpecialFeature = xmlResponseElem
    
    m_objContext.SetComplete
    
UpdateSpecialFeatureExit:
    
    Set objICustomerFinancialDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    
    Exit Function

UpdateSpecialFeatureVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerFinancialTxBO_UpdateSpecialFeature = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateSpecialFeatureExit

End Function


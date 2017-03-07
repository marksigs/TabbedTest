VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CustomerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Customer Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String
'
'Dependencies:  List any other dependent components
'               e.g. CustomerDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     30/06/99    Created
'JLD    30/11/99    CustomerDO.UpdatePersonalDetails requires CUSTOMER and CUSTOMERVERSION
'                    blocks, so send the whole XML string
'SR     09/12/99    Modified methods SaveCustomerAddress and CreateCustomerAddress so that
'                   they return the key values generated
'IW     25/02/00    SYS0327 - Modified SaveCustomerTelephone to return success xml even when
'                   there is nothing to save
''SR    23/03/00    New Method 'CreateFullCustomerDetails'.
'SR    02/05/00     New method UpdateCustomerVersin
'SR     26/05/00    New method 'CreateCustomerDetailsAndEmpt'
'SR     26/05/00    Method 'UpdatePersonalDetailsAndEmpt' is complete
'BG     01/06/00    Method 'SavePersonalDetailsAndAddress' added
'MC     20/06/00    Amended CreateFullCustomerDetails to add new Import data
'JLD    27/06/00    don't create a customerlock if global param 'Locking' is 0
'SR     03/07/00    New method - Save Tenancy
'MC     07/07/00    SYS1168 Allow SaveDependant to update the 'Dependant Relative Of' key field
'BG     26/07/00    Added CreatePropertyInsuranceDetails
'BG     27/07/00    Added GetPropertyInsuranceDetails method
'SR     28/07/00    New method SaveDependant - to cater to saving multiple records
'SR     01-08-00    Modified UpdateCustomerAddress - tenancy validation now comapres with either 'TF' or 'TU'.
'                   Delete all the previous address records, based on the global parameter
'                   'AddressValidationYears' and optional tag 'DELETEPREVIOUSADDRESSES' in the request
' JLD   04/08/00    SYS1372 - allow saveDependant to work on one DEPENDANT node passed in
' JLD   09/08/00    SYS1402 - DeletePreviousCustomerAddresses re-visited. No longer look for the
'                   global parameter in the method. The GUI does that and just passes 'DELETEPREVIOUSADDRESSES'
'                   if the previous addresses need deleting.
'SR     15/08/00    SYS1412 : modified SaveTenancy - include the functionality for deletion of tenancy
'SR     01/09/00    Modified DeleteCustomerAddress - Added an additional input param. This is passed to DO method
'CL     21/12/00    SYS1755 Changes made to implement Contact History
'SR     05/06/01    SYS2325 New method 'UpdateCustomer'
'JLD    23/10/01    SYS2839 Check for optional XML node 'LOCKING' in CreateCustomerDetails.
'HMA    09/11/04    BM0174 Do not delete the tenancy records in UpdateCustomerAddress
'------------------------------------------------------------------------------------------
'BMids Specific History:
'
'Prog   Date        Description
'MDC    13/05/2002  BMIDS00004 - Amend CreateCustomerDetails to enable creation of new Customer Versions
'MV     12/08/2002  BMIDS00311 - Core Ref : SYS1705 Removed SafeRef ;Modified ICustomerTxBO_SavePersonalDetailsAndAddress
'MV     13/08/2002  BMIDS00331 - Core Ref : SYS2372 Amended ICustomerTxBO_CreateCustomerDetails
'GHun   15/08/2002  BMIDS00006 - CAWP1 Customer Account Download
'GHun   01/10/2002  BMIDS00443 - CreateCustomerDetails: use Omiga 3 Customer number if available
'HMA    12/08/2004  BMIDS836   - Updated CreateNewCustomerVersion to set NewToECustomerInd
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'HM     08/09/2005  MAR29   ICustomerTxBO_SaveWrapUpDetails were added
'PSC    17/10/2005  MAR57 - Amend CreateCustomerDetails
'HMA    20/10/2005  MAR23   Added CreateKYCCheck
'PSC    09/01/2006  MAR1001 Amend CreateCustomerDetails to get OTHERSYSTEMCUSTOMERNUMBER correctly
'------------------------------------------------------------------------------------------
'EPSOM Specific History:
'
'Prog   Date        Description
'IK     17/05/2006  EP561   add IsCheckDigitRequired test
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerTxBO
Implements ObjectControl

 '=============================================
 'Constant Declaration Section
 '=============================================

    Private Const cstrROOT_NODE_NAME = "CUSTOMER"
    Private Const cstrALIASTABLE_NAME = "ALIAS"
    Private Const cstrTHIRDPARTY_COMPONENT = "omTP"
    Private Const cstrCUSTEMPLOYMENT_COMPONENT = "omCE"
    Private Const cstrBASE_COMPONENT = "omBase"
    Private Const cstrTENANCYTABLE_NAME = "TENANCY"
    Private m_objContext As ObjectContext
    Private m_objErrAssist As ErrAssist
    Private m_objXmlAssist As XMLAssist             ' Remove if not required
    Private Const cstrROOT_CONTACTHISTORY = "CONTACTHISTORY"

Private Function ICustomerTxBO_CreateAlias(ByVal vxmlRequestElement As IXMLDOMElement, _
                                           ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_CreateAliasVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement

    Dim objICustomerDO As ICustomerDO

    Set xmlResponseElem = vxmlRequestElement.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vxmlRequestElement.getElementsByTagName("ALIAS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ALIAS" & " tag not found"
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    Dim intNodeIndex As Integer

    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        Set xmlTempResponseElem = objICustomerDO.CreateAlias(objElement)
        
        objErrAssist.CheckResponse xmlTempResponseElem, xmlResponseElem, True
        vxmlGeneratedKeys.appendChild xmlTempResponseElem.selectSingleNode("ALIASSEQUENCENUMBER")
    Next

    Set ICustomerTxBO_CreateAlias = xmlResponseElem.cloneNode(True)
       
   If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
   End If
  
ICustomerTxBO_CreateAliasExit:
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    Exit Function

ICustomerTxBO_CreateAliasVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_CreateAliasExit
    
End Function

Private Function ICustomerTxBO_CreateCustomerAddress( _
                            ByVal vXMLRequest As IXMLDOMElement, _
                            ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest   xml Request node containing data to be persisted
'               vxmlGeneratedKeys : node to attach the generated keys
' return:       xmlResponse node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_CreateCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateCustomerAddress"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

'    Dim xmlOut as DOMDocument
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERADDRESS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "CUSTOMERADDRESS" & " tag not found"
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
   
    Dim intNodeIndex As Integer

    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        Set xmlNode = objICustomerDO.CreateCustomerAddress(objElement)
        
        ' Add the response from DO to GeneratedKeys node
        vxmlGeneratedKeys.appendChild xmlNode
    Next
    
    Set ICustomerTxBO_CreateCustomerAddress = xmlResponseElem.cloneNode(True)

ICustomerTxBO_CreateCustomerAddressExit:

    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set objElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerTxBO_CreateCustomerAddressVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_CreateCustomerAddressExit
    
End Function

Private Function ICustomerTxBO_CreateFullCustomerDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateFullCustomerDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlEmpElement As IXMLDOMElement
    Dim xmlCustomerOutElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objICustomerBO As ICustomerBO
    Dim objCustomerBO As CustomerBO
'    Dim objICustEmploymentBO As ICustomerEmploymentBO
    Dim objCustEmploymentBO As CustomerEmploymentBO
    Dim objCustFinBO As CustomerFinancialBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
   
    If objContext Is Nothing Then
        Set objICustomerBO = New CustomerBO
        Set objCustomerBO = New CustomerBO
        Set objCustEmploymentBO = New CustomerEmploymentBO
        Set objCustFinBO = New CustomerFinancialBO
    Else
        Set objICustomerBO = objContext.CreateInstance(App.Title & ".CustomerBO")
        Set objCustomerBO = objContext.CreateInstance(App.Title & ".CustomerBO")
        Set objCustEmploymentBO = objContext.CreateInstance(gstrCUSTOMEREMPLOYMENT_COMPONENT & ".CustomerEmploymentBO")
        Set objCustFinBO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
    
    '   Create Customer Details and add retreive the keys
    Dim strResponse As String, strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    
    strResponse = objCustomerBO.CreateCustomerDetails(vXMLRequest.xml)
    Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    
    strCustomerNumber = _
            objXmlAssist.GetTagValue(xmlTempDoc.documentElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = _
        objXmlAssist.GetTagValue(xmlTempDoc.documentElement, "CUSTOMERVERSIONNUMBER")
    
    
    '   Create Alias Record
    Set xmlTempResponseNode = xmlOut.createElement("GENERATEDKEYS")
    Set xmlElement = vXMLRequest.selectSingleNode(".//ALIAS")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "ALIAS", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        Set xmlTempResponseNode = objICustomerBO.ValidateAlias(vXMLRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        Set xmlTempResponseNode = objICustomerBO.CreateAlias(vXMLRequest, xmlTempResponseNode)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
                
    ' Save Employment
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "EMPLOYMENT", _
                                strCustomerNumber, strCustomerVersionNumber)
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "EMPLOYEDDETAILS", _
                            strCustomerNumber, strCustomerVersionNumber, False)
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "EARNEDINCOME", _
                            strCustomerNumber, strCustomerVersionNumber, False)
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "NETPROFIT", _
                            strCustomerNumber, strCustomerVersionNumber, False)
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "SELFEMPLOYEDDETAILS", _
                            strCustomerNumber, strCustomerVersionNumber, False)
    'MC 20/06/2000. If it exists copy NetProfit element as a child of SelfEmployedDetails
    Set xmlEmpElement = vXMLRequest.selectSingleNode(".//EMPLOYMENT")
    Set xmlElement = vXMLRequest.selectSingleNode(".//SELFEMPLOYEDDETAILS")
    If Not (xmlElement Is Nothing Or xmlEmpElement Is Nothing) Then
        objXmlAssist.CopyNode "NETPROFIT", "NETPROFIT", xmlEmpElement, xmlElement, True
    End If
    
    strResponse = objCustEmploymentBO.SaveEmploymentDetails(vXMLRequest.xml)
    Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    
    ' Save Other Income Details
    Call SetInputForCreateFullCustomerDetails(vXMLRequest, "UNEARNEDINCOME", _
                                strCustomerNumber, strCustomerVersionNumber)
    strResponse = objCustEmploymentBO.SaveOtherIncome(vXMLRequest.xml)
    Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    
    ' Save LoansLiabilities
    Set xmlElement = vXMLRequest.selectSingleNode(".//LOANSLIABILITIES")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "LOANSLIABILITIES", _
                                strCustomerNumber, strCustomerVersionNumber)
        
        strResponse = objCustFinBO.CreateLiability(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    ' Save BankCreditCard details
    Set xmlElement = vXMLRequest.selectSingleNode(".//BANKCREDITCARD")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "BANKCREDITCARD", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        strResponse = objCustFinBO.CreateBankCard(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    'MC 14/06/2000. Save Regular Outgoings details
    Set xmlElement = vXMLRequest.selectSingleNode(".//REGULAROUTGOINGS")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "REGULAROUTGOINGS", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        strResponse = objCustFinBO.CreateRegularOutgoings(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    'MC 14/06/2000. Save Arrears History details
    Set xmlElement = vXMLRequest.selectSingleNode(".//ARREARSHISTORY")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "ARREARSHISTORY", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        strResponse = objCustFinBO.CreateArrearsHistory(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    'MC 14/06/2000. Save Bankruptcy History details
    Set xmlElement = vXMLRequest.selectSingleNode(".//BANKRUPTCYHISTORY")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "BANKRUPTCYHISTORY", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        strResponse = objCustFinBO.CreateBankruptcyHistory(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    'MC 14/06/2000. Save CCJ History details
    Set xmlElement = vXMLRequest.selectSingleNode(".//CCJHISTORY")
    If Not xmlElement Is Nothing Then
        Call SetInputForCreateFullCustomerDetails(vXMLRequest, "CCJHISTORY", _
                                strCustomerNumber, strCustomerVersionNumber)
                                
        strResponse = objCustFinBO.CreateCCJHistory(vXMLRequest.xml)
        Set xmlTempDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlTempDoc.documentElement, xmlResponseElem, True
    End If
    
    ' Add Customer Keys to the response element
    Set xmlCustomerOutElem = xmlOut.createElement("CUSTOMER")
    xmlResponseElem.appendChild xmlCustomerOutElem
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCustomerNumber
    xmlCustomerOutElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCustomerVersionNumber
    xmlCustomerOutElem.appendChild xmlElement
    
    Set ICustomerTxBO_CreateFullCustomerDetails = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTempDoc = Nothing
    Set objICustomerBO = Nothing
    Set objCustomerBO = Nothing
    Set objCustEmploymentBO = Nothing
    Set objCustFinBO = Nothing
    Set objErrAssist = Nothing
    
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

    Set ICustomerTxBO_CreateFullCustomerDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Sub SetInputForCreateFullCustomerDetails( _
                ByVal vXMLRequest As IXMLDOMElement, _
                ByVal strTableName As String, _
                ByVal strCustNumber As String, ByVal strCustVersionNumber As String, _
                Optional ByVal bErrorIfTagNotFound As Boolean = True)

On Error GoTo ErrHandle
    
    Const strFunctionName As String = "SetInputForCreateFullCustomerDetails"
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlCustNumberElem As IXMLDOMElement
    Dim xmlCustVersionNumberElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As ErrAssist
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Set objErrAssist = New ErrAssist
    
    Set xmlNodeList = vXMLRequest.selectNodes(".//" & strTableName)
    If xmlNodeList.length = 0 Then
        If bErrorIfTagNotFound Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, strTableName & " tag not found"
        Else
            'Exit subroutine without raising an error
            Exit Sub
        End If
    Else
        For Each xmlElement In xmlNodeList
    
            Set xmlCustNumberElem = xmlElement.selectSingleNode(".//CUSTOMERNUMBER")
            If xmlCustNumberElem Is Nothing Then
                Set xmlCustNumberElem = xmlOut.createElement("CUSTOMERNUMBER")
            End If
            xmlCustNumberElem.Text = strCustNumber
            
            Set xmlCustVersionNumberElem = xmlElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
            If xmlCustVersionNumberElem Is Nothing Then
                Set xmlCustVersionNumberElem = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
            End If
            xmlCustVersionNumberElem.Text = strCustVersionNumber
            
        Next
    End If
    
    Set xmlElement = Nothing
    Set xmlCustNumberElem = Nothing
    Set xmlCustVersionNumberElem = Nothing
    Set xmlOut = Nothing
    Set xmlNodeList = Nothing
    Exit Sub

ErrHandle:
    Set xmlElement = Nothing
    Set xmlCustNumberElem = Nothing
    Set xmlCustVersionNumberElem = Nothing
    Set xmlOut = Nothing
    Set xmlNodeList = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ICustomerTxBO_CreateCustomerDetailsAndEmpt( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateCustomerDetailsAndEmpt"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlElement As IXMLDOMElement, xmlElem2 As IXMLDOMElement
    Dim xmlEmloymentNode As IXMLDOMNode, xmlEmployedDetailsNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlGeneratedKeys As IXMLDOMNode
    Dim xmlCustomerNumberNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
   
    Dim objCustEmpBO As CustomerEmploymentBO
   
    Dim strCustomerNumber As String, strCustVersionNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlGeneratedKeys = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlGeneratedKeys
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = "CUSTOMER" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("CUSTOMER").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  " CUSTOMER" & " tag not found"
    End If
       
    ' Save Customer details and the numbers generated into local variables
    ICustomerTxBO_CreateCustomerDetails vXMLRequest, xmlGeneratedKeys
    strCustomerNumber = xmlGeneratedKeys.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustVersionNumber = xmlGeneratedKeys.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    
    ' Save employment details
    If objContext Is Nothing Then
        Set objCustEmpBO = New CustomerEmploymentBO
    Else
        Set objCustEmpBO = _
            objContext.CreateInstance(cstrCUSTEMPLOYMENT_COMPONENT & ".CustomerEmploymentBO")
    End If
    
    ' Append nodes of Primary Key fields to tags EMPLOYMENT and EMPLOYEDDETAILS and assign
    ' values for customer number and customer version number
    Set xmlEmloymentNode = vXMLRequest.selectSingleNode(".//EMPLOYMENT")
    Set xmlEmployedDetailsNode = vXMLRequest.selectSingleNode(".//EMPLOYEDDETAILS")
    
    Set xmlElem2 = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElem2.Text = strCustomerNumber
    xmlEmloymentNode.appendChild xmlElem2.cloneNode(True)
    xmlEmployedDetailsNode.appendChild xmlElem2
    
    Set xmlElem2 = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElem2.Text = strCustVersionNumber
    xmlEmloymentNode.appendChild xmlElem2.cloneNode(True)
    xmlEmployedDetailsNode.appendChild xmlElem2
    
    Set xmlElem2 = xmlOut.createElement("EMPLOYMENTSEQUENCENUMBER")
    xmlEmloymentNode.appendChild xmlElem2.cloneNode(True)
    xmlEmployedDetailsNode.appendChild xmlElem2
    
    Dim strResponse As String
    strResponse = objCustEmpBO.SaveEmploymentDetails(vXMLRequest.xml)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'Save Employed details
    'Assign value to EmploymetSequenceNumber from values generated above
    xmlEmployedDetailsNode.selectSingleNode("EMPLOYMENTSEQUENCENUMBER").Text = _
            xmlDoc.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
    
    strResponse = objCustEmpBO.SaveEmployedDetails(vXMLRequest.xml)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'Append the generated customer keys to Response Element and return
    xmlResponseElem.appendChild xmlGeneratedKeys
    Set ICustomerTxBO_CreateCustomerDetailsAndEmpt = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set xmlElement = Nothing
    Set xmlElem2 = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set objErrAssist = Nothing
    Set xmlGeneratedKeys = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objCustEmpBO = Nothing
    
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

    Set ICustomerTxBO_CreateCustomerDetailsAndEmpt = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

'BMIDS00004 MDC 17/05/2002
Private Function ICustomerTxBO_CreateNewCustomerVersion(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateNewCustomerVersionVbErr
    
Const strFunctionName As String = "ICustomerTxBO_CreateNewCustomerVersion"

Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequestNode As IXMLDOMNode
Dim xmlCreateNode As IXMLDOMNode
Dim xmlCustomerNode As IXMLDOMNode
Dim xmlCustomerAddressGuidNodeList As IXMLDOMNodeList
Dim xmlGUIDNode As IXMLDOMNode
Dim xmlCustomerAddressList As IXMLDOMNode
Dim xmlKeysNode As IXMLDOMNode
Dim xmlEmpList As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim objCustEmpBO As CustomerEmploymentBO
Dim objICustBO As ICustomerBO
Dim xmlTempResponseNode As IXMLDOMNode

Dim strResponse As String
Dim strNewCustomerVersion As String
Dim strExistingCustomerVersion As String
Dim strCustomerNumber As String
Dim lngErrNo As Long

'BMIDS836
Dim strToEInd As String
Dim xmlCustomerToENode As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objCustEmpBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    Set objICustBO = m_objContext.CreateInstance(App.Title & ".CustomerBO")
    
    'Initialise the response
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strCustomerNumber = xmlGetMandatoryNodeText(vXMLRequest, ".//CUSTOMERNUMBER")
    strExistingCustomerVersion = xmlGetMandatoryNodeText(vXMLRequest, ".//CUSTOMERVERSIONNUMBER")
    
    'BMIDS836 Get ToE indicator passed in the Request.
    strToEInd = xmlGetNodeText(vXMLRequest, ".//NEWTOECUSTOMERIND")
        
    Set xmlTempResponseNode = xmlDoc.createElement("RESPONSE")
    
    'Get Customer data
    objICustBO.GetCustomerDetails vXMLRequest, xmlTempResponseNode
    errCheckXMLResponseNode xmlTempResponseNode, xmlResponseElem, True
    Set xmlCustomerNode = xmlGetMandatoryNode(xmlTempResponseNode, "CUSTOMER")
    
    'Remove AddressGuid from CustomerAddress node in preparation for creating a new customer version
    Set xmlCustomerAddressList = xmlGetNode(xmlCustomerNode, "CUSTOMERVERSION/CUSTOMERADDRESSLIST")
    If Not xmlCustomerAddressList Is Nothing Then
        Set xmlCustomerAddressGuidNodeList = xmlCustomerAddressList.selectNodes(".//ADDRESSGUID")
        For Each xmlGUIDNode In xmlCustomerAddressGuidNodeList
            xmlGUIDNode.Text = ""
        Next
    End If
    
    'BMIDS836 Set ToE indicator
    Set xmlCustomerToENode = xmlGetNode(xmlCustomerNode, "CUSTOMERVERSION/NEWTOECUSTOMERIND")
    If Not xmlCustomerToENode Is Nothing Then
        xmlCustomerToENode.Text = strToEInd
    End If
        
    'Build request for CreateCustomerDetails
    Set xmlTempRequestNode = vXMLRequest.cloneNode(False)
    Set xmlCreateNode = xmlDoc.createElement("CREATE")
    xmlTempRequestNode.appendChild xmlCreateNode
    xmlCreateNode.appendChild xmlCustomerNode
    'Do not need to create a lock as customer will already be locked when this method is called
    Set xmlTempNode = xmlDoc.createElement("LOCKING")
    xmlTempNode.Text = "0"
    xmlCreateNode.appendChild xmlTempNode
    
    'Create a new customer version
    ICustomerTxBO_CreateCustomerDetails xmlTempRequestNode, xmlTempResponseNode
    errCheckXMLResponseNode xmlTempResponseNode, xmlResponseElem, True

    'Save the keys of the created customer version for the response
    Set xmlKeysNode = xmlGetMandatoryNode(xmlTempResponseNode, "CUSTOMERKEY")
    xmlResponseElem.appendChild xmlKeysNode
   
    'Copy employment data to new customer version
    strNewCustomerVersion = xmlGetMandatoryNodeText(xmlKeysNode, "CUSTOMERVERSIONNUMBER")
    Set xmlTempRequestNode = vXMLRequest.cloneNode(False)
    Set xmlEmpList = xmlDoc.createElement("EMPLOYMENTLIST")
    xmlTempRequestNode.appendChild xmlEmpList
    
    Set xmlCustomerNode = xmlDoc.createElement("EXISTINGCUSTOMERVERSION")
    Set xmlTempNode = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlTempNode.Text = strCustomerNumber
    xmlCustomerNode.appendChild xmlTempNode
    Set xmlTempNode = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlTempNode.Text = strExistingCustomerVersion
    xmlCustomerNode.appendChild xmlTempNode
    xmlEmpList.appendChild xmlCustomerNode
    
    Set xmlCustomerNode = xmlDoc.createElement("NEWCUSTOMERVERSION")
    Set xmlTempNode = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlTempNode.Text = strCustomerNumber
    xmlCustomerNode.appendChild xmlTempNode
    Set xmlTempNode = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlTempNode.Text = strNewCustomerVersion
    xmlCustomerNode.appendChild xmlTempNode
    xmlEmpList.appendChild xmlCustomerNode
    
    strResponse = objCustEmpBO.CopyEmploymentDetails(xmlTempRequestNode.xml)
    errCheckXMLResponse strResponse, True
    
    'Set the response and commit the transaction
    Set ICustomerTxBO_CreateNewCustomerVersion = xmlResponseElem
    m_objContext.SetComplete
    
CreateNewCustomerVersionExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlCreateNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerAddressGuidNodeList = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlCustomerAddressList = Nothing
    Set xmlKeysNode = Nothing
    Set xmlEmpList = Nothing
    Set xmlTempNode = Nothing
    Set objCustEmpBO = Nothing
    Set objICustBO = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewCustomerVersionVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateNewCustomerVersion = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateNewCustomerVersionExit

End Function
'BMIDS00004 MDC 17/05/2002 - End

Private Function ICustomerTxBO_CreatePropertyInsuranceDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes ICustomerDO_CreatePropertyInsuranceDetails
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreatePropertyInsuranceDetails"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement

    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objICustomerDO.CreatePropertyInsuranceDetails xmlElement
    
    Set ICustomerTxBO_CreatePropertyInsuranceDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreatePropertyInsuranceDetailsExit:

    Set objICustomerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreatePropertyInsuranceDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePropertyInsuranceDetailsExit

End Function

Private Function ICustomerTxBO_DeleteAlias(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform delete
' pass:         vxmlRequestElem  xml Request node which identifies instance to be deleted
'               vxmlOutNode      xml node to which the output is to be appended
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_DeleteAliasVbErr

    Const strFunctionName As String = "ICustomerTxBO_DeleteAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("ALIAS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                oeMissingElement & " ALIAS tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        Call objICustomerDO.DeleteAlias(objElement)
    Next

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set ICustomerTxBO_DeleteAlias = xmlResponseElem

ICustomerTxBO_DeleteAliasExit:
    
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlResponseElem = Nothing
'    Set xmlOut = Nothing
    
    Exit Function

ICustomerTxBO_DeleteAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_DeleteAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_DeleteAliasExit
End Function

Private Function ICustomerTxBO_SaveDependant(ByVal vXMLRequest As IXMLDOMElement, _
                                ByVal vxmlGeneratedKeys As IXMLDOMElement, _
                                Optional ByRef rbotOperationType As Variant) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the dependant Data after finding the operation to be performed on it
' pass:         vstrXMLRequest  xml Request data to be passed to the methos
' return:       nil
' Error :      Invalid Paramenter
'------------------------------------------------------------------------------------------

On Error GoTo SaveDependantVbErr

    Const strFunctionName As String = "ICustomerTxBO_SaveDependant"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode, xmlCustomerElem As IXMLDOMNode
    
    Dim objIClassDef As IomCustClassDef
    Dim objICustomerDO As ICustomerDO
    Dim objICustomerBO As ICustomerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponseElem
    
    ' Get a list of nodes that need to be passed down to the data object
    If vXMLRequest.nodeName = "DEPENDANT" Then
        Set xmlNode = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.selectNodes(".//DEPENDANT")

        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                            oeMissingPrimaryTag, "ALIAS" & " tag not found"
        End If
    End If
    
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    ' For each relevant node, determine the operation to be performed and
    ' call the appropriate method from DO
    Set objIClassDef = New omCustClassDef
    Set xmlDoc = objIClassDef.LoadDependantData
    Set objICustomerBO = New CustomerBO
    
    Dim intNodeIndex As Integer
    Dim botOperationType As BOOPERATIONTYPE

    If Not (xmlNode Is Nothing) Then
        botOperationType = m_objXmlAssist.DetermineOperationEx(xmlNode, xmlDoc)
            
        If Not IsMissing(rbotOperationType) Then
            rbotOperationType = botOperationType
        End If
        
        Select Case botOperationType
            Case booCreate
                Set xmlTempResponseNode = objICustomerBO.ValidateDependant(xmlNode)
                m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                ' Create dependant record and add the generated sequence number to xmlGeneratedKeys
                Set xmlTempResponseNode = objICustomerDO.CreateDependant(xmlNode)
                
                Set xmlCustomerElem = xmlOut.createElement("CUSTOMER")
                vxmlGeneratedKeys.appendChild xmlCustomerElem
                
                Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
                xmlElement.Text = xmlNode.selectSingleNode(".//CUSTOMERNUMBER").Text
                xmlCustomerElem.appendChild xmlElement
                
                Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
                xmlElement.Text = xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
                xmlCustomerElem.appendChild xmlElement
                ' Dependant sequence number
                xmlCustomerElem.appendChild xmlTempResponseNode.cloneNode(True)
            Case booUpdate
                Set xmlTempResponseNode = objICustomerBO.ValidateDependant(xmlNode)
                m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                objICustomerDO.UpdateDependant xmlNode
            Case booDelete
                objICustomerBO.DeleteDependant xmlNode
            Case Else
                ' N/A
        End Select
    Else
        For intNodeIndex = 0 To xmlNodeList.length - 1
            Set xmlNode = xmlNodeList(intNodeIndex)
            botOperationType = m_objXmlAssist.DetermineOperationEx(xmlNode, xmlDoc)
            
            'SR : 28/07/00 : if the OperationType is passed as argument, assign operation type of the
            '                first Dependant to it. This was done by the method previously when it was
            '                catering to saving just one row.
            If (intNodeIndex = 0 And Not IsMissing(rbotOperationType)) Then
                rbotOperationType = botOperationType
            End If
            
            Select Case botOperationType
                Case booCreate
                    Set xmlTempResponseNode = objICustomerBO.ValidateDependant(xmlNode)
                    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                    ' Create dependant record and add the generated sequence number to xmlGeneratedKeys
                    Set xmlTempResponseNode = objICustomerDO.CreateDependant(xmlNode)
                    
                    Set xmlCustomerElem = xmlOut.createElement("CUSTOMER")
                    vxmlGeneratedKeys.appendChild xmlCustomerElem
                    
                    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
                    xmlElement.Text = xmlNode.selectSingleNode(".//CUSTOMERNUMBER").Text
                    xmlCustomerElem.appendChild xmlElement
                    
                    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
                    xmlElement.Text = xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
                    xmlCustomerElem.appendChild xmlElement
                    ' Dependant sequence number
                    xmlCustomerElem.appendChild xmlTempResponseNode.cloneNode(True)
                Case booUpdate
                    Set xmlTempResponseNode = objICustomerBO.ValidateDependant(xmlNode)
                    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                    objICustomerDO.UpdateDependant xmlNode
                Case booDelete
                    objICustomerBO.DeleteDependant xmlNode
                Case Else
                    ' N/A
            End Select
        Next intNodeIndex
    End If
    
    Set ICustomerTxBO_SaveDependant = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveDependantExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlElement = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlDoc = Nothing
    
    Set objIClassDef = Nothing
    Set objICustomerDO = Nothing
    Set objICustomerBO = Nothing
    
    Exit Function
SaveDependantVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SaveDependant = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveDependantExit

End Function

Private Function ICustomerTxBO_SavePersonalDetailsAndAddress(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  BG:01/06/00     Updates Customer details including email and address
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SavePersonalDetailsAndAddressVbErr

    Const strFunctionName As String = "ICustomerTxBO_SavePersonalDetailsAndAddress"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode
    
    Dim objICustomerBO As ICustomerBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
    
    If m_objContext Is Nothing Then
        Set objICustomerBO = New CustomerBO
    Else
        Set objICustomerBO = m_objContext.CreateInstance(App.Title & ".CustomerBO")
    End If
    
    Set xmlTempResponse = objICustomerBO.UpdateCustomerPersonalDetails(vXMLRequest)
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    
    Set xmlTempResponse = objICustomerBO.SaveCustomerAddress(vXMLRequest, vxmlGeneratedKeys)
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    
    Set ICustomerTxBO_SavePersonalDetailsAndAddress = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SavePersonalDetailsAndAddressExit:
    
    Set objICustomerBO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponse = Nothing
    
    Exit Function

SavePersonalDetailsAndAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SavePersonalDetailsAndAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SavePersonalDetailsAndAddressExit

End Function

Private Function ICustomerTxBO_SaveWrapUpDetails(ByVal vXMLRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveWrapUpDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing SaveCustomerTelephone data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_SaveWrapUpDetailsVbErr

    Const strFunctionName As String = "ICustomerTxBO_SaveWrapUpDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNodeList2 As IXMLDOMNodeList     ' PSC 18/10/2005 MAR57
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    Dim objIClassDef As IomCustClassDef
    Dim botOperationType As BOOPERATIONTYPE
    Dim xmlTempResponse As IXMLDOMNode

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERWRAPUPDETAILS")
    ' PSC 18/10/2005 MAR57
    Set xmlNodeList2 = vXMLRequest.getElementsByTagName("CUSTOMERVERSION")
    ' PSC 18/10/2005 MAR57
    If xmlNodeList.length = 0 And xmlNodeList2.length = 0 Then
        ' No need to raise error just bail out of the routine
        Set ICustomerTxBO_SaveWrapUpDetails = xmlResponseElem.cloneNode(True) 'SYS0327
        Exit Function
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    Set objIClassDef = New omCustClassDef
    
    ' For each relevant node that we are passed call the data object Create/Update
    ' PSC 18/10/2005 MAR57 - Start
    If xmlNodeList.length > 0 Then
        For Each xmlNode In xmlNodeList
            botOperationType = objXmlAssist.DetermineOperationEx(xmlNode, objIClassDef.LoadWrapUpDetailsData)
            Select Case botOperationType
                Case booUpdate: objICustomerDO.UpdateWrapUpDetails xmlNode
                Case booCreate: objICustomerDO.CreateWrapUpDetails xmlNode
                Case booDelete: objICustomerDO.DeleteWrapUpDetails xmlNode
                Case Else
                    ' N/A
            End Select
        Next
    End If
    
    'update customer version
    If xmlNodeList2.length > 0 Then
        For Each xmlNode In xmlNodeList2
            Set xmlTempResponse = ICustomerTxBO_UpdateCustomerVersion(xmlNode)
        Next
    End If
    ' PSC 18/10/2005 MAR57 - End


    errCheckXMLResponseNode xmlTempResponse, xmlResponseElem, True
    
    Set ICustomerTxBO_SaveWrapUpDetails = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_SaveWrapUpDetailsExit:
    Set objICustomerDO = Nothing
    Set objIClassDef = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList2 = Nothing  ' PSC 18/10/2005 MAR57



    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ICustomerTxBO_SaveWrapUpDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SaveWrapUpDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_SaveWrapUpDetailsExit

End Function

Private Function ICustomerTxBO_UpdateAlias(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequestElem :  XML Element containing update data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_UpdateAliasVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("ALIAS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                oeMissingElement & " ALIAS tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        Call objICustomerDO.UpdateAlias(objElement)
    Next

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set ICustomerTxBO_UpdateAlias = xmlResponseElem
    
ICustomerTxBO_UpdateAliasExit:
    
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerTxBO_UpdateAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdateAliasExit

End Function

Private Function ICustomerTxBO_DeleteCustomerAddress(ByVal vXMLRequest As IXMLDOMElement, _
                                    Optional ByVal blnDeleteMortgageAccount As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               Table 'CustomerAddress' and the corresponding record in 'Address'
'               invokes appropriate Data Object(s) to perform delete
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_DeleteCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerTxBO_DeleteCustomerAddress"

    Dim objSQLAssist As New SQLAssist
    Dim objIADOAssist As IADOAssist

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempTableElem As IXMLDOMElement
    Dim objElement As IXMLDOMElement

    Dim objICustomerDO As ICustomerDO
    Dim objIAddressDO As IAddressDO
    
    Dim xmlResponseElem As IXMLDOMElement
    
    'BMIDS00006
    Dim xmlDoc          As FreeThreadedDOMDocument40
    Dim xmlAddressGuid  As IXMLDOMNode
    'BMIDS00006 End

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERADDRESS")

    If xmlNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                oeMissingPrimaryTag, "CUSTOMERADDRESS tag not found"
    End If

    Dim intNodeIndex As Integer

    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
        Set objIAddressDO = New AddressDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
        Set objIAddressDO = m_objContext.CreateInstance(App.Title & ".AddressDO")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strCustomerAddrSeqNo As String, strCondition As String
    Dim strAddressGuid  As String, strNatureOfOccupancy As String
    Dim strTemp As String, strTenancyValidation As String
    Dim blnRecordFound As Boolean, blnTenancyTobeDeleted As Boolean

    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)

        'Find the values of CustomerNumber, CustomerVersionNumber, CustomerAddrSeqNo
        If Not objICustomerDO.bGetKeysOfCustomerAddress(objElement, strCustomerNumber, _
                                         strCustomerVersionNumber, strCustomerAddrSeqNo) Then

            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 519, _
                            "for CUSTOMERADDRESS"

        End If

        strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                      & " AND CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                      & " AND CUSTOMERADDRESSSEQUENCENUMBER = " & strCustomerAddrSeqNo

        'Find Nature of Occupancy for this address. If it is Tenancy (T), and a record
        'exist in 'Tenancy' table then delete the corresponding record.
        strNatureOfOccupancy = m_objXmlAssist.GetTagValue(vXMLRequest, "NATUREOFOCCUPANCY")
        blnTenancyTobeDeleted = False
        If strNatureOfOccupancy <> "" Then
            strTemp = " GROUPNAME = " & objSQLAssist.FormatString("NatureOfOccupancy") _
                      & " AND VALUEID = " & strNatureOfOccupancy
            Call objIADOAssist.GetValueFromTable("COMBOVALIDATION", strTemp, _
                                "VALIDATIONTYPE", strTenancyValidation, blnRecordFound)
            If blnRecordFound Then
                If strTenancyValidation = "TF" Or strTenancyValidation = "TU" Then
                    blnTenancyTobeDeleted = True
                End If
            End If
        End If
        If blnTenancyTobeDeleted Then
            If objIADOAssist.CheckRecordExists("TENANCY", strCondition) Then
                ' Add Tenancy Node and its keys to vxmlRequest and pass it for deletion
                Set xmlTempTableElem = vXMLRequest.ownerDocument.createElement("TENANCY")
                vXMLRequest.appendChild xmlTempTableElem

                Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERNUMBER")
                xmlElement.Text = strCustomerNumber
                xmlTempTableElem.appendChild xmlElement

                Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
                xmlElement.Text = strCustomerVersionNumber
                xmlTempTableElem.appendChild xmlElement

                Set xmlElement = vXMLRequest.ownerDocument.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
                xmlElement.Text = strCustomerAddrSeqNo
                xmlTempTableElem.appendChild xmlElement

                objICustomerDO.DeleteTenancy xmlTempTableElem
            End If
        End If

        'Find the AddressGuid required deletion of record from table 'Address'. First check
        'whether it is mentioned in input else get it from the table 'CustomerAddress'
        strAddressGuid = m_objXmlAssist.GetTagValue(vXMLRequest, "ADDRESSGUID")
        If strAddressGuid = "" Then
            If IsNull(objIADOAssist.GetValueFromTable("CUSTOMERADDRESS", strCondition, _
                                           "ADDRESSGUID", strAddressGuid)) = True Then
                 m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 209
            End If
            'BMIDS00006 Append the AddressGUID to the request as it is required in CustomerDO
            Set xmlDoc = New FreeThreadedDOMDocument40
            xmlDoc.validateOnParse = False
            xmlDoc.setProperty "NewParser", True
            Set xmlAddressGuid = xmlDoc.createElement("ADDRESSGUID")
            vXMLRequest.selectSingleNode("CUSTOMERADDRESS").appendChild xmlAddressGuid
            xmlAddressGuid.Text = strAddressGuid
            'BMIDS00006 END
        End If

        'Delete the record from the table 'CustomerAddress'
        Call objICustomerDO.DeleteCustomerAddress(vXMLRequest, blnDeleteMortgageAccount)

        'Add Address node and Key Element to Request node; pass it for deletion of address record
        Set xmlTempTableElem = vXMLRequest.ownerDocument.createElement("ADDRESS")
        vXMLRequest.appendChild xmlTempTableElem

        Set xmlElement = vXMLRequest.ownerDocument.createElement("ADDRESSGUID")
        xmlElement.Text = strAddressGuid
        xmlTempTableElem.appendChild xmlElement

        objIAddressDO.DeleteAddress xmlTempTableElem

    Next intNodeIndex

    Set ICustomerTxBO_DeleteCustomerAddress = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_DeleteCustomerAddressExit:

    Set objICustomerDO = Nothing
    Set objIAddressDO = Nothing
    Set objIADOAssist = Nothing
    
    Set objSQLAssist = Nothing

    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlTempTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing

    'BMIDS00006
    Set xmlDoc = Nothing
    Set xmlAddressGuid = Nothing
    'BMIDS00006 End

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

ICustomerTxBO_DeleteCustomerAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_DeleteCustomerAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_DeleteCustomerAddressExit
End Function

Private Function ICustomerTxBO_UpdateCustomer( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ICustomerTxBO_UpdateCustomerVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateCustomer"

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList, objElement As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.selectNodes(".//CUSTOMER")

    If xmlNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                oeMissingElement & " CUSTOMER tag not found"
    End If
        
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    Dim intNodeIndex As Integer
    
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objICustomerDO.UpdateCustomer objElement
    Next intNodeIndex

   If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set ICustomerTxBO_UpdateCustomer = xmlResponseElem
    
ICustomerTxBO_UpdateCustomerExit:

    Set objICustomerDO = Nothing
    
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    Exit Function

ICustomerTxBO_UpdateCustomerVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateCustomer = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdateCustomerExit

End Function

Private Function ICustomerTxBO_UpdateCustomerAddress(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               CustomerAddress, and the corresponding records in tables 'Address', 'Tenancy'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest   xml Request node containing update data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_UpdateCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateCustomerAddress"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objElement As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement

    Dim objICustomerDO As ICustomerDO
    Dim objIAddressBO As IAddressBO
    Dim objComboDO As ComboDO
    Dim objIADOAssist As IADOAssist
    Dim objGlobalParameterBO As GlobalParameterBO

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.selectNodes(".//CUSTOMERADDRESS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                oeMissingElement & " CUSTOMERADDRESS tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
        Set objIAddressBO = New AddressBO
        Set objComboDO = New ComboDO
        Set objIADOAssist = New ADOAssist
        Set objGlobalParameterBO = New GlobalParameterBO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
        Set objIAddressBO = objContext.CreateInstance(App.Title & ".AddressBO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objGlobalParameterBO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    End If
    
    ' Fetch the data required for deleting the previous addresses, if required
    Dim strDeletePreviousAddresses As String
    strDeletePreviousAddresses = objXmlAssist.GetTagValue(vXMLRequest, "DELETEPREVIOUSADDRESSES")

    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)

        'Update the Address Record
        Set xmlTableElem = objElement.selectSingleNode("ADDRESS")
        If Not xmlTableElem Is Nothing Then
            objIAddressBO.UpdateAddress xmlTableElem
        End If

        'BM0174 do not delete any tenancy records as a new customer version has not yet been created.
        
'        'If the customer was a Tenant at this address, but no longer,
'        'delete the corresponding record from the table 'Tenncy'
'        Dim strNatureOfOccupancy As String, strCustomerNumber As String
'        Dim strCustomerVersionNumber As String, strCustAddressSeqNumber As String
'        Dim blnTenancyDelete As Boolean
'
'        blnTenancyDelete = False
'        strNatureOfOccupancy = objXmlAssist.GetTagValue(vXMLRequest, "NATUREOFOCCUPANCY")
'
'        If strNatureOfOccupancy <> "" Then
'            If Not (objComboDO.IsItemInValidation("NatureOfOccupancy", strNatureOfOccupancy, "TF") Or _
'                    objComboDO.IsItemInValidation("NatureOfOccupancy", strNatureOfOccupancy, "TU")) Then
'                blnTenancyDelete = True
'            End If
'        Else
'            blnTenancyDelete = True
'        End If
'
'        If blnTenancyDelete Then
'            ' Fetch the key values of customer address
'            If Not objICustomerDO.bGetKeysOfCustomerAddress(vXMLRequest, strCustomerNumber, _
'                                    strCustomerVersionNumber, strCustAddressSeqNumber) Then
'
'                    objErrAssist.ThrowError TypeName(Me), strFunctionName, 519, _
'                                                        "for CUSTOMERADDRESS"
'            End If
'
'            'Check whether a corresponding Tenancy record exists - If so delete
'            Dim strCondition As String, strXML As String
'            strCondition = "CUSTOMERNUMBER =" & objSQLAssist.FormatString(strCustomerNumber) _
'                            & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
'                            & " And CUSTOMERADDRESSSEQUENCENUMBER = " & strCustAddressSeqNumber
'
'            If objIADOAssist.CheckRecordExists("TENANCY", strCondition) Then
'                strXML = "<TENANCY><CUSTOMERNUMBER>" & strCustomerNumber & "</CUSTOMERNUMBER>" _
'                         & "<CUSTOMERVERSIONNUMBER>" & strCustomerVersionNumber & "</CUSTOMERVERSIONNUMBER>" _
'                         & "<CUSTOMERADDRESSSEQUENCENUMBER>" & strCustAddressSeqNumber _
'                         & "</CUSTOMERADDRESSSEQUENCENUMBER></TENANCY>"
'
'                Set xmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
'                Call objICustomerDO.DeleteTenancy(xmlDoc.documentElement)
'            End If
'        End If

        objICustomerDO.UpdateCustomerAddress objElement
        If strDeletePreviousAddresses = "1" Then
            Set xmlTempResponseElem = DeletePreviousCustomerAddresses(objElement, vXMLRequest)
            m_objErrAssist.CheckResponse xmlTempResponseElem, xmlResponseElem, True
        End If
    Next
    
   If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set ICustomerTxBO_UpdateCustomerAddress = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_UpdateCustomerAddressExit:

    Set objICustomerDO = Nothing
    Set objIAddressBO = Nothing
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing

    Set xmlNodeList = Nothing
    Set xmlDoc = Nothing
    Set objElement = Nothing
    Set xmlTableElem = Nothing

    Exit Function

ICustomerTxBO_UpdateCustomerAddressVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdateCustomerAddressExit


End Function

Private Function DeletePreviousCustomerAddresses(ByVal objElement As IXMLDOMElement, _
                                                 ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
'--------------------------------------------------------------------------------------
'Description :
'--------------------------------------------------------------------------------------

On Error GoTo DeleteErr
    
    Const strFunctionName As String = "DeletePreviousCustomerAddresses"

    Dim objICustomerDO As ICustomerDO
    Dim objIClassDef As IomCustClassDef
    Dim objIDOAssist As IDOAssist

    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strAddressSequenceNumber As String
    Dim strDateMovedIn As String, dtDateMovedIn As Date, lngYears As Long
    
    Dim xmlCustomerAddressElem As IXMLDOMNode, xmlElem As IXMLDOMNode, xmlRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode, xmlNode As IXMLDOMNode, xmlResponseElem As IXMLDOMElement
    Dim xmlElem2 As IXMLDOMElement, xmlAddressGuid As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strDateMovedIn = m_objXmlAssist.GetTagValue(objElement, "DATEMOVEDIN")
    If strDateMovedIn = "" Then
        Set DeletePreviousCustomerAddresses = xmlResponseElem
        GoTo DeleteExit
    End If
        
    ' Find the customer addresses which have higher sequence than the current one
    strCustomerNumber = objElement.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustomerVersionNumber = objElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    strAddressSequenceNumber = objElement.selectSingleNode(".//CUSTOMERADDRESSSEQUENCENUMBER").Text
    
    If strCustomerNumber = "" Or strCustomerVersionNumber = "" Or strAddressSequenceNumber = "" Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, "CustomerAddress key data not found"
    End If
    
    Dim strAddlCondition As String
    strAddlCondition = "CUSTOMERADDRESSSEQUENCENUMBER > " & strAddressSequenceNumber
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objICustomerDO = New CustomerDO
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    Set xmlRequest = m_objXmlAssist.GetRequestNodeEx(vXMLRequest)
    Set xmlCustomerAddressElem = xmlDoc.createElement("CUSTOMERADDRESS")
    xmlRequest.appendChild xmlCustomerAddressElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElem.Text = strCustomerNumber
    xmlCustomerAddressElem.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlElem.Text = strCustomerVersionNumber
    xmlCustomerAddressElem.appendChild xmlElem
        
    ' Find all the address with higher AddressSequenceNumber
    Set objIClassDef = New omCustClassDef
    On Error GoTo FindCustomerVbErr
    Set xmlTempResponse = objIDOAssist.FindListMultiple(xmlCustomerAddressElem, _
                                objIClassDef.LoadCustomerAddressData, vstrAdditionalCodition:=strAddlCondition)
    On Error GoTo DeleteErr
    
    If Not xmlTempResponse Is Nothing Then
        Set xmlCustomerList = xmlTempResponse.selectNodes("CUSTOMERADDRESS")
        If xmlCustomerList.length > 0 Then
            Set xmlElem = xmlDoc.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
            xmlCustomerAddressElem.appendChild xmlElem
            
            Set xmlElem2 = xmlDoc.createElement("NATUREOFOCCUPANCY")
            xmlCustomerAddressElem.appendChild xmlElem2
            
            Set xmlAddressGuid = xmlDoc.createElement("ADDRESSGUID")
            xmlCustomerAddressElem.appendChild xmlAddressGuid
        End If
        
        ' Delete each address retreived
        For Each xmlNode In xmlCustomerList
            xmlElem.Text = xmlNode.selectSingleNode(".//CUSTOMERADDRESSSEQUENCENUMBER").Text
            xmlElem2.Text = xmlNode.selectSingleNode(".//NATUREOFOCCUPANCY").Text
            xmlAddressGuid.Text = xmlNode.selectSingleNode(".//ADDRESSGUID").Text
            Set xmlTempResponse = ICustomerTxBO_DeleteCustomerAddress(xmlRequest, True)
            m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
                       
        Next xmlNode
    End If
    
    Set DeletePreviousCustomerAddresses = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteExit:
    Set objIDOAssist = Nothing
    Set objICustomerDO = Nothing
    Set objIClassDef = Nothing
    
    Set xmlCustomerAddressElem = Nothing
    Set xmlCustomerList = Nothing
    Set xmlElem = Nothing
    Set xmlElem2 = Nothing
    Set xmlAddressGuid = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlDoc = Nothing
        
    Exit Function
    
FindCustomerVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

DeleteErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set DeletePreviousCustomerAddresses = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit
    
End Function

Private Function ICustomerTxBO_DeleteTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteTenancy multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteTenancyVbErr

    Const strFunctionName As String = "ICustomerTxBO_DeleteTenancy"

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
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "TENANCY" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("TENANCY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "TENANCY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerDO.DeleteTenancy xmlElement
    Else
        ' For each relevant node call the data object DeleteTenancy
        For Each xmlElement In xmlNodeList
            objICustomerDO.DeleteTenancy xmlElement
        Next
    End If
    
    Set ICustomerTxBO_DeleteTenancy = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteTenancyExit:

    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteTenancyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_DeleteTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteTenancyExit
End Function

Private Function DetermineOperationOnTenancy(xmlTenancyElem As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:  Determines the operation to be performed on 'Tenancy' in 'SaveTenancy'
'               method
' pass:         xmlElement contating the Tenancy Data
' return:       string representing the operation (CREATE/UPDATE)
' Error :
'------------------------------------------------------------------------------------------
On Error GoTo DetermineOperationOnTenancyVbErr
    Const strFunctionName As String = "DetermineOperationOnTenancy"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strCondition As String, strCustAddrSeqNumber As String

    Dim objICustomerDO As ICustomerDO
    Dim objIClassDef As IomCustClassDef
    Dim objIADOAssist As IADOAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objICustomerDO = New CustomerDO
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    Call objICustomerDO.bGetKeysOfCustomerAddress(xmlTenancyElem, strCustomerNumber, _
                                strCustomerVersionNumber, strCustAddrSeqNumber)

    strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                    & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                    & " And CUSTOMERADDRESSSEQUENCENUMBER = " & strCustAddrSeqNumber

    If objIADOAssist.CheckRecordExists("TENANCY", strCondition) Then
        ' SR 15/08/00 - SYS1412 : include the functionality for deletion of tenancy
        Set objIClassDef = New omCustClassDef
        Dim operationOnTenancy As BOOPERATIONTYPE
        operationOnTenancy = m_objXmlAssist.DetermineOperationEx(xmlTenancyElem, objIClassDef.LoadTenancyData)
        
        If operationOnTenancy = booDelete Then
            DetermineOperationOnTenancy = "DELETE"
        Else
            DetermineOperationOnTenancy = "UPDATE"
        End If
    Else
        DetermineOperationOnTenancy = "CREATE"
    End If

    Set objICustomerDO = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Exit Function
    
DetermineOperationOnTenancyVbErr:
    Set objICustomerDO = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error forCalling object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICustomerTxBO_SaveTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the Tenancy Data after finding the operation to be performed on it
'               and also makes the changes required in table 'ThirdParty'
' pass:         vXMLRequest  xml Request node representing the data to be persisted
' return:       nil
' Error :      Invalid Paramenter
'------------------------------------------------------------------------------------------

On Error GoTo ICustomerTxBO_SaveTenancyVbErr

    Const strFunctionName As String = "ICustomerTxBO_SaveTenancy"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlDoc As FreeThreadedDOMDocument40 'document to load the output from ThirdPartyBO
    Dim xmlTPResponseElem As IXMLDOMElement 'Response element in output from ThirdPartyBO
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIADOAssist As ADOAssist, objSQLAssist As New SQLAssist
    Dim xmlTenancyElem As IXMLDOMElement
    Dim xmlDirectoryGuidElem As IXMLDOMElement
    Dim xmlThirdPartyGuidElem As IXMLDOMElement

    Dim strDirectoryGUID As String
    Dim strDirAddressInd As String
    Dim strThirdPartyType As String, strThirdPartyGUID As String

    Dim objICustomerDO As ICustomerDO
 '   Dim objThirdPartyBO As ThirdPartyBO

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'If TenancyXML is NULL raise error
    Set xmlTenancyElem = vXMLRequest.getElementsByTagName("TENANCY").Item(0)
    If xmlTenancyElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, "Invalid Parameter "
    End If

    'Set DirectoryGuid appropriately before making any change to the table 'Tenancy'
    Set xmlDirectoryGuidElem = xmlTenancyElem.getElementsByTagName("DIRECTORYGUID").Item(0)
    strDirAddressInd = objXmlAssist.GetTagValue(xmlTenancyElem, "DIRECTORYADDRESSINDICATOR")

    If (Not xmlDirectoryGuidElem Is Nothing) And strDirAddressInd = "0" Then
        xmlDirectoryGuidElem.Text = ""
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
'        Set objThirdPartyBO = New ThirdPartyBO
        Set objIADOAssist = New ADOAssist
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
'        Set objThirdPartyBO = objContext.CreateInstance(cstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    'Get the ThirdPartyType; It might not be mentioned in XML
    'If so, get it from Table (Requird for next If loop).
    strThirdPartyType = objXmlAssist.GetTagValue(vXMLRequest, "THIRDPARTYTYPE")
    If strThirdPartyType = "" Then
        strThirdPartyGUID = objXmlAssist.GetTagValue(xmlTenancyElem, "THIRDPARTYGUID")
        
        Dim strCondition As String, strTemp As String

        strCondition = "THIRDPARTYGUID =" & objSQLAssist.FormatString(strThirdPartyGUID)
        If Not IsNull(objIADOAssist.GetValueFromTable("THIRDPARTY", strCondition, _
                                    "THIRDPARTYTYPE", strTemp)) Then
            strThirdPartyType = strTemp
        End If
    End If

    'Determine the action to be performed on 'Tenancy' and do accordingly
    Select Case DetermineOperationOnTenancy(xmlTenancyElem)
        Case "CREATE"
            objICustomerDO.CreateTenancy xmlTenancyElem
        Case "UPDATE"
            objICustomerDO.UpdateTenancy xmlTenancyElem
        Case " DELETE"
            objICustomerDO.DeleteTenancy xmlTenancyElem
    End Select

    Set ICustomerTxBO_SaveTenancy = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_SaveTenancyExit:
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
  '  Set objThirdPartyBO = Nothing
    Set objICustomerDO = Nothing
    Set xmlTenancyElem = Nothing
    Set xmlDirectoryGuidElem = Nothing
    Set xmlThirdPartyGuidElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ICustomerTxBO_SaveTenancyVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SaveTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_SaveTenancyExit

End Function

Private Function ICustomerTxBO_SaveAlias(ByVal vXMLRequestElem As IXMLDOMElement, _
                                         ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the Alias Data after finding the operation to be performed on it
' pass:         vstrXMLRequest  xml Request data to be passed to the methos
' return:       nil
' Error :      Invalid Paramenter
'------------------------------------------------------------------------------------------

On Error GoTo ICustomerTxBO_SaveAliasVbErr

    Const strFunctionName As String = "ICustomerTxBO_SaveAlias"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim xmlAliasDef As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlAliasPersonElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlAliasElem As IXMLDOMElement
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objICustomerDO As ICustomerDO
    Dim objICustomerBO As New ICustomerBO
    Dim objIClassDef As IomCustClassDef, xmlDoc As FreeThreadedDOMDocument40

    Set xmlResponseElem = vXMLRequestElem.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequestElem.ownerDocument.getElementsByTagName("ALIAS")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ALIAS" & " tag not found"
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    Dim intNodeIndex As Integer

    ' For each relevant node, determine the operation to be performed and
    ' call the appropriate method from DO
    Set objIClassDef = New omCustClassDef
    Set xmlDoc = objIClassDef.LoadAliasData
    
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Set objICustomerBO = New CustomerBO
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set xmlAliasPersonElem = xmlNodeList.Item(intNodeIndex)
        Select Case objXmlAssist.DetermineOperationEx(xmlAliasPersonElem, xmlDoc)
            Case booCreate
                Set xmlTempResponseNode = objICustomerBO.ValidateAlias(xmlAliasPersonElem)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem
    
                Set xmlTempResponseNode = objICustomerDO.CreateAlias(xmlAliasPersonElem)
                ' Add CustomerNumber, VersionNumber and AliasSeq number to GeneratedKeys
                Set xmlAliasElem = vXMLRequestElem.ownerDocument.createElement("ALIAS")
                vxmlGeneratedKeys.appendChild xmlAliasElem
                
                strCustomerNumber = objXmlAssist.GetTagValue(xmlAliasPersonElem, "CUSTOMERNUMBER")
                strCustomerVersionNumber = _
                        objXmlAssist.GetTagValue(xmlAliasPersonElem, "CUSTOMERVERSIONNUMBER")
                
                Set xmlElement = vXMLRequestElem.ownerDocument.createElement("CUSTOMERNUMBER")
                xmlElement.Text = strCustomerNumber
                xmlAliasElem.appendChild xmlElement
                
                Set xmlElement = vXMLRequestElem.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
                xmlElement.Text = strCustomerVersionNumber
                xmlAliasElem.appendChild xmlElement
                
                xmlAliasElem.appendChild xmlTempResponseNode.selectSingleNode("ALIASSEQUENCENUMBER")
            Case booUpdate
                Set xmlTempResponseNode = objICustomerBO.ValidateAlias(xmlAliasPersonElem)
                objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem
                
                Call objICustomerDO.UpdateAlias(xmlAliasPersonElem)
            Case booDelete
                Call objICustomerDO.DeleteAlias(xmlAliasPersonElem)
            Case Else
            ' N/A
        End Select
    Next

    Set ICustomerTxBO_SaveAlias = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_SaveAliasExit:

    Set xmlNodeList = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerBO = Nothing
    Set objICustomerDO = Nothing
    Set xmlAliasPersonElem = Nothing
    Set xmlAliasElem = Nothing
    Set xmlAliasDef = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerTxBO_SaveAliasVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SaveAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_SaveAliasExit
End Function

Private Function ICustomerTxBO_CreateDependant(ByVal vXMLRequest As IXMLDOMElement, _
                                               ByVal vxmlGeneratedKeys As IXMLDOMElement) _
                                               As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vXMLRequest  xml Request node containing data to be persisted
'               vxmlGeneratedKeys : node to attach the generated keys to
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_CreateDependantVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateDependant"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlDependantOutElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlTemp As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    If vXMLRequest.nodeName = "DEPENDANT" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("DEPENDANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DEPENDANT" & " tag not found"
        End If
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    Dim intNodeIndex As Integer
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    
    If Not (xmlElement Is Nothing) Then
        ' Create dependant record
        Set xmlNode = objICustomerDO.CreateDependant(xmlElement)
        
        ' Add relevant nodes to GeneratedKeys - Output
        strCustomerNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERVERSIONNUMBER")
        
        Set xmlDependantOutElem = vxmlGeneratedKeys.ownerDocument.createElement("DEPENDANT")
        Set xmlElement = vxmlGeneratedKeys.ownerDocument.createElement("CUSTOMERNUMBER")
        xmlElement.Text = strCustomerNumber
        xmlDependantOutElem.appendChild xmlElement
        
        Set xmlElement = vxmlGeneratedKeys.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
        xmlElement.Text = strCustomerVersionNumber
        xmlDependantOutElem.appendChild xmlElement
        
        xmlDependantOutElem.appendChild xmlNode
        vxmlGeneratedKeys.appendChild xmlDependantOutElem
    Else
        ' For each relevant node add key and call the data object create
        For Each xmlElement In xmlNodeList
            ' Create dependant record
            Set xmlNode = objICustomerDO.CreateDependant(xmlElement)
            
            ' Add relevant nodes to GeneratedKeys - Output
            strCustomerNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERNUMBER")
            strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlElement, "CUSTOMERVERSIONNUMBER")
            
            Set xmlDependantOutElem = vxmlGeneratedKeys.ownerDocument.createElement("DEPENDANT")
            Set xmlElement = vxmlGeneratedKeys.ownerDocument.createElement("CUSTOMERNUMBER")
            xmlElement.Text = strCustomerNumber
            xmlDependantOutElem.appendChild xmlElement
            
            Set xmlElement = vxmlGeneratedKeys.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
            xmlElement.Text = strCustomerVersionNumber
            xmlDependantOutElem.appendChild xmlElement
            
            xmlDependantOutElem.appendChild xmlNode
            vxmlGeneratedKeys.appendChild xmlDependantOutElem
        Next
    End If
    
    Set ICustomerTxBO_CreateDependant = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_CreateDependantExit:

    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlDependantOutElem = Nothing
    Set xmlElement = Nothing
    Set xmlTemp = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ICustomerTxBO_CreateDependantVbErr:
   If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_CreateDependantExit
   

End Function

Private Function ICustomerTxBO_CreateTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateTenancy multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateTenancyVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateTenancy"

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
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "TENANCY" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("TENANCY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "TENANCY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerDO.CreateTenancy xmlElement
    Else
        ' For each relevant node call the data object CreateTenancy
        For Each xmlElement In xmlNodeList
            objICustomerDO.CreateTenancy xmlElement
        Next
    End If
    
    Set ICustomerTxBO_CreateTenancy = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateTenancyExit:

    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateTenancyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateTenancyExit
End Function

Private Function ICustomerTxBO_CreateCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) _
                                                            As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vXMLRequest  xml Request node containing data to be persisted
' return:       xml Node giving the response (whether the method is successful or not)
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_CreateCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateCustomerTelephone"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

'    Dim xmlOut as DOMDocument
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement

    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERTELEPHONENUMBER")

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "CUSTOMERTELEPHONENUMBER tag not found"
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    Dim intNodeIndex As Integer

    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objICustomerDO.CreateCustomerTelephone objElement
    Next

    Set ICustomerTxBO_CreateCustomerTelephone = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_CreateCustomerTelephoneExit:
'   Set xmlOut = nothing
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerTxBO_CreateCustomerTelephoneVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_CreateCustomerTelephoneExit
End Function

Private Function ICustomerTxBO_UpdatePersonalDetailsAndEmpt(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo UpdateVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdatePersonalDetailsAndEmpt"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
   
    Dim objCustEmpBO As CustomerEmploymentBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = "CUSTOMER" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("CUSTOMER").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  " CUSTOMER" & " tag not found"
    End If
       
    ' Save Customer details and the numbers generated into local variables
    ICustomerTxBO_UpdateCustomerDetails vXMLRequest, xmlResponseElem
    
    ' Save employment details
    If objContext Is Nothing Then
        Set objCustEmpBO = New CustomerEmploymentBO
    Else
        Set objCustEmpBO = _
            objContext.CreateInstance(cstrCUSTEMPLOYMENT_COMPONENT & ".CustomerEmploymentBO")
    End If
       
    Dim strResponse As String
    strResponse = objCustEmpBO.SaveEmploymentDetails(vXMLRequest.xml)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'Save Employed details
    strResponse = objCustEmpBO.SaveEmployedDetails(vXMLRequest.xml)
    Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set ICustomerTxBO_UpdatePersonalDetailsAndEmpt = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:

    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set objErrAssist = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdatePersonalDetailsAndEmpt = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function

Private Function ICustomerTxBO_UpdateCustomerVersion( _
        ByVal vXMLRequest As IXMLDOMElement, _
        Optional ByVal bTaxDetailsOnly As Boolean = False) As IXMLDOMNode

On Error GoTo ICustomerTxBO_UpdateCustomerVersionVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateCustomerVersion"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    If vXMLRequest.nodeName = "CUSTOMERVERSION" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERVERSION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "CUSTOMERVERSION" & " tag not found"
        End If
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    If Not (xmlElement Is Nothing) Then
        objICustomerDO.UpdateCustomerVersion xmlElement
    Else
        ' For each relevant node that we are passed call the data object UpdateDependant
        For Each xmlElement In xmlNodeList
            objICustomerDO.UpdateCustomerVersion xmlElement, bTaxDetailsOnly
        Next
    End If

    Set ICustomerTxBO_UpdateCustomerVersion = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerTxBO_UpdateCustomerVersionExit:
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerTxBO_UpdateCustomerVersionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateCustomerVersion = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdateCustomerVersionExit
End Function

Private Function ICustomerTxBO_UpdateDependant(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateDependant multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing UpdateDependant data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_UpdateDependantVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateDependant"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPreviousKeyNode As IXMLDOMNode   'SYS1168
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    If vXMLRequest.nodeName = "DEPENDANT" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("DEPENDANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DEPENDANT" & " tag not found"
        End If
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    'SYS1168
    If Not (xmlElement Is Nothing) Then
        Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
        If Not (xmlPreviousKeyNode Is Nothing) Then
            objICustomerDO.DeleteDependant xmlPreviousKeyNode
            objICustomerDO.CreateDependant xmlElement
        Else
            objICustomerDO.UpdateDependant xmlElement
        End If
    Else
        ' For each relevant node that we are passed call the data object UpdateDependant
        For Each xmlElement In xmlNodeList
            Set xmlPreviousKeyNode = xmlElement.selectSingleNode("PREVIOUSKEY")
            If Not (xmlPreviousKeyNode Is Nothing) Then
                objICustomerDO.DeleteDependant xmlPreviousKeyNode
                objICustomerDO.CreateDependant xmlElement
            ElseIf xmlElement.parentNode.nodeName <> "PREVIOUSKEY" Then
                objICustomerDO.UpdateDependant xmlElement
            End If
        Next
    End If

    Set ICustomerTxBO_UpdateDependant = xmlResponseElem.cloneNode(True)

ICustomerTxBO_UpdateDependantExit:
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPreviousKeyNode = Nothing

    Exit Function

ICustomerTxBO_UpdateDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdateDependantExit
End Function

Private Function ICustomerTxBO_DeleteDependant(ByVal vXMLRequest As IXMLDOMElement, _
                                Optional ByVal rstrPersonGUID As Variant) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteDependant a single instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform DeleteDependant
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be DeleteDependantd
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_DeleteDependantVbErr

    Const strFunctionName As String = "ICustomerTxBO_DeleteDependant"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    If vXMLRequest.nodeName = "DEPENDANT" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("DEPENDANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DEPENDANT" & " tag not found"
        End If
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    If Not (xmlElement Is Nothing) Then
        objICustomerDO.DeleteDependant xmlElement, rstrPersonGUID:=rstrPersonGUID
    Else
        ' For each relevant node that we are passed call the data object DeleteDependant
        For Each xmlElement In xmlNodeList
            objICustomerDO.DeleteDependant xmlElement, rstrPersonGUID:=rstrPersonGUID
        Next
    End If
    
    Set ICustomerTxBO_DeleteDependant = xmlResponseElem.cloneNode(True)
    
ICustomerTxBo_DeleteDependantExit:

    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerTxBO_DeleteDependantVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_DeleteDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBo_DeleteDependantExit

End Function

Private Function ICustomerTxBO_DeleteAllCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) _
                                            As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vXMLRequest  xml Request node containing data to which identifies
'                            instance of the persistant data to be retrieved
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_DeleteAllCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerTxBO_DeleteAllCustomerTelephone"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
'   Dim xmlOut as new DomDocument
    
    Dim objICustomerDO As ICustomerDO

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    ' Pass through to the data object
    Call objICustomerDO.DeleteAllCustomerTelephone(vXMLRequest)

    Set ICustomerTxBO_DeleteAllCustomerTelephone = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerTxBO_DeleteAllCustomerTelephoneExit:
    
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
'    Set xmlOut = Nothing
    
    Exit Function

ICustomerTxBO_DeleteAllCustomerTelephoneVbErr:

If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_DeleteAllCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_DeleteAllCustomerTelephoneExit
End Function

Private Function ICustomerTxBO_UpdatePersonalDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdatePersonalDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing UpdatePersonalDetails data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_UpdatePersonalDetailsVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdatePersonalDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlResponseElem As IXMLDOMElement
        
    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMER")
    If xmlNodeList.length = 0 Then
        Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERVERSION")
    End If

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "CUSTOMER/CUSTOMERVERSION tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    ' For each relevant node that we are passed call the data object UpdatePersonalDetails
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objICustomerDO.UpdatePersonalDetails objElement
    Next

    Set ICustomerTxBO_UpdatePersonalDetails = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_UpdatePersonalDetailsExit:
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlResponseElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

ICustomerTxBO_UpdatePersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdatePersonalDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_UpdatePersonalDetailsExit
End Function

Private Function ICustomerTxBO_UpdatePropertyInsuranceDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate ICustomerDO_UpdatePropertyInsuranceDetails to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdatePropertyInsuranceDetails"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objICustomerDO.UpdatePropertyInsuranceDetails xmlElement
    
    Set ICustomerTxBO_UpdatePropertyInsuranceDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdatePropertyInsuranceDetailsExit:
    
    Set objICustomerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdatePropertyInsuranceDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatePropertyInsuranceDetailsExit

End Function

Private Function ICustomerTxBO_UpdateTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateTenancy multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTenancyVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateTenancy"

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
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "TENANCY" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("TENANCY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "TENANCY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerDO.UpdateTenancy xmlElement
    Else
        ' For each relevant node call the data object UpdateTenancy
        For Each xmlElement In xmlNodeList
            objICustomerDO.UpdateTenancy xmlElement
        Next
    End If
    
    Set ICustomerTxBO_UpdateTenancy = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateTenancyExit:
    
    Set objICustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateTenancyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateTenancyExit
End Function

Private Function ICustomerTxBO_SaveCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveCustomerTelephone multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing SaveCustomerTelephone data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerTxBO_SaveCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerTxBO_SaveCustomerTelephone"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
'    Dim xmlOut as DOMDocument
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    Dim objIClassDef As IomCustClassDef
    Dim botOperationType As BOOPERATIONTYPE

    Set xmlResponseElem = vXMLRequest.ownerDocument.createElement("RESPONSE")
'   xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = vXMLRequest.getElementsByTagName("CUSTOMERTELEPHONENUMBER")
    If xmlNodeList.length = 0 Then
        ' No need to raise error just bail out of the routine
        Set ICustomerTxBO_SaveCustomerTelephone = xmlResponseElem.cloneNode(True) 'SYS0327
        Exit Function
    End If

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    Set objIClassDef = New omCustClassDef
    
    ' For each relevant node that we are passed call the data object Create/Update
    For Each xmlNode In xmlNodeList
        botOperationType = objXmlAssist.DetermineOperationEx(xmlNode, objIClassDef.LoadCustomerTelephoneData)
        Select Case botOperationType
            Case booUpdate: objICustomerDO.UpdateCustomerTelephone xmlNode
            Case booCreate: objICustomerDO.CreateCustomerTelephone xmlNode
            Case booDelete: objICustomerDO.DeleteCustomerTelephone xmlNode
            Case Else
                ' N/A
        End Select
    Next
    
    Set ICustomerTxBO_SaveCustomerTelephone = xmlResponseElem.cloneNode(True)
    
ICustomerTxBO_SaveCustomerTelephoneExit:
    Set objICustomerDO = Nothing
    Set objIClassDef = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ICustomerTxBO_SaveCustomerTelephoneVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_SaveCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ICustomerTxBO_SaveCustomerTelephoneExit
End Function

Private Sub ICustomerTxBO_CreateCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, _
                                                    ByVal vobjResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateCustomerDetailsVbErr

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    
    Dim objCustomerDO As ICustomerDO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objCustomerNTxBO As ICustomerNTxBO
    
    ' PSC 11/10/2005 MAR57 - Start
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim objCustomerBO As ICustomerBO
    
    Dim blnUpdateCRSCustomer As Boolean
    Dim strUseAdminCRSUpdate As String
    Dim strOtherSystemCustNo As String
    ' PSC 11/10/2005 MAR57 - End
    
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strNumber As String
    
    Dim lngNumber As Long
    
    Dim blnCreateNewCustomer As Boolean
     
    If objContext Is Nothing Then
        Set objCustomerDO = New CustomerDO
        Set objGlobalParamDO = New GlobalParameterDO
        Set objCustomerNTxBO = New CustomerNTxBO
    Else
        Set objCustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
        Set objGlobalParamDO = objContext.CreateInstance(cstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objCustomerNTxBO = objContext.CreateInstance(App.Title & ".CustomerNTxBO")
    End If
    
    'BMIDS00004 MDC 13/05/2002
    'Enable creation of new customer versions
    
    strCustomerNumber = xmlGetNodeText(vobjRequestNode, "//CUSTOMER/CUSTOMERNUMBER")
    
    ' PSC 19/10/2005 MAR57
    strUseAdminCRSUpdate = objGlobalParamDO.GetCurrentParameterByType("UseAdminCRSUpdate", "BOOLEAN")
    
    If Len(Trim$(strCustomerNumber)) = 0 Then
        
        'BMIDS00443 Admin system can return Omiga 3 customer number as OMIGACUSTOMERNUMBER
        If Not vobjRequestNode.selectSingleNode("//CUSTOMER/OMIGACUSTOMERNUMBER") Is Nothing Then
            strCustomerNumber = Trim(xmlGetNodeText(vobjRequestNode, "//CUSTOMER/OMIGACUSTOMERNUMBER"))
        Else
        'BMIDS00443 End
        
            'Generate a new customer number
            lngNumber = objCustomerNTxBO.GetNextCustomerNumber
                
            ' APS UNIT TEST REF 84 - Prefix the Next Number with leading zeros
            ' MH Changed DO so that it determines when a number is valid
            ' MH . This was done to reduce the overhead of getting a number.
            strNumber = Format(lngNumber, "0000000")
    
            'IK_17/05/2006_EP561
            If IsCheckDigitRequired Then
                strCustomerNumber = strNumber & GenerateCheckDigit(strNumber)
            Else
                strCustomerNumber = strNumber
            End If
            'IK_17/05/2006_EP561_ends
            
        End If 'BMIDS00443
        
        strCustomerVersionNumber = "1"
        blnCreateNewCustomer = True
    Else
        blnCreateNewCustomer = False
    End If
    
    objCustomerDO.CreateCustomerDetails _
        vobjRequestNode, _
        vobjResponseNode, _
        strCustomerNumber, _
        strCustomerVersionNumber, _
        blnCreateNewCustomer
        
    strCustomerVersionNumber = xmlGetNodeText(vobjResponseNode, ".//CUSTOMERVERSIONNUMBER")
    'BMIDS00004 MDC 13/05/2002 - End
        
    ' add customernumber to customer request node, for create lock function
    Set objXmlNode = vobjRequestNode.selectSingleNode(".//CUSTOMER")
    Set objXmlElem = vobjRequestNode.ownerDocument.createElement("CUSTOMERNUMBER")
    objXmlElem.Text = strCustomerNumber
    objXmlNode.appendChild objXmlElem
    
    ' create customer lock if global parameter allows
    ' OR if there is a 'LOCKING' node in the XML use that first
    Dim strLockingReqd As String
    Dim objXmlLockNode As IXMLDOMNode
    Set objXmlLockNode = vobjRequestNode.selectSingleNode(".//LOCKING")
    If Not objXmlLockNode Is Nothing Then
        strLockingReqd = objXmlLockNode.Text
    Else
        strLockingReqd = objGlobalParamDO.GetCurrentParameterByType("Locking", "BOOLEAN")
    End If
    
    If strLockingReqd = "1" Then
        objCustomerDO.CreateCustomerLock vobjRequestNode, vobjResponseNode
    End If
    
    ' PSC 11/10/2005 MAR57 - Start
    blnUpdateCRSCustomer = xmlGetNodeAsBoolean(vobjRequestNode, ".//CUSTOMER/UPDATECRSCUSTOMER")

    If blnUpdateCRSCustomer And strUseAdminCRSUpdate = "1" Then
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        
        Set xmlTempRequest = vobjRequestNode.cloneNode(True)
        Set xmlRoot = xmlTempDoc.createElement("CUSTOMERS")
        xmlTempRequest.appendChild xmlRoot
        
        Set xmlCustomer = xmlTempDoc.createElement("CUSTOMER")
        xmlRoot.appendChild xmlCustomer
        
        Set xmlNode = xmlTempDoc.createElement("CUSTOMERNUMBER")
        xmlNode.Text = strCustomerNumber
        xmlCustomer.appendChild xmlNode
        
        Set xmlNode = xmlTempDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlNode.Text = strCustomerVersionNumber
        xmlCustomer.appendChild xmlNode
        
        Set objCustomerBO = m_objContext.CreateInstance(App.Title & ".CustomerBO")
        Set xmlTempResponse = objCustomerBO.UpdateCRSCustomer(xmlTempRequest)
        errCheckXMLResponseNode xmlTempResponse, vobjResponseNode, True
        
        ' PSC 09/01/2006 MAR1001
        strOtherSystemCustNo = xmlGetNodeText(xmlTempResponse, "CUSTOMERLIST/CUSTOMER/@OTHERSYSTEMCUSTOMERNUMBER")
    End If
    ' PSC 11/10/2005 MAR57 - End
    
    ' return generated keys in response
    Set objXmlElem = vobjResponseNode.ownerDocument.createElement("CUSTOMERKEY")
    Set objXmlNode = vobjResponseNode.appendChild(objXmlElem)
    
    Set objXmlElem = vobjResponseNode.ownerDocument.createElement("CUSTOMERNUMBER")
    objXmlElem.Text = strCustomerNumber
    objXmlNode.appendChild objXmlElem
    
    Set objXmlElem = vobjResponseNode.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")
    objXmlElem.Text = strCustomerVersionNumber
    objXmlNode.appendChild objXmlElem
    
    ' PSC 11/10/2005 MAR57 - Start
    If Len(strOtherSystemCustNo) > 0 Then
        Set objXmlElem = vobjResponseNode.ownerDocument.createElement("OTHERSYSTEMCUSTOMERNUMBER")
        objXmlElem.Text = strOtherSystemCustNo
        objXmlNode.appendChild objXmlElem
    End If
    ' PSC 11/10/2005 MAR57 - End
    

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objCustomerDO = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlLockNode = Nothing
    Set objContext = Nothing
    
    ' PSC 11/10/2005 MAR57 - Start
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlRoot = Nothing
    Set xmlCustomer = Nothing
    Set xmlNode = Nothing
    Set xmlTempDoc = Nothing
    Set objCustomerBO = Nothing
    ' PSC 11/10/2005 MAR57 - End
    
    Exit Sub
    
CreateCustomerDetailsVbErr:
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Set objCustomerDO = Nothing
    Set objXmlElem = Nothing
    Set objXmlLockNode = Nothing
    Set objContext = Nothing
        
    ' re-raise error
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub ICustomerTxBO_UpdateCustomerDetails( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByVal vobjXmlResponse As IXMLDOMNode)

On Error GoTo UpdateCustomerDetailsVbErr

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objCustomerDO As ICustomerDO
    
    If objContext Is Nothing Then
        Set objCustomerDO = New CustomerDO
    Else
        Set objCustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objCustomerDO.UpdateCustomerDetails vobjXmlRequest, vobjXmlResponse
    
    Set objCustomerDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objContext = Nothing
    
    Exit Sub
    
UpdateCustomerDetailsVbErr:
    
    Set objCustomerDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Set objContext = Nothing

    ' re-raise the error back up the calling chain
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

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

Private Function ICustomerTxBO_UpdateContactHistory(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactHistoryVbErr

    Const strFunctionName As String = "ICustomerTxBO_UpdateContactHistory"
    
 
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_CONTACTHISTORY Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
    End If
        
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_CONTACTHISTORY & " tag not found"
    End If
   
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    
    objICustomerDO.UpdateContactHistory xmlElement
    
    Set ICustomerTxBO_UpdateContactHistory = xmlResponseElem
    
    m_objContext.SetComplete
    
UpdateContactHistoryExit:
    
    Set objICustomerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
      
    
    Exit Function

UpdateContactHistoryVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_UpdateContactHistory = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateContactHistoryExit

End Function

Private Function ICustomerTxBO_CreateContactHistory(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactHistoryVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateContactHistory"
    
         
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlKeys As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_CONTACTHISTORY Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_CONTACTHISTORY & " tag not found"
    End If
      
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    
    Set xmlKeys = objICustomerDO.CreateContactHistory(xmlElement)
    
    xmlResponseElem.appendChild xmlKeys
    
    Set ICustomerTxBO_CreateContactHistory = xmlResponseElem
    
    m_objContext.SetComplete
    
CreateContactHistoryExit:

    Set objICustomerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeys = Nothing
    
    
    Exit Function

CreateContactHistoryVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateContactHistory = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateContactHistoryExit

End Function

'MAR23 Add function

Private Function ICustomerTxBO_CreateKYCCheck(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateKYCCheckVbErr

    Const strFunctionName As String = "ICustomerTxBO_CreateKYCCheck"
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlKeys As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = "KNOWYOURCUSTOMERCHECK" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("KNOWYOURCUSTOMERCHECK").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  "KNOWYOURCUSTOMERCHECK" & " tag not found"
    End If
      
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
   
    objICustomerDO.CreateKYCCheck xmlElement

    Set ICustomerTxBO_CreateKYCCheck = xmlResponseElem
    
    m_objContext.SetComplete
    
CreateKYCCheckExit:

    Set objICustomerDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeys = Nothing
    
    
    Exit Function

CreateKYCCheckVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerTxBO_CreateKYCCheck = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateKYCCheckExit

End Function




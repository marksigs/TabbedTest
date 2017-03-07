VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CustomerEmploymentTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      CustomerEmploymentTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  List any other dependent components
'               e.g. CustomerEmploymentDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MCS    09/11/99    Created
'JLD    12/06/00    Added method GetCurrentEmploymentAndIncome for homezone use.
'SR     01/08/00    SYS1256 - Modified UpdateEmploymentDetails and new private sub DeletePreviousEmployment
'MC     02/08/00    SYS1368 - Check that UNEARNEDINCOMELIST has been found in SaveCurrentEmploymentAndIncome
'SR     07/08/00    SYS1383 - Modified method UpdateEmploymentDetails
'SR     17/08/00    SYS1257  Modified private method DeletePreviousEmployments - Do not explicityly call
'                         DeleteThirdParty.
'SR     30/08/00    SYS1420 - Modified method UpdateEmploymentDetails - delete linked records
'                   when employment status is changed.
'JLD    04/09/00    SYS1511 - slight error in deleting employment details.
'JLD    23/10/01    SYS2839 Request node required in SaveCurrentEmploymentAndIncome
'---------------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerEmploymentTxBO


'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    28/06/2002  BMIDS00052 Versions of Employment
'MO     03/07/2002  BMIDS00172 Object variable or with block variable not set error within
'                   function CopyEmploymentDetails()
'MDC    30/07/2002  BMIDS00265 Create new ThirdParty record when copying employment data
'SA     13/11/2002  BMIDS00693 CopyEmploymentDetails changed to remove another ThirdPartyGuid
'SA     13/11/2002  BMIDS00929 CopyEmploymentDetails - Need to create new Accountant record as we do thirdparty.
'------------------------------------------------------------------------------------------

Private Sub UpdateEmploymentField(ByVal vxmlRequest As IXMLDOMNode, _
                                  ByVal vstrField As String)

On Error GoTo UpdateEmploymentFieldVbErr

    Const strFunctionName As String = "UpdateEmploymentField"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlElement As IXMLDOMElement
    Dim strFieldValue               As String, _
        strCustomerNumber           As String, _
        strCustomerVersionNumber    As String, _
        strEmploymentSequenceNumber As String
    Dim blnFound As Boolean
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If

    '
    ' Update the EMPLOYMENT table with ACCOUNTANTGUID if appropriate
    '
    strFieldValue = objXmlAssist.GetTagValue(vxmlRequest, vstrField, blnFound)
    If blnFound Then
        strCustomerNumber = objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERVERSIONNUMBER")
        strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlRequest, "EMPLOYMENTSEQUENCENUMBER")

        Set xmlElement = vxmlRequest.ownerDocument.createElement("EMPLOYMENT")
        xmlElement.appendChild(vxmlRequest.ownerDocument.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber
        xmlElement.appendChild(vxmlRequest.ownerDocument.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber
        xmlElement.appendChild(vxmlRequest.ownerDocument.createElement("EMPLOYMENTSEQUENCENUMBER")).Text = strEmploymentSequenceNumber
        xmlElement.appendChild(vxmlRequest.ownerDocument.createElement(vstrField)).Text = strFieldValue

        objICustomerEmploymentDO.UpdateEmploymentDetails xmlElement
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEmploymentFieldExit:
    Set xmlElement = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    
    Exit Sub

UpdateEmploymentFieldVbErr:
    Set xmlElement = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'BMIDS00004 MDC 17/05/2002
Private Function ICustomerEmploymentTxBO_CopyEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CopyEmploymentDetails is used to copy employment data from one customer version
'               to another.
' pass:         vxmlRequest  xml Request data stream containing key data idenitifying
'               records to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CopyEmploymentDetailsVbErr
    
    Const strFunctionName As String = "ICustomerEmploymentTxBO_CopyEmploymentDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    Dim xmlEmploymentNode As IXMLDOMNode
    Dim xmlEmploymentListNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempNodeList As IXMLDOMNodeList
    Dim xmlEmploymentNodeList As IXMLDOMNodeList
    Dim xmlCustomerNodelist As IXMLDOMNodeList
    Dim xmlCustomerNode As IXMLDOMNode
    Dim objICustomerEmploymentBO As ICustomerEmploymentBO
    
    'BMIDS00265 MDC 30/07/2002
    Dim xmlGUIDNode As IXMLDOMNode
    Dim xmlGUIDNodeList As IXMLDOMNodeList
    'BMIDS00265 MDC 30/07/2002 - End
    'BMIDS00929 SA 14/11/02 {
    Dim xmlAccountantRequest As IXMLDOMElement
    Dim xmlAccountantElem As IXMLDOMElement
    Dim xmlAccountantGuid As IXMLDOMNode
    Dim xmlAccountantResponse As IXMLDOMElement
    'BMIDS00929 }

    Dim strNewCustomerVersionNumber As String
    Dim lngErrNo As Long
    Dim blnFound As Boolean
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objICustomerEmploymentBO = objContext.CreateInstance(App.Title & ".CustomerEmploymentBO")
    
    'Initialise Response
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlKeysNode = xmlOut.createElement("GENERATEDKEYS")

    'Get new customer version number (for later use)
    strNewCustomerVersionNumber = objXmlAssist.GetMandatoryElementText(vxmlRequest, ".//NEWCUSTOMERVERSION/CUSTOMERVERSIONNUMBER")
    
    'Create request for FindEmploymentList
    Set xmlTempRequestNode = vxmlRequest.cloneNode(False)
    Set xmlEmploymentListNode = xmlDoc.createElement("EMPLOYMENTLIST")
    Set xmlEmploymentNode = xmlDoc.createElement("EMPLOYMENT")
    Set xmlTempNode = objXmlAssist.GetMandatoryNode(vxmlRequest, ".//EXISTINGCUSTOMERVERSION")
    objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlTempNode, xmlEmploymentNode, False
    objXmlAssist.CopyNode "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER", xmlTempNode, xmlEmploymentNode, False
    xmlTempRequestNode.appendChild xmlEmploymentListNode
    xmlEmploymentListNode.appendChild xmlEmploymentNode

    'Call FindEmploymentList to retrieve a list of employment records to be copied
    Set xmlTempResponseNode = objICustomerEmploymentBO.FindEmploymentList(xmlTempRequestNode)
    'Handle record not found
    lngErrNo = objErrAssist.CheckResponse(xmlTempResponseNode, xmlResponseElem, False)
    If lngErrNo = 0 Then
        blnFound = True
    Else
        lngErrNo = objErrAssist.GetOmigaErrorNumber(lngErrNo)
        If lngErrNo = oeRecordNotFound Then
            blnFound = False
        Else
            objErrAssist.ThrowError App.Title, strFunctionName, lngErrNo
        End If
    End If
    
    If blnFound Then
        'Process each Employment found
        Set xmlEmploymentNodeList = xmlTempResponseNode.selectNodes(".//EMPLOYMENT")
        For Each xmlEmploymentNode In xmlEmploymentNodeList
            'Call GetEmploymentDetails for this employment record
            Set xmlTempRequestNode = vxmlRequest.cloneNode(False)
            xmlTempRequestNode.appendChild xmlEmploymentNode
            Set xmlTempResponseNode = objICustomerEmploymentBO.GetEmploymentDetails(xmlTempRequestNode)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            'Remove duplicate EMPLOYMENT records not required as children of EMPLOYEDDETAILS
            Set xmlTempNodeList = xmlTempResponseNode.selectNodes(".//EMPLOYEDDETAILS/EMPLOYMENT")
            For Each xmlTempNode In xmlTempNodeList
                xmlTempNode.parentNode.removeChild xmlTempNode
            Next
            
            'Promote SELFEMPLOYEDDETAILS and CONTRACTDETAILS nodes
            Set xmlTempNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//SELFEMPLOYEDDETAILS", False)
            If Not xmlTempNode Is Nothing Then
                objXmlAssist.PromoteNode xmlTempResponseNode, xmlTempNode, True, True
            End If
            Set xmlTempNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//CONTRACTDETAILS", False)
            If Not xmlTempNode Is Nothing Then
                objXmlAssist.PromoteNode xmlTempResponseNode, xmlTempNode, True, True
            End If
            
            'BMIDS00052 MDC 28/06/2002 - Copy SHARESOWNEDINDICATOR to EMPLOYEDDETAILS node
            Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//EMPLOYEDDETAILS")
            'BMIDS00172 MO 03/07/2002 - xmlTempNode may not exist
            If Not xmlTempNode Is Nothing Then
                objXmlAssist.CopyNode "SHARESOWNEDINDICATOR", "SHARESOWNEDINDICATOR", xmlEmploymentNode, xmlTempNode, False
            End If
            'BMIDS00052 MDC 28/06/2002 - End
            
            'Update Customer Version Number with new Customer Version Number
            Set xmlCustomerNodelist = xmlTempResponseNode.selectNodes(".//CUSTOMERVERSIONNUMBER")
            For Each xmlCustomerNode In xmlCustomerNodelist
                xmlCustomerNode.Text = strNewCustomerVersionNumber
            Next
            'BMIDS00929 SA 14/11/2002 Need to get accountant details if there are any {
            If objXmlAssist.GetTagValue(xmlEmploymentNode, "ACCOUNTANTGUID") <> "" Then
                'Build up request to call GetAccountantDetails
                Set xmlAccountantRequest = vxmlRequest.cloneNode(False)
                Set xmlAccountantElem = xmlDoc.createElement("ACCOUNTANT")
                Set xmlAccountantGuid = xmlDoc.createElement("ACCOUNTANTGUID")
                
                xmlAccountantGuid.Text = objXmlAssist.GetTagValue(xmlEmploymentNode, "ACCOUNTANTGUID")
                xmlAccountantElem.appendChild xmlAccountantGuid
                xmlAccountantRequest.appendChild xmlAccountantElem
                
                Set xmlAccountantResponse = objICustomerEmploymentBO.GetAccountantDetails(xmlAccountantRequest)
                objErrAssist.CheckResponse xmlAccountantResponse, xmlResponseElem, True
                'Add details onto Employment
                Dim xmlEmpNode As IXMLDOMNode
                Set xmlEmpNode = xmlTempResponseNode.selectSingleNode("EMPLOYMENT")
                xmlEmpNode.appendChild xmlAccountantResponse.selectSingleNode("ACCOUNTANT")
            End If
            'BMIDS00929 }
            
            'BMIDS00265 MDC 30/07/2002
            'Remove Guids from Request so that new records are created
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/THIRDPARTY/THIRDPARTYGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            'BMIDS00693 SA Got to remove the ThirdPartyGuid held under the Employment node too. (updates Employment table)
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/THIRDPARTYGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            'BMIDS00693 }
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/THIRDPARTY/ADDRESS/ADDRESSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/CONTACTDETAILSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNodeList = xmlTempResponseNode.selectNodes(".//EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS/CONTACTDETAILSGUID")
            For Each xmlGUIDNode In xmlGUIDNodeList
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            Next
            'BMIDS00265 MDC 30/07/2002 - End
            'BMIDS00929 SA 14/11/02 Remove any guids associated with accountants so new ones will be created {
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/THIRDPARTYGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTYGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/ADDRESSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/ADDRESS/ADDRESSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/CONTACTDETAILSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNode = objXmlAssist.GetNode(xmlTempResponseNode, ".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/CONTACTDETAILS/CONTACTDETAILSGUID", False)
            If Not xmlGUIDNode Is Nothing Then
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            End If
            Set xmlGUIDNodeList = xmlTempResponseNode.selectNodes(".//EMPLOYMENT/ACCOUNTANT/THIRDPARTY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS/CONTACTDETAILSGUID")
            For Each xmlGUIDNode In xmlGUIDNodeList
                xmlGUIDNode.parentNode.removeChild xmlGUIDNode
            Next
            'BMIDS00929 SA 14/11/02 }
            'Call CreateEmploymentDetails (for the new Customer Version)
            Set xmlTempResponseNode = objICustomerEmploymentBO.CreateEmploymentDetails(xmlTempResponseNode)
            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
            'Save sequence number value for response
            Set xmlTempNode = objXmlAssist.GetMandatoryNode(xmlTempResponseNode, "GENERATEDKEYS")
            objXmlAssist.CopyNode "EMPLOYMENTSEQUENCENUMBER", "EMPLOYMENTSEQUENCENUMBER", xmlTempNode, xmlKeysNode, False
            
        Next
    End If
    
    'Return key values
    xmlResponseElem.appendChild xmlKeysNode

    Set ICustomerEmploymentTxBO_CopyEmploymentDetails = xmlResponseElem
    objContext.SetComplete
    
CopyEmploymentDetailsExit:

    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlEmploymentNode = Nothing
    Set xmlEmploymentListNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempNodeList = Nothing
    Set xmlEmploymentNodeList = Nothing
    Set xmlCustomerNodelist = Nothing
    Set xmlCustomerNode = Nothing
    Set objICustomerEmploymentBO = Nothing
    
    'BMIDS00265 MDC 31/07/2002
    Set xmlGUIDNode = Nothing
    Set xmlGUIDNodeList = Nothing
    'BMIDS00265 MDC 31/07/2002 - End
    
    'BMIDS00929 SA 14/11/02 {
    Set xmlAccountantRequest = Nothing
    Set xmlAccountantElem = Nothing
    Set xmlAccountantGuid = Nothing
    Set xmlAccountantResponse = Nothing
    'BMIDS00929 }
    
    Exit Function

CopyEmploymentDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CopyEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    objContext.SetAbort
    
    ' go to clean-up section
    Resume CopyEmploymentDetailsExit

End Function
'BMIDS00004 MDC 17/05/2002 - End

Private Function ICustomerEmploymentTxBO_CreateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateAccountantDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateAccountantDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "ACCOUNTANT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("ACCOUNTANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "ACCOUNTANT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If

    ' Creat a 'GENERATEDKEYS' child node to the main response XML
    Set xmlKeysNode = xmlResponseElem.appendChild(xmlOut.createElement("GENERATEDKEYS"))
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateAccountantDetails xmlElement, xmlKeysNode
    Else
        ' For each relevant node call the data object CreateAccountantDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateAccountantDetails xmlElement, xmlKeysNode
        Next
    End If
    
    Set ICustomerEmploymentTxBO_CreateAccountantDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateAccountantDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateAccountantDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateAccountantDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateContractDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateContractDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "CONTRACTDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CONTRACTDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CONTRACTDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateContractDetails xmlElement
    Else
        ' For each relevant node call the data object CreateContractDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateContractDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_CreateContractDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateContractDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateContractDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateContractDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateContractDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateEmployedDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateEmployedDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EMPLOYEDDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYEDDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EMPLOYEDDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateEmployedDetails xmlElement
    Else
        ' For each relevant node call the data object CreateEmployedDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateEmployedDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_CreateEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateEmployedDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEmployedDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateEmploymentDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateEmploymentDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYMENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EMPLOYMENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If

    ' Creat a 'GENERATEDKEYS' child node to the main response XML
    Set xmlKeysNode = xmlResponseElem.appendChild(xmlOut.createElement("GENERATEDKEYS"))
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateEmploymentDetails xmlElement, xmlKeysNode
    Else
        ' For each relevant node call the data object CreateEmploymentDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateEmploymentDetails xmlElement, xmlKeysNode
        Next
    End If
    
    Set ICustomerEmploymentTxBO_CreateEmploymentDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateEmploymentDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateEmploymentDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateNetProfitDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateNetProfitDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "NETPROFIT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("NETPROFIT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "NETPROFIT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateNetProfitDetails xmlElement
    Else
        ' For each relevant node call the data object CreateNetProfitDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateNetProfitDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "NETMONTHLYINCOME"
    Set ICustomerEmploymentTxBO_CreateNetProfitDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNetProfitDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateNetProfitDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateNetProfitDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNetProfitDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateOtherIncome multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateOtherIncome"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "UNEARNEDINCOME" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("UNEARNEDINCOME")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "UNEARNEDINCOME" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateOtherIncome xmlElement
    Else
        ' For each relevant node call the data object CreateOtherIncome
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateOtherIncome xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_CreateOtherIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOtherIncomeExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateOtherIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOtherIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_CreateSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateSelfEmployedDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateSelfEmployedDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "SELFEMPLOYEDDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("SELFEMPLOYEDDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "SELFEMPLOYEDDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateSelfEmployedDetails xmlElement
    Else
        ' For each relevant node call the data object CreateSelfEmployedDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateSelfEmployedDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_CreateSelfEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateSelfEmployedDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateSelfEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateSelfEmployedDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_CreateTaxDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateTaxDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_CreateTaxDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "INCOMESUMMARY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("INCOMESUMMARY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "INCOMESUMMARY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.CreateTaxDetails xmlElement
    Else
        ' For each relevant node call the data object CreateTaxDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.CreateTaxDetails xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_CreateTaxDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateTaxDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateTaxDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_CreateTaxDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateTaxDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_DeleteEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteEmploymentDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_DeleteEmploymentDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYMENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EMPLOYMENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.DeleteEmploymentDetails xmlElement
    Else
        ' For each relevant node call the data object DeleteEmploymentDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.DeleteEmploymentDetails xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_DeleteEmploymentDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteEmploymentDetailsExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteEmploymentDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_DeleteEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteEmploymentDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_DeleteOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteOtherIncome multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_DeleteOtherIncome"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "UNEARNEDINCOME" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("UNEARNEDINCOME")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "UNEARNEDINCOME" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.DeleteOtherIncome xmlElement
    Else
        ' For each relevant node call the data object DeleteOtherIncome
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.DeleteOtherIncome xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_DeleteOtherIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteOtherIncomeExit:

    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteOtherIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_DeleteOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteOtherIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_SaveCurrentEmploymentAndIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveCurrentEmploymentAndIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_SaveCurrentEmploymentAndIncome"

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
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIClassDef As IomCEClassDef
    Dim botOperationType As BOOPERATIONTYPE

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Initialise
    '
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYMENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EMPLOYMENT" & " tag not found"
        End If
    End If
   
    Set objIClassDef = New omCEClassDef
    
    'save EMPLOYMENTDETAILS
    Set xmlClassDefDoc = objIClassDef.LoadEmploymentData
    Set xmlElement = vxmlRequest.getElementsByTagName("EMPLOYMENT").Item(0)

    If Not (xmlElement Is Nothing) Then
        ' Call the appropriate DO method for this element
        botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
        Select Case botOperationType
            Case booUpdate: Set xmlTempResponseNode = ICustomerEmploymentTxBO_UpdateEmploymentDetails(vxmlRequest)   'JLD
                            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            Case booCreate: Set xmlTempResponseNode = ICustomerEmploymentTxBO_CreateEmploymentDetails(xmlElement)
                            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                            ' SR : 07/08/00 - Add the generated keys to the Main Response
                            xmlResponseElem.appendChild xmlTempResponseNode.selectSingleNode(".//GENERATEDKEYS")
            Case booDelete: Set xmlTempResponseNode = ICustomerEmploymentTxBO_DeleteEmploymentDetails(xmlElement)
                            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            Case Else
                ' N/A
        End Select
    End If
    
    ' SR 29/08/00 SYS1420 - Employed details has already been created/updated by Create/Update Employment Details.
    '                       Do not do that again
'    'save EmployedDetails
'    Set xmlClassDefDoc = objIClassDef.LoadEmployedDetailsData
'    Set xmlElement = vxmlRequest.getElementsByTagName("EMPLOYEDDETAILS").Item(0)
'
'    If Not (xmlElement Is Nothing) Then
'        ' Call the appropriate DO method for this element
'        botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
'        Select Case botOperationType
'            Case booUpdate: Set xmlTempResponseNode = ICustomerEmploymentTxBO_UpdateEmployedDetails(xmlElement)
'                            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'            Case booCreate: Set xmlTempResponseNode = ICustomerEmploymentTxBO_CreateEmployedDetails(xmlElement)
'                            objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'            Case booDelete: 'xmlTempResponseNode = ICustomerEmploymentTxBO_DeleteEmployedDetails(xmlElement)
'                            'objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'            Case Else
'                ' N/A
'        End Select
'    End If
    
    'save other income details
    'loop around a list of entries
    Set xmlElement = vxmlRequest.getElementsByTagName("UNEARNEDINCOMELIST").Item(0)
    'SYS1368 MC 02/08/2000. Check that UNEARNEDINCOMELIST has been found
    If Not xmlElement Is Nothing Then
        Set xmlNodeList = xmlElement.getElementsByTagName("UNEARNEDINCOME")
        Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData
        Dim intNodeIndex As Integer
        Dim xmlUnearnedIncomeElem As IXMLDOMElement
        
        For intNodeIndex = 0 To xmlNodeList.length - 1
            Set xmlUnearnedIncomeElem = xmlNodeList.Item(intNodeIndex)
            
            'Select Case objXmlAssist.DetermineOperationEx(vxmlRequest, objIClassDef.LoadUnearnedIncomeData)
            Select Case objXmlAssist.DetermineOperationEx(xmlUnearnedIncomeElem, xmlClassDefDoc)
                Case booCreate:
                    Set xmlTempResponseNode = ICustomerEmploymentTxBO_CreateOtherIncome(xmlUnearnedIncomeElem)
                    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
                Case booUpdate:
                    Set xmlTempResponseNode = ICustomerEmploymentTxBO_UpdateOtherIncome(xmlUnearnedIncomeElem)
                    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                Case booDelete:
                    Set xmlTempResponseNode = ICustomerEmploymentTxBO_DeleteOtherIncome(xmlUnearnedIncomeElem)
                    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                
                Case Else
                    ' N/A
            End Select
        Next
    End If
    
    Set ICustomerEmploymentTxBO_SaveCurrentEmploymentAndIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveCurrentEmploymentAndIncomeExit:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlClassDefDoc = Nothing
    
    Exit Function

SaveCurrentEmploymentAndIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_SaveCurrentEmploymentAndIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveCurrentEmploymentAndIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_SaveEarnedIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveEarnedIncome multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_SaveEarnedIncome"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objIClassDef As IomCEClassDef
    Dim botOperationType As BOOPERATIONTYPE

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Initialise
    '
    If vxmlRequest.nodeName = "EARNEDINCOME" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EARNEDINCOME")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EARNEDINCOME" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
   
    '
    ' Main
    '
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData

    If Not (xmlElement Is Nothing) Then
        ' Call the appropriate DO method for this element
        botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
        Select Case botOperationType
            Case booUpdate: objICustomerEmploymentDO.UpdateEarnedIncome xmlElement
            Case booCreate: objICustomerEmploymentDO.CreateEarnedIncome xmlElement
            Case booDelete: ' Not valid - deleted via DeleteEmployment
                            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                                "EARNEDINCOME cannot be deleted in this way"
            Case Else
                ' N/A
        End Select
    Else
        ' For each relevant node call the appropriate data object method
        For Each xmlElement In xmlNodeList
            botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
            Select Case botOperationType
                Case booUpdate: objICustomerEmploymentDO.UpdateEarnedIncome xmlElement
                Case booCreate: objICustomerEmploymentDO.CreateEarnedIncome xmlElement
                Case booDelete: ' Not valid - deleted via DeleteEmployment
                                objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                                    "EARNEDINCOME cannot be deleted in this way"
                Case Else
                    ' N/A
            End Select
        Next
    End If
    
    UpdateEmploymentField vxmlRequest, "NETMONTHLYINCOME"
    Set ICustomerEmploymentTxBO_SaveEarnedIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveEarnedIncomeExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlClassDefDoc = Nothing
    
    Exit Function

SaveEarnedIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_SaveEarnedIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveEarnedIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateAccountantDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAccountantDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateAccountantDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "ACCOUNTANT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("ACCOUNTANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "ACCOUNTANT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateAccountantDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateAccountantDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateAccountantDetails xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_UpdateAccountantDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateAccountantDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateAccountantDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateAccountantDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateAccountantDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateContractDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContractDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateContractDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "CONTRACTDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CONTRACTDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CONTRACTDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateContractDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateContractDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateContractDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_UpdateContractDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateContractDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateContractDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateContractDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateContractDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateEarnedIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateEarnedIncome multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEarnedIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateEarnedIncome"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EARNEDINCOME" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EARNEDINCOME")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EARNEDINCOME" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateEarnedIncome xmlElement
    Else
        ' For each relevant node call the data object UpdateEarnedIncome
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateEarnedIncome xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_UpdateEarnedIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEarnedIncomeExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateEarnedIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateEarnedIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEarnedIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateEmployedDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateEmployedDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EMPLOYEDDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYEDDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "EMPLOYEDDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateEmployedDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateEmployedDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateEmployedDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_UpdateEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEmployedDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEmployedDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo UpdateEmploymentDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateEmploymentDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim objGlobalParameterBO As GlobalParameterBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "EMPLOYMENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("EMPLOYMENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "EMPLOYMENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
        Set objGlobalParameterBO = New GlobalParameterBO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
        Set objGlobalParameterBO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    End If
    
    ' Fetch the data required for deleting the previous employment, if required
    Dim strDeletePreviousEmployment As String, intEmploymentValidationYears As Integer
    strDeletePreviousEmployment = objXmlAssist.GetTagValue(vxmlRequest, "DELETEPREVIOUSEMPLOYMENT")

    If strDeletePreviousEmployment = "1" Then
        Dim strResponse As String
        strResponse = objGlobalParameterBO.GetCurrentParameter("EmploymentValidationYears")
        Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

        objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        intEmploymentValidationYears = xmlDoc.selectSingleNode("//AMOUNT").Text
    End If
    
    Dim strEmploymentStatus As String, strNewEmploymentStatus As String
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strEmpSequenceNumber As String, strMaintainStatusChange As String
    
    strMaintainStatusChange = objXmlAssist.GetTagValue(vxmlRequest, "MAINTAINSTATUSCHANGE")
    
    If Not xmlElement Is Nothing Then
        ' SR 18/08/00 SYS1420 : if the employment status is going to change, remove related
        '                       records from all the tables linked to Employment
        Set xmlNode = objICustomerEmploymentDO.GetEmploymentData(vxmlRequest)
        strEmploymentStatus = xmlNode.selectSingleNode(".//EMPLOYMENTSTATUS").Text
        
        Set xmlNode = vxmlRequest.selectSingleNode(".//EMPLOYMENTSTATUS")
        If Not xmlNode Is Nothing Then
            If strEmploymentStatus <> xmlNode.Text Then
                'Check whether the related records from all the liked tables are to be deleted (based on
                'the tag MAINTAINSTATUSCHANGE). If required delete.
                
                If strMaintainStatusChange = "1" Then
                    strCustomerNumber = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
                    strCustomerVersionNumber = vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
                    strEmpSequenceNumber = vxmlRequest.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
                    
                    DeleteEmploymentLinkedData strCustomerNumber, strCustomerVersionNumber, _
                                               strEmpSequenceNumber, strEmploymentStatus
                End If
            End If
        End If
        
        objICustomerEmploymentDO.UpdateEmploymentDetails xmlElement
        ' SR 01/08/00 - Delete the Employment details with higher employment sequence number, if required
        
        If strDeletePreviousEmployment = "1" Then
            DeletePreviousEmployments xmlElement, vxmlRequest, intEmploymentValidationYears
        End If
    Else
        ' For each relevant node call the data object UpdateEmploymentDetails
        For Each xmlElement In xmlNodeList
        
            ' SR 18/08/00 SYS1420 : if the employment status is going to change, remove related
            '                       records from all the tables linked to Employment
            Set xmlNode = objICustomerEmploymentDO.GetEmploymentData(vxmlRequest)
            strEmploymentStatus = xmlNode.selectSingleNode(".//EMPLOYMENTSTATUS").Text
        
            Set xmlNode = vxmlRequest.selectSingleNode(".//EMPLOYMENTSTATUS")
            If Not xmlNode Is Nothing Then
                If strEmploymentStatus <> xmlNode.Text Then
                    If strMaintainStatusChange = "1" Then
                        strCustomerNumber = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
                        strCustomerVersionNumber = vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
                        strEmpSequenceNumber = vxmlRequest.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
                        
                        DeleteEmploymentLinkedData strCustomerNumber, strCustomerVersionNumber, _
                                                   strEmpSequenceNumber, strEmploymentStatus
                    End If
                End If
            End If
        
            objICustomerEmploymentDO.UpdateEmploymentDetails xmlElement
            If strDeletePreviousEmployment = "1" Then
                DeletePreviousEmployments xmlElement, vxmlRequest, intEmploymentValidationYears
            End If
        Next
    End If
    
    Set ICustomerEmploymentTxBO_UpdateEmploymentDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateEmploymentDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objGlobalParameterBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateEmploymentDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateEmploymentDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateEmploymentDetailsExit

End Function

Private Sub DeleteEmploymentLinkedData(ByVal strCustomerNumber As String, _
                                       ByVal strCustomerVersionNumber As String, _
                                       ByVal strEmpSequenceNumber As String, _
                                       ByVal strEmploymentStatus As String)
' header ----------------------------------------------------------------------------------
' Description: Deletes all the records liked to a particular record in Employment.
'              Called from UpdateEmployment when Employment status is changed during the
'              of Employment data.
'------------------------------------------------------------------------------------------
On Error GoTo errHandle
    
    Const strFunctionName As String = "DeleteEmploymentLinkedData"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlTableElem As IXMLDOMNode, xmlElement As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument40, xmlClassDefDoc As FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim objIDOAssist As IDOAssist
    Dim objComboDO As ComboDO
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIClassDef As IomCEClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objComboDO = New ComboDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    ' Table : NETPROFIT
    Set xmlTableElem = xmlDoc.createElement("NETPROFIT")
    
    Set xmlElement = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCustomerNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCustomerVersionNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("EMPLOYMENTSEQUENCENUMBER")
    xmlElement.Text = strEmpSequenceNumber
    xmlTableElem.appendChild xmlElement
    
    Set objIClassDef = New omCEClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNetProfitData
    
    objIDOAssist.Delete xmlTableElem, xmlClassDefDoc

    'Table : EARNEDINCOME
    objXmlAssist.ChangeNodeName xmlTableElem, "NETPROFIT", "EARNEDINCOME"
    Set xmlClassDefDoc = objIClassDef.LoadEarnedIncomeData

     objIDOAssist.DeleteAll xmlTableElem, xmlClassDefDoc

    'Table : UNEARNEDINCOME
    objXmlAssist.ChangeNodeName xmlTableElem, "EARNEDINCOME", "UNEARNEDINCOME"
    Set xmlClassDefDoc = objIClassDef.LoadUnearnedIncomeData
    
    objIDOAssist.DeleteAll xmlTableElem, xmlClassDefDoc
     
    ' Table : EMPLOYEDDETAILS / SELF-EMPLOYEDDETAILS
    Dim strEmpValidation As String
    strEmpValidation = objComboDO.GetFirstComboValidation("EmploymentStatus", strEmploymentStatus)
        
    Select Case strEmpValidation
        Case "E"        ' Employed
            objXmlAssist.ChangeNodeName xmlTableElem, "UNEARNEDINCOME", "EMPLOYEDDETAILS"
            Set xmlClassDefDoc = objIClassDef.LoadEmployedDetailsData
            objIDOAssist.Delete xmlTableElem, xmlClassDefDoc
        
        Case "S", "C" 'Self employed
            objXmlAssist.ChangeNodeName xmlTableElem, "UNEARNEDINCOME", "SELFEMPLOYEDDETAILS"
            Set xmlClassDefDoc = objIClassDef.LoadSelfEmployedDetailsData
            objIDOAssist.Delete xmlTableElem, xmlClassDefDoc
        
        Case "R"        ' Retired
            ' For Homezone, related data for 'Retired' is stored in UnearnedIncome,
            ' this was already taken care of
            
    End Select
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub
errHandle:
    Dim lngErrNo As Long
    lngErrNo = objErrAssist.GetOmigaErrorNumber(Err.Number)
    If lngErrNo = oeNoRowsAffected Or lngErrNo = oeNoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    
    Set objIDOAssist = Nothing
    Set objIClassDef = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' Re-raise the error
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub DeletePreviousEmployments(ByVal objElement As IXMLDOMElement, _
                           ByVal vxmlRequest As IXMLDOMElement, _
                           ByVal intEmploymentValidationYears As Integer)
'--------------------------------------------------------------------------------------
'Description :
'--------------------------------------------------------------------------------------
On Error GoTo DeleteErr
    
    Const strFunctionName As String = "DeletePreviousEmployments"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim objIClassDef As IomCEClassDef
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist, objErrAssist As New ErrAssist

    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strEmploymentSequenceNumber As String
    Dim strStartDate As String, dtStartDate As Date, lngYears As Long
    
    Dim xmlEmploymentElem As IXMLDOMNode, xmlElem As IXMLDOMNode, xmlRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode, xmlResponseElem As IXMLDOMElement, xmlTempResponse As IXMLDOMElement
    Dim xmlEmploymentList As IXMLDOMNodeList
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
        
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strStartDate = objXmlAssist.GetTagValue(objElement, "DATESTARTEDORESTABLISHED")
    If strStartDate = "" Then GoTo DeleteExit
        
    'Calculate the years lapsed between system date and dateMovedIn
    dtStartDate = CDate(strStartDate)
    lngYears = DateDiff("YYYY", dtStartDate, Date)
    dtStartDate = DateSerial(Year(Date), Month(dtStartDate), Day(dtStartDate))
    If dtStartDate > Date Then lngYears = lngYears - 1
    ' if years lapsed is less than the global parameter (AddressValidationYears)
    If lngYears < intEmploymentValidationYears Then GoTo DeleteExit
    
    ' Find the employment records which have higher sequence than the current one
    strCustomerNumber = objElement.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustomerVersionNumber = objElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    strEmploymentSequenceNumber = objElement.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
    
    If strCustomerNumber = "" Or strCustomerVersionNumber = "" Or strEmploymentSequenceNumber = "" Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, "Employment key data not found"
    End If
    
    Dim strAddlCondition As String
    strAddlCondition = "EMPLOYMENTSEQUENCENUMBER > " & strEmploymentSequenceNumber
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    ' Build request to fetch / delete Employment records
    Set xmlRequest = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlEmploymentElem = xmlDoc.createElement("EMPLOYMENT")
    xmlRequest.appendChild xmlEmploymentElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElem.Text = strCustomerNumber
    xmlEmploymentElem.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlElem.Text = strCustomerVersionNumber
    xmlEmploymentElem.appendChild xmlElem
    
    Set objIClassDef = New omCEClassDef
    Set xmlTempResponse = objIDOAssist.FindListMultiple(xmlRequest, objIClassDef.LoadEmploymentData, _
                                        vstrAdditionalCodition:=strAddlCondition)
    
    ' SR SYS1383: If no records are found, exit the function, but do not raise error
    If xmlTempResponse Is Nothing Then GoTo DeleteExit
        
    Set xmlEmploymentList = xmlTempResponse.selectNodes("EMPLOYMENT")
    If xmlEmploymentList.length > 0 Then
        Set xmlElem = xmlDoc.createElement("EMPLOYMENTSEQUENCENUMBER")
        xmlEmploymentElem.appendChild xmlElem
    End If
    Dim strResponse As String
    For Each xmlNode In xmlEmploymentList
        xmlElem.Text = xmlNode.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
        objICustomerEmploymentDO.DeleteEmploymentDetails xmlEmploymentElem
        
    Next xmlNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteExit:
    Set objIDOAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set objIClassDef = Nothing
    
    Set xmlEmploymentElem = Nothing
    Set xmlEmploymentList = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    
    Set xmlDoc = Nothing
    
    Exit Sub
DeleteErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    Set objIDOAssist = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set objIClassDef = Nothing
    
    Set xmlEmploymentElem = Nothing
    Set xmlEmploymentList = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlDoc = Nothing

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise the error
    Err.Raise Err.Number, Err.Source, Err.Description
         
End Sub

Private Function ICustomerEmploymentTxBO_UpdateNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateNetProfitDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNetProfitDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateNetProfitDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "NETPROFIT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("NETPROFIT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "NETPROFIT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateNetProfitDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateNetProfitDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateNetProfitDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "NETMONTHLYINCOME"
    Set ICustomerEmploymentTxBO_UpdateNetProfitDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateNetProfitDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateNetProfitDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateNetProfitDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNetProfitDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateOtherIncome multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherIncomeVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateOtherIncome"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "UNEARNEDINCOME" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("UNEARNEDINCOME")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "UNEARNEDINCOME" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateOtherIncome xmlElement
    Else
        ' For each relevant node call the data object UpdateOtherIncome
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateOtherIncome xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_UpdateOtherIncome = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateOtherIncomeExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateOtherIncomeVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateOtherIncome = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateOtherIncomeExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateSelfEmployedDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSelfEmployedDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateSelfEmployedDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "SELFEMPLOYEDDETAILS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("SELFEMPLOYEDDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "SELFEMPLOYEDDETAILS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateSelfEmployedDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateSelfEmployedDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateSelfEmployedDetails xmlElement
        Next
    End If

    UpdateEmploymentField vxmlRequest, "ACCOUNTANTGUID"
    Set ICustomerEmploymentTxBO_UpdateSelfEmployedDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateSelfEmployedDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateSelfEmployedDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateSelfEmployedDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateSelfEmployedDetailsExit

End Function

Private Function ICustomerEmploymentTxBO_UpdateTaxDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateTaxDetails multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTaxDetailsVbErr

    Const strFunctionName As String = "ICustomerEmploymentTxBO_UpdateTaxDetails"

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
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "INCOMESUMMARY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("INCOMESUMMARY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "INCOMESUMMARY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
    Else
        Set objICustomerEmploymentDO = objContext.CreateInstance(App.Title & ".CustomerEmploymentDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objICustomerEmploymentDO.UpdateTaxDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateTaxDetails
        For Each xmlElement In xmlNodeList
            objICustomerEmploymentDO.UpdateTaxDetails xmlElement
        Next
    End If
    
    Set ICustomerEmploymentTxBO_UpdateTaxDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateTaxDetailsExit:
    
    Set objICustomerEmploymentDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateTaxDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerEmploymentTxBO_UpdateTaxDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateTaxDetailsExit

End Function


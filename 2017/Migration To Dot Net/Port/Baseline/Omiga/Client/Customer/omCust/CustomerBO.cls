VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Customer Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (CustomerTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. CustomerTxBO, CustomerDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
' SR    15-11-99    New
'JLD    30/11/99    Set return value from SaveAlias()
'                   Get global parameters MinAge and MaxAge using GlobalParameterDO.GetCurrentParameterByType()
'                   DateAdd() function to use "yyyy" for year additions
'SR     09-12-99    Modified methods SaveCustomerAddress, CreateCustomerAddress so that they
'                   return the key values generated
'SR     17-02-00    Modified method ValidateCustomerTelephone - AQR - SYS0192
'IW     22-02-00    SYS0290 - incorrect call to ErrAssist function (returning string instead of expected object)
'                       Revised code from: ICustomerBO_SaveCustomerTelephone = objErrAssist.CreateErrorResponse
'                       To: Set ICustomerBO_SaveCustomerTelephone = objErrAssist.CreateErrorResponseEx
'SR     28-02-00    AQR - SYS0085. Added Private Sub 'ValidateTelephone' and modified
'                   'IcustomerBO_ValidateCustomerTelephone' to call this
'SR     07-02-00    Modified methods validateTelephoneNumber
'SR     13/03/00    Check whether 'Locking is required in this system' in methods in CreateCustomerLock
'                   and DeleteCustomerLock
'SR     23/03/00    New Method 'CreateFullCustomerDetails'.
'SR     02/05/00    New Method 'UpdateCustomerVersion'
'BG     24/05/00    New Method 'GetPersonalDetails'.
'SR     26/05/00    New Method 'CreateCustomerDetailsAndEmpt'
'SR     26/05/00    Method 'UpdatePersonalDetailsAndEmpt' is complete
'BG     01/06/00    Method 'SavePersonalDetailsAndAddress' added
'MC     02/06/00    Added SaveTaxDetails method
'SR     02/06/00    Added an optional parameter to method 'FindCustomerAddressList' used to sort
'                   the output in the order of CustomerAddressSequenceNumber
'BG     07/06/00    Added method GetPersonalDetailsAndAddress which Gets customer address and personal details
'                   from CustomerBO.FindCustomerAddressList and CustomerBO.GetPersonalDetails
'SR     03/07/00    New method - SaveTenancy
'BG     26/07/00    Added CreatePropertyInsuranceDetails
'BG     27/07/00    Added GetPropertyInsuranceDetails method
'SR     28/07/00    Modified method SaveDependant - moved the fucntionality to TxBO so that it can cater to
'                   saving multiple records
'SR     01/08/00    SYS1360 - Modified ValidateCustomerDetails - Call ErrAssist.ThrowError rather than RaiseError
'BG     03/08/00    Not an error for there to be no records
'SR     16/08/00    SYS1174 - make UserId attribute of the request tag optional. This is so in HomeZone
'SR     13/09/00    Modified error-handling in FindAliasList : Do not abort the transaction, if no record is
'                   found
'SR     20/10/00    SYS1645 - Modify validateCustomerDetails (Private sub). Now, the Age validation is same both
'                   for createCustomerDetails and updatePersonalDetails
'AS     27/11/00    CORE000019 : ACTION attribute removed OPERATION attribute added
'CL     21/12/00    SYS1755 Changes made to implement Contact History
'PSC    09/01/01    SYS1750 Amendments to FindCustomer to enable search on
'                   OtherSystemCustomerNumber and combine searching on an administration system.
'                   Add GetCustomerFromLegacyNumber()
'PSC    22/01/01    SYS1866 Correct default validation Tag Value to be CUSTOMER
'PSC    20/02/01    SYS1750 Amend FindCustomer to return the number of customers found
'PSC    20/02/01    SYS1750 Amend FindCustomer return the CustomerVersionNumber and Salutation correctly
'PSC    21/02/01    SYS1750 Amend GetAndSynchroniseCustomerDetails to clone the request node rather than
'                   creating one
'PSC    27/02/01    SYS1750 Amend GetAndSynchroniseCustomerDetails to pass down other system customer number
'                   as customer number for AdminBO call
'PSC    27/02/01    SYS1750 Amend FindCustomer to copy customer version accross if omiga and admin customers
'                   match and admin takes priority
'CL     09/03/01    SYS1885 amend FindContactHistoryList to show new Surname sort order
'CL     20/03/01    SYS2116 amend FindContactHistoryList to sort correctly
'SR     05/06/01    SYS2325 New method 'UpdateCustomer'
'SR     13/06/01    SYS2362 Modified methods 'ICustomerBO_FindCustomer' and 'ICustomerBO_GetAndSynchroniseCustomerDetails'
'DRC    10/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'JR     03/12/01    SYS2459 (omiplus24) re-applied telephone changes to ValidateTelephone and FindCustomerValidate
'JLD    29/01/02    SYS3921 Error in FindContactHistoryList.
'SG     26/03/02    SYS4317 Amended input XML for ODITransformer call within ICustomerBO_GetAndSynchroniseCustomerDetails
'DRC    10/04/02    SYS4373 - Check the AdminSystemState in FindCustomer
'DS     29/05/02    SYS4788 If there is no Omiga customer list, don't try and merge as the list object is set to nothing.
'HMA    12/10/04    BM0539  Display useful error message when duplicate customer lock is attempted.
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5    Add new function GetNumberOfCopiesForKFI
'HM     08/09/2005  MAR29   ICustomerTxBO_SaveWrapUpDetails, Save/ValidateWrapUpDetails were added
'HMA    10/10/2005  MAR135  Update SaveTaxDetails.
'PSC    06/10/2005  MAR57   Amend FindCustomer to allow search of the Omiga database to be optional
'                           Amend UpdateCRSCustomer to cater for customers only request
'PSC                        Amend GetAndSynchroniseCustomerDetails to do data check
'MV     20/10/2005  MAR148  Amended ICustomerBO_SaveTaxDetails()
'HMA    20/10/2005  MAR23   Added GetCustomerNumbers. Added CreateKYCCheck
'MV     21/10/2005  MAR148  Amended ICustomerBO_SaveTaxDetails()
'PSC    25/10/2005  MAR300  Correct processing of critical data check response in RunCustomerDataCheckAndUpdate
'PSC    16/11/2005  MAR590  Correct UpdateCRSCustomer to retrieve response correctly
'PSC    29/11/2005  MAR727  Amend RunCustomerDataCheckAndUpdate to use OTHERSYSTEMCUSTOMERNUMBER
'                           Amend UpdateCustomerWithAdminDetails to make admin details non mandatory
'PSC    20/12/2005  MAR917  Correct GetAndSynchroniseCustomerDetails for multiple customers and correct
'                           RunCustomerDataCheckAndUpdate to process CDRules response
'GHun   23/12/2005  MAR947  Changed ProcessLaunchCustomer to handle addresses as per XSD
'PSC    12/01/2006  MAR1031 Amend RunCustomerDataCheckAndUpdate to cater for no address changes
'PSC    24/01/2006  MAR1102 Amend RunCustomerDataCheckAndUpdate to update last amended data and time
'PSC    17/02/2006  MAR1211 Amend UpdateCustomerWithAdminDetails to include new fields
'                           Amend RunCustomerDataCheckAndUpdate to always update customer data from Admin
'PSC    17/02/2006  MAR1211 Amend RunCustomerDataCheckAndUpdate to detect changed customers correctly
'PSC    03/04/2006  MAR1573 Amend CreateKYCCheck to use CRUD
'------------------------------------------------------------------------------------------
'EPSOM Specific History:
'
'Prog   Date        Description
'PE     09/06/2006  EP702   Changed ICustomerBO_ValidateCustomerDetails to use 'Date' instead of 'Now'
'HMA    12/07/2006  EP903   Use Age on Application Date to verify Customer Age
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerBO
Implements ObjectControl

Private Const cstrTABLE_ALIASNAME = "ALIAS"
Private Const cstrTABLE_TENANCYNAME = "TENANCY"
Private Const cstrTABLE_DEPENDANTNAME = "DEPENDANT"
Private Const cstrTABLE_CUSTADDRESSNAME = "CUSTOMERADDRESS"
Private Const cstrROOT_NODE_NAME = "CUSTOMER"
Private Const cstrROOT_CONTACTHISTORY = "CONTACTHISTORY"
Private Const cstrTABLE_REMOVEDTOECUSTOMER = "REMOVEDTOECUSTOMER"   ' BMIDS758

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Private Const cstrROOT_CONTACTLOG = "CONTACTLOG"

Public Function SaveTaxDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveTaxDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveTaxDetailsVbErr

    Const strFunctionName As String = "SaveTaxDetails"

    
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
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveTaxDetails(xmlIn.documentElement, vblnDoValidation:=True)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveTaxDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveTaxDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveTaxDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveTaxDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveTaxDetailsExit
    
End Function

Public Function SaveTenancy(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the Tenancy Data after finding the operation to be performed on it
' pass:         vstrXMLRequest  xml Request data to be passed to the methos
' return:       nil
' Error :      Invalid Paramente
'------------------------------------------------------------------------------------------

On Error GoTo SaveTenancyVbErr

    Const strFunctionName As String = "SaveTenancy"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIScriptInterface As IScriptInterface


    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'If TenancyXML is NULL raise error
    If vstrXMLRequest = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If
    
    ' Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                                    strFunctionName, xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    SaveTenancy = xmlOut.xml

SaveTenancyExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function
SaveTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' go to clean-up section
    Resume SaveTenancyExit

End Function


Private Function ICustomerBO_CreateContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactHistoryVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateContactHistory"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
              
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateContactHistory(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set ICustomerBO_CreateContactHistory = xmlResponseElem

    m_objContext.SetComplete
    
CreateContactHistoryExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
     

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

    Set ICustomerBO_CreateContactHistory = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateContactHistoryExit

End Function

'BMIDS00004 MDC 13/05/2002
Private Function ICustomerBO_CreateNewCustomerVersion(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes ICustomerTxBO_CreateNewCustomerVersion
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewCustomerVersionVbErr
    
Const strFunctionName As String = "ICustomerBO_CreateNewCustomerVersion"


Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateNewCustomerVersion(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_CreateNewCustomerVersion = xmlTempResponseNode

    m_objContext.SetComplete
    
CreateNewCustomerVersionExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
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

    Set ICustomerBO_CreateNewCustomerVersion = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateNewCustomerVersionExit

End Function
'BMIDS00004 MDC 13/05/2002 - End

Private Function ICustomerBO_CreatePropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes ICustomerTxBO_CreatePropertyInsuranceDetails
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyInsuranceDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreatePropertyInsuranceDetails"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreatePropertyInsuranceDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_CreatePropertyInsuranceDetails = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreatePropertyInsuranceDetailsExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set ICustomerBO_CreatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePropertyInsuranceDetailsExit

End Function

Private Function ICustomerBO_FindContactHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindContactHistoryListVbErr
    
    Const strFunctionName As String = "ICustomerBO_FindContactHistoryList"
       
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlUserList As IXMLDOMElement
    Dim xmlUnitList As IXMLDOMElement
    Dim xmlUserDoc As FreeThreadedDOMDocument40
    Dim xmlUnitDoc As FreeThreadedDOMDocument40
    Dim objOrganisationBO As OrganisationBO
    Dim xmlTempResponse As IXMLDOMNode
    Dim xslDoc As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    Dim xmlContactHistoryList As IXMLDOMNode
    Dim cstrFunctionName As String
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
                
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
        
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
                
    If vxmlRequest.nodeName = cstrROOT_CONTACTHISTORY Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_CONTACTHISTORY & " tag not found"
    End If
             
    Set xmlDataNode = objICustomerDO.FindContactHistoryList(xmlRequestNode)
      
    Set xmlTempRequestNode = m_objXmlAssist.GetRequestNodeEx(vxmlRequest) 'JLD SYS3921
    
    Set xmlTempResponse = AddUserAndUnitToHistory(xmlTempRequestNode, xmlDataNode)
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True

    Set xmlContactHistoryList = m_objXmlAssist.GetNode(xmlTempResponse, ".//CONTACTHISTORYLIST")
              
    Dim intSortOrder As Integer
    
    intSortOrder = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//SORTORDER")
    
    If intSortOrder = 3 Then
        Set xslDoc = New FreeThreadedDOMDocument40
        xslDoc.validateOnParse = False
        xslDoc.setProperty "NewParser", True
        xslDoc.async = False

        strXslPath = App.Path & "\FindContactHistoryList.xslt"
        strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
        blnLoadedOk = xslDoc.load(strXslPath)
        
        If Not blnLoadedOk Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
        End If
        
        Set xmlTransformedDoc = New FreeThreadedDOMDocument40
        xmlTransformedDoc.validateOnParse = False
        xmlTransformedDoc.setProperty "NewParser", True
        xmlTransformedDoc.async = False
        
        Dim xmlContactHistoryDoc As FreeThreadedDOMDocument40
        Set xmlContactHistoryDoc = New FreeThreadedDOMDocument40
        xmlContactHistoryDoc.validateOnParse = False
        xmlContactHistoryDoc.setProperty "NewParser", True
        xmlContactHistoryDoc.async = False
        
        xmlContactHistoryDoc.loadXML (xmlContactHistoryList.xml)
        
        xmlContactHistoryDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
        xmlTempResponse.replaceChild xmlTransformedDoc.documentElement, xmlContactHistoryList
    End If
    
    m_objContext.SetComplete
    
    Set ICustomerBO_FindContactHistoryList = xmlTempResponse

FindContactHistoryListExit:

    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCopyNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlUserList = Nothing
    Set xmlUnitList = Nothing
    Set xmlUserDoc = Nothing
    Set xmlUnitDoc = Nothing
    Set objOrganisationBO = Nothing
    Set xmlTempResponse = Nothing
    
    

    Exit Function

FindContactHistoryListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_FindContactHistoryList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindContactHistoryListExit

End Function

Private Function ICustomerBO_GetContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetContactHistoryVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetContactHistory"

       
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlUserList As IXMLDOMElement
    Dim xmlUnitList As IXMLDOMElement
    Dim xmlUserDoc As FreeThreadedDOMDocument40
    Dim xmlUnitDoc As FreeThreadedDOMDocument40
    Dim objOrganisationBO As OrganisationBO
    Dim xmlTempResponse As IXMLDOMNode
            
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
        
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
                
    If vxmlRequest.nodeName = cstrROOT_CONTACTHISTORY Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_CONTACTHISTORY & " tag not found"
    End If
             
    Set xmlDataNode = objICustomerDO.GetContactHistory(xmlRequestNode)
      
    Set xmlTempRequestNode = m_objXmlAssist.GetRequestNodeEx(xmlRequestNode)
    
    Set xmlTempResponse = AddUserAndUnitToHistory(xmlTempRequestNode, xmlDataNode)
                   
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
   
    m_objContext.SetComplete
    
    Set ICustomerBO_GetContactHistory = xmlTempResponse

GetContactHistoryExit:

    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCopyNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlUserList = Nothing
    Set xmlUnitList = Nothing
    Set xmlUserDoc = Nothing
    Set xmlUnitDoc = Nothing
    Set objOrganisationBO = Nothing
    Set xmlTempResponse = Nothing
        
   

    Exit Function

GetContactHistoryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetContactHistory = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetContactHistoryExit

End Function

Private Function ICustomerBO_GetCustomerFromLegacyNumber(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerFromLegacyNumberVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetCustomerFromLegacyNumber"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetCustomerFromLegacyNumber(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerBO_GetCustomerFromLegacyNumber = xmlResponseElem
    
    m_objContext.SetComplete
    
GetCustomerFromLegacyNumberExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetCustomerFromLegacyNumberVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetCustomerFromLegacyNumber = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetCustomerFromLegacyNumberExit

End Function

Private Function ICustomerBO_GetCustomerNumbers(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    Set xmlDataNode = objICustomerDO.GetCustomerNumbers(vxmlRequest)

    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
        
    Set ICustomerBO_GetCustomerNumbers = xmlResponseElem
    
    Set objICustomerDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

End Function

Private Function ICustomerBO_GetPropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyInsuranceDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetPropertyInsuranceDetails"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetPropertyInsuranceDetails(xmlRequestNode)

'BG 03/08/00 Not an error for there to be no records
    
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
        
    Set ICustomerBO_GetPropertyInsuranceDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetPropertyInsuranceDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetPropertyInsuranceDetailsVbErr:

    'BG 03/08/00 Not an error for there to be no records

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetPropertyInsuranceDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPropertyInsuranceDetailsExit

End Function

Private Function ICustomerBO_SaveTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveCustomerTelephone a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request data stream containing data SaveCustomerTelephone
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveTenancyVbErr

    Const strFunctionName As String = "ICustomerBO_SaveTenancy"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call corresponding function in Business Transaction Object
    Set xmlTempResponseNode = objICustomerTxBO.SaveTenancy(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_SaveTenancy = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerBO_SaveTenancyExit:

    '   clean-up section, release object references
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing

    Exit Function

ICustomerBO_SaveTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ICustomerBO_SaveTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ICustomerBO_SaveTenancyExit
End Function

Public Function FindCustomer(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieve a list of customers matching search criteria
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       FindCustomer   xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element

'------------------------------------------------------------------------------------------
On Error GoTo FindCustomerVbErr
    
    Const strFunctionName As String = "FindCustomer"


    '   request block
    '   <REQUEST USERID="UserId" USERTYPE"UserType" UNIT="Unit" CHANNEL="Channel">
    '       <SEARCH>
    '           <CUSTOMER>
    '             <SURNAME>option</SURNAME>
    '             <FIRSTFORENAME>option</FIRSTFORENAME>
    '             <SECONDFORENAME>option</SECONDFORENAME>
    '             <OTHERFORENAME>option</OTHERFORENAME>
    '             <DATEOFBIRTH>option</DATEOFBIRTH>
    '             <POSTCODE>option</POSTCODE>
    '             <BUILDINGORHOUSENAME>option</BUILDINGORHOUSENAME>
    '             <BUILDINGORHOUSENUMBER>option</BUILDINGORHOUSENUMBER>
    '             <FLATNUMBER>option</FLATNUMBER>
    '             <CUSTOMERTELEPHONENUMBER>option</CUSTOMERTELEPHONENUMBER>
    '             <OTHERSYSTEMCUSTOMERNUMBER>option</OTHERSYSTEMCUSTOMERNUMBER>
    '           </CUSTOMER>
    '       </SEARCH>
    '   </REQUEST>
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlRequestElem As IXMLDOMElement
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    objXmlOut.async = False

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)

    Set objXmlRequestElem = objXmlAssist.GetMandatoryNode(objXmlIn, "REQUEST")

    ' Validate the search criteria
    FindCustomerValidate objXmlIn

    ' Get the User id
    Dim strUserId As String
    
    'SR SYS1174: User Id is not mandatory
    strUserId = objXmlAssist.GetAttributeValue(objXmlRequestElem, "REQUEST", "USERID")
    
    ' Reset the distribution channel if this user does not have
    ' their searches restricted
    
    If Len(strUserId) <> 0 Then
        If IsRestrictedByChannel(strUserId) = False Then
            objXmlRequestElem.setAttribute "CHANNEL", ""
        End If
    End If
    
    'BMIDS00410 Moved above the call to FindCustomer, otherwise warnings are overwritten
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    'BMIDS00410 End
    
    ICustomerBO_FindCustomer objXmlRequestElem, objXmlResponseNode

    FindCustomer = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindCustomerExit:

    Set objXmlIn = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlRequestElem = Nothing
    
    Set objXmlOut = Nothing
    Set objXmlResponseNode = Nothing
    Set objXmlResponseElem = Nothing
    
    Set objXmlAssist = Nothing
    
    Set objContext = Nothing
    
    Exit Function

FindCustomerVbErr:
    
    Dim objErrAssist As New ErrAssist

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    FindCustomer = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing

    '   go to clean-up section
    Resume FindCustomerExit

End Function

Public Function GetCustomerDetails(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Gets all the details for a particular customer including:
'               Customer Details, Customer Address Details, Telephone Details and
'               Areas of Interest
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       GetCustomerDetails  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------

On Error GoTo GetCustomerDetailsVbErr
    
    Const strFunctionName As String = "GetCustomerDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXMLElement As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    
    objXmlOut.async = False

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
       
    '-----------------------------------------------------------
    ' Perform hard coded validation
    ' See if Customer Number Customer Version Number are present
    '-----------------------------------------------------------
    
    Dim strCustomerNumber As String
    Dim strCustVersionNum As String
    
    strCustomerNumber = _
        objXmlAssist.GetMandatoryElementText(objXmlIn, "//CUSTOMER/CUSTOMERNUMBER")
        
    strCustVersionNum = _
        objXmlAssist.GetMandatoryElementText(objXmlIn, "//CUSTOMER/CUSTOMERVERSIONNUMBER")
    
    ICustomerBO_GetCustomerDetails objXmlIn.documentElement, objXmlResponseElem

    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    GetCustomerDetails = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetCustomerDetailsExit:
        
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlDataNode = Nothing
    Set objXMLElement = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    Exit Function

GetCustomerDetailsVbErr:

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetCustomerDetails = objErrAssist.CreateErrorResponse
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetCustomerDetailsExit
    
End Function

Public Function GetAndSynchroniseCustomerDetails(ByVal vstrXMLRequest) As String

On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
    
    Const strFunctionName As String = "GetAndSynchroniseCustomerDetails"

    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlIn As FreeThreadedDOMDocument40
    
    Dim objxmlRequestNode As IXMLDOMNode
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    objXmlOut.async = False

    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objxmlRequestNode = m_objXmlAssist.GetMandatoryNode(objXmlIn, "REQUEST")

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
           
    ICustomerBO_GetAndSynchroniseCustomerDetails objxmlRequestNode, objXmlResponseNode
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    GetAndSynchroniseCustomerDetails = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetAndSynchroniseCustomerDetailsExit:
    
    Set objxmlRequestNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlResponseNode = Nothing
    
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    
    
    Exit Function
    
GetAndSynchroniseCustomerDetailsVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    GetAndSynchroniseCustomerDetails = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetAndSynchroniseCustomerDetailsExit

End Function

Public Function CreateNewCustomerVersion(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a new customer version based on an existing one
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       CreateNewCustomerVersion  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of customer key data
'               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------

On Error GoTo CreateNewCustomerVersionVbErr
    
Const strFunctionName As String = "CreateNewCustomerVersion"

Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
Dim objErrAssist As New ErrAssist


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    'Initialise response
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    xmlIn.loadXML vstrXMLRequest
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateNewCustomerVersion(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewCustomerVersion = xmlTempResponseNode.xml
    objContext.SetComplete

CreateNewCustomerVersionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing

    Exit Function

CreateNewCustomerVersionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewCustomerVersion = objErrAssist.CreateErrorResponse

    objContext.SetAbort

    ' go to clean-up section
    Resume CreateNewCustomerVersionExit

End Function

Public Function CreateCustomerDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates the individual table entries for customer details
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       CreateCustomerDetails  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerDetailsVbErr
    
    Const strFunctionName As String = "CreateCustomerDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objxmlRequestNode As IXMLDOMNode
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    Dim objCustomerTxBO As ICustomerTxBO
    
    Dim strCustomerKeys As String
    
    objXmlOut.async = False

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objxmlRequestNode = objXmlAssist.GetMandatoryNode(objXmlIn, "REQUEST")
      
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
    
    ValidateCustomerDetails objxmlRequestNode
    
    If objContext Is Nothing Then
        Set objCustomerTxBO = New CustomerTxBO
    Else
        Set objCustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    objCustomerTxBO.CreateCustomerDetails objxmlRequestNode, objXmlResponseNode
    
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Append CustomerKey to the Response
    objXmlResponseElem.appendChild objXmlResponseNode.selectSingleNode(".//CUSTOMERKEY")
    
    CreateCustomerDetails = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCustomerDetailsExit:
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    Set objxmlRequestNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlResponseNode = Nothing
    Set objCustomerTxBO = Nothing
    Set objContext = Nothing

    Exit Function

CreateCustomerDetailsVbErr:

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    CreateCustomerDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume CreateCustomerDetailsExit

End Function

Public Function UpdateCustomerDetails(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Updates the customer details
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       UpdateCustomerDetails  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'
'------------------------------------------------------------------------------------------

On Error GoTo UpdateCustomerDetailsVbErr
    
    Const strFunctionName As String = "UpdateCustomerDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlRequestElem As IXMLDOMNode
    Dim objXmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objCustomerTxBO As ICustomerTxBO
    
    objXmlOut.async = False

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlRequestElem = objXmlAssist.GetMandatoryNode(objXmlIn, "REQUEST")

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    objXmlOut.appendChild objXmlResponseElem
        
    ' Validate the Personal Details
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerDetails(objXmlRequestElem)
    objErrAssist.CheckResponse xmlTempResponseNode, objXmlResponseElem, True
    
'     ValidateCustomerDetails objXmlRequestElem
    
    If objContext Is Nothing Then
        Set objCustomerTxBO = New CustomerTxBO
    Else
        Set objCustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
    
    objCustomerTxBO.UpdateCustomerDetails objXmlRequestElem, objXmlResponseElem
    
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

'    Dim strStatus As String
'
'    strStatus = "(test 10.3)" & _
'                ", xml parsing (input), xml response" & _
'                ", customer, customerversion only" & _
'                ", one ADO cmd object in CustomerDO.UpdateCustomerDetailsEx" & _
'                ", from 5 random customers"
'
'    objXmlResponseElem.setAttribute "STATUS", strStatus
    
    UpdateCustomerDetails = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateCustomerDetailsExit:
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlRequestElem = Nothing
    Set objXmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Set objCustomerTxBO = Nothing

    Exit Function

UpdateCustomerDetailsVbErr:

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    UpdateCustomerDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume UpdateCustomerDetailsExit

End Function

Public Function CreateCustomerLock(ByVal vstrXMLRequest As String, _
                                   Optional ByVal strLockRequired As String = "") As String

On Error GoTo CreateCustomerLockVbErr
    
    Const strFunctionName As String = "CreateCustomerLock"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    '  Create Response Element
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
    
    '-----------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String  'Local Variable, different from the PassedOne
    
    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem() Then
            strLockingReqd = 1
        Else
            objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
            CreateCustomerLock = objXmlOut.xml
            GoTo CreateCustomerLockExit
        End If
    ElseIf strLockRequired = "0" Then
        objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
        CreateCustomerLock = objXmlOut.xml
        GoTo CreateCustomerLockExit
    Else
        strLockingReqd = "1"
    End If

    
    objXmlOut.async = False

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' validate here?

    ICustomerBO_CreateCustomerLock objXmlIn.documentElement, objXmlResponseNode, strLockingReqd
    
     '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    CreateCustomerLock = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateCustomerLockExit:
   
    Set objXmlResponseElem = Nothing
    Set objXmlResponseNode = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

    Exit Function

CreateCustomerLockVbErr:

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    CreateCustomerLock = objErrAssist.CreateErrorResponse
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume CreateCustomerLockExit

End Function

Public Function DeleteCustomerLock(ByVal vstrXMLRequest As String, _
                                   Optional ByVal strLockRequired As String) As String

On Error GoTo DeleteCustomerLockVbErr
    
    Const strFunctionName As String = "DeleteCustomerLock"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    objXmlOut.async = False

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
    
    '-----------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String  ' This is different from the input parameter to this function
    
    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem() Then
            strLockingReqd = 1
        Else
            objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
            DeleteCustomerLock = objXmlOut.xml
            GoTo DeleteCustomerLockExit
        End If
    ElseIf strLockRequired = "0" Then
        objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
        DeleteCustomerLock = objXmlOut.xml
        GoTo DeleteCustomerLockExit
    Else
        strLockingReqd = "1"
    End If

    ' validate here?
    ICustomerBO_DeleteCustomerLock objXmlIn.documentElement, objXmlResponseNode, strLockingReqd
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    DeleteCustomerLock = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteCustomerLockExit:
   
    Set objXmlResponseElem = Nothing
    Set objXmlResponseNode = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing

    Exit Function

DeleteCustomerLockVbErr:

    If objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    DeleteCustomerLock = objErrAssist.CreateErrorResponse
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume DeleteCustomerLockExit

End Function

Public Function CreateCustomerDetailsAndEmpt( _
                        ByVal vstrXMLRequest As String) As String
' header -------------------------------------------------------------------------------------
' description:  Creates records for Personal details and Emplyoment for the customer passed in
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'<REQUEST ACTION=""CREATE"">
'   <CREATE>
'       <CUSTOMER>
'           <CUSTOMERVERSION>
'               <TITLE>1</TITLE>
'               <TITLEOTHER>Dr.</TITLEOTHER>
'               <SURNAME>Alba</SURNAME>
'               <FIRSTFORENAME>First Fore</FIRSTFORENAME>
'               <OTHERFORENAMES>Other Fore</OTHERFORENAMES>
'               <PREFERREDNAME>Preferred Name</PREFERREDNAME>
'               <DATEOFBIRTH>01-01-1970</DATEOFBIRTH>
'               <ALIASINDICATOR>0</ALIASINDICATOR>
'               <MARITALSTATUS>0</MARITALSTATUS>
'               <GENDER>1</GENDER>
'
'               <EMPLOYMENT>
'                   <EMPLOYMENTSTATUS>10</EMPLOYMENTSTATUS>
'                   <EMPLOYMENTTYPE>6</EMPLOYMENTTYPE>
'                   <OTHEREMPLOYMENTSTATUS></OTHEREMPLOYMENTSTATUS>
'               </EMPLOYMENT>
'
'               <EMPLOYEDDETAILS>
'                   <SHARESOWNEDINDICATOR></SHARESOWNEDINDICATOR>
'               </EMPLOYEDDETAILS>
'           </CUSTOMERVERSION>
'       </CUSTOMER>
'   </CREATE>
'</REQUEST>

'---------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateCustomerDetailsAndEmpt"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
'    Set xmlTempResponseNode = IOmiga4BO_Validate(xmlIn.documentElement, bomCreate)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateCustomerDetailsAndEmpt(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateCustomerDetailsAndEmpt = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateCustomerDetailsAndEmpt = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function UpdatePersonalDetailsAndEmpt( _
                        ByVal vstrXMLRequest As String) As String
' header -------------------------------------------------------------------------------------
' description:  Creates records for Personal details and Emplyoment for the customer passed in
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'<REQUEST ACTION=""CREATE"">
'   <UPDATE>
'       <CUSTOMER>
'           <CUSTOMERNUMER>..</CUSTOMERNUMER>
'           <CHANNELID>..</CHANNELID>
'           <CUSTOMERVERSION>
'               <CUSTOMERNUMER>..</CUSTOMERNUMER>
'               <CUSTOMERVERSIONNUMER>..</CUSTOMERVERSIONNUMER>
'               <TITLE>1</TITLE>
'               <TITLEOTHER>Dr.</TITLEOTHER>
'               <SURNAME>Alba</SURNAME>
'               <FIRSTFORENAME>First Fore</FIRSTFORENAME>
'               <OTHERFORENAMES>Other Fore</OTHERFORENAMES>
'               <PREFERREDNAME>Preferred Name</PREFERREDNAME>
'               <DATEOFBIRTH>01-01-1970</DATEOFBIRTH>
'               <ALIASINDICATOR>0</ALIASINDICATOR>
'               <MARITALSTATUS>0</MARITALSTATUS>
'               <GENDER>1</GENDER>
'
'               <DIRECTORYADDRESSINDICATOR>1</DIRECTORYADDRESSINDICATOR>
'
'               <EMPLOYMENT>
'                   <CUSTOMERNUMER>..</CUSTOMERNUMER>
'                   <CUSTOMERVERSIONNUMER>..</CUSTOMERVERSIONNUMER>
'                   <EMPLOYMENTSEQUENCENUMBER>..</EMPLOYMENTSEQUENCENUMBER>
'                   <EMPLOYMENTSTATUS>10</EMPLOYMENTSTATUS>
'                   <EMPLOYMENTTYPE>6</EMPLOYMENTTYPE>
'                   <OTHEREMPLOYMENTSTATUS></OTHEREMPLOYMENTSTATUS>
'               </EMPLOYMENT>
'
'               <EMPLOYEDDETAILS>
'                   <CUSTOMERNUMER>..</CUSTOMERNUMER>
'                   <CUSTOMERVERSIONNUMER>..</CUSTOMERVERSIONNUMER>
'                   <EMPLOYMENTSEQUENCENUMBER>..</EMPLOYMENTSEQUENCENUMBER>
'                   <SHARESOWNEDINDICATOR></SHARESOWNEDINDICATOR>
'               </EMPLOYEDDETAILS>
'           </CUSTOMERVERSION>
'       </CUSTOMER>
'   </UPDATE>
'</REQUEST>

'---------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "UpdatePersonalDetailsAndEmpt"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
'    Set xmlTempResponseNode = IOmiga4BO_Validate(xmlIn.documentElement, bomCreate)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdatePersonalDetailsAndEmpt(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdatePersonalDetailsAndEmpt = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdatePersonalDetailsAndEmpt = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Public Function CreateFullCustomerDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'Request format
' <REQUEST USERID=""USER0001"">
'   <CREATE>
'   <CUSTOMER>
'       <CUSTOMERNUMBER/>
'       <CUSTOMERVERSION>
'            <CUSTOMERNUMBER/>
'            <CUSTOMERVERSIONNUMBER/>
'            <CORRESPONDENCESALUTATION/>
'               .........
'               .........
'            <CUSTOMERADDRESSLIST>
'               <CUSTOMERADDRESS>
'                   <ADDRESSGUID/>
'                   <ADDRESSTYPE>2</ADDRESSTYPE>
'                   <ADDRESS>
'                       <ADDRESSGUID/>
'                       <BUILDINGORHOUSENUMBER>45</BUILDINGORHOUSENUMBER>
'                       <STREET>Malvern Road</STREET>"
'                       .........
'                   </ADDRESS>"
'                </CUSTOMERADDRESS>
'            </CUSTOMERADDRESSLIST>
'            <CUSTOMERTELEPHONENUMBERLIST>
'                <CUSTOMERTELEPHONENUMBER>
'                   <TELEPHONESEQUENCENUMBER/>
'                   ......
'                   ......
'                </CUSTOMERTELEPHONENUMBER>
'            </CUSTOMERTELEPHONENUMBERLIST>
'            <AREASOFINTERESTLIST>
'               <AREASOFINTEREST>"
'                   <INTERESTAREA>1</INTERESTAREA>
'               </AREASOFINTEREST>
'               <AREASOFINTEREST>
'                   <INTERESTAREA>2</INTERESTAREA>
'                </AREASOFINTEREST>
'            </AREASOFINTERESTLIST>
'            <ALIASPERSONLIST>
'               <ALIAS>
'                   <ALIASSEQUENCENUMBER></ALIASSEQUENCENUMBER>
'                    <ALIASTYPE>1</ALIASTYPE>
'                     ...........
'                     ................
'                    <PERSON>
'                        <PERSONGUID></PERSONGUID>
'                        <DATEOFBIRTH></DATEOFBIRTH>
'                          ...........
'                    </PERSON>
'               </ALIAS>
'            </ALIASPERSONLIST>
'            <EMPLOYMENTLIST>
'               <EMPLOYMENT>
'                   <EMPLOYMENTSEQUENCENUMBER></EMPLOYMENTSEQUENCENUMBER>
'                    ..............
'               </EMPLOYMENT>
'            </EMPLOYMENTLIST>
'            <UNEARNEDINCOMELIST>
'               <UNEARNEDINCOME>
'                   <UNEARNEDINCOMESEQUENCENUMBER></UNEARNEDINCOMESEQUENCENUMBER>
'                   ..............
'               </UNEARNEDINCOME>
'            </UNEARNEDINCOMELIST>
'            <LOANSLIABILITIESLIST>
'               <LOANSLIABILITIES>
'                   <SEQUENCENUMBER></SEQUENCENUMBER>
'                   <ADDITIONALDETAILS>Nothing</ADDITIONALDETAILS>
'                   ..........
'                   ..........
'               </LOANSLIABILITIES>
'            </LOANSLIABILITIESLIST>
'            <BANKCREDITCARDLIST>
'               <BANKCREDITCARD>
'                   <SEQUENCENUMBER></SEQUENCENUMBER>
'                   <ADDITIONALDETAILS>Nothing</ADDITIONALDETAILS>
'                   ................
'               </BANKCREDITCARD>
'            </BANKCREDITCARDLIST>
'         </CUSTOMERVERSION>"
'       </CUSTOMER>
'       </CREATE>
'    </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateFullCustomerDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIScriptInterface As IScriptInterface
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

   ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                        strFunctionName, xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateFullCustomerDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateFullCustomerDetails = xmlTempResponseNode.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objIScriptInterface = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

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

    CreateFullCustomerDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function

Public Function CreateAlias(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateAliasVbErr

    Const strFunctionName As String = "CreateAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIScriptInterface As IScriptInterface
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlGeneratedKeys As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateAlias(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

   ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlGeneratedKeys = xmlOut.createElement("ALIASSEQUENCENUMBERSLIST")
    
    Set xmlTempResponseNode = ICustomerBO_CreateAlias(xmlIn.documentElement, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    xmlResponseElem.appendChild xmlGeneratedKeys
    CreateAlias = xmlResponseElem.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateAliasExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objIScriptInterface = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlGeneratedKeys = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CreateAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateAliasExit

End Function

Private Function ICustomerBO_CreateAlias(ByVal vxmlRequest As IXMLDOMElement, _
                                         ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_CreateAliasVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateAlias"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
'   Call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateAlias(vxmlRequest, vxmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_CreateAlias = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_CreateAliasExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_CreateAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_CreateAliasExit

End Function

Public Function UpdateAlias(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               Alias and corresponding Person (if any)
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAliasVbErr

    Const strFunctionName As String = "UpdateAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateAlias(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
                                                            
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateAlias(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    UpdateAlias = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateAliasExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

UpdateAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '  go to clean-up section
    Resume UpdateAliasExit

End Function

Public Function DeleteAlias(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAliasVbErr

    Const strFunctionName As String = "DeleteAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
          
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_DeleteAlias(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    DeleteAlias = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteAliasExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlIn = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

DeleteAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume DeleteAliasExit

End Function

Public Function GetAlias(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetAliasVbErr

    Const strFunctionName As String = "GetAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetAlias(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetAlias = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetAliasExit:

    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    ' go to clean-up section
    Resume GetAliasExit

End Function

Private Function ICustomerBO_CreateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates a single instance of the persistant data associated with this
'               invokes appropriate transactioning (Tx) business object
' pass:         vXMLRequest  xml Request node containing data to be persistes
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_CreateCustomerAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateCustomerAddress"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateCustomerAddress(vxmlRequest, vxmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_CreateCustomerAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_CreateCustomerAddressExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_CreateCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_CreateCustomerAddressExit
End Function

Private Function ICustomerBO_CreateCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateCustomerTelephone a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_CreateCustomerTelephoneVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateCustomerTelephone"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateCustomerTelephone(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_CreateCustomerTelephone = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_CreateCustomerTelephoneExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_CreateCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_CreateCustomerTelephoneExit


End Function

Private Function ICustomerBO_CreateDependant(ByVal vxmlRequest As IXMLDOMElement, _
                                       ByVal vxmlGeneratedKeys As IXMLDOMElement) _
                                       As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateDependant a single instance of the persistant data associated with this
'               business object
' pass:         vXMLRequest  xml Request data stream containing data to be persisted
'               vxmlGeneratedKeys : node to attach the generated keys to
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_CreateDependantVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateDependant"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Create function
    Set xmlTempResponseNode = objICustomerTxBO.CreateDependant(vxmlRequest, vxmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_CreateDependant = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_CreateDependantExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_CreateDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_CreateDependantExit

End Function

Private Function ICustomerBO_CreateFullCustomerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateFullCustomerDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   call Business Transaction Object
    Set xmlTempResponseNode = objICustomerTxBO.CreateFullCustomerDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_CreateFullCustomerDetails = xmlTempResponseNode

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
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
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateFullCustomerDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function ICustomerBO_CreateCustomerDetailsAndEmpt( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateCustomerDetailsAndEmpt"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   call Business Transaction Object CreateTenancy function
    Set xmlTempResponseNode = objICustomerTxBO.CreateCustomerDetailsAndEmpt(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_CreateCustomerDetailsAndEmpt = xmlTempResponseNode

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
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
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_CreateCustomerDetailsAndEmpt = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function ICustomerBO_CreateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateTenancy a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateTenancyVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateTenancy"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   call Business Transaction Object CreateTenancy function
    Set xmlTempResponseNode = objICustomerTxBO.CreateTenancy(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_CreateTenancy = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateTenancyExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set ICustomerBO_CreateTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateTenancyExit
End Function

Private Function ICustomerBO_DeleteAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_DeleteAliasVbErr
    
    Const strFunctionName As String = "ICustomerBO_DeleteAlias"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Delete function
    Set xmlTempResponseNode = objICustomerTxBO.DeleteAlias(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_DeleteAlias = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_DeleteAliasExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICustomerTxBO = Nothing
    Exit Function

ICustomerBO_DeleteAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_DeleteAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_DeleteAliasExit
End Function

Private Function ICustomerBO_DeleteAllCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) _
                                                        As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:       xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_DeleteAllCustomerTelephoneVbErr
    
    Const strFunctionName As String = "ICustomerBO_DeleteAllCustomerTelephone"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '------------------------------------------------------------------------------------------
    ' validate parameters
    '------------------------------------------------------------------------------------------
    Dim strTagValue, strTagName As String
    strTagName = "CUSTOMERNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If

    strTagName = "CUSTOMERVERSIONNUMBER"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    If Len(strTagValue) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Expected " & strTagName & " tag"
    End If
        
    '   Call Business Transaction Object Delete function
    Set xmlTempResponseNode = objICustomerTxBO.DeleteAllCustomerTelephone(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_DeleteAllCustomerTelephone = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_DeleteAllCustomerTelephoneExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICustomerTxBO = Nothing
    
    Exit Function

ICustomerBO_DeleteAllCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_DeleteAllCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_DeleteAllCustomerTelephoneExit

End Function

Private Function ICustomerBO_DeleteCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  DeleteCustomerAddress a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vXMLRequest  xml Request node containing data to which identifies
'               instance to be Deleted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_DeleteCustomerAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_DeleteCustomerAddress"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Delete function
    Set xmlTempResponseNode = objICustomerTxBO.DeleteCustomerAddress(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_DeleteCustomerAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_DeleteCustomerAddressExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICustomerTxBO = Nothing
    Exit Function

ICustomerBO_DeleteCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_DeleteCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_DeleteCustomerAddressExit

End Function

Private Function ICustomerBO_DeleteDependant(ByVal vxmlRequest As IXMLDOMElement, _
                                            Optional ByVal rstrPersonGUID As Variant) _
                                                                As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteDependant a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vXMLRequest  xml Request node containing data to which identifies
'               instance to be Deleted
'               rstrPersonGUID - GUID of the person associated with the dependant
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_DeleteDependantVbErr
    
    Const strFunctionName As String = "ICustomerBO_DeleteDependant"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Delete function
    Set xmlTempResponseNode = objICustomerTxBO.DeleteDependant(vxmlRequest, rstrPersonGUID)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_DeleteDependant = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_DeleteDependantExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICustomerTxBO = Nothing
    Exit Function

ICustomerBO_DeleteDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_DeleteDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_DeleteDependantExit
End Function

Private Function ICustomerBO_DeleteTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteTenancy a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteTenancyVbErr
    
    Const strFunctionName As String = "ICustomerBO_DeleteTenancy"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
           
    '   call Business Transaction Object DeleteTenancy function
    Set xmlTempResponseNode = objICustomerTxBO.DeleteTenancy(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerBO_DeleteTenancy = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteTenancyExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    Set ICustomerBO_DeleteTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteTenancyExit
End Function

Private Function ICustomerBO_FindAliasList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds the data for multiple instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo ICustomerBO_FindAliasListVbErr
    
    Const strFunctionName As String = "ICustomerBO_FindAliasList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    ' Validate to check that Customer Number is mentioned (for all customers) in the Input String
    Set xmlNodeList = vxmlRequest.selectNodes("//" & cstrTABLE_ALIASNAME)
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_ALIASNAME & " tag not found"
    Else
        For Each xmlNode In xmlNodeList
            If objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER") = "" Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 256
            End If
        Next xmlNode
    End If
            
    Set xmlDataNode = objICustomerDO.FindAliasList(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerBO_FindAliasList = xmlResponseElem.cloneNode(True)

ICustomerBO_FindAliasListExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_FindAliasListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_FindAliasList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_FindAliasListExit

End Function

Private Function ICustomerBO_FindCustomerAddressList( _
            ByVal vxmlRequest As IXMLDOMElement, _
            Optional ByVal blnOrderBySequenceNUmber As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_FindCustomerAddressListVbErr
    
    Const strFunctionName As String = "ICustomerBO_FindCustomerAddressList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
     ' Validate to check that Customer Number is mentioned (for all customers) in the Input String
     
    If vxmlRequest.nodeName <> cstrTABLE_CUSTADDRESSNAME Then
        Set xmlNodeList = vxmlRequest.selectNodes("//" & cstrTABLE_CUSTADDRESSNAME)
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, cstrTABLE_CUSTADDRESSNAME & " tag not found"
        Else
            For Each xmlNode In xmlNodeList
                If objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER") = "" Then
                    objErrAssist.ThrowError TypeName(Me), strFunctionName, 256
                End If
            Next xmlNode
        End If
    End If
    
    Set xmlDataNode = objICustomerDO.FindCustomerAddressList(vxmlRequest, blnOrderBySequenceNUmber)
    'SYS0785 NOT an error for there to be NO Addresses for this customer.
    'MCS 25/05/00
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerBO_FindCustomerAddressList = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
ICustomerBO_FindCustomerAddressListExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_FindCustomerAddressListVbErr:
    
    'SYS0785 NOT an error for there to be NO Addresses for this customer.
    'MCS 25/05/00
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_FindCustomerAddressList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_FindCustomerAddressListExit

End Function

Private Function ICustomerBO_FindDependantList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds the data for multiple instances of the persistant data associated with
'               Dependant
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_FindDependantListVbErr
    
    Const strFunctionName As String = "ICustomerBO_FindDependantList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMElement
        
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    ' Check that 'DEPENDANT' tag is mentioned
    Set xmlRequestNode = vxmlRequest.getElementsByTagName("DEPENDANT").Item(0)

    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "DEPENDANT tag not found"
    End If
           
    Set xmlDataNode = objICustomerDO.FindDependantList(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerBO_FindDependantList = xmlResponseElem.cloneNode(True)

ICustomerBO_FindDependantListExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_FindDependantListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_FindDependantList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_FindDependantListExit

End Function

Private Function ICustomerBO_GetAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetAliasVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetAlias"


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
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_ALIASNAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_ALIASNAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_ALIASNAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetAlias(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerBO_GetAlias = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetAliasExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetAliasExit

End Function

Private Function ICustomerBO_GetCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vXMLRequest     xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetCustomerAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetCustomerAddress"


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
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTADDRESSNAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTADDRESSNAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTADDRESSNAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetCustomerAddress(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerBO_GetCustomerAddress = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetCustomerAddressExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetCustomerAddressExit

End Function

Private Function ICustomerBO_GetCustomerPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instances of the persistant data associated with
'               CustomerPersonalDetails
' pass:         vXMLRequest     xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetCustomerPersonalDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetCustomerPersonalDetails"


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
    Dim xmlElement As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    If vxmlRequest.nodeName = "CUSTOMER" Or vxmlRequest.nodeName = "CUSTOMERVERSION" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("CUSTOMER").Item(0)
        If xmlRequestNode Is Nothing Then
            Set xmlRequestNode = vxmlRequest.getElementsByTagName("CUSTOMERVERSION").Item(0)
        End If
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMER/CUSTOMERVERSION tag not found"
    End If
        
    ' Make sure that both CustomerNumber and CustomerVersionNumber are mentioned in Input
    
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    
    strCustomerNumber = objXmlAssist.GetTagValue(xmlRequestNode, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(xmlRequestNode, "CUSTOMERVERSIONNUMBER")
    
    If strCustomerNumber = "" Or strCustomerVersionNumber = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If
    
    '
    ' Get personal details
    ' If the Input's root node is mentioned as 'CUSTOMERVERSION', create a node with name
    ' 'CUSTOMER' containing values for CustomerNumber and CustomerVersionNumber
    
    If xmlRequestNode.nodeName = "CUSTOMERVERSION" Then
        Set xmlElement = vxmlRequest.ownerDocument.createElement("CUSTOMER")
        xmlElement.appendChild xmlRequestNode.selectSingleNode("CUSTOMERNUMBER")
        xmlElement.appendChild xmlRequestNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
        
        xmlRequestNode.appendChild xmlElement
    End If
    
    Set xmlDataNode = objICustomerDO.GetPersonalDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
    
    ' Get telephone list
    ' Create a node with name 'CUSTOMRETELEPHONENUMBER' contatining CustomerNumber
    ' and CustomerVersionNumber in it.
    '
    Set xmlElement = vxmlRequest.ownerDocument.createElement("CUSTOMERTELEPHONENUMBER")
    xmlElement.appendChild xmlRequestNode.selectSingleNode("//CUSTOMERNUMBER")
    xmlElement.appendChild xmlRequestNode.selectSingleNode("//CUSTOMERVERSIONNUMBER")
        
    xmlRequestNode.appendChild xmlElement
    
    On Error GoTo FindCustomerTelephoneListErr
    
    Set xmlDataNode = Nothing
    Set xmlDataNode = objICustomerDO.FindCustomerTelephoneList(xmlRequestNode)
    
    On Error GoTo ICustomerBO_GetCustomerPersonalDetailsVbErr
    
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.firstChild.appendChild xmlDataNode
    End If
           
    Set ICustomerBO_GetCustomerPersonalDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetCustomerPersonalDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindCustomerTelephoneListErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ' It is possible that no telephone number records exist for the customer
        Resume Next
    End If

ICustomerBO_GetCustomerPersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetCustomerPersonalDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetCustomerPersonalDetailsExit
End Function

Private Function ICustomerBO_GetComponentData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data from specified columns & tables associated with Customer.
' pass:         vXMLRequest     xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved and criteria
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetComponentDataVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetComponentData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    Set xmlDataNode = objICustomerDO.GetComponentData(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
           
    Set ICustomerBO_GetComponentData = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetComponentDataExit:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetComponentDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetComponentData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetComponentDataExit
End Function

Private Function ICustomerBO_GetDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instances of the persistant data associated with
'               Dependant
' pass:         vXMLRequest     xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetDependantVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetDependant"


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
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_DEPENDANTNAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_DEPENDANTNAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_DEPENDANTNAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetDependant(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerBO_GetDependant = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetDependantExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetDependantExit

End Function

Private Function ICustomerBO_GetPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPersonalDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetPersonalDetails"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetPersonalDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ICustomerBO_GetPersonalDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetPersonalDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetPersonalDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetPersonalDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPersonalDetailsExit
End Function

Private Function ICustomerBO_GetPersonalDetailsAndAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets customer address and personal details from CustomerBO.FindCustomerAddressList
'               and CustomerBO.GetPersonalDetails
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPersonalDetailsAndAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetPersonalDetailsAndAddress"

    Const cstrROOT_NODE_NAME = "CUSTOMER"
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerBO As ICustomerBO
    Dim objXmlAssist As New XMLAssist
      
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerBO = New CustomerBO
    Else
        Set objICustomerBO = m_objContext.CreateInstance(App.Title & ".CustomerBO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerBO.GetPersonalDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
    
    objXmlAssist.ChangeNodeName xmlRequestNode, "CUSTOMER", "CUSTOMERADDRESS"
    
    Set xmlDataNode = objICustomerBO.FindCustomerAddressList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
            
    Set ICustomerBO_GetPersonalDetailsAndAddress = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetPersonalDetailsAndAddressExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerBO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetPersonalDetailsAndAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetPersonalDetailsAndAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPersonalDetailsAndAddressExit

End Function

Private Function ICustomerBO_GetTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instances of the persistant data associated with
'               Tenancy
' pass:         vXMLRequest     xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetTenancyVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetTenancy"


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
    
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_TENANCYNAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_TENANCYNAME).Item(0)
    End If
              
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_TENANCYNAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.GetTenancy(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set ICustomerBO_GetTenancy = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ICustomerBO_GetTenancyExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetTenancyExit

End Function

Private Function ICustomerBO_SaveAlias(ByVal vxmlRequest As IXMLDOMElement, _
                                       ByVal xmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request nodecontaining data to be saved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveAliasVbErr

    Const strFunctionName As String = "ICustomerBO_SaveAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlAliasElem As IXMLDOMElement
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If

    Set xmlAliasElem = vxmlRequest.getElementsByTagName("ALIAS").Item(0)

    If xmlAliasElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If

    Set xmlTempResponseNode = objICustomerTxBO.SaveAlias(vxmlRequest, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                   
    Set ICustomerBO_SaveAlias = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerBO_SaveAliasExit:

    '   clean-up section, release object references
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlAliasElem = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

ICustomerBO_SaveAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ICustomerBO_SaveAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ICustomerBO_SaveAliasExit
End Function

Private Function ICustomerBO_SaveCustomerAddress(ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request nodecontaining data to be saved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerBO_SaveCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCustomerAddressElem As IXMLDOMElement
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlCustomerAddressElem = vxmlRequest.getElementsByTagName("CUSTOMERADDRESS").Item(0)

    If xmlCustomerAddressElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If

    'Get the values of AddressSequence Number and decide the operation to be performed
    ' on CustomerAddress
    Dim strCustAddressSeqNumber As String

    strCustAddressSeqNumber = objXmlAssist.GetTagValue(xmlCustomerAddressElem, _
                                    "CUSTOMERADDRESSSEQUENCENUMBER")

    'If CustomerAddressSequenceNumber is null create customer address else update
    If strCustAddressSeqNumber = "" Then
        Set xmlTempResponseNode = ICustomerBO_CreateCustomerAddress(vxmlRequest, vxmlGeneratedKeys)
    Else
         Set xmlTempResponseNode = ICustomerBO_UpdateCustomerAddress(vxmlRequest)
    End If
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                   
    Set ICustomerBO_SaveCustomerAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerBO_SaveCustomerAddressExit:

    '   clean-up section, release object references
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCustomerAddressElem = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

ICustomerBO_SaveCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_SaveCustomerAddress = objErrAssist.CreateErrorResponseEx

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ICustomerBO_SaveCustomerAddressExit
End Function

Private Function ICustomerBO_SavePersonalDetailsAndAddress(ByVal vxmlRequest As IXMLDOMElement, _
                    ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  BG:01/06/00     Updates Customer details including email and address
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
    On Error GoTo SavePersonalDetailsAndAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_SavePersonalDetailsAndAddress"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    Set xmlTempResponseNode = objICustomerTxBO.SavePersonalDetailsAndAddress(vxmlRequest, vxmlGeneratedKeys)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerBO_SavePersonalDetailsAndAddress = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SavePersonalDetailsAndAddressExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set ICustomerBO_SavePersonalDetailsAndAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SavePersonalDetailsAndAddressExit

End Function

Private Function ICustomerBO_SaveCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveCustomerTelephone a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request nodecontaining data to be saved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerBO_SaveCustomerTelephone"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'   Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'   xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call corresponding function in Business Transaction Object
    Set xmlTempResponseNode = objICustomerTxBO.SaveCustomerTelephone(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_SaveCustomerTelephone = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerBO_SaveCustomerTelephoneExit:

    '   clean-up section, release object references
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing

    Exit Function

ICustomerBO_SaveCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_SaveCustomerTelephone = objErrAssist.CreateErrorResponseEx

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ICustomerBO_SaveCustomerTelephoneExit
End Function

Private Function ICustomerBO_SaveWrapUpDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveWrapUpDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request nodecontaining data to be saved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveWrapUpDetailsVbErr

    Const strFunctionName As String = "ICustomerBO_SaveWrapUpDetails"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call corresponding function in Business Transaction Object
    Set xmlTempResponseNode = objICustomerTxBO.SaveWrapUpDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_SaveWrapUpDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

ICustomerBO_SaveWrapUpDetailsExit:

    '   clean-up section, release object references
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing

    Exit Function

ICustomerBO_SaveWrapUpDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_SaveWrapUpDetails = objErrAssist.CreateErrorResponseEx

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume ICustomerBO_SaveWrapUpDetailsExit
End Function

Private Function ICustomerBO_SaveTaxDetails(ByVal vxmlRequest As IXMLDOMElement, Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveTaxDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveTaxDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_SaveTaxDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerEmploymentTxBO As ICustomerEmploymentTxBO
    Dim objICustomerTxBO As ICustomerTxBO
    Dim objIScriptInterface As IScriptInterface
    Dim strUpdateCustomerVersion As String       ' MAR135

    Dim blnUpdate As Boolean
    blnUpdate = False

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerEmploymentTxBO = New CustomerEmploymentTxBO
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerEmploymentTxBO = objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentTxBO")
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    ' Determine if Creating or Updating
    Dim strAction As String
    ' AS 27/11/00 CORE000019 : ACTION attribute removed OPERATION attribute added
    strAction = vxmlRequest.getAttribute("OPERATION")
    If strAction = "CREATE" Then
        '   call Business Transaction Object SaveTaxDetails function
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.CreateTaxDetails(vxmlRequest)
    Else
        Set xmlTempResponseNode = objICustomerEmploymentTxBO.UpdateTaxDetails(vxmlRequest)
    End If
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    'Update tax details held in CustomerVersion
    'MAR135 Only update the Customer Version if the indicator is not present.
    strUpdateCustomerVersion = xmlGetAttributeText(vxmlRequest, "UPDATECUSTOMERVERSION")
    If (strUpdateCustomerVersion <> "False") Then
        Set xmlTempResponseNode = objICustomerTxBO.UpdateCustomerVersion(vxmlRequest, True)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
   
    Set ICustomerBO_SaveTaxDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveTaxDetailsExit:

    Set objErrAssist = Nothing
    Set objICustomerEmploymentTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function

SaveTaxDetailsVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        blnUpdate = True
        Resume Next
    End If

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_SaveTaxDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveTaxDetailsExit

End Function

Private Function ICustomerBO_UpdateAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_UpdateVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateAlias"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdateAlias(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdateAlias = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateAliasExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateAliasExit

End Function

Private Function ICustomerBO_UpdateContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactHistoryVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateContactHistory"

    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
    
    Set xmlTempResponseNode = objICustomerTxBO.UpdateContactHistory(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerBO_UpdateContactHistory = xmlResponseElem

    m_objContext.SetComplete
    
UpdateContactHistoryExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
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

    Set ICustomerBO_UpdateContactHistory = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateContactHistoryExit

End Function

Private Function ICustomerBO_UpdateCustomer( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ICustomerBO_UpdateCustomerVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCustomer"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
    
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdateCustomer(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdateCustomer = xmlResponseElem

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateCustomerExit:

    Set objErrAssist = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateCustomerVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCustomer = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateCustomerExit

End Function

Private Function ICustomerBO_UpdateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_UpdateCustomerAddressVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCustomerAddress"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdateCustomerAddress(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdateCustomerAddress = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateCustomerAddressExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateCustomerAddressExit

End Function

Private Function ICustomerBO_UpdatePersonalDetailsAndEmpt(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdatePersonalDetailsAndEmpt"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   call Business Transaction Object CreateTenancy function
    Set xmlTempResponseNode = objICustomerTxBO.UpdatePersonalDetailsAndEmpt(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set ICustomerBO_UpdatePersonalDetailsAndEmpt = xmlTempResponseNode

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
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
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdatePersonalDetailsAndEmpt = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function

Private Function ICustomerBO_UpdateCustomerPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_UpdateCustomerPersonalDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCustomerPersonalDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    ' Validate the Personal Details
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Update Personal Details
    Set xmlTempResponseNode = ICustomerBO_UpdatePersonalDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Save Customer Telephone
    Set xmlTempResponseNode = ICustomerBO_SaveCustomerTelephone(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
    Set ICustomerBO_UpdateCustomerPersonalDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateCustomerPersonalDetailsExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateCustomerPersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCustomerPersonalDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateCustomerPersonalDetailsExit

End Function

Private Function ICustomerBO_UpdateCustomerVersion( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ICustomerBO_UpdateCustomerVersionVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCustomerVersion"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdateCustomerVersion(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdateCustomerVersion = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateCustomerVersionExit:

    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateCustomerVersionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCustomerVersion = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateCustomerVersionExit

End Function

Private Function ICustomerBO_UpdateDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_UpdateDependantVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateDependant"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdateDependant(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdateDependant = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdateDependantExit:

    Set objErrAssist = Nothing
'    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdateDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateDependant = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdateDependantExit
End Function

Private Function ICustomerBO_UpdatePersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request node containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_UpdatePersonalDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdatePersonalDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
        
    '   Call Business Transaction Object Update function
    Set xmlTempResponseNode = objICustomerTxBO.UpdatePersonalDetails(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set ICustomerBO_UpdatePersonalDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
ICustomerBO_UpdatePersonalDetailsExit:

    Set objErrAssist = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ICustomerBO_UpdatePersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdatePersonalDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_UpdatePersonalDetailsExit

End Function

Private Function ICustomerBO_UpdatePropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes ICustomerTxBO_UpdatePropertyInsuranceDetails
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePropertyInsuranceDetailsVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdatePropertyInsuranceDetails"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerTxBO.UpdatePropertyInsuranceDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerBO_UpdatePropertyInsuranceDetails = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdatePropertyInsuranceDetailsExit:

    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set ICustomerBO_UpdatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatePropertyInsuranceDetailsExit

End Function

Private Function ICustomerBO_UpdateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateTenancy a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTenancyVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateTenancy"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICustomerTxBO As ICustomerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
    
    Set xmlTempResponseNode = objICustomerTxBO.UpdateTenancy(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ICustomerBO_UpdateTenancy = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateTenancyExit:

    Set objErrAssist = Nothing
    Set objICustomerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set ICustomerBO_UpdateTenancy = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateTenancyExit
End Function

Private Function ICustomerBO_ValidateAlias(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for Alias Data
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vxmlTableElement  xml Request node containing details of action to be
'                                 performed and data required
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateAliasVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

'    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPersonElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIPersonClassDef As New IPersonClassDef

    Set xmlResponseElem = vxmlTableElement.ownerDocument.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'If Alias Type is not mentioned Raise error
    If objXmlAssist.GetTagValue(vxmlTableElement, "ALIASTYPE") = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 265
    End If

    'If all the non-key values in Person are Null raise error 264
    Set xmlPersonElem = vxmlTableElement.getElementsByTagName("PERSON").Item(0)
    
    Set objIPersonClassDef = New PersonClassdef
    If Not xmlPersonElem Is Nothing Then
        If objXmlAssist.AreAllNonKeyValuesNullEx(xmlPersonElem, objIPersonClassDef.LoadPersonData) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 264
        End If
    End If
    
    Set ICustomerBO_ValidateAlias = xmlResponseElem.cloneNode(True)
    
ICustomerBO_ValidateAliasExit:
    'clean-up code
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlPersonElem = Nothing
    Set xmlResponseElem = Nothing

Exit Function

ICustomerBO_ValidateAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateAlias = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_ValidateAliasExit
    
End Function

Public Function FindAliasList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindAliasListVbErr

    Const strFunctionName As String = "FindAliasList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_FindAliasList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    FindAliasList = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

FindAliasListExit:

    Set xmlResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

FindAliasListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindAliasList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        ' SR 12/09/00 : Complete the transaction, if error raises is RecordNotFound
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
           objContext.SetComplete
        Else
            objContext.SetComplete
        End If
    End If

    '   go to clean-up section
    Resume FindAliasListExit

End Function

Public Function SaveAlias(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the Alias Data after finding the operation to be performed on it
' pass:         vstrXMLRequest  xml Request data to be passed to the methos
' Return:       nil
' Error:        Invalid Paramente
'------------------------------------------------------------------------------------------

On Error GoTo SaveAliasVbErr

    Const strFunctionName As String = "SaveAlias"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlGeneratedKeys As IXMLDOMElement

'    If AliasXML is NULL raise error
    If vstrXMLRequest = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, "Invalid Parameter "
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlGeneratedKeys = xmlOut.createElement("ALIASLIST")
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveAlias(xmlIn.documentElement, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    If Not xmlGeneratedKeys.firstChild Is Nothing Then
        xmlResponseElem.appendChild xmlGeneratedKeys.firstChild
    End If
    
    SaveAlias = xmlOut.xml
    
SaveAliasExit:

'       clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    
    Exit Function
SaveAliasVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveAlias = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

'     go to clean-up section
    Resume SaveAliasExit

End Function

Public Function FindCustomerAddressList(ByVal vstrXMLRequest As String, _
                        Optional ByVal blnOrderBySequenceNUmber As Boolean = False) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'               blnOrderBySequenceNUmber - Whether the list is to be sorted in the Ascending
'                                           order of AddressSequenceNumber
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomerAddressListVbErr

    Const strFunctionName As String = "FindCustomerAddressList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_FindCustomerAddressList(xmlIn.documentElement, blnOrderBySequenceNUmber)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    FindCustomerAddressList = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

FindCustomerAddressListExit:

    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

FindCustomerAddressListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCustomerAddressList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume FindCustomerAddressListExit

End Function

Public Function GetComponentData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data from specified columns of the view associated with Customer.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'RequestFormat :
'              <REQUEST ACTION="SEARCH>
'                   <TABLE>CUSTOMERVERSION
'                       <CRITERIA>
'                           <CUSTOMERNUMBER>1234<CUSTOMERNUMBER>
'                           <CUSTOMERVERSIONNUMBER>1<CUSTOMERVERSIONNUMBER>
'                       </CRITERIA>
'                       <OUTPUT ALLFIELDS = "FALSE">
'                           <FIELDS>
'                               <FIELD>SURNAME</FIELD>
'                               <FIELD>TITLE</FIELD>
'                           </FIELDS>
'                       <OUTPUT>
'               </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo GetGetComponentDataVbErr

    Const strFunctionName As String = "GetComponentData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetComponentData(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetComponentData = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetGetComponentDataExit:
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetGetComponentDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetComponentData = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume GetGetComponentDataExit

End Function

Public Function GetCustomerAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerAddressVbErr

    Const strFunctionName As String = "GetCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlOut.xml
    GetCustomerAddress = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetCustomerAddressExit:
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCustomerAddress = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume GetCustomerAddressExit

End Function

Private Function ICustomerBO_ValidateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) _
                                            As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for Customer Address and Address Data
'               if hard-coded rules completed OK return
' pass:         vXMLRequest  xml Request node containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateCustomerAddressVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvertAssist As New ConvertAssist
    Dim objSQLAssist As New SQLAssist
    Dim objIADOAssist As ADOAssist
    
    Dim xmlResponseElem As IXMLDOMElement

    Dim strDateMovedIn As String
    Dim strDateMovedOut As String

    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '  Validate Dates - Date Moved Out, if mentioned, must not be prior to Date Moved In
    strDateMovedOut = objXmlAssist.GetTagValue(vxmlRequest, "DATEMOVEDOUT")
    If strDateMovedOut <> "" Then
        strDateMovedIn = objXmlAssist.GetTagValue(vxmlRequest, "DATEMOVEDIN")
        If strDateMovedIn = "" Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 507, _
                            " Date Moved In is missing "
        End If

        If objConvertAssist.CSafeDate(strDateMovedIn) > _
                        objConvertAssist.CSafeDate(strDateMovedOut) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 135
        End If
    End If

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    'For Addres Type 1(Home) Or 2 (Correspondence), if a record already exists in
    'CustomerAddress (for this Customer Number and Customer Version Number)
    'with differnet sequence number - Raise error (only one Home / correspondence
    'address is allowed)
    Dim intAddressType As Integer
    Dim strCustomerNumber As String, _
        strCustomerVersionNumber As String, _
        strEmploymentSequenceNumber As String
    Dim vntTemp As Variant
    Dim strCondition As String

    intAddressType = Val(objXmlAssist.GetTagValue(vxmlRequest, "ADDRESSTYPE"))
    If (intAddressType = 1) Or (intAddressType = 2) Then
        strCustomerNumber = objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERNUMBER")
        strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERVERSIONNUMBER")
        strEmploymentSequenceNumber = objXmlAssist.GetTagValue(vxmlRequest, "CUSTOMERADDRESSSEQUENCENUMBER")

        If strCustomerNumber = "" Or strCustomerVersionNumber = "" Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 519
        End If

        strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                       & " And CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                       & " And ADDRESSTYPE = " & intAddressType

        vntTemp = objIADOAssist.GetValueFromTable("CUSTOMERADDRESS", strCondition, "CUSTOMERADDRESSSEQUENCENUMBER")
        If Not IsNull(vntTemp) Then
            If CStr(vntTemp) <> strEmploymentSequenceNumber Then
                ' A record already exists for this address type
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 134
            End If
        End If
    End If
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    '-----------------------------------------------------------------------------
    
    Set ICustomerBO_ValidateCustomerAddress = xmlResponseElem.cloneNode(True)
    
ICustomerBO_ValidateCustomerAddressExit:
    'clean-up code
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objConvertAssist = Nothing
    Set xmlResponseElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

Exit Function

ICustomerBO_ValidateCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateCustomerAddress = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_ValidateCustomerAddressExit

End Function

Public Function DeleteCustomerAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerAddressVbErr

    Const strFunctionName As String = "DeleteCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

     ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_DeleteCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DeleteCustomerAddress = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

DeleteCustomerAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

DeleteCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteCustomerAddress = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If

    '   go to clean-up section
    Resume DeleteCustomerAddressExit

End Function

Public Function UpdateCustomer(ByVal vstrXMLRequest As String) As String

On Error GoTo UpdateCustomerVbErr

    Const strFunctionName As String = "UpdateCustomer"


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

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCustomer(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
     
    UpdateCustomer = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
UpdateCustomerExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

UpdateCustomerVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCustomer = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '  go to clean-up section
    Resume UpdateCustomerExit
    
End Function

Public Function UpdateCustomerAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               CustomerAddress and corresponding records in tables 'Address', 'Tenancy'
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerAddressVbErr

    Const strFunctionName As String = "UpdateCustomerAddress"


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
    Dim objIScriptInterface As IScriptInterface

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
                                                            
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
     
    UpdateCustomerAddress = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateCustomerAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

UpdateCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCustomerAddress = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '  go to clean-up section
    Resume UpdateCustomerAddressExit

End Function

Public Function UpdateCustomerVersion(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with
'               CustomerVersion
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerVersionVbErr

    Const strFunctionName As String = "UpdateCustomerVersion"


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
    Dim objIScriptInterface As IScriptInterface

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerVersion(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
                                                            
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCustomerVersion(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
     
    UpdateCustomerVersion = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateCustomerVersionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

UpdateCustomerVersionVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCustomerVersion = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '  go to clean-up section
    Resume UpdateCustomerVersionExit

End Function

Public Function GetTenancy(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               the table 'Tenancy' and the respective data from the table 'ThirdParty'
' pass:         vXMLRequest  xml Request node which identifies instance of the persistant
'                            data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetTenancyVbErr

    Const strFunctionName As String = "GetTenancy"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetTenancy = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

GetTenancyExit:
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    ' go to clean-up section
    Resume GetTenancyExit

End Function

Private Function ICustomerBO_ValidateCustomerVersion( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ValidateCustomerVersionVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateCustomerVersion"


    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    '-----------------------------------------------------------------------------

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Set ICustomerBO_ValidateCustomerVersion = xmlResponseElem

ValidateCustomerVersionExit:
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateCustomerVersionVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateCustomerVersion = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ValidateCustomerVersionExit:

End Function

Private Function ICustomerBO_ValidateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  performs validation for Tenancy Data
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vXMLRequest    xml Request mode containing details of data to be validates
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateTenancyVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateTenancy"


    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'If Tenancy Type is not mentioned Raise error
    If objXmlAssist.GetTagValue(vxmlRequest, "TENANCYTYPE") = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 266
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    '-----------------------------------------------------------------------------

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Set ICustomerBO_ValidateTenancy = xmlResponseElem

ICustomerBO_ValidateTenancyExit:
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ICustomerBO_ValidateTenancyVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateTenancy = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ICustomerBO_ValidateTenancyExit:
End Function

Public Function DeleteTenancy(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteTenancy a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteTenancyVbErr

    Const strFunctionName As String = "DeleteTenancy"

    
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
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_DeleteTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteTenancy = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteTenancyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    DeleteTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteTenancyExit

End Function

Public Function CreateCustomerAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a single instance of the persistant data associated with this
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerAddressVbErr

    Const strFunctionName As String = "CreateCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIScriptInterface As IScriptInterface
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlGeneratedKeys As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

   ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlGeneratedKeys = xmlOut.createElement("CUSTOMERADDRESSLIST")
    Set xmlTempResponseNode = ICustomerBO_CreateCustomerAddress(xmlIn.documentElement, _
                                        xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    xmlResponseElem.appendChild xmlGeneratedKeys
    
    CreateCustomerAddress = xmlResponseElem.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

CreateCustomerAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    Set xmlGeneratedKeys = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CreateCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateCustomerAddress = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If

    '  go to clean-up section
    Resume CreateCustomerAddressExit

End Function

Public Function SaveCustomerAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the CustomerAddress Data after finding the operation to be performed on it
' pass:         vstrXMLRequest  xml Request data to be passed to the methos
' return:       nil
' Error :      Invalid Paramente
'------------------------------------------------------------------------------------------

On Error GoTo SaveCustomerAddressVbErr

    Const strFunctionName As String = "SaveCustomerAddress"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlGeneratedKeys As IXMLDOMElement
    
    Dim objIScriptInterface As IScriptInterface
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    'If Request String is NULL raise error
    If vstrXMLRequest = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'If CustomerAddressXML is NULL raise error
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerAddress(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlGeneratedKeys = xmlOut.createElement("CUSTOMERADDRESSLIST")
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveCustomerAddress(xmlIn.documentElement, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    xmlResponseElem.appendChild xmlGeneratedKeys
    
    SaveCustomerAddress = xmlOut.xml

SaveCustomerAddressExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlIn = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function
SaveCustomerAddressVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveCustomerAddress = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' go to clean-up section
    Resume SaveCustomerAddressExit

End Function

Public Function CreateDependant(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateDependant a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateDependantVbErr

    Const strFunctionName As String = "CreateDependant"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlGeneratedKeys As IXMLDOMElement
    
    Dim objIScriptInterface As IScriptInterface
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateDependant(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

   ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlGeneratedKeys = xmlOut.createElement("DEPENDANTSLIST")
    
    Set xmlTempResponseNode = ICustomerBO_CreateDependant(xmlIn.documentElement, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
      
    xmlResponseElem.appendChild xmlGeneratedKeys
    
    CreateDependant = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateDependantExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlGeneratedKeys = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    Exit Function

CreateDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateDependant = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If

    ' go to clean-up section
    Resume CreateDependantExit

End Function

Public Function CreateTenancy(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateTenancy a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo CreateTenancyVbErr

    Const strFunctionName As String = "CreateTenancy"

    
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
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//TENANCY/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateTenancy = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateTenancyExit:

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

CreateTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateTenancyExit

End Function

Public Function CreateCustomerTelephone(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateCustomerTelephone a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerTelephoneVbErr

    Const strFunctionName As String = "CreateCustomerTelephone"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIScriptInterface As IScriptInterface
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerTelephone(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateCustomerTelephone(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateCustomerTelephone = xmlResponseElem.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCustomerTelephoneExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objIScriptInterface = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CreateCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateCustomerTelephone = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If

    ' go to clean-up section
    Resume CreateCustomerTelephoneExit

End Function

Public Function UpdateDependant(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateDependant a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data UpdateDependant
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDependantVbErr

    Const strFunctionName As String = "UpdateDependant"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateDependant(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
                                                            
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateDependant(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateDependant = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateDependantExit:

     '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function
    
UpdateDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateDependant = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume UpdateDependantExit

End Function

Public Function SaveCustomerTelephone(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveCustomerTelephone a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request data stream containing data SaveCustomerTelephone
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveCustomerTelephoneVbErr

    Const strFunctionName As String = "SaveCustomerTelephone"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIScriptInterface As IScriptInterface
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
            
     'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerTelephone(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveCustomerTelephone(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveCustomerTelephone = xmlResponseElem.xml
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveCustomerTelephoneExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

SaveCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveCustomerTelephone = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume SaveCustomerTelephoneExit

End Function

Public Function SaveWrapUpDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveWrapUpDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request data stream containing data SaveWrapUpDetails
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveWrapUpDetailsVbErr

    Const strFunctionName As String = "SaveWrapUpDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIScriptInterface As IScriptInterface
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
            
     'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ValidateWrapUpDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveWrapUpDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveWrapUpDetails = xmlResponseElem.xml
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveWrapUpDetailsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

SaveWrapUpDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveWrapUpDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume SaveWrapUpDetailsExit

End Function
Public Function DeleteDependant(ByVal vstrXMLRequest As String, _
                           Optional ByRef rstrPersonGUID As Variant) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteDependant a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be DeleteDependantd
'               rstrPersonGUID - GUID of the person associated with the dependant
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDependantVbErr

    Const strFunctionName As String = "DeleteDependant"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_DeleteDependant(xmlIn.documentElement, rstrPersonGUID)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteDependant = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteDependantExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

DeleteDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteDependant = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume DeleteDependantExit

End Function

Public Function GetDependant(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDependantVbErr

    Const strFunctionName As String = "GetDependant"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetDependant(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetDependant = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDependantExit:

    Set xmlResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetDependantVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetDependant = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume GetDependantExit

End Function

Public Function GetCustomerPersonalDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerPersonalDetailsVbErr

    Const strFunctionName As String = "GetCustomerPersonalDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetCustomerPersonalDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
'    objXmlAssist.WriteXMLToFile xmlOut.xml
        
    GetCustomerPersonalDetails = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

GetCustomerPersonalDetailsExit:

    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetCustomerPersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCustomerPersonalDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    '   go to clean-up section
    Resume GetCustomerPersonalDetailsExit

End Function

Private Function ICustomerBO_ValidateCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK
' pass:         vXMLRequest   xml Request node containing details of action to be
'                             performed and data required
'                             This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateCustomerTelephoneVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateCustomerTelephone"


    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode     As IXMLDOMNode

    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If vxmlRequest.nodeName = "CUSTOMERTELEPHONENUMBER" Then
        Set xmlNode = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("CUSTOMERTELEPHONENUMBER")
    End If

    If Not xmlNode Is Nothing Then
        ValidateTelephone xmlNode
    Else
        For Each xmlNode In xmlNodeList
            ValidateTelephone xmlNode
        Next xmlNode
    End If
   
    Set ICustomerBO_ValidateCustomerTelephone = xmlResponseElem.cloneNode(True)
    
ICustomerBO_ValidateCustomerTelephoneExit:
    Set objErrAssist = Nothing
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlResponseElem = Nothing

Exit Function

ICustomerBO_ValidateCustomerTelephoneVbErr:
   If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateCustomerTelephone = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ICustomerBO_ValidateCustomerTelephoneExit

End Function

Private Function ValidateWrapUpDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK
' pass:         vXMLRequest   xml Request node containing details of action to be
'                             performed and data required
'                             This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateWrapUpDetailsVbErr

    Const strFunctionName As String = "ValidateWrapUpDetails"


    Dim objErrAssist As New ErrAssist
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode     As IXMLDOMNode

    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If vxmlRequest.nodeName = "WRAPUP" Then
        Set xmlNode = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("WRAPUP")
    End If

    If Not xmlNode Is Nothing Then
        ValidateTelephone xmlNode
    Else
        For Each xmlNode In xmlNodeList
            ValidateTelephone xmlNode
        Next xmlNode
    End If
   
    Set ValidateWrapUpDetails = xmlResponseElem.cloneNode(True)
    
ValidateWrapUpDetailsExit:
    Set objErrAssist = Nothing
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlResponseElem = Nothing

Exit Function

ValidateWrapUpDetailsVbErr:
   If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ValidateWrapUpDetails = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ValidateWrapUpDetailsExit

End Function

Private Function ICustomerBO_ValidateDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateDependantVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateDependant"


    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode     As IXMLDOMNode, _
        xmlPersonNode As IXMLDOMNode
        
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIPersonClassDef As IPersonClassDef
    Dim xmlPersonDef As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlNodeList = vxmlRequest.getElementsByTagName("DEPENDANT")
    
    Set objIPersonClassDef = New PersonClassdef
    Set xmlPersonDef = objIPersonClassDef.LoadPersonData
    
    For Each xmlNode In xmlNodeList
        '
        ' At least one non-key value for a person must be specified
        '
        Set xmlPersonNode = xmlNode.selectSingleNode("PERSON")
        If Not xmlPersonNode Is Nothing Then
            If objXmlAssist.AreAllNonKeyValuesNullEx(xmlPersonNode, xmlPersonDef) Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 264
            End If
        End If
    Next xmlNode

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    '-----------------------------------------------------------------------------
    
    Set ICustomerBO_ValidateDependant = xmlResponseElem.cloneNode(True)
    
ICustomerBO_ValidateDependantExit:

    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlPersonNode = Nothing
    Set xmlPersonDef = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlResponseElem = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

Exit Function

ICustomerBO_ValidateDependantVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateDependant = objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume ICustomerBO_ValidateDependantExit

End Function

Private Function ICustomerBO_ValidateCustomerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
' pass:         vXMLRequest  xml Request node containing the details of action to be
'                            performed and data required
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_ValidateCustomerDetailsVbErr

    Const strFunctionName As String = "ICustomerBO_ValidateCustomerDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim strName  As String
    Dim intChar As Integer

    Dim intMinAge As Integer
    Dim intMaxAge As Integer
    Dim intAge    As Integer
    Dim strDoB As String
    Dim dtDOB  As Date
    Dim strApplicationDate As String   ' EP903
    Dim strDate As String              ' EP903
    Dim dtDate As Date                 ' EP903
    Dim lngAge As Long                 ' EP903

    Dim objGP As New GlobalParameterDO
    Dim xmlResponseElem As IXMLDOMElement
        
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    If objContext Is Nothing Then
        Set objGP = New GlobalParameterDO
    Else
        Set objGP = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If

    '
    ' First character of surname must be alpha
    '
    strName = Trim$(objXmlAssist.GetTagValue(vxmlRequest, "SURNAME"))
    If strName <> "" Then
        intChar = Asc(UCase$(strName))

        If Not (((intChar >= 65) And (intChar <= 90))) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 166, _
                strName & " is not a valid surname"
        End If
    End If

    '
    ' First character of forenamemust be alpha
    '
    strName = Trim$(objXmlAssist.GetTagValue(vxmlRequest, "FIRSTFORENAME"))
    If strName <> "" Then
        intChar = Asc(UCase$(strName))

        If Not (((intChar >= 65) And (intChar <= 90))) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 166, _
                strName & " is not a valid forename"
        End If
    End If

    '
    ' Age must be in the correct range
    '
    'intMinAge = Val(objGP.GetCurrentParameter("MinimumAge"))
    intMinAge = objGP.GetCurrentParameterByType("MinimumAge", "AMOUNT")
    intMaxAge = objGP.GetCurrentParameterByType("MaximumAge", "AMOUNT")
    strDoB = objXmlAssist.GetTagValue(vxmlRequest, "DATEOFBIRTH")
    
    'EP903 Use Application Date, if present, to verify customer's age
    '      Make age verification consistent with other calculations.
    strApplicationDate = objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONDATE")
    
    If (Len(Trim(strApplicationDate)) > 0) Then
        strDate = strApplicationDate
    Else
        strDate = Now
    End If
    
    If (Len(Trim(strDoB)) > 0) Then
    
        'Calculate the applicant's age
        lngAge = Year(strDate) - Year(strDoB)
        
        If Month(strDate) < Month(strDoB) Then
            lngAge = lngAge - 1
        Else
            If Month(strDate) = Month(strDoB) Then
                If Day(strDate) < Day(strDoB) Then
                    lngAge = lngAge - 1
                End If
            End If
        End If
        
        'check the age
        If lngAge < intMinAge Then
            ' Too young
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 164
        End If
        
        If lngAge > intMaxAge Then
            ' Too old
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 165
        End If

    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    '-----------------------------------------------------------------------------
    
    Set ICustomerBO_ValidateCustomerDetails = xmlResponseElem.cloneNode(True)
    
ICustomerBO_ValidateCustomerDetailsExit:
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objGP = Nothing
    Set xmlResponseElem = Nothing

Exit Function

ICustomerBO_ValidateCustomerDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_ValidateCustomerDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_ValidateCustomerDetailsExit

End Function

Public Function DeleteAllCustomerTelephone(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:       xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllCustomerTelephoneVbErr

    Const strFunctionName As String = "DeleteAllCustomerTelephone"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' call FreeThreadedDOMDocument40 based function
    Set xmlTempResponseNode = ICustomerBO_DeleteAllCustomerTelephone(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DeleteAllCustomerTelephone = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteAllCustomerTelephoneExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

DeleteAllCustomerTelephoneVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAllCustomerTelephone = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume DeleteAllCustomerTelephoneExit

End Function

Public Function UpdatePersonalDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdatePersonalDetails a single instance of the persistant data associated with this
'               business object
' pass:         vstrXMLRequest  xml Request data stream containing data UpdatePersonalDetails
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePersonalDetailsVbErr

    Const strFunctionName As String = "UpdatePersonalDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
                                                            
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdatePersonalDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdatePersonalDetails = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdatePersonalDetailsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

UpdatePersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdatePersonalDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume UpdatePersonalDetailsExit

End Function

Public Function UpdateTenancy(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateTenancy a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTenancyVbErr

    Const strFunctionName As String = "UpdateTenancy"

    
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
    Set xmlDirectoryGUIDNode = xmlIn.documentElement.selectSingleNode(".//TENANCY/DIRECTORYGUID")
    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        If (Trim$(xmlDirectoryGUIDNode.Text) <> "") And Not (strDirectoryAddressIndicator = "1") Then
            ' entity is not linked to a directory record so empty the linking GUID
            xmlDirectoryGUIDNode.Text = ""
        End If
    End If

    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateTenancy(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateTenancy = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateTenancyExit:

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

UpdateTenancyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateTenancy = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateTenancyExit
    
End Function

Public Function UpdateCustomerPersonalDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateCustomerPersonalDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data UpdateCustomerPersonalDetails
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerPersonalDetailsVbErr

    Const strFunctionName As String = "UpdateCustomerPersonalDetails"


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
    Dim objIScriptInterface As IScriptInterface

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Load the Input into DOM(request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCustomerPersonalDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    UpdateCustomerPersonalDetails = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

UpdateCustomerPersonalDetailsExit:

    '   clean-up section, release object references
     Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing

    Exit Function

UpdateCustomerPersonalDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCustomerPersonalDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If

    '   go to clean-up section
    Resume UpdateCustomerPersonalDetailsExit

End Function

Public Function FindDependantList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindDependantListVbErr

    Const strFunctionName As String = "FindDependantList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

     ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_FindDependantList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
  
    FindDependantList = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindDependantListExit:

    Set xmlResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

FindDependantListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindDependantList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume FindDependantListExit

End Function

Private Function ICustomerBO_SaveDependant(ByVal vxmlRequest As IXMLDOMElement, _
                                        ByVal vxmlGeneratedKeys As IXMLDOMElement, _
                                        Optional ByRef rbotOperationType As Variant) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Saves the Dependant records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vXMLRequest  xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_SaveDependantVbErr

    Const strFunctionName As String = "ICustomerBO_SaveDependant"


    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = vxmlRequest.ownerDocument.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objICustomerTxBO As ICustomerTxBO
    
    If m_objContext Is Nothing Then
        Set objICustomerTxBO = New CustomerTxBO
    Else
        Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
    End If
    
    If Not IsMissing(rbotOperationType) Then
        Set xmlTempResponseNode = objICustomerTxBO.SaveDependant(vxmlRequest, vxmlGeneratedKeys, rbotOperationType)
    Else
        Set xmlTempResponseNode = objICustomerTxBO.SaveDependant(vxmlRequest, vxmlGeneratedKeys)
    End If
        
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
        
    Set ICustomerBO_SaveDependant = xmlResponseElem
               
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ICustomerBO_SaveDependantExit:

    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ICustomerBO_SaveDependantVbErr:
    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ICustomerBO_SaveDependant = m_objErrAssist.CreateErrorResponse

    '   go to clean-up section
    Resume ICustomerBO_SaveDependantExit
End Function

Public Function SaveDependant(ByVal vstrXMLRequest As String, _
                              Optional ByRef rbotOperationType As Variant) As String
' header ----------------------------------------------------------------------------------
' description:  Saves the Dependant records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo SaveDependantVbErr

    Const strFunctionName As String = "SaveDependant"


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
    Dim xmlGeneratedKeys As IXMLDOMElement
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
     'Validate the Input (request string)
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        
    Set xmlGeneratedKeys = xmlOut.createElement("DEPENDANTSLIST")
        
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SaveDependant(xmlIn.documentElement, _
                                                    xmlGeneratedKeys, rbotOperationType)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveDependant = xmlResponseElem.xml
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveDependantExit:
    Set objErrAssist = Nothing

    Exit Function

SaveDependantVbErr:
    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveDependant = objErrAssist.CreateErrorResponse

    '   go to clean-up section
    Resume SaveDependantExit

End Function

Private Sub ICustomerBO_FindCustomer(ByVal vobjXmlRequest As IXMLDOMNode, ByVal vobjXmlResponse As IXMLDOMNode)
On Error GoTo FindCustomerVbErr

    Const cstrFunctionName As String = "ICustomerBO_FindCustomer"
    
    
    Dim objCustomerDO As ICustomerDO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objConvAssist As ConvertAssist
    Dim xmlAdminRequestElem As IXMLDOMElement
    'Dim strAdminSystemState As String  'BMIDS00339 BMids does not use AdminSystemState
    Dim xmlAdminResponseDoc As FreeThreadedDOMDocument40
    Dim xmlAdminCustListElem As IXMLDOMNode
    Dim xmlOmigaCustListElem As IXMLDOMElement
    Dim xmlOutCustomerList As IXMLDOMElement
    Dim xmlAdminCustomerList As IXMLDOMNodeList
    Dim xmlOmigaCustomerList As IXMLDOMNodeList
    Dim xmlAdminCustomer As IXMLDOMElement
    Dim xmlOmigaCustomer As IXMLDOMElement
    Dim xslDoc As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    Dim objAdminBO As Object
  
    Dim strDoAdminSearch As String
    Dim blnDoAdminSearch As Boolean
    Dim strUseAdminDetails As String
    Dim blnUseAdminDetails As Boolean
    Dim blnDoingOmigaSearch As Boolean
    Dim blnDoingAdminSearch As Boolean
    Dim blnRetrievingGlobalParam As Boolean
    Dim strOmigaCustomerNo As String
    'Dim strAdminCustomerNo As String
    'Dim intNoOfAdminCustomers As Integer
    Dim intNoOfOmigaCustomers As Integer
    'Dim intAdminIndex As Integer
    'Dim intOmigaIndex As Integer
    Dim strDeDuplicateCustomer As String
    Dim blnDeDuplicateCustomer As Boolean
    Dim strAdminResponse As String
    'BMIDS00410
    Dim strResponseType As String
    Dim xmlMessage      As IXMLDOMNode
    Dim xmlMessageText  As IXMLDOMNode
    Dim xmlMessageType  As IXMLDOMNode
    'BMIDS00410 End
    Dim strOmigaOtherSysCustomerNo As String 'BMIDS00443
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
    
    ' PSC 06/10/2005 MAR57 - Start
    Dim strDoOmigaSearch As String
    Dim blnDoOmigaSearch As Boolean
    ' PSC 06/10/2005 MAR57 - End
    
    Set objConvAssist = New ConvertAssist
    
    If m_objContext Is Nothing Then
        Set objCustomerDO = New CustomerDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objCustomerDO = _
            m_objContext.CreateInstance(App.Title & ".CustomerDO")
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
   
    ' See if we need to do an administration system search
    strDoAdminSearch = objGlobalParamDO.GetCurrentParameterByType("FindLegacyCustomer", "BOOLEAN")
    blnDoAdminSearch = objConvAssist.CSafeBool(strDoAdminSearch)
    
    ' PSC 06/10/2005 MAR57 - Start
    strDoOmigaSearch = objGlobalParamDO.GetCurrentParameterByType("UseFindCustomer", "BOOLEAN")
    blnDoOmigaSearch = objConvAssist.CSafeBool(strDoOmigaSearch)
    ' PSC 06/10/2005 MAR57 - End
    
    ' BMIDS00339 BMIDS does not use AdminSystemState
    ' AQR SYS4373 - Check the AdminSystemState
    'strAdminSystemState = m_objXmlAssist.GetAttributeValue(vobjXmlRequest, "REQUEST", "ADMINSYSTEMSTATE")
    'blnDoAdminSearch = ((strAdminSystemState <> "null") And blnDoAdminSearch)
    'BMIDS00339 End
    
    ' Only create admin interface if we need to do admin search
    If blnDoAdminSearch = True Then
        If m_objContext Is Nothing Then
            Set objAdminBO = CreateObject(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        Else
            Set objAdminBO = _
                m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        End If
    End If

    ' PSC 06/10/2005 MAR57 - Start
    If blnDoOmigaSearch Then
        blnDoingOmigaSearch = True
        objCustomerDO.FindCustomer vobjXmlRequest, vobjXmlResponse
    End If
    ' PSC 06/10/2005 MAR57 - End
    
    blnDoingOmigaSearch = False
    
    ' Get a list of Omiga customers and remove them from the response as the list needs
    ' to be processed and attached in a different order
    Set xmlOmigaCustListElem = m_objXmlAssist.GetNode(vobjXmlResponse, ".//CUSTOMERLIST", False)

    If Not xmlOmigaCustListElem Is Nothing Then
        Set xmlOmigaCustomerList = vobjXmlResponse.selectNodes(".//CUSTOMER")
        intNoOfOmigaCustomers = xmlOmigaCustomerList.length
        vobjXmlResponse.removeChild xmlOmigaCustListElem
    End If
        
    ' If this implementation has an admin system do a search on the Admin System and
    ' merge the results with the Omiga results
    If blnDoAdminSearch = True Then
            
        ' Determine whether Omiga details or the Admin System details are to be used in
        ' the event of a customer being on both systems
        strUseAdminDetails = objGlobalParamDO.GetCurrentParameterByType("UseAdminDetails", "BOOLEAN")
        blnUseAdminDetails = objConvAssist.CSafeBool(strUseAdminDetails)
        
        ' Convert input request to attribute based XML for call to AdminInterfaceBO
        Set xmlAdminRequestElem = xmlCreateAttributeBasedResponse(vobjXmlRequest, True)
        xmlAdminRequestElem.setAttribute "OPERATION", "FindCustomer"
        
        'BMIDS00391 CUSTOMER node should be directly under REQUEST
        m_objXmlAssist.PromoteNode xmlAdminRequestElem, m_objXmlAssist.GetNode(xmlAdminRequestElem, ".//CUSTOMER"), True
        'BMIDS00391 End
        
        ' Search the administration system
        blnDoingAdminSearch = True
        strAdminResponse = objAdminBO.Request(xmlAdminRequestElem.xml)
        Set xmlAdminResponseDoc = m_objXmlAssist.load(strAdminResponse, TypeName(Me), cstrFunctionName)
        
        'BMIDS00410 Convert admin system errors to warnings
        strResponseType = xmlGetAttributeText(xmlAdminResponseDoc.documentElement, "TYPE")
        If (strResponseType <> "SUCCESS") And (strResponseType <> "WARNING") Then
            xmlAdminResponseDoc.documentElement.setAttribute "TYPE", "WARNING"
            
            Set xmlMessage = xmlAdminResponseDoc.createElement("MESSAGE")
            xmlAdminResponseDoc.documentElement.appendChild xmlMessage
            
            Set xmlMessageType = xmlAdminResponseDoc.createElement("MESSAGETYPE")
            xmlMessageType.Text = "WARNING"
            xmlMessage.appendChild xmlMessageType
            
            Set xmlMessageText = xmlAdminResponseDoc.createElement("MESSAGETEXT")
            xmlMessageText.Text = xmlGetNodeText(xmlAdminResponseDoc, "RESPONSE/ERROR/DESCRIPTION") _
                & vbCrLf & xmlGetNodeText(xmlAdminResponseDoc, "RESPONSE/ERROR/SOURCE")
                '& vbCrLf & xmlGetNodeText(xmlAdminResponseDoc, "RESPONSE/ERROR/NUMBER")
            xmlMessage.appendChild xmlMessageText
        End If
        m_objErrAssist.CheckResponse xmlAdminResponseDoc.documentElement, vobjXmlResponse, False
        'BMIDS00410 End
        
        blnDoingAdminSearch = False
        
        ' Convert the response back into element based and get a list of admin customers
        Set xmlAdminCustListElem = m_objXmlAssist.GetNode(xmlAdminResponseDoc, ".//CUSTOMERLIST", False)
        
        If Not xmlAdminCustListElem Is Nothing Then
            Set xmlAdminCustListElem = xmlMakeNodeElementBased(xmlAdminCustListElem, True, "CUSTOMERLIST")
            Set xmlAdminCustomerList = xmlAdminCustListElem.selectNodes(".//CUSTOMER")
            'intNoOfAdminCustomers = xmlAdminCustomerList.length
        End If
                    
        ' Create a list to hold the merged data
        Set xmlOutCustomerList = vobjXmlResponse.ownerDocument.createElement("CUSTOMERLIST")
            
'        intAdminIndex = 0
'        intOmigaIndex = 0
        
        'SR 12/06/01 : SYS2362 - find the value of the global parameter 'DeDuplicateCustomer'.
        '              If yes, remove the duplicates in Customer records.
        blnDeDuplicateCustomer = False
        
        blnRetrievingGlobalParam = True
        strDeDuplicateCustomer = objGlobalParamDO.GetCurrentParameterByType("DeDuplicateCustomer", "BOOLEAN")
        blnRetrievingGlobalParam = False
        
        If strDeDuplicateCustomer <> "" Then
            blnDeDuplicateCustomer = objConvAssist.CSafeBool(strDeDuplicateCustomer)
        End If
        
        If blnDeDuplicateCustomer Then
            'BMIDS00443 Deduping algorithm changed so that it is not dependent on customer lists being
            'sorted, and matches on both Omiga customer numbers and Other system customer numbers
            If (Not xmlAdminCustomerList Is Nothing) And (Not xmlOmigaCustomerList Is Nothing) Then
                    
                For Each xmlOmigaCustomer In xmlOmigaCustomerList
                                    
                    strOmigaCustomerNo = m_objXmlAssist.GetNodeValue(xmlOmigaCustomer, "CUSTOMERNUMBER", True, True)
                    strOmigaOtherSysCustomerNo = m_objXmlAssist.GetNodeValue(xmlOmigaCustomer, "OTHERSYSTEMCUSTOMERNUMBER", True, False)
                    strDeDuplicateCustomer = "CUSTOMER[(OMIGACUSTOMERNUMBER[.='" & strOmigaCustomerNo & "'] or OTHERSYSTEMCUSTOMERNUMBER[.='" & strOmigaOtherSysCustomerNo & "'])]"
                    Set xmlAdminCustomer = xmlAdminCustListElem.selectSingleNode(strDeDuplicateCustomer)
                    If Not xmlAdminCustomer Is Nothing Then
                        'Duplicate found - choose which one to remove
                        If blnUseAdminDetails Then
                            xmlOmigaCustListElem.removeChild xmlOmigaCustomer
                        Else
                            xmlAdminCustListElem.removeChild xmlAdminCustomer
                        End If
                    End If
                Next
                
                'Update customer lists as some customers may have been removed
                Set xmlOmigaCustomerList = xmlOmigaCustListElem.selectNodes(".//CUSTOMER")
                Set xmlAdminCustomerList = xmlAdminCustListElem.selectNodes(".//CUSTOMER")
            End If
            'BMIDS00443 End
                        
'            While intAdminIndex < intNoOfAdminCustomers Or intOmigaIndex < intNoOfOmigaCustomers
'
'                ' Get current Omiga Customer
'                If intNoOfOmigaCustomers > 0 And intOmigaIndex < intNoOfOmigaCustomers Then
'                    Set xmlOmigaCustomer = xmlOmigaCustomerList.Item(intOmigaIndex)
'                    strOmigaCustomerNo = m_objXmlAssist.GetNodeValue(xmlOmigaCustomer, _
'                                                                     "OTHERSYSTEMCUSTOMERNUMBER", _
'                                                                     True, False)
'                Else
'                    strOmigaCustomerNo = "ZZZZZZZZZZZZZZZZZZZZ"
'                End If
'
'                ' Get current Admin Customer
'                If intNoOfAdminCustomers > 0 And intAdminIndex < intNoOfAdminCustomers Then
'                    Set xmlAdminCustomer = xmlAdminCustomerList.Item(intAdminIndex)
'                    strAdminCustomerNo = m_objXmlAssist.GetNodeValue(xmlAdminCustomer, _
'                                                                     "OTHERSYSTEMCUSTOMERNUMBER")
'                Else
'                    strAdminCustomerNo = "ZZZZZZZZZZZZZZZZZZZZ"
'                End If
'
'                If strOmigaCustomerNo = strAdminCustomerNo Then
'
'                    ' Check which details to use
'                    If blnUseAdminDetails = False Then
'                        xmlOutCustomerList.appendChild xmlOmigaCustomer
'                    Else
'                        Set xmlNode = m_objXmlAssist.GetNode(xmlOmigaCustomer, "CUSTOMERNUMBER")
'                        xmlAdminCustomer.appendChild xmlNode
'                        Set xmlNode = m_objXmlAssist.GetNode(xmlOmigaCustomer, ".//CUSTOMERVERSIONNUMBER")
'                        xmlAdminCustomer.appendChild xmlNode
'                        xmlOutCustomerList.appendChild xmlAdminCustomer
'                    End If
'                    intOmigaIndex = intOmigaIndex + 1
'                    intAdminIndex = intAdminIndex + 1
'                ElseIf strOmigaCustomerNo < strAdminCustomerNo Then
'                    xmlOutCustomerList.appendChild xmlOmigaCustomer
'                    intOmigaIndex = intOmigaIndex + 1
'                Else
'                    xmlOutCustomerList.appendChild xmlAdminCustomer
'                    intAdminIndex = intAdminIndex + 1
'                End If
'            Wend
        
        End If

        ' Merge the results
        'DS SYS4788 If there is no Omiga customer list, don't do the loop.
        If Not xmlOmigaCustomerList Is Nothing Then
            For Each xmlOmigaCustomer In xmlOmigaCustomerList
                xmlOutCustomerList.appendChild xmlOmigaCustomer
            Next xmlOmigaCustomer
        End If
        
        'BMIDS00410 If there is no Admin customer list, don't to the loop
        If Not xmlAdminCustomerList Is Nothing Then
            For Each xmlAdminCustomer In xmlAdminCustomerList
                xmlOutCustomerList.appendChild xmlAdminCustomer
            Next xmlAdminCustomer
        End If

    Else
        ' If only an Omiga search has been carried out then use this as the merged list
        If Not xmlOmigaCustListElem Is Nothing Then
            Set xmlOutCustomerList = xmlOmigaCustListElem
        End If
    End If

    ' Sort the list into the appropriate order and attach the result to the response
    If Not xmlOutCustomerList Is Nothing Then
        Set xslDoc = New FreeThreadedDOMDocument40
        xslDoc.validateOnParse = False
        xslDoc.setProperty "NewParser", True
        xslDoc.async = False

        strXslPath = App.Path & "\FindCustomer.xslt"
        strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
        blnLoadedOk = xslDoc.load(strXslPath)
        
        If Not blnLoadedOk Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
        End If
        
        Set xmlTransformedDoc = New FreeThreadedDOMDocument40
        xmlTransformedDoc.validateOnParse = False
        xmlTransformedDoc.setProperty "NewParser", True
        xmlTransformedDoc.async = False
        
        Dim xmlCustomersDoc As FreeThreadedDOMDocument40
        Set xmlCustomersDoc = New FreeThreadedDOMDocument40
        xmlCustomersDoc.validateOnParse = False
        xmlCustomersDoc.setProperty "NewParser", True
        xmlCustomersDoc.async = False
        
        xmlCustomersDoc.loadXML (xmlOutCustomerList.xml)
        
        xmlCustomersDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
        
        ' Add the number of found customers as an attribute of the converted customerlist node
        Set xmlOutCustomerList = xmlTransformedDoc.selectSingleNode("CUSTOMERLIST")
        xmlOutCustomerList.setAttribute "TOTAL", xmlOutCustomerList.selectNodes("CUSTOMER").length
        
        vobjXmlResponse.appendChild xmlTransformedDoc.documentElement
    End If
         
    Set objCustomerDO = Nothing
    Set objCustomerDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlAdminResponseDoc = Nothing
    Set xmlAdminCustListElem = Nothing
    Set xmlOutCustomerList = Nothing
    Set xmlAdminCustomerList = Nothing
    Set xmlOmigaCustListElem = Nothing
    Set xmlOmigaCustomerList = Nothing
    Set xmlAdminCustomer = Nothing
    Set xmlOmigaCustomer = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    Set objAdminBO = Nothing
    'BMIDS00410
    Set xmlMessage = Nothing
    Set xmlMessageText = Nothing
    Set xmlMessageType = Nothing
    'BMIDS00410 End
    
    Exit Sub
    
FindCustomerVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        
        ' Resume if we have no records from Omiga and we still have to do an Admin Search or
        ' We have no records on the Admin System  but we have records from Omiga
        If (blnDoingOmigaSearch = True And blnDoAdminSearch = True) Or _
           (blnDoingAdminSearch = True And intNoOfOmigaCustomers > 0) Or _
            blnRetrievingGlobalParam Then
            Resume Next
        End If
    End If

    Set objCustomerDO = Nothing
    Set objCustomerDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlAdminResponseDoc = Nothing
    Set xmlAdminCustListElem = Nothing
    Set xmlOutCustomerList = Nothing
    Set xmlAdminCustomerList = Nothing
    Set xmlOmigaCustListElem = Nothing
    Set xmlOmigaCustomerList = Nothing
    Set xmlAdminCustomer = Nothing
    Set xmlOmigaCustomer = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    Set objAdminBO = Nothing
    'BMIDS00410
    Set xmlMessage = Nothing
    Set xmlMessageText = Nothing
    Set xmlMessageType = Nothing
    'BMIDS00410 End
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ICustomerBO_GetCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, ByVal vobjResponseNode As IXMLDOMNode)
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objICustomerDO As ICustomerDO

    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objICustomerDO.GetCustomerDetails vobjRequestNode, vobjResponseNode
    
    Set objICustomerDO = Nothing
    Set objContext = Nothing

End Sub

Private Sub ICustomerBO_GetAndSynchroniseCustomerDetails( _
    ByVal vobjRequestNode As IXMLDOMNode, _
    ByVal vobjResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Gets customer details.  If customer is only on Omiga database, get the
'               data and return it.  If there is an additional Admin system in use and the
'               customer is on that aswell, get both Omiga and Admin data, merge them then
'               update Omiga with the data and return the data held in Omiga after the update
'               If the customer is not on the Omiga system, but is on the Admin system, Create
'               them on the Omiga system and send back the newly created data.
'History:  BG 06/02/2001 SYS1866 - added functionality to deal with an additional admin system
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName As String = "ICustomerBO_GetAndSynchroniseCustomerDetails"
    
    Dim strCustomerNumber               As String
    Dim strCustVersionNum               As String
    Dim strOtherSysCustNo               As String
    Dim blnFindLegacyCustomer           As Boolean
    Dim blnUseAdminDetails              As Boolean
    Dim objGlobalParameterDO            As GlobalParameterDO
    Dim objICustomerTxBO                As ICustomerTxBO
    Dim xmlTempDoc                      As FreeThreadedDOMDocument40
    Dim xmlTempNode                     As IXMLDOMNode
    Dim xmlCreateNode                   As IXMLDOMNode
    Dim xmlTempResponseNode             As IXMLDOMNode
    Dim xmlCustomerNode                 As IXMLDOMNode
    Dim xmlTempReqNode                  As IXMLDOMNode
    Dim xmlCustomerNumberNode           As IXMLDOMNode
    Dim xmlCustomerVersionNumberNode    As IXMLDOMNode
    Dim xmlElement                      As IXMLDOMElement
    'BMIDS00443
    Dim xmlAdminResponseNode            As IXMLDOMNode
    Dim blnRetrieveOmigaCustomer        As Boolean
    Dim blnUpdateOmigaCustomer          As Boolean
    Dim blnRequireCustomerLock          As Boolean
    Dim blnOmigaCustomerExists          As Boolean
    'BMIDS00443 End
    'MO - 13/11/2002 - BMIDS00921
    Dim xmlReadOnlyAttribute            As IXMLDOMAttribute
    Dim xmlChangedData                  As IXMLDOMAttribute     ' PSC 26/10/2005 MAR300
 
    ' PSC 11/10/2005 MAR57 - Start
    Dim xmlCustomerDataList As IXMLDOMNode
    Dim xmlOmigaCustomerData As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerIn As IXMLDOMNode
    Dim xmlCriticalContext As IXMLDOMNode
    Dim xmlSearch As IXMLDOMNode
    
    Dim blnRunCustomerDataCheck As Boolean
    Dim blnLockOverride As Boolean
    ' PSC 11/10/2005 MAR57 - End
    Dim blnDataChanged As Boolean           ' PSC 26/10/2005 MAR300
    
    
On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
    
    ' PSC 11/10/2005 MAR57 - Start
    ' Validate input
    Set xmlSearch = vobjRequestNode.selectSingleNode(".//SEARCH")
    Set xmlCustomerList = xmlSearch.selectNodes(".//CUSTOMER")
    
    blnRunCustomerDataCheck = xmlGetAttributeAsBoolean(xmlSearch, "CUSTOMERDATACHECK", "0")
    
    For Each xmlCustomerIn In xmlCustomerList
        ' PSC 12/01/01 SYS1750 - Start
        strCustomerNumber = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "CUSTOMERNUMBER", False, False)
            
        strCustVersionNum = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "CUSTOMERVERSIONNUMBER", False, False)
         
        strOtherSysCustNo = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "OTHERSYSTEMCUSTOMERNUMBER", False, False)
        
        ' Must have Customer Number and Customer Version Number and/or Other System Customer Number
        If (Len(strCustomerNumber) > 0 And Len(strCustVersionNum) = 0 And Len(strOtherSysCustNo) = 0) Or _
           (Len(strCustVersionNum) > 0 And Len(strCustomerNumber) = 0) Or _
           (Len(strCustVersionNum) = 0 And Len(strCustomerNumber) = 0 And Len(strOtherSysCustNo) = 0) Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                      cstrFunctionName, _
                                      oeMissingPrimaryTag, _
                                      "CUSTOMERNUMBER = " & strCustomerNumber & _
                                      "; CUSTOMERVERSIONNUMBER = " & strCustVersionNum & _
                                      "; OTHERSYSTEMCUSTOMERNUMBER = " & strOtherSysCustNo
        End If
    Next
    
    If blnRunCustomerDataCheck Then
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        Set xmlCustomerDataList = xmlTempDoc.createElement("CUSTOMERDATALIST")
    End If
  
    ' Get and Sync for each customer passed in
    For Each xmlCustomerIn In xmlCustomerList
    
        ' PSC 20/12/2005 MAR917 - Start
        strCustomerNumber = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "CUSTOMERNUMBER", False, False)
            
        strCustVersionNum = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "CUSTOMERVERSIONNUMBER", False, False)
         
        strOtherSysCustNo = _
            m_objXmlAssist.GetNodeValue(xmlCustomerIn, "OTHERSYSTEMCUSTOMERNUMBER", False, False)
        ' PSC 20/12/2005 MAR917 - End

        'BMIDS00443
        blnRetrieveOmigaCustomer = False
        blnUpdateOmigaCustomer = False
        blnRequireCustomerLock = False
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        
        blnLockOverride = xmlGetAttributeAsBoolean(xmlCustomerIn, "NOLOCK")
        
        If Len(strOtherSysCustNo) > 0 Then
            ' Check if there is an existing Omiga customer with a matching OtherSystemCustomerNumber
            Set xmlTempReqNode = vobjRequestNode.cloneNode(False)
            Set xmlElement = xmlTempDoc.createElement("CUSTOMER")
            xmlTempReqNode.appendChild xmlElement
            xmlElement.appendChild xmlCustomerIn.selectSingleNode("OTHERSYSTEMCUSTOMERNUMBER").cloneNode(True)
            
            On Error GoTo IgnoreRecordNotFound
            Set xmlTempResponseNode = ICustomerBO_GetCustomerFromLegacyNumber(xmlTempReqNode)
            m_objErrAssist.CheckResponse xmlTempResponseNode, vobjResponseNode, True
    
            On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
            
            'Create an instance of GlobalParameterDO
            Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
            
            'Find out if Admin component is available
            blnFindLegacyCustomer = objGlobalParameterDO.GetCurrentParameterByType("FindLegacyCustomer", "BOOLEAN")
            'Find out if Admin customer details should be used in preference to Omiga details
            blnUseAdminDetails = objGlobalParameterDO.GetCurrentParameterByType("UseAdminDetails", "BOOLEAN")
                
            If Not xmlTempResponseNode.selectSingleNode("CUSTOMER") Is Nothing Then
                'The Customer already exists on Omiga
                
                If Len(strCustomerNumber) = 0 Then
                    strCustomerNumber = m_objXmlAssist.GetNodeValue(xmlTempResponseNode, "CUSTOMER/CUSTOMERVERSION/CUSTOMERNUMBER")
                    xmlCustomerIn.selectSingleNode("CUSTOMERNUMBER").Text = strCustomerNumber
                End If
                
                If Len(strCustVersionNum) = 0 Then
                    strCustVersionNum = m_objXmlAssist.GetNodeValue(xmlTempResponseNode, "CUSTOMER/CUSTOMERVERSION/CUSTOMERVERSIONNUMBER")
                    xmlCustomerIn.selectSingleNode("CUSTOMERVERSIONNUMBER").Text = strCustVersionNum
                End If
                
                blnRequireCustomerLock = True
                If blnFindLegacyCustomer And blnUseAdminDetails Then
                    blnUpdateOmigaCustomer = True
                Else
                    blnRetrieveOmigaCustomer = True
                End If
            Else    ' No existing Omiga customer with a matching OtherSystemCustomerNumber
                
                If blnFindLegacyCustomer Then
                
                    If Len(strCustomerNumber) > 0 Then
                    
                        ' Check if there is an existing Omiga customer with a matching CustomerNumber
                        Set xmlTempReqNode = vobjRequestNode.cloneNode(False)
                        Set xmlElement = xmlTempDoc.createElement("SEARCH")
                        xmlTempReqNode.appendChild xmlElement
                        Set xmlCustomerNode = xmlTempDoc.createElement("CUSTOMER")
                        xmlElement.appendChild xmlCustomerNode
                        xmlCustomerNode.appendChild xmlCustomerIn.selectSingleNode("CUSTOMERNUMBER").cloneNode(True)
                        
                        On Error GoTo IgnoreRecordNotFound
                        Set xmlTempResponseNode = FindLatestCustomerVersion(xmlTempReqNode)
                        m_objErrAssist.CheckResponse xmlTempResponseNode, vobjResponseNode, True
                
                        On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
                    
                        If Not xmlTempResponseNode.selectSingleNode("CUSTOMERVERSION") Is Nothing Then
                            blnOmigaCustomerExists = True
                        Else
                            blnOmigaCustomerExists = False
                        End If
                    Else
                        blnOmigaCustomerExists = False
                    End If
    
                    If blnOmigaCustomerExists Then
                        'The Customer already exists on Omiga
                        strCustVersionNum = m_objXmlAssist.GetNodeValue(xmlTempResponseNode, "CUSTOMERVERSION/CUSTOMERVERSIONNUMBER")
                        xmlCustomerIn.selectSingleNode("CUSTOMERVERSIONNUMBER").Text = strCustVersionNum
                        
                        blnRequireCustomerLock = True
                        If blnUseAdminDetails Then
                            blnUpdateOmigaCustomer = True
                        Else
                            blnRetrieveOmigaCustomer = True
                        End If
                        
                    Else
                        'Get customer details from the admin system
                        Set xmlAdminResponseNode = GetAdminCustomer(vobjRequestNode, strOtherSysCustNo)
                        
                        'Create a new Omiga Customer
                        Set xmlTempNode = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
                        Set xmlCreateNode = xmlTempDoc.createElement("CREATE")
            
                        If Len(strOtherSysCustNo) > 0 Then
                            xmlCreateNode.appendChild xmlAdminResponseNode
                        End If
            
                        xmlTempNode.appendChild xmlCreateNode
            
                        'Create an instance of CustomerTxBO
                        Set objICustomerTxBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerTxBO")
            
                        'Call CustomerTxBO.CreateCustomerDetails passing in the converted XML String
                        objICustomerTxBO.CreateCustomerDetails xmlTempNode, vobjResponseNode
                        m_objErrAssist.ReRaiseResponseError vobjResponseNode
            
                        Set xmlTempResponseNode = xmlTempDoc.createElement("RESPONSE")
            
                        Set xmlCreateNode = xmlTempDoc.createElement("SEARCH")
                        Set xmlTempReqNode = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
                        xmlTempDoc.appendChild xmlTempReqNode
                        xmlTempReqNode.appendChild xmlCreateNode
                        Set xmlTempNode = m_objXmlAssist.GetNode(vobjResponseNode, ".//CUSTOMERKEY")
                        xmlCreateNode.appendChild xmlTempNode
                        m_objXmlAssist.ChangeNodeName xmlTempNode, "CUSTOMERKEY", "CUSTOMER"
            
                        'Call GetCustomer Details to get the refreshed data and pass it back
                        ICustomerBO_GetCustomerDetails xmlTempReqNode, vobjResponseNode
                    End If
                End If
            End If
        Else    'No OtherSystemCustomerNumber exists
            blnRetrieveOmigaCustomer = True
            blnRequireCustomerLock = True
        End If
        'BMIDS00443 End
        
        If blnRequireCustomerLock And Not blnLockOverride Then
            'Create a customer lock and check if any errors occurred
            Set xmlTempReqNode = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
            Set xmlCustomerNode = xmlTempDoc.createElement("CUSTOMER")
            xmlTempReqNode.appendChild xmlCustomerNode
            Set xmlCustomerNumberNode = xmlTempDoc.createElement("CUSTOMERNUMBER")
            xmlCustomerNumberNode.Text = strCustomerNumber
            xmlCustomerNode.appendChild xmlCustomerNumberNode
            
            On Error Resume Next
            ICustomerBO_CreateCustomerLock xmlTempReqNode, vobjResponseNode
        
            If Err.Number <> 0 Then
                If Err.Number = omiga4DuplicateKey Then
                    'MO - 13/11/2002 - BMIDS00921 - Start
                    'Set xmlElement = vobjResponseNode.ownerDocument.createAttribute("READONLY")
                    'xmlElement.Text = "1"
                    'vobjResponseNode.Attributes.setNamedItem xmlElement
                    Set xmlReadOnlyAttribute = vobjResponseNode.ownerDocument.createAttribute("READONLY")
                    xmlReadOnlyAttribute.Text = "1"
                    vobjResponseNode.Attributes.setNamedItem xmlReadOnlyAttribute
                    Set xmlReadOnlyAttribute = Nothing
                    'MO - 13/11/2002 - BMIDS00921 - End
                Else
                    On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
                    Err.Raise Err.Number, Err.Source, Err.Description
                End If
            End If
            On Error GoTo GetAndSynchroniseCustomerDetailsVbErr
        End If
        
        If blnRetrieveOmigaCustomer Or blnRunCustomerDataCheck Then
            'Retrieve Omiga Customer
            Set xmlTempReqNode = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
            Set xmlCreateNode = xmlTempDoc.createElement("SEARCH")
            xmlTempReqNode.appendChild xmlCreateNode
            Set xmlCustomerNode = xmlTempDoc.createElement("CUSTOMER")
            xmlCreateNode.appendChild xmlCustomerNode
            Set xmlCustomerNumberNode = xmlTempDoc.createElement("CUSTOMERNUMBER")
            xmlCustomerNumberNode.Text = strCustomerNumber
            xmlCustomerNode.appendChild xmlCustomerNumberNode
            Set xmlCustomerVersionNumberNode = xmlTempDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlCustomerVersionNumberNode.Text = strCustVersionNum
            xmlCustomerNode.appendChild xmlCustomerVersionNumberNode
    
            'Call GetCustomer Details to get the refreshed data and pass it back
            ICustomerBO_GetCustomerDetails xmlTempReqNode, vobjResponseNode
            
            ' If a customer data check is required copy the details for use later
            If blnRunCustomerDataCheck Then
                Set xmlOmigaCustomerData = vobjResponseNode.selectSingleNode("CUSTOMER")
                xmlCustomerDataList.appendChild xmlOmigaCustomerData
            End If
        End If
        
        If blnUpdateOmigaCustomer And Not blnRunCustomerDataCheck Then
            Set xmlAdminResponseNode = GetAdminCustomer(vobjRequestNode, strOtherSysCustNo)
            'Update Omiga Customer using admin details
            UpdateCustomerWithAdminDetails vobjRequestNode, vobjResponseNode, xmlAdminResponseNode
        End If
    Next
    
    If blnRunCustomerDataCheck Then
        Set xmlAdminResponseNode = GetAdminCustomer(vobjRequestNode)
        Set xmlCriticalContext = vobjRequestNode.selectSingleNode(".//CRITICALDATACONTEXT")
        ' PSC 26/10/2005 MAR300 - Start
        blnDataChanged = RunCustomerDataCheckAndUpdate(vobjRequestNode.cloneNode(True), xmlCriticalContext, xmlCustomerDataList, xmlAdminResponseNode)
        
        If blnDataChanged Then
            Set xmlChangedData = vobjResponseNode.ownerDocument.createAttribute("DATACHANGED")
            xmlChangedData.Text = "1"
            vobjResponseNode.Attributes.setNamedItem xmlChangedData
        End If
        ' PSC 26/10/2005 MAR300 - End
    End If
    ' PSC 11/10/2005 MAR57 - End

    
GetAndSynchroniseCustomerDetailsVbErr:
        
    Set objGlobalParameterDO = Nothing
    Set objICustomerTxBO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempNode = Nothing
    Set xmlCreateNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempReqNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNumberNode = Nothing
    Set xmlElement = Nothing
    Set xmlAdminResponseNode = Nothing
    
    ' PSC 11/10/2005 MAR57 - Start
    Set xmlCustomerDataList = Nothing
    Set xmlOmigaCustomerData = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomerIn = Nothing
    Set xmlCriticalContext = Nothing
    ' PSC 11/10/2005 MAR57 - End
    Set xmlChangedData = Nothing        ' PSC 26/10/2005 MAR300


    If Err.Number <> 0 Then
        ' re-raise error
        ' PSC 20/12/2005 MAR917
        Err.Raise Err.Number, cstrFunctionName & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Exit Sub

'BMIDS00443
IgnoreRecordNotFound:
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
'BMIDS00443 End

End Sub

Private Sub ICustomerBO_CreateCustomerLock( _
    ByVal vobjXmlRequestNode As IXMLDOMNode, _
    ByVal vobjXmlResponseNode As IXMLDOMNode, _
    Optional ByVal strLockRequired As String = "")
                                
On Error GoTo CreateCustomerLockVbErr

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    'BM0539  Get customer number for error message
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim strCustomerNumber As String
    strCustomerNumber = objXmlAssist.GetMandatoryElementText(vobjXmlRequestNode, ".//CUSTOMERNUMBER")
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    If strLockRequired = "" Then
        If Not IsLockRequiredInThisSystem Then
            GoTo CreateCustomerLockExit
        End If
    ElseIf strLockRequired = "0" Then
        GoTo CreateCustomerLockExit
    End If
    
    Dim objCustomerDO As ICustomerDO
    If objContext Is Nothing Then
        Set objCustomerDO = New CustomerDO
    Else
        Set objCustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objCustomerDO.CreateCustomerLock vobjXmlRequestNode, vobjXmlResponseNode
            
CreateCustomerLockExit:
    Set objCustomerDO = Nothing
    Set objContext = Nothing
    Set objErrAssist = Nothing  ' BM0539
    Set objXmlAssist = Nothing  ' BM0539
    
    Exit Sub
    
CreateCustomerLockVbErr:
        
    Set objCustomerDO = Nothing
    Set objContext = Nothing
    
    'BM0539 Display useful error message when duplicate key is detected
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
        Err.Description = "Customer " & strCustomerNumber & " is already locked to another application."
    End If
    
    ' re-raise error
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub ICustomerBO_DeleteCustomerLock( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByVal vobjXmlResponse As IXMLDOMNode, _
    Optional ByVal strLockRequired As String = "")

On Error GoTo DeleteCustomerLockVbErr

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
     '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem Then
            GoTo DeleteCustomerLockExit
        End If
    ElseIf strLockRequired = "0" Then
        GoTo DeleteCustomerLockExit
    End If
    
    Dim objXmlAssist As New XMLAssist
    Dim objCustomerDO As ICustomerDO

    If objContext Is Nothing Then
        Set objCustomerDO = New CustomerDO
    Else
        Set objCustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    objCustomerDO.DeleteCustomerLock vobjXmlRequest, vobjXmlResponse

DeleteCustomerLockExit:
    Set objCustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing
    
    Exit Sub
    
DeleteCustomerLockVbErr:
        
    Set objCustomerDO = Nothing
    Set objXmlAssist = Nothing
    Set objContext = Nothing
        
    ' re-raise error
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub


Private Sub ValidateCustomerDetails(ByVal vobjXmlRequest As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Performs the hard coded validation of Customer Details
' pass:         vobjXmlRequest  Request xml Document
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCustomerDetailsVbErr
    
    Const strFunctionName As String = "ValidateCustomerDetails"
     
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objCustomerElem As IXMLDOMElement
    Dim objCustVersionElem As IXMLDOMElement
    
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objCustomerDO As ICustomerDO
    
    Dim objGeneralAssist As New GeneralAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
'    Dim objScriptInterface As New ScriptInterface
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
        Set objCustomerDO = New CustomerDO
    Else
        Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objCustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If

    ' Get minimum and maximum ages
    Dim varMinimumAge As Variant, varMaximumAge As Variant
    Dim intMinimumAge As Integer, intMaximumAge As Integer
    
    Dim lngAge As Long
   
    ' Validate customer
    Set objCustomerElem = _
        objXmlAssist.GetMandatoryNode(vobjXmlRequest, "//CUSTOMER")
    
'    varMinimumAge = objGlobalParamDO.GetCurrentParameterByTypeSp("MinimumAge", "AMOUNT")
    varMinimumAge = objGlobalParamDO.GetCurrentParameterByType("MinimumAge", "AMOUNT")
    If Not IsEmpty(varMinimumAge) Then
        intMinimumAge = varMinimumAge
    End If
    
'    varMaximumAge = objGlobalParamDO.GetCurrentParameterByTypeSp("MaximumAge", "AMOUNT")
    varMaximumAge = objGlobalParamDO.GetCurrentParameterByType("MaximumAge", "AMOUNT")
    If Not IsEmpty(varMaximumAge) Then
        intMaximumAge = varMaximumAge
    End If
    
    Set objGlobalParamDO = Nothing
    
    ' Add any derived data
'    Set objXmlDerived = _
'        objXmlAssist.load(objCustomerDO.AddDerivedData(objCustomerElem.xml), TypeName(Me), strFunctionName)
'
'    objXmlAssist.ReplaceNode objCustomerElem, objXmlDerived.documentElement

    ' Derived values from Customer Version
    
    Set objCustVersionElem = _
        objXmlAssist.GetMandatoryNode(vobjXmlRequest, "//CUSTOMER/CUSTOMERVERSION")
    
    Dim strSurname As String, strFirstForename As String, strDoB As String
    ' SR 19/10/00 SYS1645 : Change the way the DoB is validated against the min and max age allowed
    '   objCustomerDO.CustomerDetailsDerivedData objCustomerElem
        
    strFirstForename = objXmlAssist.GetElementText(objCustVersionElem, "FIRSTFORENAME")
    strSurname = objXmlAssist.GetElementText(objCustVersionElem, "SURNAME")
    strDoB = objXmlAssist.GetElementText(objCustVersionElem, "DATEOFBIRTH")
    'strAge = objXmlAssist.GetElementText(objCustVersionElem, "AGE")
    
    'EP903  Use Application Date if present to verify customer's age
    Dim strApplicationDate As String
    Dim strDate As String
    strApplicationDate = objXmlAssist.GetElementText(objCustVersionElem, "APPLICATIONDATE")
    
    If (Len(strApplicationDate) > 0) Then
        strDate = strApplicationDate
    Else
        strDate = Now
    End If
        
    ' First character of Surname and Forename must be alpha
    If objGeneralAssist.IsAlpha(Left$(strSurname, 1)) = False Or _
       objGeneralAssist.IsAlpha(Left$(strFirstForename, 1)) = False Then
       objErrAssist.ThrowError TypeName(Me), strFunctionName, 166
    End If
    
'    If Len(strAge) > 0 And IsNumeric(strAge) = True Then
'        ' Can't be younger than minimum age
'        If CInt(strAge) < intMinimumAge Then
'            objErrAssist.ThrowError TypeName(Me), strFunctionName, 164
'        End If
'
'        ' Must be older than maximum age
'        If CInt(strAge) > intMaximumAge Then
'            objErrAssist.ThrowError TypeName(Me), strFunctionName, 165
'        End If
'    End If

    ' SR 19/10/00 SYS1645 : Change the way the DoB is validated against the min and max age allowed. Now the validation
    '                       is same as that done for UpdateCustomerDetails (ICustomerBO_ValidateCustomerDetails)
        
    ' MO  23/10/2002 BMIDS00639 : Changed the code, as this doesnt calculate ages properly, as it only works on years
    '                               not years, months and dates. START
        
    'If strDoB <> "" Then ' Date of birth exists in XML
    '    Dim dtDOB  As Date
    '    dtDOB = CDate(strDoB)
    '
    '    If (Now < DateAdd("yyyy", intMinimumAge, dtDOB)) Then
    '        ' Too young
    '        objErrAssist.ThrowError TypeName(Me), strFunctionName, 164
    '    End If
    '
    '    If (Now > DateAdd("yyyy", intMaximumAge, dtDOB)) Then
    '        ' Too old
    '        objErrAssist.ThrowError TypeName(Me), strFunctionName, 165
    '    End If
    'End If
    
    'BMIDS00838 Only do age validation when the date of birth is not blank
    'EP903 Use Application Date instead of today's date to verify age.
    If (Len(Trim(strDoB)) > 0) Then
    
        'calculate the applicant's age
        lngAge = Year(strDate) - Year(strDoB)
        
        If Month(strDate) < Month(strDoB) Then
            lngAge = lngAge - 1
        Else
            If Month(strDate) = Month(strDoB) Then
                If Day(strDate) < Day(strDoB) Then
                    lngAge = lngAge - 1
                End If
            End If
        End If
        
        'check the age
        If lngAge < intMinimumAge Then
            ' too young
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 164
        End If
        
        If lngAge > intMaximumAge Then
            ' too old
            ' BM0170 Error number 165 must be used, not 164
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 165
        End If
        ' MO  23/10/2002 BMIDS00639 : END
    End If 'BMIDS00838
    
'    objScriptInterface.RunScript App.Title, TypeName(Me), strFunctionName, vstrXMLRequest
     
    Set objCustVersionElem = Nothing
    Set objXmlAssist = Nothing
    Set objCustomerDO = Nothing
    Set objErrAssist = Nothing
    Set objCustomerElem = Nothing
    Set objGeneralAssist = Nothing
    Set objGlobalParamDO = Nothing
'    Set objScriptInterface = Nothing
   
    Exit Sub
    
ValidateCustomerDetailsVbErr:
    
    Set objCustVersionElem = Nothing
    Set objXmlAssist = Nothing
    Set objCustomerDO = Nothing
    Set objErrAssist = Nothing
    Set objCustomerElem = Nothing
    Set objGeneralAssist = Nothing
    Set objGlobalParamDO = Nothing
'    Set objScriptInterface = Nothing

    '   re-raise error to calling function
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Sub FindCustomerValidate(ByVal vobjXmlDoc As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:  Performs the hard coded validation for FindCustomer
' pass:         vobjXmlRequest  Request xml Document
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomerValidateVbErr

    Const strFunctionName As String = "FindCustomerValidate"

    Dim blnHasPostCode As Boolean
    Dim blnHasSurname As Boolean
    Dim blnHasFirstForename As Boolean
    Dim blnHasDoB As Boolean
    Dim strOtherSysCustNo As String
    Dim strSecondForeName As String
    Dim strOtherForename As String
    Dim strHouseName As String
    Dim strHouseNumber As String
    Dim strFlatNumber As String
    Dim strPhoneNumber As String
    
    Dim objErrAssist As New ErrAssist
    
    ' PSC 09/01/01 SYS1750 - Start
    strOtherSysCustNo = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                   ".//OTHERSYSTEMCUSTOMERNUMBER", _
                                                   False, False)
                                                                                                   
    strSecondForeName = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                   ".//SECONDFORENAME", _
                                                   False, False)
                                                   
    strOtherForename = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                   ".//OTHERFORENAME", _
                                                   False, False)
                                                   
    strHouseName = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                               ".//BUILDINGORHOUSENAME", _
                                               False, False)
                                                   
    strHouseNumber = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                 ".//BUILDINGORHOUSENUMBER", _
                                                 False, False)
                                                   
    strFlatNumber = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                ".//FLATNUMBER", _
                                                False, False)
                                                   
    'JR - SYS2459 (Omiplus24) add Country/Area Code
    strPhoneNumber = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                ".//COUNTRYCODE", _
                                                False, False)
                                                   
    strPhoneNumber = strPhoneNumber & m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                ".//AREACODE", _
                                                False, False)
                
    strPhoneNumber = m_objXmlAssist.GetNodeValue(vobjXmlDoc.documentElement, _
                                                 ".//CUSTOMERTELEPHONENUMBER", _
                                                 False, False)
                                                   
                                                                                                                                                 
        '   Post Code required with any other property details
    If vobjXmlDoc.documentElement.getElementsByTagName("POSTCODE").length > 0 Then
        If Len(vobjXmlDoc.getElementsByTagName("POSTCODE").Item(0).Text) > 0 Then
            blnHasPostCode = True
        End If
    End If
        
    '   Surname, etc. criteria
    If vobjXmlDoc.getElementsByTagName("SURNAME").length > 0 Then
        If Len(vobjXmlDoc.getElementsByTagName("SURNAME").Item(0).Text) > 0 Then
            blnHasSurname = True
        End If
    End If
    
    If vobjXmlDoc.getElementsByTagName("FIRSTFORENAME").length > 0 Then
        If Len(vobjXmlDoc.getElementsByTagName("FIRSTFORENAME").Item(0).Text) > 0 Then
            blnHasFirstForename = True
        End If
    End If
    
    If vobjXmlDoc.getElementsByTagName("DATEOFBIRTH").length > 0 Then
        If Len(vobjXmlDoc.getElementsByTagName("DATEOFBIRTH").Item(0).Text) > 0 Then
            blnHasDoB = True
        End If
    End If
            
    ' Search must be done on customer number alone or on customer details
    If Len(strOtherSysCustNo) = 0 Then

        If (blnHasSurname = True And blnHasFirstForename = False) Or _
           (blnHasSurname = False And blnHasFirstForename = True) Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 549
        End If
        
        If blnHasSurname = False And blnHasDoB = True Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 550
        End If
    Else
        If blnHasSurname = True Or blnHasFirstForename = True Or _
           blnHasDoB = True Or blnHasPostCode = True Or _
           Len(strSecondForeName) > 0 Or Len(strOtherForename) > 0 Or _
           Len(strHouseName) > 0 Or Len(strHouseNumber) > 0 Or _
           Len(strFlatNumber) > 0 Or Len(strPhoneNumber) > 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 302
        End If
    End If
    ' PSC 09/01/01 SYS1750 - End
    
    
    
    Set objErrAssist = Nothing
    
    Exit Sub

FindCustomerValidateVbErr:

    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IsRestrictedByChannel(ByVal vstrUserId As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Determines if a search is to be restricted by channel
' pass:         vstrUserId  User Id
' return:       IsRestrictedByChannel True:  search is to be restricted by channel
'                                     False: search is not to be restricted by channel
'------------------------------------------------------------------------------------------
On Error GoTo IsRestrictedByChannelVbErr

    Const strFunctionName As String = "IsRestrictedByChannel"


    'nturner not needed as it is a private function
    'Dim objContext As ObjectContext
    'Set objContext = GetObjectContext()
    
    Dim objGlobalParamDO As New GlobalParameterDO
'    Dim objOrganisationDO As New OrganisationUserDO
    
    Dim objXmlResponseDoc As FreeThreadedDOMDocument40
    Dim objXmlRequestDoc As FreeThreadedDOMDocument40
    Dim objXMLElement As IXMLDOMElement
    Dim objXmlUserElement As IXMLDOMElement
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim blnRestrictByChannel As Boolean

    blnRestrictByChannel = True
    
    Dim varParam As Variant
    
    ' Check the global parameter to see if searches are restricted by channel
    
    varParam = _
        objGlobalParamDO.GetCurrentParameterByType("SearchByDistributionChannel", _
                                                   "BOOLEAN")
    
    If Not IsEmpty(varParam) Then
        blnRestrictByChannel = varParam
    End If
    
    ' If the SearchByDistributionChannel parameter is false then we need to get the
    ' user's access type to see if this restricts the search by channel
    If blnRestrictByChannel = False Then
    
'       FIXME
'        ' Get user's access type
'        Dim strAccessType As String
'        Dim strResponse As String
'
'        Set objXmlUserElement = objXmlRequestDoc.createElement("USERID")
'        objXmlUserElement.Text = vstrUserId
'        objXmlRequestDoc.appendChild objXmlUserElement
'
'        strResponse = objOrganisationDO.GetData(objXmlRequestDoc.xml)
'        Set objXmlResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'        Set objXMLElement = objXmlResponseDoc.getElementsByTagName("ACCESSTYPE").Item(0)
'        strAccessType = objXMLElement.Text
'
'        blnRestrictByChannel = IsAccessTypeRestricted(strAccessType)
          
    End If
    
    IsRestrictedByChannel = blnRestrictByChannel

    Set objGlobalParamDO = Nothing
'    Set objOrganisationDO = Nothing
    Set objXmlRequestDoc = Nothing
    Set objXMLElement = Nothing
    Set objXmlUserElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
    Exit Function
    
IsRestrictedByChannelVbErr:

    Set objGlobalParamDO = Nothing
'    Set objOrganisationDO = Nothing
    Set objXmlRequestDoc = Nothing
    Set objXMLElement = Nothing
    Set objXmlUserElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ValidateTelephone(xmlNode As IXMLDOMNode)
'-----------------------------------------------------------
'Performs hardcoded validation on telephonenumber
'Pass : xmlNode : CustomerTelephone node
'-----------------------------------------------------------
    Const strFunctionName As String = "ValidateTelephone"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    ' If usage is specified so must both telephone number and contact time
    If (objXmlAssist.GetTagValue(xmlNode, "USAGE") <> "") And _
        (objXmlAssist.GetTagValue(xmlNode, "TELEPHONENUMBER") = "") And _
        (objXmlAssist.GetTagValue(xmlNode, "COUNTRYCODE") = "") And _
        (objXmlAssist.GetTagValue(xmlNode, "AREACODE") = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 128
    End If
    ' If usage is NOT specified then neither telephone number of contact time should be specified
    If (objXmlAssist.GetTagValue(xmlNode, "USAGE") = "") And _
        ((objXmlAssist.GetTagValue(xmlNode, "TELEPHONENUMBER") <> "") And _
        (objXmlAssist.GetTagValue(xmlNode, "COUNTRYCODE") <> "") And _
        (objXmlAssist.GetTagValue(xmlNode, "AREACODE") <> "") Or _
        (objXmlAssist.GetTagValue(xmlNode, "CONTACTTIME") <> "")) Then

        objErrAssist.ThrowError TypeName(Me), strFunctionName, 128
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
End Sub

Private Function IsLockRequiredInThisSystem() As Boolean
'---------------------------------------------------------------
'This checks whether locking is required for this system.
'Return TRUE, if required else FALSE
'--------------------------------------------------------------
On Error GoTo IsLockRequiredInThisSystemVbErr

    Const strFunctionName As String = "IsLockRequiredInThisSystem"


    Dim objGlobalParamDO As GlobalParameterDO
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
                objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim strLockingReqd As String
    strLockingReqd = objGlobalParamDO.GetCurrentParameterByType("Locking", "BOOLEAN")
    
    IsLockRequiredInThisSystem = IIf(strLockingReqd = "0", False, True)
    
    Set objGlobalParamDO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set objErrAssist = Nothing
    Exit Function
IsLockRequiredInThisSystemVbErr:

    objErrAssist.AddToErrSource strFunctionName
    
    Set objGlobalParamDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set objContext = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPersonalDetails(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Retrieves the personal details for the Applicant passed in.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPersonalDetailsVbErr

    Const strFunctionName As String = "GetPersonalDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetPersonalDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPersonalDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetPersonalDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPersonalDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPersonalDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPersonalDetailsExit

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

Public Function SavePersonalDetailsAndAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  BG:01/06/00     Updates Customer details including email and address
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SavePersonalDetailsAndAddressVbErr

    Const strFunctionName As String = "SavePersonalDetailsAndAddress"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlGeneratedKeys As IXMLDOMElement

        ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_Validate(xmlIn.documentElement, bomSavePersonalDetailsAndAddress)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    Set xmlGeneratedKeys = xmlOut.createElement("CUSTOMERADDRESSLIST")
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_SavePersonalDetailsAndAddress(xmlIn.documentElement, xmlGeneratedKeys)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    xmlResponseElem.appendChild xmlGeneratedKeys
    
    SavePersonalDetailsAndAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SavePersonalDetailsAndAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    SavePersonalDetailsAndAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SavePersonalDetailsAndAddressExit
    
End Function

Private Function ICustomerBO_Validate(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal veboMethod As ICustomerBOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ICustomerBO_Validate"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim strCustomerNumber As String
    Dim strContactReason As String
    Dim strContactText As String
    Dim strUserId As String
    Dim strUnitID As String
    Dim objComboDO As ComboDO
    Dim blnIsOther As Boolean
    Dim strOtherTxt As String
    Dim xmlCustNo As IXMLDOMElement
    Dim xmlDateTime As IXMLDOMElement
    Dim xmlStatus As IXMLDOMElement
    Dim blnIsCRSContact As Boolean      'BMIDS00005 BM065
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    Select Case veboMethod
        Case bomCreateContactHistory
              
            If vxmlRequest.nodeName = cstrROOT_CONTACTHISTORY Then
                Set xmlRequestNode = vxmlRequest
            Else
                Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
            End If
            
            If xmlRequestNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, cstrROOT_CONTACTHISTORY & " tag not found"
            End If
            
            'Check that the required fields have values
            strCustomerNumber = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//CUSTOMERNUMBER")
            strContactReason = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//CONTACTREASONCODE")
            strContactText = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//CONTACTTEXT")
            strUserId = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//USERID")
            strUnitID = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//UNITID")
                                
            Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
            
            ' BMIDS00005 BM065 Validation of other reason is not applicable for CRS Contacts (validation type C)
            blnIsCRSContact = objComboDO.IsItemInValidation("ContactHistoryReason", strContactReason, "C")
            If Not blnIsCRSContact Then
            ' BMIDS00005 End
            
                blnIsOther = objComboDO.IsItemInValidation("ContactHistoryReason", strContactReason, "O")
                
                strOtherTxt = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//OTHERREASONTEXT", False, False)
                
                If blnIsOther = True Then
                    If Len(strOtherTxt) = 0 Then
                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 292, "OTHERREASONTEXT must be entered"
                    End If
                Else
                    If Len(strOtherTxt) <> 0 Then
                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 292, "OTHERREASONTEXT must be empty"
                    End If
                End If
            End If  'BMIDS00005

        Case bomUpdateContactHistory
            If vxmlRequest.nodeName = cstrROOT_CONTACTHISTORY Then
                Set xmlRequestNode = vxmlRequest
            Else
                Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_CONTACTHISTORY).Item(0)
            End If
            
            If xmlRequestNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, cstrROOT_CONTACTHISTORY & " tag not found"
            End If
            
            'BMIDS00005 BM065 Edit also saves contact details, reason and other reason, so
            'there should be at least 3 child elements in the node
                
            'If xmlRequestNode.childNodes.length <> 3 Then
            If xmlRequestNode.childNodes.length < 3 Then
            'BMIDS00005 End
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 292
            Else
            
                Set xmlCustNo = m_objXmlAssist.GetNode(xmlRequestNode, ".//CUSTOMERNUMBER")
                Set xmlDateTime = m_objXmlAssist.GetNode(xmlRequestNode, ".//CONTACTHISTORYDATETIME")
                Set xmlStatus = m_objXmlAssist.GetNode(xmlRequestNode, ".//STATUSINDICATOR")
               
            End If
        Case Else
                              
            If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
                Set xmlRequestNode = vxmlRequest
            Else
                ' PSC 22/01/01 SYS1866
                Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
            End If
            If xmlRequestNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
            End If
    End Select
                 
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
        
        Set ICustomerBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function GetPersonalDetailsAndAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets customer address and personal details from CustomerBO.FindCustomerAddressList
'               and CustomerBO.GetPersonalDetails
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPersonalDetailsAndAddressVbErr

    Const strFunctionName As String = "GetPersonalDetailsAndAddress"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetPersonalDetailsAndAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPersonalDetailsAndAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetPersonalDetailsAndAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPersonalDetailsAndAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPersonalDetailsAndAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPersonalDetailsAndAddressExit

End Function

Public Function CreatePropertyInsuranceDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "CreatePropertyInsuranceDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_Validate(xmlIn.documentElement, bomCreatePropertyInsuranceDetails)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreatePropertyInsuranceDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreatePropertyInsuranceDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreatePropertyInsuranceDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CreatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePropertyInsuranceDetailsExit

End Function

Public Function UpdatePropertyInsuranceDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "UpdatePropertyInsuranceDetails"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_Validate(xmlIn.documentElement, bomUpdatePropertyInsuranceDetails)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdatePropertyInsuranceDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdatePropertyInsuranceDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdatePropertyInsuranceDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    UpdatePropertyInsuranceDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdatePropertyInsuranceDetailsExit
    
End Function

Public Function GetPropertyInsuranceDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyInsuranceDetailsVbErr

    Const strFunctionName As String = "GetPropertyInsuranceDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetPropertyInsuranceDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPropertyInsuranceDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetPropertyInsuranceDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPropertyInsuranceDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPropertyInsuranceDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPropertyInsuranceDetailsExit

End Function

Public Function FindContactHistoryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindContactHistoryListVbErr

    Const strFunctionName As String = "FindContactHistoryList"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
       
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
          
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_FindContactHistoryList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
      
    FindContactHistoryList = xmlResponseElem.xml
       
    m_objContext.SetComplete

FindContactHistoryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindContactHistoryListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindContactHistoryList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindContactHistoryListExit

End Function

Public Function GetContactHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetContactHistoryVbErr

    Const strFunctionName As String = "GetContactHistory"
        
        
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
       
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetContactHistory(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetContactHistory = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetContactHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetContactHistoryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetContactHistory = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetContactHistoryExit

End Function

Public Function CreateContactHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactHistoryVbErr

    Const strFunctionName As String = "CreateContactHistory"
    
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_Validate(xmlIn.documentElement, bomCreateContactHistory)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateContactHistory(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateContactHistory = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateContactHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    
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

    CreateContactHistory = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateContactHistoryExit

End Function

Public Function UpdateContactHistory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactHistoryVbErr

    Const strFunctionName As String = "UpdateContactHistory"
    
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_Validate(xmlIn.documentElement, bomUpdateContactHistory)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateContactHistory(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateContactHistory = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateContactHistoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
 
    
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

    UpdateContactHistory = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateContactHistoryExit
    
End Function

Private Function AddUserAndUnitToHistory(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vxmlContactHistory As IXMLDOMElement) As IXMLDOMNode
On Error GoTo AddUserAndUnitToHistoryVbErr
    
    Const strFunctionName As String = "AddUserAndUnitToHistory"
    
      
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlUserList As IXMLDOMElement
    Dim xmlUnitList As IXMLDOMElement
    Dim xmlUserDoc As FreeThreadedDOMDocument40
    Dim xmlUnitDoc As FreeThreadedDOMDocument40
    Dim objOrganisationBO As OrganisationBO
    Dim xmlContHistList As IXMLDOMNodeList
    Dim xmlContHist As IXMLDOMElement
    Dim xmlUserID As IXMLDOMElement
    Dim xmlUnitID As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    
    Dim strUserResponse As String
    Dim strUnitResponse As String
    Dim strUserName As String
    Dim strUnitName As String
    Dim strPattern As String
    Dim xmlContHistDoc As FreeThreadedDOMDocument40
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
          
    Set xmlContHistDoc = New FreeThreadedDOMDocument40
    xmlContHistDoc.validateOnParse = False
    xmlContHistDoc.setProperty "NewParser", True
    Set xmlCopyNode = vxmlContactHistory.cloneNode(True)
    xmlContHistDoc.appendChild xmlCopyNode
              
    'Create new document and append the Request node
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempRequestNode = vxmlRequestNode.cloneNode(True)
    xmlTempDoc.appendChild xmlTempRequestNode
                 
    'Create element for the UserID
    Set xmlUserList = xmlTempDoc.createElement("USERLIST")
    Set xmlUnitList = xmlTempDoc.createElement("UNITLIST")
      
    'Set the xml to the CONTACTHISTORY nodes
    Set xmlContHistList = xmlCopyNode.selectNodes("//CONTACTHISTORY")
               
    For Each xmlContHist In xmlContHistList
        Set xmlUserID = m_objXmlAssist.GetNode(xmlContHist, ".//USERID")
        Set xmlUnitID = m_objXmlAssist.GetNode(xmlContHist, ".//UNITID")
        
        strPattern = "USERID[.=" & """" & xmlUserID.Text & """" & "]"
        Set xmlTempNode = xmlUserList.selectSingleNode(strPattern)
        
        If xmlTempNode Is Nothing Then
            xmlUserList.appendChild xmlUserID
        End If
          
        strPattern = ".//UNITID[.=" & """" & xmlUnitID.Text & """" & "]"
        Set xmlTempNode = xmlUnitList.selectSingleNode(strPattern)
       
        If xmlTempNode Is Nothing Then
            xmlUnitList.appendChild xmlUnitID
        End If
    Next
    
    xmlTempRequestNode.appendChild xmlUserList
    
    Set objOrganisationBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
             
    strUserResponse = objOrganisationBO.FindUserNameList(xmlTempRequestNode.xml)
    Set xmlUserDoc = m_objXmlAssist.load(strUserResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlUserDoc.documentElement, xmlResponseElem, True
   
    xmlTempRequestNode.replaceChild xmlUnitList, xmlUserList
    
    strUnitResponse = objOrganisationBO.FindUnitNameList(xmlTempRequestNode.xml)
    Set xmlUnitDoc = m_objXmlAssist.load(strUnitResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlUnitDoc.documentElement, xmlResponseElem, True
       
    'Set the xml to the CONTACTHISTORY nodes
    Set xmlContHistList = vxmlContactHistory.selectNodes("//CONTACTHISTORY")
                   
    'Loop throught the xmlContHistList and get the FORENAME,SURNAME & UNITNAME text
    'using the UserID and UnitID
    For Each xmlContHist In xmlContHistList
        Set xmlUserID = m_objXmlAssist.GetNode(xmlContHist, ".//USERID")
        Set xmlUnitID = m_objXmlAssist.GetNode(xmlContHist, ".//UNITID")
        
        'Get FORENAME text and assign to strUserName
        strPattern = ".//USER[USERID = " & """" & xmlUserID.Text & """" & "]/USERFORENAME"
        strUserName = m_objXmlAssist.GetNodeValue(xmlUserDoc, strPattern, False)
        
        'Get SURNAME (and concatenate it to strUserName)
        strPattern = ".//USER[USERID = " & """" & xmlUserID.Text & """" & "]/USERSURNAME"
        strUserName = strUserName & " " & m_objXmlAssist.GetNodeValue(xmlUserDoc, strPattern, False)
        
        If strUserName = " " Then
            strUserName = "Unknown"
        End If
        
        'Get UNITNAME and assign it to strUnitName
        strPattern = ".//UNIT[UNITID = " & """" & xmlUnitID.Text & """" & "]/UNITNAME"
        strUnitName = m_objXmlAssist.GetNodeValue(xmlUnitDoc, strPattern, False)
        
        If Len(strUnitName) = 0 Then
            strUnitName = "Unknown"
        End If
       
        'Create USERNAME element
        Set xmlElement = xmlTempDoc.createElement("USERNAME")
        xmlElement.Text = strUserName
        xmlContHist.insertBefore xmlElement, xmlUnitID
        
        'Create UNITNAME element
        Set xmlElement = xmlTempDoc.createElement("UNITNAME")
        xmlElement.Text = strUnitName
        xmlContHist.insertBefore xmlElement, xmlUnitID.nextSibling
     Next
     
     xmlResponseElem.appendChild vxmlContactHistory
   
     Set AddUserAndUnitToHistory = xmlResponseElem

AddUserAndUnitToHistoryExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlCopyNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlUserList = Nothing
    Set xmlUnitList = Nothing
    Set xmlUserDoc = Nothing
    Set xmlUnitDoc = Nothing
    Set objOrganisationBO = Nothing
    Set xmlContHistList = Nothing
    Set xmlContHist = Nothing
    Set xmlUserID = Nothing
    Set xmlUnitID = Nothing
    Set xmlElement = Nothing
    Set xmlTempNode = Nothing
       

    Exit Function

AddUserAndUnitToHistoryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set AddUserAndUnitToHistory = m_objErrAssist.CreateErrorResponseEx()
       
    ' go to clean-up section
    Resume AddUserAndUnitToHistoryExit

End Function

Public Function GetCustomerFromLegacyNumber(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerFromLegacyNumberVbErr

    Const strFunctionName As String = "GetCustomerFromLegacyNumber"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetCustomerFromLegacyNumber(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCustomerFromLegacyNumber = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetCustomerFromLegacyNumberExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetCustomerFromLegacyNumberVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCustomerFromLegacyNumber = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetCustomerFromLegacyNumberExit

End Function

Private Function ConvertCustomerToOmigaFormat(ByVal vxmlCustomer As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Converts a customer from the administration system into the correct format
'               to create a customer in Omiga
' pass:         vxmlCustomer            xml Containing Customer to convert
'
' return:       n/a
'------------------------------------------------------------------------------------------

On Error GoTo ConvertCustomerToOmigaFormatVbErr

    Const cstrFunctionName As String = "ConvertCustomerToOmigaFormat"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlCustomerCopy As IXMLDOMElement
    Dim xmlCustomerVersElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlTableList As IXMLDOMElement
    Dim xmlHomeAddress As IXMLDOMElement
    Dim xmlCorrAddress As IXMLDOMElement
    Dim objComboDO  As ComboDO
    
    Dim strHomeAddressValue As String
    Dim strCorrespAddressValue As String
    Dim strOmigaCustomerNumber As String    'BMIDS00443
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    Set xmlCustomerCopy = vxmlCustomer.cloneNode(True)
    
    ' Set up customer details
    Set xmlTableElem = xmlOut.createElement("CUSTOMER")
    xmlOut.appendChild xmlTableElem
    Set xmlElement = xmlOut.createElement("OTHERSYSTEMCUSTOMERNUMBER")
    xmlElement.Text = m_objXmlAssist.GetNodeValue(xmlCustomerCopy, ".//CUSTOMERNUMBER")
    xmlTableElem.appendChild xmlElement
    
    'BMIDS00443 Admin Customer may have Omiga3 customer number which should be used if available
    strOmigaCustomerNumber = Trim(m_objXmlAssist.GetNodeValue(xmlCustomerCopy, ".//OMIGACUSTOMERNUMBER", False, False))
    If Len(strOmigaCustomerNumber) > 0 Then
        Set xmlElement = xmlOut.createElement("OMIGACUSTOMERNUMBER")
        xmlElement.Text = strOmigaCustomerNumber
        xmlTableElem.appendChild xmlElement
    End If
    'BMIDS00443 End
    
    ' PSC 19/10/2005 MAR57
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//OTHERSYSTEMTYPE", False)
    
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CHANNELID", False)
    
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement
    End If
    
    ' Set up customer version number and add it under CUSTOMER
    Set xmlCustomerVersElem = xmlOut.createElement("CUSTOMERVERSION")
    xmlTableElem.appendChild xmlCustomerVersElem
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CONTACTEMAILADDRESS", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
        
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//DATEOFBIRTH", False)
    
    If Not xmlElement Is Nothing Then
        'BM0308 Dates only containing separator charactors should be treated as blank
        'e.g. "  /  /    " should be converted to ""
        If Len(Replace(xmlElement.Text, " ", vbNullString)) <= 2 Then
            xmlElement.Text = vbNullString
        End If
        'BM0308 End
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
     Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//SURNAME", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
   
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//FIRSTFORENAME", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If

    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//SECONDFORENAME", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//OTHERFORENAMES", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//GENDER", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
   
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//TITLE", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
   
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//MOTHERSMAIDENNAME", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//MARITALSTATUS", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//NATIONALINSURANCENUMBER", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//MEMBEROFSTAFF", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    'KW
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//NATIONALITY", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    'KW End
    
    ' PSC 13/10/2005 MAR57 - Start
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERSTATUS", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//SPECIALNEEDS", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERKYCSTATUS", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERKYCADDRESSFLAG", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERKYCIDFLAG", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERCATEGORY", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//MAILSHOTREQUIRED", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CUSTOMERDORMANTACCOUNTIND", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//PROSPECTPASSWORDTAKEN", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    ' PSC 13/10/2005 MAR57 - End
    
    ' Set up addresses
    Set xmlHomeAddress = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CURRENTADDRESS", False)
    Set xmlCorrAddress = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//CORRESPONDENCEADDRESS", False)
    
    If Not (xmlHomeAddress Is Nothing And xmlCorrAddress Is Nothing) Then
                        
        ' Add CUSTOMERADDRESSLIST under CUSTOMERVERSION
        Set xmlTableList = xmlOut.createElement("CUSTOMERADDRESSLIST")
        xmlCustomerVersElem.appendChild xmlTableList
        
        ' Set up Current Address
        If Not xmlHomeAddress Is Nothing Then
        
            Set xmlTableElem = xmlOut.createElement("CUSTOMERADDRESS")
            xmlTableList.appendChild xmlTableElem
            
            strHomeAddressValue = objComboDO.GetFirstComboValueId("CustomerAddressType", "H")
    
            Set xmlElement = xmlOut.createElement("ADDRESSTYPE")
            xmlElement.Text = strHomeAddressValue
            xmlTableElem.appendChild xmlElement
    
            xmlTableElem.appendChild xmlHomeAddress
            m_objXmlAssist.ChangeNodeName xmlHomeAddress, "CURRENTADDRESS", "ADDRESS"
        End If
            
        ' Set up Correspondence Address
        If Not xmlCorrAddress Is Nothing Then
        
            Set xmlTableElem = xmlOut.createElement("CUSTOMERADDRESS")
            xmlTableList.appendChild xmlTableElem
            
            strCorrespAddressValue = objComboDO.GetFirstComboValueId("CustomerAddressType", "C")
    
            Set xmlElement = xmlOut.createElement("ADDRESSTYPE")
            xmlElement.Text = strCorrespAddressValue
            xmlTableElem.appendChild xmlElement
    
            xmlTableElem.appendChild xmlCorrAddress
            m_objXmlAssist.ChangeNodeName xmlCorrAddress, "CORRESPONDENCEADDRESS", "ADDRESS"
        End If
    End If
    
    ' Set up phone numbers
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//TELEPHONENUMBERLIST", False)
    
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
        m_objXmlAssist.ChangeNodeName xmlElement, "TELEPHONENUMBERLIST", "CUSTOMERTELEPHONENUMBERLIST"
        m_objXmlAssist.ChangeNodeName xmlElement, "TELEPHONENUMBERDETAILS", "CUSTOMERTELEPHONENUMBER"
    End If
    
    ' BMIDS864  Set up customer verification
    Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerCopy, ".//VERIFICATIONLIST", False)
 
    If Not xmlElement Is Nothing Then
        xmlCustomerVersElem.appendChild xmlElement
    End If
    ' BMIDS864  End

    Set ConvertCustomerToOmigaFormat = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set xmlElement = Nothing
    Set xmlCustomerVersElem = Nothing
    Set xmlCustomerCopy = Nothing
    Set xmlTableElem = Nothing
    Set xmlTableList = Nothing
    Set xmlHomeAddress = Nothing
    Set xmlCorrAddress = Nothing
    Set objComboDO = Nothing
    
    
    Exit Function
    
ConvertCustomerToOmigaFormatVbErr:

    
    Set xmlOut = Nothing
    Set xmlElement = Nothing
    Set xmlCustomerVersElem = Nothing
    Set xmlCustomerCopy = Nothing
    Set xmlTableElem = Nothing
    Set xmlTableList = Nothing
    Set xmlHomeAddress = Nothing
    Set xmlCorrAddress = Nothing
    Set objComboDO = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'BMIDS00005 CRWP2 BM065 GetCRSContactData
Private Function ICustomerBO_GetCRSContactData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves contact log data for a customer data from BMids ICBS/CRS system
' pass:         vxmlRequest xml Request node containing criteria for data to be retrieved
' return:       xml Response node containing contacts for a customer
'------------------------------------------------------------------------------------------
On Error GoTo GetCRSContactDataVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetCRSContactData"

    
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlAdminResponse    As FreeThreadedDOMDocument40
    Dim xmlNewContact       As FreeThreadedDOMDocument40
    Dim xmlRequestElem      As IXMLDOMElement
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlAdminRequestElem As IXMLDOMElement
    Dim xmlContactsNodeList As IXMLDOMNodeList
    Dim xmlContactElem      As IXMLDOMElement
    Dim xmlFindContacts     As IXMLDOMElement
    Dim xmlContactHistory   As IXMLDOMElement
    Dim objAdminBO          As AdminInterfaceBO
    Dim strAdminResponse    As String
    Dim strCustomerNumber   As String
    Dim strSortOrder        As String
    Dim strResponse         As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strCustomerNumber = m_objXmlAssist.GetMandatoryAttribute(vxmlRequest.selectSingleNode(cstrROOT_CONTACTLOG), "OMIGACUSTOMERNUMBER")
    strSortOrder = m_objXmlAssist.GetMandatoryAttribute(vxmlRequest.selectSingleNode(cstrROOT_CONTACTLOG), "SORTORDER")

    Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")

    Set xmlAdminRequestElem = vxmlRequest.cloneNode(True)
    xmlAdminRequestElem.setAttribute "OPERATION", "GetCRSContactData"

    ' Query the admin system
    strAdminResponse = objAdminBO.Request(xmlAdminRequestElem.xml)
    errCheckXMLResponse strAdminResponse, True
    
    Set xmlAdminResponse = New FreeThreadedDOMDocument40
    xmlAdminResponse.validateOnParse = False
    xmlAdminResponse.setProperty "NewParser", True
    Set xmlAdminResponse = m_objXmlAssist.load(strAdminResponse, TypeName(Me), strFunctionName)
    
    Set xmlContactsNodeList = xmlAdminResponse.selectNodes("RESPONSE/CONTACTLOG/CONTACT")
    If xmlContactsNodeList.length > 0 Then
        ' Save each contact that was returned by the admin system to the database
        For Each xmlContactElem In xmlContactsNodeList
            Set xmlRequestElem = vxmlRequest.cloneNode(False)
            xmlRequestElem.removeAttribute "OPERATION"
            xmlRequestElem.appendChild xmlContactElem
            Set xmlNewContact = xmlCreateElementRequestFromNode(xmlRequestElem, "CONTACT", True, "CONTACTHISTORY")
            xmlChangeNodeName xmlNewContact.selectSingleNode("REQUEST/CONTACTHISTORY/OMIGACUSTOMERNUMBER"), "OMIGACUSTOMERNUMBER", "CUSTOMERNUMBER"
            xmlChangeNodeName xmlNewContact.selectSingleNode("REQUEST/CONTACTHISTORY/CONTACTREASON"), "CONTACTREASON", "CONTACTREASONCODE"
            xmlChangeNodeName xmlNewContact.selectSingleNode("REQUEST/CONTACTHISTORY/CONTACTDETAILS"), "CONTACTDETAILS", "CONTACTTEXT"
            xmlChangeNodeName xmlNewContact.selectSingleNode("REQUEST/CONTACTHISTORY/OTHERREASON"), "OTHERREASON", "OTHERREASONTEXT"
            xmlChangeNodeName xmlNewContact.selectSingleNode("REQUEST/CONTACTHISTORY/DATE"), "DATE", "CONTACTHISTORYDATETIME"
            CreateContactHistory xmlNewContact.xml
        Next
    Else
        ' Return error message oeCustNoContactsFound if no contacts were found
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeCustNoContactsFound
    End If
    
    'Create a REQUEST for calling FindContactHistoryList
    Set xmlFindContacts = vxmlRequest.cloneNode(False)
    Set xmlContactHistory = xmlOut.createElement("CONTACTHISTORY")
    
    Set xmlContactElem = xmlOut.createElement("CUSTOMERNUMBER")
    xmlContactElem.Text = strCustomerNumber
    xmlContactHistory.appendChild xmlContactElem
    
    Set xmlContactElem = xmlOut.createElement("SORTORDER")
    xmlContactElem.Text = strSortOrder
    xmlContactHistory.appendChild xmlContactElem
    
    xmlFindContacts.appendChild xmlContactHistory
    
    strResponse = FindContactHistoryList(xmlFindContacts.xml)
    Set xmlOut = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
    Set ICustomerBO_GetCRSContactData = xmlOut.documentElement
    
    m_objContext.SetComplete
    
GetCRSContactDataExit:
    
    Set xmlOut = Nothing
    Set xmlAdminResponse = Nothing
    Set xmlNewContact = Nothing
    Set xmlRequestElem = Nothing
    Set xmlResponseElem = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlContactsNodeList = Nothing
    Set xmlContactElem = Nothing
    Set xmlFindContacts = Nothing
    Set xmlContactHistory = Nothing
    Set objAdminBO = Nothing
        

    Exit Function

GetCRSContactDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetCRSContactData = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetCRSContactDataExit

End Function
'BMIDS00005 End

'BMIDS00005 CRWP2 BM065 UpdateCRSContactLog
Private Function ICustomerBO_UpdateCRSContactLog(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates the BMids ICBS/CRS system with contact log data created in Omiga
' pass:         vxmlRequest xml Request node containing application number to update contacts for
' return:       xml Response node containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCRSContactLogVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCRSContactLog"

    
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlContact          As IXMLDOMElement
    Dim xmlAdminRequest     As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlContacts         As IXMLDOMNode
    Dim objICustomerDO      As ICustomerDO
    Dim objAdminBO          As AdminInterfaceBO
    Dim strAdminResponse    As String
    Dim strUserId           As String
    Dim strUnitID           As String
    Dim blnContactsFound    As Boolean          ' BMIDS00325
    Dim objICustomerTxBO    As ICustomerTxBO    ' PSC 21/10/2002 BMIDS00030
    Dim xmlCustomerNo       As IXMLDOMNode      ' PSC 21/10/2002 BMIDS00030
    Dim xmlContactDateTime  As IXMLDOMNode      ' PSC 21/10/2002 BMIDS00030
    Dim xmlRoot             As IXMLDOMNode      ' PSC 21/10/2002 BMIDS00030
    Dim xmlNode             As IXMLDOMNode      ' PSC 21/10/2002 BMIDS00030
    
    ' PSC 21/10/2002 BMIDS00030
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
      
    'Query the database to find contacts for the customer(s) on the application
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    Set xmlTempResponseNode = objICustomerDO.FindContactsPerApplication(vxmlRequest)
    
    'BMIDS00325 Handle no xml returned
    If xmlTempResponseNode Is Nothing Then
        blnContactsFound = False
    'BMIDS00325 End
    'BMIDS00060 Added check to see if any contacts have been found
    ElseIf xmlTempResponseNode.selectNodes("CONTACT").length = 0 Then
    'BMIDS00060 End
        blnContactsFound = False    'BMIDS00325
    Else
        blnContactsFound = True     'BMIDS00325
        ' Convert the output to the correct format for the admin system
        strUserId = xmlGetAttributeText(vxmlRequest, "USERID")
        strUnitID = xmlGetAttributeText(vxmlRequest, "UNITID")
        Set xmlContacts = xmlCreateAttributeBasedResponse(xmlTempResponseNode, True)
        For Each xmlContact In xmlContacts.selectNodes("CONTACT")
            xmlContact.setAttribute "USERID", strUserId
            xmlContact.setAttribute "UNITID", strUnitID
        Next
        Set xmlAdminRequest = vxmlRequest.cloneNode(False)
        xmlAdminRequest.setAttribute "OPERATION", "UpdateCRSContactLog"
        xmlAdminRequest.appendChild xmlContacts
        
        'Update the admin system with the contacts
        Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        strAdminResponse = objAdminBO.Request(xmlAdminRequest.xml)
        
        errCheckXMLResponse strAdminResponse, True
        
        ' PSC 21/10/2002 BMIDS00030 - Start
        ' Set up generic request to update the contacts to show they
        ' have been interfaced
        Set xmlRoot = xmlOut.createElement("CONTACTHISTORY")
        Set xmlCustomerNo = xmlOut.createElement("CUSTOMERNUMBER")
        xmlRoot.appendChild xmlCustomerNo
        Set xmlContactDateTime = xmlOut.createElement("CONTACTHISTORYDATETIME")
        xmlRoot.appendChild xmlContactDateTime
        Set xmlNode = xmlOut.createElement("INTERFACEDTOEXTERNALSYSTEM")
        xmlNode.Text = "1"
        xmlRoot.appendChild xmlNode
        
        Set objICustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
        
        For Each xmlContact In xmlContacts.selectNodes("CONTACT")
            xmlCustomerNo.Text = xmlGetMandatoryAttributeText(xmlContact, "OMIGACUSTOMERNUMBER")
            xmlContactDateTime.Text = xmlGetMandatoryAttributeText(xmlContact, "CONTACTDATE")
        
            Set xmlNode = objICustomerTxBO.UpdateContactHistory(xmlRoot)
            errCheckXMLResponseNode xmlNode, , True
        Next
        
        xmlOut.appendChild xmlNode
        ' PSC 21/10/2002 BMIDS00030 - End
    End If
    
    'BMIDS00060 dont call the admin system when there are no contacts to update
    If Not blnContactsFound Then    'BMIDS00325
        ' There were no contacts to update so return success
        Set xmlContact = xmlOut.createElement("RESPONSE")
        xmlContact.setAttribute "TYPE", "SUCCESS"
        xmlOut.appendChild xmlContact
    End If
    'BMIDS00060
    
    Set ICustomerBO_UpdateCRSContactLog = xmlOut.documentElement
    
    m_objContext.SetComplete
    
UpdateCRSContactLogExit:

    Set xmlOut = Nothing
    Set xmlContact = Nothing
    Set xmlAdminRequest = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlContacts = Nothing
    Set objICustomerDO = Nothing
    Set objAdminBO = Nothing
    Set objICustomerTxBO = Nothing      ' PSC 21/10/2002 BMIDS00030
    Set xmlCustomerNo = Nothing         ' PSC 21/10/2002 BMIDS00030
    Set xmlContactDateTime = Nothing    ' PSC 21/10/2002 BMIDS00030
    Set xmlRoot = Nothing               ' PSC 21/10/2002 BMIDS00030

    
    Exit Function

UpdateCRSContactLogVbErr:

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCRSContactLog = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateCRSContactLogExit

End Function
'BMIDS00005 End

'BMIDS00005 CRWP2 BM065 GetCRSContactData
Public Function GetCRSContactData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calls ICustomerBO_GetCRSContactData to get CRS contact data from the admin system
' pass:         vstrXmlRequest  xml Request data stream containing criteria for data to be retrieved
' return:                       xml Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo GetCRSContactDataVbErr

    Const strFunctionName As String = "GetCRSContactData"
        
        
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetCRSContactData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCRSContactData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetCRSContactDataExit:

    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetCRSContactDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCRSContactData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetCRSContactDataExit

End Function
'BMIDS00005 End

'BMIDS00005 CRWP2 BM065 UpdateCRSContactLog
Public Function UpdateCRSContactLog(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calls ICustomerBO_UpdateCRSContactLog to send updated contact logs to the admin system
' pass:         vstrXmlRequest  xml Request data string containing the application number to update contacts for
' return:                       xml Response data string containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCRSContactLogVbErr

    Const strFunctionName As String = "UpdateCRSContactLog"
    
        
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                                  
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCRSContactLog(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateCRSContactLog = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateCRSContactLogExit:

    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

UpdateCRSContactLogVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCRSContactLog = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateCRSContactLogExit
    
End Function
'BMIDS00005 End

'BMIDS00026 IWP2 BM060 UpdateCRSCustomer
Public Function UpdateCRSCustomer(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calls ICustomerBO_UpdateCRSCustomer to send customer details to the admin system
' pass:         vstrXmlRequest  xml Request data stream containing the application number to update customer(s) details for
' return:                       xml Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCRSCustomerVbErr

    Const strFunctionName As String = "UpdateCRSCustomer"
    
        
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                                  
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_UpdateCRSCustomer(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateCRSCustomer = xmlResponseElem.xml
    
    m_objContext.SetComplete

UpdateCRSCustomerExit:

    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

UpdateCRSCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCRSCustomer = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateCRSCustomerExit
    
End Function
'BMIDS00026 End

Private Function ICustomerBO_UpdateCRSCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates the BMids ICBS/CRS system with customer details created in Omiga
' pass:         vxmlRequest xml Request node containing application number to update contacts for
' return:       xml Response node containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCRSCustomerVbErr
    
    Const strFunctionName As String = "ICustomerBO_UpdateCRSCustomer"

    
    Dim xmlOut                      As FreeThreadedDOMDocument40
    Dim xmlCustomer                 As IXMLDOMElement
    Dim xmlAdminRequest             As IXMLDOMElement
    Dim xmlTempResponseNode         As IXMLDOMNode
    Dim xmlCustomers                As IXMLDOMNode
    Dim objICustomerDO              As ICustomerDO
    Dim objAdminBO                  As AdminInterfaceBO
    Dim strAdminResponse            As String
    Dim strShortName                As String
    Dim strInitials                 As String
    Dim strSurname                  As String
    Dim strPosition                 As String
    Dim strCustomerRole             As String
    Dim intNumberOfApplicants       As Integer
    Dim intCustomerSequenceNumber   As Integer
    Dim xmlTempNode                 As IXMLDOMElement
    Dim strDateMovedIn              As String
    
    ' PSC 16/11/2005 MAR590 - Start
    Dim xmlAdminCustomers As IXMLDOMNodeList
    Dim strMessageType As String
    ' PSC 16/11/2005 MAR590 - End
    
    'PSC 06/10/2005 MAR57 - Start
    Dim xmlCustomerNo As IXMLDOMNode
    Dim xmlCustUpdateRequest As IXMLDOMNode
    Dim xmlCustUpdateResponse As IXMLDOMNode
    Dim xmlCustUpdate As IXMLDOMNode
    Dim objCustTxBO As ICustomerTxBO
    Dim strPassword As String

    Set xmlCustomerNo = vxmlRequest.selectSingleNode("CUSTOMERS/CUSTOMER/CUSTOMERNUMBER")
    strPassword = xmlGetNodeText(vxmlRequest, "CUSTOMERS/PASSWORD")
    'PSC 06/10/2005 MAR57 - End
    
    'Query the database to find all customers on the application
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
    
    'PSC 06/10/2005 MAR57 - Start
    If xmlCustomerNo Is Nothing Then
        Set xmlTempResponseNode = objICustomerDO.FindCustomersPerApplication(vxmlRequest)
        Set xmlCustomers = xmlCreateAttributeBasedResponse(xmlTempResponseNode, True)
    Else
        Set xmlCustomers = objICustomerDO.FindCustomerCRSData(vxmlRequest)
    End If
    'PSC 06/10/2005 MAR57 - End
        
    ' Convert the output to the correct format for the admin system
    For Each xmlCustomer In xmlCustomers.selectNodes("CUSTOMER")
    
        ' derive and add the short name
        strSurname = xmlGetAttributeText(xmlCustomer, "SURNAME")
        strInitials = xmlGetAttributeText(xmlCustomer, "INITIALS")
        strShortName = DeriveShortName(strSurname, strInitials)
        
        xmlSetAttributeValue xmlCustomer, "SHORTNAME", strShortName
        
        ' derive and add the position within the application
        'PSC 06/10/2005 MAR57 - Start
        If Len(strPassword) > 0 Then
            xmlSetAttributeValue xmlCustomer, "PASSWORD", strPassword
        End If
        
        If xmlCustomerNo Is Nothing Then
            intNumberOfApplicants = xmlGetAttributeAsInteger(xmlCustomer, "NUMBEROFAPPLICANTS")
            strCustomerRole = xmlGetAttributeText(xmlCustomer, "CUSTOMERROLE")
            intCustomerSequenceNumber = xmlGetAttributeAsInteger(xmlCustomer, "CUSTOMERSEQUENCENUMBER")
            strPosition = DerivePositionWithinApplication(intNumberOfApplicants, strCustomerRole, intCustomerSequenceNumber)
            
            xmlSetAttributeValue xmlCustomer, "POSITIONWITHINAPPLICATION", strPosition
        End If
        'PSC 06/10/2005 MAR57 - End
        
        'Remove attributes that do not need to be passed to the admin system
        xmlCustomer.removeAttribute "CUSTOMERVERSIONNUMBER"
        xmlCustomer.removeAttribute "CUSTOMERROLE"
        xmlCustomer.removeAttribute "CUSTOMERSEQUENCENUMBER"
        xmlCustomer.removeAttribute "NUMBEROFAPPLICANTS"
        
        Set xmlTempNode = xmlCustomer.selectSingleNode(".//ADDRESS")
        
        strDateMovedIn = m_objXmlAssist.GetAttributeFromNode(xmlTempNode, "DATEMOVEDIN")
        If strDateMovedIn = "" Then
            xmlTempNode.setAttribute "DATEMOVEDIN", Now()
        End If
    Next
    
    Set xmlAdminRequest = vxmlRequest.cloneNode(False)
    xmlAdminRequest.setAttribute "OPERATION", "UpdateCRSCustomer"
    xmlAdminRequest.appendChild xmlCustomers
    
    'Update the admin system with the customer details
    Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    strAdminResponse = objAdminBO.Request(xmlAdminRequest.xml)
    errCheckXMLResponse strAdminResponse, True
    
    Set xmlOut = m_objXmlAssist.load(strAdminResponse, TypeName(Me), strFunctionName)

    ' PSC 16/11/2005 MAR590 - Start
    strMessageType = xmlGetAttributeText(vxmlRequest, "MessageType")
    
    If strMessageType = "CreateCustomer" Then
        Set xmlCustUpdateRequest = vxmlRequest.cloneNode(False)

        Set xmlAdminCustomers = xmlOut.selectNodes("RESPONSE/CUSTOMERLIST/CUSTOMER")
        
        For Each xmlCustomer In xmlAdminCustomers
            Set xmlCustUpdate = xmlOut.createElement("CUSTOMER")
            xmlCustUpdateRequest.appendChild xmlCustUpdate
            Set xmlTempNode = xmlOut.createElement("CUSTOMERNUMBER")
            xmlTempNode.Text = xmlGetAttributeText(xmlCustomer, "OMIGACUSTOMERNUMBER")
            xmlCustUpdate.appendChild xmlTempNode
            Set xmlTempNode = xmlOut.createElement("OTHERSYSTEMCUSTOMERNUMBER")
            xmlTempNode.Text = xmlGetAttributeText(xmlCustomer, "OTHERSYSTEMCUSTOMERNUMBER")
            xmlCustUpdate.appendChild xmlTempNode
        Next
    
        Set objCustTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
        Set xmlCustUpdateResponse = objCustTxBO.UpdateCustomer(xmlCustUpdateRequest)
        errCheckXMLResponseNode xmlCustUpdateResponse, , True
    End If
    ' PSC 16/11/2005 MAR590 - end
    
    Set ICustomerBO_UpdateCRSCustomer = xmlOut.documentElement

    m_objContext.SetComplete
    
UpdateCRSCustomerExit:

    Set xmlOut = Nothing
    Set xmlCustomer = Nothing
    Set xmlAdminRequest = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCustomers = Nothing
    Set objICustomerDO = Nothing
    Set objAdminBO = Nothing

    'PSC 06/10/2005 MAR57 - Start
    Set xmlCustomerNo = Nothing
    Set xmlCustUpdateRequest = Nothing
    Set xmlCustUpdateResponse = Nothing
    Set xmlCustUpdate = Nothing
    Set objCustTxBO = Nothing
    'PSC 06/10/2005 MAR57 - End
    
    Set xmlAdminCustomers = Nothing         ' PSC 16/11/2005 MAR590

    
    Exit Function

UpdateCRSCustomerVbErr:

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_UpdateCRSCustomer = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateCRSCustomerExit

End Function
'BMIDS00026 IWP2 BM060
Private Function DeriveShortName(ByVal vstrSurname As String, ByVal vstrInitials As String) As String
' header ----------------------------------------------------------------------------------
' description:  Derives a shortened version of customer's surname and initials
'               The short name is the 1st 18 characters of the surname followed by a space
'               and the 1st initial, followed by a space and the 2nd initial
' pass:         vstrSurname string containing the cutomer's surname
'               vstrInitials string containing the customer's initials
' return:       string containing the derived short name
'------------------------------------------------------------------------------------------
    Dim strShortName    As String
    Dim strInitial      As String
    
    strShortName = Left(Trim(vstrSurname), 18)

    ' Add first initial if it exists
    strInitial = Left(Trim(vstrInitials), 1)
    If Len(strInitial) > 0 Then
        strShortName = strShortName & " " & strInitial
        ' Add second initial if it exists
        strInitial = Mid(Trim(vstrInitials), 2, 1)
        If Len(strInitial) > 0 Then
            strShortName = strShortName & " " & strInitial
        End If
    End If
    
    DeriveShortName = strShortName
End Function
'BMIDS00026 End

'BMIDS00026 IWP2 BM060
Private Function DerivePositionWithinApplication(ByVal vintNumberOfApplicants, ByVal vstrCustomerRole As String, ByVal vintCustomerSequenceNumber As Integer) As String
' header ----------------------------------------------------------------------------------
' description:  Derives a customer's position within the application based on Omiga 3 rules
' pass:         vintNumberOfApplicants integer containing the number of applicants on an application
'               vstrCustomeRole string containing the customer's role on an application. Either "A" or "G"
'               vintCustomerSequenceNumber integer containing the customer sequence number
' return:       string containing the derived position within the application
'------------------------------------------------------------------------------------------
    Dim strPosition As String

    'BMIDS00252 Added extra condition missing from spec
    'If vintNumberOfApplicants = 1 Then
    If (vintNumberOfApplicants = 1) And (vstrCustomerRole <> "G") Then
        strPosition = "SOW"
    Else
        If vstrCustomerRole = "G" Then
            strPosition = "IND"
        Else
            If vintCustomerSequenceNumber = 1 Then
                strPosition = "JAF"
            ElseIf vintCustomerSequenceNumber = 2 Then
                strPosition = "JA2"
            Else
                strPosition = "JAO"
            End If
        End If
    End If
    
    DerivePositionWithinApplication = strPosition
End Function
'BMIDS00026 End

'MO     21/08/2002  BMIDS00218 - Start
Public Function GetNextCustomerNumber() As String

On Error GoTo GetNextCustomerNumberVbErr

    Const strFunctionName As String = "GetNextCustomerNumber"

        
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetNextCustomerNumber()
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetNextCustomerNumber = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetNextCustomerNumberExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNextCustomerNumberVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNextCustomerNumber = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetNextCustomerNumberExit

End Function

Private Function ICustomerBO_GetNextCustomerNumber() As IXMLDOMNode

On Error GoTo GetNextCustomerNumberVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetNextCustomerNumber"


    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerNTxBO As ICustomerNTxBO
    Dim lngNewCustomerNumber As Long
    Dim xmlCustomerElem As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objICustomerNTxBO = New CustomerNTxBO
    Else
        Set objICustomerNTxBO = m_objContext.CreateInstance(App.Title & ".CustomerNTxBO")
    End If
                    
    'get the next customer number
    lngNewCustomerNumber = objICustomerNTxBO.GetNextCustomerNumber()
    
    'create a customer node to return the new number in
    Set xmlCustomerElem = xmlOut.createElement("CUSTOMER")
    
    'set the CUSTOMERNUMBER attribute to the new number
    xmlSetAttributeValue xmlCustomerElem, "CUSTOMERNUMBER", lngNewCustomerNumber
    
    xmlResponseElem.appendChild xmlCustomerElem
        
    Set ICustomerBO_GetNextCustomerNumber = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetNextCustomerNumberExit:
    
    Set objICustomerNTxBO = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNextCustomerNumberVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetNextCustomerNumber = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetNextCustomerNumberExit
    
End Function
'MO     21/08/2002  BMIDS00218 - End

'BMIDS00443 Split into a separate function as GetAndSynchroniseCustomerDetails was too long
Private Function GetAdminCustomer(ByVal vxmlRequest As IXMLDOMNode, Optional ByVal vstrOtherSysCustNo As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Download a customer from the admin system using the OtherSystemCustomerNumber
' pass:         vxmlRequest         xml Request element
'               vstrOtherSysCustNo  String containing Admin system customer number
' return:       robjResponseNode is passed by reference
'------------------------------------------------------------------------------------------
    Const cstrFunctionName  As String = "GetAdminCustomer"
    Dim xmlAdminRequestDoc  As FreeThreadedDOMDocument40
    Dim xmlAdminResponseDoc As FreeThreadedDOMDocument40
    Dim xmlAdminRequestElem As IXMLDOMElement
    Dim xmlCustomerElem     As IXMLDOMElement
    Dim xmlElementBasedNode As IXMLDOMNode
    Dim xmlTempConvertNode  As IXMLDOMNode
    Dim objAdminBO          As Object
    Dim strAdminResponse    As String
    ' PSC 11/10/2005 MAR57 - Start
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerIn As IXMLDOMNode
    Dim xmlCustomers As IXMLDOMNode
    Dim xmlCustomerOut As IXMLDOMNode
    ' PSC 11/10/2005 MAR57 - End

    
On Error GoTo GetAdminCustomerVbErr
    
    Set xmlAdminRequestDoc = New FreeThreadedDOMDocument40
    xmlAdminRequestDoc.validateOnParse = False
    xmlAdminRequestDoc.setProperty "NewParser", True
    
    'Create Request element
    Set xmlAdminRequestElem = vxmlRequest.cloneNode(False)
    xmlAdminRequestDoc.appendChild xmlAdminRequestElem
    xmlAdminRequestElem.setAttribute "OPERATION", "GetCustomerDetails"
    
    ' PSC 11/10/2005 MAR57 - Start
    If Len(vstrOtherSysCustNo) > 0 Then
        'Create Customer element
        Set xmlCustomerElem = xmlAdminRequestDoc.createElement("CUSTOMER")
        xmlAdminRequestElem.appendChild xmlCustomerElem
        xmlCustomerElem.setAttribute "CUSTOMERNUMBER", vstrOtherSysCustNo
    Else
        Set xmlCustomerList = vxmlRequest.selectNodes(".//CUSTOMER")
        For Each xmlCustomerIn In xmlCustomerList
            Set xmlCustomerElem = xmlAdminRequestDoc.createElement("CUSTOMER")
            xmlAdminRequestElem.appendChild xmlCustomerElem
            xmlCustomerElem.setAttribute "CUSTOMERNUMBER", xmlGetMandatoryNodeText(xmlCustomerIn, "OTHERSYSTEMCUSTOMERNUMBER")
        Next
    End If
    ' PSC 11/10/2005 MAR57 - End
    
    'Create an instance of AdminBO and call Request method
    Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    strAdminResponse = objAdminBO.Request(xmlAdminRequestDoc.xml)
    
    m_objErrAssist.CheckXMLResponse strAdminResponse, True
    Set xmlAdminResponseDoc = m_objXmlAssist.load(strAdminResponse, TypeName(Me), cstrFunctionName)
    'm_objErrAssist.CheckResponse xmlAdminResponseDoc.documentElement, vobjResponseNode, True
               
    'Convert attribute based node to element based
    Set xmlElementBasedNode = xmlMakeNodeElementBased(xmlAdminResponseDoc.documentElement, True, "")

    ' PSC 11/10/2005 MAR57 - Start
    ' Other system customer number is empty so customers have been passed in the request
    If Len(vstrOtherSysCustNo) = 0 Then
        Set xmlCustomers = xmlAdminRequestDoc.createElement("CUSTOMERS")
        
        Set xmlCustomerList = xmlElementBasedNode.selectNodes("CUSTOMER")
        For Each xmlCustomerOut In xmlCustomerList
            Set xmlTempConvertNode = ConvertCustomerToOmigaFormat(xmlCustomerOut)
            xmlCustomers.appendChild xmlTempConvertNode
       Next
        Set GetAdminCustomer = xmlCustomers
    Else
        'Convert XML to normal Omiga format in order to pass it to CreateCustomerDetails
        Set xmlTempConvertNode = ConvertCustomerToOmigaFormat(xmlElementBasedNode)
        Set GetAdminCustomer = xmlTempConvertNode
    End If
    ' PSC 11/10/2005 MAR57 - End
    
GetAdminCustomerVbErr:
    
    Set xmlAdminRequestDoc = Nothing
    Set xmlAdminResponseDoc = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlElementBasedNode = Nothing
    Set xmlTempConvertNode = Nothing
    Set objAdminBO = Nothing
    ' PSC 11/10/2005 MAR57 - Start
    Set xmlCustomerList = Nothing
    Set xmlCustomerIn = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomerOut = Nothing
    ' PSC 11/10/2005 MAR57 - End
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
'BMIDS00443 End

'BMIDS00443 Split into a separate sub as GetAndSynchroniseCustomerDetails was too long
Private Sub UpdateCustomerWithAdminDetails(ByVal vobjRequestNode As IXMLDOMNode, ByRef robjResponseNode As IXMLDOMNode, ByVal vxmlAdminNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Use details of a customer from the admin system to update the Omiga version
'               of the same customer.
' pass:         vobjRequestNode     xml Request element
'               robjResponseNode    xml Response element
'               vxmlAdminNode       xml node containing Admin customer details
'
' return:       robjResponseNode is passed by reference
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName              As String = "UpdateCustomerWithAdminDetails"
    Dim xmlTempDoc                      As FreeThreadedDOMDocument40
    Dim xmlTempReqNode                  As IXMLDOMNode
    Dim xmlTempNode                     As IXMLDOMNode
    Dim xmlAdminParentNode              As IXMLDOMNode
    Dim xmlOmigaParentNode              As IXMLDOMNode
    Dim xmlCustomerNumberNode           As IXMLDOMNode
    Dim xmlCustomerVersionNumberNode    As IXMLDOMNode
    Dim xmlAdminChildNode               As IXMLDOMNode
    Dim xmlOmigaChildNode               As IXMLDOMNode
    Dim xmlSeqNoNode                    As IXMLDOMNode
    Dim xmlAddressGuid                  As IXMLDOMNode
    Dim objICustomerTxBO                As ICustomerTxBO
    Dim xmlAttrib                       As IXMLDOMNode
    Dim strPattern                      As String
    Dim strUpdate                       As String
    Dim i                               As Integer

'    'Create a customer lock and check if any errors occurred
'On Error Resume Next
'    ICustomerBO_CreateCustomerLock vobjRequestNode, robjResponseNode
'
'    If Err.Number <> 0 Then
'        If Err.Number = omiga4DuplicateKey Then
'            Set xmlAttrib = robjResponseNode.ownerDocument.createAttribute("READONLY")
'            xmlAttrib.Text = "1"
'            robjResponseNode.Attributes.setNamedItem xmlAttrib
'        Else
'On Error GoTo UpdateCustomerWithAdminDetailsVBErr
'            Err.Raise Err.Number, Err.Source, Err.Description
'        End If
'    End If

On Error GoTo UpdateCustomerWithAdminDetailsVBErr
    ICustomerBO_GetCustomerDetails vobjRequestNode, robjResponseNode

    'Set up the Request to pass into CustomerBO.UpdatePersonalDetails
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempReqNode = m_objXmlAssist.GetRequestNodeEx(vobjRequestNode)
    xmlTempDoc.appendChild xmlTempReqNode
    Set xmlTempNode = robjResponseNode.selectSingleNode(".//CUSTOMER")
    xmlTempReqNode.appendChild xmlTempNode.cloneNode(True)

    Set xmlTempNode = xmlTempDoc.createElement("ADDRESSGUID")

    'Point to CUSTOMER nodes
    strPattern = ".//CUSTOMER"
    Set xmlAdminParentNode = vxmlAdminNode
    Set xmlOmigaParentNode = m_objXmlAssist.GetNode(xmlTempReqNode, strPattern)
    Set xmlCustomerNumberNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, "CUSTOMERNUMBER")

    ' PSC 29/11/2005 MAR727 - Start
    strPattern = "OTHERSYSTEMCUSTOMERNUMBER"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = "OTHERSYSTEMTYPE"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    'BMIDS00443 BMids admin system does not return ChannelID
    'strPattern = "CHANNELID"
    'Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern)
    'Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
    'xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode

    'Point to CUSTOMERVERSION node
    strPattern = ".//CUSTOMERVERSION"
    Set xmlAdminParentNode = m_objXmlAssist.GetNode(vxmlAdminNode, strPattern, False)
    Set xmlOmigaParentNode = m_objXmlAssist.GetNode(xmlTempReqNode, strPattern)
    Set xmlCustomerVersionNumberNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, "CUSTOMERVERSIONNUMBER")
    
    strPattern = ".//CONTACTEMAILADDRESS"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//DATEOFBIRTH"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//SURNAME"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//FIRSTFORENAME"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//SECONDFORENAME"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//GENDER"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//TITLE"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//MOTHERSMAIDENNAME"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//MARITALSTATUS"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//NATIONALINSURANCENUMBER"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    strPattern = ".//MEMBEROFSTAFF"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If

    'KW
    strPattern = ".//NATIONALITY"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
   'KW End
   ' PSC 29/11/2005 MAR727 - End
   
    ' PSC 17/02/2006 MAR1211 - Start
    strPattern = ".//CUSTOMERSTATUS"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
       
    strPattern = ".//SPECIALNEEDS"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
       
    strPattern = ".//CUSTOMERKYCSTATUS"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//CUSTOMERKYCADDRESSFLAG"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//CUSTOMERKYCIDFLAG"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//CUSTOMERCATEGORY"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//MAILSHOTREQUIRED"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//CUSTOMERDORMANTACCOUNTIND"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    
    strPattern = ".//PROSPECTPASSWORDTAKEN"
    Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
    
    If Not xmlAdminChildNode Is Nothing Then
        If Len(xmlAdminChildNode.Text) > 0 Then
            Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern)
            xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
        End If
    End If
    ' PSC 17/02/2006 MAR1211 - End
    
    'Point to CUSTOMERADDRESSLIST node
    strPattern = ".//CUSTOMERADDRESSLIST"
    
    'BMIDS00792 CUSTOMERADDRESSLIST node may not exist in Admin XML
    Set xmlAdminParentNode = m_objXmlAssist.GetNode(vxmlAdminNode, strPattern, False)
    Set xmlOmigaParentNode = m_objXmlAssist.GetNode(xmlTempReqNode, strPattern)

    ' BMIDS00443 Remove Addresses that are totally blank
    DeleteBlankAddressXML xmlAdminParentNode

    For i = 1 To 2  'BMIDS00443 The block below was duplicated twice, so it is now inside a loop
        strPattern = ".//CUSTOMERADDRESS[ADDRESSTYPE = '" & CStr(i) & "']"
        'BMIDS00792
        If Not xmlAdminParentNode Is Nothing Then
            Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
        Else
            Set xmlAdminChildNode = Nothing
        End If
        'BMIDS00792 End
        Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern, False)
    
        If Not xmlOmigaChildNode Is Nothing Then
            If Not xmlAdminChildNode Is Nothing Then
                Set xmlSeqNoNode = m_objXmlAssist.GetNode(xmlOmigaChildNode, "CUSTOMERADDRESSSEQUENCENUMBER")
                Set xmlAddressGuid = m_objXmlAssist.GetNode(xmlOmigaChildNode, "ADDRESSGUID")
                xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
                Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, ".//CUSTOMERADDRESS[ADDRESSTYPE = '" & CStr(i) & "']")
    
                'append Customernumber and customerversionnumber tags
                xmlOmigaChildNode.appendChild xmlCustomerVersionNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlCustomerNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlSeqNoNode.cloneNode(True)
    
                Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, ".//CUSTOMERADDRESS[ADDRESSTYPE = '" & CStr(i) & "']/ADDRESS")
                xmlOmigaChildNode.appendChild xmlAddressGuid.cloneNode(True)
            End If
        Else
            If Not xmlAdminChildNode Is Nothing Then
                xmlOmigaParentNode.appendChild xmlAdminChildNode.cloneNode(True)
                Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, ".//CUSTOMERADDRESS[ADDRESSTYPE = '" & CStr(i) & "']")
    
                'append Customernumber and customerversionnumber tags
                xmlOmigaChildNode.appendChild xmlCustomerVersionNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlCustomerNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlTempNode.cloneNode(True)
            End If
        End If
    Next 'BMIDS00443

    'Point to CUSTOMERTELEPHONENUMBERLIST node
    strPattern = ".//CUSTOMERTELEPHONENUMBERLIST"
    
    'BMIDS00792 CUSTOMERTELEPHONENUMBERLIST node may not exist in Admin XML
    Set xmlAdminParentNode = m_objXmlAssist.GetNode(vxmlAdminNode, strPattern, False)
    Set xmlOmigaParentNode = m_objXmlAssist.GetNode(xmlTempReqNode, strPattern)

    For i = 1 To 5  'BMIDS00443 The block below was duplicated 5 times, so it is now inside a loop
        
        strPattern = ".//CUSTOMERTELEPHONENUMBER[USAGE = '" & CStr(i) & "']"
        'BMIDS00792
        If Not xmlAdminParentNode Is Nothing Then
            Set xmlAdminChildNode = m_objXmlAssist.GetNode(xmlAdminParentNode, strPattern, False)
        Else
            Set xmlAdminChildNode = Nothing
        End If
        'BMIDS00792 End
        Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, strPattern, False)
    
        If Not xmlOmigaChildNode Is Nothing Then
            If Not xmlAdminChildNode Is Nothing Then
                Set xmlSeqNoNode = m_objXmlAssist.GetNode(xmlOmigaChildNode, "TELEPHONESEQUENCENUMBER")
                xmlOmigaParentNode.replaceChild xmlAdminChildNode.cloneNode(True), xmlOmigaChildNode
                Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlTempReqNode, ".//CUSTOMERTELEPHONENUMBER[USAGE = '" & CStr(i) & "']", False)
                'append Customernumber and customerversionnumber tags
                xmlOmigaChildNode.appendChild xmlCustomerVersionNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlCustomerNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlSeqNoNode.cloneNode(True)
            End If
        Else
            If Not xmlAdminChildNode Is Nothing Then
                xmlOmigaParentNode.appendChild xmlAdminChildNode.cloneNode(True)
                Set xmlOmigaChildNode = m_objXmlAssist.GetNode(xmlOmigaParentNode, "CUSTOMERTELEPHONENUMBER[USAGE = '" & CStr(i) & "']", False)
                'append Customernumber and customerversionnumber tags
                xmlOmigaChildNode.appendChild xmlCustomerVersionNumberNode.cloneNode(True)
                xmlOmigaChildNode.appendChild xmlCustomerNumberNode.cloneNode(True)
            End If
        End If
   
    Next    'BMIDS00443

    'Create an instance of CustomerTxBO
    Set objICustomerTxBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerTxBO")

    'Call SavePersonalDetailsAndAddress passing in the converted XML String
    strUpdate = SavePersonalDetailsAndAddress(xmlTempReqNode.xml)
    Set xmlTempDoc = m_objXmlAssist.load(strUpdate, TypeName(Me), cstrFunctionName)
    m_objErrAssist.CheckResponse xmlTempDoc.documentElement, robjResponseNode, True

    'clear the response node
    robjResponseNode.removeChild robjResponseNode.firstChild
    'SR 29/11/2002 : BMIDS01048 - ReadOnly attribute is not to be removed
    'robjResponseNode.Attributes.removeNamedItem "READONLY"
    'SR 29/11/2002 : BMIDS01048 - End
    ICustomerBO_GetCustomerDetails vobjRequestNode, robjResponseNode

UpdateCustomerWithAdminDetailsVBErr:
    Set xmlTempDoc = Nothing
    Set xmlTempReqNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlAdminParentNode = Nothing
    Set xmlOmigaParentNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNumberNode = Nothing
    Set xmlAdminChildNode = Nothing
    Set xmlOmigaChildNode = Nothing
    Set xmlSeqNoNode = Nothing
    Set xmlAddressGuid = Nothing
    Set objICustomerTxBO = Nothing
    Set xmlAttrib = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub
'BMIDS00443 End

'BMIDS00443
Private Sub DeleteBlankAddressXML(ByRef rxmlAddresses As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Delete CustomerAddress nodes that contain no data
' pass:         rxmlAddresses   xml node containing CustomerAddresses
' return:       rxmlAddresses is passed by reference
'------------------------------------------------------------------------------------------
    Dim xmlAddress  As IXMLDOMNode
    Dim strAddress  As String
    
On Error GoTo DeleteBlankAddressXMLVBErr
    
    For Each xmlAddress In rxmlAddresses.selectNodes("CUSTOMERADDRESS")
        strAddress = xmlGetNodeText(xmlAddress, "ADDRESS/BUILDINGORHOUSENAME") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/BUILDINGORHOUSENUMBER") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/FLATNUMBER") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/STREET") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/DISTRICT") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/TOWN") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/COUNTY") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/COUNTRY") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/POSTCODE") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/DELIVERYPOINTSUFFIX") & _
            xmlGetNodeText(xmlAddress, "ADDRESS/MAILSORTCODE")
        If Len(Trim(strAddress)) = 0 Then
            rxmlAddresses.removeChild xmlAddress
        End If
    Next
    
DeleteBlankAddressXMLVBErr:
    Set xmlAddress = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub
'BMIDS00443 End

'BMIDS00443
Private Function FindLatestCustomerVersion(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Find the lastest version of a customer
' pass:         vxmlRequest     xml element containing the request with CustomerNumber
' return:       xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLatestCustomerVersionVbErr
    
    Const strFunctionName As String = "FindLatestCustomerVersion"

    
    Dim xmlRequestNode  As IXMLDOMNode
    Dim xmlDataNode     As IXMLDOMNode
    Dim xmlOut          As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICustomerDO  As ICustomerDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objICustomerDO = m_objContext.CreateInstance(App.Title & ".CustomerDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objICustomerDO.FindLatestCustomerVersion(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set FindLatestCustomerVersion = xmlResponseElem
    
    m_objContext.SetComplete
    
FindLatestCustomerVersionExit:
    
    Set xmlRequestNode = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindLatestCustomerVersionVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set FindLatestCustomerVersion = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindLatestCustomerVersionExit

End Function
'BMIDS00443 End

'BMIDS758  Add function
Public Function GetRemovedToECustomers(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetRemovedToECustomersVbErr

    Const strFunctionName As String = "GetRemovedToECustomers"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetRemovedToECustomers(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetRemovedToECustomers = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

GetRemovedToECustomersExit:

    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetRemovedToECustomersVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetRemovedToECustomers = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    '   go to clean-up section
    Resume GetRemovedToECustomersExit

End Function

Private Function ICustomerBO_GetRemovedToECustomers(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetRemovedToECustomersVbErr
    
    Const strFunctionName As String = "ICustomerBO_GetRemovedToECustomers"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
        
    Dim objICustomerDO As ICustomerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
    
    If vxmlRequest.nodeName = cstrTABLE_REMOVEDTOECUSTOMER Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_REMOVEDTOECUSTOMER).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_REMOVEDTOECUSTOMER & " tag not found"
    End If
 
    Set xmlDataNode = objICustomerDO.GetRemovedToECustomers(vxmlRequest)
    
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerBO_GetRemovedToECustomers = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    
ICustomerBO_GetRemovedToECustomersExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

ICustomerBO_GetRemovedToECustomersVbErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ICustomerBO_GetRemovedToECustomers = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ICustomerBO_GetRemovedToECustomersExit

End Function

'BBG573  Add new function
'PSC 28/06/2005 MAR5 - Start
Public Function GetNumberOfCopiesForKFI(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'               blnOrderBySequenceNUmber - Whether the list is to be sorted in the Ascending
'                                           order of AddressSequenceNumber
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfCopiesForKFIVbErr
    Dim strFunctionName As String
    strFunctionName = "GetNumberOfCopiesForKFI"
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    ' Delegate to FreeThreadedDomDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetNumberOfCopiesForKFI(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetNumberOfCopiesForKFI = xmlOut.xml
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
GetNumberOfCopiesForKFIExit:
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Exit Function
GetNumberOfCopiesForKFIVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetNumberOfCopiesForKFI = objErrAssist.CreateErrorResponse
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    '   go to clean-up section
    Resume GetNumberOfCopiesForKFIExit
End Function

Private Function ICustomerBO_GetNumberOfCopiesForKFI(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               Alias
' pass:         vstrXMLRequest  xml Request node containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ICustomerBO_GetNumberOfCopiesForKFIVbErr
    
    Dim strFunctionName As String
    strFunctionName = "ICustomerBO_GetNumberOfCopiesForKFI"
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim strCustomerNumber As String
    Dim strTempCustomer As String
    Dim iCustomerRoleType As Integer
    Dim iAddressType As Integer
    
    Dim arrPrint() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim xmlTempDoc As New FreeThreadedDOMDocument40
    Dim xmlAddressListElem As IXMLDOMElement
    Dim xmlAddressList As IXMLDOMNodeList
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerNameList As IXMLDOMNodeList
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlSaveAddress As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlCustomerNumberElem As IXMLDOMElement
    Dim iTotalAddresses As Integer
    Dim arrAddressGuid() As String
    Dim arrPostCode() As String
    Dim arrBuildingOrHouseNumber() As String
    Dim arrBuildingOrHouseName() As String
    Dim arrFlatNumber() As String
    Dim arrStreet() As String
    Dim iNumberOfKFIs As Integer
    Dim bFound As Boolean

    Dim objICustomerDO As ICustomerDO
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = objContext.CreateInstance(App.Title & ".CustomerDO")
    End If
     ' Validate to check that Customer Number is mentioned (for all customers) in the Input String
    If vxmlRequest.nodeName <> cstrTABLE_CUSTADDRESSNAME Then
        Set xmlNodeList = vxmlRequest.selectNodes("//" & cstrTABLE_CUSTADDRESSNAME)
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, cstrTABLE_CUSTADDRESSNAME & " tag not found"
        Else
            For Each xmlNode In xmlNodeList
                If objXmlAssist.GetTagValue(xmlNode, "CUSTOMERNUMBER") = "" Then
                    objErrAssist.ThrowError TypeName(Me), strFunctionName, 256
                End If
            Next xmlNode
        End If
    End If
    Set xmlDataNode = objICustomerDO.GetCustomerAddressListAndRole(vxmlRequest)
    
    'Now that the customer data is available, calculate the number of KFI copies required.
    'Get an address for each applicant (use correspondence in preference to current)
    'Send a KFI to each different address.
    
    Set xmlAddressListElem = xmlTempDoc.createElement("ADDRESSLLIST")
    Set xmlCustomerNameList = xmlDataNode.selectNodes("//CUSTOMERNAME")
    Set xmlNodeList = xmlDataNode.selectNodes("//CUSTOMERADDRESS")
   
    For i = 0 To (xmlNodeList.length - 1)
        Set xmlSaveAddress = Nothing
        Set xmlNode = xmlNodeList.Item(i)
        Set xmlCustomerNumberElem = xmlNodeList.Item(i).selectSingleNode("CUSTOMERNUMBER")
        strCustomerNumber = xmlGetNodeText(xmlNode, "CUSTOMERNUMBER")
        iCustomerRoleType = xmlGetNodeAsInteger(xmlNode, "CUSTOMERROLETYPE")
        
        'If this customer is an Applicant, save the address and customer number
        If (GetValidationTypeForValueID("CustomerRoleType", iCustomerRoleType) = "A") Then
            
            bFound = False
            For j = 0 To (xmlNodeList.length - 1)
                Set xmlTempNode = xmlNodeList.Item(j)
                strTempCustomer = xmlGetNodeText(xmlTempNode, "CUSTOMERNUMBER")
                
                If (strTempCustomer = strCustomerNumber) Then
                    'If this is the first address for this customer or it is the correspondence address, save it.
                    iAddressType = xmlGetNodeAsInteger(xmlTempNode, "ADDRESSTYPE")
                    If ((bFound = False) Or _
                        (GetValidationTypeForValueID("CustomerAddressType", iAddressType) = "C")) Then
                        Set xmlSaveAddress = xmlTempNode.selectSingleNode("ADDRESS")
          
                        'If the customer number has not already been added, add it now
                        If objXmlAssist.GetNode(xmlSaveAddress, "CUSTOMERNUMBER", False) Is Nothing Then
                            objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", xmlCustomerNumberElem, xmlSaveAddress, False
                        End If
                        
                        bFound = True
                    End If
                End If
            Next
            'Append the address for this applicant to the list
            
            'Remove the ADDRESSGUID node so that the ADDRESS nodes can be compared
            Dim xmlElem As IXMLDOMElement
            If Not objXmlAssist.GetNode(xmlSaveAddress, "ADDRESSGUID", False) Is Nothing Then
                Set xmlElem = xmlSaveAddress.selectSingleNode("ADDRESSGUID")
                xmlSaveAddress.removeChild xmlElem
            End If
                 
            objXmlAssist.CopyNode "ADDRESS", "ADDRESS", xmlSaveAddress, xmlAddressListElem, True
           

        End If
    Next
    
    'Remove any duplicate ADDRESS nodes
    objXmlAssist.RemoveDuplicates xmlAddressListElem
    
    Set xmlAddressList = xmlAddressListElem.selectNodes("ADDRESS")
    
    iTotalAddresses = xmlAddressList.length
    
    ReDim arrPrint(iTotalAddresses)
    For i = 0 To (iTotalAddresses - 1)
        arrPrint(i) = True
    Next

    'Compare the addresses and remove any duplicates
    'The following fields must match for the address to be the same
    ' Post Code
    ' BuildingOrHouseNumber
    ' BuildingOrHouseName
    ' FlatNumber
    ' Street
    
    'Convert the text to upper case before comparing
       
    ReDim arrAddressGuid(iTotalAddresses)
    ReDim arrBuildingOrHouseName(iTotalAddresses)
    ReDim arrBuildingOrHouseNumber(iTotalAddresses)
    ReDim arrFlatNumber(iTotalAddresses)
    ReDim arrStreet(iTotalAddresses)
    ReDim arrPostCode(iTotalAddresses)
   
    For i = 0 To (xmlAddressList.length - 1)
        Set xmlAddressNode = xmlAddressList.Item(i)
    
        arrAddressGuid(i) = UCase(xmlGetNodeText(xmlAddressNode, "ADDRESSGUID"))
        arrBuildingOrHouseName(i) = UCase(xmlGetNodeText(xmlAddressNode, "BUILDINGORHOUSENAME"))
        arrBuildingOrHouseNumber(i) = UCase(xmlGetNodeText(xmlAddressNode, "BUILDINGORHOUSENUMBER"))
        arrFlatNumber(i) = UCase(xmlGetNodeText(xmlAddressNode, "FLATNUMBER"))
        arrStreet(i) = UCase(xmlGetNodeText(xmlAddressNode, "STREET"))
        arrPostCode(i) = UCase(xmlGetNodeText(xmlAddressNode, "POSTCODE"))
    
    Next
    
    For j = 0 To (iTotalAddresses - 1)
        For k = 0 To (iTotalAddresses - 1)
        
            If (j <> k) Then
            
                If ((arrPostCode(j) = arrPostCode(k)) And _
                    (arrBuildingOrHouseNumber(j) = arrBuildingOrHouseNumber(k)) And _
                    (arrBuildingOrHouseName(j) = arrBuildingOrHouseName(k)) And _
                    (arrFlatNumber(j) = arrFlatNumber(k)) And _
                    (arrStreet(j) = arrStreet(k))) Then
                    
                    'Addresses match
                    'Copy the customer number from j and set the flag for j to be False
                    If arrPrint(k) = True Then
                    
                        objXmlAssist.CopyNode "CUSTOMERNUMBER", "CUSTOMERNUMBER", _
                                xmlAddressList.Item(j), _
                                xmlAddressList.Item(k), True
                   
                        arrPrint(j) = False
                    End If
                End If
           
            End If
                
        Next
    Next
    
    'Now we have
    'xmlAddressList which is a list af all the addresses together with the relevant customer numbers.
    'arrPrint() holds a flag which indicates if a KFI should go to this address
    
    'Calculate the number of KFIs needed
    iNumberOfKFIs = 0
    For i = 0 To (UBound(arrPrint) - 1)
        If arrPrint(i) = True Then
            iNumberOfKFIs = iNumberOfKFIs + 1
        End If
    Next
    
    'Formulate the response XML
    
'    <RESPONSE>
'        <NUMBEROFCOPIES>n</NUMBEROFCOPIES>
'        <ADDRESSLIST>
'            <ADDRESS>
'                <BUILDINGORHOUSENAME>...</BUILDINGORHOUSENAME>
'                <BUILDINGORHOUSENUMBER>...</BUILDINGORHOUSENUMBER>
'                <FLATNUMBER>...</FLATNUMBER>
'                <STREET>...</STREET>
'                <DISTRICT>...</DISTRICT>
'                <TOWN>...</TOWN>
'                <COUNTY>...</COUNTY>
'                <COUNTRY>...</COUNTRY>
'                <POSTCODE>...</POSTCODE>
'                <DELIVERYPOINTSUFFIX>...</DELIVERYPOINTSUFFIX>
'                <MAILSORTCODE>...</MAILSORTCODE>
'                <PAFINDICATOR>...</PAFINDICATOR>
'                <BFPO>...</BFPO>
'                <CUSTOMERNUMBER>...</CUSTOMERNUMBER>
'                <CUSTOMERNAME>
'                   <CUSTOMERNUMBER>...</CUSTOMERNUMBER>
'                   <CUSTOMERVERSIONNUMBER>...</CUSTOMERVERSIONNUMBER>
'                   <CORRESPONDENCESALUTATION>...</CORRESPONDENCESALUTATION>
'                   <FIRSTFORENAME>...</FIRSTFORENAME>
'                   <OTHERFORENAMES>...</OTHERFORENAMES>
'                   <SECONDFORENAME>...</SECONDFORENAME>
'                   <SURNAME>...</SURNAME>
'                   <TITLE>...</TITLE>
'                </CUSTOMERNAME>
'                <CUSTOMERNAME>
'                        ...
'                </CUSTOMERNAME>
'                    ...
'            </ADDRESS>
'            <ADDRESS>
'                .....
'            </ADDRESS>
'            ...
'        </ADDRESSLIST>
'    </RESPONSE>
    
    Set xmlTempNode = xmlTempDoc.createElement("NUMBEROFCOPIES")
    xmlTempNode.Text = Format(iNumberOfKFIs)
    objXmlAssist.CopyNode "NUMBEROFCOPIES", "NUMBEROFCOPIES", xmlTempNode, xmlResponseElem, False
    
    Set xmlTempNode = xmlTempDoc.createElement("ADDRESSLIST")
    
    For i = 0 To iTotalAddresses - 1
        If (arrPrint(i) = True) Then
            
            'Get the CustomerName information for each customer with this address
            Set xmlCustomerList = xmlAddressList.Item(i).selectNodes("CUSTOMERNUMBER")
            For j = 0 To (xmlCustomerList.length - 1)
                strCustomerNumber = xmlCustomerList.Item(j).Text
                
                bFound = False
                k = 0
                While (k < xmlCustomerNameList.length) And (bFound = False)
                
                    If (xmlGetNodeText(xmlCustomerNameList.Item(k), "CUSTOMERNUMBER") = strCustomerNumber) Then
                        bFound = True
                        objXmlAssist.CopyNode "CUSTOMERNAME", "CUSTOMERNAME", _
                                                xmlCustomerNameList.Item(k), _
                                                xmlAddressList.Item(i), True
                        
                    End If
                    k = k + 1
                Wend
            Next
            
            'Add this ADDRESS node to the response
            objXmlAssist.CopyNode "ADDRESS", "ADDRESS", xmlAddressList.Item(i), xmlTempNode, True

        End If
    Next
    
    objXmlAssist.CopyNode "ADDRESSLIST", "ADDRESSLIST", xmlTempNode, xmlResponseElem, True
     
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set ICustomerBO_GetNumberOfCopiesForKFI = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    
ICustomerBO_GetNumberOfCopiesForKFIExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objICustomerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempDoc = Nothing
    Set xmlAddressListElem = Nothing
    Set xmlAddressList = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomerNameList = Nothing
    Set xmlAddressNode = Nothing
    Set xmlSaveAddress = Nothing
    Set xmlTempNode = Nothing
    Set xmlCustomerNumberElem = Nothing
    Exit Function
ICustomerBO_GetNumberOfCopiesForKFIVbErr:
    
    'NOT an error for there to be NO Addresses for this customer.
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set ICustomerBO_GetNumberOfCopiesForKFI = objErrAssist.CreateErrorResponseEx()
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    ' go to clean-up section
    Resume ICustomerBO_GetNumberOfCopiesForKFIExit
End Function
'PSC 28/06/2005 MAR5 - End
'PSC 11/10/2005 MAR57 Start
'PSC 26/10/2005 MAR300
Private Function RunCustomerDataCheckAndUpdate(ByVal vxmlRequest As IXMLDOMNode, _
                                          ByVal vmlCriticalDataContext As IXMLDOMNode, _
                                          ByVal vxmlOmigaCustomerData As IXMLDOMNode, _
                                          ByVal vxmlAdminCustomerData As IXMLDOMNode) As Boolean
                                 
On Error GoTo RunCustomerDataCheckAndUpdateVbErr
    
    Const cstrFunctionName  As String = "RunCustomerDataCheckAndUpdate"
    
    Dim xmlCriticalDataCheckDoc As FreeThreadedDOMDocument40
    Dim xmlCriticalDataCheckElem As IXMLDOMElement
    Dim xmlBefore As IXMLDOMNode
    Dim xmlAfter As IXMLDOMNode
    Dim xmlAppBefore As IXMLDOMNode
    Dim xmlAppAfter As IXMLDOMNode
    Dim xmlAdminCustomerList As IXMLDOMNodeList
    Dim xmlAdminCustomer As IXMLDOMNode
    Dim xmlOmigaCustomer As IXMLDOMNode
    Dim xmlAppCustRole As IXMLDOMNode
    Dim xmlContextData As IXMLDOMNode
    Dim xmlChangedCustomers As IXMLDOMNodeList
    Dim xmlCriticalDataCustomer As IXMLDOMNode          ' PSC 20/02/2006 MAR1211
    Dim xmlChangedAddresses As IXMLDOMNodeList
    Dim xmlChangedAddress As IXMLDOMNode
    Dim xmlAddress As IXMLDOMNode
    Dim xmlLastAmendedDate As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlCriticalDataCustomers As IXMLDOMNodeList     ' PSC 20/02/2006 MAR1211
    
    ' PSC 29/11/2005 MAR727 - Start
    Dim xmlAddresses As IXMLDOMNodeList            ' PSC 20/12/2005 MAR917
    Dim xslDoc As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    ' PSC 29/11/2005 MAR727 - End
   
    Dim omTMBO As Object
    
    Dim strOtherSystemCustomerNumber As String
    Dim strSearchPattern As String
    Dim strResponse As String
    Dim strAddressType As String
    
    ' PSC 29/11/2005 MAR727 - Start
    Dim strAddressTypesSearch As String
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
    ' PSC 29/11/2005 MAR727 - End
    
    ' PSC 20/12/2005 MAR917 - Start
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlUpdateCustResponse As IXMLDOMNode
    Dim xmlSearchNode As IXMLDOMNode
    ' PSC 20/12/2005 MAR917 - End
    
    Set xmlCriticalDataCheckDoc = New FreeThreadedDOMDocument40
    xmlCriticalDataCheckDoc.validateOnParse = False
    xmlCriticalDataCheckDoc.setProperty "NewParser", True
    
    Set xmlCriticalDataCheckElem = vxmlRequest.cloneNode(False)
    xmlCriticalDataCheckDoc.appendChild xmlCriticalDataCheckElem
    xmlCriticalDataCheckElem.setAttribute "OPERATION", "CriticalDataCheck"
    xmlCriticalDataCheckElem.setAttribute "CREDITCHECK_TEST", "YES"
    xmlCriticalDataCheckElem.setAttribute "COSTMODELLING_TEST", "YES"
    
    Set xmlBefore = xmlCriticalDataCheckDoc.createElement("BEFORE")
    xmlCriticalDataCheckElem.appendChild xmlBefore
    
    Set xmlAfter = xmlCriticalDataCheckDoc.createElement("AFTER")
    xmlCriticalDataCheckElem.appendChild xmlAfter
    
    If Not vmlCriticalDataContext Is Nothing Then
        Set xmlContextData = vmlCriticalDataContext.cloneNode(True)
    Else
        Set xmlContextData = xmlCriticalDataCheckDoc.createElement("CRITICALDATACONTEXT")
    End If
    
    xmlCriticalDataCheckElem.appendChild xmlContextData
   
    xmlSetAttributeValue xmlContextData, "COMPONENT", "None"
    xmlSetAttributeValue xmlContextData, "METHOD", "None"
   
    Set xmlAppBefore = xmlCriticalDataCheckDoc.createElement("APPLICATION")
    xmlBefore.appendChild xmlAppBefore
    
    Set xmlAppAfter = xmlCriticalDataCheckDoc.createElement("APPLICATION")
    xmlAfter.appendChild xmlAppAfter
    
    Set xmlAdminCustomerList = vxmlAdminCustomerData.selectNodes(".//CUSTOMER")
    
    For Each xmlAdminCustomer In xmlAdminCustomerList
        Set xmlAppCustRole = xmlCriticalDataCheckDoc.createElement("APPLICATIONCUSTOMERROLE")
        xmlAppAfter.appendChild xmlAppCustRole
        xmlAppCustRole.appendChild xmlCreateAttributeBasedResponse(xmlAdminCustomer, True)
    
        strOtherSystemCustomerNumber = xmlGetMandatoryNodeText(xmlAdminCustomer, "OTHERSYSTEMCUSTOMERNUMBER")
        ' PSC 29/11/2005 MAR727
        strSearchPattern = ".//CUSTOMER[OTHERSYSTEMCUSTOMERNUMBER='" & strOtherSystemCustomerNumber & "']"
        Set xmlOmigaCustomer = xmlGetMandatoryNode(vxmlOmigaCustomerData, strSearchPattern)
        
        Set xmlAppCustRole = xmlCriticalDataCheckDoc.createElement("APPLICATIONCUSTOMERROLE")
        xmlAppBefore.appendChild xmlAppCustRole
        xmlAppCustRole.appendChild xmlCreateAttributeBasedResponse(xmlOmigaCustomer, True)
    Next
    
    ' PSC 29/11/2005 MAR727 - Start
    Set xslDoc = New FreeThreadedDOMDocument40
    xslDoc.validateOnParse = False
    xslDoc.setProperty "NewParser", True
    xslDoc.async = False

    strXslPath = App.Path & "\RunCustomerDataCheckAndUpdate.xslt"
    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
    blnLoadedOk = xslDoc.load(strXslPath)
    
    If Not blnLoadedOk Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
    End If
    
    Set xmlTransformedDoc = New FreeThreadedDOMDocument40
    xmlTransformedDoc.validateOnParse = False
    xmlTransformedDoc.setProperty "NewParser", True
    xmlTransformedDoc.async = False

    xmlCriticalDataCheckDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
    ' PSC 29/11/2005 MAR727 - End

    Set omTMBO = m_objContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTMBO")
    strResponse = omTMBO.omTMRequest(xmlTransformedDoc.xml)
    Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
    ' PSC 25/10/2005 MAR300
    ' PSC 17/02/2006 MAR1211 - Update customer data on all customers
    ' PSC 20/02/2006 MAR1211 - Start
    Set xmlCriticalDataCustomers = xmlTempResponse.selectNodes("RESPONSE/CUSTOMERS/CUSTOMER")
    Set xmlChangedCustomers = xmlTempResponse.selectNodes("RESPONSE/CUSTOMERS/CUSTOMER[@DATACHANGED='1']")
    ' PSC 20/02/2006 MAR1211 - End
    
    ' PSC 20/02/2006 MAR1211
    For Each xmlCriticalDataCustomer In xmlCriticalDataCustomers
        ' PSC 20/12/2005 MAR917
        ' PSC 20/02/2006 MAR1211
        strOtherSystemCustomerNumber = xmlGetMandatoryAttributeText(xmlCriticalDataCustomer, "OTHERSYSTEMCUSTOMERNUMBER")
        ' PSC 29/11/2005 MAR727
        strSearchPattern = ".//CUSTOMER[OTHERSYSTEMCUSTOMERNUMBER='" & strOtherSystemCustomerNumber & "']"
        Set xmlAdminCustomer = xmlGetMandatoryNode(vxmlAdminCustomerData, strSearchPattern)
        ' PSC 20/12/2005 MAR917
        Set xmlOmigaCustomer = xmlGetMandatoryNode(vxmlRequest, strSearchPattern)
        
        ' PSC 29/11/2005 MAR727 - Start
        ' PSC 20/02/2006 MAR1211
        Set xmlChangedAddresses = xmlCriticalDataCustomer.selectNodes("ADDRESS[@ADDRESSCHANGED='1']")
        
        For Each xmlChangedAddress In xmlChangedAddresses
            ' PSC 20/12/2005 MAR917
            strAddressType = xmlGetMandatoryAttributeText(xmlChangedAddress, "ADDRESSTYPE")
            
            If Len(strAddressTypesSearch) > 0 Then
                strAddressTypesSearch = strAddressTypesSearch & " and "
            End If
            
            strAddressTypesSearch = strAddressTypesSearch & "ADDRESSTYPE!='" & strAddressType & "'"
            
        Next
        
        ' Drop addresses that haven't changed
        ' PSC 20/12/2005 MAR917
        ' PSC 12/01/2006 MAR1031 - Start
        If Len(strAddressTypesSearch) = 0 Then
            strSearchPattern = ".//CUSTOMERADDRESS"
        Else
            strSearchPattern = ".//CUSTOMERADDRESS[" & strAddressTypesSearch & "]"
        End If
        
        Set xmlAddresses = xmlAdminCustomer.selectNodes(strSearchPattern)
        ' PSC 12/01/2006 MAR1031 - End
        
        For Each xmlAddress In xmlAddresses
            xmlAddress.parentNode.removeChild xmlAddress
        Next
        
        ' Set the last amended date on remaining addresses as they have changed
        ' PSC 20/12/2005 MAR917
        Set xmlAddresses = xmlAdminCustomer.selectNodes(".//CUSTOMERADDRESS")
        
        For Each xmlAddress In xmlAddresses
            ' PSC 25/10/2005 MAR300
            Set xmlLastAmendedDate = xmlGetNode(xmlAddress, "LASTAMENDEDDATE")
            
            If xmlLastAmendedDate Is Nothing Then
                Set xmlLastAmendedDate = xmlCriticalDataCheckDoc.createElement("LASTAMENDEDDATE")
                xmlAddress.appendChild xmlLastAmendedDate
            End If
            
            ' PSC 24/01/2006 MAR1102
            xmlLastAmendedDate.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
        Next
        ' PSC 29/11/2005 MAR727 - End
        
        ' Need to Update Address
        ' PSC 20/12/2005 MAR917 - Start
        Set xmlTempRequest = vxmlRequest.cloneNode(False)
        Set xmlSearchNode = xmlCriticalDataCheckDoc.createElement("SEARCH")
        xmlTempRequest.appendChild xmlSearchNode
        xmlSearchNode.appendChild xmlOmigaCustomer.cloneNode(True)
        Set xmlUpdateCustResponse = xmlCriticalDataCheckDoc.createElement("RESPONSE")
        
        UpdateCustomerWithAdminDetails xmlTempRequest, xmlUpdateCustResponse, xmlAdminCustomer
        ' PSC 20/12/2005 MAR917 - End

    Next
    
    ' PSC 26/10/2005 MAR300 - Start
    If xmlChangedCustomers.length > 0 Then
        RunCustomerDataCheckAndUpdate = True
    End If
    ' PSC 26/10/2005 MAR300 - End
     
RunCustomerDataCheckAndUpdateVbErr:

    Set xmlCriticalDataCheckDoc = Nothing
    Set xmlCriticalDataCheckElem = Nothing
    Set xmlBefore = Nothing
    Set xmlAfter = Nothing
    Set xmlAppBefore = Nothing
    Set xmlAppAfter = Nothing
    Set xmlAdminCustomerList = Nothing
    Set xmlAdminCustomer = Nothing
    Set xmlOmigaCustomer = Nothing
    Set xmlAppCustRole = Nothing
    Set xmlContextData = Nothing
    Set omTMBO = Nothing
    Set xmlChangedCustomers = Nothing
    Set xmlCriticalDataCustomer = Nothing       ' PSC 20/02/2006 MAR1211
    Set xmlChangedAddresses = Nothing
    Set xmlChangedAddress = Nothing
    Set xmlAddress = Nothing
    Set xmlLastAmendedDate = Nothing
    Set xmlCriticalDataCustomers = Nothing      ' PSC 20/02/2006 MAR1211
    
    ' PSC 29/11/2005 MAR727 - Start
    Set xmlAddresses = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    ' PSC 29/11/2005 MAR727 - End
    
    ' PSC 20/12/2005 MAR917 - Start
    Set xmlTempRequest = Nothing
    Set xmlUpdateCustResponse = Nothing
    Set xmlSearchNode = Nothing
    ' PSC 20/12/2005 MAR917 - End
    
    If Err.Number <> 0 Then
        ' PSC 20/12/2005 MAR917
        Err.Raise Err.Number, cstrFunctionName & "." & Err.Source, Err.Description
    End If
End Function

Public Function ProcessLaunchCustomer(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a single instance of the persistant data associated with this
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ProcessLaunchCustomerVbErr

    Const strFunctionName As String = "ProcessLaunchCustomer"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objIScriptInterface As IScriptInterface
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCustomerOut As IXMLDOMNode
    

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    'Validate the Input (request string)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    ' Run validation checks
    Set xmlTempResponseNode = ICustomerBO_ValidateCustomerAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

   ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
       
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_ProcessLaunchCustomer(xmlIn.documentElement)
    Set xmlResponseElem = xmlTempResponseNode
    
    ProcessLaunchCustomer = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ProcessLaunchCustomerExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIScriptInterface = Nothing
    Set xmlCustomerOut = Nothing

    Exit Function

ProcessLaunchCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ProcessLaunchCustomer = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    '  go to clean-up section
    Resume ProcessLaunchCustomerExit

End Function

Private Function ICustomerBO_ProcessLaunchCustomer(ByVal vxmlRequestNode As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates a single instance of the persistant data associated with this
'               invokes appropriate transactioning (Tx) business object
' pass:         vXMLRequest  xml Request node containing data to be persistes
' return:       xml Response data stream containing results of operation
'                    either: TYPE="SUCCESS"
'                    or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ProcessLaunchCustomerExit
    
    Const strFunctionName As String = "ICustomerBO_ProcessLaunchCustomer"

    Dim xmlCustomerIn As IXMLDOMNode
    Dim xmlTempReqNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlOmigaCustomer As IXMLDOMNode
    Dim xmlReadOnlyAttribute As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objCustomerTxBO As ICustomerTxBO
    
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    
    Dim strOtherSystemCustomerNo As String
    Dim strCustomerVersionNo As String
    Dim strCustomerNo As String
    Dim lngErrNo As Long
    Dim blnNewCustomer As Boolean
    'MAR947 GHun
    Dim xmlCAddress As IXMLDOMElement
    Dim xmlAddress As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    'MAR947 End
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    
    Set xmlResponseNode = xmlTempDoc.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponseNode, "TYPE", "SUCCESS"
    
    Set xmlCustomerIn = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    strOtherSystemCustomerNo = xmlGetMandatoryAttributeText(xmlCustomerIn, "CUSTOMERNUMBER")
    strCustomerNo = xmlGetAttributeText(xmlCustomerIn, "OMIGACUSTOMERNUMBER")
    
    blnNewCustomer = False
    
    ' Check if there is an existing Omiga customer with a matching OtherSystemCustomerNumber
    Set xmlTempReqNode = vxmlRequestNode.cloneNode(False)
    Set xmlCustomer = xmlTempDoc.createElement("CUSTOMER")
    xmlTempReqNode.appendChild xmlCustomer
    Set xmlNode = xmlTempDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
    xmlNode.Text = strOtherSystemCustomerNo
    xmlCustomer.appendChild xmlNode
            
    Set xmlTempResponseNode = ICustomerBO_GetCustomerFromLegacyNumber(xmlTempReqNode)
    lngErrNo = errCheckXMLResponseNode(xmlTempResponseNode, xmlResponseNode, False)
    
    If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
        errRaiseXMLResponseNode xmlTempResponseNode
    End If
    
    If lngErrNo = 0 Then
        Set xmlOmigaCustomer = xmlTempResponseNode.selectSingleNode("CUSTOMER")
        strCustomerNo = xmlGetMandatoryNodeText(xmlOmigaCustomer, "CUSTOMERNUMBER")
        strCustomerVersionNo = xmlGetMandatoryNodeText(xmlOmigaCustomer, "CUSTOMERVERSION/CUSTOMERVERSIONNUMBER")
    End If
    
    ' If there is an Omiga customer try to lock it
    
    If Len(strCustomerNo) > 0 Then
        'Create a customer lock and check if any errors occurred
        Set xmlTempReqNode = m_objXmlAssist.GetRequestNodeEx(vxmlRequestNode)
        Set xmlCustomer = xmlTempDoc.createElement("CUSTOMER")
        xmlTempReqNode.appendChild xmlCustomer
        Set xmlNode = xmlTempDoc.createElement("CUSTOMERNUMBER")
        xmlNode.Text = strCustomerNo
        xmlCustomer.appendChild xmlNode
        
        On Error Resume Next
        ICustomerBO_CreateCustomerLock xmlTempReqNode, xmlResponseNode
    
        If Err.Number <> 0 Then
            If errGetOmigaErrorNumber(Err.Number) = oeDuplicateKey Then
                Set xmlReadOnlyAttribute = xmlTempDoc.createAttribute("READONLY")
                xmlReadOnlyAttribute.Value = "1"
                xmlResponseNode.Attributes.setNamedItem xmlReadOnlyAttribute
                Set xmlReadOnlyAttribute = Nothing
            Else
                On Error GoTo ProcessLaunchCustomerExit
                Err.Raise Err.Number, Err.Source, Err.Description
            End If
        End If
        On Error GoTo ProcessLaunchCustomerExit
    Else ' Create new version of the customer
    
        'MAR947 GHun Copy the attributes from the address node (if it exists) to the parent node
        For Each xmlCAddress In xmlCustomerIn.selectNodes("*[name()='CURRENTADDRESS' or name()='CORRESPONDENCEADDRESS']")
            Set xmlAddress = xmlCAddress.selectSingleNode("ADDRESS")
            If Not xmlAddress Is Nothing Then
                For Each xmlAttrib In xmlAddress.Attributes
                    If Len(xmlAttrib.Value) > 0 Then
                        xmlCAddress.setAttribute xmlAttrib.Name, xmlAttrib.Value
                    End If
                Next
                xmlCAddress.removeChild xmlAddress
            End If
        Next
        'MAR947 End
        
        Set xmlCustomer = xmlMakeNodeElementBased(xmlCustomerIn, True, "CUSTOMER")

        Set xmlOmigaCustomer = ConvertCustomerToOmigaFormat(xmlCustomer)
        
        Set xmlTempReqNode = vxmlRequestNode.cloneNode(False)
        Set xmlNode = xmlTempDoc.createElement("CREATE")
        xmlTempReqNode.appendChild xmlNode
        xmlNode.appendChild xmlOmigaCustomer
        
        Set objCustomerTxBO = m_objContext.CreateInstance(App.Title & ".CustomerTxBO")
        objCustomerTxBO.CreateCustomerDetails xmlTempReqNode, xmlResponseNode
        Set xmlOmigaCustomer = xmlResponseNode.selectSingleNode("CUSTOMERKEY")
        strCustomerNo = xmlGetMandatoryNodeText(xmlOmigaCustomer, "CUSTOMERNUMBER")
        strCustomerVersionNo = xmlGetMandatoryNodeText(xmlOmigaCustomer, "CUSTOMERVERSIONNUMBER")
        blnNewCustomer = True
    End If
    
    Set xmlCustomer = xmlTempDoc.createElement("CUSTOMER")
    xmlResponseNode.appendChild xmlCustomer
    Set xmlNode = xmlTempDoc.createElement("CUSTOMERNUMBER")
    xmlNode.Text = strCustomerNo
    xmlCustomer.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlNode.Text = strCustomerVersionNo
    xmlCustomer.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
    xmlNode.Text = strOtherSystemCustomerNo
    xmlCustomer.appendChild xmlNode
    
    If blnNewCustomer Then
        xmlSetAttributeValue xmlCustomer, "NEWCUSTOMER", "1"
    End If
   
    Set ICustomerBO_ProcessLaunchCustomer = xmlResponseNode


ProcessLaunchCustomerExit:
        
    Set xmlCustomerIn = Nothing
    Set xmlTempReqNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCustomer = Nothing
    Set xmlOmigaCustomer = Nothing
    Set xmlReadOnlyAttribute = Nothing
    Set xmlNode = Nothing
    Set objCustomerTxBO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlResponseNode = Nothing
    'MAR947 GHun
    Set xmlCAddress = Nothing
    Set xmlAddress = Nothing
    Set xmlAttrib = Nothing
    'MAR947 End
       
    If Err.Number <> 0 Then
        ' re-raise error
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    
    Exit Function

End Function

'PSC  11/10/2005 MAR57 End

'MAR23  Add function
Public Function GetCustomerNumbers(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the customer version  number and othersystemcustomernumber
' pass:         vXMLRequest  xml Request node which identifies instance of the persistant
'                            data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerNumbersVbErr

    Const strFunctionName As String = "GetCustomerNumbers"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Call xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICustomerBO_GetCustomerNumbers(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetCustomerNumbers = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

GetCustomerNumbersExit:
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetCustomerNumbersVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCustomerNumbers = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    ' go to clean-up section
    Resume GetCustomerNumbersExit

End Function

'MAR23 Added CreateKYCCheck
Public Function CreateKYCCheck(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateKYCCheckVbErr

    Const strFunctionName As String = "CreateKYCCheck"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = ICustomerBO_CreateKYCCheck(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateKYCCheck = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateKYCCheckExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
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

    CreateKYCCheck = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateKYCCheckExit

End Function

Private Function ICustomerBO_CreateKYCCheck(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateKYCCheckVbErr
    
    Const strFunctionName As String = "ICustomerBO_CreateKYCCheck"
        
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' PSC 03/04/2006 MAR1573 - Start
    Dim xmlTempRequest As IXMLDOMNode
    Dim objCRUDTxBO As Object
    Dim strResponse As String
    ' PSC 03/04/2006 MAR1573 - End
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' PSC 03/04/2006 MAR1573 - Start
    Set objCRUDTxBO = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDTxBO")
    
    Set xmlTempRequest = xmlCreateAttributeBasedResponse(vxmlRequest, True)
    xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "CREATE"
    xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCustSchema"
    strResponse = objCRUDTxBO.omRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    ' PSC 03/04/2006 MAR1573 - End

    Set ICustomerBO_CreateKYCCheck = xmlResponseElem

    m_objContext.SetComplete
    
CreateKYCCheckExit:

    ' PSC 03/04/2006 MAR1573 - Start
    Set xmlTempRequest = Nothing
    Set objCRUDTxBO = Nothing
    ' PSC 03/04/2006 MAR1573 - End
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
     

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

    Set ICustomerBO_CreateKYCCheck = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateKYCCheckExit

End Function



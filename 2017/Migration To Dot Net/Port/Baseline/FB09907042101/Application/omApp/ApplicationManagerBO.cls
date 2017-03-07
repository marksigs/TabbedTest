VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationManagerBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Dependencies:  ApplicationManagerDO, ApplicationManagerTxBO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date      Description
'SR    12/01/00  Restructured
'SR    13/03/00 Check whether System requires Locking in the methods : CreateLock,
'                DeleteLock, LockCustomersForApplication, UnlockApplicationAndCustomers
'SR    31/05/00  SYS0794 - Modified method 'FindCustomersForApplication'
'MS    10/06/00  Raise Errors other than recordNotFound correctly
'MC    27/06/00  Add RunAIPScoreAndDownload method and associated private methods
'SR    13/09/00  SYS1138, SYS1435 New Method SaveCustomerAndEmploymentDetails
'AS    15/11/00  CORE000002: Added FindApplicationsList method
'DJP   19/12/00  Added FindApplicationOwnershipList
'ASm   12/01/01  SYS1817: oeInvalidNumberOfApplicants, oeInvalidNumberOfGuarantors
'                error enums removed and replaced by there respective error numbers
'PSC   12/01/01  SYS1749 Amend FindBusinessForCustomer to search administration system
'                Add ImportAccountsIntoApplication()
'PSC   22/02/01  SYS1749 Amend FindBusinessForCustomer to return combo text for type of application
'PSC   26/02/01  SYS1749 Amend FindBusinessForCustomer to only do admin search if there is a
'                OtherSystemCustomerNumber present
'PSC   26/02/01  SYS1749 Amend FindBusinessForCustomer so that OtherSystemCustomerNumber is not
'                mandatory
'PSC   08/03/01  SYS1879 Add GetApplicationSummaryData
'PSC   09/03/01  SYS1879 Amend GetApplicationSummaryData to return Application Number,
'                Type of Application and Address fields correctly
'PSC   14/03/01  SYS2067 Amend GetApplicationSummaryData to bring back correct employment
'                details
'SR    13/06/01  SYS2362 Modified method 'IApplicationManagerBO_FindBusinessForCustomer'
'JR    03/07/01  SYSOmiplus24, Add AddTelephoneNumberNode and modified GetApplicationSummaryData
'DRC   08/10/01  SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'PSC   12/12/01  SYS3318 Return Net Monthly Income in GetApplicationSummaryData
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog  Date      Description
'PSC   11/10/05  MAR57 - Amend FindBusinessForCustomer to use global parameter UseAdminFindBusForCustomer
'HMA   20/02/06  MAR1040 Changed GetApplicationSummaryData to retrieve IncomeSummary fields.
'------------------------------------------------------------------------------------------
Option Explicit

Implements IApplicationManagerBO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
 
Private Const cstrTABLE_NAME As String = "APPLICATION"
Private Const cstrTABLE_CUSTLOCKAPPLLOCK As String = "CUSTOMERLOCKAPPLICATIONLOCK"
Private Const cstrTABLE_CUSTOMER As String = "CUSTOMER"
Private Const cstrTABLE_APPLLOCK As String = "APPLICATIONLOCK"
Private Const cstrTABLE_CUSTOMERROLE As String = "CUSTOMERROLE"
Private Const cstrCOMPONENT_NAME As String = "omApp"
Private Const cstrGetAddData As String = "GETADDDATAFORACCDOWNLOAD"

Private Enum SERVICE_TYPE
    stCREDIT_CHECK
    stRISK_ASSESSMENT
    stOM4TO3DOWNLOAD
End Enum

Public Function AddCustomerToApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'
' pass:         vstrXMLRequest
' return:
'   AddCustomerToApplication    xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerToApplicationVbErr
    
    Const strFunctionName As String = "AddCustomerToApplication"


    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
                                                                                               
    Set xmlTempResponseNode = IApplicationManagerBO_Validate( _
                            xmlIn.documentElement, bomAddCustomerToApplication)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
                IApplicationManagerBO_AddCustomerToApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    AddCustomerToApplication = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

AddCustomerToApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

AddCustomerToApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    AddCustomerToApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AddCustomerToApplicationExit
    
End Function

Public Function FindCustomersForApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Find all customers attached to the specified application fact find.
' pass:
'   vstrXMLRequest
'                   Format:
'                   <REQUEST>
'                       <APPLICATIONNUMBER>
'                       </APPLICATIONNUMBER>
'                       <APPLICATIONFACTFINDNUMBER>
'                       </APPLICATIONFACTFINDNUMBER>
'                   </REQUEST>
' return:
'   FindCustomersForApplication
'                   xml Response data stream containing results of operation
'                   either: TYPE="SUCCESS" and xml representation of data
'                   or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomersForApplicationVbErr

    Const strFunctionName As String = "FindCustomersForApplication"


    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objCustomerBO As CustomerBO
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim xmlTemp As IXMLDOMElement
    Dim xmlCustRoleInput As IXMLDOMElement
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlCustomerVersionNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objXmlRequestNode As IXMLDOMNode

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    objXmlOut.appendChild objXmlResponseElem

    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objCustomerBO = New CustomerBO
    Else
        Set objIApplManagerDO = _
            m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objCustomerBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If

    '------------------------------------------------------------------------------------------
    ' validate parameters
    '------------------------------------------------------------------------------------------
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    strApplicationNumber = _
        m_objXmlAssist.GetTagValue(objXmlIn.documentElement, "APPLICATIONNUMBER")
    strApplicationFFNumber = _
        m_objXmlAssist.GetTagValue(objXmlIn.documentElement, "APPLICATIONFACTFINDNUMBER")

    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFFNumber) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, _
            "Values for both APPLICATIONNUMBER and APPLICATIONFACTFINDNUMBER must be specified"
    End If

    '------------------------------------------------------------------------------------------
    ' get CustomerRole list
    '------------------------------------------------------------------------------------------
    Set xmlCustRoleInput = objXmlDoc.createElement("CUSTOMERROLE")
    Set xmlTemp = objXmlDoc.createElement("APPLICATIONNUMBER")
    xmlTemp.Text = strApplicationNumber
    xmlCustRoleInput.appendChild xmlTemp
    
    Set xmlTemp = objXmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTemp.Text = strApplicationFFNumber
    xmlCustRoleInput.appendChild xmlTemp
    
    Set xmlTemp = objIApplManagerDO.FindCustomerRoleList(xmlCustRoleInput)
    
    '------------------------------------------------------------------------------------------
    ' for each customer in the list, get CustomerVersion and CustomerAddressList
    '------------------------------------------------------------------------------------------
    Set objXmlNodeList = _
        xmlTemp.getElementsByTagName("CUSTOMERROLE")
    Dim intNumCustomers As Integer

    intNumCustomers = objXmlNodeList.length
    Dim intCustomer As Integer

    Dim strRequestXml As String, strResponseXml As String
    Dim lngErrNumber As Long

    For intCustomer = 0 To intNumCustomers - 1

        '------------------------------------------------------------------------------------------
        ' get CustomerVersion and insert it into the CustomerRoleList xml
        '------------------------------------------------------------------------------------------
     
        Set objXmlRequestNode = objXmlNodeList.Item(intCustomer).cloneNode(True)
        strRequestXml = objXmlRequestNode.xml
        
        ' Modify the Request string, as required by CustomerBO
        strRequestXml = Replace(strRequestXml, "CUSTOMERROLE", "CUSTOMER")
        strRequestXml = Replace(strRequestXml, "/CUSTOMERROLE", "/CUSTOMER")
        Dim blnRaiseRecordNotFoundError As Boolean
        blnRaiseRecordNotFoundError = True
        strResponseXml = objCustomerBO.GetCustomerPersonalDetails(strRequestXml)
        Set objXmlDoc = m_objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
 '       lngErrNumber = m_objErrAssist.GetErrorNumberFromResponse(objXmlDoc.documentElement, True)
        m_objErrAssist.CheckResponse objXmlDoc.documentElement, objXmlResponseElem, True
        m_objXmlAssist.AttachResponseData objXmlResponseElem, objXmlDoc.documentElement
        blnRaiseRecordNotFoundError = False
        
        'MS 10/06/00 Raise Errors other than recordNotFound correctly
'        If lngErrNumber <> 0 And lngErrNumber <> oeRecordNotfound Then
'            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngErrNumber
            
'        End If
        
        Set objXmlCustomerVersionNode = _
            objXmlNodeList.Item(intCustomer).appendChild(objXmlDoc.documentElement.firstChild)
        '-------------------------------------------------------------------------------------
        ' Get CustomerAddress list and insert it into the CustomerVersion xml.
        ' N.B. Ignore "record not found" error
        '-------------------------------------------------------------------------------------
        Set objXmlElem = objXmlDoc.createElement("ADDRESSTYPE")
        objXmlElem.Text = "1"   ' Home / Main
        objXmlRequestNode.appendChild objXmlElem
        strRequestXml = objXmlRequestNode.xml

        ' Modify the Request string, as required by CustomerBO for CustomerAddress
        strRequestXml = Replace(strRequestXml, "CUSTOMERROLE", "CUSTOMERADDRESS")
        strRequestXml = Replace(strRequestXml, "/CUSTOMERROLE", "/CUSTOMERADDRESS")
        strResponseXml = objCustomerBO.FindCustomerAddressList(strRequestXml)

        Set objXmlDoc = m_objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
        lngErrNumber = m_objErrAssist.GetErrorNumberFromResponse(objXmlDoc.documentElement, True)

        If lngErrNumber = 0 Then
            ' SR 31/05/00 SYS0794 - Append customerAddress data only if they exist
            If Not objXmlDoc.documentElement.firstChild Is Nothing Then
                objXmlCustomerVersionNode.appendChild objXmlDoc.documentElement.firstChild
            End If
        Else
            If lngErrNumber <> oeRecordNotFound Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngErrNumber
            End If
        End If
    Next

    objXmlResponseElem.appendChild xmlTemp

    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"

    FindCustomersForApplication = objXmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
FindCustomersForApplicationExit:

    Set objXmlNode = Nothing
    Set objXmlCustomerVersionNode = Nothing
    Set objXmlNodeList = Nothing
    Set objXmlResponseElem = Nothing
    Set xmlTemp = Nothing
    Set objXmlElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objIApplManagerDO = Nothing
    Set objXmlDoc = Nothing

    Exit Function

FindCustomersForApplicationVbErr:
    
     'MS 10/06/00 Raise Errors other than recordNotFound correctly
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        If blnRaiseRecordNotFoundError = True Then
            Resume Next
        End If
    End If
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    FindCustomersForApplication = m_objErrAssist.CreateErrorResponse
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindCustomersForApplicationExit

End Function

Public Function DeleteAllCustomerLockApplicationLock(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete all instances of the persistant data associated with the
'               table 'CustomerLockApplicationLock'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "DeleteAllCustomerLockApplicationLock"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
        '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        DeleteAllCustomerLockApplicationLock = xmlResponseElem.xml
        GoTo DeleteExit
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
     Set xmlTempResponseNode = _
            IApplicationManagerBO_DeleteAllCustomerLockApplicationLock( _
                                                    xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteAllCustomerLockApplicationLock = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAllCustomerLockApplicationLock = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit
End Function

Public Function DeleteLock(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the
'               table 'ApplicationLock'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLockVbErr

    Const strFunctionName As String = "DeleteLock"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        DeleteLock = xmlResponseElem.xml
        GoTo DeleteLockExit
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
     Set xmlTempResponseNode = _
            IApplicationManagerBO_DeleteLock(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteLock = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteLockExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLockVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteLock = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLockExit
End Function

Public Function CreateCustomerLockApplicationLock(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       Create          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateCustomerLockApplicationLock"

        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        CreateCustomerLockApplicationLock = xmlResponseElem.xml
        GoTo CreateExit
    End If


    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, _
                         IApplicationManagerBOMethod.bomCreateCustomerLockApplicationLock)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
     m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
        IApplicationManagerBO_CreateCustomerLockApplicationLock(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateCustomerLockApplicationLock = xmlResponseElem.xml
    
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

    CreateCustomerLockApplicationLock = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function IsStaffApplication(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo IsStaffApplicationVbErr

    Const strFunctionName As String = "IsStaffApplication"


    Dim objConvAssist As ConvertAssist
    Dim objErrAssist As ErrAssist
    
    Set objErrAssist = New ErrAssist
    Set objConvAssist = New ConvertAssist
    
    Dim xmlIn  As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim objXmlResponseElem  As IXMLDOMElement
    Dim xmlCustomersForApplication  As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCustomerNodelist As IXMLDOMNodeList
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlTempRespElem  As IXMLDOMElement
    
    Dim strResponse  As String
    
    Dim blnMemberOfStaff  As Boolean
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set objXmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild objXmlResponseElem
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    strResponse = FindCustomersForApplication(xmlIn.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    Set xmlCustomersForApplication = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
    Set xmlCustomerNodelist = xmlCustomersForApplication.selectNodes("./RESPONSE/CUSTOMER/CUSTOMERVERSION")
    
    If xmlCustomerNodelist.length > 0 Then
        
        For Each xmlCustomerNode In xmlCustomerNodelist
            
            blnMemberOfStaff = False
            
            blnMemberOfStaff = objConvAssist.CSafeBool(m_objXmlAssist.GetNodeValue(xmlCustomerNode, "MEMBEROFSTAFF", False))
            
            If blnMemberOfStaff Then
                Exit For
            End If
            
        Next
        
    End If
    
    Set xmlTempRespElem = xmlOut.createElement("MEMBEROFSTAFF")
    objXmlResponseElem.appendChild xmlTempRespElem
    
    If blnMemberOfStaff Then
        xmlTempRespElem.Text = "1"    'true
    Else
        xmlTempRespElem.Text = "0"    'false
    End If
    
    IsStaffApplication = objXmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set objConvAssist = Nothing
    Set xmlIn = Nothing
    Set objXmlResponseElem = Nothing
    Set xmlCustomersForApplication = Nothing
    Set xmlOut = Nothing
    Set xmlCustomerNodelist = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempRespElem = Nothing
    
    
    Exit Function

IsStaffApplicationVbErr:

    objErrAssist.AddToErrSource strFunctionName
    
    'Set objGlobalParamDO = Nothing
    Set objErrAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ValidateCustomerLockApplicationLock(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vXMLRequest  xml Request node containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateCustomerLockApplicationLock"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objSqlAssist As SQLAssist
    Dim objADOAssist As ADOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objNodeList As IXMLDOMNodeList
    Dim objNode As IXMLDOMNode

    Dim lngErrorNo As Long

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objNodeList = vxmlRequest.getElementsByTagName("CUSTOMERNUMBER")
    
    Dim strCondition As String
    Set objSqlAssist = New SQLAssist
    
    ' Check that each customer doesn't have any locks already
    For Each objNode In objNodeList
        Dim strResponse As String

        strCondition = "CUSTOMERNUMBER=" & objSqlAssist.FormatString(objNode.Text)
        If objADOAssist.CheckRecordExists(cstrTABLE_CUSTLOCKAPPLLOCK, strCondition) Then
             objErrAssist.ThrowError TypeName(Me), strFunctionName, 190
        End If
    Next

    '   clean-up code
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objNodeList = Nothing
    Set objNode = Nothing

Exit Sub

ValidateVbErr:

    If Err.Number = oeRecordNotFound Then
        lngErrorNo = Err.Number
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objNodeList = Nothing
    Set objNode = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function FindCustomerLockApplicationLockList( _
                                    ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindCustomerLockApplicationLockList"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = _
            IApplicationManagerBO_FindCustomerLockApplicationLockList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindCustomerLockApplicationLockList = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCustomerLockApplicationLockList = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function CreateLock(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with the
'               Table 'ApplicationLock'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "CreateLock"

        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        CreateLock = xmlResponseElem.xml
        GoTo CreateLockExit
    End If

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, _
                                    IApplicationManagerBOMethod.bomCreateApplicationLock)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
     m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
                IApplicationManagerBO_CreateLock(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateLock = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateLockExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateLockVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateLock = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLockExit

End Function

Public Function GetLockData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:       GetLockData     xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLockDataVbErr

    Const strFunctionName As String = "GetLockData"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationManagerBO_GetLockData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLockData = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetLockDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLockDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLockData = m_objErrAssist.CreateErrorResponse

'    If Not m_objContext Is Nothing Then
'        m_objContext.setcomplete
'    End If
    
    ' go to clean-up section
    Resume GetLockDataExit

End Function

Public Function DeleteCustomerLockApplicationLock(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the table
'               'CustomerLockApplicationLock'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "DeleteCustomerLockApplicationLock"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        DeleteCustomerLockApplicationLock = xmlResponseElem.xml
        GoTo DeleteExit
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_DeleteCustomerLockApplicationLock(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteCustomerLockApplicationLock = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteCustomerLockApplicationLock = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit
End Function

Public Function LockCustomersForApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   lock all customers attached to the specified application number
' pass:
'   vstrXMLRequest  XML data stream in format:
'                       <REQUEST>
'                           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                           <CUSTOMERNUMBER></CUSTOMERNUMBER>
'                       </REQUEST>
' return:
'   LockCustomersForApplication
'                   xml Response data stream containing results of operation
'                   either: TYPE="SUCCESS" and xml representation of data
'                   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo LockCustomersForApplicationVbErr
    
    Const strFunctionName As String = "LockCustomersForApplication"


    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlMessageElem As IXMLDOMElement
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        LockCustomersForApplication = xmlResponseElem.xml
        GoTo LockCustomersForApplicationExit
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                                    strFunctionName, xmlIn.documentElement)
     m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
        IApplicationManagerBO_LockCustomersForApplication(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    LockCustomersForApplication = xmlTempResponseNode.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

LockCustomersForApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlMessageElem = Nothing

    Exit Function

LockCustomersForApplicationVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    LockCustomersForApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume LockCustomersForApplicationExit

End Function

Public Function FindBusinessForCustomer(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Find the Business Type and other related for the given customer
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'               Format <REQUEST>
'                           <CUSTOMER>
'                               <CUSTOMERNUMBER> </CUSTOMERNUMBER>
'                               <CUSTOMERVERSIONNUMBER> </CUSTOMERVERSIONNUMBER>
'                            </CUSTOMER>
'                       </REQUEST>
'
' return:       FindBusinessForCustomer    xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
 On Error GoTo FindBusinessForCustomerVbErr

    Const strFunctionName As String = "FindBusinessForCustomer"
    
       
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
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationManagerBO_FindBusinessForCustomer(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindBusinessForCustomer = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindBusinessForCustomerExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindBusinessForCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindBusinessForCustomer = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindBusinessForCustomerExit

End Function

Public Function MaintainCustomerRoles(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo MaintainCustomerRolesVbErr
    
    Const strFunctionName As String = "MaintainCustomerRoles"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
     Set xmlTempResponseNode = IApplicationManagerBO_MaintainCustomerRoles(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    MaintainCustomerRoles = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

MaintainCustomerRolesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

MaintainCustomerRolesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    MaintainCustomerRoles = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume MaintainCustomerRolesExit


End Function

Public Function UpdateCustomerRoles(ByVal vstrXMLRequest As String) As String
' header --------------------------------------------------------------------
' description:
'
' pass:
'   vstrXMLRequest
' return:
'   UpdateCustomerRoles
'       xml Response data stream containing results of operation
'       either: TYPE="SUCCESS"
'       or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'----------------------------------------------------------------------------
On Error GoTo UpdateCustomerRolesVbErr
    
    Const strFunctionName As String = "UpdateCustomerRoles"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationManagerBO_UpdateCustomerRoles(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    UpdateCustomerRoles = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateCustomerRolesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateCustomerRolesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateCustomerRoles = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCustomerRolesExit
    
End Function

Public Function DeleteCustomerFromApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Remove the link between a customer and the application
' pass:
'   vstrXMLRequest
' return:
'   DeleteCustomerFromApplication
'       xml Response data stream containing results of operation
'       either: TYPE="SUCCESS"
'       or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerFromApplicationVbErr
    
    Const strFunctionName As String = "DeleteCustomerFromApplication"


    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Validate the Input
    Set xmlTempResponseNode = IApplicationManagerBO_Validate( _
                            xmlIn.documentElement, bomDeleteCustomerFromApplication)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_DeleteCustomerFromApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    DeleteCustomerFromApplication = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteCustomerFromApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteCustomerFromApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteCustomerFromApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCustomerFromApplicationExit

End Function

Private Function IApplicationManagerBO_AddCustomerToApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerToApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_AddCustomerToApplication"

 
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.AddCustomerToApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_AddCustomerToApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
AddCustomerToApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

AddCustomerToApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_AddCustomerToApplication = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AddCustomerToApplicationExit

End Function

Private Function IApplicationManagerBO_CreateCustomerLockApplicationLock( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with the table
'               'CustomerLockApplicationLock'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_CreateCustomerLockApplicationLock"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because so methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String
    
    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_CreateCustomerLockApplicationLock = xmlResponseElem
            GoTo CreateExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_CreateCustomerLockApplicationLock = xmlResponseElem
        GoTo CreateExit
    End If

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.CreateCustomerLockApplicationLock(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_CreateCustomerLockApplicationLock = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationManagerTxBO = Nothing
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

    Set IApplicationManagerBO_CreateCustomerLockApplicationLock = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationManagerBO_CreateCustomerRole( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_CreateCustomerRole"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIApplManagerTxBO.CreateCustomerRole(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_CreateCustomerRole = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplManagerTxBO = Nothing
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

    Set IApplicationManagerBO_CreateCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IApplicationManagerBO_CreateFullApplicationDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_CreateFullApplicationDetails"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.CreateFullApplicationDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_CreateFullApplicationDetails = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationManagerTxBO = Nothing
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

    Set IApplicationManagerBO_CreateFullApplicationDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationManagerBO_CreateLock( _
                ByVal vxmlRequest As IXMLDOMElement, _
                Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with the table
'               'CustomerLockApplicationLock'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_CreateLock"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String
    
    If strLockRequired = "" Then
        If Not IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_CreateLock = xmlResponseElem
            GoTo CreateLockExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_CreateLock = xmlResponseElem
        GoTo CreateLockExit
    End If

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    ' Add LockDateTime and type to the request
    AddLockTypeAndDate vxmlRequest
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.CreateLock(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_CreateLock = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateLockExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateLockVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_CreateLock = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLockExit
End Function

Private Function IApplicationManagerBO_DeleteAllCustomerLockApplicationLock( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_DeleteAllCustomerLockApplicationLock"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String

    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_DeleteAllCustomerLockApplicationLock = _
                                                    xmlResponseElem
            GoTo DeleteExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_DeleteAllCustomerLockApplicationLock = _
                                    xmlResponseElem
        GoTo DeleteExit
    End If


    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.DeleteAllCustomerLockApplicationLock(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_DeleteAllCustomerLockApplicationLock = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_DeleteAllCustomerLockApplicationLock = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit
End Function

Private Function IApplicationManagerBO_DeleteCustomerLockApplicationLock( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the
'               table 'CustomerLockApplicationLock'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_DeleteCustomerLockApplicationLock"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String

    If strLockRequired = "" Then
        If IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_DeleteCustomerLockApplicationLock = xmlResponseElem
            GoTo DeleteExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_DeleteCustomerLockApplicationLock = xmlResponseElem
        GoTo DeleteExit
    End If
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.DeleteCustomerLockApplicationLock(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_DeleteCustomerLockApplicationLock = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_DeleteCustomerLockApplicationLock = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IApplicationManagerBO_DeleteCustomerFromApplication( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerFromApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_DeleteCustomerFromApplication"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.DeleteCustomerFromApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_DeleteCustomerFromApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteCustomerFromApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

DeleteCustomerFromApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_DeleteCustomerFromApplication = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCustomerFromApplicationExit
End Function

Private Function IApplicationManagerBO_DeleteLock( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the
'               table 'ApplicationLock'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLockVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_DeleteLock"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String

    If strLockRequired = "" Then
        If Not IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_DeleteLock = xmlResponseElem
            GoTo DeleteLockExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_DeleteLock = xmlResponseElem
        GoTo DeleteLockExit
    End If
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.DeleteLock(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_DeleteLock = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteLockExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteLockVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_DeleteLock = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteLockExit

End Function

Private Function IApplicationManagerBO_FindCustomerLockApplicationLockList( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'CustomerLockApplicationLock'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_FindCustomerLockApplicationLockList"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIApplicationManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
    Else
        Set objIApplicationManagerDO = _
                        m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTLOCKAPPLLOCK Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTLOCKAPPLLOCK).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTLOCKAPPLLOCK & " tag not found"
    End If
        
    Set xmlDataNode = _
            objIApplicationManagerDO.FindCustomerLockApplicationLockList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationManagerBO_FindCustomerLockApplicationLockList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objIApplicationManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindCustomerLockApplicationLockList = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function IApplicationManagerBO_FindLegacyBusinessForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'CustomerRole'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLegacyBusinessForCustomerVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_FindLegacyBusinessForCustomer"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = _
                        m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMER Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMER).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTOMER & " tag not found"
    End If
        
    Set xmlDataNode = objIApplManagerDO.FindLegacyBusinessForCustomer(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationManagerBO_FindLegacyBusinessForCustomer = xmlResponseElem

FindLegacyBusinessForCustomerExit:

    Set xmlRequestNode = Nothing
    Set objIApplManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindLegacyBusinessForCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindLegacyBusinessForCustomer = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLegacyBusinessForCustomerExit


End Function

Private Function IApplicationManagerBO_GetAdditionalRequestDataForAccDownload(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets data from ApplicationFactFind and Customerversion
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetAddDataVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_GetAdditionalRequestDataForAccDownload"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlAppRequestDoc As FreeThreadedDOMDocument40
    Dim xmlAppRequestNode As IXMLDOMNode
    Dim xmlCustRequestDoc As FreeThreadedDOMDocument40
    Dim xmlCustRequestNode As IXMLDOMNode
    Dim xmlAppDataNode As IXMLDOMNode
    Dim xmlAppDataElement As IXMLDOMElement
    Dim xmlCustDataNode As IXMLDOMNode
    Dim xmlCustNode As IXMLDOMNode
    Dim xmlInCustNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFFNo As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlCustNumberNode As IXMLDOMNode
    Dim xmlCustVersionNode As IXMLDOMNode
    Dim xmlInCustList As IXMLDOMNodeList
    Dim xmlPersonElement As IXMLDOMElement
    Dim sCustomerNumber As String
    Dim sCustomerVersion As String
    
    Dim objIApplicationBO As IApplicationBO
    Dim objICustomerBO As ICustomerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIApplicationBO = New ApplicationBO
        Set objICustomerBO = New CustomerBO
    Else
        Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")
        Set objICustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If
                
    If vxmlRequest.nodeName = cstrGetAddData Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrGetAddData).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrGetAddData & " tag not found"
    End If
        
    'Get the application data
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlRequestNode, ".//APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXmlAssist.GetNode(xmlRequestNode, ".//APPLICATIONFACTFINDNUMBER")
    
    Set xmlAppRequestDoc = New FreeThreadedDOMDocument40
    xmlAppRequestDoc.validateOnParse = False
    xmlAppRequestDoc.setProperty "NewParser", True
    xmlAppRequestDoc.async = False
    Set xmlAppRequestNode = vxmlRequest.cloneNode(False)
    xmlAppRequestDoc.appendChild xmlAppRequestNode
    
    Set xmlTableElem = xmlAppRequestDoc.createElement("APPLICATION")
    xmlTableElem.appendChild xmlAppNo.cloneNode(True)
    xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    xmlAppRequestNode.appendChild xmlTableElem
    
    Set xmlAppDataElement = objIApplicationBO.GetApplicationData(xmlAppRequestDoc.documentElement)
    m_objErrAssist.CheckResponse xmlAppDataElement, xmlResponseElem, True
    
    'build the response
    Set xmlAppDataNode = xmlOut.createElement("APPLICATION")
    Set xmlTempNode = m_objXmlAssist.GetNode(xmlAppDataElement, ".//SPECIALSCHEME")
    xmlAppDataNode.appendChild xmlTempNode.cloneNode(True)
    Set xmlTempNode = m_objXmlAssist.GetNode(xmlAppDataElement, ".//MAINRESIDENCEIND")
    xmlAppDataNode.appendChild xmlTempNode.cloneNode(True)
    
    'build a request to send to GetPersonalDetails
    Set xmlCustRequestDoc = New FreeThreadedDOMDocument40
    xmlCustRequestDoc.validateOnParse = False
    xmlCustRequestDoc.setProperty "NewParser", True
    xmlCustRequestDoc.async = False
    Set xmlCustRequestNode = vxmlRequest.cloneNode(False)
    xmlCustRequestDoc.appendChild xmlCustRequestNode
    Set xmlTableElem = xmlCustRequestDoc.createElement("CUSTOMER")
    xmlCustRequestNode.appendChild xmlTableElem
    Set xmlCustNumberNode = xmlCustRequestDoc.createElement("CUSTOMERNUMBER")
    xmlTableElem.appendChild xmlCustNumberNode
    Set xmlCustVersionNode = xmlCustRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlTableElem.appendChild xmlCustVersionNode
    
    Set xmlCustDataNode = xmlOut.createElement("CUSTOMERLIST")
    Set xmlInCustList = xmlRequestNode.selectNodes(".//CUSTOMER")
    For Each xmlInCustNode In xmlInCustList
        
        sCustomerNumber = m_objXmlAssist.GetNodeValue(xmlInCustNode, "CUSTOMERNUMBER")
        sCustomerVersion = m_objXmlAssist.GetNodeValue(xmlInCustNode, "CUSTOMERVERSIONNUMBER")
        
        Set xmlCustNode = xmlOut.createElement("CUSTOMER")
        Set xmlTempElem = xmlOut.createElement("CUSTOMERNUMBER")
        xmlTempElem.Text = sCustomerNumber
        xmlCustNode.appendChild xmlTempElem
        Set xmlTempElem = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlTempElem.Text = sCustomerVersion
        xmlCustNode.appendChild xmlTempElem
        
        'Fill in the details to send to GetPersonalDetails
        xmlCustNumberNode.Text = sCustomerNumber
        xmlCustVersionNode.Text = sCustomerVersion
        
        Set xmlPersonElement = objICustomerBO.GetPersonalDetails(xmlCustRequestDoc.documentElement)
        m_objErrAssist.CheckResponse xmlPersonElement, xmlResponseElem, True
        
        Set xmlTempNode = m_objXmlAssist.GetNode(xmlPersonElement, ".//OWNMORTGAGELENDERIND")
        xmlCustNode.appendChild xmlTempNode.cloneNode(True)
        Set xmlTempNode = m_objXmlAssist.GetNode(xmlPersonElement, ".//MORTGAGEREPAID")
        xmlCustNode.appendChild xmlTempNode.cloneNode(True)
        Set xmlTempNode = m_objXmlAssist.GetNode(xmlPersonElement, ".//MORTGAGELET")
        xmlCustNode.appendChild xmlTempNode.cloneNode(True)
        Set xmlTempNode = m_objXmlAssist.GetNode(xmlPersonElement, ".//AXESSAPPLICATION")
        xmlCustNode.appendChild xmlTempNode.cloneNode(True)
        
        
        
        xmlCustDataNode.appendChild xmlCustNode
    Next
    
    
    
        
        
    xmlResponseElem.appendChild xmlAppDataNode
    xmlResponseElem.appendChild xmlCustDataNode
        
    Set IApplicationManagerBO_GetAdditionalRequestDataForAccDownload = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetAddDataExit:
    
    Set xmlRequestNode = Nothing
    Set objIApplicationBO = Nothing
    Set objICustomerBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetAddDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_GetAdditionalRequestDataForAccDownload = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetAddDataExit

End Function

Private Function IApplicationManagerBO_GetCustomerRoleData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with the table
'               CustomerRole
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_GetCustomerRoleData"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = _
                m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERROLE Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERROLE).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTOMERROLE & " tag not found"
    End If
        
    Set xmlDataNode = objIApplManagerDO.GetCustomerRoleData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationManagerBO_GetCustomerRoleData = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetDataExit:
    
    Set xmlRequestNode = Nothing
    Set objIApplManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_GetCustomerRoleData = _
                                m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function IApplicationManagerBO_GetLockData( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with the table
'               'ApplicationLock'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLockDataVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_GetLockData"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIApplicationManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIApplicationManagerDO = New ApplicationManagerDO
    Else
        Set objIApplicationManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_APPLLOCK Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_APPLLOCK).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_APPLLOCK & " tag not found"
    End If
        
    Set xmlDataNode = objIApplicationManagerDO.GetLockData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IApplicationManagerBO_GetLockData = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetLockDataExit:
    
    Set xmlRequestNode = Nothing
    Set objIApplicationManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetLockDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_GetLockData = m_objErrAssist.CreateErrorResponseEx()

'    If Not m_objContext Is Nothing Then
'        m_objContext.setcomplete
'    End If
    
    ' go to clean-up section
    Resume GetLockDataExit

End Function

Private Function IApplicationManagerBO_ImportAccountRefreshIntoApplication( _
                                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ImportAccountRefreshIntoApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_ImportAccountRefreshIntoApplication"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    
    Set xmlTempResponseNode = objIApplicationManagerTxBO.ImportAccountRefreshIntoApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode  'JD BMIDS975
    
    Set IApplicationManagerBO_ImportAccountRefreshIntoApplication = xmlResponseElem

    m_objContext.SetComplete
    
ImportAccountRefreshIntoApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ImportAccountRefreshIntoApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_ImportAccountRefreshIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountRefreshIntoApplicationExit

End Function
    

Private Function IApplicationManagerBO_ImportCustomersIntoApplication(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode

On Error GoTo ImportCustomersIntoApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_ImportCustomersIntoApplication"
        
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objConvAssist As ConvertAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objConvAssist = New ConvertAssist
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")

    Dim strAdminSystem As String
    strAdminSystem = objGlobalParamDO.GetCurrentParameterByType("FindLegacyCustomer", "BOOLEAN")
    
    ' Check if we have an administration system
    If objConvAssist.CSafeBool(strAdminSystem) = True Then
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
        
        '   call Business Transaction Object ImportCustomersIntoApplication function
        Set xmlTempResponseNode = objIApplicationManagerTxBO.ImportCustomersIntoApplication(vxmlRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    End If
   
    Set IApplicationManagerBO_ImportCustomersIntoApplication = xmlResponseElem

    m_objContext.SetComplete
    
ImportCustomersIntoApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing
    

    Exit Function

ImportCustomersIntoApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_ImportCustomersIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportCustomersIntoApplicationExit


End Function

Private Function IApplicationManagerBO_LockCustomersForApplication( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo LockCustomersForApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_LockCustomersForApplication"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    Dim blnReadOnly As Boolean
    Dim xmlMessageElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String

    If strLockRequired = "" Then
        If Not IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_LockCustomersForApplication = xmlResponseElem
            GoTo LockCustomersForApplicationExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_LockCustomersForApplication = xmlResponseElem
        GoTo LockCustomersForApplicationExit
    End If
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object's function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.LockCustomersForApplication(vxmlRequest)
     
    On Error Resume Next
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' APS 02/09/99 - ReadOnly processing. We cannot return an error if the case
    ' is already locked. Instead we should return success and set the READONLY
    ' attribute of the RESPONSE to "1" to indicate readonly mode
    
    Dim lngErrNumber As Long
    Dim strErrDescription As String
    
    ' save the error
    If Err.Number <> 0 Then
        lngErrNumber = m_objErrAssist.GetOmigaErrorNumber(Err.Number)
    End If
    strErrDescription = Err.Description
    
    ' clear the error object
    Err.Clear
    
    ' reinstate the error handler
    On Error GoTo LockCustomersForApplicationVbErr
    
    If (lngErrNumber = 182 Or _
        lngErrNumber = 183 Or _
        lngErrNumber = 190 Or _
        lngErrNumber = 7013 Or _
        lngErrNumber = 7014) Then 'BM0202 21/05/03 LDM show user and unit of who has locked the app
        blnReadOnly = True
    ElseIf lngErrNumber <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngErrNumber
    End If
    
    If blnReadOnly Then
        xmlResponseElem.setAttribute "READONLY", "1"
        Set xmlMessageElem = xmlOut.createElement("MESSAGE")
        xmlMessageElem.Text = strErrDescription
        xmlResponseElem.appendChild xmlMessageElem
    End If
     
    Set IApplicationManagerBO_LockCustomersForApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

LockCustomersForApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

LockCustomersForApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_LockCustomersForApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume LockCustomersForApplicationExit

End Function

Private Function IApplicationManagerBO_MaintainCustomerRoles( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo MaintainCustomerRolesVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_MaintainCustomerRoles"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.MaintainCustomerRoles(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_MaintainCustomerRoles = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

MaintainCustomerRolesExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

MaintainCustomerRolesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_MaintainCustomerRoles = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume MaintainCustomerRolesExit

End Function

Private Function IApplicationManagerBO_SaveCustomerAndEmploymentDetails( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo SaveVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_SaveCustomerAndEmploymentDetails"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplMgrTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplMgrTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplMgrTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplMgrTxBO.SaveCustomerAndEmploymentDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_SaveCustomerAndEmploymentDetails = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveVbExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIApplMgrTxBO = Nothing
    
    Exit Function

SaveVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_SaveCustomerAndEmploymentDetails = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveVbExit
    
End Function

Private Function IApplicationManagerBO_UnlockApplicationAndCustomers( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UnlockApplicationAndCustomersVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_UnlockApplicationAndCustomers"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '---------------------------------------------------------------------------------------
    ' Check whether locking is available for this system, if so proceed else exit the method
    ' This is done in DOM Interface method also, because some methods might directly call this,
    ' bypassing the string interface.
    '---------------------------------------------------------------------------------------
    Dim strLockingReqd_Local As String

    If strLockRequired = "" Then
        If Not IsLockRequiredInThisSystem Then
            Set IApplicationManagerBO_UnlockApplicationAndCustomers = xmlResponseElem
            GoTo UnlockApplicationAndCustomersExit
        End If
    ElseIf strLockRequired = "0" Then
        Set IApplicationManagerBO_UnlockApplicationAndCustomers = xmlResponseElem
        GoTo UnlockApplicationAndCustomersExit
    End If

    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance( _
                                    App.Title & ".ApplicationManagerTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
            objIApplicationManagerTxBO.UnlockApplicationAndCustomers(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IApplicationManagerBO_UnlockApplicationAndCustomers = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UnlockApplicationAndCustomersExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UnlockApplicationAndCustomersVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_UnlockApplicationAndCustomers = _
                                      m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UnlockApplicationAndCustomersExit
End Function

Private Function IApplicationManagerBO_UpdateCustomerRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with the
'               table 'CustomerRole'
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_UpdateCustomerRole"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplManagerTxBO = _
                m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplManagerTxBO.UpdateCustomerRole(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_UpdateCustomerRole = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:

    Set objIApplManagerTxBO = Nothing
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

    Set IApplicationManagerBO_UpdateCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IApplicationManagerBO_UpdateCustomerRoles(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerRolesVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_UpdateCustomerRoles"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object's function
    Set xmlTempResponseNode = objIApplicationManagerTxBO.UpdateCustomerRoles(vxmlRequest)
     m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_UpdateCustomerRoles = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateCustomerRolesExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateCustomerRolesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_UpdateCustomerRoles = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCustomerRolesExit

End Function

Private Function IApplicationManagerBO_Validate( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal veboMethod As IApplicationManagerBOMethod _
                            = bomApplicationManagerDefault) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IApplicationManagerBO_Validate"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIApplicationManagerDO As IApplicationManagerDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Select Case veboMethod
        Case bomCreateCustomerLockApplicationLock
            Call ValidateCustomerLockApplicationLock(vxmlRequest)
        Case IApplicationManagerBOMethod.bomAddCustomerToApplication
            Call ValidateAddCustomerToApplication(vxmlRequest)
        Case IApplicationManagerBOMethod.bomDeleteCustomerFromApplication
            Call ValidateDeleteCustomerFromApplication(vxmlRequest)
        Case IApplicationManagerBOMethod.bomCreateFullApplicationDetails
            Call ValidateFullApplicationDetails(vxmlRequest)
        Case Default
            ' do nothing
    End Select
    
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IApplicationManagerBO_Validate = xmlResponseElem
    
ValidateExit:

    Set objIApplicationManagerDO = Nothing
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

    Set IApplicationManagerBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit


End Function

Public Function CreateCustomerRole(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateCustomerRole"

        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, _
                                      IApplicationManagerBOMethod.bomCreateCustomerRole)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
        Set xmlTempResponseNode = _
            IApplicationManagerBO_CreateCustomerRole(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateCustomerRole = xmlResponseElem.xml
    
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

    CreateCustomerRole = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function DeleteCustomerRole(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the table
'               'CustomerRole'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "DeleteCustomerRole"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationManagerBO_DeleteCustomerRole(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteCustomerRole = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteCustomerRole = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IApplicationManagerBO_DeleteCustomerRole( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_DeleteCustomerRole"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplManagerTxBO As IApplicationManagerTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplManagerTxBO = New ApplicationManagerTxBO
    Else
        Set objIApplManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
    End If
           
    '   call Business Transaction Object Delete function
    Set xmlTempResponseNode = objIApplManagerTxBO.DeleteCustomerRole(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationManagerBO_DeleteCustomerRole = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteExit:

    Set objIApplManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_DeleteCustomerRole = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Public Function UpdateCustomerRole(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with the
'               table 'CustomerRole'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "UpdateCustomerRole"

    
    Dim xmlIn As FreeThreadedDOMDocument40
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
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, _
                                        IApplicationManagerBOMethod.bomUpdateCustomerRole)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationManagerBO_UpdateCustomerRole(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateCustomerRole = xmlResponseElem.xml
    
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

    UpdateCustomerRole = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
    
End Function

Public Function GetCustomerRoleData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with the
'               tabel 'CustomerRole'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetCustomerRoleData"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = _
                IApplicationManagerBO_GetCustomerRoleData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetCustomerRoleData = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetCustomerRoleData = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function FindCustomerRoleList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'CustomerRole'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindCustomerRoleList"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = _
            IApplicationManagerBO_FindCustomerRoleList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindCustomerRoleList = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindCustomerRoleList = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function IApplicationManagerBO_FindCustomerRoleList( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'CustomerRole'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_FindCustomerRoleList"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = _
                        m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERROLE Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERROLE).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTOMERROLE & " tag not found"
    End If
        
    Set xmlDataNode = objIApplManagerDO.FindCustomerRoleList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationManagerBO_FindCustomerRoleList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objIApplManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindCustomerRoleList = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Sub ValidateDeleteCustomerFromApplication(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vXMLRequest  xml Request node containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateDeleteCustomerFromApplication"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objGlobalParameterBO As GlobalParameterBO
    Dim objComboDO As ComboDO
    Dim objIDOAssist As IDOAssist
    Dim objIApplClassDef As IApplicationClassDef
    
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlAFFCounts As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objConvertAssist As ConvertAssist
    
    If objContext Is Nothing Then
        Set objGlobalParameterBO = New GlobalParameterBO
        Set objComboDO = New ComboDO
        Set objIDOAssist = New DOAssist
    Else
        Set objGlobalParameterBO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objComboDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    '----------------------------------------------------------------------------------
    ' If legacy system is available and the application is further advance, raise error
    '----------------------------------------------------------------------------------
    Dim strXmlResponse As String, strLegacyCustomer As String
    
    strXmlResponse = objGlobalParameterBO.GetCurrentParameter("FindLegacyCustomer")
    Set xmlDoc = m_objXmlAssist.load(strXmlResponse, TypeName(Me), strFunctionName)
    strLegacyCustomer = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "BOOLEAN")
    
    ' Indicates whether the interface to the organisation's legacy systems is available
    If strLegacyCustomer = "1" Then
        Dim strTypeOfApp As String
        strTypeOfApp = m_objXmlAssist.GetTagValue(vxmlRequest, "TYPEOFAPPLICATION")
        'BMIDS01010 Throw an error if TypeOfMortgage has validation types F and M, but not S or P
        'If (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "F")) Then
        If (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "F") And objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "M")) Then
            If Not (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "S") Or objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "P")) Then
        'BMIDS01010 End
                '"A customer may not be reordered within, assigned to, or removed from this type of mortgage application"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 168
            End If 'BMIDS01010
        End If
    End If
    
    '-----------------------------------------------------------------------------------------
    ' Check that the number of applicants is atleast 1 after this transaction.
    ' Add the new number of applicants and guarantors to the  request tag, so that it can be
    ' used further for updation.
    '-----------------------------------------------------------------------------------------
    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
    Dim intNumApplicants As Integer, intNumGuarantors As Integer
    
    Dim strCustomerRoleType As String
    
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & "APPLICATIONNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & "APPLICATIONFACTFINDNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData(AFFApplicantCounts)
    Set xmlAFFCounts = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
    
    Set objConvertAssist = New ConvertAssist
    intNumApplicants = objConvertAssist.CSafeInt(xmlAFFCounts.selectSingleNode("NUMBEROFAPPLICANTS").Text)
    intNumGuarantors = objConvertAssist.CSafeInt(xmlAFFCounts.selectSingleNode("NUMBEROFGUARANTORS").Text)
    
    strCustomerRoleType = vxmlRequest.selectSingleNode(".//CUSTOMERROLETYPE").Text
    
    ' Raise error, if number of applicants falls below 1 due to this transaction
    If strCustomerRoleType = "1" Then
        If intNumApplicants <= 1 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 173
        End If
        Set xmlElement = xmlApplClassDefDoc.createElement("NUMBEROFAPPLICANTS")
        xmlElement.Text = intNumApplicants - 1
        vxmlRequest.appendChild xmlElement
    ElseIf strCustomerRoleType = "2" Then
        If intNumGuarantors < 1 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                                    "Number of Guarantors cannot be less than zero"
        End If
        Set xmlElement = xmlApplClassDefDoc.createElement("NUMBEROFGUARANTORS")
        xmlElement.Text = intNumGuarantors - 1
        vxmlRequest.appendChild xmlElement
    End If

    Set xmlApplClassDefDoc = Nothing
    Set xmlAFFCounts = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    
    Set objGlobalParameterBO = Nothing
    Set objComboDO = Nothing
    Set objIDOAssist = Nothing
    Set objConvertAssist = Nothing
    
Exit Sub

ValidateVbErr:
    
    Set xmlApplClassDefDoc = Nothing
    Set xmlAFFCounts = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    
    Set objGlobalParameterBO = Nothing
    Set objComboDO = Nothing
    Set objIDOAssist = Nothing
    Set objConvertAssist = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ValidateAddCustomerToApplication(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vXMLRequest  xml Request node containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateAddCustomerToApplication"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objGlobalParameterBO As GlobalParameterBO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objComboDO As ComboDO
    Dim objApplDO As IApplicationDO
    Dim objIDOAssist As IDOAssist
    Dim objIApplClassDef As IApplicationClassDef
    
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlAFFCounts As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objConvertAssist As ConvertAssist
    
    If objContext Is Nothing Then
        Set objGlobalParameterBO = New GlobalParameterBO
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objComboDO = New ComboDO
        Set objApplDO = New ApplicationDO
        Set objIDOAssist = New DOAssist
    Else
        Set objGlobalParameterBO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objGlobalParameterDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objComboDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objApplDO = objContext.CreateInstance(cstrCOMPONENT_NAME & ".ApplicationDO")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'BMIDS00434 This validation has been moved to CR040 as it should only be treated as a warning
    ''----------------------------------------------------------------------------------
    '' If legacy system is available and the application is further advance, raise error
    ''----------------------------------------------------------------------------------
    'Dim strXmlResponse As String, strLegacyCustomer As String
    '
    'strXmlResponse = objGlobalParameterBO.GetCurrentParameter("FindLegacyCustomer")
    'Set xmlDoc = m_objXmlAssist.load(strXmlResponse, TypeName(Me), strFunctionName)
    'strLegacyCustomer = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "BOOLEAN")
    '
    '' Indicates whether the interface to the organisation's legacy systems is available
    'If strLegacyCustomer = "1" Then
    '    Dim strTypeOfApp As String
    '    strTypeOfApp = m_objXmlAssist.GetTagValue(vxmlRequest, "TYPEOFAPPLICATION")
    '    If (objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApp, "F")) Then
    '        '"A customer may not be reordered within, assigned to, or removed from this type of mortgage application"
    '        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 168
    '    End If
    'End If
    'BMIDS00434 End
    
    '-----------------------------------------------------------------------------------------
    ' Check that the number of applicants and the guarantors will not exceed the maximum limit
    ' after this transaction. Add the new number of applicants and guarantors to the  request
    ' tag, so that it can be used further for updation.
    '-----------------------------------------------------------------------------------------
    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
    Dim intNumApplicants As Integer, intNumGuarantors As Integer
    Dim intMaxApplicants As Integer, intMaxGuarantors As Integer
    
    Dim strCustomerRoleType As String
    
    Set xmlInputElem = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & "APPLICATIONNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & "APPLICATIONFACTFINDNUMBER").Text
    xmlTableElem.appendChild xmlElement
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData(AFFApplicantCounts)
    Set xmlAFFCounts = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
    
    Set objConvertAssist = New ConvertAssist
    intNumApplicants = objConvertAssist.CSafeInt(xmlAFFCounts.selectSingleNode("NUMBEROFAPPLICANTS").Text)
    intNumGuarantors = objConvertAssist.CSafeInt(xmlAFFCounts.selectSingleNode("NUMBEROFGUARANTORS").Text)
    
    strCustomerRoleType = vxmlRequest.selectSingleNode(".//CUSTOMERROLETYPE").Text
    
    Dim strResponseXml As String
    ' Raise error, if number of applicants or guarantors exceeds maximum and
    ' the node to request
    If strCustomerRoleType = "1" Then
        strResponseXml = objGlobalParameterDO.GetCurrentParameter("MaximumApplicants")
        Set xmlDoc = m_objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
        intMaxApplicants = _
            objConvertAssist.CSafeInt(m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT"))
        If intNumApplicants + 1 > intMaxApplicants Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 540
        End If
        Set xmlElement = xmlApplClassDefDoc.createElement("NUMBEROFAPPLICANTS")
        xmlElement.Text = intNumApplicants + 1
        vxmlRequest.appendChild xmlElement
    Else
        strResponseXml = objGlobalParameterDO.GetCurrentParameter("MaximumGuarantors")
        Set xmlDoc = m_objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
        intMaxGuarantors = _
            objConvertAssist.CSafeInt(m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT"))
        If intNumGuarantors + 1 > intMaxGuarantors Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 542
        End If
        Set xmlElement = xmlApplClassDefDoc.createElement("NUMBEROFGUARANTORS")
        xmlElement.Text = intNumGuarantors + 1
        vxmlRequest.appendChild xmlElement
    End If

    Set xmlApplClassDefDoc = Nothing
    Set xmlAFFCounts = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    
    Set objGlobalParameterDO = Nothing
    Set objGlobalParameterBO = Nothing
    Set objComboDO = Nothing
    Set objApplDO = Nothing
    Set objIDOAssist = Nothing
    
Exit Sub

ValidateVbErr:
    
    Set xmlApplClassDefDoc = Nothing
    Set xmlAFFCounts = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    
    Set objGlobalParameterBO = Nothing
    Set objComboDO = Nothing
    Set objApplDO = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ValidateFullApplicationDetails( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
' pass:         vXMLRequest  xml Request node containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateFullApplicationDetails"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
       
    Dim objIApplicationBO As IApplicationBO
    Dim objINewPropertyBO As INewPropertyBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIApplicationBO = New ApplicationBO
    Set objINewPropertyBO = New NewPropertyBO
    
    Set xmlTempResponseNode = objIApplicationBO.ValidateApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIApplicationBO.ValidateFinancialSummary(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlTempResponseNode = objINewPropertyBO.ValidateNewPropertyDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlTempResponseNode = objINewPropertyBO.ValidateNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ValidateFullApplicationDetails = xmlResponseElem
    
ValidateExit:
    Set objIApplicationBO = Nothing
    Set objINewPropertyBO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOut = Nothing

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

    ValidateFullApplicationDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit
    
End Function

Public Function UnlockApplicationAndCustomers(ByVal vstrXMLRequest As String) As String
' header ------------------------------------------------------------------------------
' description:
'   Remove all the locks on application and customers linked to the application
' pass:
'   vstrXMLRequest
' return:
'
' Raise Errors:
'-------------------------------------------------------------------------------------
On Error GoTo UnlockApplicationAndCustomersVbErr
    
    Const strFunctionName As String = "UnlockApplicationAndCustomers"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '------------------------------------------------------
    ' Check whether locking is available for this system.
    ' If required proceed else exit the method.
    '------------------------------------------------------
    Dim strLockingReqd As String
    
    If IsLockRequiredInThisSystem() Then
        strLockingReqd = 1
    Else
        UnlockApplicationAndCustomers = xmlResponseElem.xml
        GoTo UnlockApplicationAndCustomersExit
    End If

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                               
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
      IApplicationManagerBO_UnlockApplicationAndCustomers(xmlIn.documentElement, strLockingReqd)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UnlockApplicationAndCustomers = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UnlockApplicationAndCustomersExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UnlockApplicationAndCustomersVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UnlockApplicationAndCustomers = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UnlockApplicationAndCustomersExit
    
End Function

Private Sub AddLockTypeAndDate(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Adds the Lock Type and Lock Date to the request
' pass:         vstrXMLRequest  xml Request
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddLockTypeAndDateErr
    
    Const strFunctionName As String = "AddLockTypeAndDate"

    Dim objLockElement As IXMLDOMElement
    Dim objLockDateElem As IXMLDOMElement
    Dim objLockTypeElem As IXMLDOMElement
                    
    Dim intNoOfLocks As Integer
    Dim intLockIndex As Integer
    
    intNoOfLocks = vxmlRequest.getElementsByTagName(cstrTABLE_APPLLOCK).length
    
    For intLockIndex = 0 To intNoOfLocks - 1
        Set objLockElement = vxmlRequest.getElementsByTagName(cstrTABLE_APPLLOCK).Item(intLockIndex)
        Set objLockDateElem = objLockElement.getElementsByTagName("LOCKDATE").Item(0)

        If objLockDateElem Is Nothing Then
            Set objLockDateElem = vxmlRequest.ownerDocument.createElement("LOCKDATE")
            objLockElement.appendChild objLockDateElem
        End If
        
        ' AQR SYS0059 - Date&Time should be stored in database
        objLockDateElem.Text = Format(Now, "dd/mm/yyyy HH:MM:SS")
    
        Set objLockTypeElem = objLockElement.getElementsByTagName("TYPEOFLOCK").Item(0)
    
        If objLockTypeElem Is Nothing Then
            Set objLockTypeElem = vxmlRequest.ownerDocument.createElement("TYPEOFLOCK")
            objLockElement.appendChild objLockTypeElem
        End If
        
        If Len(objLockTypeElem.Text) = 0 Then
            objLockTypeElem.Text = "On"
        End If
    Next
    
    Set objLockElement = Nothing
    Set objLockDateElem = Nothing
    Set objLockTypeElem = Nothing

    Exit Sub

AddLockTypeAndDateErr:
    
    Set objLockElement = Nothing
    Set objLockDateElem = Nothing
    Set objLockTypeElem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IsLockRequiredInThisSystem() As Boolean
'---------------------------------------------------------------
'This checks whether locking is required for this system.
'Return TRUE, if required else FALSE
'--------------------------------------------------------------
On Error GoTo IsLockRequiredInThisSystemVbErr

    Const strFunctionName As String = "IsLockRequiredInThisSystem"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
                m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    Dim strLockingReqd As String
    strLockingReqd = objGlobalParamDO.GetCurrentParameterByType("Locking", "BOOLEAN")
    
    IsLockRequiredInThisSystem = IIf(strLockingReqd = "0", False, True)
    
    Set objGlobalParamDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Set objErrAssist = Nothing
    
    Exit Function
IsLockRequiredInThisSystemVbErr:
    objErrAssist.AddToErrSource strFunctionName
    
    Set objGlobalParamDO = Nothing
    Set objErrAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Function CreateFullApplicationDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'    <REQUEST USERID = ""USER0001"" UNITID=""UNIT1"" MACHINEID=""MSGETC"">
'    <APPLICATION>
'       <APPLICATIONNUMBER></APPLICATIONNUMBER>
'       <PACKAGENUMBER>PACK0002</PACKAGENUMBER>
'       <USERID>Srini</USERID>
'       <UNITID>UNIT1</UNITID>
'       <DISTRIBUTIONCHANNELID>500</DISTRIBUTIONCHANNELID>
'       <AMOUNTREQUESTED>500</AMOUNTREQUESTED>
'
'       <APPLICATIONVERIFICATION>
'           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'           <RESIDENCYOKINDICATOR>1</RESIDENCYOKINDICATOR>
'           <AGEOKINDICATOR>1</AGEOKINDICATOR>
'           ..........
'           ...........
'       </APPLICATIONVERIFICATION>
'
'       <FINANCIALSUMMARY>
'           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'           <ARREARSHISTORYINDICATOR>1</ARREARSHISTORYINDICATOR>
'           <BANKCARDINDICATOR>1</BANKCARDINDICATOR>
'           ...................
'           ..........
'       </FINANCIALSUMMARY>
'
'       <NEWPROPERTY>
'           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'           ........
'           .......
'       </NEWPROPERTY>
'
'       <NEWPROPERTYADDRESS>
'           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'           <ADDRESSGUID></ADDRESSGUID>
'           <ARRANGEMENTSFORACCESS></ARRANGEMENTSFORACCESS>
'           <OTHERARRANGEMENTSFORACCESS></OTHERARRANGEMENTSFORACCESS>
'           <ADDRESS>
'               <ADDRESSGUID></ADDRESSGUID>
'               <BUILDINGORHOUSENAME></BUILDINGORHOUSENAME>
'               <BUILDINGORHOUSENUMBER></BUILDINGORHOUSENUMBER>
'               .........
'               ........
'           </ADDRESS>
'       </NEWPROPERTYADDRESS>
'
'    </APPLICATION>
'    </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateFullApplicationDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, _
                                    IApplicationManagerBOMethod.bomCreateFullApplicationDetails)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_CreateFullApplicationDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateFullApplicationDetails = xmlTempResponseNode.xml
    
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

    CreateFullApplicationDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

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

Private Function IsServiceAvailable(vstServiceType As SERVICE_TYPE) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check against GlobalParameter table whether a particular service is available.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo IsServiceAvailableVbErr
    
    Const strFunctionName As String = "IsServiceAvailable"

    
    Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim strParamName As String
    
    Select Case vstServiceType
    Case stCREDIT_CHECK
        strParamName = "CreditCheckAvailable"
    Case stRISK_ASSESSMENT
        strParamName = "RiskAssessmentAvailable"
    Case stOM4TO3DOWNLOAD
        strParamName = "Omiga4To3DownloadAvailable"
    Case Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeNotImplemented, "Specified service type not handled"
    End Select
    
    IsServiceAvailable = objGlobalParamDO.GetCurrentParameterByType( _
        strParamName, "BOOLEAN")
    
    Set objGlobalParamDO = Nothing
    
    Exit Function
    
IsServiceAvailableVbErr:
    
    Set objGlobalParamDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function LogSubmissionError( _
    ByVal vstrDescription As String, _
    ByVal vstServiceType As SERVICE_TYPE)
' header ----------------------------------------------------------------------------------
' description:
'   Log a Submission process error.
' pass:
' return:   n/a
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "LogSubmissionError"

    
    Dim strMsg As String
    Dim strSource As String
    
    Select Case vstServiceType
    Case stCREDIT_CHECK
        strSource = "Credit Check"
    Case stRISK_ASSESSMENT
        strSource = "Risk Assessment"
    Case stOM4TO3DOWNLOAD
        strSource = "Om4To3Download"
    Case Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeNotImplemented, "Specified service type not handled"
    End Select
    
    strSource = App.Title & ".SubmitBO." & strFunctionName
    
    strMsg = _
        strSource & " Submission Failure - " & vbCr & _
        "Details: " & vstrDescription & vbCr
        
    App.LogEvent strMsg, vbLogEventTypeError
    
    ' FIXME Remove this eventually
    Debug.Print strMsg

End Function

Private Function GetSubmissionStatusComboId(ByVal vstrValidation As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the combo id for supplied combo validation value from combo "SubmissionStatus".
' pass:
'   vstrValidation
'       combo validation value.
' return:
'   Combo Id value.
'------------------------------------------------------------------------------------------
On Error GoTo GetSubmissionStatusComboIdVbErr
    
    Const strFunctionName As String = "GetSubmissionStatusComboId"

    
    Dim objComboDO As ComboDO
    
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    GetSubmissionStatusComboId = objComboDO.GetFirstComboValueId( _
        "SubmissionStatus", vstrValidation)
    
    Set objComboDO = Nothing
    
    Exit Function
    
GetSubmissionStatusComboIdVbErr:
    
    Set objComboDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Function FindLegacyBusinessForCustomer(ByVal vstrXMLRequest As String) As String
On Error GoTo FindLegacyBusinessForCustomerVbErr

    Const strFunctionName As String = "FindLegacyBusinessForCustomer"
        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_FindLegacyBusinessForCustomer(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    FindLegacyBusinessForCustomer = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindLegacyBusinessForCustomerExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindLegacyBusinessForCustomerVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindLegacyBusinessForCustomer = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindLegacyBusinessForCustomerExit


End Function

Public Function SaveCustomerAndEmploymentDetails(ByVal vstrXMLRequest As String) As String

On Error GoTo SaveCustomerAndEmploymentDetailsVbErr

    Const strFunctionName As String = "SaveCustomerAndEmploymentDetails"

        
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_SaveCustomerAndEmploymentDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveCustomerAndEmploymentDetails = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveCustomerAndEmploymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveCustomerAndEmploymentDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveCustomerAndEmploymentDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveCustomerAndEmploymentDetailsExit

End Function

Function FindApplicationList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindApplicationListVbErr

    Const strFunctionName As String = "FindApplicationList"
        
    
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_FindApplicationList(xmlIn.documentElement)
            
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    FindApplicationList = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindApplicationListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindApplicationListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindApplicationList = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationListExit


End Function

Private Function IApplicationManagerBO_FindApplicationList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  pass on the call to ApplicationmanagerDO.FindApplicationList
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationListVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_FindApplicationList"



    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIApplManagerDO As IApplicationManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
    Else
        Set objIApplManagerDO = _
                        m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    End If
                        
    Set xmlDataNode = objIApplManagerDO.FindApplicationList(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationManagerBO_FindApplicationList = xmlResponseElem

FindApplicationListExit:

    Set objIApplManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindApplicationListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindApplicationList = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationListExit

End Function

Function FindApplicationOwnershipList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindApplicationOwnershipListVbErr

    Const strFunctionName As String = "FindApplicationOwnershipList"
        
    
    Dim xmlIn As FreeThreadedDOMDocument40
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

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
            IApplicationManagerBO_FindApplicationOwnershipList(xmlIn.documentElement)
            
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    FindApplicationOwnershipList = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindApplicationOwnershipListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindApplicationOwnershipListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindApplicationOwnershipList = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationOwnershipListExit


End Function

Private Function IApplicationManagerBO_FindApplicationOwnershipList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Returns all application owners for the application number passed in. Needs
'               to call Organisation to obtain user/unit details, then joins that with the
'               user history data.
' pass:         vxmlRequest  xml Request data stream containing the application number
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationOwnershipListVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_FindApplicationOwnershipList"
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim strAppNo As String
    Dim objIOrganisationBO As IOrganisationBO
    Dim objIApplicationBO As IApplicationBO
    Dim xmlUsers As IXMLDOMNode
    Dim xmlUnits As IXMLDOMNode
    Dim xmlUserHistory As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
    ' Get search element, which is the application number
    strAppNo = m_objXmlAssist.GetNodeValue(vxmlRequest, "./APPLICATIONOWNERSHIP/APPLICATIONNUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingElement, "/APPLICATIONOWNERSHIP/APPLICATIONNUMBER"
    End If
    
    Set objIOrganisationBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Create UserHistory request, and get user history details
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlUserHistoryReq As IXMLDOMNode
    Dim xmlUserHistoryResp As IXMLDOMElement
    Dim xmlDom As FreeThreadedDOMDocument40
    
    Set xmlUserHistoryReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDom = xmlUserHistoryReq.ownerDocument
    Set xmlUserHistoryReq = xmlUserHistoryReq.appendChild(xmlDom.createElement("USERHISTORY"))
    Set xmlUserHistoryReq = xmlUserHistoryReq.appendChild(xmlDom.createElement("APPLICATIONNUMBER"))

    xmlUserHistoryReq.Text = strAppNo

    
    Set xmlUserHistoryResp = objIApplicationBO.FindUserHistoryList(xmlDom.documentElement)
    m_objErrAssist.CheckResponse xmlUserHistoryResp, xmlResponseElem, True
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Create UnitName request, and get Unit details
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlUnitReq As IXMLDOMNode
    Dim xmlUnitNode As IXMLDOMNode
    Dim xmlUnitResp As IXMLDOMNode
    
    Set xmlUnitReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDom = xmlUnitReq.ownerDocument
    Set xmlUnitNode = xmlUnitReq.appendChild(xmlDom.createElement("UNITLIST"))
    
    BuildNode "UNITID", xmlUserHistoryResp, xmlUnitNode
    
    Set xmlUnitResp = objIOrganisationBO.FindUnitNameList(xmlUnitReq)
    m_objErrAssist.CheckResponse xmlUnitResp, xmlResponseElem, True
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Create UserName request, and get User details
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim xmlUserReq As IXMLDOMNode
    Dim xmlUserNode As IXMLDOMNode
    Dim xmlUserResp As IXMLDOMNode
    
    Set xmlUserReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDom = xmlUserReq.ownerDocument
    Set xmlUserNode = xmlUserReq.appendChild(xmlDom.createElement("USERLIST"))
    
    BuildNode "USERID", xmlUserHistoryResp, xmlUserNode
    
    Set xmlUserResp = objIOrganisationBO.FindUserNameList(xmlUserReq)
    m_objErrAssist.CheckResponse xmlUserResp, xmlResponseElem, True
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' We now have the User, Unit, and UserHistory details, so we need to join
    ' them together into a set of rows
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlResp As IXMLDOMElement
    Dim colXML As Collection
    Dim colKeys As Collection
    
    Set colXML = New Collection
    Set colKeys = New Collection
    
    ' Need to join the following XML structures
    colXML.Add xmlUserHistoryResp
    colXML.Add xmlUnitResp
    colXML.Add xmlUserResp
    
    ' Unit key
    colKeys.Add "UNITID"
    
    ' User key
    colKeys.Add "USERID"
    
    ' Do the join
    Set xmlResp = JoinNodes("USERHISTORY", colXML, colKeys)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationManagerBO_FindApplicationOwnershipList = xmlResp

FindApplicationOwnershipListExit:

    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindApplicationOwnershipListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindApplicationOwnershipList = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindApplicationOwnershipListExit

End Function

Private Function JoinNodes(strMainTag As String, colXML As Collection, colKeys As Collection) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Joins any number of XML data streams as passed in via colXML. The first entry
'               is the main xml list, the rest are to be joined to the first.
' pass:         strMainTag contains the tag of the element to be appended to, colXML contains the main XML and
'               each piece of XML to be joined, and colKeys contains a list of keys which are indexed from the
'               first to the subsequence XML data streams.
' return:       xml Response Node
    Const strFunctionName As String = "JoinNodes"
    
    Dim nThisKey As Integer
    Dim nKeyCount As Integer
    Dim strSearch As String
    Dim strNodeKey As String
    Dim xmlMain As IXMLDOMElement
    Dim xmlMainList As IXMLDOMNodeList
    Dim xmlMainNode As IXMLDOMNode
    Dim xmlOtherNode As IXMLDOMElement
    Dim xmlOtherNodeItem As IXMLDOMElement
    Dim xmlNodeKey As IXMLDOMNode
    
    On Error GoTo JoinNodesErr
    
    ' First one contains the number of rows we want
    If colXML.Count <= 1 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingParameter, "XML collection must have more than one entry"
    End If
    
    If colKeys.Count = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingParameter, "XML key collection must have at least one entry"
    End If
    
    Set xmlMain = colXML(1).cloneNode(True)
    Set xmlMainList = xmlMain.getElementsByTagName(strMainTag)
    nKeyCount = colKeys.Count
    
    For Each xmlMainNode In xmlMainList
        
        For nThisKey = 1 To nKeyCount
            strNodeKey = colKeys(nThisKey)
            Set xmlNodeKey = m_objXmlAssist.GetNode(xmlMainNode, "./" & strNodeKey)
            
            ' GetNode should raise an error if it doesn't find the above node, but check anyway
            If Not xmlNodeKey Is Nothing Then
                Set xmlOtherNode = colXML(nThisKey + 1)
                
                ' We have the key, so look into the relevant other XML to find the node
                strSearch = "//* " & "[" & strNodeKey & " = '" & xmlNodeKey.Text & "']"
                Set xmlOtherNodeItem = xmlOtherNode.selectSingleNode(strSearch)
                
                If Not xmlOtherNodeItem Is Nothing Then
                    ' Remove the original to avoid unnecessary duplication
                    xmlMainNode.removeChild xmlNodeKey
                    ' Append the found node
                    xmlMainNode.appendChild xmlOtherNodeItem.cloneNode(True)
                End If
            End If
        Next
    Next

    Set JoinNodes = xmlMain
    
    Exit Function
JoinNodesErr:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub BuildNode(strTag As String, xmlRequest As IXMLDOMElement, xmlOutElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Takes an XML request tag and returns an XML stream containing a list of xml elements
'               taken from the request

' pass:         strTag - the tag to search for, xmlRequest - the request to be searched, xmlOutElem -
'               the XML to be appended to
' return:       xml Response Node
    Const strFunctionName As String = "BuildNode"
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    On Error GoTo BuildNodeErr
        
    If Not xmlRequest Is Nothing And Not xmlOutElem Is Nothing Then
        Set xmlNodeList = xmlRequest.getElementsByTagName(strTag)
        
        For Each xmlNode In xmlNodeList
           xmlOutElem.appendChild xmlNode.cloneNode(True)
        Next
    
    End If
    
    Exit Sub

BuildNodeErr:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IApplicationManagerBO_FindBusinessForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindBusinessForCustomerVbErr
    
    Const cstrFunctionName As String = "IApplicationManagerBO_FindBusinessForCustomer"

    
    
    ' XSL to sort by DATECREATED and APPLICATIONNUMBERORACCOUNTNUMBER
    ' BMIDS00980 Application and Account number now stored in separate nodes
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationManagerDO As IApplicationManagerDO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objAdminBO As Object
    Dim xmlApplicationList As IXMLDOMNodeList
    Dim xmlAccountList As IXMLDOMNodeList
    Dim xmlAdminRequestElem As IXMLDOMElement
    Dim xmlAdminResponseDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlMergedNode As IXMLDOMNode
    Dim xmlApplicationListElem As IXMLDOMElement
    Dim xmlAccountListElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objConvAssist As ConvertAssist
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    Dim xslDoc As FreeThreadedDOMDocument40
    
    Dim strDoAdminSearch As String
    Dim blnDoAdminSearch As String
    Dim blnDoingOmigaSearch As Boolean
    Dim blnDoingAdminSearch As Boolean
    Dim strAdminResponse As String
    Dim intNoOfApplications As Integer
    Dim intNoOfAccounts As Integer
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objConvAssist = New ConvertAssist
    
    Set objIApplicationManagerDO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    
    ' See if we need to do an administration system search
    ' PSC 11/10/2005 MAR57
    strDoAdminSearch = objGlobalParamDO.GetCurrentParameterByType("UseAdminFindBusForCustomer", "BOOLEAN")
    blnDoAdminSearch = objConvAssist.CSafeBool(strDoAdminSearch)
    
    If blnDoAdminSearch = True Then
        Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    End If
                       
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMER Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMER).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTOMER & " tag not found"
    End If
        
    blnDoingOmigaSearch = True
    Set xmlDataNode = objIApplicationManagerDO.FindOmigaBusinessForCustomer(xmlRequestNode.xml)
    blnDoingOmigaSearch = False
    
    ' Get the list of applications
    Set xmlApplicationListElem = xmlDataNode
    
    If Not xmlApplicationListElem Is Nothing Then
        m_objXmlAssist.ChangeNodeName xmlDataNode, "OMIGABUSINESS", "EXISTINGBUSINESS"
        Set xmlApplicationList = xmlDataNode.selectNodes(".//EXISTINGBUSINESS")
        intNoOfApplications = xmlApplicationList.length
    End If

    Dim strOtherSysCustomerNo
    strOtherSysCustomerNo = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//OTHERSYSTEMCUSTOMERNUMBER", _
                                                        False, False)
    
    ' Search admin system if required and we have a admin system customer number
    If blnDoAdminSearch = True And Len(strOtherSysCustomerNo) > 0 Then
    
        ' Convert input request to attribute based XML for call to AdminInterfaceBO
        Set xmlAdminRequestElem = xmlCreateAttributeBasedResponse(vxmlRequest, True)
        Set xmlElement = m_objXmlAssist.GetNode(xmlAdminRequestElem, ".//CUSTOMER")
        xmlElement.setAttribute "CUSTOMERNUMBER", strOtherSysCustomerNo
        
        xmlAdminRequestElem.setAttribute "OPERATION", "FindBusinessForCustomer"
        
        ' Search the administration system
        blnDoingAdminSearch = True
        strAdminResponse = objAdminBO.REQUEST(xmlAdminRequestElem.xml)
        Set xmlAdminResponseDoc = m_objXmlAssist.load(strAdminResponse, TypeName(Me), cstrFunctionName)
        m_objErrAssist.CheckResponse xmlAdminResponseDoc.documentElement, vxmlRequest, True
        blnDoingAdminSearch = False
        
        ' Convert the response back into element based and get a list of admin accounts
        Set xmlNode = m_objXmlAssist.GetNode(xmlAdminResponseDoc, ".//MORTGAGEACCOUNTLIST", False)
        
        If Not xmlNode Is Nothing Then
            Set xmlNode = xmlMakeNodeElementBased(xmlNode, True, "MORTGAGEACCOUNTLIST")
            m_objXmlAssist.ChangeNodeName xmlNode, "MORTGAGEACCOUNT", "EXISTINGBUSINESS"
            Set xmlAccountList = xmlNode.selectNodes(".//EXISTINGBUSINESS")
            intNoOfAccounts = xmlAccountList.length
        End If
    End If
    
    ' If there are applications and/or accounts merge and sort into the appropriate order
    If intNoOfApplications > 0 Or intNoOfAccounts > 0 Then
        Set xmlMergedNode = xmlOut.createElement("EXISTINGBUSINESSLIST")
        
        If intNoOfApplications > 0 Then
            For Each xmlElement In xmlApplicationList
                ' SR 12/06/01 : SYS2362 Add new element 'TYPE' which will have values 'Account' or 'Application'
                '               based on source of the record
                Set xmlNode = xmlOut.createElement("TYPE")
                xmlNode.Text = "Application"
                xmlElement.appendChild xmlNode
                ' SR 12/06/01
                
                xmlMergedNode.appendChild xmlElement
            Next
        End If
    
        If intNoOfAccounts > 0 Then
            For Each xmlElement In xmlAccountList
                ' SR 12/06/01 : SYS2362 Add new element 'TYPE' which will have values 'Account' or 'Application'
                '               based on source of the record
                Set xmlNode = xmlOut.createElement("TYPE")
                xmlNode.Text = "Account"
                xmlElement.appendChild xmlNode
                ' SR 12/06/01

                xmlMergedNode.appendChild xmlElement
            Next
        End If
                  
        ' Sort the list into the appropriate order and attach the result to the response
        Set xslDoc = New FreeThreadedDOMDocument40
        xslDoc.validateOnParse = False
        xslDoc.setProperty "NewParser", True
        xslDoc.async = False
        
        strXslPath = App.Path & "\FindBusinessForCustomer.xslt"
        strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
        blnLoadedOk = xslDoc.load(strXslPath)
        
        If Not blnLoadedOk Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
        End If
    
        Set xmlTransformedDoc = New FreeThreadedDOMDocument40
        xmlTransformedDoc.validateOnParse = False
        xmlTransformedDoc.setProperty "NewParser", True
        xmlTransformedDoc.async = False
        xmlMergedNode.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
        xmlResponseElem.appendChild xmlTransformedDoc.documentElement
    End If
            
    m_objContext.SetComplete
    
    Set IApplicationManagerBO_FindBusinessForCustomer = xmlResponseElem

FindBusinessForCustomerExit:

    Set xmlRequestNode = Nothing
    Set objIApplicationManagerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindBusinessForCustomerVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        
        ' Resume if we have no records from Omiga and we still have to do an Admin Search or
        ' We have no records on the Admin System  but we have records from Omiga
        If (blnDoingOmigaSearch = True And blnDoAdminSearch = True) Or _
           (blnDoingAdminSearch And intNoOfApplications > 0) Then
            Resume Next
        End If
    End If


    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_FindBusinessForCustomer = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindBusinessForCustomerExit

End Function

Public Function ImportAccountsIntoApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ImportAccountsIntoApplicationVbErr

    Const strFunctionName As String = "ImportAccountsIntoApplication"
        
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
    
    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, bomImportAccountsIntoApplication)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationManagerBO_ImportAccountsIntoApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ImportAccountsIntoApplication = xmlResponseElem.xml
    
    m_objContext.SetComplete

ImportAccountsIntoApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

ImportAccountsIntoApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ImportAccountsIntoApplication = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountsIntoApplicationExit

End Function

Private Function IApplicationManagerBO_ImportAccountsIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ImportAccountsIntoApplicationVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_ImportAccountsIntoApplication"
        
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationManagerTxBO As IApplicationManagerTxBO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objConvAssist As ConvertAssist
    
    'BMIDS918 GHun
    Const cstrTABLE_CUSTOMERLIST        As String = "CUSTOMERLIST"
    Dim objAdminBO                      As Object
    Dim objComboDO                      As ComboDO
    Dim objCustomerBO                   As CustomerBO
    Dim objIApplicationNTxBO            As IApplicationNTxBO
    Dim objCustFinancialBO              As CustomerFinancialBO
    Dim xmlCustomerList                 As IXMLDOMNodeList
    Dim xmlMortgageAccountList          As IXMLDOMNodeList
    Dim xmlAppNo                        As IXMLDOMNode
    Dim xmlAppFFNo                      As IXMLDOMNode
    Dim xmlOtherSystemAccountNumber     As IXMLDOMNode
    Dim xmlMortgageAccountNode          As IXMLDOMNode
    Dim xmlTempRequestNode              As IXMLDOMNode
    Dim xmlCustomerNumberNode           As IXMLDOMNode
    Dim xmlCustomerVersionNumberNode    As IXMLDOMNode
    Dim xmlCustomerNode                 As IXMLDOMNode
    Dim xmlListNode                     As IXMLDOMNode
    Dim xmlAddCustomer                  As IXMLDOMNode
    Dim xmlAddDataAppNode               As IXMLDOMNode
    Dim xmlElemBasedAdmin               As IXMLDOMNode
    Dim xmlAddDataRequestNode           As IXMLDOMNode
    Dim xmlElement                      As IXMLDOMElement
    Dim xmlAdminRequestElem             As IXMLDOMElement
    Dim xmlTableElem                    As IXMLDOMElement
    Dim xmlTempElement                  As IXMLDOMElement
    Dim xmlMortgageAccountListElement   As IXMLDOMElement
    Dim xmlListElem                     As IXMLDOMElement
    Dim xmlCustomerElem                 As IXMLDOMElement
    Dim xmlGetAddDataElem               As IXMLDOMElement
    Dim xmlAddDataElement               As IXMLDOMElement
    Dim xmlMortgageAccountDoc           As FreeThreadedDOMDocument40
    Dim xmlAddDataRequestDoc            As FreeThreadedDOMDocument40
    Dim xmlAdminResponseDoc             As FreeThreadedDOMDocument40
    Dim xmlTempRequestDoc               As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc              As FreeThreadedDOMDocument40
    Dim xmlCustomerAddressDoc           As FreeThreadedDOMDocument40
    Dim strAccountNo                    As String
    Dim strTypeOfApplication            As String
    Dim strAdminCustomerNumber          As String
    Dim strOmigaCustomerNumber          As String
    Dim strPattern                      As String
    Dim blnFurtherAdvance               As Boolean
    Dim strResponse                     As String
    'BMIDS918 End
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objConvAssist = New ConvertAssist
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")

    Dim strAdminSystem As String
    strAdminSystem = objGlobalParamDO.GetCurrentParameterByType("FindLegacyCustomer", "BOOLEAN")
       
    ' Check if we have an administration system
    If objConvAssist.CSafeBool(strAdminSystem) = True Then
    
        'BMIDS918 GHun Call GetAccountDetails here before the transaction begins, so that
        'saving the results have a larger slice of the transaction
        
        ' Check if the element passed in is the one with the table name else
        ' look below the node that is passed in
        If vxmlRequest.nodeName = cstrTABLE_CUSTOMERLIST Then
            Set xmlElement = vxmlRequest
        Else
            Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERLIST).Item(0)
        End If
        
        If xmlElement Is Nothing Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeMissingPrimaryTag, _
                                      cstrTABLE_CUSTOMERLIST & " tag not found"
        End If
        
        strTypeOfApplication = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//TYPEOFAPPLICATION", True)
        strAccountNo = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//ACCOUNTNUMBER", True, False)
        Set xmlAppNo = m_objXmlAssist.GetNode(vxmlRequest, ".//APPLICATIONNUMBER", True)
        Set xmlAppFFNo = m_objXmlAssist.GetNode(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER", True)
        
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        
        Set xmlTempRequestDoc = New FreeThreadedDOMDocument40
        xmlTempRequestDoc.validateOnParse = False
        xmlTempRequestDoc.setProperty "NewParser", True
        xmlTempRequestDoc.async = False
        Set xmlTempRequestNode = vxmlRequest.cloneNode(False)
        
        'Is this a Further Advance?
        blnFurtherAdvance = objComboDO.IsItemInValidation("TypeOfMortgage", strTypeOfApplication, "F")
        If blnFurtherAdvance = True Then
        
            If Len(strAccountNo) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), _
                                          strFunctionName, _
                                          oeMissingElementValue, _
                                          "Further advance requires account number"
            Else
                'BMIDS01043 Update OtherSystemAccountNumber for further advances. This needs to be done before calling
                'the admin system, so if anything goes wrong the OtherSystemAccountNumber is available for retrying
                Set xmlTableElem = xmlTempRequestDoc.createElement("APPLICATION")
                xmlTableElem.appendChild xmlAppNo.cloneNode(True)
                xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
                
                Set xmlOtherSystemAccountNumber = xmlTempRequestDoc.createElement("OTHERSYSTEMACCOUNTNUMBER")
                xmlOtherSystemAccountNumber.Text = strAccountNo
                xmlTableElem.appendChild xmlOtherSystemAccountNumber
                
                Set objIApplicationNTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationNTxBO")
                Set xmlTempElement = objIApplicationNTxBO.UpdateApplication(xmlTableElem)
                m_objErrAssist.CheckResponse xmlTempElement, xmlResponseElem, True
                'BMIDS01043 End
            End If
        End If
        
        '=========================================================================================
        ' Retrieve a list of Mortgage Accounts for all listed Customers
        '=========================================================================================
        Set xmlCustomerList = xmlElement.selectNodes("CUSTOMER")
        If xmlCustomerList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeMissingPrimaryTag, _
                                      cstrTABLE_CUSTOMER & " tag not found"
        End If
        
        'Initialise MortgageAccount document for later use
        Set xmlMortgageAccountDoc = New FreeThreadedDOMDocument40
        xmlMortgageAccountDoc.validateOnParse = False
        xmlMortgageAccountDoc.setProperty "NewParser", True
        xmlMortgageAccountDoc.async = False
        Set xmlMortgageAccountListElement = xmlMortgageAccountDoc.createElement("MORTGAGEACCOUNTLIST")
        
        ' PSC 16/10/2002 BMIDS00575
        xmlMortgageAccountDoc.appendChild xmlMortgageAccountListElement
               
        'Build Request
        Set xmlListElem = xmlTempRequestDoc.createElement("MORTGAGEACCOUNTLIST")
        xmlTempRequestNode.appendChild xmlListElem
        Set xmlTableElem = xmlTempRequestDoc.createElement("MORTGAGEACCOUNT")
        xmlListElem.appendChild xmlTableElem
        Set xmlCustomerNumberNode = xmlTempRequestDoc.createElement("CUSTOMERNUMBER")
        xmlTableElem.appendChild xmlCustomerNumberNode
        Set xmlCustomerVersionNumberNode = xmlTempRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
        xmlTableElem.appendChild xmlCustomerVersionNumberNode
        'BMIDS00006 GHun
        xmlTempRequestDoc.appendChild xmlListElem 'BMIDS00425
        'xmlTempRequestDoc.appendChild xmlTableElem
        'BMIDS00006 End
        
        Set objCustFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
        
        For Each xmlCustomerNode In xmlCustomerList
            'Find existing Mortgage Accounts
            Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlCustomerNode, ".//CUSTOMERNUMBER")
            xmlCustomerNumberNode.Text = xmlTempElement.Text
            Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlCustomerNode, ".//CUSTOMERVERSIONNUMBER")
            xmlCustomerVersionNumberNode.Text = xmlTempElement.Text
            strResponse = objCustFinancialBO.FindMortgageAccountList(xmlTempRequestDoc.xml)
            Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
                    
            'Append all Mortgage Accounts to overall list. Deduplication not required.
            'This list is used later to determine if an account to be imported already exists
            Set xmlMortgageAccountList = xmlTempResponseDoc.documentElement.selectNodes(".//MORTGAGEACCOUNT")
            For Each xmlMortgageAccountNode In xmlMortgageAccountList
                xmlMortgageAccountListElement.appendChild xmlMortgageAccountNode
            Next
            
        Next
            
        'JD BMIDS830 Get more information to send to GetAccountDetails
        Set xmlAddDataRequestDoc = New FreeThreadedDOMDocument40
        xmlAddDataRequestDoc.validateOnParse = False
        xmlAddDataRequestDoc.setProperty "NewParser", True
        xmlAddDataRequestDoc.async = False
        Set xmlAddDataRequestNode = vxmlRequest.cloneNode(False)
        xmlAddDataRequestDoc.appendChild xmlAddDataRequestNode
        
        Set xmlGetAddDataElem = xmlAddDataRequestDoc.createElement(cstrGetAddData) ' SR 06/08/2004
        xmlAddDataRequestNode.appendChild xmlGetAddDataElem
        Set xmlTableElem = xmlAddDataRequestDoc.createElement("APPLICATION") 'JD BMIDS830
        xmlTableElem.appendChild xmlAppNo.cloneNode(True)
        xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
        xmlGetAddDataElem.appendChild xmlTableElem
        Set xmlTableElem = xmlAddDataRequestDoc.createElement("CUSTOMERLIST")
        Set xmlListNode = xmlGetAddDataElem.appendChild(xmlTableElem)
        For Each xmlCustomerNode In xmlCustomerList
            Set xmlTableElem = xmlAddDataRequestDoc.createElement("CUSTOMER")
            xmlListNode.appendChild xmlTableElem
            Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlCustomerNode, ".//CUSTOMERNUMBER")
            xmlTableElem.appendChild xmlTempElement.cloneNode(True)
            Set xmlTempElement = m_objXmlAssist.GetMandatoryNode(xmlCustomerNode, ".//CUSTOMERVERSIONNUMBER")
            xmlTableElem.appendChild xmlTempElement.cloneNode(True)
        Next
        Set xmlAddDataElement = IApplicationManagerBO_GetAdditionalRequestDataForAccDownload(xmlAddDataRequestDoc.documentElement)
        m_objErrAssist.CheckResponse xmlAddDataElement, xmlResponseElem, True
        'BMIDS830 end
        
        '=========================================================================================
        ' Call Admin System for list of Accounts belonging to listed Customers
        '=========================================================================================
                   
        ' Convert input request to attribute based XML for call to AdminInterfaceBO
        Set xmlAdminRequestElem = xmlCreateAttributeBasedResponse(vxmlRequest, True)
        xmlAdminRequestElem.setAttribute "OPERATION", "GetAccountDetails"
        
        ' AdminBO's input customer number is the Other System CustomerNumber
        Set xmlCustomerList = xmlAdminRequestElem.selectNodes(".//CUSTOMER")
        For Each xmlCustomerElem In xmlCustomerList
            strAdminCustomerNumber = xmlGetAttributeText(xmlCustomerElem, "OTHERSYSTEMCUSTOMERNUMBER")
            
            ' PSC 23/10/2002 BMIDS00465 - Start
            ' Only send customers that are known to the admin system
            If Len(Trim$(strAdminCustomerNumber)) > 0 Then
                strOmigaCustomerNumber = xmlGetAttributeText(xmlCustomerElem, "CUSTOMERNUMBER")
                xmlCustomerElem.setAttribute "CUSTOMERNUMBER", strAdminCustomerNumber
                xmlCustomerElem.setAttribute "OMIGACUSTOMERNUMBER", strOmigaCustomerNumber
                ' DRC 25/02/2004 BMIDS718
                If Len(Trim$(strAccountNo)) > 0 Then
                    xmlCustomerElem.setAttribute "BMACCOUNTNUMBER", strAccountNo
                End If
                ' JD BMIDS830 add additional customer data
                strPattern = ".//CUSTOMER[CUSTOMERNUMBER = '" & strOmigaCustomerNumber & "']"
                Set xmlAddCustomer = m_objXmlAssist.GetNode(xmlAddDataElement, strPattern, False) ' SR 06/08/2004 : BMIDS830
                If Not xmlAddCustomer Is Nothing Then
                    xmlCustomerElem.setAttribute "MORTGAGELENDERBM", m_objXmlAssist.GetElementText(xmlAddCustomer, "OWNMORTGAGELENDERIND")
                    xmlCustomerElem.setAttribute "MORTGAGEREPAID", m_objXmlAssist.GetElementText(xmlAddCustomer, "MORTGAGEREPAID")
                    xmlCustomerElem.setAttribute "MORTGAGELET", m_objXmlAssist.GetElementText(xmlAddCustomer, "MORTGAGELET")
                    xmlCustomerElem.setAttribute "AXESS", m_objXmlAssist.GetElementText(xmlAddCustomer, "AXESSAPPLICATION")
                End If
            Else
                xmlCustomerElem.parentNode.removeChild xmlCustomerElem
            End If
            ' PSC 23/10/2002 BMIDS00465 - End
        
        Next
        ' DRC 25/02/2004 BMIDS718
        Set xmlTempElement = xmlAdminRequestElem.selectSingleNode(".//APPLICATION")
        xmlTempElement.setAttribute "MORTGAGESTATUSTYPE", strTypeOfApplication
        ' DRC 25/02/2004 BMIDS718 - End
        ' JD BMIDS830 add additional app data
        Set xmlAddDataAppNode = m_objXmlAssist.GetNode(xmlAddDataElement, ".//APPLICATION")
        If Not xmlAddDataAppNode Is Nothing Then
            xmlTempElement.setAttribute "MORTGAGESCHEME", m_objXmlAssist.GetElementText(xmlAddDataAppNode, "SPECIALSCHEME")
            xmlTempElement.setAttribute "MAINRESIDENCE", m_objXmlAssist.GetElementText(xmlAddDataAppNode, "MAINRESIDENCEIND")
        End If
        
        Set objAdminBO = m_objContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        strResponse = objAdminBO.REQUEST(xmlAdminRequestElem.xml)
        Set objAdminBO = Nothing
        Set xmlAdminResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlAdminResponseDoc.documentElement, xmlResponseElem, True
        
        ' If no records are found set the element based to the returned response else
        ' convert the response into element based. Work around until
        ' xmlMakeNodeElementBased is fixed
        Set xmlTempElement = m_objXmlAssist.GetNode(xmlAdminResponseDoc, ".//RESPONSE")
        If xmlTempElement.getAttribute("TYPE") <> "SUCCESS" Then
            Set xmlElemBasedAdmin = xmlAdminResponseDoc.documentElement
        Else
            Set xmlElemBasedAdmin = xmlMakeNodeElementBased(xmlAdminResponseDoc.documentElement, True, "")
        End If
        
        'Get the CustomerAddress once before hand, rather than once for each MortgageAccount
        Set xmlCustomerAddressDoc = New FreeThreadedDOMDocument40
        xmlCustomerAddressDoc.validateOnParse = False
        xmlCustomerAddressDoc.setProperty "NewParser", True
        Set xmlTempElement = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        xmlCustomerAddressDoc.appendChild xmlTempElement
        
        Set xmlTableElem = xmlCustomerAddressDoc.createElement("CUSTOMERADDRESS")
        xmlTempElement.appendChild xmlTableElem
        
        Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERNUMBER")
        xmlTableElem.appendChild xmlElement.cloneNode(True)
        
        Set xmlElement = m_objXmlAssist.GetNode(vxmlRequest, ".//CUSTOMERVERSIONNUMBER")
        xmlTableElem.appendChild xmlElement.cloneNode(True)
        
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        strResponse = objCustomerBO.FindCustomerAddressList(xmlCustomerAddressDoc.xml)
        Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
        
        vxmlRequest.appendChild xmlTempResponseDoc.selectSingleNode("RESPONSE/CUSTOMERADDRESSLIST")
        
        'Pass Admin response
        m_objXmlAssist.ChangeNodeName xmlElemBasedAdmin, "RESPONSE", "ADMINSYSTEMRESPONSE", False
        vxmlRequest.appendChild xmlElemBasedAdmin
        'Pass Mortage Account list
        vxmlRequest.appendChild xmlMortgageAccountDoc.documentElement
        'Pass through the further advance status so that it does not have to be requeried
        If blnFurtherAdvance Then
            vxmlRequest.setAttribute "IsFurtherAdvance", "1"
        Else
            vxmlRequest.setAttribute "IsFurtherAdvance", "0"
        End If
        
        'BMIDS918 End
    
        Set objIApplicationManagerTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationManagerTxBO")
        
        '   call Business Transaction Object ImportAccountsIntoApplication function
        Set xmlTempResponseNode = objIApplicationManagerTxBO.ImportAccountsIntoApplication(vxmlRequest)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    End If
   
    Set IApplicationManagerBO_ImportAccountsIntoApplication = xmlResponseElem

    m_objContext.SetComplete
    
ImportAccountsIntoApplicationExit:

    Set objIApplicationManagerTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing
    'BMIDS918 GHun
    Set objAdminBO = Nothing
    Set objComboDO = Nothing
    Set objIApplicationNTxBO = Nothing
    Set xmlElement = Nothing
    Set xmlAdminRequestElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFFNo = Nothing
    Set xmlOtherSystemAccountNumber = Nothing
    Set xmlTempElement = Nothing
    Set xmlCustomerList = Nothing
    Set xmlMortgageAccountDoc = Nothing
    Set xmlMortgageAccountList = Nothing
    Set xmlMortgageAccountNode = Nothing
    Set xmlMortgageAccountListElement = Nothing
    Set xmlListElem = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCustomerVersionNumberNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlAddDataRequestDoc = Nothing
    Set xmlAddDataRequestNode = Nothing
    Set xmlGetAddDataElem = Nothing
    Set xmlListNode = Nothing
    Set xmlAddCustomer = Nothing
    Set xmlAddDataAppNode = Nothing
    Set xmlAdminResponseDoc = Nothing
    Set xmlElemBasedAdmin = Nothing
    Set objCustFinancialBO = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlAddDataElement = Nothing
    Set objCustomerBO = Nothing
    Set xmlCustomerAddressDoc = Nothing
    'BMIDS918 End

    Exit Function

ImportAccountsIntoApplicationVbErr:
    'BMIDS918 GHun Ignore record not found errors
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    'BMIDS918 End

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_ImportAccountsIntoApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountsIntoApplicationExit

End Function

Public Function ImportAccountRefreshIntoApplication(ByVal vstrXMLRequest As String) As String

On Error GoTo ImportAccountRefreshIntoApplicationVbErr

    Const strFunctionName As String = "ImportAccountRefreshIntoApplication"
      
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
                                                                                                 
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IApplicationManagerBO_ImportAccountRefreshIntoApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ImportAccountRefreshIntoApplication = xmlResponseElem.xml
    
    m_objContext.SetComplete

ImportAccountRefreshIntoApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    Exit Function

ImportAccountRefreshIntoApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ImportAccountRefreshIntoApplication = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportAccountRefreshIntoApplicationExit

End Function

Public Function GetApplicationSummaryData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationSummaryDataVbErr

    Const strFunctionName As String = "GetApplicationSummaryData"
        
        
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
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationManagerBO_GetApplicationSummaryData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetApplicationSummaryData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetApplicationSummaryDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetApplicationSummaryDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetApplicationSummaryData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetApplicationSummaryDataExit

End Function

Private Function IApplicationManagerBO_GetApplicationSummaryData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationSummaryDataVbErr
    
    Const strFunctionName As String = "IApplicationManagerBO_GetApplicationSummaryData"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objCustEmploymentBO As CustomerEmploymentBO
    Dim objIApplicationBO As IApplicationBO
    Dim objIntermediaryBO  As IntermediaryBO
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCustomerRoleList As IXMLDOMNodeList
    Dim xmlCustomerRole As IXMLDOMElement
    Dim xmlOutApplication As IXMLDOMElement
    Dim xmlOutApplicant As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlEmployment As IXMLDOMElement
    Dim xmlCustomerNo As IXMLDOMElement
    Dim xmlCustomerVersNo As IXMLDOMElement
    Dim xmlOutEmployment As IXMLDOMElement
    Dim xmlAppNo As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    Dim xmlTypeOfApp As IXMLDOMElement
    Dim xmlCustomerRoleDoc As FreeThreadedDOMDocument40
    Dim xmlPhoneNumber As IXMLDOMElement
    Dim xmlIntermediaryGuid As IXMLDOMElement
    Dim xmlCountryCode As IXMLDOMElement
    Dim xmlAreaCode As IXMLDOMElement
    
    Dim xmlIncomeRequest As IXMLDOMElement               ' MAR1040
    Dim xmlOutIncomeSummary As IXMLDOMElement            ' MAR1040
    
    Dim strResponse As String
    Dim lngErrorNo As Long

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objCustEmploymentBO = _
        m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    Set objIApplicationBO = m_objContext.CreateInstance(App.Title & ".ApplicationBO")
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    Set objIntermediaryBO = m_objContext.CreateInstance(gstrINTERMEDIARY & ".IntermediaryBO")
                            
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Set xmlOutApplication = xmlOut.createElement("APPLICATION")
       
    strResponse = FindCustomersForApplication(vxmlRequest.xml)
    Set xmlCustomerRoleDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlCustomerRoleDoc.documentElement, xmlResponseElem, True
        
    ' Set up customer employment request
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    ' PSC 14/03/01 SYS2067 - Start
    Set xmlElement = xmlOut.createElement("EMPLOYMENTLIST")
    xmlTempRequest.appendChild xmlElement
    Set xmlEmployment = xmlOut.createElement("EMPLOYMENT")
    xmlElement.appendChild xmlEmployment
    ' PSC 14/03/01 SYS2067 - End
    
    Set xmlCustomerNo = xmlOut.createElement("CUSTOMERNUMBER")
    xmlEmployment.appendChild xmlCustomerNo
    Set xmlCustomerVersNo = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlEmployment.appendChild xmlCustomerVersNo
        
    Set xmlCustomerRoleList = xmlCustomerRoleDoc.selectNodes(".//CUSTOMERROLE")
    
    For Each xmlCustomerRole In xmlCustomerRoleList
                               
        ' Create new applicant in the output
        Set xmlOutApplicant = xmlOut.createElement("APPLICANT")
        xmlOutApplication.appendChild xmlOutApplicant
        
        ' Copy customer details to the output
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERNUMBER")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERVERSIONNUMBER")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERORDER")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//FIRSTFORENAME")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//SECONDFORENAME")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//OTHERFORENAMES")
        xmlOutApplicant.appendChild xmlElement
                
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//SURNAME")
        xmlOutApplicant.appendChild xmlElement
        
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//DATEOFBIRTH")
        xmlOutApplicant.appendChild xmlElement
        
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//TITLE")
        xmlOutApplicant.appendChild xmlElement

        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//TITLEOTHER")
        xmlOutApplicant.appendChild xmlElement

        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//MEMBEROFSTAFF")
        xmlOutApplicant.appendChild xmlElement
        
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//MARITALSTATUS")
        xmlOutApplicant.appendChild xmlElement
        
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERROLETYPE")
        xmlOutApplicant.appendChild xmlElement
        
        ' Set the customer number for this customer
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERNUMBER")
        xmlCustomerNo.Text = xmlElement.Text
        
        ' Set the customer version number number for this customer
        Set xmlElement = m_objXmlAssist.GetNode(xmlCustomerRole, ".//CUSTOMERVERSIONNUMBER")
        xmlCustomerVersNo.Text = xmlElement.Text

        ' Get list of employment details
        strResponse = objCustEmploymentBO.FindEmploymentList(xmlTempRequest.xml)
        Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
        On Error Resume Next
        
        m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, xmlResponseElem, True
        
        m_objErrAssist.SaveErr

        On Error GoTo GetApplicationSummaryDataVbErr
        
        lngErrorNo = m_objErrAssist.LastErr
        
        ' Record not found is OK
        If m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And _
           lngErrorNo <> 0 Then
            m_objErrAssist.ReRaise
        End If
                           
        ' Find the main employment
        Set xmlOutEmployment = m_objXmlAssist.GetNode(xmlResponseDoc, ".//EMPLOYMENT[MAINSTATUS = '1']", False)
        
        ' If main employment details present copy them to output
        If Not xmlOutEmployment Is Nothing Then
            Set xmlElement = m_objXmlAssist.GetNode(xmlOutEmployment, ".//OCCUPATIONTYPE")
            xmlOutApplicant.appendChild xmlElement
        
            Set xmlElement = m_objXmlAssist.GetNode(xmlOutEmployment, ".//EMPLOYMENTSTATUS", False)
            xmlOutApplicant.appendChild xmlElement

            Set xmlElement = m_objXmlAssist.GetNode(xmlOutEmployment, ".//DATESTARTEDORESTABLISHED", False)
            xmlOutApplicant.appendChild xmlElement
            
            ' PSC 12/12/01 SYS3318 - Start
            Set xmlElement = m_objXmlAssist.GetNode(xmlOutEmployment, ".//NETMONTHLYINCOME", False)
            xmlOutApplicant.appendChild xmlElement
            ' PSC 12/12/01 SYS3318 - End
        End If
        
        'MAR1040
        'Get Income SummaryData
        Set xmlIncomeRequest = xmlOut.createElement("INCOMESUMMARY")
        xmlIncomeRequest.appendChild xmlCustomerNo.cloneNode(True)
        xmlIncomeRequest.appendChild xmlCustomerVersNo.cloneNode(True)
    
        strResponse = objCustEmploymentBO.GetIncomeSummary(xmlIncomeRequest.xml)
        Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
        Set xmlOutIncomeSummary = m_objXmlAssist.GetNode(xmlResponseDoc, ".//INCOMESUMMARY", False)
        If Not xmlOutIncomeSummary Is Nothing Then
            Set xmlElement = m_objXmlAssist.GetNode(xmlOutIncomeSummary, ".//NETALLOWABLEANNUALINCOME", False)
            xmlOutApplicant.appendChild xmlElement

            Set xmlElement = m_objXmlAssist.GetNode(xmlOutIncomeSummary, ".//NETCONFIRMEDALLOWABLEINCOME", False)
            xmlOutApplicant.appendChild xmlElement
        End If
       
    Next
    
    ' Get the application data
    Set xmlTempResponse = objIApplicationDO.GetApplicationData(vxmlRequest)
    
    Set xmlIntermediaryGuid = m_objXmlAssist.GetNode(xmlTempResponse, ".//INTERMEDIARYGUID")
    
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONNUMBER")
    xmlOutApplication.appendChild xmlAppNo
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONFACTFINDNUMBER")
    xmlOutApplication.appendChild xmlElement
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//TYPEOFBUYER")
    xmlOutApplication.appendChild xmlElement
   
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONDATE")
    xmlOutApplication.appendChild xmlElement
       
    Set xmlTypeOfApp = m_objXmlAssist.GetNode(xmlTempResponse, ".//TYPEOFAPPLICATION")
    xmlOutApplication.appendChild xmlTypeOfApp
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONRECOMMENDEDDATE")
    xmlOutApplication.appendChild xmlElement
    
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONAPPROVALDATE")
    xmlOutApplication.appendChild xmlElement
    
    'BMIDS744
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//OPTOUTINDICATOR")
    xmlOutApplication.appendChild xmlElement
    
    ' Clone the request and add in type of application to get the mortgage property address data
    Set xmlTempRequest = vxmlRequest.cloneNode(True)
    Set xmlElement = m_objXmlAssist.GetNode(xmlTempRequest, ".//APPLICATION")
    xmlElement.appendChild xmlTypeOfApp.cloneNode(True)
        
    Set xmlTempResponse = objIApplicationBO.GetMortgagePropertyAddress(xmlTempRequest)
    
    On Error Resume Next
    
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True

    m_objErrAssist.SaveErr
    
    On Error GoTo GetApplicationSummaryDataVbErr
    
    lngErrorNo = m_objErrAssist.LastErr
    
    ' Record not found is OK
    If m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And _
       lngErrorNo <> 0 Then
       m_objErrAssist.ReRaise
    End If
    
    ' If an address was found
    If m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
                
        ' Copy Address nodes accross to output
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//POSTCODE")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//FLATNUMBER")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//BUILDINGORHOUSENAME")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//BUILDINGORHOUSENUMBER")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//STREET")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//DISTRICT")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//TOWN")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//COUNTY")
        xmlOutApplication.appendChild xmlElement
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//COUNTRY")
        xmlOutApplication.appendChild xmlElement
        
        ' Get the ACCESSTELEPHONENUMBER from the output
        Set xmlPhoneNumber = m_objXmlAssist.GetNode(xmlTempResponse, ".//ACCESSTELEPHONENUMBER", False)
        
        ' If ACCESSTELEPHONENUMBER is present then move it to output as TELEPHONENUMBER
        ' Else try to get the number from the customer details
        If Not xmlPhoneNumber Is Nothing Then
            Set xmlElement = xmlOut.createElement("TELEPHONENUMBER")
            xmlElement.Text = xmlPhoneNumber.Text
            xmlOutApplication.appendChild xmlElement
        
            'JR - SYSOmiplus24, output CountryCode and AreaCode if they exist
            Set xmlCountryCode = m_objXmlAssist.GetNode(xmlTempResponse, ".//COUNTRYCODE", False)
            If Not xmlCountryCode Is Nothing Then
                Set xmlElement = xmlOut.createElement("COUNTRYCODE")
                xmlElement.Text = xmlCountryCode.Text
                xmlOutApplication.appendChild xmlElement
            End If
            
            Set xmlAreaCode = m_objXmlAssist.GetNode(xmlTempResponse, ".//AREACODE", False)
            If Not xmlAreaCode Is Nothing Then
                Set xmlElement = xmlOut.createElement("AREACODE")
                xmlElement.Text = xmlAreaCode.Text
                xmlOutApplication.appendChild xmlElement
            End If
        Else
            Dim strCustomerNo As String
            Dim strCustomerVersNo As String
            Dim strPattern As String
            
            strCustomerNo = m_objXmlAssist.GetNodeValue(xmlCustomerRoleDoc, ".//CUSTOMERNUMBER")
            strCustomerVersNo = m_objXmlAssist.GetNodeValue(xmlCustomerRoleDoc, ".//CUSTOMERVERSIONNUMBER")
            
            'JR - Omiplus24, include call to AddTelephoneNumberNode method to replace existing code
            
            ' Find the preferred phone number to contact
            AddTelephoneNumberNode xmlCustomerRoleDoc, xmlOutApplication, strCustomerNo, strCustomerVersNo, "1"
            
            'Check if TelephoneNumber exists
            Set xmlElement = m_objXmlAssist.GetNode(xmlOutApplication, ".//TELEPHONENUMBER", False)
            
            ' If no preferred phone number get the first one
            If xmlElement Is Nothing Then
                AddTelephoneNumberNode xmlCustomerRoleDoc, xmlOutApplication, strCustomerNo, strCustomerVersNo
            End If
        End If
    End If
    
    ' Create Request for Application Priority
    Set xmlElement = xmlOut.createElement("APPLICATIONPRIORITY")
    xmlElement.appendChild xmlAppNo.cloneNode(True)
    
    On Error Resume Next
    
    Set xmlTempResponse = objIApplicationDO.GetApplicationPriority(xmlElement)
    
    m_objErrAssist.SaveErr
    
    On Error GoTo GetApplicationSummaryDataVbErr
    
    lngErrorNo = m_objErrAssist.LastErr
    
    ' Record not found is OK
    If m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound And _
       lngErrorNo <> 0 Then
       m_objErrAssist.ReRaise
    End If
    
    ' If application priority found
    If m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
        
        ' Copy application priority to output
        Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponse, ".//APPLICATIONPRIORITYVALUE")
        xmlOutApplication.appendChild xmlElement
    
    End If
    
    ' Get the intermediary data if the Intermediary Guid is present
    If Len(xmlIntermediaryGuid.Text) > 0 Then
        Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlElement = xmlOut.createElement("INTERMEDIARY")
        xmlElement.appendChild xmlIntermediaryGuid
        xmlTempRequest.appendChild xmlElement

        strResponse = objIntermediaryBO.GetIndividualIntermediary(xmlTempRequest.xml)
        Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlResponseDoc.documentElement, xmlResponseElem, True

        Set xmlElement = m_objXmlAssist.GetNode(xmlResponseDoc, ".//INTERMEDIARYPANELID")
        xmlOutApplication.appendChild xmlElement
    End If
    
    xmlResponseElem.appendChild xmlOutApplication

    Set IApplicationManagerBO_GetApplicationSummaryData = xmlResponseElem
    
    m_objContext.SetComplete
    
GetApplicationSummaryDataExit:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
    
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicationDO = Nothing
    Set objCustEmploymentBO = Nothing
    Set objIApplicationBO = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlCustomerRoleList = Nothing
    Set xmlCustomerRole = Nothing
    Set xmlOutApplication = Nothing
    Set xmlOutApplicant = Nothing
    Set xmlTempRequest = Nothing
    Set xmlElement = Nothing
    Set xmlEmployment = Nothing
    Set xmlCustomerNo = Nothing
    Set xmlCustomerVersNo = Nothing
    Set xmlOutEmployment = Nothing
    Set xmlAppNo = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTypeOfApp = Nothing
    Set xmlCustomerRoleDoc = Nothing
    Set xmlIntermediaryGuid = Nothing
    'JR SYS2459
    Set xmlAreaCode = Nothing
    Set xmlCountryCode = Nothing
    

    Exit Function

GetApplicationSummaryDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationManagerBO_GetApplicationSummaryData = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetApplicationSummaryDataExit

End Function

Private Sub AddTelephoneNumberNode(ByVal xmlNodeToSearch As IXMLDOMNode, _
                                   ByVal vxmlResponseNode As IXMLDOMNode, _
                                   ByVal strCustomerNo As String, _
                                   ByVal strCustomerVersNo As String, _
                                   Optional strPreferredContact As String)

On Error GoTo AddTelephoneNumberNodeErr
   
    Dim strPattern As String
    Const strFunctionName As String = "AddTelephoneNumberNode"
    
    Dim intLoop As Integer
    
    Dim vArray As Variant
    
    Dim xmlElement As IXMLDOMElement
    
    vArray = Array("COUNTRYCODE", "AREACODE", "TELEPHONENUMBER")

    For intLoop = 0 To 2
    
        strPattern = ".//CUSTOMERTELEPHONENUMBER[CUSTOMERNUMBER = '"
        strPattern = strPattern & strCustomerNo & "' and CUSTOMERVERSIONNUMBER = '"
        strPattern = strPattern & strCustomerVersNo & "'"
    
        If strPreferredContact <> "" Then
            strPattern = strPattern & " and PREFERREDMETHODOFCONTACT = '" & strPreferredContact & "'"
        End If
        
        strPattern = strPattern & "]/" & vArray(intLoop)
    
        Set xmlElement = m_objXmlAssist.GetNode(xmlNodeToSearch, strPattern, False)
        If Not xmlElement Is Nothing Then
            vxmlResponseNode.appendChild xmlElement
        End If
    Next

    Set xmlElement = Nothing

    Exit Sub

AddTelephoneNumberNodeErr:

    Set xmlElement = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function ImportCustomersIntoApplication(ByVal vstrXMLRequest As String) As String

On Error GoTo ImportCustomersIntoApplicationVbErr

    Const strFunctionName As String = "ImportCustomersIntoApplication"
    
        
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
    
'    Set xmlTempResponseNode = IApplicationManagerBO_Validate(xmlIn.documentElement, bomImportCustomersIntoApplication)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationManagerBO_ImportCustomersIntoApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ImportCustomersIntoApplication = xmlResponseElem.xml
    
    m_objContext.SetComplete

ImportCustomersIntoApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

ImportCustomersIntoApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ImportCustomersIntoApplication = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ImportCustomersIntoApplicationExit

End Function
Public Function GetAdditionalRequestDataForAccDownload(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get specific application and customer information to pass to GetAccountDetails
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetAddDataVbErr

    Const strFunctionName As String = "GetAdditionalRequestDataForAccDownload"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationManagerBO_GetAdditionalRequestDataForAccDownload(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAdditionalRequestDataForAccDownload = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetAddDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAddDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAdditionalRequestDataForAccDownload = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetAddDataExit

End Function
